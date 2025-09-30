# Write a robust Markdown -> DOCX converter module to a downloadable file.
# The file provides: markdown_to_docx(...) with Pandoc-first pipeline and an HTML fallback.
from pathlib import Path

code = r'''# markdown_to_docx.py
# -*- coding: utf-8 -*-
"""
A robust Markdown → DOCX converter with an optional top logo.

Design goals
------------
- **Very robust**: prefers Pandoc (via `pypandoc`) for best fidelity
  (tables, footnotes, task lists/checkboxes, GFM, definition lists,
  math to native Word equations, etc.).
- **Graceful fallback**: if Pandoc is unavailable, falls back to
  Python-Markdown → HTML → DOCX (via `html2docx`). This covers most
  common Markdown, though advanced features (e.g., complex tables,
  math) are best with Pandoc.
- **AI output friendly**: pre-processes slightly messy Markdown
  (e.g., unclosed code fences).
- **Logo insertion**: optional PNG (or other image) at the very top,
  with configurable width.
- **Resource resolution**: tries to resolve relative image paths.

Install
-------
Recommended (Pandoc path — highest fidelity):

    pip install pypandoc
    # and install Pandoc itself (one of):
    #  - https://pandoc.org/installing.html
    #  - brew install pandoc
    #  - choco install pandoc
    #  - winget install --id JohnMacFarlane.Pandoc
    # (Optionally: pypandoc.download_pandoc() can fetch a local copy.)

Fallback path (no Pandoc):

    pip install markdown html2docx python-docx pygments

Usage
-----
    from markdown_to_docx import markdown_to_docx

    markdown_text = \"\"\"# Title

    Some text with **bold**, tables, and `code`.

    ```python
    print("hello")
    ```

    - [x] task
    - [ ] another task
    \"\"\"

    markdown_to_docx(
        markdown_text,
        "out.docx",
        logo_path="logo.png",             # optional
        logo_width_inches=1.5,            # optional
        add_toc=True,                     # optional
        toc_depth=3,                      # optional
        reference_docx=None,              # optional Pandoc docx template
        base_path=".",                    # where to resolve relative images
        prefer_pandoc=True,               # try Pandoc first
        highlight_style="pygments"        # Pandoc code highlighting style
    )

Command line (optional):
    python markdown_to_docx.py -i input.md -o out.docx --logo logo.png --toc

Notes
-----
- Pandoc writer for DOCX natively converts LaTeX-style math in Markdown
  to Word equations (OMML). The fallback HTML path does not support
  equation rendering.
- For best control of fonts/styles, provide a Pandoc reference DOCX
  (created from a Word template) via `reference_docx`.
"""

from __future__ import annotations

import io
import os
import re
import sys
import shutil
import tempfile
from pathlib import Path
from typing import Optional, Sequence, List


def _ensure_closed_code_fences(text: str) -> str:
    """
    Ensures any unclosed triple backtick/tilde code fences are closed.

    This is a light heuristic to tolerate AI-generated Markdown that
    sometimes omits a closing fence.
    """
    lines = text.split("\n")
    fence = None  # '```' or '~~~'
    for ln in lines:
        s = ln.strip()
        if s.startswith("```"):
            if fence is None:
                fence = "```"
            elif fence == "```":
                fence = None
        elif s.startswith("~~~"):
            if fence is None:
                fence = "~~~"
            elif fence == "~~~":
                fence = None
    if fence is not None:
        # append a matching closer
        lines.append(fence)
    return "\n".join(lines)


def _preprocess_markdown(md: str) -> str:
    """Normalize line endings, strip BOM, detab, and close unbalanced code fences."""
    if md is None:
        md = ""
    # Normalize newlines
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    # Strip BOM if present
    md = md.lstrip("\ufeff")
    # Detab (AI output often uses tabs inconsistently)
    md = md.replace("\t", "    ")
    # Fix any unclosed fences
    md = _ensure_closed_code_fences(md)
    return md


def _attempt_download_pandoc() -> bool:
    """
    Try to download a private copy of pandoc via pypandoc if the system
    doesn't have it. Returns True if pandoc is usable afterwards.
    """
    try:
        import pypandoc  # type: ignore
        try:
            # Will raise OSError if pandoc isn't found
            _ = pypandoc.get_pandoc_path()
            return True
        except OSError:
            # Try to download a local copy (may fail without internet)
            try:
                pypandoc.download_pandoc()
                _ = pypandoc.get_pandoc_path()
                return True
            except Exception:
                return False
    except Exception:
        return False


def _pandoc_available() -> bool:
    try:
        import pypandoc  # type: ignore
        try:
            _ = pypandoc.get_pandoc_path()
            return True
        except OSError:
            return False
    except Exception:
        return False


def _inject_logo_for_pandoc(logo_path: Path, width_in: float) -> str:
    """
    Build a Pandoc/Markdown image line with width attribute (implicit figure).
    Pandoc will embed the image in DOCX.
    """
    # Use POSIX path to be safe across platforms in markdown text
    posix = logo_path.resolve().as_posix()
    # Use implicit figure with width attribute in inches
    return f'![]({posix}){{width={width_in:.3f}in}}\n\n'


def _inject_logo_for_html(logo_path: Path, width_in: float) -> str:
    """
    Build an HTML <img> tag to prepend when using the HTML fallback.
    Convert inches to pixels at 96 DPI for broad compatibility.
    """
    px = int(round(width_in * 96))
    # Absolute path is safer for local embedding
    src = logo_path.resolve().as_uri() if hasattr(logo_path, "as_uri") else str(logo_path.resolve())
    return f'<p><img src="{src}" width="{px}" /></p>\n\n'


def _run_pandoc_pipeline(
    markdown_text: str,
    output_file_path: Path,
    logo_path: Optional[Path],
    logo_width_inches: float,
    add_toc: bool,
    toc_depth: int,
    reference_docx: Optional[Path],
    base_path: Optional[Path],
    highlight_style: Optional[str],
) -> None:
    """
    Convert via pypandoc → DOCX for highest fidelity.
    """
    import pypandoc  # type: ignore

    # Compose markdown (insert logo at top if provided)
    md = _preprocess_markdown(markdown_text)
    if logo_path:
        md = _inject_logo_for_pandoc(logo_path, logo_width_inches) + md

    # Write to a temp .md file to allow resource resolution
    with tempfile.TemporaryDirectory() as td:
        tmp_md = Path(td) / "input.md"
        tmp_md.write_text(md, encoding="utf-8")

        # Pandoc extensions for robust GFM markdown
        # Avoid deprecated '+smart' in Pandoc ≥ 3.0
        from_fmt = (
            "gfm"
            "+footnotes"
            "+definition_lists"
            "+task_lists"
            "+pipe_tables"
            "+table_captions"
            "+strikeout"
            "+superscript"
            "+subscript"
            "+yaml_metadata_block"
            "+tex_math_dollars"
            "+implicit_figures"
            "+autolink_bare_uris"
            "+emoji"
            "+raw_html"
        )

        extra_args: List[str] = ["--from", from_fmt, "--standalone"]

        # Resource path(s) for images referenced in the markdown
        resource_paths: List[str] = []
        if base_path:
            resource_paths.append(str(base_path.resolve()))
        if logo_path:
            resource_paths.append(str(logo_path.resolve().parent))
        # Always include current working directory last
        resource_paths.append(str(Path.cwd().resolve()))
        extra_args.extend(["--resource-path", os.pathsep.join(resource_paths)])

        # Optional: template for styles
        if reference_docx:
            extra_args.extend(["--reference-doc", str(reference_docx.resolve())])

        # Optional: TOC
        if add_toc:
            extra_args.extend(["--toc", "--toc-depth", str(int(toc_depth))])

        # Optional: syntax highlighting theme
        if highlight_style:
            extra_args.extend(["--highlight-style", str(highlight_style)])

        # Ensure output directory exists
        output_file_path.parent.mkdir(parents=True, exist_ok=True)

        # Convert
        pypandoc.convert_file(
            str(tmp_md),
            to="docx",
            outputfile=str(output_file_path),
            extra_args=extra_args,
        )


def _run_html_fallback_pipeline(
    markdown_text: str,
    output_file_path: Path,
    logo_path: Optional[Path],
    logo_width_inches: float,
    add_toc: bool,
    toc_depth: int,
    base_path: Optional[Path],
) -> None:
    """
    Fallback: Python-Markdown → HTML → DOCX (html2docx).
    """
    # Local imports so the module can be imported even if deps aren't present.
    import markdown as md  # type: ignore
    from html2docx import html2docx  # type: ignore

    body = _preprocess_markdown(markdown_text)

    # If TOC requested, insert marker recognized by Python-Markdown's 'toc' extension.
    if add_toc and "[TOC]" not in body:
        body = "[TOC]\n\n" + body

    # Prepend logo as an HTML <img>, sized in pixels (robust for html2docx)
    if logo_path:
        body = _inject_logo_for_html(logo_path, logo_width_inches) + body

    extensions = [
        "extra",
        "admonition",
        "attr_list",
        "codehilite",
        "def_list",
        "fenced_code",
        "footnotes",
        "md_in_html",
        "sane_lists",
        "smarty",
        "tables",
        "toc",
    ]
    extension_configs = {
        "codehilite": {"guess_lang": False, "use_pygments": True},
        "toc": {"permalink": True, "toc_depth": f"2-{int(toc_depth)}"},
    }

    html = md.markdown(
        body,
        extensions=extensions,
        extension_configs=extension_configs,
        output_format="xhtml1",
    )

    # Convert HTML → DOCX
    # html2docx.html2docx returns either a python-docx Document
    # or bytes (depending on version). Handle both.
    base_url = str(base_path.resolve()) if base_path else ""
    result = html2docx(html, title=None, base_url=base_url)  # type: ignore

    out_path = output_file_path
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Try to save via python-docx Document API; if not, assume bytes-like
    saved = False
    try:
        # python-docx Document has .save()
        result.save(str(out_path))  # type: ignore[attr-defined]
        saved = True
    except Exception:
        try:
            # html2docx might return bytes; write directly
            data = bytes(result) if not isinstance(result, (bytes, bytearray)) else result
            with open(out_path, "wb") as f:
                f.write(data)  # type: ignore[arg-type]
            saved = True
        except Exception as e:
            raise RuntimeError(f"html2docx fallback could not write output: {e}") from e

    if not saved:
        raise RuntimeError("html2docx fallback did not produce a savable document.")


def markdown_to_docx(
    markdown_text: str,
    output_file_path: str,
    *,
    logo_path: Optional[str] = None,
    logo_width_inches: float = 1.5,
    add_toc: bool = False,
    toc_depth: int = 3,
    reference_docx: Optional[str] = None,
    base_path: Optional[str] = None,
    prefer_pandoc: bool = True,
    highlight_style: Optional[str] = "pygments",
) -> None:
    """
    Convert Markdown text to a well-formatted .docx file.

    Parameters
    ----------
    markdown_text : str
        The Markdown content to convert. YAML front matter (metadata) is supported.
    output_file_path : str
        Path to write the resulting .docx file.
    logo_path : Optional[str], default None
        Path to a PNG (or other image) to insert at the very top.
    logo_width_inches : float, default 1.5
        Width of the logo in inches.
    add_toc : bool, default False
        If True, generates a table of contents.
    toc_depth : int, default 3
        Depth of the table of contents (when enabled).
    reference_docx : Optional[str], default None
        Pandoc reference .docx template to control styles (Pandoc pipeline only).
    base_path : Optional[str], default None
        Base path to resolve relative images/links in the Markdown.
    prefer_pandoc : bool, default True
        Try to use Pandoc first for best fidelity. If unavailable (or
        conversion fails), the function falls back to HTML → DOCX.
    highlight_style : Optional[str], default "pygments"
        Pandoc syntax highlighting style (e.g., "pygments", "tango",
        "kate", "monochrome"). Set to None to disable.

    Raises
    ------
    RuntimeError if both Pandoc and the HTML fallback fail.
    """
    out = Path(output_file_path)
    if out.suffix.lower() != ".docx":
        out = out.with_suffix(".docx")

    logo_p: Optional[Path] = Path(logo_path).expanduser().resolve() if logo_path else None
    if logo_p and not logo_p.exists():
        raise FileNotFoundError(f"Logo file not found: {logo_p}")

    ref_docx_p: Optional[Path] = Path(reference_docx).expanduser().resolve() if reference_docx else None
    if ref_docx_p and not ref_docx_p.exists():
        raise FileNotFoundError(f"Reference DOCX not found: {ref_docx_p}")

    base_p: Optional[Path] = Path(base_path).expanduser().resolve() if base_path else None
    if base_p and not base_p.exists():
        raise FileNotFoundError(f"Base path not found: {base_p}")

    last_error: Optional[Exception] = None

    # Pandoc preferred
    if prefer_pandoc:
        try:
            if _pandoc_available() or _attempt_download_pandoc():
                _run_pandoc_pipeline(
                    markdown_text=markdown_text,
                    output_file_path=out,
                    logo_path=logo_p,
                    logo_width_inches=logo_width_inches,
                    add_toc=add_toc,
                    toc_depth=toc_depth,
                    reference_docx=ref_docx_p,
                    base_path=base_p,
                    highlight_style=highlight_style,
                )
                return
        except Exception as e:
            last_error = e  # keep and try fallback

    # Fallback: HTML path
    try:
        _run_html_fallback_pipeline(
            markdown_text=markdown_text,
            output_file_path=out,
            logo_path=logo_p,
            logo_width_inches=logo_width_inches,
            add_toc=add_toc,
            toc_depth=toc_depth,
            base_path=base_p,
        )
        return
    except Exception as e:
        if last_error is None:
            last_error = e

    raise RuntimeError(
        "Failed to convert Markdown to DOCX using both Pandoc and the HTML fallback. "
        f"Last error: {last_error}"
    )


def _read_file_text(p: Path) -> str:
    return p.read_text(encoding="utf-8")


def _cli() -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert Markdown to DOCX (Pandoc preferred; HTML fallback)."
    )
    parser.add_argument("-i", "--input", help="Input Markdown file path. If omitted, reads STDIN.")
    parser.add_argument("-o", "--output", required=True, help="Output DOCX file path.")
    parser.add_argument("--logo", help="Optional logo image (PNG recommended) to insert at top.")
    parser.add_argument("--logo-width", type=float, default=1.5, help="Logo width in inches (default: 1.5).")
    parser.add_argument("--toc", action="store_true", help="Include a table of contents.")
    parser.add_argument("--toc-depth", type=int, default=3, help="TOC depth (default: 3).")
    parser.add_argument("--reference-docx", help="Pandoc reference DOCX template path.")
    parser.add_argument("--base-path", help="Base path for resolving relative image paths.")
    parser.add_argument(
        "--no-pandoc",
        action="store_true",
        help="Do not use Pandoc; force HTML fallback.",
    )
    parser.add_argument(
        "--highlight-style",
        default="pygments",
        help="Pandoc code highlight style (e.g., pygments, tango, kate, monochrome). Use '' to disable.",
    )

    args = parser.parse_args()

    # Read markdown
    if args.input:
        md_text = _read_file_text(Path(args.input).expanduser().resolve())
    else:
        md_text = sys.stdin.read()

    try:
        markdown_to_docx(
            md_text,
            args.output,
            logo_path=args.logo,
            logo_width_inches=float(args.logo_width),
            add_toc=bool(args.toc),
            toc_depth=int(args.toc_depth),
            reference_docx=args.reference_docx,
            base_path=args.base_path,
            prefer_pandoc=not bool(args.no_pandoc),
            highlight_style=(args.highlight_style or None),
        )
    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(_cli())
'''
Path('/mnt/data/markdown_to_docx.py').write_text(code, encoding='utf-8')
print("Wrote /mnt/data/markdown_to_docx.py")
