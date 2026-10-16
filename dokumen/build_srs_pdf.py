#!/usr/bin/env python3
"""Prepare print-ready HTML from the SRS Markdown and embedded diagrams."""

from __future__ import annotations

import base64
import html
import re
import shutil
import subprocess
from pathlib import Path


HERE = Path(__file__).resolve().parent
SOURCE = HERE / "SRS_Sistem_Pendataan_TSP_CSR_Samarinda.md"
SVG = HERE / "use_case_tsp_csr.svg"
HTML_OUTPUT = HERE / "SRS_Sistem_Pendataan_TSP_CSR_Samarinda_render.html"


def inline(text: str) -> str:
    image_pattern = r"!\[([^\]]*)\]\(([^)]+)\)"
    images: dict[str, str] = {}

    def embed_image(match: re.Match[str]) -> str:
        image_path = (HERE / match.group(2)).resolve()
        converter = shutil.which("rsvg-convert")
        if not converter:
            raise SystemExit("rsvg-convert diperlukan untuk menyematkan diagram.")
        raster = subprocess.run(
            [converter, "-f", "png", "-w", "680", str(image_path)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        ).stdout
        image_data = base64.b64encode(raster).decode("ascii")
        token = f"IMAGE_PLACEHOLDER_{len(images)}"
        images[token] = (
            f'<img class="diagram" alt="{html.escape(match.group(1))}" '
            f'src="data:image/png;base64,{image_data}">'
        )
        return token

    text = re.sub(image_pattern, embed_image, text)

    result = html.escape(text, quote=False)
    result = re.sub(r"`(.+?)`", r"<code>\1</code>", result)
    result = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", result)
    result = re.sub(r"\*(.+?)\*", r"<em>\1</em>", result)
    result = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r'<a href="\2">\1</a>', result)
    for token, image_html in images.items():
        result = result.replace(token, image_html)
    return result


def render_markdown(source: str) -> str:
    lines = source.splitlines()
    out: list[str] = []
    paragraph: list[str] = []
    list_open = False
    i = 0

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            text = " ".join(part.strip() for part in paragraph)
            if text.startswith("> "):
                out.append(f'<blockquote>{inline(text[2:])}</blockquote>')
            elif text.startswith("!["):
                out.append(f'<figure class="diagram-block">{inline(text)}</figure>')
            else:
                out.append(f"<p>{inline(text)}</p>")
            paragraph = []

    def close_list() -> None:
        nonlocal list_open
        if list_open:
            out.append("</ul>")
            list_open = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            flush_paragraph()
            close_list()
            i += 1
            continue

        if stripped.startswith("|"):
            flush_paragraph()
            close_list()
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [cell.strip() for cell in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", cell) for cell in cells):
                    rows.append(cells)
                i += 1
            if rows:
                out.append('<table><thead><tr>')
                out.extend(f"<th>{inline(cell)}</th>" for cell in rows[0])
                out.append("</tr></thead><tbody>")
                for row in rows[1:]:
                    out.append("<tr>")
                    for cell in row:
                        out.append(f"<td>{inline(cell)}</td>")
                    out.append("</tr>")
                out.append("</tbody></table>")
            continue

        heading = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if heading:
            flush_paragraph()
            close_list()
            level = len(heading.group(1))
            title = heading.group(2)
            css_class = ' class="diagram-heading"' if title.startswith(("7.4 ", "8.2 ")) else ""
            out.append(f"<h{level}{css_class}>{inline(title)}</h{level}>")
            i += 1
            continue

        if stripped.startswith("- "):
            flush_paragraph()
            if not list_open:
                out.append("<ul>")
                list_open = True
            out.append(f"<li>{inline(stripped[2:])}</li>")
            i += 1
            continue

        close_list()
        if stripped.startswith("> "):
            flush_paragraph()
            paragraph.append(stripped)
        else:
            paragraph.append(stripped)
        i += 1

    flush_paragraph()
    close_list()
    return "\n".join(out)


def main() -> None:
    body = render_markdown(SOURCE.read_text(encoding="utf-8"))
    # A source-file reference is useful in the Markdown project, but not in a
    # standalone PDF delivered outside this workspace.
    body = re.sub(r"<p>Sumber diagram yang dapat diedit:.*?</p>", "", body)
    page = f"""<!doctype html>
<html lang="id"><head><meta charset="utf-8"><title>SRS TSP/CSR Samarinda</title>
<style>
@page {{ size: A4; margin: 16mm 15mm 17mm; }}
* {{ box-sizing: border-box; }}
body {{ font-family: Arial, Helvetica, sans-serif; color: #172033; font-size: 9.4pt; line-height: 1.42; }}
h1 {{ color: #12335b; font-size: 22pt; line-height: 1.15; margin: 0 0 7pt; }}
h1 + h2 {{ font-size: 15pt; border: 0; margin: 0 0 16pt; }}
h2 {{ color: #12335b; font-size: 15pt; border-bottom: 1px solid #dbe4ee; padding-bottom: 4pt; margin: 21pt 0 8pt; page-break-after: avoid; }}
h3 {{ color: #1d4f7a; font-size: 11.5pt; margin: 14pt 0 5pt; page-break-after: avoid; }}
h4 {{ color: #334155; font-size: 10pt; margin: 10pt 0 4pt; page-break-after: avoid; }}
p {{ margin: 4pt 0 8pt; orphans: 3; widows: 3; }}
ul {{ margin: 4pt 0 10pt; padding-left: 20pt; }}
li {{ margin: 0 0 3pt; }}
blockquote {{ margin: 10pt 0; padding: 9pt 12pt; background: #f1f6fb; border-left: 3pt solid #2563eb; color: #334155; }}
table {{ width: 100%; border-collapse: collapse; table-layout: auto; margin: 7pt 0 12pt; font-size: 8pt; }}
thead {{ display: table-header-group; }}
th {{ background: #eaf1f8; color: #12335b; font-weight: 700; text-align: left; }}
th, td {{ border: 1px solid #cbd5e1; padding: 4pt 5pt; vertical-align: top; overflow-wrap: anywhere; }}
tr {{ page-break-inside: avoid; }}
code {{ font-family: Menlo, Consolas, monospace; font-size: .92em; background: #f1f5f9; padding: 1pt 2pt; }}
a {{ color: #1d4ed8; text-decoration: none; }}
.diagram-heading {{ page-break-before: always; }}
.diagram-block {{ margin: 8pt auto; text-align: center; page-break-inside: avoid; }}
.diagram {{ width: 100%; max-height: 230mm; object-fit: contain; }}
</style></head><body>{body}</body></html>"""

    HTML_OUTPUT.write_text(page, encoding="utf-8")
    print(f"HTML siap untuk diekspor ke PDF: {HTML_OUTPUT}")


if __name__ == "__main__":
    main()
