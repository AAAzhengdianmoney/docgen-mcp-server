"""
DocGen MCP Server — Generate Word (.docx) and PDF documents via MCP.
Cross-platform: Windows, macOS, and Linux. CJK font auto-detection.

Built with FastMCP. Install: pip install mcp python-docx fpdf2
"""
from __future__ import annotations

import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from fpdf import FPDF
from fpdf.enums import WrapMode
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    name="DocGenerator",
    instructions="Generate Word (.docx) and PDF documents",
)

# ── CJK font detection ────────────────────────────────────────────
_CJK_FONT = None

_CJK_CANDIDATES = [
    # Windows
    "C:/Windows/Fonts/msyh.ttc",       # Microsoft YaHei
    "C:/Windows/Fonts/msyhbd.ttc",     # Microsoft YaHei Bold
    "C:/Windows/Fonts/simsun.ttc",     # SimSun
    "C:/Windows/Fonts/simhei.ttf",     # SimHei
    # macOS
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    # Linux
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
]

for p in _CJK_CANDIDATES:
    if Path(p).exists():
        _CJK_FONT = p
        break


# ── Tools ──────────────────────────────────────────────────────────

@mcp.tool()
def generate_docx(
    output_path: str,
    title: str = "",
    sections: list[dict] = [],
) -> str:
    """Generate a Word (.docx) document with title and sections.

    Parameters:
        output_path: Absolute path to save the .docx file.
        title: Document title (centered, large heading).
        sections: List of section dicts. Each: {"heading": "...", "body": "...", "style": "normal|bullet|numbered"}.
    """
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(11)

    if title:
        p = doc.add_heading(title, level=0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for s in sections:
        heading = s.get("heading", "")
        body = s.get("body", "")
        sn = s.get("style", "normal")
        if heading:
            doc.add_heading(heading, level=1)
        if not body:
            continue
        if sn == "bullet":
            for line in body.strip().split("\n"):
                if line.strip():
                    doc.add_paragraph(line.strip(), style="List Bullet")
        elif sn == "numbered":
            for line in body.strip().split("\n"):
                if line.strip():
                    doc.add_paragraph(line.strip(), style="List Number")
        else:
            doc.add_paragraph(body)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return f"OK: {out.resolve()} ({out.stat().st_size} bytes)"


class _PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "", 8)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


@mcp.tool()
def generate_pdf(
    output_path: str,
    title: str = "",
    sections: list[dict] = [],
) -> str:
    """Generate a PDF document with title and sections.

    Parameters:
        output_path: Absolute path to save the .pdf file.
        title: Document title.
        sections: List of section dicts. Each: {"heading": "...", "body": "..."}.
    """
    pdf = _PDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    if _CJK_FONT:
        pdf.add_font("CJK", "", _CJK_FONT)
        font = "CJK"
    else:
        font = "Helvetica"

    if title:
        pdf.set_font(font, "", 18)
        pdf.cell(0, 10, title, align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(5)

    for s in sections:
        heading = s.get("heading", "")
        body = s.get("body", "")
        if heading:
            pdf.set_font(font, "", 14)
            pdf.cell(0, 8, heading, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(2)
        if not body:
            continue
        pdf.set_font(font, "", 11)
        bw = pdf.w - pdf.l_margin - pdf.r_margin
        for line in body.strip().split("\n"):
            if not line.strip():
                continue
            lines = pdf.multi_cell(
                bw, 6, line.strip(),
                dry_run=True, output="LINES", wrapmode=WrapMode.CHAR,
            )
            for ln in lines:
                pdf.cell(bw, 6, ln, new_x="LMARGIN", new_y="NEXT")

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        tmppath = Path(f.name)
    pdf.output(str(tmppath))
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(tmppath.read_bytes())
    return f"OK: {out.resolve()} ({out.stat().st_size} bytes)"


if __name__ == "__main__":
    mcp.run(transport="stdio")
