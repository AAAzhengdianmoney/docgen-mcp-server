# DocGen MCP Server

Generate **Word (.docx)** and **PDF** documents directly from Claude Code and any MCP-compatible client. Cross-platform: Windows, macOS, Linux. Built with FastMCP.

## Features

- **DOCX & PDF** — title, headings, body text, bullet/numbered lists, tables, images, dividers
- **Rich text** — `**bold**` and `*italic*` inline formatting in body text (italic → underline for CJK)
- **Tables** — structured data with `style: "table"` and `{headers, rows}`
- **Images** — embed local images with optional width control
- **Dividers** — visual section separators
- **Headers & footers** — per-document page header/footer text
- **Multi-level headings** — h1/h2/h3 via `level` field
- **Multi-font** — 17+ CJK fonts auto-discovered per platform, per-section font selection
- **`list_fonts`** tool — see what fonts are available on your system
- **Zero config** — CJK fonts auto-detected on all platforms
- **Font aliases** — `"kai"`=楷体, `"song"`=宋体, `"fang"`=仿宋, `"hei"`=黑体, `"li"`=隶书, etc.
- **Lightweight** — only `python-docx` + `fpdf2` + `mcp`

## Install

```bash
pip install -r requirements.txt
# or
pip install mcp python-docx fpdf2
```

## Configure Claude Code

**Option A — User scope** (available in all projects):

```bash
claude mcp add docgen --scope user -- python path/to/server.py
```

**Option B — Project scope** (`.mcp.json`):

```json
{
  "mcpServers": {
    "docgen": {
      "command": "python",
      "args": ["path/to/server.py"]
    }
  }
}
```

Then restart Claude Code.

## Usage

### generate_docx

```
mcp__docgen__generate_docx
  output_path: /path/to/output.docx
  title: "My Document"
  title_font: "hei"
  heading_font: "hei"
  header_text: "Page header"
  footer_text: "Page footer"
  sections: [
    {"heading": "Section 1", "body": "Hello **world**.", "font": "kai"},
    {"heading": "Key Points", "body": "Point A\nPoint B", "style": "bullet"},
    {"heading": "Stats", "style": "table", "table": {"headers": ["Name","Val"], "rows": [["A",1]]}},
    {"style": "divider"},
    {"heading": "Photo", "style": "image", "image": "/path/to/img.png", "image_width": 80}
  ]
```

| Param | Required | Notes |
|-------|----------|-------|
| `output_path` | ✅ | Absolute path, directories auto-created |
| `title` | No | Centered large heading |
| `title_font` | No | Font for the main title (use alias like `"hei"`, `"dengbold"`) |
| `heading_font` | No | Default font for all section headings. Each section can override with its own `heading_font` |
| `header_text` | No | Page header text (centered, top of each page) |
| `footer_text` | No | Page footer text (appended before page number) |
| `sections` | No | Array of section dicts |
| `sections[].heading` | No | Section heading text |
| `sections[].body` | No | Body text. Supports `**bold**` and `*italic*` (italic → underline for CJK) |
| `sections[].style` | No | `"normal"` (default), `"bullet"`, `"numbered"`, `"table"`, `"divider"`, `"image"` |
| `sections[].level` | No | Heading level 1/2/3 (default: 1) |
| `sections[].font` | No | Font for body text. Use `list_fonts()` to see available fonts. Aliases: `"kai"`, `"fang"`, `"hei"`, etc. |
| `sections[].heading_font` | No | Font for this section's heading (overrides global `heading_font`) |
| `sections[].table` | No | `{"headers": [...], "rows": [[...], ...]}` for table style |
| `sections[].table_font` | No | Font for table cells (falls back to `font`) |
| `sections[].image` | No | Path to a local image file for image style |
| `sections[].image_width` | No | Image width in mm (optional) |

### generate_pdf

```
mcp__docgen__generate_pdf
  output_path: /path/to/output.pdf
  title: "My Report"
  title_font: "hei"
  heading_font: "hei"
  header_text: "Confidential"
  footer_text: "Acme Corp"
  sections: [
    {"heading": "Summary", "body": "Some text here.", "font": "kai"},
    {"style": "table", "table": {"headers": ["A","B"], "rows": [["1","2"]]}}
  ]
```

Same parameters as generate_docx. In PDF, bold/italic markers are stripped (CJK fonts lack bold/italic glyph variants).

### list_fonts

```
mcp__docgen__list_fonts
```
Returns all available CJK fonts on the system, grouped by platform, with alias names.

## CJK Font Support

Auto-discovered fonts (Windows example):

| Font | Alias | Style |
|------|-------|-------|
| 微软雅黑 | `yahei` | Modern sans-serif body |
| 黑体 | `hei` | Bold headings (default) |
| 宋体 | `song` | Classic serif |
| 楷体 | `kai` | Calligraphic |
| 仿宋 | `fang` | Semi-cursive |
| 隶书 | `li` | Ancient stone-carving |
| 幼圆 | `you` | Rounded, friendly |
| 等线 | `deng` | Neutral geometric |
| 华文楷体 | `stkaiti` | Alternative calligraphic |
| 华文仿宋 | `stfangsong` | Elegant semi-cursive |
| 华文细黑 | `stxihei` | Light sans-serif |

Plus: 华文宋体, MingLiU, 等线 Bold/Light variants. macOS/Linux fonts auto-detected.

Use `mcp__docgen__list_fonts` to see what's available on your system.

If no CJK font is found, falls back to Helvetica/Arial (ASCII-only).

## Development Notes

### CJK Caveats

- **Italic → Underline**: CJK fonts have no italic glyph files. DOCX renders `*italic*` as underlined text instead.
- **Bold in PDF**: PDF generator strips `**bold**` markers (CJK fonts lack bold variants); they render correctly in DOCX.
- **MCP arg bug**: Long CJK strings in MCP tool arguments trigger a serialization bug (#64506). Use `sections_file` pointing to a disk JSON file for production documents.

## Example Prompt

> Generate a PDF report titled "Q2 Review" with three sections: summary, results, and next steps. Add a table with KPIs and use page headers.

## License

MIT
