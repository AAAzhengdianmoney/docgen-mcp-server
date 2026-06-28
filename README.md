# DocGen MCP Server

Generate **Word (.docx)** and **PDF** documents directly from Claude Code and any MCP-compatible client. Cross-platform: Windows, macOS, Linux. Built with FastMCP.

## Features

- **DOCX** — title, headings, body text, bullet/numbered lists
- **PDF** — auto page numbers, CJK (Chinese/Japanese/Korean) font support
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
  title_font: "hei"           # optional, override title font
  heading_font: "hei"         # optional, default font for all headings
  sections: [
    {"heading": "Section 1", "body": "Hello world.", "style": "normal", "font": "kai"},
    {"heading": "Key Points", "body": "Point A\nPoint B", "style": "bullet"}
  ]
```

| Param | Required | Notes |
|-------|----------|-------|
| `output_path` | ✅ | Absolute path, directories auto-created |
| `title` | No | Centered large heading |
| `title_font` | No | Font for the main title (use alias like `"hei"`, `"dengbold"`) |
| `heading_font` | No | Default font for all section headings. Each section can override with its own `heading_font` |
| `sections` | No | Array of `{heading, body, style, font, heading_font}` |
| `sections[].style` | No | `"normal"` (default), `"bullet"`, `"numbered"` |
| `sections[].font` | No | Font for this section's body text. Use `list_fonts()` to see available fonts. Aliases: `"kai"`, `"fang"`, `"hei"`, `"song"`, `"li"`, `"you"`, `"deng"`, `"yahei"`, etc. |
| `sections[].heading_font` | No | Font for this section's heading (overrides global `heading_font`) |

### generate_pdf

```
mcp__docgen__generate_pdf
  output_path: /path/to/output.pdf
  title: "My Report"
  title_font: "hei"
  heading_font: "hei"
  sections: [
    {"heading": "Summary", "body": "Some text here.", "font": "kai"}
  ]
```

Same font parameters as generate_docx.

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

## Example Prompt

> Generate a PDF report titled "Q2 Review" with three sections: summary, results, and next steps.

## License

MIT
