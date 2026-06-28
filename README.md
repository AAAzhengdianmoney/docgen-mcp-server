# DocGen MCP Server

Generate **Word (.docx)** and **PDF** documents directly from Claude Code and any MCP-compatible client. Cross-platform: Windows, macOS, Linux. Built with FastMCP.

## Features

- **DOCX** — title, headings, body text, bullet/numbered lists
- **PDF** — auto page numbers, CJK (Chinese/Japanese/Korean) font support
- **Zero config** — CJK fonts auto-detected on all platforms
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
  sections: [
    {"heading": "Section 1", "body": "Hello world.", "style": "normal"},
    {"heading": "Key Points", "body": "Point A\nPoint B", "style": "bullet"}
  ]
```

| Param | Required | Notes |
|-------|----------|-------|
| `output_path` | ✅ | Absolute path, directories auto-created |
| `title` | No | Centered large heading |
| `sections` | No | Array of `{heading, body, style}` |
| `sections[].style` | No | `"normal"` (default), `"bullet"`, `"numbered"` |

### generate_pdf

```
mcp__docgen__generate_pdf
  output_path: /path/to/output.pdf
  title: "My Report"
  sections: [
    {"heading": "Summary", "body": "Some text here."}
  ]
```

## CJK Font Support

Tested with:
- **Windows**: Microsoft YaHei (微软雅黑), SimSun, SimHei
- **macOS**: PingFang, STHeiti, Hiragino Sans GB
- **Linux**: WQY ZenHei/MicroHei, Noto Sans CJK

If no CJK font is found, falls back to Helvetica (ASCII-only).

## Example Prompt

> Generate a PDF report titled "Q2 Review" with three sections: summary, results, and next steps.

## License

MIT
