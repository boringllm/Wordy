# Word Template MCP Server

An MCP server that produces Word documents (.docx) which **strictly follow an arbitrary Word template** — cover page, styles, fonts (police), colors, tables, headers/footers, page setup — from a **markdown report**.

The key design decision: the output document is *the template file itself*, opened, selectively cleared, and refilled. Every style, theme color, font, numbering definition, header, footer and page setup is preserved byte-for-byte, whatever form the template takes. Markdown content is then mapped onto the template's **own named styles**, never onto hard-coded formatting.

## Tools

### 1. `inspect_template(template_path)`
Call this **first**. Returns a complete machine-readable description of the template:

- every paragraph / character / table **style** with its font name, size, color (#RRGGBB), bold/italic
- **theme** colors (accent1..6, hyperlink, ...) and major/minor theme fonts
- **sections**: page size, orientation, margins, header/footer text
- **body_blocks**: every top-level block (paragraph/table) with index, style, text preview, and page/section-break flags — used to decide what to keep
- **cover_page_hint**: suggested `keep_blocks` value (everything up to the first page/section break)
- **placeholders**: all `{{name}}` tokens found in body, headers and footers
- **suggested_style_map**: recommended markdown→style mapping for this template

### 2. `create_document(template_path, output_path, markdown_path | markdown_text, keep_blocks, page_break_before_content, placeholders, style_map)`
Renders the markdown into a copy of the template and writes it to `output_path`.

- `keep_blocks`: `0` = drop template body (keep styles/headers/setup only), `N` = keep the first N blocks (cover page, TOC...), `-1` = keep everything and append.
- `placeholders`: fills `{{name}}` tokens in kept blocks **and** headers/footers, preserving run formatting even when a token is split across runs.
- `style_map`: override which template style each markdown element uses (`h1`..`h6`, `paragraph`, `bullet_list`, `ordered_list`, `table`, `quote`, `code_block`, `caption`, `hyperlink_char_style`). Defaults come from the template itself (e.g. generated tables reuse the style of the template's own tables).
- Markdown supported: headings 1–6, bold/italic/strikethrough, inline code, fenced code blocks, nested lists (uses `List Bullet 2/3`-type styles when defined), GFM tables, blockquotes, real Word hyperlinks, horizontal rules, local images (auto-scaled to text width).
- Returns counts, the exact template styles used, and `warnings` listing any fallback.

### 3. `validate_document(document_path, template_path)`
Verifies the output conforms: every style used exists in the template, no foreign fonts, no leftover `{{placeholders}}`. Returns `{conforms, issues, stats}`.

## Agent workflow

```
inspect_template  →  choose keep_blocks (cover_page_hint), placeholders, style_map
create_document   →  check warnings
validate_document →  confirm conforms: true
```

## Installation

```
pip install -r requirements.txt
```

## Running

**stdio** (default — for Claude Desktop / Claude Code and most MCP clients):

```
python server.py
```

**Streamable HTTP**:

```
python server.py --transport http --host 127.0.0.1 --port 8848
# endpoint: http://127.0.0.1:8848/mcp
```

## Client configuration

Claude Desktop / generic stdio client (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "word-template": {
      "command": "python",
      "args": ["C:\\temp\\LLM\\Wordy\\server.py"]
    }
  }
}
```

Claude Code:

```
claude mcp add word-template -- python C:\temp\LLM\Wordy\server.py
```

HTTP client:

```json
{
  "mcpServers": {
    "word-template": {
      "type": "http",
      "url": "http://127.0.0.1:8848/mcp"
    }
  }
}
```

## Files

- `server.py` — MCP server (FastMCP), tool contracts, stdio + streamable HTTP transports
- `docx_engine.py` — template inspection, markdown→docx rendering, validation
- `examples/make_sample_template.py` — generates a demo corporate template (cover page, styled headings, header/footer placeholders, styled table)
- `examples/sample_template.docx` — the demo template
- `examples/sample_report.md` — a demo markdown report
- `examples/sample_output.docx` — result of rendering the report through the template

## Notes & limitations

- Placeholder syntax is `{{name}}` (spaces tolerated: `{{ name }}`).
- Ordered lists rely on the template's `List Number` style; numbering continuation across separate lists follows the template's numbering definition.
- Remote (http) images are not downloaded; they are inserted as italic `[image: alt]` markers and reported in `warnings`.
- Templates with localized style names work: styles are matched by name **and** style id, and `style_map` lets the agent target any exact style surfaced by `inspect_template`.
