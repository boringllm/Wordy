"""
Word Template MCP Server
========================

An MCP server that lets an agent produce Word documents (.docx) that strictly
follow an arbitrary Word template (cover page, styles, fonts, colors, tables,
headers/footers, page setup), from a markdown report.

Intended agent workflow (3 tools):

  1. inspect_template   -> learn everything about the template: available
                           styles (with fonts/sizes/colors), theme colors,
                           sections, body structure (numbered blocks), cover
                           page boundary, {{placeholders}}, and a suggested
                           markdown->style mapping.
  2. create_document    -> render the markdown into a copy of the template,
                           keeping the first N body blocks (cover page),
                           filling placeholders, and mapping every markdown
                           element onto the template's own named styles.
  3. validate_document  -> verify the produced file only uses styles/fonts
                           defined by the template and has no leftover
                           placeholders.

Transports:
  stdio (default):   python server.py
  streamable HTTP:   python server.py --transport http [--host H] [--port P]
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import Any, Optional

from mcp.server.fastmcp import FastMCP

import docx_engine

mcp = FastMCP(
    "word-template",
    instructions=(
        "Produces .docx documents that strictly follow a Word template. "
        "Always call inspect_template first to discover the template's "
        "styles, cover-page boundary (cover_page_hint.keep_blocks), and "
        "placeholders; then call create_document; then validate_document."
    ),
)


@mcp.tool()
def inspect_template(template_path: str) -> dict:
    """Analyze a Word template (.docx) and return a complete, machine-readable
    description of it. ALWAYS call this before create_document.

    Args:
        template_path: Absolute path to the .docx template file.

    Returns a JSON object with:
      - styles: every style defined in the template, grouped by type
        (paragraph / character / table), each with name, style_id, builtin
        flag, base_style, and for text styles the effective font (name,
        size_pt, bold, italic, color as #RRGGBB; null = inherited).
      - theme: theme colors (dk1, lt1, accent1..6, hlink, ...) and the
        theme's major/minor fonts.
      - defaults: font of the 'Normal' style (document default text).
      - sections: page size (cm), orientation, margins, and the text found
        in headers/footers (including first-page variants).
      - body_blocks: ordered list of every top-level block in the template
        body (paragraphs and tables) with index, style name, a text preview,
        and flags contains_page_break / contains_section_break. Use these
        indexes to decide the keep_blocks value for create_document.
      - cover_page_hint: {detected, keep_blocks, reason} — a suggestion for
        how many leading blocks constitute the cover page (everything up to
        and including the first page/section break).
      - placeholders: every {{name}} token found in the body, headers and
        footers. Fill these via the placeholders argument of create_document.
      - suggested_style_map: recommended mapping of markdown elements
        (h1..h6, paragraph, bullet_list, ordered_list, table, quote,
        code_block, caption, hyperlink_char_style) to this template's actual
        style names. Pass it (possibly adjusted) as style_map to
        create_document.
    """
    return docx_engine.inspect_template(template_path)


@mcp.tool()
def create_document(
    template_path: str,
    output_path: str,
    markdown_path: Optional[str] = None,
    markdown_text: Optional[str] = None,
    keep_blocks: int = 0,
    page_break_before_content: Optional[bool] = None,
    placeholders: Optional[dict] = None,
    style_map: Optional[dict] = None,
) -> dict:
    """Create a new .docx from a Word template and a markdown report. The
    output is a byte-level copy of the template (all styles, theme, fonts,
    colors, numbering, headers/footers and page setup preserved) in which the
    markdown content is rendered using the template's own named styles.

    Args:
        template_path: Absolute path to the .docx template.
        output_path: Absolute path where the new .docx will be written
            (directories are created if needed; overwrites existing file).
        markdown_path: Path to the markdown report file (UTF-8). Provide
            exactly one of markdown_path or markdown_text.
        markdown_text: The markdown report content as a string. Provide
            exactly one of markdown_path or markdown_text.
        keep_blocks: How many leading body blocks of the template to keep in
            the output (use body_blocks / cover_page_hint from
            inspect_template to choose):
              0  = discard the whole template body, keep only styles, theme,
                   headers/footers and page setup (default);
              N  = keep blocks 0..N-1, e.g. the cover page and/or a TOC;
              -1 = keep the entire template body and append the report after.
        page_break_before_content: Insert a page break between the kept
            template blocks and the rendered markdown. Default (null) = auto:
            inserted only if blocks are kept and the last kept block does not
            already end with a page/section break.
        placeholders: Mapping used to fill {{name}} tokens, e.g.
            {"title": "Q3 Report", "date": "2026-07-09"}. Replaced in kept
            body blocks AND in all headers/footers, preserving the runs'
            original formatting. Unknown tokens are left untouched (and will
            be flagged by validate_document).
        style_map: Optional overrides mapping markdown elements to template
            style NAMES (must exist in the template, otherwise a warning is
            returned and defaults are used). Recognized keys: "h1".."h6",
            "paragraph", "bullet_list", "ordered_list", "table" (a TABLE
            style), "quote", "code_block", "caption",
            "hyperlink_char_style" (a CHARACTER style). Start from
            inspect_template's suggested_style_map.

    Markdown support: headings 1-6, paragraphs, bold/italic/strikethrough,
    inline code, fenced code blocks, ordered/unordered lists (nested; deeper
    levels use 'List Bullet 2/3'-type styles when the template defines them),
    GFM tables (rendered with the template's table style, header row bold),
    blockquotes, hyperlinks (real Word hyperlinks using the template's
    Hyperlink character style), horizontal rules, and images
    (![alt](path), resolved relative to the markdown file, auto-scaled to
    the text column width).

    Returns: {output_path, kept_template_blocks, placeholders_replaced,
    elements_rendered (counts), template_styles_used, warnings}. Inspect
    `warnings` — it lists any style fallbacks or missing images.
    """
    if (markdown_path is None) == (markdown_text is None):
        raise ValueError(
            "Provide exactly one of markdown_path or markdown_text.")
    base_dir = None
    if markdown_path is not None:
        if not os.path.isfile(markdown_path):
            raise FileNotFoundError("Markdown file not found: %s" % markdown_path)
        with open(markdown_path, "r", encoding="utf-8") as f:
            markdown_text = f.read()
        base_dir = os.path.dirname(os.path.abspath(markdown_path))
    return docx_engine.create_document(
        template_path=template_path,
        output_path=output_path,
        markdown_text=markdown_text,
        keep_blocks=keep_blocks,
        page_break_before_content=page_break_before_content,
        placeholders=placeholders,
        style_map=style_map,
        markdown_base_dir=base_dir,
    )


@mcp.tool()
def validate_document(document_path: str, template_path: str) -> dict:
    """Check that a generated .docx strictly conforms to a Word template.
    Call this after create_document to confirm template fidelity.

    Args:
        document_path: Absolute path to the .docx to validate.
        template_path: Absolute path to the reference .docx template.

    Checks performed:
      - every paragraph and character style used in the document (including
        inside tables) is defined in the template;
      - no unreplaced {{placeholder}} tokens remain;
      - no fonts foreign to the template (its styles + theme fonts) are
        applied to runs.

    Returns: {conforms: bool, issues: [human-readable problems],
    stats: {paragraphs, tables, paragraph_styles_used,
    character_styles_used}}. If conforms is false, fix by re-running
    create_document with a corrected style_map or placeholders.
    """
    return docx_engine.validate_document(document_path, template_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Word Template MCP server")
    parser.add_argument("--transport", choices=["stdio", "http"],
                        default="stdio",
                        help="stdio (default) or http (streamable HTTP)")
    parser.add_argument("--host", default="127.0.0.1",
                        help="HTTP host (default 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8848,
                        help="HTTP port (default 8848)")
    args = parser.parse_args()

    if args.transport == "http":
        mcp.settings.host = args.host
        mcp.settings.port = args.port
        print("Word Template MCP server on http://%s:%d/mcp"
              % (args.host, args.port), file=sys.stderr)
        mcp.run(transport="streamable-http")
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
