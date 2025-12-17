# Aspose.Words MCP Server

![Python Version](https://img.shields.io/badge/python-3.11%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Overview
**Aspose.Words MCP Server** is a FastMCP-based MCP server built on top of [Aspose.Words for Python via .NET](https://products.aspose.com/words/python-net/). It automates Microsoft Word document creation and editing and exposes operations as MCP tools that any MCP-compatible client can call. Supported transports: `stdio`, `streamable-http`, `sse`.

## Features

- Create documents; read/write text, headings, and paragraphs
- Text formatting (font, size, style, color, underline)
- Page and section breaks; page setup (margins, orientation, paper size, columns)
- Lists (bulleted/numbered)
- Tables: create, auto-fit, merge cells, alignment, shading, paddings, column widths, header styling
- Footnotes and endnotes: add, convert, anchor-based operations, validation
- Comments: get by author, by paragraph, all comments
- Document properties: read/write (title, author, subject, keywords)
- Protection: protect/unprotect, partial editing restrictions
- Bookmarks, hyperlinks
- Watermarks (text/image)
- Export as Base64 (DOCX, PDF, etc.), advanced export options
- Render page to image (PNG, etc.)
- In-memory document management: copy, save as, list, delete, merge

## Requirements

- Python 3.11+
- [Aspose.Words for Python via .Net](https://products.aspose.com/words/python-net/). This library is a [commercial product](https://purchase.aspose.com/buy/words/python).  
You'll need to obtain a valid license for Aspose.Words. The package will install this dependency, but you're responsible for complying with Aspose's licensing terms.

## Installation

```bash
pip install aspose-words-mcp
```

From source (download repo and install requirements):

```bash
git clone https://github.com/aspose-words/Aspose.Words-MCP-Server
cd Aspose.Words-MCP-Server
pip install -r requirements.txt
```

## Command Line Interface

After installation, the CLI command is available:

```bash
aspose-words-mcp
```

By default, the server runs with the `stdio` transport.

Run without installation:

```bash
python mcp_server.py
```

## Transports and Configuration

Supported MCP transports: `stdio`, `streamable-http`, `sse`.

### Environment Variables

- `MCP_TRANSPORT` — `stdio` | `streamable-http` | `sse` (default `stdio`)
- `MCP_HOST` — host address (default `0.0.0.0`)
- `MCP_PORT` — port (default `8080`)
- `MCP_PATH` — HTTP path for `streamable-http` (default `/mcp`)
- `MCP_SSE_PATH` — events path for `sse` (default `/sse`)
- `LOG_LEVEL` — logging level (`INFO`, `DEBUG`, ...)

### Aspose.Words License

The Aspose.Words license is applied when the server starts. The effective path to the license file is resolved with the following precedence:

1) The `license_path` argument of `run_server(..., license_path=None)`
2) The `ASPOSE_WORDS_LICENSE_PATH` environment variable

If no license is provided or the file is not accessible, the server runs in Evaluation mode.

Example of setting the environment variable:

```bash
export ASPOSE_WORDS_LICENSE_PATH='/path/to/aspose.words.lic'
```

### HTTP/SSE Run Example

```bash
export MCP_TRANSPORT=streamable-http   # or sse
export MCP_HOST=0.0.0.0
export MCP_PORT=8080
export MCP_PATH=/mcp                   # for streamable-http
export MCP_SSE_PATH=/sse               # for sse
aspose-words-mcp
```

On start, the server prints the listening address.

## Tools

See full list and signatures in `mcp_server.py` (function `register_tools`) and tests in `tests/features/*`.

Main tool categories:

- content/reading: create document, insert/delete/read text, headings, lists, HTML/Markdown
- layout: pages, breaks, columns, headers/footers, page numbering
- tables: create and format tables
- watermarks: watermarks
- links/bookmarks: hyperlinks and bookmarks
- properties: document properties
- protection: protection and restrictions
- comments/notes: comments, footnotes/endnotes
- export/render: export, page rendering

## Example Workflow via an MCP Client

Sequence of tool calls (names match the server):

1. `create_document` → get `doc_id`
2. `add_heading` (e.g., levels 1–3)
3. `add_paragraph` / `insert_text_end`
4. `add_table_end` or `add_table_at_paragraph`
5. `add_watermark_text` or `add_watermark_image_base64`
6. `export_base64` (e.g., `fmt="pdf"`) — get file as Base64

## Integration with MCP Clients

- Claude Desktop MCP: add this server with `streamable-http` or `sse` transport and the URL printed by the server at startup.
- Any MCP (JSON) clients — configure the matching transport and path.

## License

This package is licensed under the MIT License. However, it depends on Aspose.Words for Python via .Net library, which is proprietary, closed-source library.

⚠️ You must obtain valid license for Aspose.Words for Python via .Net library. This repository does not include or distribute any proprietary components.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Use of third-party trademarks or logos is subject to those third-party policies.