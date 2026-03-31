# Changelog

All notable changes to this project will be documented in this file.

The format is based on "Keep a Changelog", and versioning adheres to Semantic Versioning (SemVer).

## [Unreleased]

### Updated
- Integration with Aspose.Words 26.3.0.
- Exposed new MCP capabilities for AI summarization and paragraph run joining as part of the 26.3.0 release alignment.
- Removed `TableStyle.bidi` required no repository changes because the MCP server did not use that API.
- Updated `qodana.yaml` for native run compatibility by removing the deprecated `ide` field.

## [0.1.8] - 2026-03-19
### Updated
- Integration with Aspose.Words 26.2.0.
- No MCP tool behavior changes were required for this release; update aligns dependency with upstream API.

## [0.1.7] - 2026-03-19
### Updated
- Integration with Aspose.Words 26.1.0.

## [0.1.0] - 2025-12-02
### Added
- Initial version of the MCP server for Aspose.Words automation (based on FastMCP).
- Integration with Aspose.Words 25.11.0.
- Core capabilities: content creation and editing, tables, styles, comments and footnotes, watermarks, links, layout/formatting, export, and file read/write.
- Server startup via the `aspose-words-mcp` script.
- Configuration via environment variables (`.env`).
