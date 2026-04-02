# Changelog

All notable changes to this project will be documented in this file.

The format is based on "Keep a Changelog", and versioning adheres to Semantic Versioning (SemVer).

## [26.1.0] - 2026-04-02
### Added
- Exposed Docling export through the advanced export MCP surface.
- Added merge control for preserving only the first imported section.

### Changed
- Aligned server versioning with Aspose.Words 26.1.0.

## [25.12.0] - 2026-04-01
### Added
- Exposed regex-aware replacement through the existing text replacement MCP surface.

### Changed
- Aligned server versioning with Aspose.Words 25.12.0.

## [25.11.0] - 2025-12-02
### Added
- Initial version of the MCP server for Aspose.Words automation (based on FastMCP).
- Integration with Aspose.Words 25.11.0.
- Core capabilities: content creation and editing, tables, styles, comments and footnotes, watermarks, links, layout/formatting, export, and file read/write.
- Server startup via the `aspose-words-mcp` script.
- Configuration via environment variables (`.env`).
