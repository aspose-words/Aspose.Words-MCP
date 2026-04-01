# Changelog

All notable changes to this project will be documented in this file.

The format is based on "Keep a Changelog", and versioning adheres to Semantic Versioning (SemVer).

## [Unreleased]
### Added
- `replace_regex_to_images_base64` for regex-driven image exports that return an `images` list with `base64`, `mime`, and `ext` fields.
### Changed
- Aligned the MCP server package metadata and Aspose.Words dependency to 25.12.0.
- Extended `replace_text` with `use_regex` and `whole_word` while preserving the existing `search_text`, `replacement_text`, `replace_all`, and `case_sensitive` arguments.
- Routed regex replacement flows through the Aspose.Words 25.12.0 low-code regex replacement API.
### Fixed
- Made regex replacement errors explicit instead of silently degrading when invalid patterns are supplied.

## [25.11.0] - 2025-12-02
### Added
- Initial version of the MCP server for Aspose.Words automation (based on FastMCP).
- Integration with Aspose.Words 25.11.0.
- Core capabilities: content creation and editing, tables, styles, comments and footnotes, watermarks, links, layout/formatting, export, and file read/write.
- Server startup via the `aspose-words-mcp` script.
- Configuration via environment variables (`.env`).
