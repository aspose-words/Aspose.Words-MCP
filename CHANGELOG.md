# Changelog

All notable changes to this project will be documented in this file.

The format is based on "Keep a Changelog", and versioning adheres to Semantic Versioning (SemVer).

## [26.4.0] - 2026-04-18
### Added
- Exposed tagged-PDF custom-node-ID workflow through existing MCP surfaces.

### Changed
- Aligned server versioning with Aspose.Words 26.4.0.

### Skipped
- Funnel-chart data labels were not exposed because this MCP server has no chart authoring/editing workflow to extend.
- Chart leader-line style rendering was not exposed because this MCP server has no chart authoring/editing workflow to extend.
- Locale-aware default axis titles were not exposed because this MCP server has no chart authoring/editing workflow to extend.

## [26.3.0] - 2026-04-04
### Added
- Exposed opt-in run joining through the existing text replacement MCP surface.

### Changed
- Aligned server versioning with Aspose.Words 26.3.0.

### Skipped
- `OpenAiModel(name)` / `OpenAiModel(name, api_key)` were not exposed because this MCP server has no AI tool surface, and exposing them would require new secret/configuration infrastructure.
- `TableStyle.bidi` removal has no server-visible impact because this API is not used in this repository and is not exposed through MCP tools.

## [26.2.0] - 2026-04-02
### Added
- Exposed text-shaping control for advanced PDF export through the existing export MCP surface.

### Changed
- Aligned server versioning with Aspose.Words 26.2.0.

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
