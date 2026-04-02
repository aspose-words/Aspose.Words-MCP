# AGENTS.md

## Purpose

This repository is updated by coding agents during Aspose.Words release alignment work.  
When making changes, preserve the existing repository structure, conventions, and architectural patterns.

## Required implementation rules

### 1. Use explicit API only

Do not use `getattr()` or `hasattr()` for Aspose.Words API access.

All Aspose.Words API used by the MCP server must be referenced explicitly and directly.  
If the server is being updated for a target Aspose.Words release, assume that the target API is the intended and supported API surface.

Do not add dynamic attribute checks to support old or uncertain API shapes.  
Do not add compatibility shims based on `getattr()`, `hasattr()`, or similar reflective fallbacks.

### 2. Do not hide library usage errors

Do not use `try`/`except` to hide, suppress, or silently bypass incorrect Aspose.Words API usage.

If the MCP server uses the library incorrectly, that failure must remain visible during development and testing.  
Do not catch exceptions only to continue execution with fallback behavior, partial behavior, or silent degradation.

Allowed exception handling is narrow and explicit only when it is required for:
- test assertions,
- user-input validation,
- external I/O boundaries,
- conversion of an error into a clearer MCP-facing error when the underlying failure remains explicit and testable.

Do not wrap Aspose.Words API calls in broad `try`/`except` blocks for defensive compatibility.

### 3. Tests must be updated with server changes

Any change to MCP server code must be accompanied by the necessary test updates.

If server behavior changes, tool coverage changes, tool parameters change, schemas change, manifests change, descriptions change, or new functionality is exposed, update existing tests and add new tests where required.

Do not leave server changes without corresponding test coverage adjustments.

### 4. Server versioning

The MCP server version must match the version of the Aspose.Words library it is aligned with.
When updating the server for a new Aspose.Words release, ensure that the version in `pyproject.toml` (and any other version-tracking files) is updated to match the target library version.

## Additional coding expectations

- Prefer extending an existing tool when the new capability fits naturally within it.
- Add a new tool only when the functionality is clearly distinct and improves the MCP surface.
- Keep all changes directly grounded in the release notes and API diff for the target release.
- Remove temporary files and folders created during the task before finishing the work.
- Do not leave behind ad hoc artifacts such as extra `__tests__` directories, scratch scripts, repro files, or one-off debugging helpers unless the task explicitly requires them to remain in the repository.
- Do not make unrelated refactors.
- Do not make cosmetic-only edits.
- Do not modify unrelated files.
- Preserve naming conventions, project layout, and code style already used in the repository.
- Follow the repository Ruff configuration in pyproject.toml for formatting and linting expectations.
- Update `CHANGELOG.md` when required by the task instructions.
- README.md updates should be concise and integrated into existing sections. Avoid adding separate version-specific sections (e.g., `### 25.12.0 regex workflows`) or lengthy examples (e.g., `#### Regex replacement example`).
