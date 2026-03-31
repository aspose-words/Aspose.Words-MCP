# Context
Swarm: default

## Project Context
- Language: Python 3.11+
- Framework: FastMCP-based MCP server
- Build command: python -m compileall mcp_server.py core aliases tests
- Test command: pytest <specific test file>
- Lint command: ruff check . && ruff format --check .
- Entry points: `mcp_server.py`, `core/`, `tests/features/`, `tests/server/`

## Decisions
- Align only to 26.3.0 release-note and API-diff changes that are relevant to MCP server capabilities: prevents unrelated feature expansion.
- Expose exactly two new release-aligned surfaces first: AI summarization and options-based paragraph run merging: these are the only newly added public APIs in scope that the server does not already use.
- Treat removed `TableStyle.bidi` as verification-only unless code evidence appears later: reality check found no current usage.

## SME Cache
### python
- Confidence: HIGH
- Expose a narrow AI summarization capability using direct model construction with explicit configuration and no silent fallback behavior.
- Expose a targeted run-merging capability using the new options-based paragraph API and keep mutations scoped to selected content.
- Mock AI model construction and summarization in tests; do not require live external credentials in CI.

## Patterns
- Tool registration pattern: core logic lives under `core/`; user-facing MCP tool wiring lives in `mcp_server.py`.
- Test pattern: add focused feature tests for each newly exposed capability rather than broad suite changes.
- Release alignment pattern: update README and CHANGELOG when new server-facing capability is exposed.

## Project Governance
- MUST use explicit Aspose.Words API access only; do not use `getattr()` or `hasattr()` for Aspose.Words API access.
- MUST NOT hide Aspose.Words API usage errors with broad defensive exception handling.
- MUST update tests whenever server behavior or exposed tools change.
- SHOULD prefer extending an existing tool when the new capability fits naturally within it.
- SHOULD keep changes grounded in the release notes and API diff for the target release.
- SHOULD avoid unrelated refactors, cosmetic-only edits, and unrelated file modifications.
- SHOULD follow repository Ruff formatting/lint rules.
- SHOULD update `CHANGELOG.md` when required by the task instructions.

## Agent Activity

| Tool | Calls | Success | Failed | Avg Duration |
|------|-------|---------|--------|--------------|
| read | 309 | 309 | 0 | 6ms |
| grep | 100 | 100 | 0 | 15ms |
| test_runner | 96 | 96 | 0 | 650ms |
| task | 74 | 74 | 0 | 70980ms |
| apply_patch | 64 | 64 | 0 | 899ms |
| glob | 53 | 53 | 0 | 16ms |
| bash | 50 | 50 | 0 | 1599ms |
| pre_check_batch | 28 | 28 | 0 | 4009ms |
| lint | 28 | 28 | 0 | 4035ms |
| update_task_status | 17 | 17 | 0 | 4ms |
| declare_scope | 12 | 12 | 0 | 2ms |
| diff | 11 | 11 | 0 | 16ms |
| check_gate_status | 11 | 11 | 0 | 2ms |
| imports | 7 | 7 | 0 | 55ms |
| todo_extract | 6 | 6 | 0 | 2ms |
| save_plan | 2 | 2 | 0 | 11ms |
| symbols | 2 | 2 | 0 | 4ms |
| write_retro | 1 | 1 | 0 | 3ms |
