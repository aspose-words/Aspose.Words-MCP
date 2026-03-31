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
| read | 428 | 428 | 0 | 6ms |
| test_runner | 158 | 158 | 0 | 407ms |
| grep | 133 | 133 | 0 | 16ms |
| task | 111 | 111 | 0 | 61811ms |
| glob | 78 | 78 | 0 | 17ms |
| bash | 76 | 76 | 0 | 1378ms |
| apply_patch | 72 | 72 | 0 | 803ms |
| lint | 37 | 37 | 0 | 3169ms |
| pre_check_batch | 35 | 35 | 0 | 3325ms |
| update_task_status | 27 | 27 | 0 | 4ms |
| diff | 17 | 17 | 0 | 29ms |
| declare_scope | 17 | 17 | 0 | 2ms |
| check_gate_status | 16 | 16 | 0 | 2ms |
| imports | 13 | 13 | 0 | 54ms |
| todo_extract | 10 | 10 | 0 | 2ms |
| save_plan | 7 | 7 | 0 | 10ms |
| knowledgeRecall | 6 | 6 | 0 | 3ms |
| symbols | 2 | 2 | 0 | 4ms |
| write_retro | 1 | 1 | 0 | 3ms |
| doc_scan | 1 | 1 | 0 | 2ms |
| detect_domains | 1 | 1 | 0 | 2ms |
