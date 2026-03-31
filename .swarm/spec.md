# Aspose.Words MCP Server 26.3.0 Release Alignment

## Feature Description
Align the MCP server with the Aspose.Words 26.3.0 release so the server exposes newly added public API that is relevant to document workflows, removes any dependence on removed API, and keeps release-facing behavior, tests, and documentation consistent with the target version.

## User Scenarios

### Scenario 1: Use newly added AI-backed summarization capability
**Given** a user has a document and valid credentials for the supported AI-backed summarization flow
**When** the user invokes the MCP server capability for document summarization
**Then** the server must summarize the document using the newly supported direct model-construction API and return or save the resulting document output with clear error reporting on invalid inputs or external service failures.

### Scenario 2: Normalize fragmented paragraph runs
**Given** a document contains adjacent text runs that should be merged when formatting-equivalent content is normalized
**When** the user invokes the MCP server capability for joining runs in a paragraph with selected merge options
**Then** the server must merge eligible runs using the new options-based API and report the resulting document changes without modifying unrelated content.

### Scenario 3: Trust release-aligned server behavior
**Given** the MCP server targets Aspose.Words 26.3.0
**When** the server is built, tested, and documented
**Then** the exposed capabilities, dependency versioning, and release notes in the repository must match the 26.3.0 public API changes relevant to this server.

## Functional Requirements
- **FR-001** The server MUST expose at least one user-facing capability that exercises the newly supported direct model-construction API for AI-backed document summarization.
- **FR-002** The AI-backed summarization capability MUST require explicit user-provided configuration for the external model invocation and MUST surface failures clearly instead of silently degrading behavior.
- **FR-003** The server MUST expose a user-facing capability that exercises the new options-based paragraph run-joining API.
- **FR-004** The run-joining capability MUST allow users to control the newly added merge-option behaviors that affect whether redundant, insignificant, or spacing-related differences are ignored.
- **FR-005** The run-joining capability MUST scope its document mutations to the user-selected target content and MUST report the outcome of the operation.
- **FR-006** The repository MUST not reference removed public API that is no longer available in the target release.
- **FR-007** The repository MUST update automated tests to cover any changed or newly exposed server behavior introduced for the 26.3.0 alignment.
- **FR-008** The repository MUST update release-facing documentation to describe the newly exposed 26.3.0-aligned capabilities.

## Success Criteria
- **SC-001** A user can invoke an MCP workflow that uses the direct AI model-construction release addition without requiring unsupported fallback behavior.
- **SC-002** A user can invoke an MCP workflow that uses the new options-based paragraph run-joining release addition and observe deterministic merge behavior.
- **SC-003** Automated tests pass for all modified server behaviors introduced for the 26.3.0 alignment.
- **SC-004** Repository documentation and release history reflect the 26.3.0 alignment and the newly exposed capabilities.
- **SC-005** No repository code or tests rely on removed API from the prior release surface.

## Key Entities
- Document
- Paragraph
- Text run
- Summarization request
- Summarization result
- Merge options

## Edge Cases and Failure Modes
- Invalid or missing external service credentials
- Unsupported or empty model configuration
- Empty document or target content that cannot be summarized meaningfully
- Paragraph selection that is out of range or contains no mergeable runs
- Merge-option combinations that produce no document changes
- External service or library errors during summarization
- Existing repository surfaces that appear release-aligned by version but do not yet expose the new capabilities
