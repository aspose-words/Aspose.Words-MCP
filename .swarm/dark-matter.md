## Dark Matter: Hidden Couplings

Found 20 file pairs that frequently co-change but have no import relationship:

| File A | File B | NPMI | Co-Changes | Lift |
|--------|--------|------|------------|------|
| core/utils/license.py | tests/features/test_content_creation.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/comments.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/content.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/export.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/io.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/layout.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/links.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/notes.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/properties.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/protection.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/reading.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/store.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/styles.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/tables.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/utils/docs_util.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | core/watermarks.py | 1.000 | 3 | 6.67 |
| aliases/filename_alias_tools.py | tests/unit/test_store.py | 1.000 | 3 | 6.67 |
| core/comments.py | core/content.py | 1.000 | 3 | 6.67 |
| core/comments.py | core/export.py | 1.000 | 3 | 6.67 |
| core/comments.py | core/io.py | 1.000 | 3 | 6.67 |

These pairs likely share an architectural concern invisible to static analysis.
Consider adding explicit documentation or extracting the shared concern.