from __future__ import annotations

import datetime as _dt

import aspose.words as aw

from core.utils.docs_util import ensure_path


def _resolve_compare_datetime(date_time_iso: str | None) -> _dt.datetime:
    if date_time_iso is None:
        return _dt.datetime.now()

    normalized_value = str(date_time_iso).strip()
    if not normalized_value:
        raise ValueError('date_time_iso must be a non-empty ISO 8601 datetime string')

    if normalized_value.endswith('Z'):
        normalized_value = f'{normalized_value[:-1]}+00:00'

    try:
        return _dt.datetime.fromisoformat(normalized_value)
    except ValueError as exc:
        raise ValueError('date_time_iso must be a valid ISO 8601 datetime string') from exc


def compare_documents(
    doc_id: str,
    compare_to_doc_id: str,
    author: str = 'MCP',
    date_time_iso: str | None = None,
    compare_list_definitions: bool = False,
) -> bool:
    if author is None or not str(author).strip():
        raise ValueError('author must be a non-empty string')
    if not isinstance(compare_list_definitions, bool):
        raise ValueError('compare_list_definitions must be a boolean')

    resolved_datetime = _resolve_compare_datetime(date_time_iso)
    source_path = ensure_path(str(doc_id))
    target_path = ensure_path(str(compare_to_doc_id))

    source_doc = aw.Document(str(source_path))
    target_doc = aw.Document(str(target_path))
    options = aw.comparing.CompareOptions()
    options.advanced_options.compare_list_definitions = compare_list_definitions

    source_doc.compare(
        document=target_doc,
        author=str(author),
        date_time=resolved_datetime,
        options=options,
    )
    source_doc.save(str(source_path))
    return True
