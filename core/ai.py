from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

import aspose.words as aw

from core.utils.docs_util import ensure_path

_SUMMARY_LENGTH_MAP: Dict[str, aw.ai.SummaryLength] = {
    'very_short': aw.ai.SummaryLength.VERY_SHORT,
    'short': aw.ai.SummaryLength.SHORT,
    'medium': aw.ai.SummaryLength.MEDIUM,
    'long': aw.ai.SummaryLength.LONG,
    'very_long': aw.ai.SummaryLength.VERY_LONG,
}


def _require_non_empty(value: Optional[str], field_name: str) -> str:
    normalized = (value or '').strip()
    if not normalized:
        raise ValueError(f'{field_name} is required')
    return normalized


def _require_safe_doc_id(doc_id: str) -> str:
    doc_id_value = _require_non_empty(doc_id, 'doc_id')
    if doc_id_value in ('.', './', '.\\'):
        raise ValueError('doc_id must be a simple identifier, not a path')
    doc_id_path = Path(doc_id_value)
    if doc_id_path.is_absolute() or any(part in ('.', '..') for part in doc_id_path.parts):
        raise ValueError('doc_id must not contain path traversal segments')
    if any(sep in doc_id_value for sep in ('/', '\\')):
        raise ValueError('doc_id must be a simple identifier, not a path')
    return doc_id_value


def _resolve_summary_length(summary_length: Optional[str]) -> Tuple[str, aw.ai.SummaryLength]:
    normalized = (summary_length or '').strip().lower().replace('-', '_') or 'short'
    if normalized not in _SUMMARY_LENGTH_MAP:
        supported = ', '.join(_SUMMARY_LENGTH_MAP.keys())
        raise ValueError(
            f'Unsupported summary_length: {summary_length}. Supported values: {supported}'
        )
    return normalized, _SUMMARY_LENGTH_MAP[normalized]


def _resolve_output_path(
    source_path: Path, summary_length: str, output_name: Optional[str]
) -> Path:
    output_base_name = (output_name or '').strip()
    if output_base_name:
        safe_name = Path(output_base_name).name
        if not safe_name.lower().endswith('.docx'):
            safe_name = f'{safe_name}.docx'
    else:
        safe_name = f'{source_path.stem}.summary.{summary_length}.docx'
    return source_path.parent / safe_name


def summarize_document(
    doc_id: str,
    model_name: str,
    api_key: str,
    summary_length: Optional[str] = 'short',
    output_name: Optional[str] = None,
    model_factory: Callable[[str, str], aw.ai.OpenAiModel] = aw.ai.OpenAiModel,
    summarize_call: Optional[
        Callable[[aw.ai.OpenAiModel, aw.Document, aw.ai.SummarizeOptions], aw.Document]
    ] = None,
) -> Dict[str, str]:
    doc_id_value = _require_safe_doc_id(doc_id)
    model_name_value = _require_non_empty(model_name, 'model_name')
    api_key_value = _require_non_empty(api_key, 'api_key')

    normalized_summary_length, summary_length_enum = _resolve_summary_length(summary_length)

    source_path = ensure_path(doc_id_value)
    output_path = _resolve_output_path(source_path, normalized_summary_length, output_name)
    resolved_source_path = source_path.resolve()
    resolved_output_path = output_path.resolve()
    if (
        resolved_output_path == resolved_source_path
        or str(resolved_output_path).casefold() == str(resolved_source_path).casefold()
    ):
        raise ValueError(
            'output_name must not match the source document filename; '
            'choose a different output_name'
        )

    source_doc = aw.Document(str(source_path))

    model = model_factory(model_name_value, api_key_value)

    options = aw.ai.SummarizeOptions()
    options.summary_length = summary_length_enum

    if summarize_call is None:
        summary_doc = model.summarize(source_document=source_doc, options=options)
    else:
        summary_doc = summarize_call(model, source_doc, options)

    summary_doc.save(str(output_path), aw.SaveFormat.DOCX)

    return {
        'sourceDocId': doc_id_value,
        'sourcePath': str(source_path),
        'outputFilename': output_path.name,
        'outputPath': str(output_path),
        'modelName': model_name_value,
        'summaryLength': normalized_summary_length,
    }
