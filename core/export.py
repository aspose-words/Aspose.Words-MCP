from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import aspose.words as aw

from core.utils.docs_util import ensure_path, ensure_resources_dir

_CUSTOM_NODE_IDS_SIDECAR_SUFFIX = '.custom_node_ids.json'
_ASPOSE_EVALUATION_EXACT_PARAGRAPHS = {
    'evaluation only. created with aspose.words. copyright 2003-2026 aspose pty ltd.',
    'this document was truncated here because it was created in the evaluation mode.',
}
_ASPOSE_EVALUATION_FOOTER_PREFIX = (
    'created with an evaluation copy of aspose.words. to remove all limitations, '
    'you can use free temporary license '
)
_ASPOSE_EVALUATION_FOOTER_SUFFIX = 'https://products.aspose.com/words/temporary-license/'


def _custom_node_ids_sidecar_path(doc_path: Path) -> Path:
    return doc_path.with_suffix(f'{doc_path.suffix}{_CUSTOM_NODE_IDS_SIDECAR_SUFFIX}')


def _normalize_paragraph_text(paragraph_text: str) -> str:
    return paragraph_text.rstrip('\r\n')


def _is_evaluation_only_paragraph(paragraph_text: str) -> bool:
    normalized_text = _normalize_paragraph_text(paragraph_text).strip().lower()
    if normalized_text == '':
        return False
    if normalized_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS:
        return True
    return normalized_text.startswith(
        _ASPOSE_EVALUATION_FOOTER_PREFIX
    ) and normalized_text.endswith(_ASPOSE_EVALUATION_FOOTER_SUFFIX)


def _evaluation_paragraph_indices(paragraph_texts: list[str]) -> set[int]:
    normalized_texts = [
        _normalize_paragraph_text(paragraph_text).strip().lower()
        for paragraph_text in paragraph_texts
    ]
    evaluation_indices: set[int] = set()

    leading_index: Optional[int] = None
    for index, normalized_text in enumerate(normalized_texts):
        if normalized_text != '':
            leading_index = index
            break

    if leading_index is not None:
        leading_text = normalized_texts[leading_index]
        if leading_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS:
            evaluation_indices.add(leading_index)

    trailing_index = len(normalized_texts) - 1
    while trailing_index >= 0:
        trailing_text = normalized_texts[trailing_index]
        if trailing_text == '':
            trailing_index -= 1
            continue

        is_evaluation_footer = trailing_text.startswith(
            _ASPOSE_EVALUATION_FOOTER_PREFIX
        ) and trailing_text.endswith(_ASPOSE_EVALUATION_FOOTER_SUFFIX)
        if trailing_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS or is_evaluation_footer:
            evaluation_indices.add(trailing_index)
            trailing_index -= 1
            continue

        break

    return evaluation_indices


def _collect_user_paragraphs(doc: aw.Document) -> list[aw.Paragraph]:
    paragraph_nodes = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    paragraph_texts: list[str] = []
    paragraphs: list[aw.Paragraph] = []
    for index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[index].as_paragraph()
        paragraphs.append(paragraph)
        paragraph_texts.append(paragraph.to_string(aw.SaveFormat.TEXT) or '')

    evaluation_indices = _evaluation_paragraph_indices(paragraph_texts)

    user_paragraphs: list[aw.Paragraph] = []
    for index, paragraph in enumerate(paragraphs):
        if index in evaluation_indices:
            continue
        user_paragraphs.append(paragraph)
    return user_paragraphs


def _resolve_sidecar_entry_ordinal(
    entry_key_ordinal: int, entry_payload: dict[str, object], user_paragraphs: list[aw.Paragraph]
) -> Optional[int]:
    payload_ordinal_raw = entry_payload.get('user_paragraph_ordinal')
    payload_ordinal = int(payload_ordinal_raw) if payload_ordinal_raw is not None else None
    if payload_ordinal is None or payload_ordinal == entry_key_ordinal:
        return entry_key_ordinal

    normalized_text_raw = entry_payload.get('normalized_text')
    if isinstance(normalized_text_raw, str) and normalized_text_raw != '':
        key_candidate_matches = False
        payload_candidate_matches = False

        if 0 <= entry_key_ordinal < len(user_paragraphs):
            key_candidate_text = (
                user_paragraphs[entry_key_ordinal].to_string(aw.SaveFormat.TEXT) or ''
            )
            key_candidate_matches = (
                _normalize_paragraph_text(key_candidate_text) == normalized_text_raw
            )

        if 0 <= payload_ordinal < len(user_paragraphs):
            payload_candidate_text = (
                user_paragraphs[payload_ordinal].to_string(aw.SaveFormat.TEXT) or ''
            )
            payload_candidate_matches = (
                _normalize_paragraph_text(payload_candidate_text) == normalized_text_raw
            )

        if payload_candidate_matches and not key_candidate_matches:
            return payload_ordinal

    return entry_key_ordinal


def _reapply_persisted_custom_node_ids(doc: aw.Document, doc_id: str, doc_path: Path) -> None:
    sidecar_path = _custom_node_ids_sidecar_path(doc_path)
    if not sidecar_path.exists():
        return

    sidecar_data = json.loads(sidecar_path.read_text(encoding='utf-8'))
    if not isinstance(sidecar_data, dict):
        raise ValueError('Custom node ID sidecar must be a JSON object')

    sidecar_doc_id = sidecar_data.get('doc_id')
    if sidecar_doc_id is not None and str(sidecar_doc_id) != doc_id:
        return

    paragraph_custom_node_ids = sidecar_data.get('paragraph_custom_node_ids')
    if paragraph_custom_node_ids is None:
        return
    if not isinstance(paragraph_custom_node_ids, dict):
        raise ValueError('paragraph_custom_node_ids must be a JSON object')

    user_paragraphs = _collect_user_paragraphs(doc)
    for user_paragraph_ordinal_raw, entry_payload_raw in paragraph_custom_node_ids.items():
        if not isinstance(entry_payload_raw, dict):
            raise ValueError('paragraph_custom_node_ids entries must be JSON objects')

        custom_node_id_raw = entry_payload_raw.get('custom_node_id')
        if custom_node_id_raw is None:
            raise ValueError('paragraph_custom_node_ids entries must include custom_node_id')

        user_paragraph_ordinal = int(user_paragraph_ordinal_raw)
        resolved_ordinal = _resolve_sidecar_entry_ordinal(
            user_paragraph_ordinal, entry_payload_raw, user_paragraphs
        )
        if resolved_ordinal is None:
            continue

        normalized_text_raw = entry_payload_raw.get('normalized_text')
        if isinstance(normalized_text_raw, str) and normalized_text_raw != '':
            if resolved_ordinal < 0 or resolved_ordinal >= len(user_paragraphs):
                continue
            resolved_text = user_paragraphs[resolved_ordinal].to_string(aw.SaveFormat.TEXT) or ''
            if _normalize_paragraph_text(resolved_text) != normalized_text_raw:
                continue

        if resolved_ordinal < 0 or resolved_ordinal >= len(user_paragraphs):
            continue
        paragraph = user_paragraphs[resolved_ordinal]
        paragraph.custom_node_id = int(custom_node_id_raw)


def with_svg_embed_options() -> Any:
    sso = aw.saving.SvgSaveOptions()
    sso.export_embedded_images = True
    ensure_resources_dir('svg', sso)
    return sso


def export_markdown(doc: Any) -> bytes:
    import tempfile as _tmp
    from pathlib import Path

    data: bytes = b''
    with _tmp.NamedTemporaryFile(suffix='.md', delete=True) as tf:
        tmp_path = Path(tf.name)
        doc.save(str(tmp_path), aw.SaveFormat.MARKDOWN)
        data = tmp_path.read_bytes()
    return data


def build_pdf_opts(options: Dict[str, Any]) -> Any:
    pdf_opts = aw.saving.PdfSaveOptions()
    comp = (options or {}).get('compliance')
    if comp:
        m = {
            'PDF_A1A': aw.saving.PdfCompliance.PDF_A1A,
            'PDF_A1B': aw.saving.PdfCompliance.PDF_A1B,
        }
        key = str(comp).upper()
        key_norm = key.replace('_A_', 'A')
        if key_norm in m:
            pdf_opts.compliance = m[key_norm]
    if (options or {}).get('export_document_structure') is True:
        pdf_opts.export_document_structure = True
    return pdf_opts


def build_html_opts(fmt_key: str, embed_resources: bool) -> Any:
    if fmt_key == 'html_fixed':
        opts_hf = aw.saving.HtmlFixedSaveOptions()
        ensure_resources_dir('html', opts_hf)
        return opts_hf
    if fmt_key == 'mhtml':
        opts = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
    else:
        opts = aw.saving.HtmlSaveOptions()
    opts.export_images_as_base64 = bool(embed_resources)
    if not embed_resources:
        ensure_resources_dir('html', opts)
    return opts


def export(doc_id: str, fmt: str = 'docx') -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'docx').lower()
    if fmt_l == 'pdf':
        save_format = aw.SaveFormat.PDF
        mime = 'application/pdf'
        ext = 'pdf'
    elif fmt_l == 'rtf':
        save_format = aw.SaveFormat.RTF
        mime = 'application/rtf'
        ext = 'rtf'
    else:
        save_format = aw.SaveFormat.DOCX
        mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ext = 'docx'
    out = BytesIO()
    doc.save(out, save_format)
    data = out.getvalue()
    return data, mime, ext


def render_page(
    doc_id: str, page_index: int = 0, fmt: str = 'png', dpi: int = 150
) -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'png').lower()
    if fmt_l in ('jpeg', 'jpg'):
        save_format = aw.SaveFormat.JPEG
        mime = 'image/jpeg'
        ext = 'jpg'
    elif fmt_l == 'svg':
        save_format = aw.SaveFormat.SVG
        mime = 'image/svg+xml'
        ext = 'svg'
    elif fmt_l == 'tiff':
        save_format = aw.SaveFormat.TIFF
        mime = 'image/tiff'
        ext = 'tiff'
    elif fmt_l == 'png':
        save_format = aw.SaveFormat.PNG
        mime = 'image/png'
        ext = 'png'
    else:
        raise ValueError(f'Unsupported render format: {fmt}')
    single = doc.extract_pages(int(page_index), 1)
    out = BytesIO()
    if fmt_l == 'svg':
        sso = aw.saving.SvgSaveOptions()
        sso.export_embedded_images = True
        ensure_resources_dir('svg', sso)
        single.save(out, sso)
    else:
        iso = aw.saving.ImageSaveOptions(save_format)
        iso.horizontal_resolution = float(dpi)
        iso.vertical_resolution = float(dpi)
        single.save(out, iso)
    return out.getvalue(), mime, ext


def export_advanced(
    doc_id: str, fmt: str = 'docx', options: Optional[Dict[str, Any]] = None
) -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'docx').lower()
    opts = options or {}
    specs: Dict[str, Dict[str, Any]] = {
        'html': {
            'mime': 'text/html',
            'ext': 'html',
            'save_format': aw.SaveFormat.HTML,
            'builder': lambda: build_html_opts('html', bool(opts.get('embed_resources', True))),
        },
        'mhtml': {
            'mime': 'message/rfc822',
            'ext': 'mhtml',
            'save_format': aw.SaveFormat.MHTML,
            'builder': lambda: build_html_opts('mhtml', bool(opts.get('embed_resources', True))),
        },
        'html_fixed': {
            'mime': 'text/html',
            'ext': 'html',
            'save_format': aw.SaveFormat.HTML_FIXED,
            'builder': lambda: build_html_opts(
                'html_fixed', bool(opts.get('embed_resources', True))
            ),
        },
        'epub': {'mime': 'application/epub+zip', 'ext': 'epub', 'save_format': aw.SaveFormat.EPUB},
        'odt': {
            'mime': 'application/vnd.oasis.opendocument.text',
            'ext': 'odt',
            'save_format': aw.SaveFormat.ODT,
        },
        'md': {'mime': 'text/markdown', 'ext': 'md', 'custom': True},
        'markdown': {'mime': 'text/markdown', 'ext': 'md', 'custom': True},
        'svg': {
            'mime': 'image/svg+xml',
            'ext': 'svg',
            'save_format': aw.SaveFormat.SVG,
            'builder': lambda: with_svg_embed_options(),
        },
        'pdf': {
            'mime': 'application/pdf',
            'ext': 'pdf',
            'save_format': aw.SaveFormat.PDF,
            'builder': lambda: build_pdf_opts(opts),
        },
        'docling': {
            'mime': 'application/json',
            'ext': 'json',
            'save_format': aw.SaveFormat.DOCLING,
            'builder': lambda: _build_docling_opts(),
        },
    }
    spec = specs.get(fmt_l)
    if not spec:
        raise ValueError(f'Unsupported export format: {fmt}')
    if fmt_l == 'pdf' and opts.get('enable_text_shaping') is True:
        doc.layout_options.enable_text_shaping = True
    if spec.get('custom'):
        data = export_markdown(doc)
        return data, spec['mime'], spec['ext']
    should_reapply_custom_node_ids = (
        fmt_l == 'pdf' and opts.get('export_document_structure') is True
    )
    if should_reapply_custom_node_ids:
        _reapply_persisted_custom_node_ids(doc, doc_id, file_path)

    save_opts = spec.get('builder')() if spec.get('builder') else None
    out = BytesIO()
    if save_opts is not None:
        doc.save(out, save_opts)
    else:
        doc.save(out, spec['save_format'])
    return out.getvalue(), spec['mime'], spec['ext']


def _build_docling_opts() -> Any:
    docling_opts = aw.saving.DoclingSaveOptions()
    docling_opts.save_format = aw.SaveFormat.DOCLING
    return docling_opts
