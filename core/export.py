from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import aspose.words as aw

from core.utils.docs_util import ensure_path, ensure_resources_dir


def _normalize_paragraph_text(text: str) -> str:
    return ' '.join(text.split())


def _apply_pdf_custom_node_id_metadata(doc: Any, file_path: Path) -> None:
    sidecar_path = file_path.with_suffix('.custom_node_id.json')
    if not sidecar_path.exists():
        return

    with sidecar_path.open('r', encoding='utf-8') as sidecar_file:
        records = json.load(sidecar_file)

    if not isinstance(records, list):
        raise ValueError('Custom-node-ID sidecar must be a JSON list of records')

    paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

    claimed_paragraph_indices: set[int] = set()

    for record_index, record in enumerate(records):
        if not isinstance(record, dict):
            raise ValueError(
                f'Custom-node-ID sidecar record at index {record_index} must be an object'
            )

        kind = record.get('kind')
        paragraph_index = record.get('paragraph_index')
        expected_text = record.get('expected_text')
        custom_node_id = record.get('custom_node_id')

        if kind != 'paragraph':
            raise ValueError(
                f'Unsupported custom-node-ID record kind at index {record_index}: {kind!r}'
            )
        if not isinstance(paragraph_index, int):
            raise ValueError(
                f'Custom-node-ID paragraph_index at index {record_index} must be an integer'
            )
        if paragraph_index < 0:
            raise ValueError(
                f'Custom-node-ID paragraph_index at index {record_index} must be non-negative'
            )
        if not isinstance(expected_text, str):
            raise ValueError(
                f'Custom-node-ID expected_text at index {record_index} must be a string'
            )
        if not isinstance(custom_node_id, int):
            raise ValueError(
                f'Custom-node-ID custom_node_id at index {record_index} must be an integer'
            )

        target_paragraph = None
        target_paragraph_index: Optional[int] = None
        normalized_expected_text = _normalize_paragraph_text(expected_text)

        if paragraph_index < paragraphs.count:
            indexed_paragraph = paragraphs[paragraph_index]
            indexed_text = indexed_paragraph.to_string(aw.SaveFormat.TEXT) or ''
            if (
                _normalize_paragraph_text(indexed_text) == normalized_expected_text
                and paragraph_index not in claimed_paragraph_indices
            ):
                target_paragraph = indexed_paragraph
                target_paragraph_index = paragraph_index

        if target_paragraph is None:
            matches: list[tuple[int, Any]] = []
            for current_index in range(paragraphs.count):
                if current_index in claimed_paragraph_indices:
                    continue
                paragraph = paragraphs[current_index]
                paragraph_text = paragraph.to_string(aw.SaveFormat.TEXT) or ''
                if _normalize_paragraph_text(paragraph_text) == normalized_expected_text:
                    matches.append((current_index, paragraph))

            if len(matches) == 1:
                target_paragraph_index, target_paragraph = matches[0]
            elif len(matches) == 0:
                raise ValueError(
                    'Unable to resolve paragraph for custom-node-ID record '
                    f'at index {record_index}: no text match found'
                )
            else:
                best_distance = min(abs(index - paragraph_index) for index, _ in matches)
                nearest_matches = [
                    (index, paragraph)
                    for index, paragraph in matches
                    if abs(index - paragraph_index) == best_distance
                ]

                if len(nearest_matches) == 1:
                    target_paragraph_index, target_paragraph = nearest_matches[0]
                else:
                    raise ValueError(
                        'Unable to resolve paragraph for custom-node-ID record '
                        f'at index {record_index}: multiple text matches found'
                    )

        if target_paragraph_index is None:
            raise ValueError(
                'Unable to resolve paragraph for custom-node-ID record '
                f'at index {record_index}: paragraph index resolution failed'
            )

        if target_paragraph_index in claimed_paragraph_indices:
            raise ValueError(
                'Unable to resolve paragraph for custom-node-ID record '
                f'at index {record_index}: paragraph already matched by earlier record'
            )

        claimed_paragraph_indices.add(target_paragraph_index)

        target_paragraph.custom_node_id = custom_node_id


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
        _apply_pdf_custom_node_id_metadata(doc, file_path)
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
    if fmt_l == 'pdf':
        _apply_pdf_custom_node_id_metadata(doc, file_path)
    if spec.get('custom'):
        data = export_markdown(doc)
        return data, spec['mime'], spec['ext']
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
