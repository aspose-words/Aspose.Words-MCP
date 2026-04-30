import base64
import json
from pathlib import Path

import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import mcp_server as srv
from core.utils.docs_util import ensure_path


def _png_1x1_b64():
    return (
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABQMB'
        'gYJ4V7wAAAAASUVORK5CYII='
    )


@pytest.mark.parametrize(
    'fmt,expected_ext,expected_mime_prefix',
    [
        ('png', 'png', 'image/'),
        ('jpeg', 'jpg', 'image/'),
        ('svg', 'svg', 'image/'),
        ('tiff', 'tiff', 'image/'),
    ],
)
def test_watermarks_and_render(fmt, expected_ext, expected_mime_prefix):
    r = srv.tool_create_document('p0-wm.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Body')
    srv.tool_add_watermark_text(did, text='CONFIDENTIAL')
    srv.tool_add_watermark_image_base64(did, image_base64=_png_1x1_b64())
    out = srv.tool_render_page_base64(did, page_index=0, fmt=fmt, dpi=120)
    raw = base64.b64decode(out['base64'])
    assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
    assert out['ext'] == expected_ext
    assert out['mime'].startswith(expected_mime_prefix)


@pytest.mark.parametrize(
    'fmt,expected_ext,expected_mime',
    [
        ('html', 'html', 'text/html'),
        ('html_fixed', 'html', 'text/html'),
        ('mhtml', 'mhtml', 'message/'),
        ('epub', 'epub', 'application/epub+zip'),
        ('odt', 'odt', 'application/vnd.oasis.opendocument.text'),
        ('md', 'md', 'text/markdown'),
        ('svg', 'svg', 'image/svg+xml'),
        ('pdf', 'pdf', 'application/pdf'),
        ('docling', 'json', 'application/json'),
    ],
)
def test_export_base64_advanced_formats(fmt, expected_ext, expected_mime):
    r = srv.tool_create_document('p0-adv-export.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, f'Export {fmt}')
    export_calls = [{'embed_resources': True}] if fmt == 'html' else [None]
    if fmt == 'pdf':
        export_calls = [None, {'enable_text_shaping': True}]

    for options in export_calls:
        out = srv.tool_export_base64_advanced(did, fmt=fmt, options=options)
        assert set(out) >= {'base64', 'mime', 'ext'}
        raw = base64.b64decode(out['base64'])
        assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
        if fmt == 'docling':
            assert out['ext'] == 'json'
            assert out['mime'] == 'application/json'
            payload = raw.decode('utf-8')
            parsed = json.loads(payload)
            assert isinstance(parsed, (dict, list))
        assert out['ext'] == expected_ext
        if expected_mime.endswith('/'):
            assert out['mime'].startswith(expected_mime)
        else:
            assert out['mime'] == expected_mime


def test_bookmarks_and_hyperlinks():
    r = srv.tool_create_document('p0-links.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'P0')
    srv.tool_add_bookmark_at_paragraph(did, name='BM_TEST', paragraph_index=0)
    srv.tool_insert_hyperlink_at_paragraph(
        did, paragraph_index=0, text='Example', target='https://example.com', external=True
    )
    xml = srv.tool_get_xml(did)['xml']
    assert 'BM_TEST' in xml
    assert 'example.com' in xml


def _write_custom_node_id_sidecar(doc_id: str, payload: str) -> Path:
    doc_path = ensure_path(doc_id)
    sidecar_path = doc_path.with_suffix('.custom_node_id.json')
    sidecar_path.write_text(payload, encoding='utf-8')
    return sidecar_path


def test_export_pdf_rejects_malformed_custom_node_id_sidecar_json():
    r = srv.tool_create_document('p0-pdf-bad-sidecar-json.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Alpha')

    _write_custom_node_id_sidecar(did, '{"kind": "paragraph"')

    with pytest.raises(json.JSONDecodeError):
        srv.tool_export_base64_advanced(did, fmt='pdf')


def test_export_pdf_rejects_invalid_custom_node_id_record_fields():
    r = srv.tool_create_document('p0-pdf-invalid-record-fields.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Alpha')

    _write_custom_node_id_sidecar(
        did,
        json.dumps(
            [
                {
                    'kind': 'paragraph',
                    'paragraph_index': '0',
                    'expected_text': 'Alpha',
                    'custom_node_id': 7,
                }
            ]
        ),
    )

    with pytest.raises(ValueError, match='paragraph_index .* must be an integer'):
        srv.tool_export_base64_advanced(did, fmt='pdf')


def test_export_pdf_rejects_ambiguous_custom_node_id_paragraph_resolution():
    r = srv.tool_create_document('p0-pdf-ambiguous-resolution.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Duplicate')
    srv.tool_add_paragraph(did, 'Spacer')
    srv.tool_add_paragraph(did, 'Duplicate')

    _write_custom_node_id_sidecar(
        did,
        json.dumps(
            [
                {
                    'kind': 'paragraph',
                    'paragraph_index': 2,
                    'expected_text': 'Duplicate',
                    'custom_node_id': 7,
                }
            ]
        ),
    )

    with pytest.raises(ValueError, match='multiple text matches found'):
        srv.tool_export_base64_advanced(did, fmt='pdf')


def test_export_pdf_rejects_oversized_custom_node_id_assignment():
    r = srv.tool_create_document('p0-pdf-oversized-custom-node-id.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Alpha')

    _write_custom_node_id_sidecar(
        did,
        json.dumps(
            [
                {
                    'kind': 'paragraph',
                    'paragraph_index': 0,
                    'expected_text': 'Alpha',
                    'custom_node_id': 2**65,
                }
            ]
        ),
    )

    with pytest.raises((OverflowError, RuntimeError, ValueError)):
        srv.tool_export_base64_advanced(did, fmt='pdf')


def test_export_pdf_rejects_duplicate_paragraph_claims_in_sidecar_mapping():
    r = srv.tool_create_document('p0-pdf-duplicate-paragraph-claims.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Alpha')

    _write_custom_node_id_sidecar(
        did,
        json.dumps(
            [
                {
                    'kind': 'paragraph',
                    'paragraph_index': 0,
                    'expected_text': 'Alpha',
                    'custom_node_id': 7,
                },
                {
                    'kind': 'paragraph',
                    'paragraph_index': 0,
                    'expected_text': 'Alpha',
                    'custom_node_id': 9,
                },
            ]
        ),
    )

    with pytest.raises(ValueError, match='no text match found'):
        srv.tool_export_base64_advanced(did, fmt='pdf')


def test_export_pdf_applies_persisted_paragraph_custom_node_id_and_emits_marker(monkeypatch):
    original_build_pdf_opts = srv._export.build_pdf_opts

    def _build_pdf_opts_without_text_compression(options):
        pdf_opts = original_build_pdf_opts(options)
        pdf_opts.text_compression = aw.saving.PdfTextCompression.NONE
        pdf_opts.export_document_structure = True
        return pdf_opts

    monkeypatch.setattr(
        srv._export,
        'build_pdf_opts',
        _build_pdf_opts_without_text_compression,
    )

    r = srv.tool_create_document('p0-pdf-custom-node-id-positive.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'Paragraph with exported custom node id marker')
    srv.tool_set_paragraph_custom_node_id(did, paragraph_index=0, custom_node_id=4242)

    out = srv.tool_export_base64_advanced(did, fmt='pdf')
    pdf_data = base64.b64decode(out['base64'])

    assert b'XXAsposeWords/CustomId' in pdf_data
