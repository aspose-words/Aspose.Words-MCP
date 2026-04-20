import base64
import json

import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv


def _png_1x1_b64():
    return (
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AABQMB'
        'gYJ4V7wAAAAASUVORK5CYII='
    )


def _has_struct_tree_root(pdf_bytes: bytes) -> bool:
    return b'/StructTreeRoot' in pdf_bytes or b'/Type/StructTreeRoot' in pdf_bytes


def _has_marked_true(pdf_bytes: bytes) -> bool:
    return b'/MarkInfo<</Marked true>>' in pdf_bytes or b'/Marked true' in pdf_bytes


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


def test_export_base64_advanced_pdf_tagged_control_shape_and_markers():
    created = srv.tool_create_document('p0-adv-export-tagged-pdf.docx')
    doc_id = created['docId']

    srv.tool_add_heading(doc_id, 'Tagged PDF Validation', level=1, custom_node_id=1101)
    srv.tool_add_paragraph(
        doc_id,
        'Paragraph for tagged structure verification.',
        custom_node_id=2202,
    )

    default_pdf = srv.tool_export_base64_advanced(doc_id, fmt='pdf')
    default_pdf_explicit_false = srv.tool_export_base64_advanced(
        doc_id,
        fmt='pdf',
        tagged_pdf_preserve_custom_node_ids=False,
    )
    tagged_pdf = srv.tool_export_base64_advanced(
        doc_id,
        fmt='pdf',
        tagged_pdf_preserve_custom_node_ids=True,
    )

    for export_out in (default_pdf, default_pdf_explicit_false, tagged_pdf):
        assert set(export_out) == {'base64', 'mime', 'ext'}
        assert export_out['ext'] == 'pdf'
        assert export_out['mime'] == 'application/pdf'

    default_pdf_bytes = base64.b64decode(default_pdf['base64'])
    default_pdf_explicit_false_bytes = base64.b64decode(default_pdf_explicit_false['base64'])
    tagged_pdf_bytes = base64.b64decode(tagged_pdf['base64'])

    assert len(default_pdf_bytes) > 0
    assert len(default_pdf_explicit_false_bytes) > 0
    assert len(tagged_pdf_bytes) > 0

    assert _has_struct_tree_root(default_pdf_bytes) is False
    assert _has_marked_true(default_pdf_bytes) is False
    assert _has_struct_tree_root(default_pdf_explicit_false_bytes) is False
    assert _has_marked_true(default_pdf_explicit_false_bytes) is False

    assert _has_struct_tree_root(tagged_pdf_bytes) is True
    assert _has_marked_true(tagged_pdf_bytes) is True


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
