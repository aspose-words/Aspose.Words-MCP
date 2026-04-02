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
