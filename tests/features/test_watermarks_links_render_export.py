import base64

import pytest

pytest.importorskip('aspose.words')
import aspose.words as aw

import mcp_server as srv
from core.export import replace_regex_to_images_base64
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
    ],
)
def test_export_base64_advanced_formats(fmt, expected_ext, expected_mime):
    r = srv.tool_create_document('p0-adv-export.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, f'Export {fmt}')
    options = None
    if fmt == 'html':
        options = {'embed_resources': True}
    elif fmt == 'pdf':
        options = {'compliance': 'PDF_A_1B'}
    out = srv.tool_export_base64_advanced(did, fmt=fmt, options=options)
    raw = base64.b64decode(out['base64'])
    assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
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


def test_replace_regex_to_images_base64_success_and_shape():
    response = srv.tool_create_document('regex-images-shape.docx')
    doc_id = response['docId']
    srv.tool_add_paragraph(doc_id, 'Token-100 Token-200')

    result = replace_regex_to_images_base64(
        doc_id=doc_id,
        pattern=r'Token-\d+',
        replacement_text='Replaced',
        fmt='png',
        dpi=140,
        case_sensitive=True,
    )

    assert 'images' in result
    assert isinstance(result['images'], list)
    assert len(result['images']) > 0

    for image_payload in result['images']:
        assert set(image_payload.keys()) == {'base64', 'mime', 'ext'}
        assert image_payload['mime'] == 'image/png'
        assert image_payload['ext'] == 'png'
        decoded = base64.b64decode(image_payload['base64'])
        assert isinstance(decoded, (bytes, bytearray))
        assert len(decoded) > 0


def test_replace_regex_to_images_base64_preserves_all_images_from_api():
    response = srv.tool_create_document('regex-images-multi.docx')
    doc_id = response['docId']
    srv.tool_add_paragraph(doc_id, 'Order-100')
    srv.tool_add_page_break_end(doc_id)
    srv.tool_add_paragraph(doc_id, 'Order-200')

    file_path = ensure_path(doc_id)
    save_options = aw.saving.ImageSaveOptions(aw.SaveFormat.PNG)
    save_options.horizontal_resolution = 110.0
    save_options.vertical_resolution = 110.0
    raw_streams = aw.lowcode.Replacer.replace_to_images_regex(
        str(file_path), save_options, r'Order-\d+', 'Updated'
    )

    result = replace_regex_to_images_base64(
        doc_id=doc_id,
        pattern=r'Order-\d+',
        replacement_text='Updated',
        fmt='png',
        dpi=110,
        case_sensitive=True,
    )

    assert len(result['images']) == len(raw_streams)


def test_replace_regex_to_images_base64_svg_is_explicitly_unsupported():
    response = srv.tool_create_document('regex-images-svg.docx')
    doc_id = response['docId']
    srv.tool_add_paragraph(doc_id, 'Item-100 Item-200')

    with pytest.raises(ValueError, match='Unsupported replace-to-images format: svg'):
        replace_regex_to_images_base64(
            doc_id=doc_id,
            pattern=r'Item-\d+',
            replacement_text='Replaced',
            fmt='svg',
            dpi=120,
            case_sensitive=True,
        )


def test_replace_regex_to_images_base64_invalid_pattern_is_explicit():
    response = srv.tool_create_document('regex-images-invalid.docx')
    doc_id = response['docId']
    srv.tool_add_paragraph(doc_id, 'Token-300')

    with pytest.raises(RuntimeError) as exc_info:
        replace_regex_to_images_base64(
            doc_id=doc_id,
            pattern='([a',
            replacement_text='X',
            fmt='png',
            dpi=120,
            case_sensitive=True,
        )

    message = str(exc_info.value).strip()
    assert message
    message_lower = message.lower()
    assert (
        '([a' in message
        or 'regex' in message_lower
        or 'pattern' in message_lower
        or 'parse' in message_lower
        or 'invalid' in message_lower
    )
