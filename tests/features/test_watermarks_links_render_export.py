import base64
import json
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path

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
        ('xps', 'xps', 'application/vnd.ms-xpsdocument'),
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


def test_pdf_export_generate_form_field_scripts_option():
    r = srv.tool_create_document('p0-pdf-form-field-scripts.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'PDF form field scripts')

    out = srv.tool_export_base64_advanced(
        did, fmt='pdf', options={'generate_form_field_scripts': True}
    )

    raw = base64.b64decode(out['base64'])
    assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
    assert out['ext'] == 'pdf'
    assert out['mime'] == 'application/pdf'

    export_source = Path('core/export.py').read_text(encoding='utf-8')
    assert 'pdf_opts.generate_form_field_scripts = True' in export_source
    assert 'getattr(pdf_opts' not in export_source
    assert 'hasattr(pdf_opts' not in export_source


def _patch_pdf_export_aw(monkeypatch, tmp_path):
    events = []
    saves = []
    certificate_holder = object()

    class FakePdfCompliance:
        PDF_A1A = 'PDF_A1A'
        PDF_A1B = 'PDF_A1B'

    class FakePdfSaveOptions:
        def __init__(self):
            self.compliance = None
            self.generate_form_field_scripts = False
            self.digital_signature_details = None

    class FakePdfDigitalSignatureDetails:
        def __init__(self, certificate_holder_arg, reason, location, signature_date):
            events.append(
                (
                    'signature_details',
                    (certificate_holder_arg, reason, location, signature_date),
                )
            )
            self.certificate_holder = certificate_holder_arg
            self.reason = reason
            self.location = location
            self.signature_date = signature_date
            self.timestamp_settings = None

    class FakePdfDigitalSignatureTimestampSettings:
        def __init__(self, server_url, user_name, password, timeout=None):
            events.append(('timestamp_settings', (server_url, user_name, password, timeout)))
            self.server_url = server_url
            self.user_name = user_name
            self.password = password
            self.timeout = timeout

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, password):
            events.append(('certificate_holder', (file_name, password)))
            return certificate_holder

    class FakeSaveFormat:
        HTML = 'HTML'
        MHTML = 'MHTML'
        HTML_FIXED = 'HTML_FIXED'
        EPUB = 'EPUB'
        ODT = 'ODT'
        SVG = 'SVG'
        PDF = 'PDF'
        XPS = 'XPS'
        DOCLING = 'DOCLING'

    class FakeDocument:
        def __init__(self, path):
            self.path = path
            self.layout_options = types.SimpleNamespace(enable_text_shaping=False)

        def save(self, out, save_options):
            saves.append((self, save_options))
            out.write(b'pdf bytes')

    fake_aw = types.SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        digitalsignatures=types.SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=types.SimpleNamespace(
            PdfCompliance=FakePdfCompliance,
            PdfSaveOptions=FakePdfSaveOptions,
            PdfDigitalSignatureDetails=FakePdfDigitalSignatureDetails,
            PdfDigitalSignatureTimestampSettings=FakePdfDigitalSignatureTimestampSettings,
        ),
    )
    monkeypatch.setattr(srv._export, 'aw', fake_aw)
    monkeypatch.setattr(srv._export, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')
    return events, saves, certificate_holder, FakePdfSaveOptions


def test_pdf_export_assigns_digital_signature_details(monkeypatch, tmp_path):
    events, saves, certificate_holder, fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    certificate_path = tmp_path / 'ml-dsa-signing.pfx'
    certificate_path.write_text('certificate bytes')
    signature_date = '2026-08-10T12:34:56Z'

    out = srv.tool_export_base64_advanced(
        'doc-1',
        fmt='pdf',
        options={
            'digital_signature': {
                'certificate_path': f' {certificate_path} ',
                'certificate_passphrase': 'certificate-secret',
                'reason': 'Release approval',
                'location': 'Build server',
                'signature_date': signature_date,
            }
        },
    )

    assert base64.b64decode(out['base64']) == b'pdf bytes'
    assert out['ext'] == 'pdf'
    assert out['mime'] == 'application/pdf'
    assert len(saves) == 1
    saved_options = saves[0][1]
    assert isinstance(saved_options, fake_pdf_save_options)
    assert saved_options.digital_signature_details.reason == 'Release approval'
    assert saved_options.digital_signature_details.location == 'Build server'
    assert saved_options.digital_signature_details.certificate_holder is certificate_holder
    assert saved_options.digital_signature_details.signature_date == datetime(
        2026, 8, 10, 12, 34, 56, tzinfo=timezone.utc
    )
    assert saved_options.digital_signature_details.timestamp_settings is None
    assert events == [
        ('certificate_holder', (str(certificate_path), 'certificate-secret')),
        (
            'signature_details',
            (
                certificate_holder,
                'Release approval',
                'Build server',
                datetime(2026, 8, 10, 12, 34, 56, tzinfo=timezone.utc),
            ),
        ),
    ]


def test_pdf_export_assigns_timestamp_settings_and_preserves_pdf_options(
    monkeypatch, tmp_path
):
    events, saves, certificate_holder, _fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    certificate_path = tmp_path / 'timestamp-signing.pfx'
    certificate_path.write_text('certificate bytes')

    srv.tool_export_base64_advanced(
        'doc-2',
        fmt='pdf',
        options={
            'compliance': 'pdf_a_1b',
            'generate_form_field_scripts': True,
            'enable_text_shaping': True,
            'digital_signature': {
                'certificate_path': str(certificate_path),
                'timestamp': {
                    'server_url': ' https://timestamp.example/tsa ',
                    'user_name': 'timestamp-user',
                    'password': 'timestamp-secret',
                    'timeout_seconds': '45',
                },
            },
        },
    )

    saved_document, saved_options = saves[0]
    assert saved_document.layout_options.enable_text_shaping is True
    assert saved_options.compliance == 'PDF_A1B'
    assert saved_options.generate_form_field_scripts is True
    assert saved_options.digital_signature_details.reason == ''
    assert saved_options.digital_signature_details.location == ''
    assert saved_options.digital_signature_details.timestamp_settings.timeout == timedelta(
        seconds=45
    )
    assert events == [
        (
            'timestamp_settings',
            (
                'https://timestamp.example/tsa',
                'timestamp-user',
                'timestamp-secret',
                timedelta(seconds=45),
            ),
        ),
        ('certificate_holder', (str(certificate_path), '')),
        (
            'signature_details',
            (certificate_holder, '', '', saved_options.digital_signature_details.signature_date),
        ),
    ]


def test_pdf_export_assigns_default_timestamp_timeout(monkeypatch, tmp_path):
    events, saves, _certificate_holder, _fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    certificate_path = tmp_path / 'timestamp-default.pfx'
    certificate_path.write_text('certificate bytes')

    srv.tool_export_base64_advanced(
        'doc-2-default',
        fmt='pdf',
        options={
            'digital_signature': {
                'certificate_path': str(certificate_path),
                'timestamp': {
                    'server_url': 'https://timestamp.example/tsa',
                    'user_name': 'timestamp-user',
                    'password': 'timestamp-secret',
                },
            },
        },
    )

    saved_options = saves[0][1]
    assert saved_options.digital_signature_details.timestamp_settings.timeout is None
    assert events[0] == (
        'timestamp_settings',
        ('https://timestamp.example/tsa', 'timestamp-user', 'timestamp-secret', None),
    )


@pytest.mark.parametrize(
    'digital_signature,error_type,error_message',
    [
        ({}, ValueError, 'digital_signature.certificate_path must be non-empty'),
        (
            {'certificate_path': '   '},
            ValueError,
            'digital_signature.certificate_path must be non-empty',
        ),
        (
            {'certificate_path': 'missing.pfx'},
            FileNotFoundError,
            'Certificate file not found',
        ),
    ],
)
def test_pdf_export_rejects_invalid_certificate_before_signing(
    monkeypatch, tmp_path, digital_signature, error_type, error_message
):
    events, _saves, _certificate_holder, _fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    if digital_signature.get('certificate_path') == 'missing.pfx':
        digital_signature['certificate_path'] = str(tmp_path / 'missing.pfx')

    with pytest.raises(error_type, match=error_message):
        srv.tool_export_base64_advanced(
            'doc-3',
            fmt='pdf',
            options={'digital_signature': digital_signature},
        )

    assert events == []


@pytest.mark.parametrize('signature_date', ['', 'not-a-date'])
def test_pdf_export_rejects_invalid_signature_date_before_signing(
    monkeypatch, tmp_path, signature_date
):
    events, _saves, _certificate_holder, _fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    certificate_path = tmp_path / 'signature-date-validation.pfx'
    certificate_path.write_text('certificate bytes')

    with pytest.raises(ValueError, match='digital_signature.signature_date'):
        srv.tool_export_base64_advanced(
            'doc-3-date',
            fmt='pdf',
            options={
                'digital_signature': {
                    'certificate_path': str(certificate_path),
                    'signature_date': signature_date,
                }
            },
        )

    assert events == []


@pytest.mark.parametrize(
    'timestamp,error_message',
    [
        ({}, 'digital_signature.timestamp.server_url must be non-empty'),
        ({'server_url': ''}, 'digital_signature.timestamp.server_url must be non-empty'),
        (
            {'server_url': 'https://timestamp.example/tsa', 'timeout_seconds': ''},
            'digital_signature.timestamp.timeout_seconds must be greater than 0',
        ),
        (
            {'server_url': 'https://timestamp.example/tsa', 'timeout_seconds': 0},
            'digital_signature.timestamp.timeout_seconds must be greater than 0',
        ),
    ],
)
def test_pdf_export_rejects_invalid_timestamp_before_signing(
    monkeypatch, tmp_path, timestamp, error_message
):
    events, _saves, _certificate_holder, _fake_pdf_save_options = _patch_pdf_export_aw(
        monkeypatch, tmp_path
    )
    certificate_path = tmp_path / 'timestamp-validation.pfx'
    certificate_path.write_text('certificate bytes')

    with pytest.raises(ValueError, match=error_message):
        srv.tool_export_base64_advanced(
            'doc-4',
            fmt='pdf',
            options={
                'digital_signature': {
                    'certificate_path': str(certificate_path),
                    'timestamp': timestamp,
                }
            },
        )

    assert events == []


def test_registered_export_base64_advanced_forwards_digital_signature_options(monkeypatch):
    captured_tool_functions = {}
    export_calls = []

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_export_base64_advanced(doc_id, fmt, options=None):
        export_calls.append({'doc_id': doc_id, 'fmt': fmt, 'options': options})
        return {'base64': 'cGRm', 'mime': 'application/pdf', 'ext': 'pdf'}

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(
        srv,
        'tool_export_base64_advanced',
        fake_tool_export_base64_advanced,
    )
    options = {
        'digital_signature': {
            'certificate_path': '/certificates/signing.pfx',
            'timestamp': {'server_url': 'https://timestamp.example/tsa'},
        }
    }

    srv.register_tools()
    registered_export = captured_tool_functions['export_base64_advanced']
    response = registered_export('doc-id', fmt='pdf', options=options)

    assert response == {'base64': 'cGRm', 'mime': 'application/pdf', 'ext': 'pdf'}
    assert export_calls == [{'doc_id': 'doc-id', 'fmt': 'pdf', 'options': options}]


def test_xps_export_with_maximum_compression():
    r = srv.tool_create_document('p0-xps-compression.docx')
    did = r['docId']
    srv.tool_add_paragraph(did, 'XPS maximum compression export')

    out = srv.tool_export_base64_advanced(did, fmt='xps', options={'compression_level': 'maximum'})

    raw = base64.b64decode(out['base64'])
    assert isinstance(raw, (bytes, bytearray)) and len(raw) > 0
    assert out['ext'] == 'xps'
    assert out['mime'] == 'application/vnd.ms-xpsdocument'


@pytest.mark.parametrize(
    'compression_level,expected',
    [
        (None, None),
        ('normal', 'NORMAL'),
        ('MAXIMUM', 'MAXIMUM'),
        (' fast ', 'FAST'),
        ('super_fast', 'SUPER_FAST'),
        ('super-fast', 'SUPER_FAST'),
        ('super fast', 'SUPER_FAST'),
    ],
)
def test_xps_export_assigns_explicit_compression_level(
    monkeypatch, tmp_path, compression_level, expected
):
    saves = []

    class FakeCompressionLevel:
        NORMAL = 'NORMAL'
        MAXIMUM = 'MAXIMUM'
        FAST = 'FAST'
        SUPER_FAST = 'SUPER_FAST'

    class FakeXpsSaveOptions:
        def __init__(self):
            self.compression_level = None

    class FakeSaveFormat:
        HTML = 'HTML'
        MHTML = 'MHTML'
        HTML_FIXED = 'HTML_FIXED'
        EPUB = 'EPUB'
        ODT = 'ODT'
        SVG = 'SVG'
        PDF = 'PDF'
        XPS = 'XPS'
        DOCLING = 'DOCLING'

    class FakeDocument:
        def __init__(self, path):
            self.path = path
            self.layout_options = types.SimpleNamespace(enable_text_shaping=False)

        def save(self, out, save_options):
            saves.append(save_options)
            out.write(b'xps bytes')

    fake_aw = types.SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        saving=types.SimpleNamespace(
            CompressionLevel=FakeCompressionLevel,
            XpsSaveOptions=FakeXpsSaveOptions,
        ),
    )
    monkeypatch.setattr(srv._export, 'aw', fake_aw)
    monkeypatch.setattr(srv._export, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')

    options = {} if compression_level is None else {'compression_level': compression_level}
    out = srv.tool_export_base64_advanced('doc-1', fmt='xps', options=options)

    assert base64.b64decode(out['base64']) == b'xps bytes'
    assert out['ext'] == 'xps'
    assert out['mime'] == 'application/vnd.ms-xpsdocument'
    assert len(saves) == 1
    assert isinstance(saves[0], FakeXpsSaveOptions)
    assert saves[0].compression_level == expected


@pytest.mark.parametrize('compression_level', ['smallest', '', '   '])
def test_xps_export_rejects_invalid_compression_before_save(
    monkeypatch, tmp_path, compression_level
):
    class FakeCompressionLevel:
        NORMAL = 'NORMAL'
        MAXIMUM = 'MAXIMUM'
        FAST = 'FAST'
        SUPER_FAST = 'SUPER_FAST'

    class FakeXpsSaveOptions:
        def __init__(self):
            self.compression_level = None

    class FakeSaveFormat:
        HTML = 'HTML'
        MHTML = 'MHTML'
        HTML_FIXED = 'HTML_FIXED'
        EPUB = 'EPUB'
        ODT = 'ODT'
        SVG = 'SVG'
        PDF = 'PDF'
        XPS = 'XPS'
        DOCLING = 'DOCLING'

    class FakeDocument:
        def __init__(self, path):
            self.path = path
            self.layout_options = types.SimpleNamespace(enable_text_shaping=False)

        def save(self, out, save_options):
            raise AssertionError('Document.save must not be called for invalid compression_level')

    fake_aw = types.SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        saving=types.SimpleNamespace(
            CompressionLevel=FakeCompressionLevel,
            XpsSaveOptions=FakeXpsSaveOptions,
        ),
    )
    monkeypatch.setattr(srv._export, 'aw', fake_aw)
    monkeypatch.setattr(srv._export, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')

    with pytest.raises(ValueError, match='Unsupported XPS compression_level') as exc_info:
        srv.tool_export_base64_advanced(
            'doc-1', fmt='xps', options={'compression_level': compression_level}
        )

    message = str(exc_info.value)
    assert 'normal' in message
    assert 'maximum' in message
    assert 'fast' in message
    assert 'super_fast' in message
