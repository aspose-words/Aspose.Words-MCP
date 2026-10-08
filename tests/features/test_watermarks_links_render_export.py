import base64
import datetime
import json
from pathlib import Path
from types import SimpleNamespace

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
    if fmt == 'xps':
        export_calls = [None, {'compression_level': 'maximum'}]

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


def test_tool_export_base64_advanced_forwards_digital_signature_payload(monkeypatch):
    export_calls = []
    digital_signature = {
        'certificate_path': str(Path('certificates') / 'signing.pfx'),
        'certificate_password': 'certificate-secret',
        'reason': 'Approval',
        'location': 'Remote',
        'signature_date': '2026-10-08T12:30:00Z',
        'hash_algorithm': 'sha512',
        'timestamp_settings': {
            'server_url': 'https://tsa.example.test',
            'user_name': 'tsa-user',
            'password': 'tsa-secret',
            'timeout': 15,
        },
    }

    def fake_export_advanced(doc_id, fmt='docx', options=None):
        export_calls.append({'doc_id': doc_id, 'fmt': fmt, 'options': options})
        return b'fake-pdf', 'application/pdf', 'pdf'

    monkeypatch.setattr(srv._export, 'export_advanced', fake_export_advanced)

    response = srv.tool_export_base64_advanced(
        'doc-id',
        fmt='pdf',
        options={'digital_signature': digital_signature},
    )

    assert response == {
        'base64': base64.b64encode(b'fake-pdf').decode('utf-8'),
        'mime': 'application/pdf',
        'ext': 'pdf',
    }
    assert export_calls == [
        {
            'doc_id': 'doc-id',
            'fmt': 'pdf',
            'options': {'digital_signature': digital_signature},
        }
    ]


@pytest.mark.parametrize(
    'hash_algorithm,expected_hash',
    [
        ('sha256', 'sha256-constant'),
        ('sha384', 'sha384-constant'),
        ('sha512', 'sha512-constant'),
        ('ripe_md160', 'ripe-md160-constant'),
        ('SHA-256', 'sha256-constant'),
        ('RIPE-MD160', 'ripe-md160-constant'),
    ],
)
def test_pdf_export_assigns_digital_signature_details_before_save(
    monkeypatch, tmp_path, hash_algorithm, expected_hash
):
    from core import export as _export

    source_path = tmp_path / 'source.docx'
    source_path.write_text('document')
    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate')
    holder = object()
    events = []

    class FakeSaveFormat:
        HTML = 'html'
        MHTML = 'mhtml'
        HTML_FIXED = 'html-fixed'
        EPUB = 'epub'
        ODT = 'odt'
        SVG = 'svg'
        PDF = 'pdf'
        DOCLING = 'docling'
        XPS = 'xps'

    class FakePdfDigitalSignatureHashAlgorithm:
        SHA256 = 'sha256-constant'
        SHA384 = 'sha384-constant'
        SHA512 = 'sha512-constant'
        RIPE_MD160 = 'ripe-md160-constant'

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, passphrase):
            events.append(('holder', (file_name, passphrase)))
            return holder

    class FakePdfDigitalSignatureTimestampSettings:
        def __init__(self, *args):
            events.append(('timestamp', args))

    class FakePdfDigitalSignatureDetails:
        def __setattr__(self, name, value):
            events.append((f'details:{name}', value))
            object.__setattr__(self, name, value)

    class FakePdfSaveOptions:
        def __setattr__(self, name, value):
            events.append((f'pdf_opts:{name}', value))
            object.__setattr__(self, name, value)

    class FakeDocument:
        def __init__(self, file_path):
            events.append(('document', file_path))

        def save(self, output, save_options):
            events.append(('save', save_options.digital_signature_details))
            assert save_options.digital_signature_details.certificate_holder is holder
            assert save_options.digital_signature_details.reason == 'Approval'
            assert save_options.digital_signature_details.location == 'Remote'
            assert save_options.digital_signature_details.signature_date == datetime.datetime(
                2026, 10, 8, 12, 30, tzinfo=datetime.timezone.utc
            )
            assert save_options.digital_signature_details.hash_algorithm == expected_hash
            output.write(b'fake-pdf')

    fake_aw = SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        digitalsignatures=SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=SimpleNamespace(
            PdfSaveOptions=FakePdfSaveOptions,
            PdfDigitalSignatureDetails=FakePdfDigitalSignatureDetails,
            PdfDigitalSignatureHashAlgorithm=FakePdfDigitalSignatureHashAlgorithm,
            PdfDigitalSignatureTimestampSettings=FakePdfDigitalSignatureTimestampSettings,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)
    monkeypatch.setattr(_export, 'ensure_path', lambda doc_id: source_path)

    data, mime, ext = _export.export_advanced(
        'doc-id',
        fmt='pdf',
        options={
            'digital_signature': {
                'certificate_path': str(certificate_path),
                'certificate_password': 'certificate-secret',
                'reason': 'Approval',
                'location': 'Remote',
                'signature_date': '2026-10-08T12:30:00Z',
                'hash_algorithm': hash_algorithm,
            }
        },
    )

    assert data == b'fake-pdf'
    assert mime == 'application/pdf'
    assert ext == 'pdf'
    assert ('holder', (str(certificate_path), 'certificate-secret')) in events
    assert ('details:certificate_holder', holder) in events
    assert ('details:reason', 'Approval') in events
    assert ('details:location', 'Remote') in events
    assert ('details:hash_algorithm', expected_hash) in events
    assert not [event for event in events if event[0] == 'timestamp']


@pytest.mark.parametrize(
    'timestamp_settings,expected_args',
    [
        (
            {'server_url': 'https://tsa.example.test'},
            ('https://tsa.example.test', '', ''),
        ),
        (
            {
                'server_url': 'https://tsa.example.test',
                'user_name': 'tsa-user',
                'password': 'tsa-secret',
                'timeout': 12.5,
            },
            (
                'https://tsa.example.test',
                'tsa-user',
                'tsa-secret',
                datetime.timedelta(seconds=12.5),
            ),
        ),
    ],
)
def test_pdf_export_builds_digital_signature_timestamp_settings(
    monkeypatch, tmp_path, timestamp_settings, expected_args
):
    from core import export as _export

    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate')
    holder = object()
    holder_calls = []
    timestamp_instances = []

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, passphrase):
            holder_calls.append((file_name, passphrase))
            return holder

    class FakePdfDigitalSignatureDetails:
        pass

    class FakePdfDigitalSignatureTimestampSettings:
        def __init__(self, *args):
            self.args = args
            timestamp_instances.append(self)

    fake_aw = SimpleNamespace(
        digitalsignatures=SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=SimpleNamespace(
            PdfDigitalSignatureDetails=FakePdfDigitalSignatureDetails,
            PdfDigitalSignatureTimestampSettings=FakePdfDigitalSignatureTimestampSettings,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)

    details = _export.build_pdf_digital_signature_details(
        {
            'certificate_path': str(certificate_path),
            'timestamp_settings': timestamp_settings,
        }
    )

    assert details.certificate_holder is holder
    assert holder_calls == [(str(certificate_path), '')]
    assert details.timestamp_settings is timestamp_instances[0]
    assert timestamp_instances[0].args == expected_args


@pytest.mark.parametrize(
    'digital_signature,expected_error,expected_message',
    [
        ({}, ValueError, 'certificate_path must be non-empty'),
        ({'certificate_path': ''}, ValueError, 'certificate_path must be non-empty'),
        ({'certificate_path': '   '}, ValueError, 'certificate_path must be non-empty'),
        ({'certificate_path': 'missing'}, FileNotFoundError, 'Certificate file not found'),
        (
            {'certificate_path': 'cert', 'signature_date': ''},
            ValueError,
            'signature_date must be a non-empty ISO 8601 string',
        ),
        (
            {'certificate_path': 'cert', 'signature_date': 'not-a-date'},
            ValueError,
            'signature_date must be a valid ISO 8601 string',
        ),
        (
            {'certificate_path': 'cert', 'hash_algorithm': ''},
            ValueError,
            'hash_algorithm must be a non-empty string',
        ),
        (
            {'certificate_path': 'cert', 'hash_algorithm': 'sha1'},
            ValueError,
            'Unsupported digital_signature.hash_algorithm',
        ),
        (
            {'certificate_path': 'cert', 'timestamp_settings': {}},
            ValueError,
            'timestamp_settings.server_url must be non-empty',
        ),
        (
            {'certificate_path': 'cert', 'timestamp_settings': {'server_url': '   '}},
            ValueError,
            'timestamp_settings.server_url must be non-empty',
        ),
        (
            {
                'certificate_path': 'cert',
                'timestamp_settings': {'server_url': 'https://tsa.example.test', 'timeout': 0},
            },
            ValueError,
            'timestamp_settings.timeout must be a positive number',
        ),
        (
            {
                'certificate_path': 'cert',
                'timestamp_settings': {
                    'server_url': 'https://tsa.example.test',
                    'timeout': 'soon',
                },
            },
            ValueError,
            'timestamp_settings.timeout must be a positive number',
        ),
        (
            {
                'certificate_path': 'cert',
                'timestamp_settings': {
                    'server_url': 'https://tsa.example.test',
                    'timeout': 'nan',
                },
            },
            ValueError,
            'timestamp_settings.timeout must be a positive number',
        ),
    ],
)
def test_pdf_digital_signature_validation_errors(
    monkeypatch, tmp_path, digital_signature, expected_error, expected_message
):
    from core import export as _export

    certificate_path = tmp_path / 'cert'
    certificate_path.write_text('certificate')
    resolved_signature = {
        key: (str(certificate_path) if value == 'cert' else value)
        for key, value in digital_signature.items()
    }

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, passphrase):
            return object()

    class FakePdfDigitalSignatureDetails:
        pass

    class FakePdfDigitalSignatureHashAlgorithm:
        SHA256 = 'sha256-constant'
        SHA384 = 'sha384-constant'
        SHA512 = 'sha512-constant'
        RIPE_MD160 = 'ripe-md160-constant'

    class FakePdfDigitalSignatureTimestampSettings:
        def __init__(self, *args):
            self.args = args

    fake_aw = SimpleNamespace(
        digitalsignatures=SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=SimpleNamespace(
            PdfDigitalSignatureDetails=FakePdfDigitalSignatureDetails,
            PdfDigitalSignatureHashAlgorithm=FakePdfDigitalSignatureHashAlgorithm,
            PdfDigitalSignatureTimestampSettings=FakePdfDigitalSignatureTimestampSettings,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)

    with pytest.raises(expected_error, match=expected_message):
        _export.build_pdf_digital_signature_details(resolved_signature)


def test_pdf_digital_signature_aspose_failures_propagate(monkeypatch, tmp_path):
    from core import export as _export

    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate')

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, passphrase):
            raise RuntimeError('certificate failure')

    fake_aw = SimpleNamespace(
        digitalsignatures=SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=SimpleNamespace(),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)

    with pytest.raises(RuntimeError, match='certificate failure'):
        _export.build_pdf_digital_signature_details({'certificate_path': str(certificate_path)})


@pytest.mark.parametrize(
    'failing_api,expected_message',
    [
        ('details', 'details failure'),
        ('timestamp', 'timestamp failure'),
        ('save', 'save failure'),
    ],
)
def test_pdf_digital_signature_details_timestamp_and_save_failures_propagate(
    monkeypatch, tmp_path, failing_api, expected_message
):
    from core import export as _export

    source_path = tmp_path / 'source.docx'
    source_path.write_text('document')
    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate')

    class FakeSaveFormat:
        HTML = 'html'
        MHTML = 'mhtml'
        HTML_FIXED = 'html-fixed'
        EPUB = 'epub'
        ODT = 'odt'
        SVG = 'svg'
        PDF = 'pdf'
        DOCLING = 'docling'
        XPS = 'xps'

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name, passphrase):
            return object()

    class FakePdfDigitalSignatureDetails:
        def __init__(self):
            if failing_api == 'details':
                raise RuntimeError('details failure')

    class FakePdfDigitalSignatureTimestampSettings:
        def __init__(self, *args):
            if failing_api == 'timestamp':
                raise RuntimeError('timestamp failure')

    class FakePdfSaveOptions:
        pass

    class FakeDocument:
        def __init__(self, file_path):
            self.file_path = file_path

        def save(self, output, save_options):
            if failing_api == 'save':
                raise RuntimeError('save failure')
            output.write(b'fake-pdf')

    fake_aw = SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        digitalsignatures=SimpleNamespace(CertificateHolder=FakeCertificateHolder),
        saving=SimpleNamespace(
            PdfSaveOptions=FakePdfSaveOptions,
            PdfDigitalSignatureDetails=FakePdfDigitalSignatureDetails,
            PdfDigitalSignatureTimestampSettings=FakePdfDigitalSignatureTimestampSettings,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)
    monkeypatch.setattr(_export, 'ensure_path', lambda doc_id: source_path)

    options = {
        'digital_signature': {
            'certificate_path': str(certificate_path),
            'timestamp_settings': {'server_url': 'https://tsa.example.test'},
        }
    }

    with pytest.raises(RuntimeError, match=expected_message):
        _export.export_advanced('doc-id', fmt='pdf', options=options)


def test_pdf_export_uses_explicit_digital_signature_api_references():
    export_source = Path('core/export.py').read_text(encoding='utf-8')
    assert 'aw.digitalsignatures.CertificateHolder.create' in export_source
    assert 'aw.saving.PdfDigitalSignatureDetails()' in export_source
    assert 'details.certificate_holder = holder' in export_source
    assert 'details.timestamp_settings =' in export_source
    assert 'pdf_opts.digital_signature_details =' in export_source
    assert 'aw.saving.PdfDigitalSignatureTimestampSettings(' in export_source
    assert 'aw.saving.PdfDigitalSignatureHashAlgorithm.SHA256' in export_source
    assert 'aw.saving.PdfDigitalSignatureHashAlgorithm.SHA384' in export_source
    assert 'aw.saving.PdfDigitalSignatureHashAlgorithm.SHA512' in export_source
    assert 'aw.saving.PdfDigitalSignatureHashAlgorithm.RIPE_MD160' in export_source
    assert 'getattr(' not in export_source
    assert 'hasattr(' not in export_source


@pytest.mark.parametrize(
    'compression_level,expected_level',
    [
        ('normal', 'normal-constant'),
        ('maximum', 'maximum-constant'),
        ('fast', 'fast-constant'),
        ('super_fast', 'super-fast-constant'),
        ('SUPER-FAST', 'super-fast-constant'),
    ],
)
def test_xps_export_assigns_compression_level_before_save(
    monkeypatch, tmp_path, compression_level, expected_level
):
    from core import export as _export

    saved = {}

    class FakeSaveFormat:
        HTML = 'html'
        MHTML = 'mhtml'
        HTML_FIXED = 'html-fixed'
        EPUB = 'epub'
        ODT = 'odt'
        SVG = 'svg'
        PDF = 'pdf'
        DOCLING = 'docling'
        XPS = 'xps'

    class FakeCompressionLevel:
        NORMAL = 'normal-constant'
        MAXIMUM = 'maximum-constant'
        FAST = 'fast-constant'
        SUPER_FAST = 'super-fast-constant'

    class FakeXpsSaveOptions:
        def __init__(self):
            self.compression_level = None

    class FakeDocument:
        def __init__(self, file_path):
            saved['file_path'] = file_path

        def save(self, output, save_options):
            saved['save_options_type'] = type(save_options)
            saved['compression_level'] = save_options.compression_level
            output.write(b'fake-xps')

    fake_aw = SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        saving=SimpleNamespace(
            XpsSaveOptions=FakeXpsSaveOptions,
            CompressionLevel=FakeCompressionLevel,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)
    monkeypatch.setattr(_export, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')

    data, mime, ext = _export.export_advanced(
        'doc-1', fmt='xps', options={'compression_level': compression_level}
    )

    assert data == b'fake-xps'
    assert mime == 'application/vnd.ms-xpsdocument'
    assert ext == 'xps'
    assert saved['save_options_type'] is FakeXpsSaveOptions
    assert saved['compression_level'] == expected_level


@pytest.mark.parametrize(
    'compression_level,expected_message',
    [
        ('', 'non-empty string'),
        ('   ', 'non-empty string'),
        ('smallest', 'Unsupported XPS compression_level'),
    ],
)
def test_xps_export_rejects_invalid_compression_level(
    monkeypatch, tmp_path, compression_level, expected_message
):
    from core import export as _export

    class FakeSaveFormat:
        HTML = 'html'
        MHTML = 'mhtml'
        HTML_FIXED = 'html-fixed'
        EPUB = 'epub'
        ODT = 'odt'
        SVG = 'svg'
        PDF = 'pdf'
        DOCLING = 'docling'
        XPS = 'xps'

    class FakeCompressionLevel:
        NORMAL = 'normal-constant'
        MAXIMUM = 'maximum-constant'
        FAST = 'fast-constant'
        SUPER_FAST = 'super-fast-constant'

    class FakeXpsSaveOptions:
        def __init__(self):
            self.compression_level = None

    class FakeDocument:
        def __init__(self, file_path):
            self.file_path = file_path

        def save(self, output, save_options):
            raise AssertionError('Document.save should not be called for invalid XPS options')

    fake_aw = SimpleNamespace(
        Document=FakeDocument,
        SaveFormat=FakeSaveFormat,
        saving=SimpleNamespace(
            XpsSaveOptions=FakeXpsSaveOptions,
            CompressionLevel=FakeCompressionLevel,
        ),
    )
    monkeypatch.setattr(_export, 'aw', fake_aw)
    monkeypatch.setattr(_export, 'ensure_path', lambda doc_id: tmp_path / f'{doc_id}.docx')

    with pytest.raises(ValueError, match=expected_message):
        _export.export_advanced(
            'doc-1', fmt='xps', options={'compression_level': compression_level}
        )


def test_xps_export_uses_explicit_aspose_api_references():
    export_source = Path('core/export.py').read_text(encoding='utf-8')
    assert 'aw.saving.XpsSaveOptions()' in export_source
    assert 'xps_opts.compression_level = levels[key]' in export_source
    assert 'aw.saving.CompressionLevel.NORMAL' in export_source
    assert 'aw.saving.CompressionLevel.MAXIMUM' in export_source
    assert 'aw.saving.CompressionLevel.FAST' in export_source
    assert 'aw.saving.CompressionLevel.SUPER_FAST' in export_source
    assert 'getattr(xps_opts' not in export_source
    assert 'hasattr(xps_opts' not in export_source
