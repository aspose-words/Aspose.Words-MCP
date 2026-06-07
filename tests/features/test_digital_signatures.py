import types
from pathlib import Path

import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv
from core import signatures as _signatures


def test_list_digital_signatures_extracts_26_5_metadata(monkeypatch, tmp_path):
    source_path = tmp_path / 'signed.docx'
    source_path.write_text('signed document')
    signature_events: list[tuple[str, object]] = []

    class FakeDigitalSignature:
        def __init__(
            self,
            application_version: str,
            color_depth: int,
            horizontal_resolution: int,
            office_version: str,
            vertical_resolution: int,
            windows_version: str,
        ) -> None:
            self.application_version = application_version
            self.color_depth = color_depth
            self.horizontal_resolution = horizontal_resolution
            self.office_version = office_version
            self.vertical_resolution = vertical_resolution
            self.windows_version = windows_version

    class FakeDigitalSignatureUtil:
        @staticmethod
        def load_signatures(file_name: str):
            signature_events.append(('load_signatures', file_name))
            return [
                FakeDigitalSignature(
                    '26.5.0-app',
                    32,
                    300,
                    'Office 2024',
                    600,
                    'Windows 11',
                )
            ]

    def fake_ensure_path(doc_id: str) -> Path:
        signature_events.append(('ensure_path', doc_id))
        return source_path

    monkeypatch.setattr(_signatures, 'ensure_path', fake_ensure_path)
    monkeypatch.setattr(
        _signatures.aw,
        'digitalsignatures',
        types.SimpleNamespace(DigitalSignatureUtil=FakeDigitalSignatureUtil),
    )

    signatures_response = _signatures.list_digital_signatures('doc-id')

    assert signature_events == [
        ('ensure_path', 'doc-id'),
        ('load_signatures', str(source_path)),
    ]
    assert signatures_response == [
        {
            'application_version': '26.5.0-app',
            'color_depth': 32,
            'horizontal_resolution': 300,
            'office_version': 'Office 2024',
            'vertical_resolution': 600,
            'windows_version': 'Windows 11',
        }
    ]


def test_list_digital_signatures_returns_empty_list_for_empty_collection(
    monkeypatch, tmp_path
):
    source_path = tmp_path / 'unsigned.docx'
    source_path.write_text('unsigned document')
    signature_events: list[tuple[str, object]] = []

    class FakeDigitalSignatureUtil:
        @staticmethod
        def load_signatures(file_name: str):
            signature_events.append(('load_signatures', file_name))
            return []

    def fake_ensure_path(doc_id: str) -> Path:
        signature_events.append(('ensure_path', doc_id))
        return source_path

    monkeypatch.setattr(_signatures, 'ensure_path', fake_ensure_path)
    monkeypatch.setattr(
        _signatures.aw,
        'digitalsignatures',
        types.SimpleNamespace(DigitalSignatureUtil=FakeDigitalSignatureUtil),
    )

    signatures_response = _signatures.list_digital_signatures('doc-id')

    assert signature_events == [
        ('ensure_path', 'doc-id'),
        ('load_signatures', str(source_path)),
    ]
    assert signatures_response == []


@pytest.mark.parametrize(
    'certificate_path, expected_error',
    [('', ValueError), ('   ', ValueError), ('missing', FileNotFoundError)],
)
def test_sign_document_validates_certificate_path_before_aspose_calls(
    monkeypatch, tmp_path, certificate_path, expected_error
):
    source_path = tmp_path / 'source.docx'
    source_path.write_text('unsigned document')
    aspose_calls: list[str] = []

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name: str, passphrase: str):
            aspose_calls.append(f'holder:{file_name}:{passphrase}')
            raise AssertionError('certificate holder must not be created for invalid paths')

    class FakeDigitalSignatureUtil:
        @staticmethod
        def sign(src_file_name: str, dst_file_name: str, cert_holder, sign_options):
            aspose_calls.append(f'sign:{src_file_name}:{dst_file_name}')
            raise AssertionError('digital signature must not be attempted for invalid paths')

    monkeypatch.setattr(_signatures, 'ensure_path', lambda doc_id: source_path)
    monkeypatch.setattr(
        _signatures.aw,
        'digitalsignatures',
        types.SimpleNamespace(
            CertificateHolder=FakeCertificateHolder,
            DigitalSignatureUtil=FakeDigitalSignatureUtil,
            SignOptions=object,
        ),
    )

    invalid_certificate_path = (
        str(tmp_path / 'missing-certificate.pfx')
        if certificate_path == 'missing'
        else certificate_path
    )

    with pytest.raises(expected_error):
        _signatures.sign_document('doc-id', invalid_certificate_path)

    assert aspose_calls == []


def test_sign_document_assigns_26_5_sign_options_and_replaces_source(
    monkeypatch, tmp_path
):
    source_path = tmp_path / 'source.docx'
    source_path.write_text('unsigned document')
    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate bytes')
    cert_holder = object()
    aspose_events: list[tuple[str, object]] = []

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name: str, passphrase: str):
            aspose_events.append(('holder', (file_name, passphrase)))
            return cert_holder

    class FakeSignOptions:
        def __setattr__(self, name: str, assigned_object: object) -> None:
            aspose_events.append((f'assign:{name}', assigned_object))
            object.__setattr__(self, name, assigned_object)

    class FakeDigitalSignatureUtil:
        @staticmethod
        def sign(
            src_file_name: str, dst_file_name: str, cert_holder, sign_options
        ) -> None:
            aspose_events.append(('sign', (src_file_name, dst_file_name, cert_holder)))
            assert sign_options.application_version == '26.5.0-app'
            assert sign_options.color_depth == 32
            assert sign_options.horizontal_resolution == 300
            assert sign_options.office_version == 'Office 2024'
            assert sign_options.vertical_resolution == 600
            assert sign_options.windows_version == 'Windows 11'
            Path(dst_file_name).write_text('signed document')

    monkeypatch.setattr(_signatures, 'ensure_path', lambda doc_id: source_path)
    monkeypatch.setattr(
        _signatures.aw,
        'digitalsignatures',
        types.SimpleNamespace(
            CertificateHolder=FakeCertificateHolder,
            DigitalSignatureUtil=FakeDigitalSignatureUtil,
            SignOptions=FakeSignOptions,
        ),
    )

    sign_succeeded = _signatures.sign_document(
        'doc-id',
        str(certificate_path),
        'test-passphrase',
        application_version='26.5.0-app',
        color_depth=32,
        horizontal_resolution=300,
        office_version='Office 2024',
        vertical_resolution=600,
        windows_version='Windows 11',
    )

    signed_path = source_path.with_name(f'{source_path.stem}.signed{source_path.suffix}')
    assert sign_succeeded is True
    assert source_path.read_text() == 'signed document'
    assert not signed_path.exists()
    assert aspose_events == [
        ('holder', (str(certificate_path), 'test-passphrase')),
        ('assign:application_version', '26.5.0-app'),
        ('assign:color_depth', 32),
        ('assign:horizontal_resolution', 300),
        ('assign:office_version', 'Office 2024'),
        ('assign:vertical_resolution', 600),
        ('assign:windows_version', 'Windows 11'),
        ('sign', (str(source_path), str(signed_path), cert_holder)),
    ]


def test_sign_document_omitted_optional_sign_options_preserve_defaults(
    monkeypatch, tmp_path
):
    source_path = tmp_path / 'source.docx'
    source_path.write_text('unsigned document')
    certificate_path = tmp_path / 'certificate.pfx'
    certificate_path.write_text('certificate bytes')
    cert_holder = object()
    aspose_events: list[tuple[str, object]] = []

    class FakeCertificateHolder:
        @staticmethod
        def create(file_name: str, passphrase: str):
            aspose_events.append(('holder', (file_name, passphrase)))
            return cert_holder

    class FakeSignOptions:
        def __init__(self) -> None:
            object.__setattr__(self, 'application_version', 'default-app')
            object.__setattr__(self, 'color_depth', 24)
            object.__setattr__(self, 'horizontal_resolution', 96)
            object.__setattr__(self, 'office_version', 'default-office')
            object.__setattr__(self, 'vertical_resolution', 96)
            object.__setattr__(self, 'windows_version', 'default-windows')

        def __setattr__(self, name: str, assigned_object: object) -> None:
            if name in {
                'color_depth',
                'horizontal_resolution',
                'vertical_resolution',
            } and assigned_object is None:
                raise AssertionError(f'{name} must preserve its integer default')
            aspose_events.append((f'assign:{name}', assigned_object))
            object.__setattr__(self, name, assigned_object)

    class FakeDigitalSignatureUtil:
        @staticmethod
        def sign(
            src_file_name: str, dst_file_name: str, cert_holder, sign_options
        ) -> None:
            aspose_events.append(('sign', (src_file_name, dst_file_name, cert_holder)))
            assert sign_options.application_version == 'default-app'
            assert sign_options.color_depth == 24
            assert sign_options.horizontal_resolution == 96
            assert sign_options.office_version == 'default-office'
            assert sign_options.vertical_resolution == 96
            assert sign_options.windows_version == 'default-windows'
            Path(dst_file_name).write_text('signed document')

    monkeypatch.setattr(_signatures, 'ensure_path', lambda doc_id: source_path)
    monkeypatch.setattr(
        _signatures.aw,
        'digitalsignatures',
        types.SimpleNamespace(
            CertificateHolder=FakeCertificateHolder,
            DigitalSignatureUtil=FakeDigitalSignatureUtil,
            SignOptions=FakeSignOptions,
        ),
    )

    sign_succeeded = _signatures.sign_document('doc-id', str(certificate_path))

    signed_path = source_path.with_name(f'{source_path.stem}.signed{source_path.suffix}')
    assert sign_succeeded is True
    assert source_path.read_text() == 'signed document'
    assert not signed_path.exists()
    assert aspose_events == [
        ('holder', (str(certificate_path), '')),
        ('sign', (str(source_path), str(signed_path), cert_holder)),
    ]


def test_tool_sign_document_forwards_26_5_sign_options(monkeypatch):
    signature_calls = []
    certificate_path = str(Path('certificates') / 'signing.pfx')

    def fake_sign_document(
        doc_id: str,
        certificate_path: str,
        certificate_passphrase: str = '',
        application_version: str | None = None,
        color_depth: int | None = None,
        horizontal_resolution: int | None = None,
        office_version: str | None = None,
        vertical_resolution: int | None = None,
        windows_version: str | None = None,
    ) -> bool:
        signature_calls.append(
            {
                'doc_id': doc_id,
                'certificate_path': certificate_path,
                'certificate_passphrase': certificate_passphrase,
                'application_version': application_version,
                'color_depth': color_depth,
                'horizontal_resolution': horizontal_resolution,
                'office_version': office_version,
                'vertical_resolution': vertical_resolution,
                'windows_version': windows_version,
            }
        )
        return True

    monkeypatch.setattr(srv._signatures, 'sign_document', fake_sign_document)

    signing_response = srv.tool_sign_document(
        doc_id='doc-id',
        certificate_path=certificate_path,
        certificate_passphrase='certificate-secret',
        application_version='26.5.0-app',
        color_depth=32,
        horizontal_resolution=300,
        office_version='Office 2024',
        vertical_resolution=600,
        windows_version='Windows 11',
    )

    assert signing_response == {}
    assert signature_calls == [
        {
            'doc_id': 'doc-id',
            'certificate_path': certificate_path,
            'certificate_passphrase': 'certificate-secret',
            'application_version': '26.5.0-app',
            'color_depth': 32,
            'horizontal_resolution': 300,
            'office_version': 'Office 2024',
            'vertical_resolution': 600,
            'windows_version': 'Windows 11',
        }
    ]


def test_tool_get_digital_signatures_wraps_signatures_and_forwards_doc_id(
    monkeypatch,
):
    signature_calls = []
    expected_signatures = [
        {
            'application_version': '26.5.0-app',
            'color_depth': 32,
            'horizontal_resolution': 300,
            'office_version': 'Office 2024',
            'vertical_resolution': 600,
            'windows_version': 'Windows 11',
        }
    ]

    def fake_list_digital_signatures(doc_id: str) -> list[dict[str, object]]:
        signature_calls.append(doc_id)
        return expected_signatures

    monkeypatch.setattr(
        srv._signatures,
        'list_digital_signatures',
        fake_list_digital_signatures,
    )

    signatures_response = srv.tool_get_digital_signatures('doc-id')

    assert signatures_response == {'signatures': expected_signatures}
    assert signature_calls == ['doc-id']


def test_registered_sign_document_forwards_26_5_sign_options(monkeypatch):
    captured_tool_functions = {}
    tool_sign_calls = []
    certificate_path = str(Path('certificates') / 'signing.pfx')

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_sign_document(
        doc_id: str,
        certificate_path: str,
        certificate_passphrase: str = '',
        application_version: str | None = None,
        color_depth: int | None = None,
        horizontal_resolution: int | None = None,
        office_version: str | None = None,
        vertical_resolution: int | None = None,
        windows_version: str | None = None,
    ):
        tool_sign_calls.append(
            {
                'doc_id': doc_id,
                'certificate_path': certificate_path,
                'certificate_passphrase': certificate_passphrase,
                'application_version': application_version,
                'color_depth': color_depth,
                'horizontal_resolution': horizontal_resolution,
                'office_version': office_version,
                'vertical_resolution': vertical_resolution,
                'windows_version': windows_version,
            }
        )
        return {'signed': True}

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(srv, 'tool_sign_document', fake_tool_sign_document)

    srv.register_tools()
    registered_sign_document = captured_tool_functions['sign_document']
    registered_response = registered_sign_document(
        doc_id='doc-id',
        certificate_path=certificate_path,
        certificate_passphrase='certificate-secret',
        application_version='26.5.0-app',
        color_depth=32,
        horizontal_resolution=300,
        office_version='Office 2024',
        vertical_resolution=600,
        windows_version='Windows 11',
    )

    assert registered_response == {'signed': True}
    assert tool_sign_calls == [
        {
            'doc_id': 'doc-id',
            'certificate_path': certificate_path,
            'certificate_passphrase': 'certificate-secret',
            'application_version': '26.5.0-app',
            'color_depth': 32,
            'horizontal_resolution': 300,
            'office_version': 'Office 2024',
            'vertical_resolution': 600,
            'windows_version': 'Windows 11',
        }
    ]


def test_registered_get_digital_signatures_delegates_and_returns_response(
    monkeypatch,
):
    captured_tool_functions = {}
    tool_signature_calls = []
    expected_response = {
        'signatures': [
            {
                'application_version': '26.5.0-app',
                'color_depth': 32,
                'horizontal_resolution': 300,
                'office_version': 'Office 2024',
                'vertical_resolution': 600,
                'windows_version': 'Windows 11',
            }
        ]
    }

    class FakeMcp:
        def tool(self, description=None):
            def capture_tool(function_to_register):
                captured_tool_functions[function_to_register.__name__] = function_to_register
                return function_to_register

            return capture_tool

    def fake_tool_get_digital_signatures(doc_id: str) -> dict[str, object]:
        tool_signature_calls.append(doc_id)
        return expected_response

    monkeypatch.setattr(srv, 'mcp', FakeMcp())
    monkeypatch.setattr(
        srv,
        'tool_get_digital_signatures',
        fake_tool_get_digital_signatures,
    )

    srv.register_tools()
    registered_get_digital_signatures = captured_tool_functions['get_digital_signatures']
    registered_response = registered_get_digital_signatures('doc-id')

    assert registered_response == expected_response
    assert tool_signature_calls == ['doc-id']
