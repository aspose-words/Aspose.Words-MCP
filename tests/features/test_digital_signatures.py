import base64
import types
from pathlib import Path

import pytest

pytest.importorskip('aspose.words')
import mcp_server as srv
from core import signatures as _signatures


class FakeSignOptions:
    metadata_fields = {
        'application_version',
        'color_depth',
        'horizontal_resolution',
        'office_version',
        'vertical_resolution',
        'windows_version',
    }
    instances = []

    def __init__(self):
        self.metadata_assignments = {}
        self.application_version = None
        self.color_depth = None
        self.horizontal_resolution = None
        self.office_version = None
        self.vertical_resolution = None
        self.windows_version = None
        self.metadata_assignments.clear()
        FakeSignOptions.instances.append(self)

    def __setattr__(self, name, assigned_metadata):
        object.__setattr__(self, name, assigned_metadata)
        if name in self.metadata_fields and 'metadata_assignments' in self.__dict__:
            self.metadata_assignments[name] = assigned_metadata


def test_sign_document_forwards_all_metadata(monkeypatch, tmp_path):
    certificate_create_calls = []
    sign_calls = []
    cert_holder = object()
    source_path = Path(tmp_path / 'source.docx')
    destination_path = Path(tmp_path / 'destination.docx')

    def fake_create(cert_bytes, passphrase):
        certificate_create_calls.append({'cert_bytes': cert_bytes, 'passphrase': passphrase})
        return cert_holder

    def fake_sign(src_file_name, dst_file_name, cert_holder, sign_options, /):
        sign_calls.append(
            {
                'src_file_name': src_file_name,
                'dst_file_name': dst_file_name,
                'cert_holder': cert_holder,
                'sign_options': sign_options,
            }
        )

    FakeSignOptions.instances.clear()
    fake_digital_signatures = types.SimpleNamespace(
        SignOptions=FakeSignOptions,
        CertificateHolder=types.SimpleNamespace(create=fake_create),
        DigitalSignatureUtil=types.SimpleNamespace(sign=fake_sign),
    )
    monkeypatch.setattr(_signatures, 'ensure_path', lambda source_doc_id: source_path)
    monkeypatch.setattr(_signatures, 'docx_path', lambda destination_doc_id: destination_path)
    monkeypatch.setattr(_signatures.aw, 'digitalsignatures', fake_digital_signatures)

    expected_passphrase = 'test-passphrase'
    returned_doc_id = _signatures.sign_document(
        source_doc_id='source-id',
        destination_doc_id='destination-id',
        certificate_bytes=b'certificate-bytes',
        certificate_passphrase=expected_passphrase,
        application_version='Aspose.Words 26.5.0',
        color_depth=32,
        horizontal_resolution=300,
        office_version='16.0',
        vertical_resolution=300,
        windows_version='10.0',
    )

    assert returned_doc_id == 'destination-id'
    assert len(FakeSignOptions.instances) == 1
    sign_options = FakeSignOptions.instances[0]
    assert sign_options.metadata_assignments == {
        'application_version': 'Aspose.Words 26.5.0',
        'color_depth': 32,
        'horizontal_resolution': 300,
        'office_version': '16.0',
        'vertical_resolution': 300,
        'windows_version': '10.0',
    }
    assert sign_options.application_version == 'Aspose.Words 26.5.0'
    assert sign_options.color_depth == 32
    assert sign_options.horizontal_resolution == 300
    assert sign_options.office_version == '16.0'
    assert sign_options.vertical_resolution == 300
    assert sign_options.windows_version == '10.0'
    assert certificate_create_calls == [
        {'cert_bytes': b'certificate-bytes', 'passphrase': expected_passphrase}
    ]
    assert sign_calls == [
        {
            'src_file_name': str(source_path),
            'dst_file_name': str(destination_path),
            'cert_holder': cert_holder,
            'sign_options': sign_options,
        }
    ]


def test_sign_document_omitted_metadata_remains_unset(monkeypatch, tmp_path):
    certificate_create_calls = []
    sign_calls = []
    cert_holder = object()
    source_path = Path(tmp_path / 'source.docx')
    destination_path = Path(tmp_path / 'destination.docx')

    def fake_create(cert_bytes, passphrase):
        certificate_create_calls.append({'cert_bytes': cert_bytes, 'passphrase': passphrase})
        return cert_holder

    def fake_sign(src_file_name, dst_file_name, cert_holder, sign_options, /):
        sign_calls.append(
            {
                'src_file_name': src_file_name,
                'dst_file_name': dst_file_name,
                'cert_holder': cert_holder,
                'sign_options': sign_options,
            }
        )

    FakeSignOptions.instances.clear()
    fake_digital_signatures = types.SimpleNamespace(
        SignOptions=FakeSignOptions,
        CertificateHolder=types.SimpleNamespace(create=fake_create),
        DigitalSignatureUtil=types.SimpleNamespace(sign=fake_sign),
    )
    monkeypatch.setattr(_signatures, 'ensure_path', lambda source_doc_id: source_path)
    monkeypatch.setattr(_signatures, 'docx_path', lambda destination_doc_id: destination_path)
    monkeypatch.setattr(_signatures.aw, 'digitalsignatures', fake_digital_signatures)

    expected_passphrase = 'test-passphrase'
    returned_doc_id = _signatures.sign_document(
        source_doc_id='source-id',
        destination_doc_id='destination-id',
        certificate_bytes=b'certificate-bytes',
        certificate_passphrase=expected_passphrase,
    )

    assert returned_doc_id == 'destination-id'
    assert len(FakeSignOptions.instances) == 1
    sign_options = FakeSignOptions.instances[0]
    assert sign_options.metadata_assignments == {}
    assert sign_options.application_version is None
    assert sign_options.color_depth is None
    assert sign_options.horizontal_resolution is None
    assert sign_options.office_version is None
    assert sign_options.vertical_resolution is None
    assert sign_options.windows_version is None
    assert certificate_create_calls == [
        {'cert_bytes': b'certificate-bytes', 'passphrase': expected_passphrase}
    ]
    assert sign_calls == [
        {
            'src_file_name': str(source_path),
            'dst_file_name': str(destination_path),
            'cert_holder': cert_holder,
            'sign_options': sign_options,
        }
    ]


def test_tool_sign_document_decodes_certificate_and_forwards_metadata(monkeypatch):
    sign_calls = []

    def fake_sign_document(
        *,
        source_doc_id,
        destination_doc_id,
        certificate_bytes,
        certificate_passphrase,
        application_version=None,
        color_depth=None,
        horizontal_resolution=None,
        office_version=None,
        vertical_resolution=None,
        windows_version=None,
    ):
        sign_calls.append(
            {
                'source_doc_id': source_doc_id,
                'destination_doc_id': destination_doc_id,
                'certificate_bytes': certificate_bytes,
                'certificate_passphrase': certificate_passphrase,
                'application_version': application_version,
                'color_depth': color_depth,
                'horizontal_resolution': horizontal_resolution,
                'office_version': office_version,
                'vertical_resolution': vertical_resolution,
                'windows_version': windows_version,
            }
        )
        return destination_doc_id

    monkeypatch.setattr(srv._signatures, 'sign_document', fake_sign_document)

    certificate_bytes = b'certificate-bytes'
    certificate_base64 = base64.b64encode(certificate_bytes).decode('utf-8')
    response = srv.tool_sign_document(
        source_doc_id='source-id',
        destination_doc_id='destination-id',
        certificate_base64=certificate_base64,
        certificate_passphrase='test-passphrase',
        application_version='Aspose.Words 26.5.0',
        color_depth=32,
        horizontal_resolution=300,
        office_version='16.0',
        vertical_resolution=300,
        windows_version='10.0',
    )

    assert response == {'docId': 'destination-id'}
    assert sign_calls == [
        {
            'source_doc_id': 'source-id',
            'destination_doc_id': 'destination-id',
            'certificate_bytes': certificate_bytes,
            'certificate_passphrase': 'test-passphrase',
            'application_version': 'Aspose.Words 26.5.0',
            'color_depth': 32,
            'horizontal_resolution': 300,
            'office_version': '16.0',
            'vertical_resolution': 300,
            'windows_version': '10.0',
        }
    ]
