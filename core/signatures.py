from __future__ import annotations

from typing import Optional

import aspose.words as aw

from core.utils.docs_util import docx_path, ensure_path


def sign_document(
    source_doc_id: str,
    destination_doc_id: str,
    certificate_bytes: bytes,
    certificate_passphrase: str,
    application_version: Optional[str] = None,
    color_depth: Optional[int] = None,
    horizontal_resolution: Optional[int] = None,
    office_version: Optional[str] = None,
    vertical_resolution: Optional[int] = None,
    windows_version: Optional[str] = None,
) -> str:
    source_path = ensure_path(str(source_doc_id))
    destination_path = docx_path(str(destination_doc_id))
    sign_options = aw.digitalsignatures.SignOptions()

    if application_version is not None:
        sign_options.application_version = application_version
    if color_depth is not None:
        sign_options.color_depth = color_depth
    if horizontal_resolution is not None:
        sign_options.horizontal_resolution = horizontal_resolution
    if office_version is not None:
        sign_options.office_version = office_version
    if vertical_resolution is not None:
        sign_options.vertical_resolution = vertical_resolution
    if windows_version is not None:
        sign_options.windows_version = windows_version

    cert_holder = aw.digitalsignatures.CertificateHolder.create(
        certificate_bytes,
        certificate_passphrase,
    )
    aw.digitalsignatures.DigitalSignatureUtil.sign(
        str(source_path),
        str(destination_path),
        cert_holder,
        sign_options,
    )
    return destination_doc_id
