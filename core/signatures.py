from __future__ import annotations

from pathlib import Path
from typing import Optional

import aspose.words as aw

from core.utils.docs_util import ensure_path


def list_digital_signatures(doc_id: str) -> list[dict[str, object]]:
    source_path = ensure_path(doc_id)
    signatures = aw.digitalsignatures.DigitalSignatureUtil.load_signatures(str(source_path))

    return [
        {
            'application_version': signature.application_version,
            'color_depth': signature.color_depth,
            'horizontal_resolution': signature.horizontal_resolution,
            'office_version': signature.office_version,
            'vertical_resolution': signature.vertical_resolution,
            'windows_version': signature.windows_version,
        }
        for signature in signatures
    ]


def sign_document(
    doc_id: str,
    certificate_path: str,
    certificate_password: str = '',
    application_version: Optional[str] = None,
    color_depth: Optional[int] = None,
    horizontal_resolution: Optional[int] = None,
    office_version: Optional[str] = None,
    vertical_resolution: Optional[int] = None,
    windows_version: Optional[str] = None,
) -> bool:
    source_path = ensure_path(doc_id)
    if not certificate_path.strip():
        raise ValueError('certificate_path must be non-empty')

    cert_path = Path(certificate_path)
    if not cert_path.exists():
        raise FileNotFoundError(f'Certificate file not found: {certificate_path}')

    holder = aw.digitalsignatures.CertificateHolder.create(
        str(cert_path),
        certificate_password,
    )
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

    signed_path = source_path.with_name(f'{source_path.stem}.signed{source_path.suffix}')
    aw.digitalsignatures.DigitalSignatureUtil.sign(
        src_file_name=str(source_path),
        dst_file_name=str(signed_path),
        cert_holder=holder,
        sign_options=sign_options,
    )
    signed_path.replace(source_path)
    return True
