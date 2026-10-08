from __future__ import annotations

import datetime
import math
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import aspose.words as aw

from core.utils.docs_util import ensure_path, ensure_resources_dir


def with_svg_embed_options() -> Any:
    sso = aw.saving.SvgSaveOptions()
    sso.export_embedded_images = True
    ensure_resources_dir('svg', sso)
    return sso


def export_markdown(doc: Any) -> bytes:
    import tempfile as _tmp
    from pathlib import Path

    data: bytes = b''
    with _tmp.NamedTemporaryFile(suffix='.md', delete=True) as tf:
        tmp_path = Path(tf.name)
        doc.save(str(tmp_path), aw.SaveFormat.MARKDOWN)
        data = tmp_path.read_bytes()
    return data


def build_pdf_opts(options: Dict[str, Any]) -> Any:
    pdf_opts = aw.saving.PdfSaveOptions()
    comp = (options or {}).get('compliance')
    if comp:
        m = {
            'PDF_A1A': aw.saving.PdfCompliance.PDF_A1A,
            'PDF_A1B': aw.saving.PdfCompliance.PDF_A1B,
        }
        key = str(comp).upper()
        key_norm = key.replace('_A_', 'A')
        if key_norm in m:
            pdf_opts.compliance = m[key_norm]
    if (options or {}).get('generate_form_field_scripts'):
        pdf_opts.generate_form_field_scripts = True
    digital_signature = (options or {}).get('digital_signature')
    if digital_signature is not None:
        pdf_opts.digital_signature_details = build_pdf_digital_signature_details(digital_signature)
    return pdf_opts


def build_pdf_digital_signature_details(options: Dict[str, Any]) -> Any:
    if not isinstance(options, dict):
        raise ValueError('digital_signature must be an object')

    certificate_path = options.get('certificate_path')
    if certificate_path is None or not str(certificate_path).strip():
        raise ValueError('digital_signature.certificate_path must be non-empty')

    cert_path = Path(str(certificate_path).strip())
    if not cert_path.exists():
        raise FileNotFoundError(f'Certificate file not found: {certificate_path}')

    certificate_password = (
        '' if options.get('certificate_password') is None else str(options['certificate_password'])
    )
    holder = aw.digitalsignatures.CertificateHolder.create(str(cert_path), certificate_password)
    details = aw.saving.PdfDigitalSignatureDetails()
    details.certificate_holder = holder

    if options.get('reason') is not None:
        details.reason = str(options['reason'])
    if options.get('location') is not None:
        details.location = str(options['location'])
    if options.get('signature_date') is not None:
        details.signature_date = _parse_pdf_signature_date(options['signature_date'])
    if options.get('hash_algorithm') is not None:
        details.hash_algorithm = _pdf_signature_hash_algorithm(options['hash_algorithm'])
    if options.get('timestamp_settings') is not None:
        details.timestamp_settings = _build_pdf_signature_timestamp_settings(
            options['timestamp_settings']
        )

    return details


def _parse_pdf_signature_date(value: Any) -> datetime.datetime:
    text = str(value).strip()
    if not text:
        raise ValueError('digital_signature.signature_date must be a non-empty ISO 8601 string')
    normalized = text[:-1] + '+00:00' if text.endswith('Z') else text
    try:
        return datetime.datetime.fromisoformat(normalized)
    except ValueError as exc:
        raise ValueError(
            'digital_signature.signature_date must be a valid ISO 8601 string'
        ) from exc


def _pdf_signature_hash_algorithm(value: Any) -> Any:
    text = str(value).strip().lower().replace('-', '_')
    normalized = text.replace('sha_', 'sha', 1)
    algorithms = {
        'sha256': aw.saving.PdfDigitalSignatureHashAlgorithm.SHA256,
        'sha384': aw.saving.PdfDigitalSignatureHashAlgorithm.SHA384,
        'sha512': aw.saving.PdfDigitalSignatureHashAlgorithm.SHA512,
        'ripe_md160': aw.saving.PdfDigitalSignatureHashAlgorithm.RIPE_MD160,
    }
    if not normalized:
        raise ValueError('digital_signature.hash_algorithm must be a non-empty string')
    if normalized not in algorithms:
        raise ValueError(
            'Unsupported digital_signature.hash_algorithm: '
            f'{value}. Expected one of: sha256, sha384, sha512, ripe_md160.'
        )
    return algorithms[normalized]


def _build_pdf_signature_timestamp_settings(options: Any) -> Any:
    if not isinstance(options, dict):
        raise ValueError('digital_signature.timestamp_settings must be an object')

    server_url = options.get('server_url')
    if server_url is None or not str(server_url).strip():
        raise ValueError('digital_signature.timestamp_settings.server_url must be non-empty')

    user_name = '' if options.get('user_name') is None else str(options['user_name'])
    password = '' if options.get('password') is None else str(options['password'])
    if options.get('timeout') is None:
        return aw.saving.PdfDigitalSignatureTimestampSettings(
            str(server_url).strip(),
            user_name,
            password,
        )

    timeout = _parse_pdf_timestamp_timeout(options['timeout'])
    return aw.saving.PdfDigitalSignatureTimestampSettings(
        str(server_url).strip(),
        user_name,
        password,
        timeout,
    )


def _parse_pdf_timestamp_timeout(value: Any) -> datetime.timedelta:
    if isinstance(value, bool):
        raise ValueError('digital_signature.timestamp_settings.timeout must be a positive number')
    try:
        seconds = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            'digital_signature.timestamp_settings.timeout must be a positive number'
        ) from exc
    if not math.isfinite(seconds) or seconds <= 0:
        raise ValueError('digital_signature.timestamp_settings.timeout must be a positive number')
    return datetime.timedelta(seconds=seconds)


def build_html_opts(fmt_key: str, embed_resources: bool) -> Any:
    if fmt_key == 'html_fixed':
        opts_hf = aw.saving.HtmlFixedSaveOptions()
        ensure_resources_dir('html', opts_hf)
        return opts_hf
    if fmt_key == 'mhtml':
        opts = aw.saving.HtmlSaveOptions(aw.SaveFormat.MHTML)
    else:
        opts = aw.saving.HtmlSaveOptions()
    opts.export_images_as_base64 = bool(embed_resources)
    if not embed_resources:
        ensure_resources_dir('html', opts)
    return opts


def build_xps_opts(options: Dict[str, Any]) -> Any:
    xps_opts = aw.saving.XpsSaveOptions()
    comp = (options or {}).get('compression_level')
    if comp is None:
        return xps_opts
    key = str(comp).strip().lower().replace('-', '_')
    if not key:
        raise ValueError('XPS compression_level must be a non-empty string')
    levels = {
        'normal': aw.saving.CompressionLevel.NORMAL,
        'maximum': aw.saving.CompressionLevel.MAXIMUM,
        'fast': aw.saving.CompressionLevel.FAST,
        'super_fast': aw.saving.CompressionLevel.SUPER_FAST,
    }
    if key not in levels:
        raise ValueError(
            'Unsupported XPS compression_level: '
            f'{comp}. Expected one of: normal, maximum, fast, super_fast.'
        )
    xps_opts.compression_level = levels[key]
    return xps_opts


def export(doc_id: str, fmt: str = 'docx') -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'docx').lower()
    if fmt_l == 'pdf':
        save_format = aw.SaveFormat.PDF
        mime = 'application/pdf'
        ext = 'pdf'
    elif fmt_l == 'rtf':
        save_format = aw.SaveFormat.RTF
        mime = 'application/rtf'
        ext = 'rtf'
    else:
        save_format = aw.SaveFormat.DOCX
        mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        ext = 'docx'
    out = BytesIO()
    doc.save(out, save_format)
    data = out.getvalue()
    return data, mime, ext


def render_page(
    doc_id: str, page_index: int = 0, fmt: str = 'png', dpi: int = 150
) -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'png').lower()
    if fmt_l in ('jpeg', 'jpg'):
        save_format = aw.SaveFormat.JPEG
        mime = 'image/jpeg'
        ext = 'jpg'
    elif fmt_l == 'svg':
        save_format = aw.SaveFormat.SVG
        mime = 'image/svg+xml'
        ext = 'svg'
    elif fmt_l == 'tiff':
        save_format = aw.SaveFormat.TIFF
        mime = 'image/tiff'
        ext = 'tiff'
    elif fmt_l == 'png':
        save_format = aw.SaveFormat.PNG
        mime = 'image/png'
        ext = 'png'
    else:
        raise ValueError(f'Unsupported render format: {fmt}')
    single = doc.extract_pages(int(page_index), 1)
    out = BytesIO()
    if fmt_l == 'svg':
        sso = aw.saving.SvgSaveOptions()
        sso.export_embedded_images = True
        ensure_resources_dir('svg', sso)
        single.save(out, sso)
    else:
        iso = aw.saving.ImageSaveOptions(save_format)
        iso.horizontal_resolution = float(dpi)
        iso.vertical_resolution = float(dpi)
        single.save(out, iso)
    return out.getvalue(), mime, ext


def export_advanced(
    doc_id: str, fmt: str = 'docx', options: Optional[Dict[str, Any]] = None
) -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'docx').lower()
    opts = options or {}
    specs: Dict[str, Dict[str, Any]] = {
        'html': {
            'mime': 'text/html',
            'ext': 'html',
            'save_format': aw.SaveFormat.HTML,
            'builder': lambda: build_html_opts('html', bool(opts.get('embed_resources', True))),
        },
        'mhtml': {
            'mime': 'message/rfc822',
            'ext': 'mhtml',
            'save_format': aw.SaveFormat.MHTML,
            'builder': lambda: build_html_opts('mhtml', bool(opts.get('embed_resources', True))),
        },
        'html_fixed': {
            'mime': 'text/html',
            'ext': 'html',
            'save_format': aw.SaveFormat.HTML_FIXED,
            'builder': lambda: build_html_opts(
                'html_fixed', bool(opts.get('embed_resources', True))
            ),
        },
        'epub': {'mime': 'application/epub+zip', 'ext': 'epub', 'save_format': aw.SaveFormat.EPUB},
        'odt': {
            'mime': 'application/vnd.oasis.opendocument.text',
            'ext': 'odt',
            'save_format': aw.SaveFormat.ODT,
        },
        'md': {'mime': 'text/markdown', 'ext': 'md', 'custom': True},
        'markdown': {'mime': 'text/markdown', 'ext': 'md', 'custom': True},
        'svg': {
            'mime': 'image/svg+xml',
            'ext': 'svg',
            'save_format': aw.SaveFormat.SVG,
            'builder': lambda: with_svg_embed_options(),
        },
        'pdf': {
            'mime': 'application/pdf',
            'ext': 'pdf',
            'save_format': aw.SaveFormat.PDF,
            'builder': lambda: build_pdf_opts(opts),
        },
        'xps': {
            'mime': 'application/vnd.ms-xpsdocument',
            'ext': 'xps',
            'save_format': aw.SaveFormat.XPS,
            'builder': lambda: build_xps_opts(opts),
        },
        'docling': {
            'mime': 'application/json',
            'ext': 'json',
            'save_format': aw.SaveFormat.DOCLING,
            'builder': lambda: _build_docling_opts(),
        },
    }
    spec = specs.get(fmt_l)
    if not spec:
        raise ValueError(f'Unsupported export format: {fmt}')
    if fmt_l == 'pdf' and opts.get('enable_text_shaping') is True:
        doc.layout_options.enable_text_shaping = True
    if spec.get('custom'):
        data = export_markdown(doc)
        return data, spec['mime'], spec['ext']
    builder = spec.get('builder')
    save_opts = builder() if builder is not None else None
    out = BytesIO()
    if save_opts is not None:
        doc.save(out, save_opts)
    else:
        doc.save(out, spec['save_format'])
    return out.getvalue(), spec['mime'], spec['ext']


def _build_docling_opts() -> Any:
    docling_opts = aw.saving.DoclingSaveOptions()
    docling_opts.save_format = aw.SaveFormat.DOCLING
    return docling_opts
