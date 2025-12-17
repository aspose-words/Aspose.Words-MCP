from __future__ import annotations
from io import BytesIO
from typing import Dict, Optional, Tuple, Any
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
    with _tmp.NamedTemporaryFile(suffix='.md', delete=True) as tf:
        tmp_path = Path(tf.name)
        doc.save(str(tmp_path), aw.SaveFormat.MARKDOWN)
        return tmp_path.read_bytes()

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
    return pdf_opts

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

def export(doc_id: str, fmt: str='docx') -> Tuple[bytes, str, str]:
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

def render_page(doc_id: str, page_index: int=0, fmt: str='png', dpi: int=150) -> Tuple[bytes, str, str]:
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

def export_advanced(doc_id: str, fmt: str='docx', options: Optional[Dict[str, Any]]=None) -> Tuple[bytes, str, str]:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    fmt_l = (fmt or 'docx').lower()
    opts = options or {}
    specs: Dict[str, Dict[str, Any]] = {'html': {'mime': 'text/html', 'ext': 'html', 'save_format': aw.SaveFormat.HTML, 'builder': lambda: build_html_opts('html', bool(opts.get('embed_resources', True)))}, 'mhtml': {'mime': 'message/rfc822', 'ext': 'mhtml', 'save_format': aw.SaveFormat.MHTML, 'builder': lambda: build_html_opts('mhtml', bool(opts.get('embed_resources', True)))}, 'html_fixed': {'mime': 'text/html', 'ext': 'html', 'save_format': aw.SaveFormat.HTML_FIXED, 'builder': lambda: build_html_opts('html_fixed', bool(opts.get('embed_resources', True)))}, 'epub': {'mime': 'application/epub+zip', 'ext': 'epub', 'save_format': aw.SaveFormat.EPUB}, 'odt': {'mime': 'application/vnd.oasis.opendocument.text', 'ext': 'odt', 'save_format': aw.SaveFormat.ODT}, 'md': {'mime': 'text/markdown', 'ext': 'md', 'custom': True}, 'markdown': {'mime': 'text/markdown', 'ext': 'md', 'custom': True}, 'svg': {'mime': 'image/svg+xml', 'ext': 'svg', 'save_format': aw.SaveFormat.SVG, 'builder': lambda: with_svg_embed_options()}, 'pdf': {'mime': 'application/pdf', 'ext': 'pdf', 'save_format': aw.SaveFormat.PDF, 'builder': lambda: build_pdf_opts(opts)}}
    spec = specs.get(fmt_l)
    if not spec:
        raise ValueError(f'Unsupported export format: {fmt}')
    if spec.get('custom'):
        data = export_markdown(doc)
        return data, spec['mime'], spec['ext']
    save_opts = spec.get('builder')() if spec.get('builder') else None
    out = BytesIO()
    if save_opts is not None:
        doc.save(out, save_opts)
    else:
        doc.save(out, spec['save_format'])
    return out.getvalue(), spec['mime'], spec['ext']
