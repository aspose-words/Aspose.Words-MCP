from __future__ import annotations

import uuid
from pathlib import Path
from typing import List, Optional, Tuple

import aspose.words as aw

from core.utils.docs_util import docx_path, ensure_path, get_data_dir


def create_document(name: Optional[str] = None) -> Tuple[str, str]:
    if name is None:
        name = 'hello.docx'
    doc_id = uuid.uuid4().urn.removeprefix('urn:uuid:')
    doc = aw.Document()
    file_path = docx_path(doc_id)
    doc.save(str(file_path))
    return doc_id, name


def import_from_file(filename: str) -> Tuple[str, str]:
    src_path = Path(str(filename))
    if not src_path.exists():
        raise FileNotFoundError(f'Source file not found: {filename}')
    doc = aw.Document(str(src_path))
    doc_id = uuid.uuid4().urn.removeprefix('urn:uuid:')
    dst = docx_path(doc_id)
    doc.save(str(dst), aw.SaveFormat.DOCX)
    return doc_id, src_path.name


def copy(doc_id: str) -> str:
    src = ensure_path(str(doc_id))
    new_id = uuid.uuid4().urn.removeprefix('urn:uuid:')
    dst = docx_path(new_id)
    doc = aw.Document(str(src))
    doc.save(str(dst), aw.SaveFormat.DOCX)
    return new_id


def save_as_new(doc_id: str, name: Optional[str] = None, fmt: str = 'docx') -> Tuple[str, str]:
    src_path = ensure_path(str(doc_id))
    new_id = uuid.uuid4().urn.removeprefix('urn:uuid:')
    new_name = name or f'document.{fmt or "docx"}'
    doc = aw.Document(str(src_path))
    dst_path = docx_path(new_id)
    doc.save(str(dst_path), aw.SaveFormat.DOCX)
    return new_id, new_name


def get_document_path(doc_id: str) -> str:
    return str(ensure_path(str(doc_id)))


def get_document_bytes(doc_id: str) -> bytes:
    p = ensure_path(str(doc_id))
    return p.read_bytes()


def delete(doc_id: str) -> bool:
    path = ensure_path(str(doc_id))
    path.unlink()
    return True


def document_exists(doc_id: str) -> bool:
    return docx_path(str(doc_id)).exists()


def import_header_footer_node(
    source_doc_id: str,
    destination_doc_id: str,
    header_footer_type: str,
    resolve_theme_colors: bool = False,
) -> str:
    header_footer_types = {
        'header_primary': aw.HeaderFooterType.HEADER_PRIMARY,
        'header_first': aw.HeaderFooterType.HEADER_FIRST,
        'header_even': aw.HeaderFooterType.HEADER_EVEN,
        'footer_primary': aw.HeaderFooterType.FOOTER_PRIMARY,
        'footer_first': aw.HeaderFooterType.FOOTER_FIRST,
        'footer_even': aw.HeaderFooterType.FOOTER_EVEN,
    }
    mapped_header_footer_type = header_footer_types.get(header_footer_type)
    if mapped_header_footer_type is None:
        accepted_values = ', '.join(sorted(header_footer_types.keys()))
        raise ValueError(
            f"Unsupported header_footer_type '{header_footer_type}'. "
            f"Supported values: {accepted_values}"
        )

    source_path = ensure_path(str(source_doc_id))
    destination_path = ensure_path(str(destination_doc_id))
    source_document = aw.Document(str(source_path))
    destination_document = aw.Document(str(destination_path))
    source_node = source_document.first_section.headers_footers.get_by_header_footer_type(
        mapped_header_footer_type
    )
    if source_node is None:
        raise ValueError(f"Source document does not contain header/footer '{header_footer_type}'")

    options = aw.ImportFormatOptions()
    options.resolve_theme_colors = resolve_theme_colors
    imported_node = destination_document.import_node(
        src_node=source_node,
        is_import_children=True,
        import_format_mode=aw.ImportFormatMode.KEEP_SOURCE_FORMATTING,
        import_format_options=options,
    )
    destination_document.first_section.headers_footers.add(imported_node)
    destination_document.save(str(destination_path))
    return destination_doc_id


def cleanup_data_dir() -> int:
    count = 0
    for p in get_data_dir().glob('*.docx'):
        p.unlink()
        count += 1
    return count


def merge(source_ids: List[str], append_document_with_new_page: Optional[bool] = None) -> str:
    if not source_ids:
        raise ValueError('sourceIds must be non-empty')
    first_path = ensure_path(source_ids[0])
    result_doc = aw.Document(str(first_path))
    import_format_options: Optional[aw.ImportFormatOptions] = None
    if append_document_with_new_page is not None:
        import_format_options = aw.ImportFormatOptions()
        import_format_options.append_document_with_new_page = append_document_with_new_page
    for sid in source_ids[1:]:
        p = ensure_path(sid)
        src = aw.Document(str(p))
        if import_format_options is None:
            result_doc.append_document(src, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
            continue
        result_doc.append_document(
            src,
            aw.ImportFormatMode.KEEP_SOURCE_FORMATTING,
            import_format_options,
        )
    new_id = uuid.uuid4().urn.removeprefix('urn:uuid:')
    dst = docx_path(new_id)
    result_doc.save(str(dst))
    return new_id
