from __future__ import annotations
from typing import Optional, Tuple, List
from pathlib import Path
import uuid
import aspose.words as aw
from core.utils.docs_util import docx_path, ensure_path, get_data_dir

def create_document(name: Optional[str]=None) -> Tuple[str, str]:
    if name is None:
        name = 'hello.docx'
    doc_id = str(uuid.uuid4())
    doc = aw.Document()
    file_path = docx_path(doc_id)
    doc.save(str(file_path))
    return doc_id, name

def import_from_file(filename: str) -> Tuple[str, str]:
    src_path = Path(str(filename))
    if not src_path.exists():
        raise FileNotFoundError(f'Source file not found: {filename}')
    doc = aw.Document(str(src_path))
    doc_id = str(uuid.uuid4())
    dst = docx_path(doc_id)
    doc.save(str(dst), aw.SaveFormat.DOCX)
    return doc_id, src_path.name

def copy(doc_id: str) -> str:
    src = ensure_path(str(doc_id))
    new_id = str(uuid.uuid4())
    dst = docx_path(new_id)
    doc = aw.Document(str(src))
    doc.save(str(dst), aw.SaveFormat.DOCX)
    return new_id

def save_as_new(doc_id: str, name: Optional[str]=None, fmt: str='docx') -> Tuple[str, str]:
    src_path = ensure_path(str(doc_id))
    new_id = str(uuid.uuid4())
    new_name = name or f"document.{fmt or 'docx'}"
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

def cleanup_data_dir() -> int:
    count = 0
    for p in get_data_dir().glob('*.docx'):
        p.unlink()
        count += 1
    return count

def merge(source_ids: List[str]) -> str:
    if not source_ids:
        raise ValueError('sourceIds must be non-empty')
    first_path = ensure_path(source_ids[0])
    result_doc = aw.Document(str(first_path))
    for sid in source_ids[1:]:
        p = ensure_path(sid)
        src = aw.Document(str(p))
        result_doc.append_document(src, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
    new_id = str(uuid.uuid4())
    dst = docx_path(new_id)
    result_doc.save(str(dst))
    return new_id
