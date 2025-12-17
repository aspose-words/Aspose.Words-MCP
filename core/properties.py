from __future__ import annotations
from typing import Dict, Optional
import aspose.words as aw
from core.utils.docs_util import ensure_path

def properties_get(doc_id: str) -> Dict[str, Optional[str]]:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    props = doc.built_in_document_properties
    return {
        'title': str(props.title) if props.title is not None else None,
        'author': str(props.author) if props.author is not None else None,
        'subject': str(props.subject) if props.subject is not None else None,
        'keywords': str(props.keywords) if props.keywords is not None else None,
    }

def properties_set(doc_id: str, title: Optional[str]=None, author: Optional[str]=None, subject: Optional[str]=None, keywords: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    props = doc.built_in_document_properties
    if title is not None:
        props.title = title
    if author is not None:
        props.author = author
    if subject is not None:
        props.subject = subject
    if keywords is not None:
        props.keywords = keywords
    doc.save(str(path))
    return True
