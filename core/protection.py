from __future__ import annotations
from typing import Dict, List, Optional
import aspose.words as aw
from . import content as _content
from core.utils.docs_util import ensure_path

def protect(doc_id: str, password: Optional[str]=None) -> str:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    doc.protect(aw.ProtectionType.READ_ONLY, password)
    doc.save(str(path))
    return 'ReadOnly'

def unprotect(doc_id: str, password: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    doc.unprotect(password)
    doc.save(str(path))
    return True

def protect_restrict(doc_id: str, password: Optional[str]=None, ranges: Optional[List[Dict[str, int]]]=None) -> str:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    if ranges:
        paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        for i, r in enumerate(ranges):
            pidx = r.get('paragraphIndex')
            if pidx is not None and 0 <= pidx < paras.count:
                start = max(0, int(r.get('start', 0)))
                end = max(start, int(r.get('end', 0)))
                _content.format_range(doc_id, pidx, start, end)
    doc.protect(aw.ProtectionType.READ_ONLY, password)
    doc.save(str(path))
    return str(doc.protection_type)
