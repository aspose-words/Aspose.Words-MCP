from __future__ import annotations
from typing import Dict, List, Optional
import aspose.words as aw
from core.utils.docs_util import docx_path, get_paragraph_nodes

def get_comments(doc_id: str) -> List[Dict[str, Optional[str]]]:
    path = docx_path(doc_id)
    if not path.exists():
        raise FileNotFoundError(f'Document {doc_id} not found')
    doc = aw.Document(str(path))
    nodes = doc.get_child_nodes(aw.NodeType.COMMENT, True)
    paras = get_paragraph_nodes(doc)
    res: List[Dict[str, Optional[str]]] = []
    for i in range(nodes.count):
        c = nodes[i].as_comment()
        if c.ancestor is not None:
            continue
        author = c.author
        dt = c.date_time_utc
        text = c.to_string(aw.SaveFormat.TEXT)
        par = c.get_ancestor(aw.NodeType.PARAGRAPH)
        pidx = paras.index_of(par) if par is not None else None
        res.append({
            'author': author,
            'text': text if text is not None else '',
            'date': dt.isoformat() if dt else None,
            'paragraphIndex': pidx,
        })
    return res

def get_comments_by_author(doc_id: str, author: str) -> List[Dict[str, Optional[str]]]:
    items = get_comments(doc_id)
    return [x for x in items if (x.get('author') or '') == author]

def get_comments_for_paragraph(doc_id: str, paragraph_index: int) -> List[Dict[str, Optional[str]]]:
    items = get_comments(doc_id)
    return [x for x in items if x.get('paragraphIndex') == paragraph_index]
