from typing import Any, Dict, List, Optional
from io import BytesIO
from pathlib import Path
import aspose.words as aw

def _docs():
    from core.utils import docs_util as _d
    return _d

def read_paragraphs(doc_id: str, start: Optional[int]=None, end: Optional[int]=None) -> List[str]:
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    nodes = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    texts: List[str] = []
    for i in range(nodes.count):
        para = nodes[i]
        t = para.to_string(aw.SaveFormat.TEXT) or ''
        texts.append(t)
    s = start or 0
    e = end if end is not None else len(texts)
    s = max(0, min(s, len(texts)))
    e = max(s, min(e, len(texts)))
    return texts[s:e]

def find_text(doc_id: str, text: str='', match_case: bool=False, whole_word: bool=False):
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    matches: List[int] = []
    query = text or ''
    if not query:
        return matches
    for i in range(paras.count):
        p = paras[i]
        t = p.to_string(aw.SaveFormat.TEXT) or ''
        a = _docs().normalize_text_for_match(t)
        b = _docs().normalize_text_for_match(query)
        if match_case:
            a = t
            b = query
        if whole_word:
            words = [w for w in (a or '').replace('\r', ' ').split() if w]
            if b in words:
                matches.append(i)
        elif b in a:
            matches.append(i)
    return matches

def get_text(doc_id: str) -> str:
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    return doc.get_text()

def get_xml(doc_id: str) -> str:
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    buf = BytesIO()
    doc.save(buf, aw.SaveFormat.WORD_ML)
    return buf.getvalue().decode('utf-8')

def get_outline(doc_id: str):
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    doc.update_page_layout()
    doc.update_fields()
    buf = BytesIO()
    doc.save(buf, aw.SaveFormat.DOCX)
    buf.seek(0)
    doc = aw.Document(buf)
    nodes = _docs().get_paragraph_nodes(doc)
    items: List[Dict[str, Any]] = []
    heading_map: Dict[Any, int] = {aw.StyleIdentifier.HEADING1: 1, aw.StyleIdentifier.HEADING2: 2, aw.StyleIdentifier.HEADING3: 3, aw.StyleIdentifier.HEADING4: 4, aw.StyleIdentifier.HEADING5: 5, aw.StyleIdentifier.HEADING6: 6, aw.StyleIdentifier.HEADING7: 7, aw.StyleIdentifier.HEADING8: 8, aw.StyleIdentifier.HEADING9: 9}
    for i in range(nodes.count):
        p = nodes[i]
        pobj = p.as_paragraph()
        sid = pobj.paragraph_format.style_identifier
        level = heading_map.get(sid)
        if level is None:
            sname = pobj.paragraph_format.style.name
            if isinstance(sname, str):
                low = sname.lower()
                lvl_tok = None
                for tok in low.replace('_', ' ').split():
                    if tok.isdigit():
                        lvl_tok = int(tok)
                        break
                if lvl_tok:
                    level = lvl_tok
        if level is None:
            ol = pobj.paragraph_format.outline_level
            if ol != aw.OutlineLevel.BODY_TEXT:
                for lvl in range(1, 10):
                    resolved = _docs().resolve_outline_level(lvl)
                    if ol == resolved:
                        level = lvl
                        break
        if level:
            t = p.to_string(aw.SaveFormat.TEXT) or ''
            t = t.replace('\r', '').strip()
            items.append({'text': t, 'level': int(level), 'paragraphIndex': i})
    return items

def get_info(doc_id: str) -> Dict[str, Any]:
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    text = doc.get_text()
    words = len([w for w in (text or '').split() if w])
    paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True).count
    pages = doc.page_count
    size = path.stat().st_size
    return {'sizeBytes': int(size), 'words': int(words), 'paragraphs': int(paragraphs), 'pages': pages}

def stats(doc_id: str) -> Dict[str, int]:
    path = _docs().ensure_path(doc_id)
    doc = aw.Document(str(path))
    text = doc.get_text()
    words = len([w for w in (text or '').split() if w])
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True).count
    pages = doc.page_count
    return {'words': int(words), 'paragraphs': int(paras), 'pages': int(pages)}

def list_documents() -> list:
    ids: List[str] = []
    data_dir = Path(_docs().get_data_dir())
    for p in data_dir.glob('*.docx'):
        ids.append(p.stem)
    return ids
