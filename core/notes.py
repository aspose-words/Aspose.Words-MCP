from __future__ import annotations
from typing import Any, Dict, List, Optional
import aspose.words as aw
from core.utils.docs_util import ensure_path, get_paragraph_nodes, find_paragraph_indices_by_anchor, normalize_text_for_match, safe_count, hex_to_color

def notes_add_footnote(doc_id: str, paragraph_index: int, text: str) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    builder.move_to_paragraph(paragraph_index, -1)
    builder.insert_footnote(aw.notes.FootnoteType.FOOTNOTE, text)
    doc.save(str(path))
    return True

def notes_add_endnote(doc_id: str, paragraph_index: int, text: str) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    builder.move_to_paragraph(paragraph_index, -1)
    builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, text)
    doc.save(str(path))
    return True

def notes_convert_footnotes_to_endnotes(doc_id: str) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    notes = doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    to_convert = []
    for i in range(notes.count):
        fn = notes[i]
        fno = fn.as_footnote()
        if fno.footnote_type == aw.notes.FootnoteType.FOOTNOTE:
            to_convert.append(fn)
    for fn in to_convert:
        fno = fn.as_footnote()
        txt = fno.to_string(aw.SaveFormat.TEXT) or ''
        anc = fn.get_ancestor(aw.NodeType.PARAGRAPH)
        if anc is not None:
            idx = paras.index_of(anc)
            if idx is not None and idx >= 0:
                builder.move_to_paragraph(int(idx), -1)
            else:
                builder.move_to_document_end()
        else:
            builder.move_to_document_end()
        builder.insert_footnote(aw.notes.FootnoteType.ENDNOTE, txt)
        fn.remove()
    doc.save(str(path))
    return True

def notes_style(doc_id: str, font_name: Optional[str]=None, font_size: Optional[float]=None, bold: Optional[bool]=None, italic: Optional[bool]=None, color_hex: Optional[str]=None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    notes = doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)
    for i in range(notes.count):
        n = notes[i]
        fn = n.as_footnote()
        paragraphs = fn.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        for pi in range(paragraphs.count):
            pobj = paragraphs[pi].as_paragraph()
            runs = pobj.runs
            for j in range(runs.count):
                run = runs[j]
                font = run.font
                if font_name:
                    font.name = font_name
                if font_size:
                    font.size = font_size
                if bold is not None:
                    font.bold = bold
                if italic is not None:
                    font.italic = italic
                col = hex_to_color(color_hex)
                if col is not None:
                    font.color = col
    doc.save(str(path))
    return True

def notes_list(doc_id: str) -> List[Dict[str, Any]]:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = get_paragraph_nodes(doc)
    notes = doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)
    out: List[Dict[str, Any]] = []
    total = notes.count
    for i in range(total):
        fn = notes[i]
        fno = fn.as_footnote()
        tname = 'endnote' if fno.footnote_type == aw.notes.FootnoteType.ENDNOTE else 'footnote'
        txt = normalize_text_for_match(fno.to_string(aw.SaveFormat.TEXT) or '')
        anc = fn.get_ancestor(aw.NodeType.PARAGRAPH)
        pidx = None
        if anc is not None:
            j = paras.index_of(anc)
            if j is not None and j >= 0:
                pidx = int(j)
        if pidx is None:
            pidx = max(0, paras.count - 1)
        out.append({'type': tname, 'paragraphIndex': int(pidx), 'text': txt})
    return out

def notes_add_by_anchor(doc_id: str, note_type: str, anchor_text: str, text: str, position: str='after', occurrence: int=1, match_case: bool=False, whole_word: bool=False) -> bool:
    path = ensure_path(doc_id)
    if not anchor_text:
        raise ValueError('anchor_text must be non-empty')
    doc = aw.Document(str(path))
    indices = find_paragraph_indices_by_anchor(doc, anchor_text, match_case=match_case, whole_word=whole_word)
    if not indices:
        raise ValueError('Anchor text not found')
    idx = int(occurrence or 1) - 1
    if idx < 0 or idx >= len(indices):
        raise IndexError('occurrence out of range for anchor matches')
    pidx = indices[idx]
    builder = aw.DocumentBuilder(doc)
    ntype = aw.notes.FootnoteType.FOOTNOTE if (note_type or '').lower() != 'endnote' else aw.notes.FootnoteType.ENDNOTE
    if position == 'before':
        builder.move_to_paragraph(pidx, 0)
    else:
        builder.move_to_paragraph(pidx, -1)
    builder.insert_footnote(ntype, text)
    doc.save(str(path))
    return True

def notes_delete_by_anchor(doc_id: str, note_type: str, anchor_text: str, match_case: bool=False, whole_word: bool=False, occurrence: Optional[int]=None, remove_all: bool=True) -> int:
    path = ensure_path(doc_id)
    if not anchor_text:
        return 0
    doc = aw.Document(str(path))
    indices = find_paragraph_indices_by_anchor(doc, anchor_text, match_case=match_case, whole_word=whole_word)
    if not indices:
        return 0
    if occurrence is not None:
        idx = int(occurrence) - 1
        if 0 <= idx < len(indices):
            target_indices = [indices[idx]]
        else:
            target_indices = []
    else:
        target_indices = indices if remove_all else [indices[0]]
    removed = 0
    paras = get_paragraph_nodes(doc)
    target_paras = set()
    for tidx in target_indices:
        target_paras.add(paras[tidx])
    all_notes = doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)
    total = safe_count(all_notes)
    want_end = (note_type or '').lower() == 'endnote'
    for i in range(total):
        fn = all_notes[i]
        fno = fn.as_footnote()
        is_end = fno.footnote_type == aw.notes.FootnoteType.ENDNOTE
        if want_end != is_end:
            continue
        anc = fn.get_ancestor(aw.NodeType.PARAGRAPH)
        if anc in target_paras:
            fn.remove()
            removed += 1
            if occurrence is not None:
                break
    doc.save(str(path))
    return int(removed)

def notes_validate_by_anchor(doc_id: str, note_type: str, anchor_text: str, min_count: int=1, match_case: bool=False, whole_word: bool=False) -> Dict[str, Any]:
    path = ensure_path(doc_id)
    if not anchor_text:
        return {'ok': False, 'count': 0}
    doc = aw.Document(str(path))
    indices = find_paragraph_indices_by_anchor(doc, anchor_text, match_case=match_case, whole_word=whole_word)
    if not indices:
        return {'ok': False, 'count': 0}
    paras = get_paragraph_nodes(doc)
    target = set((paras[i] for i in indices))
    all_notes = doc.get_child_nodes(aw.NodeType.FOOTNOTE, True)
    total = safe_count(all_notes)
    want_end = (note_type or '').lower() == 'endnote'
    cnt = 0
    for i in range(total):
        fn = all_notes[i]
        fno = fn.as_footnote()
        is_end = fno.footnote_type == aw.notes.FootnoteType.ENDNOTE
        if want_end != is_end:
            continue
        anc = fn.get_ancestor(aw.NodeType.PARAGRAPH)
        if anc in target:
            cnt += 1
    ok = cnt >= int(min_count or 1)
    return {'ok': bool(ok), 'count': int(cnt)}
