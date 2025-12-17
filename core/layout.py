from __future__ import annotations
from typing import Optional, Dict
import aspose.words as aw
from core.utils.docs_util import ensure_path

def header_add_text(doc_id: str, text: str, primary: bool=True) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.move_to_header_footer(aw.HeaderFooterType.HEADER_PRIMARY if primary else aw.HeaderFooterType.HEADER_FIRST)
    builder.writeln(text)
    doc.save(str(path))
    return True

def footer_add_text(doc_id: str, text: str, primary: bool=True) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY if primary else aw.HeaderFooterType.FOOTER_FIRST)
    builder.writeln(text)
    doc.save(str(path))
    return True

def add_page_numbering(doc_id: str, format_string: str='Page {PAGE} of {NUMPAGES}') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    builder.move_to_header_footer(aw.HeaderFooterType.FOOTER_PRIMARY)
    template = format_string or 'Page {PAGE} of {NUMPAGES}'
    parts = template.replace('{PAGE}', '\x01PAGE\x01').replace('{NUMPAGES}', '\x01NUMPAGES\x01')
    tokens = parts.split('\x01')
    for tok in tokens:
        if tok == 'PAGE':
            builder.insert_field('PAGE')
        elif tok == 'NUMPAGES':
            builder.insert_field('NUMPAGES')
        elif tok:
            builder.write(tok)
    doc.save(str(path))
    return True

def set_different_first_page_header_footer(doc_id: str, enabled: bool=True) -> bool:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    for i in range(doc.sections.count):
        sec = doc.sections[i]
        sec.page_setup.different_first_page_header_footer = bool(enabled)
    doc.save(str(file_path))
    return True

def set_page_setup(doc_id: str, margins: Optional[Dict[str, float]]=None, orientation: Optional[str]=None, paper: Optional[str]=None, columns: Optional[int]=None) -> bool:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    for i in range(doc.sections.count):
        sec = doc.sections[i]
        ps = sec.page_setup
        if margins:
            if 'top' in margins:
                ps.top_margin = float(margins['top'])
            if 'bottom' in margins:
                ps.bottom_margin = float(margins['bottom'])
            if 'left' in margins:
                ps.left_margin = float(margins['left'])
            if 'right' in margins:
                ps.right_margin = float(margins['right'])
        if orientation:
            o = orientation.lower()
            ps.orientation = aw.Orientation.LANDSCAPE if o.startswith('land') else aw.Orientation.PORTRAIT
        if paper:
            ps.paper_size = getattr(aw.PaperSize, str(paper).upper())
        if columns is not None:
            ps.text_columns.set_count(int(columns))
    doc.save(str(file_path))
    return True

def insert_section_break(doc_id: str, kind: str='nextPage') -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    mapping = {'nextpage': aw.BreakType.SECTION_BREAK_NEW_PAGE, 'continuous': aw.BreakType.SECTION_BREAK_CONTINUOUS, 'even': aw.BreakType.SECTION_BREAK_EVEN_PAGE, 'odd': aw.BreakType.SECTION_BREAK_ODD_PAGE}
    key = (kind or 'nextPage').replace('_', '').lower()
    btype = mapping.get(key, aw.BreakType.SECTION_BREAK_NEW_PAGE)
    builder.insert_break(btype)
    doc.save(str(path))
    return True
