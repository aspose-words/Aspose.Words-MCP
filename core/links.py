from __future__ import annotations
import aspose.words as aw
from core.utils.docs_util import ensure_path


def _move_builder_to_paragraph(doc_id: str, paragraph_index: int):
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    builder = aw.DocumentBuilder(doc)
    builder.move_to_paragraph(paragraph_index, -1)
    return path, doc, builder

def add_bookmark_at_paragraph(doc_id: str, name: str, paragraph_index: int) -> bool:
    path, doc, builder = _move_builder_to_paragraph(doc_id, paragraph_index)
    builder.start_bookmark(name)
    builder.end_bookmark(name)
    doc.save(str(path))
    return True

def insert_hyperlink_at_paragraph(doc_id: str, paragraph_index: int, text: str, target: str, external: bool=True) -> bool:
    path, doc, builder = _move_builder_to_paragraph(doc_id, paragraph_index)
    if external:
        builder.insert_hyperlink(text, target, True)
    else:
        builder.insert_hyperlink(text, f'#{target}', False)
    doc.save(str(path))
    return True
