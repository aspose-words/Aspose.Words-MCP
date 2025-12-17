import logging
from pathlib import Path
from typing import Optional, List, Any

import aspose.words as aw
from aspose.pydrawing import Color

logger = logging.getLogger(__name__)
_DATA_DIR: Optional[Path] = None

def init_data_dir(data_dir: Path) -> None:
    global _DATA_DIR
    _DATA_DIR = data_dir
    _DATA_DIR.mkdir(parents=True, exist_ok=True)
    logger.info(f'Using data directory: {_DATA_DIR}')

def get_data_dir() -> Path:
    if _DATA_DIR is None:
        raise RuntimeError('Data directory is not initialized. Call init_data_dir first.')
    return _DATA_DIR

def docx_path(doc_id: str) -> Path:
    return get_data_dir() / f'{doc_id}.docx'

def ensure_path(doc_id: str) -> Path:
    p = docx_path(doc_id)
    if not p.exists():
        raise FileNotFoundError(f'Document {doc_id} not found')
    return p

def document_exists(doc_id: str) -> bool:
    return docx_path(doc_id).exists()

def hex_to_color(color_hex: Optional[str]):
    if not color_hex:
        return None
    color_hex = str(color_hex).strip().lstrip('#')
    if len(color_hex) != 6:
        return None
    hexdigits = '0123456789abcdefABCDEF'
    if any(ch not in hexdigits for ch in color_hex):
        return None
    r = int(color_hex[0:2], 16)
    g = int(color_hex[2:4], 16)
    b = int(color_hex[4:6], 16)
    return Color.from_argb(255, r, g, b)

def ensure_resources_dir(kind: str, opts_obj: Any) -> None:
    d = get_data_dir() / f'{kind}_resources'
    d.mkdir(exist_ok=True)
    opts_obj.resources_folder = str(d)

def move_builder(doc: aw.Document, builder: aw.DocumentBuilder, where: str, paragraph_index: Optional[int]) -> None:
    if where == 'start':
        builder.move_to_document_start()
    elif where == 'paragraph':
        if paragraph_index is None or paragraph_index < 0:
            raise ValueError("paragraph_index must be provided and >= 0 when where='paragraph'")
        paragraphs = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        if paragraph_index >= paragraphs.count:
            raise IndexError('paragraph_index out of range')
        target = paragraphs[paragraph_index]
        try:
            builder.move_to(target)
        except RuntimeError:
            parent = target.parent_node
            story_paras = parent.get_child_nodes(aw.NodeType.PARAGRAPH, False)
            local_idx = 0
            for i in range(int(story_paras.count)):
                if story_paras[i] is target:
                    local_idx = i
                    break
            builder.move_to_paragraph(local_idx, -1)
    else:
        builder.move_to_document_end()

def get_paragraph_nodes(doc: aw.Document):
    return doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)

def safe_count(obj: Any) -> int:
    return int(obj.count)

def normalize_text_for_match(text: str) -> str:
    t = (text or '').replace('\r', '').replace('\x07', '').replace('\x07', '').replace('\x0c', '').strip()
    return t

def find_paragraph_indices_by_anchor(doc: aw.Document, anchor_text: str, match_case: bool=False, whole_word: bool=False) -> List[int]:
    if not anchor_text:
        return []
    needle = anchor_text if match_case else anchor_text.lower()
    nodes = get_paragraph_nodes(doc)
    indices: List[int] = []
    for i in range(safe_count(nodes)):
        p = nodes[i]
        t = p.to_string(aw.SaveFormat.TEXT) or ''
        t = normalize_text_for_match(t)
        hay = t if match_case else t.lower()
        if whole_word:
            words = hay.split()
            if needle in words:
                indices.append(i)
        elif needle in hay:
            indices.append(i)
    return indices

def resolve_heading_style_identifier(level: int):
    _MAP = {1: aw.StyleIdentifier.HEADING1, 2: aw.StyleIdentifier.HEADING2, 3: aw.StyleIdentifier.HEADING3, 4: aw.StyleIdentifier.HEADING4, 5: aw.StyleIdentifier.HEADING5, 6: aw.StyleIdentifier.HEADING6, 7: aw.StyleIdentifier.HEADING7, 8: aw.StyleIdentifier.HEADING8, 9: aw.StyleIdentifier.HEADING9}
    return _MAP[level]

def resolve_outline_level(level: int):
    _MAP = {1: aw.OutlineLevel.LEVEL1, 2: aw.OutlineLevel.LEVEL2, 3: aw.OutlineLevel.LEVEL3, 4: aw.OutlineLevel.LEVEL4, 5: aw.OutlineLevel.LEVEL5, 6: aw.OutlineLevel.LEVEL6, 7: aw.OutlineLevel.LEVEL7, 8: aw.OutlineLevel.LEVEL8, 9: aw.OutlineLevel.LEVEL9}
    return _MAP[level]
