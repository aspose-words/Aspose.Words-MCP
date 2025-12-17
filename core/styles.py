from __future__ import annotations
from typing import Optional
import aspose.words as aw
from core.utils.docs_util import ensure_path, hex_to_color

def create_style(doc_id: str, style_name: Optional[str]=None, base_style: Optional[str]=None, font_size: Optional[float]=None, font_name: Optional[str]=None, bold: Optional[bool]=None, italic: Optional[bool]=None, color_hex: Optional[str]=None) -> str:
    if not style_name:
        raise TypeError('style_name must be provided')
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    style = doc.styles.get_by_name(style_name)
    if style is None:
        style = doc.styles.add(aw.StyleType.PARAGRAPH, style_name)
    if base_style:
        style.base_style_name = base_style
    font = style.font
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
    return style_name
