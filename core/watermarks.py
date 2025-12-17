from __future__ import annotations
from typing import Optional
from io import BytesIO
import aspose.words as aw
from core.utils.docs_util import ensure_path, hex_to_color

def add_watermark_text(doc_id: str, text: str, font_name: Optional[str]=None, size: Optional[float]=None, color_hex: Optional[str]=None, diagonal: bool=True) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    opts = aw.TextWatermarkOptions()
    if font_name:
        opts.font_family = font_name
    if size:
        opts.font_size = float(size)
    col = hex_to_color(color_hex)
    if col is not None:
        opts.color = col
    opts.layout = aw.WatermarkLayout.DIAGONAL if diagonal else aw.WatermarkLayout.HORIZONTAL
    doc.watermark.set_text(text or '', opts)
    doc.save(str(path))
    return True

def add_watermark_image(doc_id: str, image_bytes: bytes, scale: Optional[float]=None, washout: bool=True) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    with BytesIO(image_bytes) as stream:
        opts = aw.ImageWatermarkOptions()
        if scale is not None:
            opts.scale = float(scale)
        opts.is_washout = bool(washout)
        doc.watermark.set_image(stream, opts)
    doc.save(str(path))
    return True
