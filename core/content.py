from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from typing import List, Optional

import aspose.words as aw

from core.utils.docs_util import (
    ensure_path,
    find_paragraph_indices_by_anchor,
    hex_to_color,
    move_builder,
    resolve_heading_style_identifier,
    resolve_outline_level,
)

_MAX_REGEX_PATTERN_LENGTH = 256
_CUSTOM_NODE_IDS_SIDECAR_SUFFIX = '.custom_node_ids.json'
_ASPOSE_EVALUATION_EXACT_PARAGRAPHS = {
    'evaluation only. created with aspose.words. copyright 2003-2026 aspose pty ltd.',
    'this document was truncated here because it was created in the evaluation mode.',
}
_ASPOSE_EVALUATION_FOOTER_PREFIX = (
    'created with an evaluation copy of aspose.words. to remove all limitations, '
    'you can use free temporary license '
)
_ASPOSE_EVALUATION_FOOTER_SUFFIX = 'https://products.aspose.com/words/temporary-license/'
_ALLOWED_REGEX_ESCAPES = {
    'd',
    'D',
    's',
    'S',
    'w',
    'W',
    't',
    'n',
    'r',
    '\\',
    '.',
    '^',
    '$',
    '|',
    '-',
    '[',
    ']',
    '{',
    '}',
    '(',
    ')',
    '*',
    '+',
    '?',
}


def _custom_node_ids_sidecar_path(doc_path: Path) -> Path:
    return doc_path.with_suffix(f'{doc_path.suffix}{_CUSTOM_NODE_IDS_SIDECAR_SUFFIX}')


def _normalize_paragraph_text(paragraph_text: str) -> str:
    return paragraph_text.rstrip('\r\n')


def _is_evaluation_only_paragraph(paragraph_text: str) -> bool:
    normalized_text = _normalize_paragraph_text(paragraph_text).strip().lower()
    if normalized_text == '':
        return False
    if normalized_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS:
        return True
    return normalized_text.startswith(
        _ASPOSE_EVALUATION_FOOTER_PREFIX
    ) and normalized_text.endswith(_ASPOSE_EVALUATION_FOOTER_SUFFIX)


def _evaluation_paragraph_indices(paragraph_texts: list[str]) -> set[int]:
    normalized_texts = [
        _normalize_paragraph_text(paragraph_text).strip().lower()
        for paragraph_text in paragraph_texts
    ]
    evaluation_indices: set[int] = set()

    leading_index: Optional[int] = None
    for index, normalized_text in enumerate(normalized_texts):
        if normalized_text != '':
            leading_index = index
            break

    if leading_index is not None:
        leading_text = normalized_texts[leading_index]
        if leading_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS:
            evaluation_indices.add(leading_index)

    trailing_index = len(normalized_texts) - 1
    while trailing_index >= 0:
        trailing_text = normalized_texts[trailing_index]
        if trailing_text == '':
            trailing_index -= 1
            continue

        is_evaluation_footer = trailing_text.startswith(
            _ASPOSE_EVALUATION_FOOTER_PREFIX
        ) and trailing_text.endswith(_ASPOSE_EVALUATION_FOOTER_SUFFIX)
        if trailing_text in _ASPOSE_EVALUATION_EXACT_PARAGRAPHS or is_evaluation_footer:
            evaluation_indices.add(trailing_index)
            trailing_index -= 1
            continue

        break

    return evaluation_indices


def _find_user_paragraph_ordinal(doc: aw.Document, paragraph_node: aw.Paragraph) -> int:
    paragraph_nodes = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    paragraph_index = paragraph_nodes.index_of(paragraph_node)
    if paragraph_index is None or paragraph_index < 0:
        raise ValueError(
            'Unable to resolve created paragraph ordinal for custom_node_id persistence'
        )

    paragraph_texts: list[str] = []
    for index in range(paragraph_nodes.count):
        paragraph = paragraph_nodes[index].as_paragraph()
        paragraph_texts.append(paragraph.to_string(aw.SaveFormat.TEXT) or '')

    evaluation_indices = _evaluation_paragraph_indices(paragraph_texts)

    user_paragraph_ordinal = 0
    for index in range(int(paragraph_index) + 1):
        if index in evaluation_indices:
            continue
        user_paragraph_ordinal += 1

    return user_paragraph_ordinal - 1


def _build_custom_node_id_sidecar_entry(
    doc: aw.Document,
    paragraph_node: aw.Paragraph,
    paragraph_text: str,
    custom_node_id: int,
) -> dict[str, int | str]:
    user_paragraph_ordinal = _find_user_paragraph_ordinal(doc, paragraph_node)
    sidecar_entry: dict[str, int | str] = {
        'custom_node_id': int(custom_node_id),
        'user_paragraph_ordinal': int(user_paragraph_ordinal),
    }

    normalized_text = _normalize_paragraph_text(paragraph_text)
    if normalized_text != '':
        sidecar_entry['normalized_text'] = normalized_text

    return sidecar_entry


def _persist_custom_node_id_sidecar(
    doc_path: Path, doc_id: str, sidecar_entry: dict[str, int | str]
) -> None:
    sidecar_path = _custom_node_ids_sidecar_path(doc_path)
    sidecar_data: dict[str, object]
    if sidecar_path.exists():
        sidecar_data = json.loads(sidecar_path.read_text(encoding='utf-8'))
    else:
        sidecar_data = {}

    paragraph_custom_node_ids = sidecar_data.get('paragraph_custom_node_ids')
    if not isinstance(paragraph_custom_node_ids, dict):
        paragraph_custom_node_ids = {}

    user_paragraph_ordinal = int(sidecar_entry['user_paragraph_ordinal'])
    paragraph_custom_node_ids[str(user_paragraph_ordinal)] = sidecar_entry

    sidecar_payload = {
        'doc_id': doc_id,
        'paragraph_custom_node_ids': paragraph_custom_node_ids,
    }
    sidecar_path.write_text(
        json.dumps(sidecar_payload, ensure_ascii=False, indent=2, sort_keys=True),
        encoding='utf-8',
    )


def validate_regex_pattern(pattern: str) -> None:
    if pattern == '':
        raise ValueError('search text must be a non-empty string')
    if len(pattern) > _MAX_REGEX_PATTERN_LENGTH:
        raise ValueError(
            f'regex pattern is too long; maximum length is {_MAX_REGEX_PATTERN_LENGTH} characters'
        )

    _reject_lookaround_and_group_operators(pattern)
    _reject_backreferences(pattern)
    _validate_allowed_subset(pattern)
    _reject_stacked_quantifiers(pattern)


def _reject_lookaround_and_group_operators(pattern: str) -> None:
    escaped = False
    for index, char in enumerate(pattern):
        if escaped:
            escaped = False
            continue
        if char == '\\':
            escaped = True
            continue
        if char == '(':
            if index + 1 < len(pattern) and pattern[index + 1] == '?':
                raise ValueError('regex lookarounds and inline group operators are not allowed')
            raise ValueError('regex groups are not allowed in public regex mode')


def _reject_backreferences(pattern: str) -> None:
    escaped = False
    for index, char in enumerate(pattern):
        if escaped:
            if char.isdigit():
                raise ValueError('regex backreferences are not allowed in public regex mode')
            escaped = False
            continue
        if char == '\\':
            escaped = True
            continue
        if char == '$' and index + 1 < len(pattern) and pattern[index + 1].isdigit():
            raise ValueError('regex backreferences are not allowed in public regex mode')


def _validate_allowed_subset(pattern: str) -> None:
    in_char_class = False
    escaped = False
    for char in pattern:
        if escaped:
            if not _is_allowed_escape(char):
                raise ValueError(f'regex escape sequence \\{char} is not allowed')
            escaped = False
            continue

        if char == '\\':
            escaped = True
            continue

        if char == '[':
            in_char_class = True
            continue
        if char == ']':
            in_char_class = False
            continue

        if in_char_class:
            continue

        if not _is_allowed_regex_char(char):
            raise ValueError(f'regex token {char!r} is not allowed in public regex mode')

    if escaped:
        raise ValueError('regex pattern cannot end with a dangling escape')
    if in_char_class:
        raise ValueError('regex character class is not closed')


def _is_allowed_escape(char: str) -> bool:
    return char in _ALLOWED_REGEX_ESCAPES


def _is_allowed_regex_char(char: str) -> bool:
    return char.isalnum() or char.isspace() or char in '._^$|*+?{}[]-,'


def _reject_stacked_quantifiers(pattern: str) -> None:
    in_char_class = False
    escaped = False
    previous_was_quantifier = False
    index = 0
    while index < len(pattern):
        char = pattern[index]

        if escaped:
            escaped = False
            previous_was_quantifier = False
            index += 1
            continue
        if char == '\\':
            escaped = True
            index += 1
            continue

        if char == '[':
            in_char_class = True
            previous_was_quantifier = False
            index += 1
            continue
        if char == ']':
            in_char_class = False
            previous_was_quantifier = False
            index += 1
            continue

        if in_char_class:
            index += 1
            continue

        if char in '*+?':
            if previous_was_quantifier:
                raise ValueError('nested or stacked regex quantifiers are not allowed')
            previous_was_quantifier = True
            index += 1
            continue

        if char == '{':
            closing_index = pattern.find('}', index + 1)
            if closing_index == -1:
                raise ValueError('regex quantifier is not closed')
            quantifier_body = pattern[index + 1 : closing_index]
            if not quantifier_body or not _is_valid_quantifier_body(quantifier_body):
                raise ValueError('regex quantifier must use only digits and comma')
            if previous_was_quantifier:
                raise ValueError('nested or stacked regex quantifiers are not allowed')
            previous_was_quantifier = True
            index = closing_index + 1
            continue

        previous_was_quantifier = False
        index += 1


def _is_valid_quantifier_body(body: str) -> bool:
    for char in body:
        if not (char.isdigit() or char == ','):
            return False
    return True


def find_heading_style_by_name(doc: aw.Document, level: int):
    lvl = level
    preferred = None
    fallback = None
    for s in doc.styles:
        if s.type != aw.StyleType.PARAGRAPH:
            continue
        name = (s.name or '').strip()
        low = name.lower()
        digit = None
        for tok in low.replace('_', ' ').split():
            if tok.isdigit():
                digit = int(tok)
                break
        if digit == lvl:
            if 'heading' in low:
                preferred = s
                break
            if fallback is None:
                fallback = s
    return preferred or fallback


def get_heading_style_object(doc: aw.Document, level: int):
    sid = resolve_heading_style_identifier(level)
    style_obj = doc.styles.get_by_style_identifier(sid)
    if style_obj is None:
        style_obj = doc.styles.get_by_name(f'Heading {level}')
        if style_obj is None:
            style_obj = find_heading_style_by_name(doc, level)
    return style_obj


def insert_text(
    doc_id: str, text: str = '', where: str = 'end', paragraph_index: Optional[int] = None
) -> bool:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.write(text)
    doc.save(str(file_path))
    return True


def replace_text(
    doc_id: str,
    search: str = '',
    replace: str = '',
    replace_all: bool = True,
    case_sensitive: bool = False,
    whole_word: bool = False,
    use_regex: bool = False,
    search_text: Optional[str] = None,
    replacement_text: Optional[str] = None,
    join_runs: bool = False,
    ignore_redundant: Optional[bool] = None,
    ignore_insignificant: Optional[bool] = None,
    ignore_spacing: Optional[bool] = None,
) -> int:
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))

    if search_text is not None and search != '':
        raise ValueError('Provide either "search" or "search_text", not both')
    if replacement_text is not None and replace != '':
        raise ValueError('Provide either "replace" or "replacement_text", not both')

    resolved_search = search_text if search_text is not None else search
    resolved_replace = replacement_text if replacement_text is not None else replace

    if resolved_search == '':
        raise ValueError('search text must be a non-empty string')

    options = aw.replacing.FindReplaceOptions()
    options.match_case = bool(case_sensitive)
    options.direction = aw.replacing.FindReplaceDirection.FORWARD
    options.find_whole_words_only = bool(whole_word)
    if not replace_all:
        options.max_matches = 1

    if use_regex:
        validate_regex_pattern(resolved_search)
        count = doc.range.replace_regex(resolved_search, resolved_replace, options)
    else:
        count = doc.range.replace(resolved_search, resolved_replace, options)

    if count == 0:
        return 0

    if join_runs:
        join_runs_options = aw.JoinRunsOptions()
        if ignore_redundant is not None:
            join_runs_options.ignore_redundant = bool(ignore_redundant)
        if ignore_insignificant is not None:
            join_runs_options.ignore_insignificant = bool(ignore_insignificant)
        if ignore_spacing is not None:
            join_runs_options.ignore_spacing = bool(ignore_spacing)

        paragraph_nodes = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
        for index in range(paragraph_nodes.count):
            paragraph = paragraph_nodes[index].as_paragraph()
            paragraph.join_runs_with_same_formatting(join_runs_options)

    doc.save(str(file_path))
    return int(count)


def read_paragraphs(doc_id: str, start: Optional[int] = None, end: Optional[int] = None):
    file_path = ensure_path(doc_id)
    doc = aw.Document(str(file_path))
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


def insert_image(
    doc_id: str,
    image_bytes: bytes = b'',
    where: str = 'end',
    width_points: Optional[float] = None,
    height_points: Optional[float] = None,
    keep_aspect: bool = False,
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, None)
    with BytesIO(image_bytes) as buf:
        if width_points is not None and height_points is not None:
            builder.insert_image(buf, float(width_points), float(height_points))
        else:
            shape = builder.insert_image(buf)
            if keep_aspect and (width_points is not None) != (height_points is not None):
                if width_points is not None:
                    ratio = shape.height / shape.width if shape.width else 1.0
                    shape.width = float(width_points)
                    shape.height = float(width_points) * ratio
                else:
                    ratio = shape.width / shape.height if shape.height else 1.0
                    shape.height = float(height_points)
                    shape.width = float(height_points) * ratio
            else:
                if width_points is not None:
                    shape.width = float(width_points)
                if height_points is not None:
                    shape.height = float(height_points)
    doc.save(str(path))
    return True


def insert_html(
    doc_id: str, html: str = '', where: str = 'end', paragraph_index: Optional[int] = None
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_html(html or '')
    doc.save(str(path))
    return True


def insert_markdown(
    doc_id: str, markdown: str = '', where: str = 'end', paragraph_index: Optional[int] = None
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.write(markdown or '')
    doc.save(str(path))
    return True


def add_heading(
    doc_id: str,
    text: str = '',
    level: int = 1,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    border_bottom: Optional[bool] = None,
    where: str = 'end',
    paragraph_index: Optional[int] = None,
    custom_node_id: Optional[int] = None,
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_paragraph()
    paragraph_with_written_text = builder.current_paragraph
    style_id = resolve_heading_style_identifier(level)
    builder.paragraph_format.style_identifier = style_id
    builder.paragraph_format.outline_level = resolve_outline_level(level)
    if font_name:
        builder.font.name = font_name
    if font_size:
        builder.font.size = font_size
    if bold is not None:
        builder.font.bold = bold
    if italic is not None:
        builder.font.italic = italic
    if border_bottom:
        builder.paragraph_format.borders.bottom.line_style = aw.LineStyle.SINGLE
    if custom_node_id is not None:
        paragraph_with_written_text.custom_node_id = int(custom_node_id)
    builder.writeln(text or '')

    sidecar_entry: Optional[dict[str, int | str]] = None
    if custom_node_id is not None:
        sidecar_entry = _build_custom_node_id_sidecar_entry(
            doc,
            paragraph_with_written_text,
            text or '',
            int(custom_node_id),
        )

    doc.save(str(path))
    if sidecar_entry is not None:
        _persist_custom_node_id_sidecar(path, doc_id, sidecar_entry)
    return True


def add_paragraph(
    doc_id: str,
    text: str = '',
    style: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    color_hex: Optional[str] = None,
    where: str = 'end',
    paragraph_index: Optional[int] = None,
    custom_node_id: Optional[int] = None,
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    if style:
        s = doc.styles.get_by_name(style)
        if s is not None:
            builder.paragraph_format.style = s
    if font_name:
        builder.font.name = font_name
    if font_size:
        builder.font.size = font_size
    if bold is not None:
        builder.font.bold = bold
    if italic is not None:
        builder.font.italic = italic
    col = hex_to_color(color_hex)
    if col is not None:
        builder.font.color = col
    paragraph_with_written_text = builder.current_paragraph
    builder.writeln(text or '')

    sidecar_entry: Optional[dict[str, int | str]] = None
    if custom_node_id is not None:
        paragraph_with_written_text.custom_node_id = int(custom_node_id)
        sidecar_entry = _build_custom_node_id_sidecar_entry(
            doc,
            paragraph_with_written_text,
            text or '',
            int(custom_node_id),
        )

    doc.save(str(path))
    if sidecar_entry is not None:
        _persist_custom_node_id_sidecar(path, doc_id, sidecar_entry)
    return True


def add_page_break(doc_id: str, where: str = 'end', paragraph_index: Optional[int] = None) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    builder.insert_break(aw.BreakType.PAGE_BREAK)
    doc.save(str(path))
    return True


def insert_list(
    doc_id: str,
    items: List[str] = None,
    kind: str = 'bullet',
    where: str = 'end',
    paragraph_index: Optional[int] = None,
) -> bool:
    if items is None:
        items = []
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    builder = aw.DocumentBuilder(doc)
    move_builder(doc, builder, where, paragraph_index)
    # Применяем список по умолчанию как в примерах: bullet или numbered
    if (kind or '').lower().startswith('number'):
        builder.list_format.apply_number_default()
    else:
        builder.list_format.apply_bullet_default()
    for it in items or []:
        builder.writeln(str(it))
    builder.list_format.remove_numbers()
    doc.save(str(path))
    return True


def insert_near_text(
    doc_id: str,
    target_text: Optional[str] = None,
    target_paragraph_index: Optional[int] = None,
    position: str = 'after',
    content_type: str = 'paragraph',
    text: Optional[str] = None,
    level: Optional[int] = None,
    items: Optional[List[str]] = None,
    kind: Optional[str] = None,
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    pidx = None
    if target_paragraph_index is not None and 0 <= target_paragraph_index < paras.count:
        pidx = int(target_paragraph_index)
    elif target_text:
        matches = find_paragraph_indices_by_anchor(doc, target_text)
        if matches:
            pidx = matches[0]
    if pidx is None:
        pidx = max(0, paras.count - 1)
    builder = aw.DocumentBuilder(doc)
    if (position or 'after') == 'before':
        builder.move_to_paragraph(pidx, 0)
    else:
        builder.move_to_paragraph(pidx, -1)
    ctype = (content_type or 'paragraph').lower()
    if ctype == 'heading':
        add_heading(
            doc_id=doc_id,
            text=text or '',
            level=level or 1,
            where='paragraph',
            paragraph_index=pidx,
        )
    elif ctype == 'list':
        insert_list(
            doc_id=doc_id,
            items=items or [],
            kind=kind or 'bullet',
            where='paragraph',
            paragraph_index=pidx,
        )
    else:
        builder.writeln(text or '')
        doc.save(str(path))
    return True


def format_range(
    doc_id: str,
    paragraph_index: int = 0,
    start: int = 0,
    end: int = 0,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color_hex: Optional[str] = None,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    para = paras[paragraph_index]
    acc = 0
    p_obj = para.as_paragraph()
    runs = p_obj.runs
    count = runs.count
    for i in range(count):
        run = runs[i]
        text = run.text or ''
        length = len(text)
        r_start = acc
        r_end = acc + length
        acc = r_end
        if length == 0:
            continue
        if r_end <= start or r_start >= end:
            continue
        font = run.font
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic
        if underline is not None:
            font.underline = aw.Underline.SINGLE if underline else aw.Underline.NONE
        col = hex_to_color(color_hex)
        if col is not None:
            font.color = col
        if font_name:
            font.name = font_name
        if font_size:
            font.size = font_size
    doc.save(str(path))
    return True


def delete_paragraph(doc_id: str, paragraph_index: int = 0) -> bool:
    path = ensure_path(doc_id)
    doc = aw.Document(str(path))
    paras = doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
    if paragraph_index < 0 or paragraph_index >= paras.count:
        raise IndexError('paragraph_index out of range')
    node = paras[paragraph_index]
    node.remove()
    doc.save(str(path))
    return True
