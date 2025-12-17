import base64
from pathlib import Path
from typing import Optional, List, Dict, Any, Callable
from core import layout as _layout
from core import content as _content
from core import export as _export
from core import properties as _properties
from core import io as _io
from core import reading as _reading

def _store_add_mapping(document_store, doc_id: str, name: Optional[str]) -> None:
    if name:
        document_store.add_document(doc_id, name)

def _find_doc_id_by_name(document_store, name: str) -> Optional[str]:
    items = document_store.get_all_documents() or {}
    for did, info in items.items():
        if info.get('name') == name:
            return did
    return None

def _resolve_doc_id_for_filename(get_docs: Callable[[], Any], document_store, filename: str) -> str:
    did = _find_doc_id_by_name(document_store, filename)
    if did:
        return did
    p = Path(str(filename))
    if p is not None and p.exists() and p.is_file():
        doc_id, name = _io.import_from_file(str(p))
        _store_add_mapping(document_store, doc_id, name)
        return doc_id
    try_path = get_docs().get_data_dir() / str(filename)
    if try_path.exists() and try_path.is_file():
        doc_id, name = _io.import_from_file(str(try_path))
        _store_add_mapping(document_store, doc_id, name)
        return doc_id
    doc_id, name = _io.create_document(name=filename)
    _store_add_mapping(document_store, doc_id, name)
    return doc_id

def build_alias_functions(get_docs: Callable[[], Any], document_store) -> Dict[str, Callable[..., Dict[str, Any]]]:

    def file_create_document(filename: str, title: Optional[str]=None, author: Optional[str]=None):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        if title is not None or author is not None:
            _properties.properties_set(did, title=title, author=author)
        _store_add_mapping(document_store, did, filename)
        return {'docId': did, 'name': filename}

    def file_get_document_info(filename: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        return _reading.get_info(did)

    def file_get_document_text(filename: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        return {'text': get_docs().get_text(did)}

    def file_get_document_outline(filename: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        return {'outline': _reading.get_outline(did)}

    def file_get_document_xml(filename: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        return {'xml': get_docs().get_xml(did)}

    def file_convert_to_pdf(filename: str, output_filename: Optional[str]=None):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        data, mime, ext = get_docs().export(did, fmt='pdf')
        b64 = base64.b64encode(data).decode('utf-8')
        out = {'base64': b64, 'mime': mime, 'ext': ext}
        if output_filename:
            Path(output_filename).write_bytes(data)
            out['output'] = str(output_filename)
        return out

    def file_add_picture(filename: str, image_path: str, width_points: Optional[float]=None, height_points: Optional[float]=None, keep_aspect: bool=False):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        try:
            data = Path(image_path).read_bytes()
        except FileNotFoundError:
            raise FileNotFoundError(f'Image not found: {image_path}')
        _content.insert_image(did, data, where='end', width_points=width_points, height_points=height_points, keep_aspect=keep_aspect)
        return {}

    def file_add_header_text(filename: str, text: str, primary: bool=True):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.header_add_text(did, text, primary=primary)
        return {}

    def file_add_footer_text(filename: str, text: str, primary: bool=True):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.footer_add_text(did, text, primary=primary)
        return {}

    def file_add_page_numbering(filename: str, format_string: str='Page {PAGE} of {NUMPAGES}'):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.add_page_numbering(did, format_string=format_string)
        return {}

    def file_set_different_first_page_header_footer(filename: str, enabled: bool=True):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.set_different_first_page_header_footer(did, enabled=enabled)
        return {}

    def file_set_page_setup(filename: str, margins: Optional[dict]=None, orientation: Optional[str]=None, paper: Optional[str]=None, columns: Optional[int]=None):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.set_page_setup(did, margins=margins, orientation=orientation, paper=paper, columns=columns)
        return {}

    def file_insert_section_break(filename: str, kind: str='nextPage'):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _layout.insert_section_break(did, kind=kind)
        return {}

    def file_insert_html_end(filename: str, html: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        _content.insert_html(did, html, where='end', paragraph_index=None)
        return {}

    def file_render_page_base64(filename: str, page_index: int=0, fmt: str='png', dpi: int=150):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        data, mime, ext = _export.render_page(did, page_index=page_index, fmt=fmt, dpi=dpi)
        return {'base64': base64.b64encode(data).decode('utf-8'), 'mime': mime, 'ext': ext}

    def file_export_base64_advanced(filename: str, fmt: str, options: Optional[dict]=None):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        data, mime, ext = _export.export_advanced(did, fmt=fmt, options=options)
        return {'base64': base64.b64encode(data).decode('utf-8'), 'mime': mime, 'ext': ext}

    def file_search_and_replace(filename: str, find_text: str, replace_text: str, replace_all: bool=True, case_sensitive: bool=False):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        count = _content.replace_text(did, search=find_text, replace=replace_text, replace_all=replace_all, case_sensitive=case_sensitive)
        return {'count': int(count)}

    def file_delete_document(filename: str):
        did = _resolve_doc_id_for_filename(get_docs, document_store, filename)
        document_store.remove_document(did)
        _io.delete(did)
        return {}

    def file_copy_document(source_filename: str, destination_filename: Optional[str]=None):
        src_id = _resolve_doc_id_for_filename(get_docs, document_store, source_filename)
        new_id = _io.copy(src_id)
        dest_name = destination_filename or f'copy_of_{source_filename}'
        _store_add_mapping(document_store, new_id, dest_name)
        return {'docId': new_id}

    def file_list_available_documents(directory: str='.'):
        p = Path(directory)
        if not p.exists() or not p.is_dir():
            return {'files': []}
        files = [str(x.name) for x in p.glob('*.docx')]
        return {'files': files}
    return {'file_create_document': file_create_document, 'file_get_document_info': file_get_document_info, 'file_get_document_text': file_get_document_text, 'file_get_document_outline': file_get_document_outline, 'file_get_document_xml': file_get_document_xml, 'file_convert_to_pdf': file_convert_to_pdf, 'file_add_picture': file_add_picture, 'file_add_header_text': file_add_header_text, 'file_add_footer_text': file_add_footer_text, 'file_add_page_numbering': file_add_page_numbering, 'file_set_different_first_page_header_footer': file_set_different_first_page_header_footer, 'file_set_page_setup': file_set_page_setup, 'file_insert_section_break': file_insert_section_break, 'file_insert_html_end': file_insert_html_end, 'file_render_page_base64': file_render_page_base64, 'file_export_base64_advanced': file_export_base64_advanced, 'file_search_and_replace': file_search_and_replace, 'file_delete_document': file_delete_document, 'file_copy_document': file_copy_document, 'file_list_available_documents': file_list_available_documents}

def register_alias_tools(mcp, functions: Dict[str, Callable[..., Dict[str, Any]]]) -> List[str]:
    registered: List[str] = []

    @mcp.tool()
    def file_create_document(filename: str, title: Optional[str]=None, author: Optional[str]=None):
        return functions['file_create_document'](filename=filename, title=title, author=author)
    registered.append('file_create_document')

    @mcp.tool()
    def file_get_document_info(filename: str):
        return functions['file_get_document_info'](filename=filename)
    registered.append('file_get_document_info')

    @mcp.tool()
    def file_get_document_text(filename: str):
        return functions['file_get_document_text'](filename=filename)
    registered.append('file_get_document_text')

    @mcp.tool()
    def file_get_document_outline(filename: str):
        return functions['file_get_document_outline'](filename=filename)
    registered.append('file_get_document_outline')

    @mcp.tool()
    def file_get_document_xml(filename: str):
        return functions['file_get_document_xml'](filename=filename)
    registered.append('file_get_document_xml')

    @mcp.tool()
    def file_convert_to_pdf(filename: str, output_filename: Optional[str]=None):
        return functions['file_convert_to_pdf'](filename=filename, output_filename=output_filename)
    registered.append('file_convert_to_pdf')

    @mcp.tool()
    def file_add_picture(filename: str, image_path: str, width_points: Optional[float]=None, height_points: Optional[float]=None, keep_aspect: bool=False):
        return functions['file_add_picture'](filename=filename, image_path=image_path, width_points=width_points, height_points=height_points, keep_aspect=keep_aspect)
    registered.append('file_add_picture')

    @mcp.tool()
    def file_add_header_text(filename: str, text: str, primary: bool=True):
        return functions['file_add_header_text'](filename=filename, text=text, primary=primary)
    registered.append('file_add_header_text')

    @mcp.tool()
    def file_add_footer_text(filename: str, text: str, primary: bool=True):
        return functions['file_add_footer_text'](filename=filename, text=text, primary=primary)
    registered.append('file_add_footer_text')

    @mcp.tool()
    def file_add_page_numbering(filename: str, format_string: str='Page {PAGE} of {NUMPAGES}'):
        return functions['file_add_page_numbering'](filename=filename, format_string=format_string)
    registered.append('file_add_page_numbering')

    @mcp.tool()
    def file_set_different_first_page_header_footer(filename: str, enabled: bool=True):
        return functions['file_set_different_first_page_header_footer'](filename=filename, enabled=enabled)
    registered.append('file_set_different_first_page_header_footer')

    @mcp.tool()
    def file_set_page_setup(filename: str, margins: Optional[dict]=None, orientation: Optional[str]=None, paper: Optional[str]=None, columns: Optional[int]=None):
        return functions['file_set_page_setup'](filename=filename, margins=margins, orientation=orientation, paper=paper, columns=columns)
    registered.append('file_set_page_setup')

    @mcp.tool()
    def file_insert_section_break(filename: str, kind: str='nextPage'):
        return functions['file_insert_section_break'](filename=filename, kind=kind)
    registered.append('file_insert_section_break')

    @mcp.tool()
    def file_insert_html_end(filename: str, html: str):
        return functions['file_insert_html_end'](filename=filename, html=html)
    registered.append('file_insert_html_end')

    @mcp.tool()
    def file_render_page_base64(filename: str, page_index: int=0, fmt: str='png', dpi: int=150):
        return functions['file_render_page_base64'](filename=filename, page_index=page_index, fmt=fmt, dpi=dpi)
    registered.append('file_render_page_base64')

    @mcp.tool()
    def file_export_base64_advanced(filename: str, fmt: str, options: Optional[dict]=None):
        return functions['file_export_base64_advanced'](filename=filename, fmt=fmt, options=options)
    registered.append('file_export_base64_advanced')

    @mcp.tool()
    def file_search_and_replace(filename: str, find_text: str, replace_text: str, replace_all: bool=True, case_sensitive: bool=False):
        return functions['file_search_and_replace'](filename=filename, find_text=find_text, replace_text=replace_text, replace_all=replace_all, case_sensitive=case_sensitive)
    registered.append('file_search_and_replace')

    @mcp.tool()
    def file_delete_document(filename: str):
        return functions['file_delete_document'](filename=filename)
    registered.append('file_delete_document')

    @mcp.tool()
    def file_copy_document(source_filename: str, destination_filename: Optional[str]=None):
        return functions['file_copy_document'](source_filename=source_filename, destination_filename=destination_filename)
    registered.append('file_copy_document')

    @mcp.tool()
    def file_list_available_documents(directory: str='.'):
        return functions['file_list_available_documents'](directory=directory)
    registered.append('file_list_available_documents')
    return registered
