import threading
import time

from core.store import DocumentStore


def test_store_basic_crud():
    store = DocumentStore()
    store.clear()
    assert store.size() == 0
    store.add_document('id1', 'file1.docx')
    assert store.size() == 1
    assert store.document_exists('id1') is True
    assert store.get_document_name('id1') == 'file1.docx'
    info = store.get_document_info('id1')
    assert isinstance(info, dict)
    assert info['name'] == 'file1.docx'
    all_docs = store.get_all_documents()
    assert 'id1' in all_docs
    assert store.remove_document('id1') is True
    assert store.remove_document('missing') is False
    assert store.size() == 0


def test_store_thread_safety_smoke():
    store = DocumentStore()
    store.clear()

    def worker(prefix: str):
        for i in range(50):
            store.add_document(f'{prefix}-{i}', f'{prefix}-{i}.docx')
            time.sleep(0.001)

    t1 = threading.Thread(target=worker, args=('A',))
    t2 = threading.Thread(target=worker, args=('B',))
    t1.start()
    t2.start()
    t1.join()
    t2.join()
    assert store.size() >= 100
    assert store.document_exists('A-0')
    assert store.document_exists('B-0')


def test_store_clear_overwrite_and_copy():
    store = DocumentStore()
    store.clear()
    store.add_document('id1', 'a.docx')
    assert store.get_document_name('id1') == 'a.docx'
    store.add_document('id1', 'b.docx')
    assert store.get_document_name('id1') == 'b.docx'
    snap = store.get_all_documents()
    snap.pop('id1')
    assert store.document_exists('id1')
    store.clear()
    assert store.size() == 0


def test_store_multithread_mixed_ops_smoke():
    store = DocumentStore()
    store.clear()

    def writer(prefix: str, count: int = 60):
        for i in range(count):
            store.add_document(f'{prefix}-{i}', f'{prefix}-{i}.docx')

    def remover(prefix: str, count: int = 30):
        for i in range(count):
            store.remove_document(f'{prefix}-{i}')

    t1 = threading.Thread(target=writer, args=('W', 80))
    t2 = threading.Thread(target=writer, args=('X', 80))
    t3 = threading.Thread(target=remover, args=('W', 40))
    t4 = threading.Thread(target=remover, args=('X', 20))
    t1.start()
    t2.start()
    t3.start()
    t4.start()
    t1.join()
    t2.join()
    t3.join()
    t4.join()
    assert store.size() >= 80
    assert store.document_exists('W-79')
    assert store.document_exists('X-79')
