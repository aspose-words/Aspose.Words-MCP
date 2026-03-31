import threading

from core.store import DocumentStore


def test_boundary_missing_ids_and_empty_store_reads_are_stable():
    store = DocumentStore()

    assert store.size() == 0
    assert store.get_document_name('missing-id') is None
    assert store.get_document_info('missing-id') is None
    assert store.document_exists('missing-id') is False
    assert store.remove_document('missing-id') is False
    assert store.get_all_documents() == {}


def test_delete_hit_then_delete_miss_preserves_boolean_contract_and_size():
    store = DocumentStore()
    store.add_document('doc-1', 'alpha.docx')

    assert store.size() == 1
    assert store.remove_document('doc-1') is True
    assert store.size() == 0
    assert store.remove_document('doc-1') is False


def test_snapshot_top_level_copy_does_not_mutate_store_keys():
    store = DocumentStore()
    store.add_document('doc-1', 'alpha.docx')

    snapshot = store.get_all_documents()
    del snapshot['doc-1']

    assert snapshot == {}
    assert store.document_exists('doc-1') is True
    assert store.get_document_name('doc-1') == 'alpha.docx'


def test_snapshot_nested_value_is_live_reference_boundary_behavior():
    store = DocumentStore()
    store.add_document('doc-1', 'alpha.docx')

    snapshot = store.get_all_documents()
    snapshot['doc-1']['name'] = 'changed-via-snapshot.docx'

    assert store.get_document_name('doc-1') == 'changed-via-snapshot.docx'


def test_property_clear_is_idempotent():
    store = DocumentStore()
    store.add_document('doc-1', 'alpha.docx')
    store.add_document('doc-2', 'beta.docx')

    store.clear()
    first_size = store.size()
    store.clear()
    second_size = store.size()

    assert first_size == 0
    assert second_size == 0
    assert store.get_all_documents() == {}


def test_property_size_tracks_unique_key_count_under_overwrite():
    store = DocumentStore()

    expected_sizes = []
    for i in range(10):
        store.add_document(f'doc-{i}', f'name-{i}.docx')
        expected_sizes.append(i + 1)

    observed_sizes = []
    for i in range(10):
        observed_sizes.append(store.size())
        # overwrite should not increase unique-key count
        store.add_document(f'doc-{i}', f'name-{i}-updated.docx')

    assert expected_sizes[-1] == 10
    assert observed_sizes == [10] * 10
    assert store.size() == 10


def test_concurrent_delete_same_id_exactly_one_hit_and_rest_miss():
    store = DocumentStore()
    store.add_document('shared-doc', 'shared.docx')

    start = threading.Barrier(8)
    results: list[bool] = []
    results_lock = threading.Lock()

    def worker() -> None:
        start.wait()
        removed = store.remove_document('shared-doc')
        with results_lock:
            results.append(removed)

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()

    assert len(results) == 8
    assert sum(1 for value in results if value is True) == 1
    assert sum(1 for value in results if value is False) == 7
    assert store.document_exists('shared-doc') is False
    assert store.size() == 0


def test_concurrent_snapshot_reads_with_writes_return_consistent_shapes():
    store = DocumentStore()
    total_docs = 100
    snapshots: list[dict[str, dict[str, str]]] = []

    def writer() -> None:
        for i in range(total_docs):
            store.add_document(f'doc-{i}', f'name-{i}.docx')

    def reader() -> None:
        for _ in range(total_docs):
            snap = store.get_all_documents()
            snapshots.append(snap)

    t_writer = threading.Thread(target=writer)
    t_reader = threading.Thread(target=reader)

    t_writer.start()
    t_reader.start()
    t_writer.join()
    t_reader.join()

    assert len(snapshots) == total_docs
    assert all(isinstance(s, dict) for s in snapshots)
    assert all(0 <= len(s) <= total_docs for s in snapshots)
    final_snapshot = store.get_all_documents()
    assert len(final_snapshot) == total_docs
    assert store.size() == total_docs
