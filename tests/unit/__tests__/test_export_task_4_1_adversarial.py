from pathlib import Path
from types import SimpleNamespace

import pytest

pytest.importorskip('aspose.words')

from core.export import export_markdown


class FakeDoc:
    def __init__(self, payloads: list[bytes], fail: bool = False):
        self._payloads = payloads
        self._index = 0
        self._fail = fail
        self.calls: list[tuple[str, object]] = []

    def save(self, path: str, save_format: object):
        self.calls.append((path, save_format))
        if self._fail:
            raise OSError('simulated save failure')
        payload = self._payloads[self._index]
        self._index += 1
        Path(path).write_bytes(payload)


@pytest.mark.parametrize(
    'payload',
    [
        b'# Header\n\nBody',
        b'',
        'emoji-🧪-مرحبا'.encode('utf-8'),
        b'zero\x00byte',
        ('A' * 12_000).encode('utf-8'),
    ],
)
def test_export_markdown_round_trip_and_temp_lifetime(monkeypatch, payload):
    monkeypatch.setattr('core.export.aw.SaveFormat', SimpleNamespace(MARKDOWN='MARKDOWN_FMT'))
    doc = FakeDoc([payload])

    result = export_markdown(doc)

    assert result == payload
    assert len(doc.calls) == 1
    saved_path, saved_format = doc.calls[0]
    assert saved_format == 'MARKDOWN_FMT'
    assert saved_path.endswith('.md')
    assert '..' not in Path(saved_path).parts
    assert Path(saved_path).exists() is False


def test_export_markdown_repeated_calls_use_distinct_temp_paths(monkeypatch):
    monkeypatch.setattr('core.export.aw.SaveFormat', SimpleNamespace(MARKDOWN='MARKDOWN_FMT'))
    doc = FakeDoc([b'first-export', b'second-export'])

    first = export_markdown(doc)
    second = export_markdown(doc)

    assert first == b'first-export'
    assert second == b'second-export'
    assert len(doc.calls) == 2
    first_path = doc.calls[0][0]
    second_path = doc.calls[1][0]
    assert first_path != second_path
    assert Path(first_path).exists() is False
    assert Path(second_path).exists() is False


def test_export_markdown_propagates_save_error(monkeypatch):
    monkeypatch.setattr('core.export.aw.SaveFormat', SimpleNamespace(MARKDOWN='MARKDOWN_FMT'))
    doc = FakeDoc([b'unused'], fail=True)

    with pytest.raises(OSError, match='simulated save failure'):
        export_markdown(doc)
