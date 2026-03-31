import logging
import threading
from typing import Dict, Optional

logger = logging.getLogger(__name__)


class DocumentStore:
    def __init__(self):
        self._store: Dict[str, Dict[str, str]] = {}
        self._lock = threading.RLock()
        logger.info('DocumentStore initialized')

    def add_document(self, doc_id: str, name: str) -> None:
        with self._lock:
            self._store[doc_id] = {'name': name}
            logger.info(f'Added document to store: {doc_id} ({name})')

    def get_document_name(self, doc_id: str) -> Optional[str]:
        name: Optional[str] = None
        with self._lock:
            doc_info = self._store.get(doc_id)
            if doc_info:
                name = doc_info['name']
        return name

    def get_document_info(self, doc_id: str) -> Optional[Dict[str, str]]:
        info: Optional[Dict[str, str]] = None
        with self._lock:
            info = self._store.get(doc_id, None)
        return info

    def document_exists(self, doc_id: str) -> bool:
        exists = False
        with self._lock:
            exists = doc_id in self._store
        return exists

    def remove_document(self, doc_id: str) -> bool:
        removed = False
        with self._lock:
            if doc_id in self._store:
                del self._store[doc_id]
                logger.info(f'Removed document from store: {doc_id}')
                removed = True
        return removed

    def get_all_documents(self) -> Dict[str, Dict[str, str]]:
        documents: Dict[str, Dict[str, str]] = {}
        with self._lock:
            documents = self._store.copy()
        return documents

    def clear(self) -> None:
        with self._lock:
            count = len(self._store)
            self._store.clear()
            logger.info(f'Cleared {count} documents from store')

    def size(self) -> int:
        size = 0
        with self._lock:
            size = len(self._store)
        return size


document_store = DocumentStore()
