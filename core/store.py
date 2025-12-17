import logging
from typing import Dict, Optional
import threading
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
        with self._lock:
            doc_info = self._store.get(doc_id)
            if doc_info:
                return doc_info['name']
            return None

    def get_document_info(self, doc_id: str) -> Optional[Dict[str, str]]:
        with self._lock:
            return self._store.get(doc_id, None)

    def document_exists(self, doc_id: str) -> bool:
        with self._lock:
            return doc_id in self._store

    def remove_document(self, doc_id: str) -> bool:
        with self._lock:
            if doc_id in self._store:
                del self._store[doc_id]
                logger.info(f'Removed document from store: {doc_id}')
                return True
            return False

    def get_all_documents(self) -> Dict[str, Dict[str, str]]:
        with self._lock:
            return self._store.copy()

    def clear(self) -> None:
        with self._lock:
            count = len(self._store)
            self._store.clear()
            logger.info(f'Cleared {count} documents from store')

    def size(self) -> int:
        with self._lock:
            return len(self._store)
document_store = DocumentStore()
