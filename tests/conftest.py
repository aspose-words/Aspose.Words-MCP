import os
import multiprocessing
import shutil
import tempfile
from pathlib import Path

import pytest
from fastmcp.utilities.tests import run_server_in_process

import mcp_server as srv
from core.utils.docs_util import init_data_dir


multiprocessing.set_start_method('spawn', force=True)

_TEMP_DIR: Path
_DATA_DIR: Path


def pytest_sessionstart(session):
    data_dir = Path(__file__).parent / 'data'
    if data_dir.exists():
        shutil.rmtree(data_dir, ignore_errors=True)
    init_data_dir(data_dir)
    global _DATA_DIR
    _DATA_DIR = data_dir
    global _TEMP_DIR
    _TEMP_DIR = Path(tempfile.mkdtemp(prefix='mcp_docs_'))
    os.environ['DOCS_DATA_DIR'] = _TEMP_DIR.as_posix()


def pytest_sessionfinish(session, exitstatus):
    shutil.rmtree(_TEMP_DIR, ignore_errors=True)


@pytest.fixture
def mcp_base_url():
    with run_server_in_process(
        srv.run_server,
        provide_host_and_port=True,
        transport='streamable-http',
        path='/mcp',
    ) as base_url:
        yield f'{base_url}/mcp'


@pytest.fixture
def mcp_client_config(mcp_base_url):
    return {'mcpServers': {'test': {'url': mcp_base_url}}}


@pytest.fixture
def result_file_path(request):
    return os.path.join(_DATA_DIR, f'{request.node.name}.docx')
