#!/bin/sh
set -e

exec python -m mcp_server \
    --transport "${MCP_TRANSPORT:-${TRANSPORT:-http}}" \
    --host "${MCP_HOST:-${HOST:-0.0.0.0}}" \
    --port "${MCP_PORT:-${PORT:-9110}}" \
    "$@"
