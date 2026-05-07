#!/bin/bash
# Convenience wrapper for launching local Codex workers.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/launch_workers.sh" --agent codex "$@"
