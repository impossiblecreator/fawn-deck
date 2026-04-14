#!/bin/bash
# Launch Claude Code worker instances in new Terminal tabs.
# Usage:
#   ./launch_workers.sh           # Launch all assigned workers (A, B, C)
#   ./launch_workers.sh A B       # Launch specific workers

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROMPT_FILE="$SCRIPT_DIR/worker_prompt.md"

if [ ! -f "$PROMPT_FILE" ]; then
    echo "Error: worker_prompt.md not found in $SCRIPT_DIR"
    exit 1
fi

PROMPT="$(cat "$PROMPT_FILE")"

# Default to all three workers if none specified
if [ $# -eq 0 ]; then
    WORKERS=(A B C)
else
    WORKERS=("$@")
fi

for WORKER in "${WORKERS[@]}"; do
    WORKER=$(echo "$WORKER" | tr '[:lower:]' '[:upper:]')
    echo "Launching Worker $WORKER..."

    osascript <<EOF
tell application "Terminal"
    activate
    do script "cd '$SCRIPT_DIR' && WORKER_ID=$WORKER claude --dangerously-skip-permissions \"$(echo "$PROMPT" | sed 's/"/\\"/g')\""
end tell
EOF
done

echo "Launched ${#WORKERS[@]} worker(s): ${WORKERS[*]}"
