#!/bin/bash
# Launch Claude Code workers in Docker containers.
# Usage:
#   ./run-claude.sh              # Launch all workers (A, B, C) in new Terminal tabs
#   ./run-claude.sh A B          # Launch specific workers
#   ./run-claude.sh -w A         # Launch a single worker in the current terminal

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
PROMPT_FILE="$PROJECT_DIR/worker_prompt.md"

DOCKER_COMMON=(
    --rm
    -v "$PROJECT_DIR:/workspace"
    -v /Users/peterfitzpatrick/Desktop:/Users/peterfitzpatrick/Desktop:ro
    -v /Users/peterfitzpatrick/Downloads:/Users/peterfitzpatrick/Downloads:ro
    -v /Users/peterfitzpatrick/.config/claude/credentials.json:/home/claude/.config/claude/credentials.json:ro
    -e CLAUDE_CODE_USE_VERTEX=1
    -e ANTHROPIC_VERTEX_PROJECT_ID=claude-472223
    -e CLOUD_ML_REGION=us-east5
    -e GOOGLE_APPLICATION_CREDENTIALS=/home/claude/.config/claude/credentials.json
    -v /Users/peterfitzpatrick/Library/Fonts:/home/claude/.local/share/fonts:ro
    -v /tmp:/tmp
    -v /var/folders:/var/folders:ro
    claude-sandbox
)

# Build the prompt argument
PROMPT_ARG=""
if [ -f "$PROMPT_FILE" ]; then
    PROMPT_ARG="$(cat "$PROMPT_FILE")"
fi

# Launch a single worker in the current terminal (legacy -w flag)
launch_here() {
    local WID="$1"
    echo -ne "\033]0;Worker $WID\007"
    if [ -n "$PROMPT_ARG" ]; then
        docker run -it --hostname "worker-$WID" -e WORKER_ID="$WID" "${DOCKER_COMMON[@]}" "$PROMPT_ARG"
    else
        docker run -it --hostname "worker-$WID" -e WORKER_ID="$WID" "${DOCKER_COMMON[@]}"
    fi
}

# Launch a worker in a new Terminal tab
launch_tab() {
    local WID="$1"
    osascript <<EOF
tell application "Terminal"
    activate
    do script "cd '$PROJECT_DIR' && '$SCRIPT_DIR/run-claude.sh' -w $WID"
end tell
EOF
}

# Parse arguments
if [ "$1" = "-w" ]; then
    # Single-worker mode: run in current terminal
    launch_here "${2:-A}"
elif [ $# -eq 0 ]; then
    # No args: launch all three in new tabs
    for W in A B C; do
        echo "Launching Worker $W..."
        launch_tab "$W"
    done
    echo "Launched 3 workers: A B C"
else
    # Specific workers in new tabs
    for W in "$@"; do
        W=$(echo "$W" | tr '[:lower:]' '[:upper:]')
        echo "Launching Worker $W..."
        launch_tab "$W"
    done
    echo "Launched $# worker(s): $*"
fi
