#!/bin/bash
# Launch Codex workers in Docker containers.
# Usage:
#   ./run-codex.sh              # Launch all workers (A, B, C) in new Terminal tabs
#   ./run-codex.sh A B          # Launch specific workers
#   ./run-codex.sh -w A         # Launch a single worker in the current terminal

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
PROMPT_FILE="$PROJECT_DIR/worker_prompt.md"
DOCKER_IMAGE="${CODEX_DOCKER_IMAGE:-codex-sandbox}"

if ! docker image inspect "$DOCKER_IMAGE" >/dev/null 2>&1; then
    echo "Docker image '$DOCKER_IMAGE' not found. Building it now..."
    docker build -t "$DOCKER_IMAGE" "$SCRIPT_DIR" || exit 1
fi

DOCKER_COMMON=(
    --rm
    -v "$PROJECT_DIR:/workspace"
    -v /tmp:/tmp
    "$DOCKER_IMAGE"
)

if [ -d "$HOME/.codex" ]; then
    DOCKER_COMMON=(-v "$HOME/.codex:/home/codex/.codex" "${DOCKER_COMMON[@]}")
fi

if [ -d "$HOME/Library/Fonts" ]; then
    DOCKER_COMMON=(-v "$HOME/Library/Fonts:/home/codex/.local/share/fonts:ro" "${DOCKER_COMMON[@]}")
fi

if [ -d "$HOME/Desktop" ]; then
    DOCKER_COMMON=(-v "$HOME/Desktop:$HOME/Desktop:ro" "${DOCKER_COMMON[@]}")
fi

if [ -d "$HOME/Downloads" ]; then
    DOCKER_COMMON=(-v "$HOME/Downloads:$HOME/Downloads:ro" "${DOCKER_COMMON[@]}")
fi

if [ -n "${OPENAI_API_KEY:-}" ]; then
    DOCKER_COMMON=(-e OPENAI_API_KEY "${DOCKER_COMMON[@]}")
fi

# Build the prompt argument.
PROMPT_ARG=""
if [ -f "$PROMPT_FILE" ]; then
    PROMPT_ARG="$(cat "$PROMPT_FILE")"
fi

# Launch a single worker in the current terminal.
launch_here() {
    local WID="$1"
    echo -ne "\033]0;Worker $WID\007"
    if [ -n "$PROMPT_ARG" ]; then
        docker run -it --hostname "worker-$WID" -e WORKER_ID="$WID" "${DOCKER_COMMON[@]}" "$PROMPT_ARG"
    else
        docker run -it --hostname "worker-$WID" -e WORKER_ID="$WID" "${DOCKER_COMMON[@]}"
    fi
}

# Launch a worker in a new Terminal tab.
launch_tab() {
    local WID="$1"
    osascript <<EOF
tell application "Terminal"
    activate
    do script "cd '$PROJECT_DIR' && '$SCRIPT_DIR/run-codex.sh' -w $WID"
end tell
EOF
}

# Parse arguments.
if [ "$1" = "-w" ]; then
    # Single-worker mode: run in current terminal.
    launch_here "${2:-A}"
elif [ $# -eq 0 ]; then
    # No args: launch all three in new tabs.
    for W in A B C; do
        echo "Launching Worker $W..."
        launch_tab "$W"
    done
    echo "Launched 3 workers: A B C"
else
    # Specific workers in new tabs.
    for W in "$@"; do
        W=$(echo "$W" | tr '[:lower:]' '[:upper:]')
        echo "Launching Worker $W..."
        launch_tab "$W"
    done
    echo "Launched $# worker(s): $*"
fi
