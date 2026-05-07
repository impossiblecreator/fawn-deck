#!/bin/bash
# Launch AI worker instances in new terminal windows.
# Uses iTerm2 if installed, otherwise falls back to Terminal.
#
# Usage:
#   ./launch_workers.sh --agent codex      # Launch all assigned workers with Codex
#   ./launch_workers.sh --agent claude     # Launch all assigned workers with Claude Code
#   ./launch_workers.sh --agent codex A B  # Launch specific workers

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROMPT_FILE="$SCRIPT_DIR/worker_prompt.md"
VENV_DIR="$SCRIPT_DIR/.venv"
AGENT="${FAWN_AGENT:-claude}"

if [ ! -f "$PROMPT_FILE" ]; then
    echo "Error: worker_prompt.md not found in $SCRIPT_DIR"
    exit 1
fi

usage() {
    sed -n '2,9p' "$0" | sed 's/^# \{0,1\}//'
}

WORKERS=()
while [ $# -gt 0 ]; do
    case "$1" in
        --agent)
            if [ -z "${2:-}" ]; then
                echo "Error: --agent requires codex or claude"
                exit 1
            fi
            AGENT="$2"
            shift 2
            ;;
        --agent=*)
            AGENT="${1#--agent=}"
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            WORKERS+=("$1")
            shift
            ;;
    esac
done

AGENT=$(echo "$AGENT" | tr '[:upper:]' '[:lower:]')
case "$AGENT" in
    codex|claude)
        ;;
    *)
        echo "Error: unsupported agent '$AGENT'. Use codex or claude."
        exit 1
        ;;
esac

if ! command -v "$AGENT" >/dev/null 2>&1; then
    echo "Error: '$AGENT' CLI not found on PATH."
    exit 1
fi

ensure_python_env() {
    local VENV_PY="$VENV_DIR/bin/python"

    if [ ! -x "$VENV_PY" ]; then
        echo "Creating Python virtual environment in .venv..."
        python3 -m venv "$VENV_DIR" || exit 1
    fi

    if ! "$VENV_PY" -c "import pptx" >/dev/null 2>&1; then
        echo "Installing Python dependencies into .venv..."
        "$VENV_PY" -m pip install -r "$SCRIPT_DIR/requirements.txt" || exit 1
    fi
}

ensure_python_env

# Default to all three workers if none specified
if [ ${#WORKERS[@]} -eq 0 ]; then
    WORKERS=(A B C)
else
    WORKERS=("${WORKERS[@]}")
fi

# Detect terminal app once
USE_ITERM=false
if osascript -e 'id of application "iTerm2"' &>/dev/null; then
    USE_ITERM=true
fi

# Create temp scripts for each worker first
TMPSCRIPTS=()
for WORKER in "${WORKERS[@]}"; do
    WORKER=$(echo "$WORKER" | tr '[:lower:]' '[:upper:]')
    TMPSCRIPT=$(mktemp /tmp/fawn_worker_XXXXXXXX)
    if [ "$AGENT" = "codex" ]; then
        cat > "$TMPSCRIPT" <<INNEREOF
#!/bin/bash
cd '$SCRIPT_DIR'
source '$VENV_DIR/bin/activate'
WORKER_ID=$WORKER codex --cd '$SCRIPT_DIR' --sandbox danger-full-access --ask-for-approval never \${CODEX_EXTRA_ARGS:-} "\$(cat '$PROMPT_FILE')"
STATUS=\$?
echo
echo "Worker $WORKER exited with status \$STATUS. Shell is staying in: \$(pwd)"
exec "\${SHELL:-/bin/bash}"
INNEREOF
    else
        cat > "$TMPSCRIPT" <<INNEREOF
#!/bin/bash
cd '$SCRIPT_DIR'
source '$VENV_DIR/bin/activate'
WORKER_ID=$WORKER claude --dangerously-skip-permissions \${CLAUDE_EXTRA_ARGS:-} "\$(cat '$PROMPT_FILE')"
STATUS=\$?
echo
echo "Worker $WORKER exited with status \$STATUS. Shell is staying in: \$(pwd)"
exec "\${SHELL:-/bin/bash}"
INNEREOF
    fi
    chmod +x "$TMPSCRIPT"
    TMPSCRIPTS+=("$TMPSCRIPT")
done

if $USE_ITERM; then
    # Build a single AppleScript that creates all windows at once
    ASCRIPT='tell application "iTerm2"
    activate
'
    for i in "${!WORKERS[@]}"; do
        WORKER=$(echo "${WORKERS[$i]}" | tr '[:lower:]' '[:upper:]')
        SCRIPT="${TMPSCRIPTS[$i]}"
        ASCRIPT+="
    create window with default profile
    delay 0.5
    tell current window
        tell current session
            write text \"$SCRIPT\"
        end tell
    end tell
"
    done
    ASCRIPT+='end tell'

    echo "Launching ${#WORKERS[@]} $AGENT worker(s) in iTerm2: ${WORKERS[*]}"
    osascript -e "$ASCRIPT"
else
    echo "Launching ${#WORKERS[@]} $AGENT worker(s) in Terminal: ${WORKERS[*]}"
    for i in "${!WORKERS[@]}"; do
        WORKER=$(echo "${WORKERS[$i]}" | tr '[:lower:]' '[:upper:]')
        SCRIPT="${TMPSCRIPTS[$i]}"
        osascript -e "tell application \"Terminal\"" \
                  -e "activate" \
                  -e "do script \"$SCRIPT\"" \
                  -e "end tell"
    done
fi

echo "Launched ${#WORKERS[@]} $AGENT worker(s): ${WORKERS[*]}"
