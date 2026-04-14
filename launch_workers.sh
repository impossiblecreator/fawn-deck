#!/bin/bash
# Launch Claude Code worker instances in new terminal windows.
# Uses iTerm2 if installed, otherwise falls back to Terminal.
#
# Usage:
#   ./launch_workers.sh           # Launch all assigned workers (A, B, C)
#   ./launch_workers.sh A B       # Launch specific workers

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROMPT_FILE="$SCRIPT_DIR/worker_prompt.md"

if [ ! -f "$PROMPT_FILE" ]; then
    echo "Error: worker_prompt.md not found in $SCRIPT_DIR"
    exit 1
fi

# Default to all three workers if none specified
if [ $# -eq 0 ]; then
    WORKERS=(A B C)
else
    WORKERS=("$@")
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
    cat > "$TMPSCRIPT" <<INNEREOF
#!/bin/bash
cd '$SCRIPT_DIR'
WORKER_ID=$WORKER claude --dangerously-skip-permissions "\$(cat '$PROMPT_FILE')"
INNEREOF
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

    echo "Launching ${#WORKERS[@]} worker(s) in iTerm2: ${WORKERS[*]}"
    osascript -e "$ASCRIPT"
else
    echo "Launching ${#WORKERS[@]} worker(s) in Terminal: ${WORKERS[*]}"
    for i in "${!WORKERS[@]}"; do
        WORKER=$(echo "${WORKERS[$i]}" | tr '[:lower:]' '[:upper:]')
        SCRIPT="${TMPSCRIPTS[$i]}"
        osascript -e "tell application \"Terminal\"" \
                  -e "activate" \
                  -e "do script \"$SCRIPT\"" \
                  -e "end tell"
    done
fi

echo "Launched ${#WORKERS[@]} worker(s): ${WORKERS[*]}"
