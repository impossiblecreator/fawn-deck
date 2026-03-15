#!/bin/bash
# Run Claude Code in a Docker container with slide design tools
# Usage: 
#   ./run-claude.sh                     # Interactive mode (Worker A by default)
#   ./run-claude.sh -w B                # Interactive mode as Worker B  
#   ./run-claude.sh -w C                # Interactive mode as Worker C

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

# Parse worker flag
WORKER_ID="A"
while getopts "w:" opt; do
    case $opt in
        w) WORKER_ID="$OPTARG" ;;
    esac
done
shift $((OPTIND-1))

# Set terminal tab title
echo -ne "\033]0;Worker $WORKER_ID\007"

DOCKER_COMMON=(
    --rm
    --hostname "worker-$WORKER_ID"
    -v "$PROJECT_DIR:/workspace"
    -v /Users/peterfitzpatrick/Desktop:/Users/peterfitzpatrick/Desktop:ro
    -v /Users/peterfitzpatrick/Downloads:/Users/peterfitzpatrick/Downloads:ro
    -v /Users/peterfitzpatrick/.config/claude/credentials.json:/home/claude/.config/claude/credentials.json:ro
    -e CLAUDE_CODE_USE_VERTEX=1
    -e ANTHROPIC_VERTEX_PROJECT_ID=claude-472223
    -e CLOUD_ML_REGION=us-east5
    -e GOOGLE_APPLICATION_CREDENTIALS=/home/claude/.config/claude/credentials.json
    -e WORKER_ID="$WORKER_ID"
    -v /tmp:/tmp
    -v /var/folders:/var/folders:ro
    claude-sandbox
)

# Use Claude Code's --name flag to label the instance
docker run -it "${DOCKER_COMMON[@]}" --name "Worker $WORKER_ID"
