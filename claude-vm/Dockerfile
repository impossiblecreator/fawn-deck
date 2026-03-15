FROM node:22-slim

# Install system dependencies including LibreOffice + poppler for slide rendering
RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    python3 \
    python3-pip \
    python3-venv \
    curl \
    fonts-dejavu-core \
    libreoffice-impress \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Install Python packages
RUN python3 -m venv /opt/venv && /opt/venv/bin/pip install \
    python-pptx \
    lxml \
    Pillow
ENV PATH="/opt/venv/bin:$PATH"

# Install Claude Code globally
RUN npm install -g @anthropic-ai/claude-code

# Create non-root user
RUN useradd -m -s /bin/bash claude
RUN mkdir -p /workspace && chown claude:claude /workspace
RUN mkdir -p /home/claude/.config/claude && chown -R claude:claude /home/claude/.config
RUN mkdir -p /home/claude/.claude && chown -R claude:claude /home/claude/.claude

USER claude
WORKDIR /workspace

ENTRYPOINT ["claude", "--dangerously-skip-permissions"]
