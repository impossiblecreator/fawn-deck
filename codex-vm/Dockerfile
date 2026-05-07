FROM node:22-slim

# Install system dependencies including LibreOffice + poppler for slide rendering.
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

# Install Python packages.
RUN python3 -m venv /opt/venv && /opt/venv/bin/pip install \
    python-pptx \
    lxml \
    Pillow
ENV PATH="/opt/venv/bin:$PATH"

# Install Codex CLI globally.
RUN npm install -g @openai/codex

# Create non-root user.
RUN useradd -m -s /bin/bash codex
RUN mkdir -p /workspace && chown codex:codex /workspace
RUN mkdir -p /home/codex/.codex && chown -R codex:codex /home/codex/.codex

USER codex
WORKDIR /workspace

ENTRYPOINT ["codex", "--cd", "/workspace", "--sandbox", "danger-full-access", "--ask-for-approval", "never"]
