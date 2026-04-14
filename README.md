# Fawn Friends Deck — Multi-Worker Slide Design

I was using an AI deck building product and was frustrated with how slow the process was. So, I created an orchestrator that allows me to build slides with three agents at once. Much faster. 

This repo manages conflicts by creating three worker powerpoint decks from a source deck. You assign slides to each worker and work on the slides. When you're done for the time being, you merge the slides into the final deck and then promote the final deck. When you want to start building again, run `setup` to reset the worker decks with the current final deck.

If you want to make manual edits, open the worker deck and make the edits yourself. 

## Dependencies

**Python 3.7+** and these packages:
```bash
pip install -r requirements.txt
```

**System dependencies:**
- [LibreOffice](https://www.libreoffice.org/) — renders slides to PNG
- `pdftoppm` (from [Poppler](https://poppler.freedesktop.org/)) — part of the render pipeline

On macOS:
```bash
brew install --cask libreoffice
brew install poppler
```

## Setup

Copy your source deck into place, then initialize worker files:

```bash
cp your_deck.pptx source/deck_original.pptx
python3 slide_manager.py setup
```

This creates `workers/worker_A.pptx`, `worker_B.pptx`, and `worker_C.pptx` — each a full copy of the source deck.

**Create a brand guide** — add a `BRAND_GUIDE.md` to the project root with your design system: fonts, colors, spacing rules, and a design loop for workers to follow. Workers will reference this before editing any slide.

## Launching Workers

Workers are Claude Code instances — one per terminal, each with a different `WORKER_ID`.

**Step 1: Assign slides (coordinator)**
```bash
python3 slide_manager.py assign A 1 2 3 4
python3 slide_manager.py assign B 5 6 7 8
python3 slide_manager.py assign C 9 10 11 12
```
Each slide can only be assigned to one worker. Conflicts are rejected.

**Step 2: Launch workers**

There are two ways to launch workers. Both do the same thing — open Terminal tabs and start the AI workers automatically. The difference is safety.

**Option A: Simple launch** — faster, no extra setup required.
```bash
./launch_workers.sh              # Launch all workers (A, B, C)
./launch_workers.sh A B          # Launch specific workers
```
The AI workers run directly on your computer. They're instructed to only edit their own slide files, and in practice they do. But there's nothing physically stopping them from touching other files on your machine if something goes wrong.

**Option B: Sandboxed launch** — requires [Docker Desktop](https://www.docker.com/products/docker-desktop/) to be installed and running.
```bash
./claude-vm/run-claude.sh        # Launch all workers (A, B, C)
./claude-vm/run-claude.sh A B    # Launch specific workers
```
Each AI worker runs in an isolated container — think of it like a separate mini-computer that can only see the project folder. Even if a worker tries to do something unexpected, it physically cannot access your other files, apps, or data. This is the safer option if you're not comfortable giving AI agents free rein on your machine.

**Step 3: Merge when all workers are done (coordinator)**
```bash
python3 slide_manager.py merge
# → output/deck_final.pptx
```

---

## Command Reference

Commands not covered in the walkthrough above.

### Check Status
```bash
python3 slide_manager.py status        # All slide assignments and worker files
```

### Unassign Slides
```bash
python3 slide_manager.py unassign A 10
```

### Render a Worker's Slides
Renders all slides assigned to a worker. Output goes to `renders/worker_A/` and Finder opens automatically:
```bash
python3 slide_manager.py render A
```

### Merge Specific Slides
```bash
python3 slide_manager.py merge 5 8 14
```
Unspecified slides are taken from the source deck unchanged.

### Add Slides
Add blank slides to all files at once. Without `--after`, slides are appended at the end. With `--after N`, the slide is inserted after slide N and all assignments are renumbered automatically:
```bash
python3 slide_manager.py add-slide              # Append 1 blank slide
python3 slide_manager.py add-slide --after 5    # Insert after slide 5 (becomes slide 6)
python3 slide_manager.py add-slide 3 --after 5  # Insert 3 slides after slide 5
```
This command is blocked if any workers have unsaved changes — merge, promote, and setup first.

### Promote
After reviewing the merged deck, promote it to become the new source of truth:
```bash
python3 slide_manager.py promote
# Backs up source/deck_original.pptx → source/deck_original.bak.pptx
# Copies output/deck_final.pptx → source/deck_original.pptx

# Then refresh all worker files from the new source:
python3 slide_manager.py setup
```

## File Structure

```
fawn-deck/
├── launch_workers.sh      # Launch workers locally (no sandbox)
├── worker_prompt.md       # Initial prompt sent to each worker
├── claude-vm/
│   └── run-claude.sh      # Launch workers in Docker (sandboxed)
├── slide_manager.py       # Coordination tool
├── slide_renderer.py      # PNG renderer
├── assignments.json       # Worker assignments (auto-managed)
├── BRAND_GUIDE.md         # Your design system and style rules (create this)
├── source/
│   └── deck_original.pptx # Source deck — never modify directly
├── workers/
│   ├── worker_A.pptx      # Worker A's copy
│   ├── worker_B.pptx      # Worker B's copy
│   └── worker_C.pptx      # Worker C's copy
├── renders/
│   ├── worker_A/          # Worker A's renders (slide_5.png, etc.)
│   ├── worker_B/
│   └── worker_C/
└── output/
    └── deck_final.pptx    # Merged result
```
