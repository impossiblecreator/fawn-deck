# Fawn Friends Deck — Multi-Worker Slide Design

I was using an AI deck building product and was frustrated with how slow the process was. So, I created an orchestrator that allows me to build slides with three agents at once. Much faster.

---

## Command Reference

```bash
python3 slide_manager.py status              # All slide assignments and worker files
python3 slide_manager.py assign A 1 2 3      # Assign slides to a worker
python3 slide_manager.py unassign A 10       # Remove a slide assignment
python3 slide_manager.py merge               # Merge all assigned slides
python3 slide_manager.py merge 5 8 14        # Merge specific slides only
python3 slide_manager.py promote             # Promote final deck to source (validates first)
python3 slide_manager.py setup               # Refresh worker files from source (validates first)
python3 slide_manager.py render A            # Render Worker A's slides to PNG (opens Finder)
python3 slide_manager.py validate            # Manually validate output/deck_final.pptx
python3 slide_manager.py add-slide           # Append a blank slide to all files
python3 slide_manager.py add-slide --after 5 # Insert a blank slide after slide 5
python3 slide_manager.py whoami              # Show this worker's identity and assignments
```

```bash
./launch_workers.sh              # Launch all workers (A, B, C) in new terminal windows
./launch_workers.sh A B          # Launch specific workers only
./claude-vm/run-claude.sh        # Launch all workers sandboxed in Docker
./claude-vm/run-claude.sh A B    # Launch specific workers sandboxed
```

`add-slide` is blocked if any workers have unsaved changes — merge, promote, and setup first.

---

## How It Works

Instead of one AI working on slides one at a time, this system splits the deck across three AI workers running in parallel — each in its own terminal window, each responsible for a different set of slides. When they're done, you merge everything back into a single deck.

The source deck is never edited directly. At the start of each round, each worker gets their own copy of the deck to work in. They make their changes, you merge those changes in, and then the merged deck becomes the new source for the next round.

Here's the full cycle:

**Setup** — Before the first round, you copy your deck into the project and run setup. This creates three worker copies of the deck, one per AI.
```bash
cp your_deck.pptx source/deck_original.pptx
python3 slide_manager.py setup
```

**Assign** — You tell the system which slides each worker is responsible for. A slide can only belong to one worker, so there are no conflicts.
```bash
python3 slide_manager.py assign A 1 2 3 4
python3 slide_manager.py assign B 5 6 7 8
python3 slide_manager.py assign C 9 10 11 12
```

**Launch** — You start the three AI workers. Each one reads your brand guide and slide assignments, then gets to work on its own copy of the deck.
```bash
./launch_workers.sh          # local
./claude-vm/run-claude.sh    # sandboxed in Docker
```

**Merge** — When the workers are done, you merge their slides back into the output deck. The system validates the result automatically — if anything is broken, the merge is blocked and you get a clear error message. The worker files are never touched by a merge, so you can fix and re-merge safely.
```bash
python3 slide_manager.py merge
```

**Promote** — Once you're happy with the merged deck, you promote it to become the new source. The system validates the deck before promoting, so a corrupted file can never overwrite your source.
```bash
python3 slide_manager.py promote
```

**Setup again** — To start a new round of editing, you run setup again. This refreshes all three worker files from the newly promoted source, validates the source first, and backs up the old worker files before overwriting them.
```bash
python3 slide_manager.py setup
```

---

## First-Time Setup

Copy your deck into place, then create the three worker files:

```bash
cp your_deck.pptx source/deck_original.pptx
python3 slide_manager.py setup
```

Create a `BRAND_GUIDE.md` in the project root with your fonts, colors, and design rules. Workers will read it before editing any slide.

---

## Each Editing Round

**1. Assign slides to workers**

Decide which slides each worker will handle and register the assignments. Each slide can only go to one worker — the system will reject any conflicts.

```bash
python3 slide_manager.py assign A 1 2 3 4
python3 slide_manager.py assign B 5 6 7 8
python3 slide_manager.py assign C 9 10 11 12
```

**2. Launch the workers**

This opens terminal tabs and starts the AI workers — one per worker ID. Each worker reads the brand guide and their assignments, then edits only their own slides in their own file.

```bash
./launch_workers.sh          # Launch all workers (A, B, C)
./launch_workers.sh A B      # Launch specific workers only
```

For a sandboxed option where each worker runs in an isolated Docker container and physically cannot access anything outside the project folder:

```bash
./claude-vm/run-claude.sh        # Launch all workers sandboxed
./claude-vm/run-claude.sh A B    # Launch specific workers sandboxed
```

**3. Merge when workers are done**

This pulls each worker's assigned slides into the output deck. Validation runs automatically — if the result has broken references, missing images, or structural errors, you'll get a clear error message and the merge is blocked. Fix the issue in the worker file and re-merge.

```bash
python3 slide_manager.py merge
# → output/deck_final.pptx
```

**4. Review and promote**

Open `output/deck_final.pptx` and check the result. When you're happy, promote it to become the new source. The system validates the deck before doing anything — if validation fails, the promotion is blocked and your source is left untouched.

```bash
python3 slide_manager.py promote
```

**5. Set up for the next round**

This refreshes the three worker files from the newly promoted source, ready for the next round of editing. The source is validated before copying, and existing worker files are backed up automatically.

```bash
python3 slide_manager.py setup
```

---

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

---

## File Structure

```
fawn-deck/
├── launch_workers.sh      # Launch workers locally (no sandbox)
├── worker_prompt.md       # Initial prompt sent to each worker
├── claude-vm/
│   └── run-claude.sh      # Launch workers in Docker (sandboxed)
├── slide_manager.py       # Coordination tool
├── slide_renderer.py      # PNG renderer
├── pptx_safe_ops.py       # Validation and rollback pipeline
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
