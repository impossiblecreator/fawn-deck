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

**Step 2: Open a terminal per worker and launch Claude Code**

Open three separate terminal tabs or windows. In each one, launch Claude Code with a different `WORKER_ID`:

```bash
# Terminal 1 — Worker A
WORKER_ID=A claude --dangerously-skip-permissions

# Terminal 2 — Worker B
WORKER_ID=B claude --dangerously-skip-permissions

# Terminal 3 — Worker C
WORKER_ID=C claude --dangerously-skip-permissions
```

> `--dangerously-skip-permissions` lets the worker edit files and run commands without asking you to approve each one. This is safe here because workers only touch their own `worker_X.pptx` file.


**Step 3: Prompt each worker instance**

Paste this into each Claude Code session:

```
You are a slide designer working on a PowerPoint deck.

Start by running:
  python3 slide_manager.py whoami

This will show your worker letter, your working file, and your assigned slides.
Do not guess your identity — whoami is the source of truth.

If you have a BRAND_GUIDE.md, read it fully before touching any slide.

For EACH slide, follow this design loop:
  a. Render the current slide and audit it in writing
  b. Write a Design Brief (layout, reading order, element positions) — NO CODE YET
  c. Only after the brief is written, implement with python-pptx
  d. Render, compare against the brief, iterate until 8/10

Do not skip the brief step. Slides without a written brief before coding are not acceptable.

When finished, merge ONLY your assigned slides using the exact command whoami printed:
  python3 slide_manager.py merge [your slide numbers]
```

**Step 4: Merge when all workers are done (coordinator)**
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
