# Fawn Friends Deck — Multi-Worker Slide Design

## Quick Start — Run These First

```bash
# 1. Confirm your identity, file, and assigned slides
python3 slide_manager.py whoami

# 2. Read the brand and design rules
# Open BRAND_GUIDE.md and read it fully before touching any slide
```

If `whoami` errors with "WORKER_ID not set", you were not launched correctly.
Tell the coordinator — do not guess your worker ID.

---

## Your Worker Identity

Your identity comes from the `WORKER_ID` environment variable set when Claude Code was launched.
Run `whoami` to see your worker letter, your working file, and your assigned slides.

**Only edit your assigned slides. Only edit your own worker file.**

---

## Tools

### Check Assignments
```bash
python3 slide_manager.py status        # All workers and their slides
python3 slide_manager.py whoami        # Just your identity and assignments
```

### Working with Images

**Never distort image aspect ratios.** In python-pptx, set width OR height — not both:

```python
# Correct — other dimension adjusts automatically
pic.width = Inches(3)

# WRONG — forces both dimensions and stretches the image
pic.width = Inches(3)
pic.height = Inches(2)
```

If a crop is needed to fit an image into a space (e.g. portrait photo into a square slot), **do not crop in code** — ask the user:

> "Slide X needs [image] cropped to [ratio]. Can you provide a cropped version, or should I adjust the layout to fit the original proportions?"

### Render Slides
```python
from slide_renderer import render_slide

# Output path is auto-inferred from the pptx filename — no need to specify it
render_slide('workers/worker_A.pptx', slide_num=5)
# → renders/worker_A/slide_5.png  (always overwrites)
```

Always view the PNG with the Read tool after rendering.

**Renderer:** LibreOffice is required. There is no fallback — if LibreOffice is missing,
rendering will fail with a clear error and install instructions. Do not attempt to work
around a render failure; fix the missing dependency first.

### Slide Manager
```bash
python3 slide_manager.py status              # See all assignments
python3 slide_manager.py render A            # Render all of Worker A's slides (opens Finder)
python3 slide_manager.py assign A 4 5 10     # Assign slides to a worker (coordinator only)
python3 slide_manager.py merge 5 14 21       # Merge ONLY your slides — always pass slide numbers
```

---

## Design Loop (follow for every slide)

See **BRAND_GUIDE.md** for the full design loop. Summary:

1. **Render** your slide and view the PNG
2. **Write a Design Brief** (layout, reading order, element positions) — no code yet
3. **Implement** with python-pptx, following the brief
4. **Re-render**, compare against the brief, iterate until 8/10
5. **Merge your slides only**: `python3 slide_manager.py merge [your slide numbers]`

Never deliver a slide you haven't visually verified.

---

## Brand Guidelines

**Read `BRAND_GUIDE.md` before editing any slide.** Key rules:

- **Font**: Neue Haas Grotesk Display — only this font, everywhere
- **Heading color**: `#123C33`
- **Body color**: `#56534F`
- **Background**: `rgb(246, 241, 233)`
- **NEVER use**: Arial, Helvetica, Roboto, Calibri, black text, or non-brand colors

---

## File Structure
```
fawn-deck/
├── CLAUDE.md               ← you are here
├── BRAND_GUIDE.md          ← read before editing
├── slide_renderer.py       ← render slides to PNG
├── slide_manager.py        ← whoami, status, assign, merge, render
├── source/
│   └── deck_original.pptx  ← NEVER modify this
├── workers/
│   ├── worker_A.pptx       ← Worker A's file
│   ├── worker_B.pptx       ← Worker B's file
│   └── worker_C.pptx       ← Worker C's file
├── renders/
│   ├── worker_A/           ← Worker A's renders
│   ├── worker_B/
│   └── worker_C/
└── output/
    └── deck_final.pptx     ← merged result
```
