---
name: pptx-render
description: Render a .pptx (or .potx) deck to per-slide PNG images, and optionally a PDF, so slides can actually be looked at. Use for visual QA of a deck — checking text overflow, overlaps, alignment, contrast, leftover placeholders — and whenever you need to see what a generated or edited presentation really looks like. Triggers: "render the deck", "show me the slides", "check the slides look right", "PNG/PDF of the pptx".
---

# PPTX render / visual QA

Slides are invisible until rendered. Never declare a deck done on the strength of
the generating code: render it, look at every image, fix, re-render the changed
slides.

```bash
python "$SKILL_DIR/render.py" deck.pptx                 # render/slide-01.png ...
python "$SKILL_DIR/render.py" deck.pptx -o qa --pdf     # also qa/deck.pdf
python "$SKILL_DIR/render.py" deck.pptx --slides 3,7-9  # only what changed
```

`SKILL_DIR` is this skill's directory (resolve it, do not assume). The script
prints absolute paths — pass those to the image viewer as-is.

## Engines

| Order | Engine | Notes |
|---|---|---|
| 1 | **PowerPoint** (COM, `pywin32`) | Best fidelity: the user's real fonts, real layout engine. Default on Windows when Office is installed. |
| 2 | **LibreOffice** (`soffice --headless`) + PyMuPDF | Fallback. Font substitution can change text widths, so overflow checks are approximate — leave ~10% slack in containers. |

If neither is present the script exits non-zero. Do not paper over that with a
screenshot of something else; say the visual check cannot be done.

- **PowerPoint is single-instance.** The script uses `DispatchEx` and closes only
  the presentation it opened, but if the user has PowerPoint open, COM may route
  into their instance. Close decks you opened; do not call `Quit` on your own.
- Export is locale-named (`Slide1.PNG` / `Слайд1.PNG`); the script renames to
  `slide-NN.png` and warns if the PNG count does not match the slide count.
- Legacy `.ppt` must be converted to `.pptx` first.
- Output dir defaults to `render/`; old `slide-*.png` there are deleted unless
  `--keep-old`.

## What to look for (in this order)

1. **Text overflow / text cut off** at a box or slide edge — most common, always visible.
2. Overlapping elements: text through shapes, lines through words, stacked cards.
3. Footers, page numbers and citations colliding with content.
4. Gaps under ~0.3" between blocks, or margins under ~0.5" from slide edges.
5. Uneven rhythm: one crowded area, one empty one.
6. Columns and repeated elements not aligned to a common edge.
7. Low-contrast text or icons (light on light, dark on dark).
8. Leftover placeholders, `TODO`, lorem ipsum, template text in a stale language.
9. Titles wrapped to two lines where the decoration below assumed one.

Then check the content itself:

```bash
python -m markitdown deck.pptx | head -80          # text per slide, in order
python -m markitdown deck.pptx | grep -iE "\bx{3,}\b|lorem|ipsum|\bTODO|\[insert"
```

## Loop discipline

- Render → inspect → fix the source (generator script, `outline.json`, or the
  slide XML) → rebuild → re-render **only the changed slides**.
- Fix the source, never the packed XML by hand, when a generator produced the deck.
- Two or three passes is normal; stop when the images are clean, not when they
  are "good enough to explain".
- A deck that opens in python-pptx or LibreOffice is not proof it opens in
  PowerPoint — if PowerPoint export fails, that is a real signal about the file.
