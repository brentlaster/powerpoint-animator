# PowerPoint auto-animator

Two tools that add sensible entrance animations to a PowerPoint slide
programmatically, so you don't have to open the animation pane and click
through every shape by hand.

Given a slide with shapes, arrows, connectors, and text boxes, the tool
infers a click-by-click reveal order that follows the diagram's spatial
flow:

  * shapes **fade** in,
  * text boxes **wipe** in from the top,
  * arrows **wipe** in the direction they point,
  * if arrow `A → B` exists, the reveal order is `A`, then `arrow + B`
    (together on one click), and so on through the chain.

Two deliverables are provided so you can pick whichever fits:

  * `auto_animate.py` — cross-platform Python; edits `.pptx` XML directly.
    No PowerPoint install required; works on macOS, Linux, and Windows.
  * `AutoAnimate.bas` — VBA macro that runs inside PowerPoint itself.
    Uses PowerPoint's animation COM API directly, so it's guaranteed to
    produce exactly what PowerPoint would have produced by hand.

Both tools only touch slides you opt in to (see "the marker").

---

## The marker

By default the tool processes only slides whose **speaker notes contain
`[auto-animate]`**. Everything else is left alone. You can animate all
slides by passing `--all` (Python) or calling `AutoAnimateAll` (VBA).

Slides that already have animations are always left alone, so you never
clobber hand-built timing.

---

## auto_animate.py (Python)

### Install

    pip install python-pptx lxml

### Run

    python3 auto_animate.py input.pptx

writes `input_animated.pptx` next to the input.

Other forms:

    python3 auto_animate.py input.pptx out.pptx           # explicit output
    python3 auto_animate.py input.pptx --all              # ignore markers
    python3 auto_animate.py input.pptx --marker "[anim]"  # change marker
    python3 auto_animate.py input.pptx -q                 # quiet

The script prints, for each processed slide, the click order it chose —
useful for sanity-checking the flow inference before you open the file.

### What it writes

For each marked, un-animated slide the script builds a valid
`<p:timing>` block containing a `<p:seq nodeType="mainSeq">` with one
click wrapper per group, each effect referencing the shape's drawingML
`id`. It also appends a `<p:bldLst>` with a `<p:bldP>` entry per animated
shape. The block is inserted in the correct schema position (before any
`<p:extLst>`).

---

## AutoAnimate.bas (VBA)

### Install

1. Open your `.pptx` in PowerPoint.
2. Press **Alt+F11** to open the VBA editor.
3. `File → Import File…` and pick `AutoAnimate.bas`.
4. Close the editor.

### Run

Press **Alt+F8** and pick a macro:

  * `AutoAnimate` — animate every slide whose notes contain `[auto-animate]`.
  * `AutoAnimateAll` — animate every slide.
  * `AutoAnimateCurrent` — animate only the slide currently shown.

(Slides that already have any animations are skipped so you don't
overwrite hand-built timing.)

---

## Inference rules

1. **Shape classification**
   * connectors (line shapes) → *arrow*
   * block arrows (`RightArrow`, `LeftArrow`, `UpArrow`, `DownArrow`,
     curved and notched variants) → *arrow*
   * text boxes → *text*
   * everything else → *shape*

2. **Arrow direction**
   * connector: bbox corners adjusted by `flipH`/`flipV`; tail→head is
     the direction of motion.
   * block arrow: the preset geometry determines direction; `flipH`/
     `flipV` reverse it.

3. **Graph edges**
   For each arrow, its tail endpoint is matched to the smallest
   non-arrow shape whose bounding box contains it (5-pt padding). If
   nothing contains it, the nearest centre within ~2.3 inches wins.
   Same for the head.

4. **Click-group ordering**
   * Phase 1: BFS from every *source* (in-degree 0, out-degree > 0) in
     reading order.
   * Phase 2: BFS from any remaining node that has any flow edge (this
     handles cycles).
   * Phase 3: dangling arrows (no detected src or dst) are appended in
     reading order.
   * Phase 4: orphan text/shapes are split: those above the topmost
     in-flow element animate *before* the flow; those below animate
     *after*. This usually puts titles first and footnotes last.

5. **Per-shape effect**
   * *shape* → Fade, 400 ms
   * *text* → Wipe down (top-to-bottom), 500 ms
   * *arrow* → Wipe in the arrow's direction, 600 ms (700 ms in Python)

6. **Click grouping**
   Each group fires on one click. The first effect in the group is a
   click-triggered effect; any additional effects (typically `arrow +
   destination shape`) run *with previous* so they play together.

---

## Known limitations

  * **Grouped shapes** inside `<p:grpSp>` are animated as a single block
    — the group's members aren't walked into. Un-group in PowerPoint
    first if you want per-element reveals.
  * **Curved / elbow connectors** are still treated as straight for
    direction detection: motion is inferred from the bounding box's two
    diagonal corners, which is fine in practice but is a rough
    approximation.
  * **Already-animated slides** are left alone entirely. If you want to
    *add* to existing timing, remove the existing animations first.
  * **Text paragraphs-by-paragraph** isn't supported; each text box
    wipes as a single unit.
  * Pictures (`<p:pic>`) are animated as plain shapes (Fade).

---

## Files in this folder

    auto_animate.py        the Python tool
    AutoAnimate.bas        the VBA macro
    build_demo_deck.py     script that generated demo_deck.pptx
    demo_deck.pptx         demo deck with 4 slides (3 marked, 1 unmarked)
    demo_deck_animated.pptx result of running the Python tool on the demo
    README.md              this file
