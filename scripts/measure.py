"""Text-measurement helper: estimates real wrapped-line height using DejaVu
Sans (matplotlib-bundled, so no extra dependency) as a conservative proxy for
Arial, so box heights come from actual text length instead of a guess.

Why this exists: eyeballing a box height (or asking "does this look about
right?") reliably produces slides where a headline wraps one line more than
expected and collides with the subtitle below it, or a caption starts before
the text above it has actually finished rendering. None of that shows up
until someone opens the file. Measuring first catches it before the file is
ever written.

Use it two ways:

1. Before laying out a slide, call block_height() on the real copy at the
   real width to get an actual required height, and size the box to that
   (with a little padding) instead of a round number.
2. Right after every D.txt()/D.band()/D.bullets() call, call assert_fits()
   with the same text/size/width/box-height you just used, so a future
   copy edit that no longer fits fails loudly at build time instead of
   silently overlapping the next element. See example_build.py.

DejaVu Sans runs slightly wider than Arial at most weights, so a fit here is
a safe fit in the real deck; it is not exact, so don't shave the resulting
number to the decimal, leave a little padding, and be extra cautious around
any wrap that looks close to a line-count boundary (small width or copy
changes can flip it, see the class of bug this module exists to catch).
"""
import functools
import os

import matplotlib
from PIL import ImageFont

FONT_REG = os.path.join(matplotlib.get_data_path(), "fonts", "ttf", "DejaVuSans.ttf")
FONT_BLD = os.path.join(matplotlib.get_data_path(), "fonts", "ttf", "DejaVuSans-Bold.ttf")
DPI = 96


@functools.lru_cache(maxsize=None)
def _font(size_pt, bold):
    return ImageFont.truetype(FONT_BLD if bold else FONT_REG, int(size_pt * DPI / 72))


def _as_runs(paragraph):
    """Normalize one paragraph entry to a list of (text, size_pt, bold) runs.
    Accepts either a single (text, size_pt, bold) tuple (uniform paragraph)
    or a list of such tuples (mixed runs, e.g. a bold lead-in followed by a
    regular continuation, which is the deck's most common callout pattern:
    R("Bottom line.  ", 15, ACCENT_INK, True), R(body_text, 15, INK))."""
    if paragraph and isinstance(paragraph[0], str):
        return [paragraph]
    return list(paragraph)


def wrap_lines(text, size_pt, bold, width_in):
    """Single-run convenience wrapper; see wrap_runs for the mixed-run case."""
    return wrap_runs([(text, size_pt, bold)], width_in)


def wrap_runs(runs, width_in):
    """Word-wrap a flowing paragraph built from one or more (text, size_pt,
    bold) runs, using each word's own run font. Wrapping the whole paragraph
    with only the first run's weight (a natural shortcut) overstates the
    width whenever a bold lead-in is followed by longer regular text, which
    can report a line count one higher than what PowerPoint will actually
    render."""
    max_w_px = width_in * DPI
    tokens = []  # (word, font)
    for text, size_pt, bold in runs:
        font = _font(size_pt, bold)
        for w in text.split(" "):
            if w:
                tokens.append((w, font))
    if not tokens:
        return []
    space_w = _font(runs[0][1], False).getbbox(" ")[2]
    lines, cur, cur_w = [], [], 0
    for word, font in tokens:
        bbox = font.getbbox(word)
        ww = bbox[2] - bbox[0]
        add_w = ww + (space_w if cur else 0)
        if cur_w + add_w <= max_w_px or not cur:
            cur.append(word)
            cur_w += add_w
        else:
            lines.append(" ".join(cur))
            cur, cur_w = [word], ww
    if cur:
        lines.append(" ".join(cur))
    return lines


def block_height(paragraphs, width_in, line_sp=1.1, sp_after_pt=4):
    """Height in inches of one or more paragraphs at the given width.

    paragraphs is a list where each item is either a single (text, size_pt,
    bold) tuple, or a list of such tuples for a paragraph with mixed runs
    (see _as_runs). line_sp and sp_after_pt should match the line_sp/sp_after
    you're passing to deck_helpers.txt() for the same text.
    """
    total = 0.0
    for paragraph in paragraphs:
        runs = _as_runs(paragraph)
        lines = wrap_runs(runs, width_in)
        max_size = max(size for _, size, _ in runs) if runs else 0
        line_h_pt = max_size * 1.2 * line_sp
        total += (line_h_pt * len(lines) + sp_after_pt) / 72.0
    return total


def assert_fits(label, paragraphs, width_in, box_h, line_sp=1.1, sp_after_pt=4):
    """Raise with a clear message if paragraphs need more height than box_h
    at width_in. Call this with the exact same text/size/bold/width/line_sp
    you just handed to deck_helpers, right after laying out each text block,
    so a build fails loudly instead of shipping a silent overlap."""
    needed = block_height(paragraphs, width_in, line_sp=line_sp, sp_after_pt=sp_after_pt)
    assert needed <= box_h, f"{label}: needs {needed:.2f}in, only {box_h:.2f}in allocated"
