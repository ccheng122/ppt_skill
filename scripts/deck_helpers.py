"""Thin python-pptx wrappers and a shared palette for building slide decks.

Import these into a build script so every deck looks like one system and you do
not re-author text boxes each time. All positions are in inches on a 13.333 x 7.5
widescreen slide. See example_build.py for usage.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---- neutrals and semantic colors (a generic, reusable base) ----
INK      = RGBColor(0x15, 0x20, 0x2A)   # primary text
INK2     = RGBColor(0x48, 0x59, 0x65)   # secondary text
FAINT    = RGBColor(0x7B, 0x8B, 0x96)   # captions, footnotes
CARD     = RGBColor(0xF4, 0xF7, 0xF7)   # panel fill
BORDER   = RGBColor(0xDB, 0xE3, 0xE7)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
GOOD     = RGBColor(0x2F, 0x7A, 0x54)
GOOD_S   = RGBColor(0xDB, 0xED, 0xE2)
WARN     = RGBColor(0x9C, 0x6A, 0x15)
WARN_S   = RGBColor(0xF1, 0xE7, 0xCE)
GREY     = RGBColor(0x60, 0x6C, 0x74)

# ---- accent options ----
# Each option is an accent, a darker accent-ink for text on a soft band, a soft fill
# for callout bands, and a hex for matching charts (see chart_helpers). Pick the one
# that fits the subject or the brand the deck mirrors, or add your own. None is a
# default house style; choose deliberately per deck.
def _rgb(h): return RGBColor((h >> 16) & 255, (h >> 8) & 255, h & 255)
PALETTES = {
    "teal":    dict(accent=_rgb(0x17706B), ink=_rgb(0x0F4E4A), soft=_rgb(0xD6E8E5), hex="#17706B"),
    "indigo":  dict(accent=_rgb(0x2B4F8A), ink=_rgb(0x1B335C), soft=_rgb(0xDDE5F3), hex="#2B4F8A"),
    "plum":    dict(accent=_rgb(0x7A3B6B), ink=_rgb(0x54284A), soft=_rgb(0xEEDDE9), hex="#7A3B6B"),
    "forest":  dict(accent=_rgb(0x2F6E4F), ink=_rgb(0x204C37), soft=_rgb(0xDBEBE1), hex="#2F6E4F"),
    "rust":    dict(accent=_rgb(0xB0502E), ink=_rgb(0x7A371E), soft=_rgb(0xF3E1D7), hex="#B0502E"),
    "slate":   dict(accent=_rgb(0x3A4A57), ink=_rgb(0x26323C), soft=_rgb(0xE2E8EC), hex="#3A4A57"),
}

# Current accent. Call use_palette(...) in your build script to set it before laying
# out slides. Defaults to a neutral slate so nothing colorful is imposed by accident.
ACCENT, ACCENT_INK, ACCENT_SOFT = PALETTES["slate"]["accent"], PALETTES["slate"]["ink"], PALETTES["slate"]["soft"]


def use_palette(name):
    """Set the active accent from PALETTES (by name) or a custom dict {accent,ink,soft}."""
    global ACCENT, ACCENT_INK, ACCENT_SOFT
    p = PALETTES[name] if isinstance(name, str) else name
    ACCENT, ACCENT_INK, ACCENT_SOFT = p["accent"], p["ink"], p["soft"]
    return p

FONT = "Arial"  # portable across PowerPoint and Google Slides


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def add_slide(prs, bg=WHITE):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = bg
    return s


def rect(s, x, y, w, h, fill, line=None, round_=True):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if round_:
        try:
            shp.adjustments[0] = 0.05
        except Exception:
            pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(0.75)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def R(t, size, color=INK, bold=False, spacing=None, caps=False):
    """A styled text run tuple. spacing is letter spacing in points; caps upper-cases."""
    return (t, size, color, bold, spacing, caps)


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        sp_after=4, line_sp=1.1):
    """runs is a list of paragraphs; each paragraph is a list of R(...) tuples."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(sp_after)
        p.space_before = Pt(0)
        p.line_spacing = line_sp
        for (t, size, color, bold, spacing, caps) in para:
            r = p.add_run()
            r.text = t
            f = r.font
            f.size = Pt(size)
            f.color.rgb = color
            f.bold = bold
            f.name = FONT
            if spacing is not None:
                r._r.get_or_add_rPr().set('spc', str(int(spacing * 100)))
            if caps:
                r._r.get_or_add_rPr().set('cap', 'all')
    return tb


def pic(s, path, x, y, w):
    return s.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))


def eyebrow(s, text, kicker=None):
    """Section label at the top of a slide, with an optional right-aligned kicker."""
    rect(s, 0.55, 0.5, 0.34, 0.05, ACCENT, round_=False)
    txt(s, 0.95, 0.4, 9.0, 0.35, [[R(text, 11.5, ACCENT, True, 1.4, True)]])
    if kicker:
        txt(s, 9.0, 0.38, 3.8, 0.35, [[R(kicker, 11, FAINT, False, 0.3)]], align=PP_ALIGN.RIGHT)


def band(s, y, runs, x=0.55, w=12.23, h=0.92, fill=None):
    """Full-width highlight bar for a verdict or callout. runs is one paragraph.

    Defaults to the active ACCENT_SOFT, resolved at call time so use_palette() applies.
    """
    rect(s, x, y, w, h, fill if fill is not None else ACCENT_SOFT)
    txt(s, x + 0.3, y, w - 0.6, h, [runs], anchor=MSO_ANCHOR.MIDDLE, line_sp=1.1)
