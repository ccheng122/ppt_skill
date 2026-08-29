"""Minimal end to end example: build charts, then lay out slides.

Run: python3 example_build.py  ->  writes example_deck.pptx in this folder.
Copy this pattern into a real build script; keep the numbers inline so the deck
is reproducible. Remember: percentages as X.X%, no em dashes anywhere.
"""
import os
from pptx.enum.text import PP_ALIGN
import chart_helpers as C
import deck_helpers as D

HERE = os.path.dirname(os.path.abspath(__file__))

# ---- 0. choose a palette that fits the deck (see PALETTES in deck_helpers) ----
PAL = D.use_palette("teal")                 # or "indigo", "plum", "forest", "rust", "slate", or a custom dict
C.GROUP_COLORS["focus"] = PAL["hex"]        # match the chart accent to the deck accent

# ---- 1. build a chart from the real data ----
C.setup()
fig, ax = C.plt.subplots(figsize=(6.1, 3.2))
ax.plot([0, 1, 2, 3], [40.0, 41.8, 45.3, 34.9], "-o", color=C.GROUP_COLORS["benchmark"], lw=2, label="Prior")
ax.plot([0, 1, 2, 3], [22.2, 20.1, 23.2, 25.9], "-o", color=C.GROUP_COLORS["focus"], lw=2.8, label="This cohort")
ax.set_ylim(0, 55)
ax.set_xlabel("Weeks since launch")
ax.set_ylabel("Edit rate")
C.style(ax, pct=True)
ax.legend(frameon=False, fontsize=9.5, loc="upper right")
chart_path = os.path.join(HERE, "example_chart.png")
C.save(fig, chart_path)

# ---- 2. lay out the slides ----
prs = D.new_deck()

# Slide 1: executive summary
s = D.add_slide(prs)
D.eyebrow(s, "Example Readout", "2026")
D.txt(s, 0.55, 1.0, 12.2, 0.7, [[D.R("The headline answer, stated first", 34, D.INK, True)]])
D.txt(s, 0.55, 1.95, 12.3, 0.6,
      [[D.R("One or two sentences of context, in the source's own words.", 15, D.INK2)]], line_sp=1.2)
D.band(s, 3.0,
       [D.R("Key point.  ", 15, D.ACCENT_INK, True),
        D.R("A verdict the reader should walk away with. Rates read as 22.2%, not 0.222.", 15, D.INK)],
       h=1.1)

# Slide 2: deep dive with a chart
s = D.add_slide(prs)
D.eyebrow(s, "Deep dive")
D.txt(s, 0.55, 1.0, 12.2, 0.7, [[D.R("What the chart shows", 30, D.INK, True)]])
D.txt(s, 0.55, 1.8, 6.0, 0.3, [[D.R("Edit rate by week", 11, D.FAINT, True, 0.6, True)]])
D.pic(s, chart_path, 0.55, 2.15, 6.4)
D.rect(s, 7.15, 2.15, 5.63, 3.4, D.WHITE, line=D.BORDER)
D.txt(s, 7.4, 2.35, 5.15, 3.0,
      [[D.R("Supporting text on the right, pulled verbatim from the source narrative when one exists.", 14, D.INK2)]],
      line_sp=1.3)

out = os.path.join(HERE, "example_deck.pptx")
prs.save(out)
print("saved:", out)
