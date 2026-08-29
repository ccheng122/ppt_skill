"""Matplotlib defaults and a shared palette so charts match the deck.

Colors align with deck_helpers.py. Import setup() once, then style(ax) each axis,
and use GROUP_COLORS to keep the same series colored the same way across charts.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

ACCENT = "#17706B"  # placeholder; match this to the deck accent in deck_helpers.py
SLATE  = "#6B7C88"  # a neutral "prior / benchmark" series
GREEN  = "#2F855A"  # a positive / context series
INK    = "#15202A"
INK2   = "#485965"
FAINT  = "#7B8B96"
GRID   = "#DBE3E7"

# Give the most important series the accent, comparators the neutral slate, and a
# context series the green. Reuse the same key across every chart so a series keeps
# its color. Set ACCENT to match the deck's chosen accent, do not assume this sample.
GROUP_COLORS = {"focus": ACCENT, "benchmark": SLATE, "context": GREEN}


def setup():
    plt.rcParams.update({
        "font.family": "DejaVu Sans",   # matplotlib always has this; the deck text uses Arial
        "font.size": 11,
        "text.color": INK,
        "axes.edgecolor": GRID,
        "axes.labelcolor": INK2,
        "xtick.color": INK2,
        "ytick.color": INK2,
        "axes.linewidth": 0.8,
        "figure.dpi": 200,
    })


def style(ax, pct=False):
    """Drop top/right spines, add a light y grid. Set pct=True for a percent y axis."""
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", color=GRID, linewidth=0.7, alpha=0.8)
    ax.set_axisbelow(True)
    if pct:
        ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")


def save(fig, path):
    fig.tight_layout(pad=0.4)
    fig.savefig(path, transparent=True)   # transparent so it sits cleanly on the slide
    plt.close(fig)
