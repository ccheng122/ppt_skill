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


def compact_number(v, currency=False):
    """Format a large number as 1.4K / 2.3M, one decimal, matching the deck's
    fixed-precision style (see SKILL.md's X.X% rule). Prefix with $ when
    currency=True. Small values (under 1000) print plain, still one decimal
    if not a whole number."""
    sign = "-" if v < 0 else ""
    v = abs(v)
    if v >= 1_000_000:
        scaled, suffix = v / 1_000_000, "M"
    elif v >= 1_000:
        scaled, suffix = v / 1_000, "K"
    else:
        scaled, suffix = None, None

    if suffix:
        scaled = round(scaled, 1)
        if scaled >= 1000 and suffix == "K":  # e.g. 999,999 rounds to 1000.0K, bump to 1.0M
            scaled, suffix = scaled / 1000, "M"
        s = f"{scaled:.1f}{suffix}"
    else:
        s = f"{v:.0f}" if float(v).is_integer() else f"{v:.1f}"
    return f"{sign}{'$' if currency else ''}{s}"


def style(ax, pct=False, compact=False, currency=False):
    """Drop top/right spines, add a light y grid, and format the y axis.

    Set pct=True for a percent axis, or compact=True for large counts
    (1.4K, 2.3M); add currency=True with compact for a $ prefix. Always pair
    this with ax.set_ylabel(...) naming the actual metric and its unit, an
    axis with no label is a common review catch.
    """
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="y", color=GRID, linewidth=0.7, alpha=0.8)
    ax.set_axisbelow(True)
    if pct:
        ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    elif compact:
        ax.yaxis.set_major_formatter(lambda v, _: compact_number(v, currency=currency))


def save(fig, path):
    fig.tight_layout(pad=0.4)
    fig.savefig(path, transparent=True)   # transparent so it sits cleanly on the slide
    plt.close(fig)
