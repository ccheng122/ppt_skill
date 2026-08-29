# slide-deck improvement backlog

Ranked by how often the gap actually bites in practice. Tackle one by one;
check off and add a one-line note on what shipped.

## Done

- [x] **Palette derived from a brand color**, `palette_from_hex()`, for a
      company-specific deck; `NEUTRAL_PALETTES` / `EXPRESSIVE_PALETTES` for
      picking a shipped one when there's no brand to match.
- [x] **Chart axis: compact large numbers** (`1.4K`, `2.3M`, optional `$`
      prefix) via `chart_helpers.compact_number()` / `style(compact=True)`,
      plus a standing rule to always label the y axis.
- [x] **Section divider + agenda slides**, `divider()` / `agenda()` in
      `deck_helpers.py`. Confirmed by real usage, not speculative: both
      appeared in two of Clara's actual past decks (a roman-numeral divider
      slide, and a numbered "01 / 02 / 03" agenda page). Only offered past
      ~8 slides or 3+ sections, and only by asking first, per SKILL.md.
- [x] **Bullet-list support**, `bullets()` in `deck_helpers.py`. One idea per
      line (full sentences fine, no need to over-clip into fragments); no
      semicolon-spliced run-ons or raw internal notation shown unexplained.
      Approved 2026-08-29 against a rewrite of the real run-on bullet from
      `Uber Eats_v7.pptx` slide 6.
- [x] **Table helper**, `table()` in `deck_helpers.py`. Accent header row,
      alternating white/card body rows, optional relative `col_widths`.
- [x] **Stat group**, `stat_group()` in `deck_helpers.py`. Rebuilt from the
      real pattern in `Product Prioritization.pptx` slide 6 (value over a
      one-line label), no icons in this version. `direction="vertical"`
      (default, stacked beside a chart) or `"horizontal"` (across the top of
      a slide) per Clara's note that it shouldn't always be a column.
- [x] **Two-chart comparison**, `two_up()` in `deck_helpers.py`. The one
      speculative item on this list, approved anyway 2026-08-29. Even 50/50
      split, caption above each chart. Watch in real use whether an even
      split actually fits, or one side usually needs more room.

## Now

(empty, everything on this backlog is built and approved)

## Later

(nothing queued, add here as new gaps show up in real use)

## Note

Every helper added is more surface area to keep visually consistent with the
"one system" rule in SKILL.md. Default to restraint; only add what a real
deck has actually needed.
