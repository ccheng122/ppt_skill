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
- [x] **Palette silently picked, never surfaced.** Caught by a real cold-start
      test (a fresh session, no conversation memory of this backlog): on a
      generic/internal deck, `SKILL.md` said "choose deliberately" but never
      said to tell or ask the user, so it just silently picked one. Fixed:
      the generic/internal branch now asks the user before building, same
      gate as divider()/agenda(). The brand-derived branch is unchanged,
      asking isn't needed there since the color comes from a real fact
      (the brand's own hex), not an aesthetic pick.
- [x] **Real text-measurement validation**, `scripts/measure.py`
      (`block_height()` / `assert_fits()`). Built after eyeballed box heights
      in a real deck (Clara's Suno take-home readout) produced three actual
      collisions: a headline that wrapped one line more than planned into the
      subtitle below it, a 40pt stat number whose caption started before the
      number had finished rendering, and a chart-card's shadow padding
      touching the callout underneath it. Also caught two false positives in
      the first draft of the checker itself (comparing a rotated shape's
      pre-rotation bounding box raw, and measuring a bold-lead-in paragraph
      as if the whole line were bold) before `assert_fits()` reached its
      current form. `SKILL.md`'s validation section now leads with this
      instead of only LibreOffice.
- [x] **`chart_card()` / `add_soft_shadow()`**, `deck_helpers.py`. A white
      rounded card with a soft drop shadow behind a chart image, spec pulled
      from a real reference deck's own chart treatment (9pt blur, 1.5pt
      downward offset, black at 20% alpha), not invented. Documented rule
      that came out of using it: a callout/band placed under the chart(s)
      needs to match their *combined outer footprint*, not the slide's full
      content width, or the edges visibly don't line up (caught as direct
      user feedback: "doesn't look visually good").
- [x] **`section_bar()`**, `deck_helpers.py`. Optional rotated left-edge
      section rail, also pulled from a real reference deck (a spine of
      rotated, per-section-colored rectangles). Gated like divider()/agenda(),
      not a default. For a single-narrative deck (most decks here), defaults
      to one accent color with a per-slide label rather than inventing
      per-slide colors with no real taxonomy behind them, a recommendation
      the user explicitly asked for and approved. Getting the rotation
      direction right (which paragraph alignment reads as "top" once the
      shape rotates -90) took reverse-engineering the reference deck's own
      XML; that reasoning is captured in the function's docstring so it
      isn't re-derived next time.
- [x] **`table()` gets a default border.** A table with no explicit line
      color visually disappears into a white slide background between its
      white body rows and the page, a real user-reported bug ("white rows
      kinda drown into the white background"). Fixed at the source in
      `table()` itself (a thin grey outline rect drawn behind it) rather than
      requiring every caller to remember to draw one.
- [x] **One type scale, not a size per call site.** A real build session let
      body-copy sizes drift to 14.5/13/13.5/12.5pt across different slides,
      each choice locally reasonable but the deck reading as inconsistent
      overall (direct user feedback). `SKILL.md` now states the fix as a
      rule: pick a small fixed scale up front (one headline size, one body
      size, one caption size) and hold every call to one of those roles.

## Now

(empty, everything on this backlog is built and approved)

## Later

(nothing queued, add here as new gaps show up in real use)

## Note

Every helper added is more surface area to keep visually consistent with the
"one system" rule in SKILL.md. Default to restraint; only add what a real
deck has actually needed.
