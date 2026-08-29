---
name: slide-deck
description: Build a polished, editable PowerPoint (.pptx) slide deck from analysis findings, a Hex app, a doc, or a readout. Use this whenever the user wants "slides", a "deck", a "presentation", a ".pptx", or asks to "turn this into slides", "make slides for [a stakeholder]", or "put this in a deck", even if they never say the word PowerPoint. Produces a real editable .pptx that opens in PowerPoint or Google Slides, with charts built from the actual data, structured as an exec summary plus one deep dive per point and sized to the content rather than padded into many thin slides. Do NOT use this for an HTML slide webpage or a scroll page; the user almost always means an editable deck they can present and tweak.
---

# Slide deck (.pptx) from findings

## What "slides" means here

An editable PowerPoint file, not an HTML page. When someone says "slides" or "a deck" they mean a real `.pptx` they can open in PowerPoint or Google Slides, present, and edit by hand. If you hand them a web page instead, they will say "wait, I thought we were making slides." Build the `.pptx` with `python-pptx` and embed charts as images.

## First, mirror the source. Do not invent.

A deck almost always mirrors an existing deliverable: a Hex app, a doc, or an analysis already in the conversation. Two rules save a lot of rework:

1. **Include only what is in the final source.** If a chart or finding was cut from the app or doc, it does not belong in the deck. When you are unsure whether something made the final cut, ask "is this in the final version?" rather than assuming. People notice extra content immediately and it reads as sloppy.

2. **When the user already wrote the narrative, use their exact words.** If the source has an intro, a summary, or takeaway text, pull it verbatim. Do not rewrite it into "slide voice." The user knows their own copy and rewording it, even to tighten it, tends to land badly. Copy the text; do not paraphrase it.

## Structure: fix the outline, then size to the content

Default the outline to an exec summary plus one deep dive per key point, with an optional recommendation slide at the end. That shape is almost always right.

- **Slide 1, executive summary.** The headline answer plus a few key points. Lead with the conclusion.
- **Then one deep dive per key point,** each anchored by a chart.
- **Optional final slide, recommendation or next steps.**

Do not treat a slide count as a target. The common failure is padding: many sparse slides that each carry a lone stat or a single sentence, which reads as thin and makes the deck feel longer than it is. Size the deck to the content instead:

- Every slide should carry a substantive point: a finding and the evidence for it (usually a chart plus its read).
- If a slide would hold only one line or one number, fold it into a neighbor.
- A chart with its interpretation is a full slide; a lone bullet is not.
- Fewer dense, well composed slides beat many thin ones. Most readouts land in 3 to 5 slides, but let the material decide.

The arc is: lead with the answer, earn trust with the setup, let the data land.

**Longer decks:** if the deck looks like it will run past ~8 slides or cover 3+ distinct sections, ask the user whether they want an `agenda()` slide up front and `divider()` slides between sections, don't add them unprompted. Skip asking entirely on a short, single-thread deck.

## Charts carry the story

Build charts from the real data with `matplotlib`, save as transparent PNG, and embed them as pictures. A ramp line or a comparison bar communicates far more than a table of the same numbers, and it is the reason to make slides instead of pasting text. Style every chart the same way so the deck reads as one system; `scripts/chart_helpers.py` sets the palette and axes for you.

- **Always label the y axis** with `ax.set_ylabel(...)` naming the actual metric and its unit ("Revenue ($)", "Signups", "Edit rate (%)"). An unlabeled axis reads as unfinished.
- **Large counts get truncated, not spelled out.** For any axis or number in the thousands or millions, use `chart_helpers.compact_number()` (or `style(ax, compact=True)` to format the whole axis): `1,400` becomes `1.4K`, `2,300,000` becomes `2.3M`, one decimal, matching the deck's fixed-precision style. Pass `currency=True` to prefix with `$` for a monetary metric (`$1.4K`, `$2.3M`). Use `style(ax, pct=True)` instead for a percent axis, the two are mutually exclusive.

## Use the bundled helpers

Three scripts exist so you do not rebuild text boxes and chart styling from scratch each time, and so every deck looks consistent:

- `scripts/deck_helpers.py` gives you the palette and thin `python-pptx` wrappers: `new_deck()`, `add_slide()`, `rect()`, `txt()`, `R()` (a styled text run), `pic()` (embed an image), `eyebrow()` (section label), `band()` (a full-width highlight bar for a verdict or callout), `table()` (with a default grey outline, see below), `chart_card()` (a white shadowed frame behind a chart image, see below), and `section_bar()` (an optional rotated left section rail, see below).
- `scripts/chart_helpers.py` gives you matplotlib defaults, the same accent colors, a `style(ax)` cleaner, `GROUP_COLORS` for consistent series colors, and `save(fig, path)` for transparent export.
- `scripts/measure.py` gives you real text-wrap measurement (`block_height()`, `assert_fits()`) so box heights come from the actual copy at the actual width instead of a guess. See "Verify layout before you hand it off" below; this is the difference between a deck that looks right in your head and one that actually renders without collisions.

Write a short build script that imports these, generates the chart PNGs, then lays out the slides. Keep the numbers and chart source in that script so the deck is reproducible. See `scripts/example_build.py` for a minimal end to end pattern.

**Chart cards, optional.** A chart looks more finished sitting on its own white card with a soft shadow than floating loose on the slide background (`chart_card(x, y, w, h)`, called right before `pic()` at the same `x, y, w`). If a highlight band or callout sits directly under one or more chart cards, size that band to the cards' *combined outer footprint* (each card's `pad` included), not the slide's full content width, or the edges visibly don't line up, which reads as sloppy even when nothing actually overlaps.

**Left section rail, optional, ask first.** `section_bar(s, label)` draws a full-height rotated color bar on the left edge, labeling which part of the story a slide is (a real pattern pulled from a client deck's per-section spine). Gate it the same way as `divider()`/`agenda()`: offer it, don't add it unprompted. For a deck that's one continuous narrative (most decks here) rather than genuinely distinct sections, keep it to a single accent color throughout with the label text changing per slide, not an arbitrary color per slide, that reads as decoration standing in for a taxonomy that doesn't exist. Using it means shifting every other element's left margin right to clear the bar; see the function's docstring for the rotation gotcha (which paragraph alignment ends up looking "top aligned" once the shape rotates, since that direction is not the alignment you'd guess).

**One type scale, not a size per call site.** Pick a small fixed set of sizes up front (e.g. one headline size, 14pt for every paragraph of body copy, 11.5pt for small caption-style labels like the eyebrow and chart captions) and hold every text call to one of those roles. Reaching for "14.5 here, 13 there, 13.5 somewhere else" as each slide gets written is the natural failure mode, it happens gradually and each individual choice looks reasonable in isolation, but the deck ends up visibly inconsistent. Headlines, a hero stat number, and a small card number are legitimately their own larger sizes; body prose and captions are not.

## Style rules that are not optional

- **No em dashes, anywhere.** This is a standing preference. Use a period, comma, semicolon, colon, or parentheses instead; write numeric ranges as "X to Y". Before handing off, scan every text run for the em dash character and confirm there are zero. (Arrows and mid dots are fine; hyphens in compound words are fine.)
- **Percentages as `X.X%` for stakeholder decks.** Format rates like `8.6%`, not `0.086`, with one decimal, unless the user says otherwise. Counts stay as counts.
- **Font: Arial.** It is present in both PowerPoint and Google Slides, so it will not silently fall back the way a Google font would.
- **Palette: choose one deliberately, do not default to any single look.**
  - **Deck for a specific company or brand:** pull their real primary brand color (from what they've shared, their site, or their style guide, don't guess) and derive the deck's accent from it with `deck_helpers.palette_from_hex("#xxxxxx")`. This keeps the deck's own hue rather than pasting a raw, unadjusted brand hex on, and it only corrects a color that would read badly at deck scale (too light or too washed out for small text). Pass the result straight to `use_palette(...)`. If the brand has several colors, use their dominant one; the deck stays one accent, it does not try to reproduce a whole brand system. If you can't confirm the exact hex (no reliable source, or the brand's identity is genuinely multicolor/gradient rather than one solid color), say so plainly and either ask the user for the exact value or fall back to a shipped palette rather than presenting a guess as confirmed.
  - **Generic or internal deck, nothing to match:** ask the user before building, don't silently pick one. Offer the neutral-leaning options first (`NEUTRAL_PALETTES`: slate, teal, forest, indigo), mention plum or rust as available if the subject calls for more warmth, and note a custom `{accent, ink, soft}` dict is also an option. Once they answer, call `use_palette(...)` with the shipped `PALETTES` name (or the custom dict) from `deck_helpers.py`. None of them is a house style, don't default to the same one every time.
  - The neutrals and the semantic good and warning colors are shared across every palette, derived or shipped. To keep charts consistent, set the chart accent to the same palette's `hex` (`GROUP_COLORS["focus"] = PAL["hex"]`). Keep it restrained: one accent, considered neutrals, and semantic good and warning reserved for verdict callouts.

## Verify layout before you hand it off

Guessing a box height and eyeballing whether text "probably fits" is how a headline ends up wrapping one line more than planned and colliding with the subtitle underneath it, or a caption starting before the text above it has actually finished. That class of bug does not show up until someone opens the file, so don't rely on eyeballing:

- As you write each text block, immediately call `measure.assert_fits(label, paragraphs, width_in, box_h, line_sp=..., sp_after_pt=...)` with the exact text, size, width, and line_sp you just gave `deck_helpers`, right after the `D.txt()`/`D.band()`/`D.bullets()` call. A box that's too short raises immediately with the shortfall, instead of failing silently. For a paragraph that mixes a bold lead-in with a regular continuation (a common callout pattern), pass a list of `(text, size, bold)` runs for that paragraph rather than one flat tuple, so the wrap width accounts for the actual mixed weights.
- Watch for wraps sitting right at a width boundary: the same copy can be one line at one width and two lines at a width just 0.05in narrower (font-metric edge cases are real, not just a theoretical risk). If a headline's line count changes with a small, plausible width tweak, don't leave it there; reword it so it wraps the same way with real margin, the same way you'd fix any other fragile edge case.
- Reopen the `.pptx` with `python-pptx` and confirm the slide count and that each chart image is embedded.
- Scan all text for the em dash character; there should be none.
- If LibreOffice is available (`soffice --headless --convert-to pdf`), render it to eyeball spacing and overlaps as a final pass on top of the measurement checks above, they catch different things (measurement catches text overflow, a render catches things measurement can't, like whether a rotated shape or an image actually lands where you meant). If it is not available, say plainly that you could not pixel check, and point the user at the busiest slide (usually the one with two charts plus text).

## Deliver it

Save the `.pptx` into the project or wherever the user keeps deliverables, not a scratch folder. Tell them the path and how to get it into Google Slides if they use that: upload to Drive, then open with Google Slides (or in Slides, File then Import slides). Keep customer names and any personal data off slides that may be shared, unless the user explicitly asks for them; refer to a dominant account as "one company (59.0% of volume)" rather than by name.
