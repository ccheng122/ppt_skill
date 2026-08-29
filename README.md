# slide-deck

A Claude Code skill that turns analysis findings, a Hex app, or a doc into an
editable PowerPoint (`.pptx`) deck, with charts built from the real data.

## What it does

Given a set of findings or a source deliverable, it builds a real, editable deck
(opens in PowerPoint or Google Slides): an executive summary plus one deep dive per
point, each anchored by a chart. It mirrors the source (only what made the final cut,
in the source's own words), sizes the deck to the content instead of padding it into
thin slides, and follows a few fixed style rules (percentages as `X.X%`, no em dashes,
a deliberately chosen accent palette).

## What's inside

- `SKILL.md` the workflow and rules.
- `scripts/deck_helpers.py` palette options plus thin `python-pptx` wrappers.
- `scripts/chart_helpers.py` matplotlib defaults so charts match the deck.
- `scripts/example_build.py` a minimal end to end pattern.

## Requires

`python-pptx` and `matplotlib`.

## Use

Place this folder in your Claude skills directory, then ask for "slides", "a deck",
or to "turn this into a presentation".
