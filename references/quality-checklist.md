# PPTX quality checklist

## Automated checks

- Open the generated file with `python-pptx`.
- Run `auto_validate_ppt()` for layout diversity, adjacent layout repetition, shape counts, text density, font hierarchy, and color count.
- Run `SemanticQAEngine` (adaptive pipeline) for extended checks: table column overflow, ragged rows, missing headers; chart empty series, single-category, label density; slide-level safe-margin violations, low whitespace, missing title, excessive elements; deck-level font-size drift, layout repetition, missing section breaks, color-palette drift; typography hierarchy (title/body size ratio).
- Render the deck at 150 DPI or higher.
- Confirm the number and dimensions of rendered pages.

## Visual checks for every page

- No text overflow, clipping, overlap, or unintended line wrapping.
- No stretched images or irrelevant crops.
- Titles, body text, labels, and footnotes have clear hierarchy.
- Text/background contrast remains readable on projected screens.
- Page number, logo, footer, and safe margins are consistent.
- Charts and tables remain editable and legible. Verify chart data series are non-empty and categories are sufficient.
- Empty placeholders and old reference content are removed.
- No default AI-style decoration: repeated rounded cards, decorative circles, gradient glows, floating shadows, `VS` badges, oversized quotation marks, or meaningless English kickers.

## Deck-level checks

- The story follows a clear sequence: context, evidence, decision, action.
- Information density varies; not every page is a bullet list.
- Adjacent pages avoid identical composition unless the reference template explicitly requires it.
- Different template profiles change geometry and reading order, not only palette and fonts.
- At least half of content pages use relevant visual evidence when the topic supports it.
- Cover, section dividers, and closing page form one visual family.
- The result follows the requested template mode and aspect ratio.

## Adaptive pipeline checks (PR4+)

- When using `run_generation_pipeline`, inspect `GenerationResult.qa_status` and `repair_log` rather than only the file path.
- In `strict` mode, blockers raise `PresentationQualityError`; in `report` mode they are recorded but never block delivery.
- Repair is capped at `max_repair_passes` (default 2); verify the final pass did not exceed the budget.
- For adaptive output, prefer `render_layout_plans` so each derived slide becomes its own page; check `RenderTraceEntry.slide_index` maps to the right slide.
- Pagination (PR6) splits overloaded `bullets` across derived slides with deterministic IDs; confirm split titles carry the continuation marker and that derived element IDs stay stable across re-generation.

### SemanticQAEngine checks (adaptive pipeline)

Run `SemanticQAEngine.check()` per slide and `SemanticQAEngine.check_deck()` across slides. Issue kinds:

- **Table QA**: column overflow, ragged rows, missing headers.
- **Chart QA**: empty series, single category, label density.
- **Slide-level**: safe-margin violation, low whitespace ratio, missing title, excessive element count.
- **Typography hierarchy**: title font size must be significantly larger than body (configurable `typography_min_ratio`).
- **Deck-level**: font-size drift across slides, layout repetition (3+ consecutive same-recipe slides), missing section breaks in long decks, color-palette drift.

Repair engine proposes whitelist actions (`reduce_font_within_limit`, `switch_layout_candidate`, `change_text_color_to_token`, `remove_empty_placeholder`). Verify proposed repairs are appropriate before re-rendering.

### Chart type validation

The renderer supports 10 chart types: `column_clustered`, `column_stacked`, `bar_clustered`, `bar_stacked`, `line`, `line_markers`, `pie`, `doughnut`, `scatter`, `area`. Verify the chosen `chart_type` string is in this set; unknown types fall back to `column_clustered`.

### Rich text and media checks

- Multi-run text (bold, italic, color, size, hyperlinks) must render correctly in PowerPoint. Verify runs are not merged or lost.
- Video nodes embed via `add_movie`; verify playback in PowerPoint and that poster frames display when video is not playing.
- Audio nodes embed via `add_movie` with audio MIME type; verify the audio icon appears and playback works.

### Headers, footers, slide numbers

- `deck_options` controls `show_slide_numbers`, `footer_text`, `header_text`, and `slide_number_format` (default `"{current}/{total}"`). Verify these appear on every slide in the correct positions (header top-left, footer bottom-center, page number bottom-right) and use the muted theme color.

## Reference-deck checks

- Compare generated pages with the corresponding exemplar pages side by side.
- Confirm fonts and brand colors exist on the target machine.
- Verify all explicit shape replacements from the plan.
- Open the final file in PowerPoint when it contains SmartArt, charts with embedded workbooks, OLE objects, media (video/audio), transitions, or animation.

## Iteration rule

Perform at least two render-and-inspect rounds. Fix issues immediately and re-render all affected pages. Deliver only after two consecutive inspection rounds find no blocking visual defects.
