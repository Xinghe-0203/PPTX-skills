# PPTX quality checklist

## Automated checks

- Open the generated file with `python-pptx`.
- Run `auto_validate_ppt()` for layout diversity, adjacent layout repetition, shape counts, text density, font hierarchy, and color count.
- Render the deck at 150 DPI or higher.
- Confirm the number and dimensions of rendered pages.

## Visual checks for every page

- No text overflow, clipping, overlap, or unintended line wrapping.
- No stretched images or irrelevant crops.
- Titles, body text, labels, and footnotes have clear hierarchy.
- Text/background contrast remains readable on projected screens.
- Page number, logo, footer, and safe margins are consistent.
- Charts and tables remain editable and legible.
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

## Reference-deck checks

- Compare generated pages with the corresponding exemplar pages side by side.
- Confirm fonts and brand colors exist on the target machine.
- Verify all explicit shape replacements from the plan.
- Open the final file in PowerPoint when it contains SmartArt, charts with embedded workbooks, OLE objects, media, transitions, or animation.

## Iteration rule

Perform at least two render-and-inspect rounds. Fix issues immediately and re-render all affected pages. Deliver only after two consecutive inspection rounds find no blocking visual defects.
