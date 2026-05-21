# Quality Gates

Run these checks before final delivery, committing, or pushing.

## Package Checks

- Final HTML entrypoint exists and is non-empty.
- Final folder contains local CSS/JS/assets and has no hidden dependency on the reference PPTX.
- Source PPTX and user-provided files were not overwritten.
- If revising, output path is a new revision unless the user explicitly requested in-place editing.
- Slide count matches the planned outline.
- Important text is editable HTML, not baked into screenshots.
- Generated assets have provenance under `assets/generated/`.
- Generated charts have code/data/source notes when external plotting is used.

## Render Checks

Render or inspect representative slides, and for substantial decks inspect every slide or a contact sheet:

- logo sharpness and exact extracted placement
- header/footer/rule alignment
- title inside extracted title zone
- Chinese wrapping and punctuation
- no overlap among logo, title, rule, footer, citation, and main content
- table and chart readability at projected slide size
- figure captions and source notes remain legible
- no placeholder text or private presenter identity unless requested
- reduced-motion and print behavior are present

## Master-Faithful Checks

When `master-faithful` mode is active:

- Compare reusable master elements against `asset-registry.json` bounds.
- Keep high-confidence layout elements within 3 px at the 1280x720 reference size.
- Document intentional deviations.
- Never replace official/institution marks with generated assets.

## Revision Checks

- Read `planning/locked-slides.json` before changing user-corrected slides.
- Preserve user-added images and manual CSS/HTML changes unless the request conflicts.
- Update `planning/revision-log.md` when using a task workspace.
- Screenshot affected slides after the revision.
