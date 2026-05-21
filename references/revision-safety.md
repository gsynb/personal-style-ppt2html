# Revision Safety

Use this reference whenever updating, polishing, continuing, or modifying an existing generated HTML deck.

## Non-Destructive Rule

Never overwrite:

- user-provided PPTX files
- source notes, papers, datasets, or figures
- previously delivered HTML folders
- HTML/CSS/JS files the user may have manually edited

If a task continues from a delivered deck, treat the newest user-edited folder as the base, copy it to a new version path, then modify only that copy unless the user explicitly asks to edit in place.

## Version Pattern

Use the computer's local date and time plus a revision number:

```text
deck-name__YYYYMMDD-HHMMSS__rNN/
```

Store versioned deck folders under `output/versions/` when using `scripts/create_workspace.py`.

## Revision Workflow

1. Identify the base folder or file.
2. Inspect user edits before changing anything: slide count, titles, removed slides, image changes, CSS changes, and hidden/locked content.
3. Read `planning/locked-slides.json`, `planning/revision-log.md`, and `planning/image-preferences.md` when present.
4. Create the new version path before edits.
5. Apply only the requested change.
6. Render or screenshot affected slides and run HTML audit.
7. Append a short entry to `planning/revision-log.md`.

## Locked Slides

When the user says a slide should remain as-is, or when they manually correct a slide's visible content/notes, record it in `planning/locked-slides.json`:

```json
{
  "slides": {
    "1": {
      "content_locked": true,
      "notes_locked": true,
      "reason": "User removed presenter identity from title slide",
      "locked_at": "2026-05-21 15:30"
    }
  }
}
```

Do not change locked visible content or notes unless the user explicitly unlocks or edits that slide.

## Image Habit Tracking

When the user inserts, crops, replaces, dims, frames, or moves an image, record:

- slide number and role
- source type: PPTX-derived asset, user image, screenshot, chart, generated decorative asset
- placement and crop
- caption/source-note habit
- border, shadow, opacity, and image-to-text ratio

Use these notes before selecting future images or generated motifs.
