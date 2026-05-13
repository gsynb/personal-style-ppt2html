# Reusable Visual Mining

Use this reference when the task is not only to reproduce a PPTX, but to extract a personal or institutional visual system that can be reused in future HTML decks.

## Pipeline

1. Parse PPTX structure for exact objects, coordinates, text, shape fills, media files, slide-layout links, and slide-master links.
2. Build a reusable visual registry:

   ```bash
   python scripts/mine_reusable_visuals.py reference.pptx -o work/asset-registry.json
   ```

3. Use the registry with `style-profile.json`:
   - `style-profile.json` captures typography, colors, layout classes, citation habits, and slide-level style signals.
   - `asset-registry.json` captures repeated or inherited visual assets such as institution logos, header rules, footer bands, repeated images, and manual repeated shapes.
4. When visual ambiguity remains, review rendered slide screenshots with a vision-capable model and the registry together.

## Registry Semantics

Each `reusable_elements[]` item includes:

- `role`: inferred semantic role, such as `institution_logo`, `header_rule`, `footer_band`, `citation_footer`, `reusable_image`, or `reusable_shape`.
- `element_type`: `image`, `shape`, or `text`.
- `source_level`: `slide_master`, `slide_layout`, or `slide`.
- `reuse_level`: `master`, `layout`, or `manual_repeat`.
- `occurrence_count` and `slide_numbers`: how often and where the element appears.
- `bounds_px` and `bounds_norm`: exact reusable placement.
- `media` / `media_sha1`: source asset identity for images.
- `confidence`: rule-based confidence, not a factual guarantee.

## Vision Review

Use vision review as a semantic layer, not as the source of truth. A good prompt asks the model to inspect screenshots and return JSON annotations that confirm or correct the registry:

```json
{
  "confirmed_elements": [
    {
      "registry_id": "reusable-001",
      "role": "institution_logo",
      "decision": "confirm",
      "reason": "Small repeated university mark in the upper-left corner."
    }
  ],
  "missed_elements": [
    {
      "role": "section_divider_background",
      "bbox_norm": [0.0, 0.0, 1.0, 1.0],
      "slide_numbers": [9],
      "reason": "Chapter divider style appears as a repeated visual motif."
    }
  ],
  "unsafe_to_reuse": [
    {
      "registry_id": "reusable-004",
      "reason": "Looks like a factual paper figure rather than a template asset."
    }
  ]
}
```

Do not ask the vision model to regenerate logos, paper figures, experimental plots, or brand marks. It should label, confirm, and flag; extracted source assets remain the reusable material.

## Current Limits

- The registry uses PPTX structure and repeated geometry. It does not prove that an element is semantically reusable.
- Repeated slide-local objects are detected by matching position, type, and asset/style identity; visually similar but differently encoded objects may need vision review.
- EMF/WMF and complex PowerPoint effects may need external conversion or screenshot-based review.
