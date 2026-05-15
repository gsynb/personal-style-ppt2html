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
   - `asset-registry.json` captures repeated or inherited visual assets such as institution logos, header rules, footer bands, repeated images, manual repeated shapes, flow arrows, connectors, and reusable text modules.
4. When visual ambiguity remains, review rendered slide screenshots with a vision-capable model and the registry together.

## Registry Semantics

Each `reusable_elements[]` item includes:

- `role`: inferred semantic role, such as `institution_logo`, `header_rule`, `footer_band`, `citation_footer`, `flow_arrow`, `flow_connector`, `text_module`, `module_box`, `reusable_image`, or `reusable_shape`.
- `element_type`: `image`, `shape`, `connector`, or `text`.
- `source_level`: `slide_master`, `slide_layout`, or `slide`.
- `reuse_level`: `master`, `layout`, or `manual_repeat`.
- `occurrence_count` and `slide_numbers`: how often and where the element appears.
- `bounds_px` and `bounds_norm`: exact reusable placement.
- `media` / `media_sha1`: source asset identity for images.
- `geometry`, `line_width`, `arrow_head`, and `arrow_tail`: OOXML shape grammar for chevrons, home plates, round rectangles, and connector arrows.
- `confidence`: rule-based confidence, not a factual guarantee.

## Flow and Module Grammar

PowerPoint does not always store arrows as images. Common flow visuals are usually preset geometries:

- `chevron`, `homePlate`, `rightArrow`, `leftArrow`, and related block-arrow shapes become `flow_arrow`.
- `p:cxnSp` connector shapes, or connector-like line geometries with `a:headEnd` / `a:tailEnd`, become `flow_connector`.
- Filled or stroked `rect`, `roundRect`, `round1Rect`, and similar boxes with text become `text_module`; the same shapes without text become `module_box`.

Use these roles as a visual grammar when rebuilding HTML: CSS chevrons, SVG connectors, grid modules, and timeline arrows should follow the extracted bounds, colors, geometry, and recurrence pattern instead of being invented from scratch.

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
- Repeated slide-local objects are detected by matching position, type, geometry, and asset/style identity; visually similar but differently encoded objects may need vision review.
- Flow and module grammar roles are kept even when they occur on only one slide. They carry lower confidence, but are useful style evidence for rebuilding arrows, process diagrams, timelines, and module panels in HTML.
- Grouped PowerPoint objects can still need vision review when their child coordinates are transformed by complex group-level geometry.
- EMF/WMF and complex PowerPoint effects may need external conversion or screenshot-based review.
