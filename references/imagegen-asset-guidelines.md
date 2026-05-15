# Imagegen Asset Guidelines

Use image generation only for reusable non-factual visual assets that support a serious academic style.

## Good Uses

- restrained cover backgrounds,
- subtle research texture overlays,
- abstract graph, lattice, molecule, crystal, or network motifs,
- chapter divider imagery,
- neutral decorative backplates behind title slides,
- transparent-background sheets of blank arrows, chevrons, connectors, tabs, and module panels derived from extracted PPTX shape grammar.

## Bad Uses

- experimental charts,
- paper figures,
- model architectures with factual labels,
- logos, seals, or institution identity,
- any image that looks like measured data or a published result.

## Prompt Shape

Use prompts like:

```text
Create a restrained academic presentation background for a materials science group meeting: white field, subtle thin-line crystal lattice and graph network motif, sparse blue accent, no text, no logos, no data chart, high-resolution clean vector-like raster style.
```

Save generated assets in the task workspace under `assets/generated/` and reference them from HTML with explicit provenance in the asset manifest.

## Capability-Aware Workflow

Use `scripts/prepare_imagegen_briefs.py` after style extraction and reusable visual mining:

```bash
python scripts/prepare_imagegen_briefs.py work/style-profile.json \
  --registry work/asset-registry.json \
  -o work/imagegen-briefs.json
```

If the active agent runtime exposes image generation, execute only briefs where `safe_to_generate` is `true`. Save the final selected images in the HTML project under `assets/generated/`. If the runtime has no image generation tool or configured API, keep `imagegen-briefs.json` as a prompt handoff and continue the HTML build without generated imagery.

For Codex environments with a built-in image generation tool, generate first, then move or copy the selected output into the project workspace. Do not leave any HTML-referenced asset only in a global generated-images directory.

By default, `prepare_imagegen_briefs.py` excludes raw slide titles from prompts to avoid leaking private or factual text into generation. Use `--include-title-cues` only when the user explicitly wants topic-aware generated motifs and the title text is safe to use.

When wiring generated assets into HTML, prefer decorative treatment:

```html
<section class="academic-slide title-slide has-generated-bg">
  <img class="generated-bg" src="./assets/generated/style-cover-backdrop.png" alt="" aria-hidden="true" loading="lazy" decoding="async">
  ...
</section>
```

Use semantic `alt` text only when the generated image carries user-approved meaning. Keep official PPTX-derived logos and header rules as separate extracted assets.

## Default Asset Pack

The default generated asset pack is intentionally small:

- `style-cover-backdrop.png`: low-opacity title slide background.
- `style-section-divider.png`: sparse divider motif.
- `style-figure-backplate.png`: subtle texture behind real figures or HTML/SVG diagrams.
- `style-flow-elements.png`: optional transparent-background sheet of non-factual chevrons, connector arrows, corner markers, and accent rules when flow grammar is detected.
- `style-module-panels.png`: optional transparent-background sheet of blank module boxes, tabs, timeline cards, and callout panels when module grammar is detected.

Use these as optional supporting assets. They should not replace extracted logos, header rules, slide-master elements, or factual research figures.

## Provenance Manifest

For each generated image, record:

- source `style-profile.json` path,
- source `asset-registry.json` path if used,
- exact prompt,
- generation tool or API path,
- output filename,
- intended HTML usage,
- note that the image is non-factual decorative support.
