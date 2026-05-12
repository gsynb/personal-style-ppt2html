# Imagegen Asset Guidelines

Use image generation only for reusable non-factual visual assets that support a serious academic style.

## Good Uses

- restrained cover backgrounds,
- subtle research texture overlays,
- abstract graph, lattice, molecule, crystal, or network motifs,
- chapter divider imagery,
- neutral decorative backplates behind title slides.

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
