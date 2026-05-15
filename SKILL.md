---
name: personal-style-ppt2html
description: Use when converting PPTX decks into serious academic HTML presentations, video-ready slide previews, reusable personal-style templates, or optional style-matched generated visual assets, especially when prior decks contain slide-layout logos, institution marks, header rules, scientific figures, citations, equations, or research-note material. Do not use for pixel-perfect PowerPoint rendering or decorative marketing decks.
---

# Personal Style PPT2HTML

Create restrained, institution-aware academic HTML presentations or research notes by extracting the user's prior PPTX style, turning it into CSS tokens, and building editable HTML around evidence-first layouts.

## Workflow

1. **Collect sources.** Identify reference PPTX files, research notes, figures, citations, unit/lab assets, and the requested output mode: `slides`, `notes`, or `recording`.
2. **Choose the path.**

   | User goal | Required command | Optional commands |
   |---|---|---|
   | Quick HTML preview | `python scripts/pptx_to_academic_html.py reference.pptx -o work/html-preview --motion none` | `--profile work/style-profile.json` |
   | Reusable style tokens | `extract_pptx_style.py` then `build_theme_css.py` | `make_asset_manifest.py` |
   | Full personal style system | `python scripts/run_pipeline.py reference.pptx -o work/pipeline --motion recording` | Vision review of screenshots plus `asset-registry.json` |
   | Generated reusable assets | `python scripts/prepare_imagegen_briefs.py work/style-profile.json --registry work/asset-registry.json -o work/imagegen-briefs.json` | Execute briefs only when image generation tools are available |

3. **Extract style.** For the full manual pipeline, run:

   ```bash
   python scripts/extract_pptx_style.py reference.pptx -o work/style-profile.json
   python scripts/make_asset_manifest.py reference.pptx -o work/reference-assets
   python scripts/mine_reusable_visuals.py reference.pptx -o work/asset-registry.json
   python scripts/prepare_imagegen_briefs.py work/style-profile.json --registry work/asset-registry.json -o work/imagegen-briefs.json
   python scripts/build_theme_css.py work/style-profile.json -o work/theme.generated.css
   ```

   For a direct PPTX-to-HTML preview that preserves reusable slide-layout elements such as logos, rules, and recurring header bars, run:

   ```bash
   python scripts/pptx_to_academic_html.py reference.pptx -o work/html-preview --profile work/style-profile.json --motion recording
   ```

   Use `--motion none` for the most conservative output.

4. **Lock the design system.** Use `style-profile.json` for font, color, title alignment, citation placement, image density, common layout classes, and whether motion should be `none`, `subtle`, `recording`, or `demo`. Do not copy bad artifacts blindly; preserve the user's serious academic taste.
5. **Mine reusable visuals.** Use `asset-registry.json` to distinguish inherited assets (`slide_master`, `slide_layout`) from manually repeated slide-local objects (`manual_repeat`). Treat the registry as evidence, not as an automatic permission to reuse factual figures.
6. **Generate optional style-matched assets.** If the current agent exposes an image generation tool or API, read `imagegen-briefs.json` and generate only the safe briefs marked `safe_to_generate`. Save project-bound outputs under the final HTML folder's `assets/generated/`, record provenance, and wire them as low-opacity cover backgrounds, section motifs, or figure backplates. Use decorative images with `alt=""` and `aria-hidden="true"` unless the user assigns semantic meaning. If image generation is unavailable, keep the briefs as prompts and continue without blocking the HTML conversion. Do not include raw slide-title cues in generation unless the user explicitly asks for topic-aware motifs and the text is safe.
7. **Build from the template.** Copy `assets/academic-html-template/`, replace placeholder content, and load generated tokens after base styles:

   ```html
   <link rel="stylesheet" href="./theme.css">
   <link rel="stylesheet" href="./components.css">
   <link rel="stylesheet" href="./theme.generated.css">
   ```

8. **Use references only as needed.**
   - `references/academic-style-rules.md`: tone, evidence discipline, visual restraint.
   - `references/html-layout-patterns.md`: slide/note patterns and responsive rules.
   - `references/institution-brand-rules.md`: school, lab, group, and logo handling.
   - `references/imagegen-asset-guidelines.md`: when generated raster assets are safe.
   - `references/animation-and-optimization.md`: restrained motion presets, video-ready output, image optimization, and animation audit rules.
   - `references/reusable-visual-mining.md`: asset registry semantics, manual-repeat detection, and optional vision review schema.
9. **Validate.** Run:

   ```bash
   python scripts/audit_html_layout.py path/to/index.html
   ```

   For substantial frontend work, open the HTML in a browser and inspect desktop, mobile, and print/PDF behavior.

## Design Rules

- Keep the first screen the actual presentation or research note, not a landing page.
- Prefer white or near-white academic surfaces, one accent color, thin rules, aligned figures, and clear citations.
- Use exact scientific labels, units, equations, figure captions, and bibliography details from user-provided sources.
- Use editable HTML, SVG, Mermaid, MathJax, or chart libraries for factual diagrams and data.
- Use image generation only for non-factual reusable assets such as restrained cover backgrounds, subtle lattice/network motifs, or chapter dividers.
- Never generate or guess institution logos, published paper figures, experimental data, or official brand marks.
- When image generation is available, derive prompts from `imagegen-briefs.json`; do not ask the model to mimic a private deck screenshot directly unless the user explicitly provided it as a style reference and the asset is non-factual.

## Motion and Optimization Rules

- Default to `none` for formal or PDF-first academic outputs; use `subtle` for live slides and `recording` for screen-recorded walkthroughs.
- Keep animation to short opacity and small transform entrances. Do not animate data, equations, official logos, header bars, or footer rules in a distracting way.
- Every animated output must include a `prefers-reduced-motion` guard and print rules that disable motion.
- Preserve local assets and use optimized image attributes such as `loading="lazy"` and `decoding="async"` for PPTX-derived images.

## PPTX Profile Signals

Treat the extracted profile as a starting point:

- `typography.dominant_font`: default font stack.
- `colors`: background, text, accent, and muted tokens.
- `slides[].title_candidate`: title position and alignment.
- `slides[].layout_class`: reusable rhythm such as `figure-explanation-slide`, `diagram-or-flow`, or `centered-title-with-footer-citation`.
- `style_summary`: evidence of image-heavy slides, bilingual slides, and citation-footer habits.
- `asset-registry.json`: reusable visual candidates, source levels, semantic roles, confidence, slide coverage, and exact bounds.
- `imagegen-briefs.json`: optional safe prompts for generated cover backgrounds, section-divider motifs, and backplate textures that match the extracted color, layout, and reusable-role evidence.

If the profile conflicts with the user's stated target, follow the user's target and document the intentional deviation.

## Image Generation Boundary

Use generated imagery as a style extension, not as a replacement for extracted PPTX assets. Official logos, institution marks, paper figures, charts, plots, equations, and factual method diagrams must remain extracted, user-provided, or editable code-native assets.

Capability rule: if a built-in image generation tool is available, generate each approved brief from `imagegen-briefs.json`, then move or copy the selected output into the final HTML workspace. If only an API/CLI fallback is available, use it only when the runtime is configured for it. If no image generation capability is available, deliver the briefs and keep the HTML functional without those assets.

## Vision Review Boundary

Use vision-capable review only after structural mining when visual semantics are ambiguous. The model can confirm roles, identify missed motifs, and flag paper figures that should not become template assets. It must not invent or regenerate official marks, logos, paper figures, plots, or factual content.

## Output Expectations

Deliver a runnable HTML folder with local assets, stable CSS, and no hidden dependency on the reference PPTX. Keep generated scratch files in a work directory; keep final HTML, CSS, JS, and final assets together.
