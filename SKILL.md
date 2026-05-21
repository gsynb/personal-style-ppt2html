---
name: personal-style-ppt2html
description: Use when converting PPTX decks into serious academic HTML presentations, faithful PPTX-style HTML slide decks, video-ready slide previews, reusable personal-style templates, non-destructive revisions, or style-matched generated reusable assets, especially when prior decks contain slide-layout logos, institution marks, header rules, scientific figures, citations, equations, charts, or research-note material. Use strict master-layout fidelity when the user asks to mimic, reproduce, copy, or match the original PPT style. Do not use for raw pixel-perfect PowerPoint rendering or decorative marketing decks.
---

# Personal Style PPT2HTML

Create restrained, institution-aware academic HTML presentations or research notes by extracting the user's prior PPTX style, preserving reusable master/layout geometry, generating safe reusable visual assets when available, and building editable HTML around evidence-first layouts. Preserve the skill's core strength: evidence from the reference PPTX controls the master layer; planning, revisions, and generated assets must fit inside that extracted system.

## Fidelity Modes

Choose the fidelity mode before building:

- **master-faithful**: Default when the user asks to mimic, reproduce, copy, match, follow, or "use the style of" a reference PPTX. Layout-level elements from the PPTX are hard constraints.
- **academic-inspired**: Use only when the user explicitly asks for loose inspiration, a redesigned deck, or when the reference PPTX is too visually broken to preserve.
- **notes** or **recording**: May adapt density and motion, but must still preserve master elements if the request references a prior PPTX style.

In `master-faithful` mode, do not approximate recurring PPTX elements from memory. Use the extracted geometry.

## Workflow

1. **Collect sources.** Identify reference PPTX files, research notes, figures, citations, unit/lab assets, requested output mode (`slides`, `notes`, or `recording`), and fidelity mode (`master-faithful` unless the user clearly wants loose inspiration). For multi-turn or revision-heavy tasks, create a task workspace first:

   ```bash
   python scripts/create_workspace.py work/my-deck --profile group-meeting --language zh --slides 12
   ```

   Use `planning/revision-log.md`, `planning/locked-slides.json`, and `planning/image-preferences.md` when continuing a deck the user may have edited.
2. **Choose the path.**

   | User goal | Required command | Optional commands |
   |---|---|---|
   | Quick HTML preview | `python scripts/pptx_to_academic_html.py reference.pptx -o work/html-preview --motion none` | `--profile work/style-profile.json` |
   | Reusable style tokens | `extract_pptx_style.py` then `build_theme_css.py` | `make_asset_manifest.py` |
   | Faithful PPTX-style HTML deck | Full manual pipeline below, then build master CSS from `asset-registry.json` exact bounds | Contact sheet + overlap/diff audit before final |
   | Full personal style system | `python scripts/run_pipeline.py reference.pptx -o work/pipeline --motion recording` | Vision review of screenshots plus `asset-registry.json` |
   | Generated reusable assets | `python scripts/prepare_imagegen_briefs.py work/style-profile.json --registry work/asset-registry.json -o work/imagegen-briefs.json` | Execute safe briefs when image generation tools are available |

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

4. **Plan content as claims.** When creating new slide content from notes, outlines, or a topic list, read `references/authoring-workflow.md`. Write a claim spine and map each claim to a proof object before laying out slides. This planning step must not override extracted master geometry.
5. **Lock the design system.** Use `style-profile.json` for font, color, title alignment, citation placement, image density, common layout classes, and whether motion should be `none`, `subtle`, `recording`, or `demo`. Do not copy bad artifacts blindly; preserve the user's serious academic taste.
6. **Mine reusable visuals.** Use `asset-registry.json` to distinguish inherited assets (`slide_master`, `slide_layout`) from manually repeated slide-local objects (`manual_repeat`). Look for logos/rules/bands plus reusable flow grammar such as `flow_arrow`, `flow_connector`, `text_module`, and `module_box`. Treat the registry as evidence, not as an automatic permission to reuse factual figures.
7. **Build the master layer first.** In `master-faithful` mode, create a reusable HTML/CSS master layer from `asset-registry.json` before writing slide content:
   - Place `institution_logo`, `header_rule`, `footer_band`, recurring gray bands, title zones, and footer/citation zones using extracted `bounds_px` or `bounds_norm`.
   - Convert the PPTX canvas, usually 1280x720 for 16:9 exports, into CSS variables such as `--ppt-logo-left`, `--ppt-logo-top`, `--ppt-header-rule-top`, and `--ppt-content-top`.
   - Keep master element placement within 3 px of extracted 1280x720 coordinates unless an explicit responsive adaptation is needed.
   - Derive safe content boxes from the master layer. Main content must not overlap logos, header rules, title text, footer bands, or citation zones.
   - Prefer extracted media for official/institution marks. Never regenerate official logos.
8. **Generate reusable visual assets when possible.** If the current agent exposes a built-in image generation tool or configured API/CLI, execute this step; do not silently skip it.
   - Read `imagegen-briefs.json` and generate every brief marked `safe_to_generate`, unless the user explicitly says not to use image generation.
   - Generate only non-factual, reusable assets: subdued cover backgrounds, section-divider motifs, lattice/network textures, blank method backplates, decorative arrows, module boxes, connector sheets, and abstract scientific patterns.
   - Save generated outputs under the final HTML folder's `assets/generated/`, record a small provenance file listing prompt source, generation date, and intended use, and wire useful assets into the deck.
   - Use decorative generated images with `alt=""` and `aria-hidden="true"` unless the user assigns semantic meaning.
   - If image generation is unavailable or blocked, keep the briefs in the work directory and explicitly report that generated reusable assets were not produced.
   - Do not generate or guess institution logos, published paper figures, charts, plots, experimental data, equations, or factual method diagrams.
9. **Map new slides to reference layout families.** For each generated slide, choose the closest extracted reference layout class (`centered-title-with-footer-citation`, `image-led-explanation`, `diagram-or-flow`, `text-and-proof`, `single-figure-with-caption`, etc.). Adapt content inside that layout family instead of inventing unrelated card/grid systems.
10. **Keep figures reproducible when data is involved.** If a slide includes CSV/Excel data, statistical tables, manuscript figures, chart screenshots, or requested chart polish, read `references/figure-reproducibility.md`. Keep chart code/data/source notes under `figures/` when using external plotting, and keep factual diagrams editable in HTML/SVG whenever practical.
11. **Build from the template only after the master layer is fixed.** Copy `assets/academic-html-template/`, replace placeholder content, and load generated tokens after base styles:

   ```html
   <link rel="stylesheet" href="./theme.css">
   <link rel="stylesheet" href="./components.css">
   <link rel="stylesheet" href="./theme.generated.css">
   ```

12. **Use references only as needed.**
   - `references/authoring-workflow.md`: claim spine, proof-object mapping, and source-to-slide rewriting.
   - `references/academic-style-rules.md`: tone, evidence discipline, visual restraint.
   - `references/html-layout-patterns.md`: slide/note patterns and responsive rules.
   - `references/institution-brand-rules.md`: school, lab, group, and logo handling.
   - `references/imagegen-asset-guidelines.md`: when generated raster assets are safe.
   - `references/animation-and-optimization.md`: restrained motion presets, video-ready output, image optimization, and animation audit rules.
   - `references/reusable-visual-mining.md`: asset registry semantics, manual-repeat detection, and optional vision review schema.
   - `references/figure-reproducibility.md`: chart/data/figure workflow and reproducibility rules.
   - `references/revision-safety.md`: versioning, locked slides, user edits, and image habit tracking.
   - `references/quality-gates.md`: final package, render, master-fidelity, and revision checks.
13. **Validate.** Run:

   ```bash
   python scripts/audit_html_layout.py path/to/index.html
   ```

   For substantial frontend work, open the HTML in a browser and inspect desktop, mobile, and print/PDF behavior.

   In `master-faithful` mode, also render or inspect a generated contact sheet and run a visual sanity pass:
   - Confirm recurring master elements match extracted bounds within 3 px at the 1280x720 reference size.
   - Confirm no overlap among logo, title, header rule, footer band, citation, and main content.
   - Confirm generated reusable assets are subtle and do not obscure scientific content.
   - List any intentional deviations from the reference PPTX in the final response.

   Before claiming completion, committing, or pushing, read `references/quality-gates.md` and verify the relevant checks with fresh evidence.

## Design Rules

- Keep the first screen the actual presentation or research note, not a landing page.
- Prefer white or near-white academic surfaces, one accent color, thin rules, aligned figures, and clear citations.
- Use exact scientific labels, units, equations, figure captions, and bibliography details from user-provided sources.
- Use editable HTML, SVG, Mermaid, MathJax, or chart libraries for factual diagrams and data.
- Use claim titles and proof objects to improve clarity, but never move or resize extracted master elements to make a claim fit.
- In `master-faithful` mode, master elements extracted from `slide_master` or `slide_layout` are hard geometry constraints, not style suggestions.
- Use image generation, when available, for non-factual reusable assets such as restrained cover backgrounds, subtle lattice/network motifs, chapter dividers, blank module panels, connector/arrow sheets, or abstract backplates.
- Never generate or guess institution logos, published paper figures, experimental data, or official brand marks.
- When image generation is available, derive prompts from `imagegen-briefs.json`; do not ask the model to mimic a private deck screenshot directly unless the user explicitly provided it as a style reference and the asset is non-factual.
- Do not place content by eye when exact extracted bounds exist. Prefer CSS variables derived from `asset-registry.json`.

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
- `asset-registry.json`: reusable visual candidates, source levels, semantic roles, confidence, slide coverage, geometry, connector arrowheads, and exact bounds.
- `imagegen-briefs.json`: optional safe prompts for generated cover backgrounds, section-divider motifs, and backplate textures that match the extracted color, layout, and reusable-role evidence.

If the profile conflicts with the user's stated target, follow the user's target and document the intentional deviation.

## Master Fidelity Rules

Apply these rules whenever `master-faithful` mode is active:

- Reproduce `institution_logo`, `header_rule`, `footer_band`, recurring top/bottom bands, and other high-confidence layout elements using exact extracted coordinates.
- Create one shared `.ppt-master` layer or equivalent component used on every slide that needs the original master.
- Keep slide titles inside the extracted title zone. If a generated title is too long, reduce font size or wrap inside the title zone; do not move the logo or header rule.
- Reserve a content safe area below the header rule and above any footer/citation band. Main content must stay inside this safe area unless a reference layout proves otherwise.
- Use the reference contact sheet or PPTX-to-HTML preview as the visual baseline. Do not infer logo size, line thickness, band position, or footer placement from memory.
- If adapting for mobile, preserve order and identity of master elements, but allow stacked layout below the header.
- If a high-confidence reusable element is hard to reproduce as editable HTML, extract and reuse it as an image asset rather than approximating it poorly.

## Revision Safety Rules

When revising a generated deck, preserve user edits by default:

- Do not overwrite source PPTX files, prior delivered HTML folders, or user-edited HTML/CSS/JS unless explicitly requested.
- Use `scripts/create_workspace.py` for multi-turn work and place revisions under `output/versions/`.
- Read `planning/locked-slides.json` before modifying previously corrected slides.
- If the user removes presenter identity, changes an image, or adjusts a layout, treat that as intentional and record it in the planning files when a workspace exists.

## Image Generation Boundary

Use generated imagery as a style extension and reusable design system component, not as a replacement for extracted PPTX assets. Official logos, institution marks, paper figures, charts, plots, equations, and factual method diagrams must remain extracted, user-provided, or editable code-native assets.

Capability rule: if a built-in image generation tool is available, generate each approved brief from `imagegen-briefs.json`, then move or copy the selected output into the final HTML workspace. This is a required step unless the user opts out. If only an API/CLI fallback is available, use it only when the runtime is configured for it. If no image generation capability is available, deliver the briefs, keep the HTML functional without those assets, and state that generation was unavailable.

Generated reusable assets must be:

- non-factual and decorative or structural;
- visually subordinate to slide content;
- stored locally under the final output folder;
- accompanied by provenance;
- excluded from official marks, paper figures, data graphics, equations, and any asset whose factual correctness matters.

## Vision Review Boundary

Use vision-capable review only after structural mining when visual semantics are ambiguous. The model can confirm roles, identify missed motifs, and flag paper figures that should not become template assets. It must not invent or regenerate official marks, logos, paper figures, plots, or factual content.

## Output Expectations

Deliver a runnable HTML folder with local assets, stable CSS, and no hidden dependency on the reference PPTX. Keep generated scratch files in a work directory; keep final HTML, CSS, JS, generated reusable assets, provenance, and final extracted assets together.

For `master-faithful` outputs, final reporting must include:

- output folder path;
- revision path and base path if this was an update to an existing deck;
- whether image generation was used and where generated assets/provenance were saved;
- validation commands run;
- contact sheet/browser/visual review status;
- any known deviations from the reference PPTX master geometry.
