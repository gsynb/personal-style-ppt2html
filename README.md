# personal-style-ppt2html

English | [中文](README.zh-CN.md)

`personal-style-ppt2html` is an Agent Skill and Python utility toolkit for converting a user's existing PowerPoint style into serious, reusable academic HTML presentations.

It is designed for people who already have a personal or institutional presentation style in `.pptx` files and want to reuse that style in HTML without switching to flashy generic web templates.

It can be used by Codex, Claude Code, or any agent/runtime that can read a `SKILL.md` folder and run local Python scripts. Codex-specific UI metadata lives in `agents/openai.yaml`; the core workflow is in `SKILL.md`, `references/`, `scripts/`, and `assets/`.

## What It Can Do

- Convert a `.pptx` deck into a navigable HTML slide preview.
- Preserve reusable slide-layout elements such as institution logos, header bars, footer rules, and recurring visual marks.
- Preserve slide-master elements and detect manually repeated visual blocks that are not stored in layouts.
- Build an `asset-registry.json` of reusable visual candidates with semantic roles, source levels, slide coverage, exact bounds, geometry, connector arrowheads, and confidence.
- Detect reusable flow grammar from PPTX shapes, including chevrons, home plates, connector arrows, text modules, and module boxes.
- Build an `imagegen-briefs.json` prompt pack for optional style-matched generated backgrounds, divider motifs, backplate textures, and transparent vector-like element sheets when the active agent has image generation tools.
- Extract typography, theme colors, slide size, title positions, citation-footers, image density, and common layout patterns from prior decks.
- Copy embedded PPTX media into local HTML assets.
- Generate CSS design tokens from the extracted style profile.
- Add restrained motion presets for live talks, screen recording, and demo videos.
- Optimize generated HTML with local assets, lazy image loading, async image decoding, reduced-motion support, and print-safe CSS.
- Provide a restrained academic HTML template for talks, group meetings, paper reading reports, thesis defenses, and research notes.
- Audit generated HTML for basic structure, print support, motion safety, and slide container correctness.
- Scaffold task workspaces with planning files for claim spines, proof objects, non-destructive revisions, locked slides, image preferences, reproducible figures, and validation evidence.

## Why This Exists

Many HTML presentation templates are visually impressive but too decorative for academic talks. Academic presentations often need a different kind of polish:

- quiet visual hierarchy,
- exact figure and citation placement,
- institution identity without marketing-style decoration,
- paper figures and method diagrams as first-class content,
- personal consistency across group meetings, seminars, and defenses.

This project treats historical PPTX files as style references. It reads their Open XML structure and turns recurring design signals into HTML.

## How It Works

PowerPoint `.pptx` files are ZIP packages containing Office Open XML.

The converter reads:

- `ppt/slides/slideN.xml` for per-slide text, images, shapes, and coordinates,
- `ppt/slides/_rels/slideN.xml.rels` for slide relationships,
- `ppt/slideLayouts/slideLayoutN.xml` for reusable layout elements,
- `ppt/slideLayouts/_rels/slideLayoutN.xml.rels` for layout-level logos and media,
- `ppt/slideMasters/slideMasterN.xml` for master-level reusable marks and rules,
- `ppt/theme/theme1.xml` for theme colors,
- `ppt/media/*` for embedded images.

This matters because logos and style bars are often not stored directly on each slide. They may be stored in the slide layout or slide master, while some repeated visual blocks are manually copied across slides. The converter follows slide-to-layout-to-master relationships, and the visual mining script also clusters repeated slide-local elements.

The generated HTML can also carry a `data-motion` preset. Motion is intentionally limited to short opacity and transform entrances, with reduced-motion and print fallbacks, so the result works for both academic presenting and screen recording.

When an agent runtime exposes image generation, the workflow can also create non-factual supporting visuals from the extracted style fingerprint. The repository does not assume that every agent has this capability: it first writes `imagegen-briefs.json`, then the agent may execute those briefs and save the selected outputs under the final HTML project's `assets/generated/` directory.

## Repository Structure

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
├── assets/
│   └── academic-html-template/
├── references/
│   ├── academic-style-rules.md
│   ├── animation-and-optimization.md
│   ├── authoring-workflow.md
│   ├── figure-reproducibility.md
│   ├── html-layout-patterns.md
│   ├── imagegen-asset-guidelines.md
│   ├── institution-brand-rules.md
│   ├── quality-gates.md
│   ├── revision-safety.md
│   └── reusable-visual-mining.md
└── scripts/
    ├── audit_html_layout.py
    ├── build_theme_css.py
    ├── create_workspace.py
    ├── extract_pptx_style.py
    ├── make_asset_manifest.py
    ├── mine_reusable_visuals.py
    ├── pptx_common.py
    ├── pptx_to_academic_html.py
    ├── prepare_imagegen_briefs.py
    ├── run_pipeline.py
    └── test_academic_html_tools.py
```

## Quick Start

Use Python 3. No required third-party package is needed for the core XML parsing path.

For multi-turn or revision-heavy work, start with a workspace:

```bash
python scripts/create_workspace.py work/my-deck --profile group-meeting --language zh --slides 12
```

For the full workflow:

```bash
python scripts/run_pipeline.py your-deck.pptx -o work/pipeline --motion recording
```

For manual control:

```bash
python scripts/extract_pptx_style.py your-deck.pptx -o work/style-profile.json
python scripts/make_asset_manifest.py your-deck.pptx -o work/reference-assets
python scripts/mine_reusable_visuals.py your-deck.pptx -o work/asset-registry.json
python scripts/prepare_imagegen_briefs.py work/style-profile.json --registry work/asset-registry.json -o work/imagegen-briefs.json
python scripts/build_theme_css.py work/style-profile.json -o work/theme.generated.css
python scripts/pptx_to_academic_html.py your-deck.pptx -o work/html-preview --profile work/style-profile.json --motion recording
python scripts/audit_html_layout.py work/html-preview/index.html
```

Then open:

```text
work/html-preview/index.html
```

The generated deck supports keyboard navigation, reduced-motion preferences, and print/PDF export.

When using `theme.generated.css` with the reusable template, link it after `theme.css` and `components.css` so extracted tokens override the base defaults.

## Motion Presets

Use `--motion none` for the most conservative academic output. Use `--motion subtle` for live browser slides, `--motion recording` for screen-recorded walkthroughs, and `--motion demo` only when a more public demo style is appropriate.

The motion system avoids looping decoration, keeps slide-layout logos and rules stable, and disables animation for reduced-motion users and print output.

## Reusable Visual Registry

Run `mine_reusable_visuals.py` when you want to understand what can become part of the user's reusable style system. It identifies inherited objects from slide layouts and slide masters, plus slide-local objects that repeat with the same geometry and visual identity.

For ambiguous cases, use rendered slide screenshots and `asset-registry.json` together with a vision-capable model. Vision should label and flag candidates, not regenerate official logos, paper figures, plots, or factual content.

## Optional Generated Assets

If the active agent has image generation capability, use `imagegen-briefs.json` to generate only safe, non-factual assets:

- `style-cover-backdrop.png`,
- `style-section-divider.png`,
- `style-figure-backplate.png`.

Save selected outputs in the final HTML project under `assets/generated/` and record the prompt/tool provenance. These generated assets should extend the extracted style; they must not replace real PPTX-derived logos, institution marks, paper figures, charts, or factual diagrams.

By default, raw slide titles are excluded from image-generation prompts to avoid leaking private or factual text. Use `--include-title-cues` only when topic-aware motifs are explicitly desired and safe.

## Install as an Agent Skill

This repository follows the filesystem-based `SKILL.md` pattern used by Agent Skills. The scripts also work manually from the command line, so the repository is useful even outside an agent environment.

### Codex

Clone the repository into your Codex skills directory:

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.codex/skills/personal-style-ppt2html
```

Then invoke it as:

```text
Use $personal-style-ppt2html to convert my reference PPTX into an academic HTML deck.
```

### Claude Code

Clone the same repository into Claude Code's personal skills directory:

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.claude/skills/personal-style-ppt2html
```

For a project-local install that can be shared through a repository, place it under the project:

```bash
mkdir -p .claude/skills
git clone https://github.com/gsynb/personal-style-ppt2html.git .claude/skills/personal-style-ppt2html
```

Then ask Claude Code for the same task in natural language, or invoke the skill directly if your Claude Code setup exposes skill commands.

## Validation

Run:

```bash
python scripts/test_academic_html_tools.py
```

The tests cover:

- PPTX style extraction,
- CSS token generation,
- media asset extraction,
- HTML auditing,
- image-generation brief preparation,
- PPTX-to-HTML conversion with slide-layout logo and header-bar preservation,
- slide-master preservation and reusable visual registry mining,
- master/layout duplicate filtering and placeholder cleanup,
- motion preset output, optimized image attributes, and animation safety warnings.
- workspace scaffolding for planning, non-destructive revisions, locked slides, and validation outputs.

## Current Limitations

- `.emf` and `.wmf` media are not browser-native. Without a local converter such as LibreOffice, Inkscape, ImageMagick, or another vector conversion tool, these assets are represented as placeholders.
- Complex PowerPoint geometry is approximated. Simple filled rectangles, bars, images, and text boxes work best.
- The converter now follows slide-layout and slide-master inheritance for common pictures, text, and shapes, but complex PowerPoint effects are still approximated.
- The reusable visual registry uses PPTX structure and repeated geometry. LLM vision review is recommended when a repeated object may be a factual figure rather than a template asset.
- HTML output is intended as a faithful, inspectable preview and reusable style base, not a pixel-perfect PowerPoint renderer.

## Privacy Note

Source `.pptx` files and generated outputs are ignored by default. Do not commit private decks, unpublished figures, or sensitive research material unless you intentionally want them in the repository.
