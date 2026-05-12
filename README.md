# personal-style-ppt2html

English | [中文](README.zh-CN.md)

`personal-style-ppt2html` is a Codex skill and utility toolkit for converting a user's existing PowerPoint style into serious, reusable academic HTML presentations.

It is designed for people who already have a personal or institutional presentation style in `.pptx` files and want to reuse that style in HTML without switching to flashy generic web templates.

## What It Can Do

- Convert a `.pptx` deck into a navigable HTML slide preview.
- Preserve reusable slide-layout elements such as institution logos, header bars, footer rules, and recurring visual marks.
- Extract typography, theme colors, slide size, title positions, citation-footers, image density, and common layout patterns from prior decks.
- Copy embedded PPTX media into local HTML assets.
- Generate CSS design tokens from the extracted style profile.
- Provide a restrained academic HTML template for talks, group meetings, paper reading reports, thesis defenses, and research notes.
- Audit generated HTML for basic structure, print support, and slide container correctness.

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
- `ppt/theme/theme1.xml` for theme colors,
- `ppt/media/*` for embedded images.

This matters because logos and style bars are often not stored directly on each slide. They are usually stored in the slide layout. The converter follows the slide-to-layout relationship and inserts those reusable elements into the HTML output.

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
│   ├── html-layout-patterns.md
│   ├── imagegen-asset-guidelines.md
│   └── institution-brand-rules.md
└── scripts/
    ├── audit_html_layout.py
    ├── build_theme_css.py
    ├── extract_pptx_style.py
    ├── make_asset_manifest.py
    ├── pptx_to_academic_html.py
    └── test_academic_html_tools.py
```

## Quick Start

Use Python 3. No required third-party package is needed for the core XML parsing path.

```bash
python scripts/extract_pptx_style.py your-deck.pptx -o work/style-profile.json
python scripts/make_asset_manifest.py your-deck.pptx -o work/reference-assets
python scripts/build_theme_css.py work/style-profile.json -o work/theme.generated.css
python scripts/pptx_to_academic_html.py your-deck.pptx -o work/html-preview --profile work/style-profile.json
python scripts/audit_html_layout.py work/html-preview/index.html
```

Then open:

```text
work/html-preview/index.html
```

The generated deck supports keyboard navigation and print/PDF export.

## Install as a Codex Skill

Clone the repository into your Codex skills directory:

```bash
git clone https://github.com/gsynb/personal-style-ppt2html.git ~/.codex/skills/personal-style-ppt2html
```

Then invoke it as:

```text
Use $personal-style-ppt2html to convert my reference PPTX into an academic HTML deck.
```

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
- PPTX-to-HTML conversion with slide-layout logo and header-bar preservation.

## Current Limitations

- `.emf` and `.wmf` media are not browser-native. Without a local converter such as LibreOffice, Inkscape, ImageMagick, or another vector conversion tool, these assets are represented as placeholders.
- Complex PowerPoint geometry is approximated. Simple filled rectangles, bars, images, and text boxes work best.
- The converter currently focuses on slide and slide-layout content. Full slide-master inheritance can be extended further if a deck stores reusable visual elements only in the master layer.
- HTML output is intended as a faithful, inspectable preview and reusable style base, not a pixel-perfect PowerPoint renderer.

## Privacy Note

Source `.pptx` files and generated outputs are ignored by default. Do not commit private decks, unpublished figures, or sensitive research material unless you intentionally want them in the repository.
