# Figure Reproducibility

Use this reference when a deck includes CSV, Excel, statistical tables, experimental data, chart screenshots, manuscript figures, or requests to improve scientific charts.

## Core Standard

Every figure or chart must support one slide claim. Prefer editable HTML/SVG for factual diagrams and reproducible code-generated figures for data-heavy charts.

## Folder Pattern

For each important generated chart, keep:

```text
figures/
  fig01_short_name/
    code/
    data/
    outputs/
    sources/
    README.md
```

Store:

- `code/`: script that regenerates the chart
- `data/`: raw or cleaned data used by the script
- `outputs/`: PNG for HTML/PPT use plus SVG/PDF when useful
- `sources/`: source notes distinguishing user-provided data, cleaned data, and illustrative values
- `README.md`: chart claim, command, output files, and source type

## Chart Choice

- trend: line chart or slope chart
- group comparison: bar, dot, or interval plot
- distribution: box, violin, histogram, or ridgeline only when readable
- relationship: scatter with direct annotation
- part-to-whole: stacked bar or small multiples; avoid pie charts unless categories are few
- mechanism: editable SVG/HTML flow, not a decorative bitmap
- literature map or taxonomy: matrix or grouped table

## Style Rules

- white or near-white background
- thin axes and light gridlines
- readable labels at slide size
- direct labels when possible
- color-blind-safe accents
- no 3D effects, glossy fills, heavy shadows, or decorative chart gradients

For formal academic decks, generated charts should look compatible with the extracted PPTX style tokens rather than introducing a separate visual brand.

## QA

Before delivery, verify:

- chart supports the slide claim
- data source is recorded
- outputs regenerate from code/data when externally generated
- chart text is readable in rendered screenshots
- aspect ratio is preserved when inserted into HTML
