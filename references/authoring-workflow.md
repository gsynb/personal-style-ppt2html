# Authoring Workflow

Use this reference when turning notes, outlines, papers, or a topic list into a new HTML deck. It complements, but does not replace, master-faithful PPTX style extraction.

## 1. Understand The Task

Identify:

- audience: lab, committee, course, seminar, public talk, or private research note
- format: slides, notes, recording, or reusable template
- language: Chinese, English, or bilingual
- expected slide count and density
- required reference PPTX, logos, figures, citations, equations, and data
- whether later edits must preserve user-modified HTML/CSS

## 2. Build A Claim Spine

Write one claim per generated slide before laying out content. A claim says what the audience should believe after that slide.

Examples:

- weak: `背景`
- better: `DFT 精确但慢，经典力场快但跨化学环境泛化弱`
- weak: `训练数据`
- better: `MLIP 的适用域主要由 DFT/AIMD/主动学习采样覆盖决定`

Keep claim titles as conclusions unless the user's reference PPTX consistently uses short topic labels.

## 3. Choose Proof Objects

Map each claim to the simplest proof object:

- comparison table for method tradeoffs
- matrix for risk, validation, or design choices
- mechanism flow for model or workflow logic
- timeline for model evolution
- equation band for formal definitions
- figure panel for paper/experiment visuals
- chart for measured or tabular data
- quote/source note for textual evidence

Main content must remain inside the safe area derived from the extracted master layer.

## 4. Rebuild As Presentation Language

- Shorten long paragraphs.
- Turn lists into grouped arguments.
- Turn methods into flows or diagrams.
- Turn dense tables into matrices or smaller evidence panels.
- Move details to speaker notes, appendix, or citations when a slide becomes crowded.

## 5. Preserve Style Evidence

Use `style-profile.json` and `asset-registry.json` as hard evidence for recurring layout geometry and as soft evidence for content rhythm. Do not invent unrelated card systems when the reference deck has a clear grammar.
