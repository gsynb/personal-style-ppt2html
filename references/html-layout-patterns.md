# HTML Layout Patterns

Choose patterns from the extracted PPTX profile and the user's task.

## Common Academic Patterns

| Pattern | Use For | Structure |
| --- | --- | --- |
| `title-slide` | opening, defense title, group meeting title | unit lockup, large title, speaker/advisor/date footer |
| `outline-slide` | agenda | centered title, numbered section list |
| `figure-explanation-slide` | paper reading, background, model comparison | figure left, claim/bullets right, citation footer |
| `method-slide` | architecture, algorithm, workflow | title, editable flow diagram, method notes |
| `results-slide` | experiments and ablations | native chart/table/figure comparison, exact units |
| `summary-slide` | conclusion and outlook | 2-3 concise takeaways |

## Implementation Rules

- Build the actual presentation or note as the first screen; do not add a landing page.
- Use stable dimensions with `aspect-ratio`, grid tracks, and fixed slide padding.
- Keep text within containers across desktop and mobile; reduce content before reducing credibility.
- Make print/PDF export a first-class path with `@media print`.
- Use generated `theme.css` tokens from `style-profile.json` before adding new colors.

## Recording Mode

For research notes rather than slides, keep the same tokens but switch to:

- a narrow reading width,
- section-level citations,
- figure blocks with captions,
- collapsible appendices only when the user asks for interactive notes.
