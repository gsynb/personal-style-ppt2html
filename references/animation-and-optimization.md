# Animation and Optimization

Use motion only when it improves a live presentation, video recording, or demo capture. For serious academic decks, motion must remain quiet and must never compete with figures, equations, citations, or institutional marks.

## Motion Presets

| Preset | Use for | Behavior |
|---|---|---|
| `none` | Manuscript-like notes, PDF export, conservative defenses | No entrance animation |
| `subtle` | Default HTML template and live academic talks | Short fade with slight upward movement |
| `recording` | Screen recording, voice-over videos, walkthroughs | Slightly slower and more legible entrance |
| `demo` | Product-style or public demo clips | Strongest allowed motion, still limited to opacity and transform |

Direct PPTX conversion supports:

```bash
python scripts/pptx_to_academic_html.py reference.pptx -o work/html-preview --profile work/style-profile.json --motion recording
```

## Motion Rules

- Animate only `opacity` and small `transform` changes. Avoid blur, spins, parallax, bounce, 3D movement, and decorative looping.
- Keep slide-layout identity elements stable. Logos, header bars, footer rules, and institution marks should not bounce or drift.
- Do not animate data marks, axes, equations, or labels in a way that changes perceived values or reading order.
- Always include `@media (prefers-reduced-motion: reduce)` that disables animation and transition.
- Disable animation inside `@media print`.
- Do not use `infinite` animation in academic slides unless the user explicitly asks for an animated simulation and the loop is scientifically meaningful.

## Output Optimization

- Use local assets so the deck can be opened offline and recorded without network stalls.
- Add `loading="lazy"` and `decoding="async"` to non-critical images. The converter does this for PPTX images.
- Prefer browser-native image formats: PNG, JPEG, SVG, WebP, or GIF. Convert EMF/WMF before final delivery when fidelity matters.
- Keep CSS and JS small and local. Avoid heavy frontend frameworks for a PPTX-derived preview unless the user asks for interactivity that justifies them.
- For video recording, use `--motion recording`, open the generated `index.html`, navigate with keyboard, and keep print/PDF export available as a non-animated fallback.

## Audit Expectations

Run:

```bash
python scripts/audit_html_layout.py work/html-preview/index.html
```

The audit should warn when animation CSS lacks a reduced-motion guard, when looping animation appears, or when an unknown `data-motion` preset is used.
