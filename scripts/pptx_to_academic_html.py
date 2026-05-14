#!/usr/bin/env python3
"""Convert a PPTX deck into a self-contained academic HTML preview."""

from __future__ import annotations

import argparse
import html
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx_common import (
    NS,
    PX_PER_INCH,
    bounds as _bounds,
    color_from_node as _color_from_node,
    gradient_from_node as _gradient_from_node,
    is_placeholder as _is_placeholder,
    read_xml as _read_xml,
    relationships as _relationships,
    rounded_box as _rounded_box,
    slide_layout_path as _slide_layout_path,
    slide_master_path as _slide_master_path,
    slide_number as _slide_number,
    target_hash as _target_hash,
    theme_colors as _theme_colors,
)

ALLOWED_MOTION = {"none", "subtle", "recording", "demo"}


def _paragraph_text(paragraph: ET.Element) -> str:
    parts: list[str] = []
    for child in list(paragraph):
        tag = child.tag.split("}")[-1]
        if tag in {"r", "fld"}:
            text = child.find("a:t", NS)
            if text is not None and text.text:
                parts.append(text.text)
        elif tag == "br":
            parts.append("\n")
    return "".join(parts).replace("\t", "    ")


def _text_and_style(shape: ET.Element, theme: dict[str, str], default_font: str) -> tuple[str, dict[str, Any]]:
    paragraphs: list[str] = []
    sizes: Counter[float] = Counter()
    fonts: Counter[str] = Counter()
    colors: Counter[str] = Counter()
    bold = False
    italic = False
    alignments: Counter[str] = Counter()

    for paragraph in shape.findall(".//a:p", NS):
        text = _paragraph_text(paragraph).strip()
        if text:
            paragraphs.append(text)
        ppr = paragraph.find("a:pPr", NS)
        if ppr is not None and ppr.get("algn"):
            alignments[ppr.get("algn", "")] += 1
        for rpr in paragraph.findall(".//a:rPr", NS) + paragraph.findall(".//a:defRPr", NS):
            if rpr.get("sz"):
                try:
                    sizes[round(int(rpr.get("sz", "0")) / 100, 1)] += 1
                except ValueError:
                    pass
            if rpr.get("b") == "1":
                bold = True
            if rpr.get("i") == "1":
                italic = True
            for font_tag in ("latin", "ea", "cs"):
                font_node = rpr.find(f"a:{font_tag}", NS)
                if font_node is not None and font_node.get("typeface"):
                    fonts[font_node.get("typeface", "")] += 1
            color = _color_from_node(rpr.find("a:solidFill", NS), theme)
            if color:
                colors[color] += 1

    style = {
        "font_size_pt": sizes.most_common(1)[0][0] if sizes else 18,
        "font_family": fonts.most_common(1)[0][0] if fonts else default_font,
        "color": colors.most_common(1)[0][0] if colors else "#262625",
        "font_weight": 650 if bold else 400,
        "font_style": "italic" if italic else "normal",
        "text_align": {"ctr": "center", "r": "right", "l": "left"}.get(
            alignments.most_common(1)[0][0], "left"
        )
        if alignments
        else "left",
    }
    return "\n".join(paragraphs), style


def _shape_visual(shape: ET.Element, theme: dict[str, str]) -> dict[str, str]:
    sppr = shape.find("p:spPr", NS)
    fill = "transparent"
    line = "transparent"
    if sppr is not None:
        fill = (
            _color_from_node(sppr.find("a:solidFill", NS), theme)
            or _gradient_from_node(sppr.find("a:gradFill", NS), theme)
            or "transparent"
        )
        line = _color_from_node(sppr.find("a:ln/a:solidFill", NS), theme, "transparent") or "transparent"
    return {"fill": fill, "line": line}


def _css_box(bounds: dict[str, float]) -> str:
    return (
        f"left:{bounds['left']:.2f}px;top:{bounds['top']:.2f}px;"
        f"width:{bounds['width']:.2f}px;height:{bounds['height']:.2f}px;"
    )


def _safe_text(text: str) -> str:
    return "<br>".join(html.escape(line) for line in text.splitlines())


def _extract_assets(zf: zipfile.ZipFile, output_assets: Path) -> None:
    output_assets.mkdir(parents=True, exist_ok=True)
    for name in zf.namelist():
        if name.startswith("ppt/media/"):
            (output_assets / Path(name).name).write_bytes(zf.read(name))


def _dedupe_signature(
    zf: zipfile.ZipFile,
    tag: str,
    bounds: dict[str, float],
    media: str = "",
    visual: dict[str, str] | None = None,
    text: str = "",
) -> str:
    if tag == "pic":
        identity = _target_hash(zf, media) or Path(media).name
    elif text:
        identity = re.sub(r"\s+", " ", text).strip().lower()[:160]
    else:
        visual = visual or {}
        identity = f"{visual.get('fill')}|{visual.get('line')}"
    return f"{tag}:{identity}:{_rounded_box(bounds)}"


def _normalize_motion(motion: str) -> str:
    if motion not in ALLOWED_MOTION:
        allowed = ", ".join(sorted(ALLOWED_MOTION))
        raise ValueError(f"Unsupported motion preset '{motion}'. Use one of: {allowed}.")
    return motion


def _is_master_placeholder_text(text: str) -> bool:
    normalized = re.sub(r"\s+", " ", text).strip().lower()
    return (
        text.startswith("单击此处")
        or text in {"‹#›"}
        or "click to edit" in normalized
        or "master title style" in normalized
        or "master text styles" in normalized
        or "edit master" in normalized
    )


def _convert_tree_elements(
    zf: zipfile.ZipFile,
    root: ET.Element,
    rels: dict[str, dict[str, str]],
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
    layer_class: str = "",
    skip_placeholders: bool = False,
    seen: set[str] | None = None,
) -> str:
    elements: list[str] = []
    seen = seen if seen is not None else set()

    for child in root.findall(".//p:cSld/p:spTree/*", NS):
        tag = child.tag.split("}")[-1]
        bounds = _bounds(child)
        if bounds is None:
            continue

        if tag == "pic":
            blip = child.find(".//a:blip", NS)
            rel_id = blip.get(f"{{{NS['r']}}}embed") if blip is not None else None
            rel = rels.get(rel_id or "")
            media = rel.get("target", "") if rel and rel.get("type", "").endswith("/image") else ""
            if not media:
                continue
            signature = _dedupe_signature(zf, tag, bounds, media=media)
            if signature in seen:
                continue
            seen.add(signature)
            filename = Path(media).name
            extension = Path(filename).suffix.lower()
            classes = f"pptx-img {layer_class}".strip()
            if extension in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}:
                elements.append(
                    f'<img class="{classes}" src="{asset_dir_name}/{html.escape(filename)}" '
                    f'style="{_css_box(bounds)}" alt="" loading="lazy" decoding="async">'
                )
            else:
                classes = f"pptx-emf {layer_class}".strip()
                elements.append(
                    f'<div class="{classes}" style="{_css_box(bounds)}">'
                    f"{html.escape(filename)}<br>EMF resource</div>"
                )
            continue

        if tag == "sp":
            if skip_placeholders and _is_placeholder(child):
                continue
            text, text_style = _text_and_style(child, theme, default_font)
            if skip_placeholders and text and _is_master_placeholder_text(text):
                continue
            visual = _shape_visual(child, theme)
            if text:
                signature = _dedupe_signature(zf, tag, bounds, visual=visual, text=text)
                if signature in seen:
                    continue
                seen.add(signature)
                font_size_px = text_style["font_size_pt"] * PX_PER_INCH / 72
                classes = f"pptx-text {layer_class}".strip()
                if bounds["top"] > 560:
                    classes += " citation"
                style = (
                    _css_box(bounds)
                    + f"font-family:'{html.escape(text_style['font_family'])}', var(--academic-font);"
                    + f"font-size:{font_size_px:.2f}px;"
                    + f"color:{text_style['color']};"
                    + f"font-weight:{text_style['font_weight']};"
                    + f"font-style:{text_style['font_style']};"
                    + f"text-align:{text_style['text_align']};"
                )
                elements.append(f'<div class="{classes}" style="{style}">{_safe_text(text)}</div>')
            else:
                if visual["fill"] != "transparent" or visual["line"] != "transparent":
                    signature = _dedupe_signature(zf, tag, bounds, visual=visual)
                    if signature in seen:
                        continue
                    seen.add(signature)
                    style = _css_box(bounds) + f"background:{visual['fill']};border:1px solid {visual['line']};"
                    classes = f"pptx-shape {layer_class}".strip()
                    elements.append(f'<div class="{classes}" style="{style}"></div>')

    return "\n".join(elements)


def _convert_layout(
    zf: zipfile.ZipFile,
    layout_path: str | None,
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
    seen: set[str] | None = None,
) -> str:
    if not layout_path:
        return ""
    root = _read_xml(zf, layout_path)
    if root is None:
        return ""
    return _convert_tree_elements(
        zf,
        root,
        _relationships(zf, layout_path),
        theme,
        default_font,
        asset_dir_name,
        layer_class="pptx-layout",
        skip_placeholders=True,
        seen=seen,
    )


def _convert_master(
    zf: zipfile.ZipFile,
    master_path: str | None,
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
    seen: set[str] | None = None,
) -> str:
    if not master_path:
        return ""
    root = _read_xml(zf, master_path)
    if root is None:
        return ""
    return _convert_tree_elements(
        zf,
        root,
        _relationships(zf, master_path),
        theme,
        default_font,
        asset_dir_name,
        layer_class="pptx-layout pptx-master",
        skip_placeholders=True,
        seen=seen,
    )


def _convert_slide(
    zf: zipfile.ZipFile,
    slide_path: str,
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
) -> str:
    number = _slide_number(slide_path)
    root = _read_xml(zf, slide_path)
    if root is None:
        return ""
    layout_path = _slide_layout_path(zf, slide_path)
    seen: set[str] = set()
    master_elements = _convert_master(
        zf,
        _slide_master_path(zf, layout_path),
        theme,
        default_font,
        asset_dir_name,
        seen=seen,
    )
    layout_elements = _convert_layout(zf, layout_path, theme, default_font, asset_dir_name, seen=seen)
    slide_elements = _convert_tree_elements(
        zf,
        root,
        _relationships(zf, slide_path),
        theme,
        default_font,
        asset_dir_name,
        seen=seen,
    )
    combined = "\n".join(element for element in [master_elements, layout_elements, slide_elements] if element)

    return (
        f'<section class="academic-slide pptx-slide" data-slide="{number}">\n'
        f'  <div class="slide-number">{number:02d}</div>\n'
        + "\n".join(f"  {element}" for element in combined.splitlines())
        + "\n</section>"
    )


def convert_pptx_to_html(
    pptx: str | Path,
    output_dir: str | Path,
    profile: str | Path | None = None,
    motion: str = "none",
) -> Path:
    pptx_path = Path(pptx).expanduser().resolve()
    out = Path(output_dir).expanduser().resolve()
    asset_dir = out / "assets"
    out.mkdir(parents=True, exist_ok=True)
    motion = _normalize_motion(motion)

    profile_data: dict[str, Any] = {}
    if profile:
        profile_path = Path(profile).expanduser().resolve()
        if profile_path.exists():
            profile_data = json.loads(profile_path.read_text(encoding="utf-8"))
    default_font = profile_data.get("typography", {}).get("dominant_font", "微软雅黑")
    colors = profile_data.get("colors", {})

    with zipfile.ZipFile(pptx_path) as zf:
        _extract_assets(zf, asset_dir)
        theme = _theme_colors(zf)
        slide_paths = sorted(
            [name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=_slide_number,
        )
        slides = [_convert_slide(zf, slide_path, theme, default_font, "assets") for slide_path in slide_paths]

    bg = colors.get("background", "#FFFFFF")
    text = colors.get("text", "#262625")
    accent = colors.get("accent", theme.get("accent1", "#4472C4"))
    muted = colors.get("muted", "#666666")
    style_css = f"""\
:root {{
  --academic-bg: {bg};
  --academic-text: {text};
  --academic-accent: {accent};
  --academic-muted: {muted};
  --academic-font: "{default_font}", "Microsoft YaHei", "Noto Sans SC", "PingFang SC", Arial, sans-serif;
}}
* {{ box-sizing: border-box; }}
[hidden] {{ display: none !important; }}
body {{
  margin: 0;
  background: #ececec;
  color: var(--academic-text);
  font-family: var(--academic-font);
}}
.deck-shell {{
  display: grid;
  min-height: 100vh;
  place-items: center;
  padding: 28px;
}}
.academic-slide {{
  position: relative;
  width: min(1280px, calc(100vw - 56px));
  aspect-ratio: 16 / 9;
  background: var(--academic-bg);
  overflow: hidden;
  box-shadow: 0 12px 34px rgba(0, 0, 0, 0.14);
}}
.pptx-slide {{
  transform-origin: top left;
}}
.pptx-text,
.pptx-img,
.pptx-shape,
.pptx-emf {{
  position: absolute;
}}
.pptx-text {{
  white-space: pre-wrap;
  line-height: 1.15;
  overflow: hidden;
}}
.pptx-text.citation {{
  color: var(--academic-muted) !important;
  line-height: 1.2;
}}
.pptx-img {{
  object-fit: contain;
}}
.pptx-emf {{
  display: grid;
  place-items: center;
  border: 1px dashed rgba(38, 38, 37, 0.28);
  color: var(--academic-muted);
  font-size: 12px;
  text-align: center;
}}
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout),
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout),
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout) {{
  will-change: opacity, transform;
}}
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="subtle"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout) {{
  animation: pptx-enter-subtle 260ms ease-out both;
}}
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="recording"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout) {{
  animation: pptx-enter-recording 420ms ease-out both;
}}
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-text:not(.pptx-layout),
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-img:not(.pptx-layout),
.academic-deck[data-motion="demo"] .academic-slide:not([hidden]) .pptx-shape:not(.pptx-layout) {{
  animation: pptx-enter-demo 520ms cubic-bezier(0.2, 0.8, 0.2, 1) both;
}}
@keyframes pptx-enter-subtle {{
  from {{ opacity: 0; transform: translateY(4px); }}
  to {{ opacity: 1; transform: translateY(0); }}
}}
@keyframes pptx-enter-recording {{
  from {{ opacity: 0; transform: translateY(6px); }}
  to {{ opacity: 1; transform: translateY(0); }}
}}
@keyframes pptx-enter-demo {{
  from {{ opacity: 0; transform: translateY(10px) scale(0.995); }}
  to {{ opacity: 1; transform: translateY(0) scale(1); }}
}}
.slide-number {{
  position: absolute;
  right: 18px;
  bottom: 12px;
  color: rgba(38, 38, 37, 0.28);
  font-size: 12px;
  z-index: 20;
}}
.controls {{
  position: fixed;
  left: 50%;
  bottom: 18px;
  z-index: 50;
  display: flex;
  align-items: center;
  gap: 10px;
  transform: translateX(-50%);
  border: 1px solid rgba(38, 38, 37, 0.14);
  background: rgba(255, 255, 255, 0.92);
  padding: 8px 12px;
  font-size: 14px;
}}
.controls button {{
  border: 1px solid rgba(38, 38, 37, 0.18);
  background: #fff;
  color: var(--academic-text);
  padding: 4px 10px;
  font: inherit;
}}
@media (prefers-reduced-motion: reduce) {{
  *,
  *::before,
  *::after {{
    animation: none !important;
    transition: none !important;
    scroll-behavior: auto !important;
  }}
  .academic-deck .academic-slide * {{
    opacity: 1 !important;
    transform: none !important;
  }}
}}
@media print {{
  body {{ background: #fff; }}
  .deck-shell {{ display: block; padding: 0; }}
  .academic-slide {{
    width: 100vw;
    height: 100vh;
    page-break-after: always;
    break-after: page;
    box-shadow: none;
  }}
  .academic-slide * {{
    animation: none !important;
    transition: none !important;
  }}
  .controls {{ display: none; }}
}}
"""
    js = """\
const deck = document.querySelector('.academic-deck');
const slides = Array.from(document.querySelectorAll('.academic-slide'));
let current = 0;
if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
  deck?.setAttribute('data-motion', 'none');
}
function show(index) {
  current = Math.max(0, Math.min(slides.length - 1, index));
  slides.forEach((slide, i) => slide.hidden = i !== current);
  document.querySelector('#counter').textContent = `${current + 1} / ${slides.length}`;
  history.replaceState(null, '', `#${current + 1}`);
}
document.querySelector('#prev').addEventListener('click', () => show(current - 1));
document.querySelector('#next').addEventListener('click', () => show(current + 1));
document.addEventListener('keydown', (event) => {
  if (['ArrowRight', 'PageDown', ' '].includes(event.key)) { event.preventDefault(); show(current + 1); }
  if (['ArrowLeft', 'PageUp'].includes(event.key)) { event.preventDefault(); show(current - 1); }
  if (event.key === 'Home') { event.preventDefault(); show(0); }
  if (event.key === 'End') { event.preventDefault(); show(slides.length - 1); }
});
window.addEventListener('beforeprint', () => slides.forEach(slide => slide.hidden = false));
window.addEventListener('afterprint', () => show(current));
function showFromHash() {
  const fromHash = Number.parseInt(location.hash.replace('#', ''), 10);
  show(Number.isFinite(fromHash) ? fromHash - 1 : 0);
}
window.addEventListener('hashchange', showFromHash);
showFromHash();
"""
    title = html.escape(pptx_path.stem)
    index = f"""<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{title} - HTML Preview</title>
    <link rel="stylesheet" href="./style.css">
  </head>
  <body>
    <main class="academic-deck deck-shell" data-motion="{html.escape(motion)}" aria-label="{title}">
      {''.join(slides)}
    </main>
    <nav class="controls" aria-label="slide controls">
      <button id="prev" type="button">Prev</button>
      <span id="counter"></span>
      <button id="next" type="button">Next</button>
    </nav>
    <script src="./deck.js"></script>
  </body>
</html>
"""
    (out / "style.css").write_text(style_css, encoding="utf-8")
    (out / "deck.js").write_text(js, encoding="utf-8")
    (out / "index.html").write_text(index, encoding="utf-8")
    return out / "index.html"


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert a PPTX into an academic HTML preview.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("-o", "--output-dir", type=Path, required=True)
    parser.add_argument("--profile", type=Path)
    parser.add_argument(
        "--motion",
        choices=sorted(ALLOWED_MOTION),
        default="none",
        help="Motion preset for browser presentation or recording output.",
    )
    args = parser.parse_args()
    index = convert_pptx_to_html(args.pptx, args.output_dir, args.profile, motion=args.motion)
    print(index)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
