#!/usr/bin/env python3
"""Convert a PPTX deck into a self-contained academic HTML preview."""

from __future__ import annotations

import argparse
import html
import json
import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree as ET

EMU_PER_INCH = 914400
PX_PER_INCH = 96

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _emu_to_px(value: str | int | None) -> float:
    try:
        return int(value or 0) / EMU_PER_INCH * PX_PER_INCH
    except ValueError:
        return 0.0


def _read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    if name not in zf.namelist():
        return None
    return ET.fromstring(zf.read(name))


def _slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _hex(value: str | None) -> str | None:
    if not value:
        return None
    value = value.strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", value):
        return f"#{value}"
    return None


def _theme_colors(zf: zipfile.ZipFile) -> dict[str, str]:
    root = _read_xml(zf, "ppt/theme/theme1.xml")
    if root is None:
        return {}
    colors: dict[str, str] = {}
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return colors
    for child in list(scheme):
        key = child.tag.split("}")[-1]
        srgb = child.find(".//a:srgbClr", NS)
        sys_color = child.find(".//a:sysClr", NS)
        color = _hex(srgb.get("val") if srgb is not None else None)
        if not color and sys_color is not None:
            color = _hex(sys_color.get("lastClr") or sys_color.get("val"))
        if color:
            colors[key] = color
    return colors


def _color_from_node(node: ET.Element | None, theme: dict[str, str], fallback: str | None = None) -> str | None:
    if node is None:
        return fallback
    srgb = node.find(".//a:srgbClr", NS)
    scheme = node.find(".//a:schemeClr", NS)
    if srgb is not None:
        return _hex(srgb.get("val")) or fallback
    if scheme is not None:
        return theme.get(scheme.get("val", ""), fallback)
    return fallback


def _gradient_from_node(node: ET.Element | None, theme: dict[str, str]) -> str | None:
    if node is None:
        return None
    stops: list[str] = []
    for gs in node.findall(".//a:gs", NS):
        color = _color_from_node(gs, theme)
        if not color:
            continue
        try:
            percent = int(gs.get("pos") or "0") / 1000
        except ValueError:
            percent = 0
        stops.append(f"{color} {percent:.1f}%")
    if not stops:
        return None
    if len(stops) == 1:
        return stops[0].split()[0]
    return f"linear-gradient(90deg, {', '.join(stops)})"


def _bounds(node: ET.Element) -> dict[str, float] | None:
    off = node.find(".//a:xfrm/a:off", NS)
    ext = node.find(".//a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    return {
        "left": _emu_to_px(off.get("x")),
        "top": _emu_to_px(off.get("y")),
        "width": _emu_to_px(ext.get("cx")),
        "height": _emu_to_px(ext.get("cy")),
    }


def _relationship_path(part_path: str) -> str:
    part = PurePosixPath(part_path)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def _resolve_target(part_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(str(PurePosixPath(part_path).parent), target))


def _relationships(zf: zipfile.ZipFile, part_path: str) -> dict[str, dict[str, str]]:
    rel_path = _relationship_path(part_path)
    root = _read_xml(zf, rel_path)
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in list(root):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if not rel_id or not target:
            continue
        rels[rel_id] = {"type": rel_type, "target": _resolve_target(part_path, target)}
    return rels


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


def _is_placeholder(shape: ET.Element) -> bool:
    return shape.find(".//p:ph", NS) is not None


def _is_master_placeholder_text(text: str) -> bool:
    return text.startswith("单击此处") or text in {"‹#›"}


def _convert_tree_elements(
    zf: zipfile.ZipFile,
    root: ET.Element,
    rels: dict[str, dict[str, str]],
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
    layer_class: str = "",
    skip_placeholders: bool = False,
) -> str:
    elements: list[str] = []

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
            filename = Path(media).name
            extension = Path(filename).suffix.lower()
            classes = f"pptx-img {layer_class}".strip()
            if extension in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}:
                elements.append(
                    f'<img class="{classes}" src="{asset_dir_name}/{html.escape(filename)}" '
                    f'style="{_css_box(bounds)}" alt="">'
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
                    style = _css_box(bounds) + f"background:{visual['fill']};border:1px solid {visual['line']};"
                    classes = f"pptx-shape {layer_class}".strip()
                    elements.append(f'<div class="{classes}" style="{style}"></div>')

    return "\n".join(elements)


def _slide_layout_path(zf: zipfile.ZipFile, slide_path: str) -> str | None:
    for rel in _relationships(zf, slide_path).values():
        if rel.get("type", "").endswith("/slideLayout"):
            return rel.get("target")
    return None


def _convert_layout(
    zf: zipfile.ZipFile,
    layout_path: str | None,
    theme: dict[str, str],
    default_font: str,
    asset_dir_name: str,
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
    layout_elements = _convert_layout(zf, _slide_layout_path(zf, slide_path), theme, default_font, asset_dir_name)
    slide_elements = _convert_tree_elements(
        zf,
        root,
        _relationships(zf, slide_path),
        theme,
        default_font,
        asset_dir_name,
    )
    combined = "\n".join(element for element in [layout_elements, slide_elements] if element)

    return (
        f'<section class="academic-slide pptx-slide" data-slide="{number}">\n'
        f'  <div class="slide-number">{number:02d}</div>\n'
        + "\n".join(f"  {element}" for element in combined.splitlines())
        + "\n</section>"
    )


def convert_pptx_to_html(pptx: str | Path, output_dir: str | Path, profile: str | Path | None = None) -> Path:
    pptx_path = Path(pptx).expanduser().resolve()
    out = Path(output_dir).expanduser().resolve()
    asset_dir = out / "assets"
    out.mkdir(parents=True, exist_ok=True)

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
  .controls {{ display: none; }}
}}
"""
    js = """\
const slides = Array.from(document.querySelectorAll('.academic-slide'));
let current = 0;
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
    <main class="academic-deck deck-shell" aria-label="{title}">
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
    args = parser.parse_args()
    index = convert_pptx_to_html(args.pptx, args.output_dir, args.profile)
    print(index)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
