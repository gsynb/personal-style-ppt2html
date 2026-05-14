#!/usr/bin/env python3
"""Extract reusable academic HTML style signals from a PPTX file."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from pptx_common import is_citation as _shared_is_citation

EMU_PER_INCH = 914400

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def _hex(value: str | None) -> str | None:
    if not value:
        return None
    value = value.strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", value):
        return f"#{value}"
    return None


def _slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    if name not in zf.namelist():
        return None
    try:
        return ET.fromstring(zf.read(name))
    except ET.ParseError:
        return None


def _emu_to_in(value: str | int | None) -> float:
    try:
        return round(int(value or 0) / EMU_PER_INCH, 4)
    except ValueError:
        return 0.0


def _aspect_label(width_emu: int, height_emu: int) -> str:
    if not width_emu or not height_emu:
        return "unknown"
    ratio = width_emu / height_emu
    known = [(16 / 9, "16:9"), (4 / 3, "4:3"), (3 / 2, "3:2")]
    best_ratio, label = min(known, key=lambda item: abs(item[0] - ratio))
    return label if abs(best_ratio - ratio) < 0.04 else f"{ratio:.2f}:1"


def _extract_theme_colors(zf: zipfile.ZipFile) -> dict[str, str]:
    theme_names = sorted(n for n in zf.namelist() if n.startswith("ppt/theme/theme") and n.endswith(".xml"))
    if not theme_names:
        return {}
    root = _read_xml(zf, theme_names[0])
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


def _shape_bounds(shape: ET.Element) -> dict[str, float] | None:
    off = shape.find(".//a:xfrm/a:off", NS)
    ext = shape.find(".//a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    return {
        "x_in": _emu_to_in(off.get("x")),
        "y_in": _emu_to_in(off.get("y")),
        "w_in": _emu_to_in(ext.get("cx")),
        "h_in": _emu_to_in(ext.get("cy")),
    }


def _shape_text(shape: ET.Element) -> str:
    parts = [node.text.strip() for node in shape.findall(".//a:t", NS) if node.text and node.text.strip()]
    return " ".join(parts)


def _run_styles(shape: ET.Element) -> tuple[Counter[str], Counter[str], Counter[float]]:
    fonts: Counter[str] = Counter()
    colors: Counter[str] = Counter()
    sizes: Counter[float] = Counter()
    for rpr in shape.findall(".//a:rPr", NS):
        if rpr.get("sz"):
            try:
                sizes[round(int(rpr.get("sz", "0")) / 100, 1)] += 1
            except ValueError:
                pass
        for tag in ("latin", "ea", "cs"):
            node = rpr.find(f"a:{tag}", NS)
            if node is not None and node.get("typeface"):
                fonts[node.get("typeface", "")] += 1
        srgb = rpr.find(".//a:solidFill/a:srgbClr", NS)
        scheme = rpr.find(".//a:solidFill/a:schemeClr", NS)
        if srgb is not None and _hex(srgb.get("val")):
            colors[_hex(srgb.get("val")) or ""] += 1
        elif scheme is not None and scheme.get("val"):
            colors[f"scheme:{scheme.get('val')}"] += 1
    return fonts, colors, sizes


def _shape_fill(shape: ET.Element) -> str | None:
    sppr = shape.find("p:spPr", NS)
    if sppr is None:
        return None
    srgb = sppr.find(".//a:solidFill/a:srgbClr", NS)
    scheme = sppr.find(".//a:solidFill/a:schemeClr", NS)
    if srgb is not None:
        return _hex(srgb.get("val"))
    if scheme is not None and scheme.get("val"):
        return f"scheme:{scheme.get('val')}"
    return None


def _is_citation(text: str) -> bool:
    return _shared_is_citation(text)


def _choose_title(text_boxes: list[dict[str, Any]], canvas: dict[str, Any]) -> dict[str, Any] | None:
    if not text_boxes:
        return None
    height = max(canvas.get("height_in", 0), 1)

    def score(box: dict[str, Any]) -> float:
        size = box.get("max_font_size_pt") or 0
        y = box["bounds"]["y_in"]
        top_bonus = max(0, (height * 0.35 - y) / height) * 20
        length_penalty = max(0, len(box["text"]) - 90) / 20
        return size + top_bonus - length_penalty

    title = max(text_boxes, key=score)
    return {
        "text": title["text"],
        "bounds": title["bounds"],
        "max_font_size_pt": title.get("max_font_size_pt"),
        "alignment": _alignment(title["bounds"], canvas),
    }


def _alignment(bounds: dict[str, float], canvas: dict[str, Any]) -> str:
    width = canvas.get("width_in") or 1
    center = bounds["x_in"] + bounds["w_in"] / 2
    if abs(center - width / 2) < width * 0.08:
        return "center"
    if bounds["x_in"] < width * 0.18:
        return "left"
    if bounds["x_in"] > width * 0.55:
        return "right"
    return "offset"


def _classify_layout(slide: dict[str, Any], canvas: dict[str, Any]) -> str:
    title = slide.get("title_candidate") or {}
    citations = slide.get("citation_candidates") or []
    if title.get("alignment") == "center" and citations:
        return "centered-title-with-footer-citation"
    if slide["image_count"] >= 2 and slide["text_box_count"] >= 2:
        return "image-led-explanation"
    if slide["image_count"] >= 1 and slide["text_box_count"] <= 2:
        return "single-figure-with-caption"
    if slide["text_box_count"] >= 5:
        return "diagram-or-flow"
    return "text-and-proof"


def _extract_slide(root: ET.Element, number: int, canvas: dict[str, Any]) -> tuple[dict[str, Any], Counter[str], Counter[str], Counter[float], Counter[str]]:
    fonts: Counter[str] = Counter()
    text_colors: Counter[str] = Counter()
    sizes: Counter[float] = Counter()
    fills: Counter[str] = Counter()
    text_boxes: list[dict[str, Any]] = []
    citation_candidates: list[str] = []

    for shape in root.findall(".//p:sp", NS):
        text = _shape_text(shape)
        bounds = _shape_bounds(shape)
        shape_fonts, shape_colors, shape_sizes = _run_styles(shape)
        fonts.update(shape_fonts)
        text_colors.update(shape_colors)
        sizes.update(shape_sizes)
        fill = _shape_fill(shape)
        if fill:
            fills[fill] += 1
        if text and bounds:
            max_size = max(shape_sizes.keys(), default=None)
            box = {"text": text, "bounds": bounds, "max_font_size_pt": max_size}
            text_boxes.append(box)
            if bounds["y_in"] > (canvas.get("height_in") or 0) * 0.72 and _is_citation(text):
                citation_candidates.append(text)

    slide = {
        "number": number,
        "shape_count": len(root.findall(".//p:sp", NS)),
        "image_count": len(root.findall(".//p:pic", NS)),
        "table_count": len(root.findall(".//a:tbl", NS)),
        "text_box_count": len(text_boxes),
        "text_preview": " ".join(box["text"] for box in text_boxes)[:260],
        "title_candidate": _choose_title(text_boxes, canvas),
        "citation_candidates": citation_candidates[:3],
    }
    slide["layout_class"] = _classify_layout(slide, canvas)
    return slide, fonts, text_colors, sizes, fills


def _top(counter: Counter[Any], limit: int = 12) -> list[dict[str, Any]]:
    return [{"value": key, "count": count} for key, count in counter.most_common(limit)]


def _choose_plain_color(counter: Counter[str], fallback: str) -> str:
    for color, _count in counter.most_common():
        if color.startswith("#"):
            return color
    return fallback


def analyze_pptx(path: str | Path) -> dict[str, Any]:
    pptx_path = Path(path).expanduser().resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    with zipfile.ZipFile(pptx_path) as zf:
        names = zf.namelist()
        presentation = _read_xml(zf, "ppt/presentation.xml")
        width_emu = height_emu = 0
        if presentation is not None:
            size = presentation.find("p:sldSz", NS)
            if size is not None:
                width_emu = int(size.get("cx", "0"))
                height_emu = int(size.get("cy", "0"))

        canvas = {
            "width_emu": width_emu,
            "height_emu": height_emu,
            "width_in": _emu_to_in(width_emu),
            "height_in": _emu_to_in(height_emu),
            "aspect_ratio_label": _aspect_label(width_emu, height_emu),
        }
        theme_colors = _extract_theme_colors(zf)
        slide_paths = sorted(
            [name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=_slide_number,
        )

        fonts: Counter[str] = Counter()
        text_colors: Counter[str] = Counter()
        font_sizes: Counter[float] = Counter()
        shape_fills: Counter[str] = Counter()
        slides: list[dict[str, Any]] = []

        for slide_path in slide_paths:
            root = _read_xml(zf, slide_path)
            if root is None:
                continue
            slide, slide_fonts, slide_colors, slide_sizes, slide_fills = _extract_slide(
                root, _slide_number(slide_path), canvas
            )
            slides.append(slide)
            fonts.update(slide_fonts)
            text_colors.update(slide_colors)
            font_sizes.update(slide_sizes)
            shape_fills.update(slide_fills)

        dominant_font = fonts.most_common(1)[0][0] if fonts else "system-ui"
        background = "#FFFFFF"
        if "#FFFFFF" not in shape_fills and theme_colors.get("lt1"):
            background = theme_colors["lt1"]
        text = _choose_plain_color(text_colors, "#202124")
        accent = theme_colors.get("accent1") or _choose_plain_color(shape_fills, "#315FA8")
        muted = "#666666"

        layout_counter = Counter(slide["layout_class"] for slide in slides)
        citation_count = sum(1 for slide in slides if slide["citation_candidates"])
        bilingual_count = sum(1 for slide in slides if re.search(r"[\u4e00-\u9fff]", slide["text_preview"]))

        return {
            "source": {
                "path": str(pptx_path),
                "slide_count": len(slides),
                "media_count": len([name for name in names if name.startswith("ppt/media/")]),
            },
            "canvas": canvas,
            "theme": {"colors": theme_colors},
            "typography": {
                "dominant_font": dominant_font,
                "font_counts": _top(fonts),
                "font_size_counts_pt": _top(font_sizes),
            },
            "colors": {
                "background": background,
                "text": text,
                "accent": accent,
                "muted": muted,
                "text_color_counts": _top(text_colors),
                "shape_fill_counts": _top(shape_fills),
            },
            "slides": slides,
            "style_summary": {
                "layout_counts": _top(layout_counter),
                "citation_footer_slides": citation_count,
                "bilingual_slides": bilingual_count,
                "image_heavy_slides": sum(1 for slide in slides if slide["image_count"] >= 2),
            },
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract style signals from a PPTX reference deck.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("-o", "--output", type=Path, help="Write the style profile JSON to this path.")
    args = parser.parse_args()

    profile = analyze_pptx(args.pptx)
    text = json.dumps(profile, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
