#!/usr/bin/env python3
"""Mine reusable visual elements from a PPTX package."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import zipfile
from collections import defaultdict
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


def _emu_to_px(value: str | int | None) -> float:
    try:
        return int(value or 0) / EMU_PER_INCH * PX_PER_INCH
    except ValueError:
        return 0.0


def _hex(value: str | None) -> str | None:
    if not value:
        return None
    value = value.strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", value):
        return f"#{value}"
    return None


def _relationship_path(part_path: str) -> str:
    part = PurePosixPath(part_path)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def _resolve_target(part_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(str(PurePosixPath(part_path).parent), target))


def _relationships(zf: zipfile.ZipFile, part_path: str) -> dict[str, dict[str, str]]:
    root = _read_xml(zf, _relationship_path(part_path))
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in list(root):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rel_id and target:
            rels[rel_id] = {"type": rel_type, "target": _resolve_target(part_path, target)}
    return rels


def _canvas(zf: zipfile.ZipFile) -> dict[str, float | str]:
    root = _read_xml(zf, "ppt/presentation.xml")
    width_emu = 12192000
    height_emu = 6858000
    if root is not None:
        size = root.find("p:sldSz", NS)
        if size is not None:
            width_emu = int(size.get("cx", str(width_emu)))
            height_emu = int(size.get("cy", str(height_emu)))
    width_px = _emu_to_px(width_emu)
    height_px = _emu_to_px(height_emu)
    return {
        "width_emu": width_emu,
        "height_emu": height_emu,
        "width_px": round(width_px, 2),
        "height_px": round(height_px, 2),
        "aspect_ratio": round(width_px / height_px, 4) if height_px else "unknown",
    }


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


def _normalized_bounds(bounds: dict[str, float], canvas: dict[str, float | str]) -> dict[str, float]:
    width = float(canvas.get("width_px") or 1)
    height = float(canvas.get("height_px") or 1)
    return {
        "left": round(bounds["left"] / width, 4),
        "top": round(bounds["top"] / height, 4),
        "width": round(bounds["width"] / width, 4),
        "height": round(bounds["height"] / height, 4),
    }


def _rounded_box(bounds: dict[str, float]) -> tuple[int, int, int, int]:
    return (
        round(bounds["left"] / 4),
        round(bounds["top"] / 4),
        round(bounds["width"] / 4),
        round(bounds["height"] / 4),
    )


def _color_from_node(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    srgb = node.find(".//a:srgbClr", NS)
    scheme = node.find(".//a:schemeClr", NS)
    if srgb is not None:
        return _hex(srgb.get("val"))
    if scheme is not None and scheme.get("val"):
        return f"scheme:{scheme.get('val')}"
    return None


def _gradient_signature(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    colors: list[str] = []
    for gs in node.findall(".//a:gs", NS):
        color = _color_from_node(gs)
        if color:
            colors.append(color)
    return "gradient:" + "|".join(colors) if colors else None


def _shape_visual(shape: ET.Element) -> dict[str, str]:
    sppr = shape.find("p:spPr", NS)
    if sppr is None:
        return {"fill": "transparent", "line": "transparent"}
    fill = (
        _color_from_node(sppr.find("a:solidFill", NS))
        or _gradient_signature(sppr.find("a:gradFill", NS))
        or "transparent"
    )
    line = _color_from_node(sppr.find("a:ln/a:solidFill", NS)) or "transparent"
    return {"fill": fill, "line": line}


def _shape_text(shape: ET.Element) -> str:
    parts = [node.text.strip() for node in shape.findall(".//a:t", NS) if node.text and node.text.strip()]
    return " ".join(parts)


def _is_placeholder(shape: ET.Element) -> bool:
    return shape.find(".//p:ph", NS) is not None


def _target_hash(zf: zipfile.ZipFile, target: str) -> str | None:
    if target not in zf.namelist():
        return None
    return hashlib.sha1(zf.read(target)).hexdigest()


def _slide_layout_path(zf: zipfile.ZipFile, slide_path: str) -> str | None:
    for rel in _relationships(zf, slide_path).values():
        if rel.get("type", "").endswith("/slideLayout"):
            return rel.get("target")
    return None


def _slide_master_path(zf: zipfile.ZipFile, layout_path: str | None) -> str | None:
    if not layout_path:
        return None
    for rel in _relationships(zf, layout_path).values():
        if rel.get("type", "").endswith("/slideMaster"):
            return rel.get("target")
    return None


def _is_citation(text: str) -> bool:
    return bool(re.search(r"\b(et al\.?|arXiv|doi|DOI|Nature|Science|Proceedings)\b", text) or re.search(r"\b(19|20)\d{2}\b", text))


def _infer_role(record: dict[str, Any], canvas: dict[str, float | str]) -> str:
    bounds = record["bounds_px"]
    width = float(canvas.get("width_px") or 1)
    height = float(canvas.get("height_px") or 1)
    left = bounds["left"]
    top = bounds["top"]
    box_width = bounds["width"]
    box_height = bounds["height"]
    width_ratio = box_width / width
    height_ratio = box_height / height

    if record["element_type"] == "image":
        if left < width * 0.16 and top < height * 0.22 and width_ratio < 0.22 and height_ratio < 0.28:
            return "institution_logo"
        return "reusable_image"
    if record["element_type"] == "text":
        if top > height * 0.72 and _is_citation(record.get("text", "")):
            return "citation_footer"
        return "repeated_text"
    if width_ratio > 0.6 and height_ratio <= 0.045:
        if top < height * 0.28:
            return "header_rule"
        if top + box_height > height * 0.72:
            return "footer_rule"
    if width_ratio > 0.7 and height_ratio > 0.045:
        if top < height * 0.3:
            return "header_band"
        if top > height * 0.64:
            return "footer_band"
    return "reusable_shape"


def _collect_tree_elements(
    zf: zipfile.ZipFile,
    part_path: str,
    source_level: str,
    slide_number: int,
    canvas: dict[str, float | str],
) -> list[dict[str, Any]]:
    root = _read_xml(zf, part_path)
    if root is None:
        return []
    rels = _relationships(zf, part_path)
    records: list[dict[str, Any]] = []

    for child in root.findall(".//p:cSld/p:spTree/*", NS):
        tag = child.tag.split("}")[-1]
        bounds = _bounds(child)
        if bounds is None:
            continue

        record: dict[str, Any] = {
            "source_level": source_level,
            "part_path": part_path,
            "slide_number": slide_number,
            "bounds_px": {key: round(value, 2) for key, value in bounds.items()},
            "bounds_norm": _normalized_bounds(bounds, canvas),
        }
        if tag == "pic":
            blip = child.find(".//a:blip", NS)
            rel_id = blip.get(f"{{{NS['r']}}}embed") if blip is not None else None
            rel = rels.get(rel_id or "")
            media = rel.get("target", "") if rel and rel.get("type", "").endswith("/image") else ""
            if not media:
                continue
            record.update(
                {
                    "element_type": "image",
                    "media": Path(media).name,
                    "media_path": media,
                    "media_sha1": _target_hash(zf, media),
                }
            )
        elif tag == "sp":
            if _is_placeholder(child):
                continue
            text = _shape_text(child)
            visual = _shape_visual(child)
            record.update(
                {
                    "element_type": "text" if text else "shape",
                    "text": text,
                    "fill": visual["fill"],
                    "line": visual["line"],
                }
            )
            if not text and visual["fill"] == "transparent" and visual["line"] == "transparent":
                continue
        else:
            continue
        record["role"] = _infer_role(record, canvas)
        records.append(record)
    return records


def _signature(record: dict[str, Any]) -> str:
    box = _rounded_box(record["bounds_px"])
    source_level = record["source_level"]
    element_type = record["element_type"]
    if element_type == "image":
        identity = record.get("media_sha1") or record.get("media") or ""
    elif element_type == "text":
        identity = re.sub(r"\s+", " ", record.get("text", "")).strip().lower()[:120]
    else:
        identity = f"{record.get('fill')}|{record.get('line')}"
    return f"{source_level}:{element_type}:{identity}:{box}"


def _reuse_level(source_level: str) -> str:
    if source_level == "slide_master":
        return "master"
    if source_level == "slide_layout":
        return "layout"
    return "manual_repeat"


def _confidence(group: list[dict[str, Any]]) -> float:
    source_level = group[0]["source_level"]
    role = group[0]["role"]
    if source_level in {"slide_master", "slide_layout"} and role in {"institution_logo", "header_rule", "footer_rule"}:
        return 0.96
    if source_level in {"slide_master", "slide_layout"}:
        return 0.9
    if len({item["slide_number"] for item in group}) >= 3:
        return 0.86
    return 0.78


def _layout_patterns(pptx_path: Path) -> list[dict[str, Any]]:
    try:
        from extract_pptx_style import analyze_pptx

        profile = analyze_pptx(pptx_path)
    except Exception:
        return []
    return profile.get("style_summary", {}).get("layout_counts", [])


def build_reusable_visual_registry(pptx: str | Path, min_occurrences: int = 2) -> dict[str, Any]:
    pptx_path = Path(pptx).expanduser().resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    with zipfile.ZipFile(pptx_path) as zf:
        canvas = _canvas(zf)
        slide_paths = sorted(
            [name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=_slide_number,
        )
        collected: list[dict[str, Any]] = []
        for slide_path in slide_paths:
            number = _slide_number(slide_path)
            layout_path = _slide_layout_path(zf, slide_path)
            master_path = _slide_master_path(zf, layout_path)
            if master_path:
                collected.extend(_collect_tree_elements(zf, master_path, "slide_master", number, canvas))
            if layout_path:
                collected.extend(_collect_tree_elements(zf, layout_path, "slide_layout", number, canvas))
            collected.extend(_collect_tree_elements(zf, slide_path, "slide", number, canvas))

    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in collected:
        groups[_signature(record)].append(record)

    reusable: list[dict[str, Any]] = []
    for group in groups.values():
        source_level = group[0]["source_level"]
        slide_numbers = sorted({item["slide_number"] for item in group})
        if source_level == "slide" and len(slide_numbers) < min_occurrences:
            continue
        first = group[0]
        reusable.append(
            {
                "id": f"reusable-{len(reusable) + 1:03d}",
                "role": first["role"],
                "element_type": first["element_type"],
                "source_level": source_level,
                "reuse_level": _reuse_level(source_level),
                "occurrence_count": len(slide_numbers),
                "slide_numbers": slide_numbers,
                "part_paths": sorted({item["part_path"] for item in group}),
                "bounds_px": first["bounds_px"],
                "bounds_norm": first["bounds_norm"],
                "fill": first.get("fill"),
                "line": first.get("line"),
                "media": first.get("media"),
                "media_sha1": first.get("media_sha1"),
                "text_preview": first.get("text", "")[:160],
                "confidence": _confidence(group),
            }
        )

    role_order = {
        "institution_logo": 0,
        "header_rule": 1,
        "header_band": 2,
        "footer_rule": 3,
        "footer_band": 4,
        "citation_footer": 5,
    }
    reusable.sort(key=lambda item: (role_order.get(item["role"], 20), -item["occurrence_count"], item["id"]))
    for index, item in enumerate(reusable, start=1):
        item["id"] = f"reusable-{index:03d}"

    return {
        "source": {
            "path": str(pptx_path),
            "slide_count": len({record["slide_number"] for record in collected}),
        },
        "canvas": canvas,
        "reusable_elements": reusable,
        "layout_patterns": _layout_patterns(pptx_path),
        "notes": [
            "Registry is based on PPTX structure and repeated geometry; use rendered screenshots or vision review for semantic confirmation.",
            "Do not regenerate institution logos or factual paper figures; reuse extracted source assets.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Mine reusable visual elements from a PPTX file.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("-o", "--output", type=Path, help="Write the reusable visual registry JSON.")
    parser.add_argument("--min-occurrences", type=int, default=2, help="Minimum slide count for slide-local repeats.")
    args = parser.parse_args()

    registry = build_reusable_visual_registry(args.pptx, min_occurrences=args.min_occurrences)
    text = json.dumps(registry, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
