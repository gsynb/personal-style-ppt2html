"""Shared helpers for reading PowerPoint Open XML packages."""

from __future__ import annotations

import hashlib
import posixpath
import re
import zipfile
from pathlib import PurePosixPath
from typing import Any
from xml.etree import ElementTree as ET

EMU_PER_INCH = 914400
PX_PER_INCH = 96

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def emu_to_px(value: str | int | None) -> float:
    try:
        return int(value or 0) / EMU_PER_INCH * PX_PER_INCH
    except ValueError:
        return 0.0


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    if name not in zf.namelist():
        return None
    try:
        return ET.fromstring(zf.read(name))
    except ET.ParseError:
        return None


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def hex_color(value: str | None) -> str | None:
    if not value:
        return None
    value = value.strip().lstrip("#").upper()
    if re.fullmatch(r"[0-9A-F]{6}", value):
        return f"#{value}"
    return None


def relationship_path(part_path: str) -> str:
    part = PurePosixPath(part_path)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def resolve_target(part_path: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(str(PurePosixPath(part_path).parent), target))


def relationships(zf: zipfile.ZipFile, part_path: str) -> dict[str, dict[str, str]]:
    root = read_xml(zf, relationship_path(part_path))
    if root is None:
        return {}
    rels: dict[str, dict[str, str]] = {}
    for rel in list(root):
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rel_id and target:
            rels[rel_id] = {"type": rel_type, "target": resolve_target(part_path, target)}
    return rels


def bounds(node: ET.Element) -> dict[str, float] | None:
    off = node.find(".//a:xfrm/a:off", NS)
    ext = node.find(".//a:xfrm/a:ext", NS)
    if off is None or ext is None:
        return None
    return {
        "left": emu_to_px(off.get("x")),
        "top": emu_to_px(off.get("y")),
        "width": emu_to_px(ext.get("cx")),
        "height": emu_to_px(ext.get("cy")),
    }


def normalized_bounds(box: dict[str, float], canvas: dict[str, float | str]) -> dict[str, float]:
    width = float(canvas.get("width_px") or 1)
    height = float(canvas.get("height_px") or 1)
    return {
        "left": round(box["left"] / width, 4),
        "top": round(box["top"] / height, 4),
        "width": round(box["width"] / width, 4),
        "height": round(box["height"] / height, 4),
    }


def rounded_box(box: dict[str, float], tolerance_px: int = 4) -> tuple[int, int, int, int]:
    return (
        round(box["left"] / tolerance_px),
        round(box["top"] / tolerance_px),
        round(box["width"] / tolerance_px),
        round(box["height"] / tolerance_px),
    )


def color_from_node(
    node: ET.Element | None,
    theme: dict[str, str] | None = None,
    fallback: str | None = None,
) -> str | None:
    if node is None:
        return fallback
    srgb = node.find(".//a:srgbClr", NS)
    scheme = node.find(".//a:schemeClr", NS)
    if srgb is not None:
        return hex_color(srgb.get("val")) or fallback
    if scheme is not None and scheme.get("val"):
        if theme is not None:
            return theme.get(scheme.get("val", ""), fallback)
        return f"scheme:{scheme.get('val')}"
    return fallback


def gradient_from_node(node: ET.Element | None, theme: dict[str, str]) -> str | None:
    if node is None:
        return None
    stops: list[str] = []
    for gs in node.findall(".//a:gs", NS):
        color = color_from_node(gs, theme)
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


def gradient_signature(node: ET.Element | None) -> str | None:
    if node is None:
        return None
    colors: list[str] = []
    for gs in node.findall(".//a:gs", NS):
        color = color_from_node(gs)
        if color:
            colors.append(color)
    return "gradient:" + "|".join(colors) if colors else None


def shape_text(shape: ET.Element) -> str:
    parts = [node.text.strip() for node in shape.findall(".//a:t", NS) if node.text and node.text.strip()]
    return " ".join(parts)


def is_placeholder(shape: ET.Element) -> bool:
    return shape.find(".//p:ph", NS) is not None


def target_hash(zf: zipfile.ZipFile, target: str) -> str | None:
    if target not in zf.namelist():
        return None
    return hashlib.sha1(zf.read(target)).hexdigest()


def theme_colors(zf: zipfile.ZipFile) -> dict[str, str]:
    root = read_xml(zf, "ppt/theme/theme1.xml")
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
        color = hex_color(srgb.get("val") if srgb is not None else None)
        if not color and sys_color is not None:
            color = hex_color(sys_color.get("lastClr") or sys_color.get("val"))
        if color:
            colors[key] = color
    return colors


def canvas(zf: zipfile.ZipFile) -> dict[str, float | str]:
    root = read_xml(zf, "ppt/presentation.xml")
    width_emu = 12192000
    height_emu = 6858000
    if root is not None:
        size = root.find("p:sldSz", NS)
        if size is not None:
            width_emu = int(size.get("cx", str(width_emu)))
            height_emu = int(size.get("cy", str(height_emu)))
    width_px = emu_to_px(width_emu)
    height_px = emu_to_px(height_emu)
    return {
        "width_emu": width_emu,
        "height_emu": height_emu,
        "width_px": round(width_px, 2),
        "height_px": round(height_px, 2),
        "aspect_ratio": round(width_px / height_px, 4) if height_px else "unknown",
    }


def slide_layout_path(zf: zipfile.ZipFile, slide_path: str) -> str | None:
    for rel in relationships(zf, slide_path).values():
        if rel.get("type", "").endswith("/slideLayout"):
            return rel.get("target")
    return None


def slide_master_path(zf: zipfile.ZipFile, layout_path: str | None) -> str | None:
    if not layout_path:
        return None
    for rel in relationships(zf, layout_path).values():
        if rel.get("type", "").endswith("/slideMaster"):
            return rel.get("target")
    return None


def is_citation(text: str) -> bool:
    return bool(
        re.search(r"\b(et al\.?|arXiv|doi|DOI|Nature|Science|Advances|Proceedings)\b", text)
        or re.search(r"\b(19|20)\d{2}\b", text)
        or re.search(r"[\u4e00-\u9fff]{1,8}等[，,]?", text)
        or re.search(r"(物理学报|化学进展|科学通报|中国科学|高等学校化学学报|中华医学|学报|期刊|研究)", text)
        or "[J]" in text
    )


def json_ready(value: Any) -> Any:
    return value
