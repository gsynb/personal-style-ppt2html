#!/usr/bin/env python3
"""Prepare safe image-generation briefs from extracted PPTX style evidence."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path
from typing import Any

HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _safe_color(value: Any, fallback: str) -> str:
    if isinstance(value, str) and HEX_RE.match(value):
        return value.upper()
    return fallback


def _colors(profile: dict[str, Any]) -> dict[str, str]:
    colors = profile.get("colors") if isinstance(profile.get("colors"), dict) else {}
    theme_colors = profile.get("theme", {}).get("colors", {}) if isinstance(profile.get("theme"), dict) else {}
    accent = colors.get("accent") or theme_colors.get("accent1")
    return {
        "background": _safe_color(colors.get("background"), "#FFFFFF"),
        "text": _safe_color(colors.get("text"), "#262625"),
        "accent": _safe_color(accent, "#315FA8"),
        "muted": _safe_color(colors.get("muted"), "#666666"),
    }


def _registry_roles(registry: dict[str, Any] | None) -> list[str]:
    if not registry:
        return []
    roles = Counter()
    for element in registry.get("reusable_elements", []):
        role = element.get("role")
        if isinstance(role, str) and role:
            roles[role] += 1
    return sorted(roles, key=lambda role: (-roles[role], role))


def _registry_geometry_summary(registry: dict[str, Any] | None, limit: int = 10) -> list[str]:
    if not registry:
        return []
    counts = Counter()
    for element in registry.get("reusable_elements", []):
        if not isinstance(element, dict):
            continue
        role = element.get("role")
        geometry = element.get("geometry")
        if role in {"flow_arrow", "flow_connector", "text_module", "module_box"} and geometry:
            counts[f"{role}:{geometry}"] += 1
    return [name for name, _ in counts.most_common(limit)]


def _layout_classes(profile: dict[str, Any]) -> list[str]:
    classes = Counter()
    for slide in profile.get("slides", []):
        layout_class = slide.get("layout_class") if isinstance(slide, dict) else None
        if isinstance(layout_class, str) and layout_class:
            classes[layout_class] += 1
    return sorted(classes, key=lambda item: (-classes[item], item))


def _title_terms(profile: dict[str, Any], limit: int = 6) -> list[str]:
    terms: list[str] = []
    seen: set[str] = set()
    for slide in profile.get("slides", []):
        if not isinstance(slide, dict):
            continue
        title = slide.get("title_candidate", {}).get("text")
        if not isinstance(title, str):
            continue
        clean = re.sub(r"\s+", " ", title).strip()
        if not clean or len(clean) > 80:
            continue
        key = clean.casefold()
        if key in seen:
            continue
        seen.add(key)
        terms.append(clean)
        if len(terms) >= limit:
            break
    return terms


def _style_fingerprint(
    profile: dict[str, Any],
    registry: dict[str, Any] | None,
    include_title_cues: bool,
) -> dict[str, Any]:
    typography = profile.get("typography", {}) if isinstance(profile.get("typography"), dict) else {}
    canvas = profile.get("canvas", {}) if isinstance(profile.get("canvas"), dict) else {}
    summary = profile.get("style_summary", {}) if isinstance(profile.get("style_summary"), dict) else {}
    return {
        "colors": _colors(profile),
        "dominant_font": typography.get("dominant_font") or "academic sans-serif",
        "aspect_ratio": canvas.get("aspect_ratio_label") or "16:9",
        "layout_classes": _layout_classes(profile),
        "reusable_roles": _registry_roles(registry),
        "reusable_geometry": _registry_geometry_summary(registry),
        "citation_footer_slides": summary.get("citation_footer_slides", 0),
        "title_terms": _title_terms(profile) if include_title_cues else [],
        "content_cue_policy": "title-cues-included" if include_title_cues else "raw-slide-titles-excluded",
    }


def _base_style_sentence(fingerprint: dict[str, Any]) -> str:
    colors = fingerprint["colors"]
    roles = ", ".join(fingerprint["reusable_roles"]) or "no explicit reusable visual roles detected"
    layouts = ", ".join(fingerprint["layout_classes"][:4]) or "balanced academic slide layouts"
    if fingerprint["title_terms"]:
        topic_clause = f"broad user-approved topic cues: {'; '.join(fingerprint['title_terms'][:4])}."
    else:
        topic_clause = "keep motifs abstract and do not infer factual subject matter from slide text."
    return (
        f"Match the extracted PPTX style: white or near-white academic canvas {colors['background']}, "
        f"primary text color {colors['text']}, restrained accent {colors['accent']}, muted support color "
        f"{colors['muted']}, {fingerprint['aspect_ratio']} landscape composition, dominant font mood "
        f"{fingerprint['dominant_font']}, layout rhythm: {layouts}, observed reusable roles: {roles}, "
        f"{topic_clause}"
    )


def _shape_grammar_sentence(fingerprint: dict[str, Any]) -> str:
    geometry = ", ".join(fingerprint.get("reusable_geometry", [])[:8])
    if not geometry:
        geometry = "rectangular modules, rounded module boxes, thin connector lines, and simple directional arrows"
    return (
        "Use the extracted reusable PPTX shape grammar as inspiration: "
        f"{geometry}. Keep shapes modular, clean, and easy to crop or place in HTML."
    )


def _negative_instruction() -> str:
    return (
        "No text, no captions, no logos, no seals, no institution marks, no watermarks, no UI chrome, "
        "no charts, no axes, no plotted data, no paper figures, no factual diagrams, no synthetic "
        "experimental results, and no decorative marketing style."
    )


def build_imagegen_briefs(
    profile: dict[str, Any],
    registry: dict[str, Any] | None = None,
    generated_dir: str = "assets/generated",
    include_title_cues: bool = False,
) -> dict[str, Any]:
    """Build image-generation briefs that an agent can execute when image tools are available."""
    fingerprint = _style_fingerprint(profile, registry, include_title_cues)
    style_sentence = _base_style_sentence(fingerprint)
    negative = _negative_instruction()
    generated = generated_dir.rstrip("/")

    briefs = [
        {
            "id": "cover-backdrop",
            "asset_type": "raster background",
            "safe_to_generate": True,
            "output_filename": f"{generated}/style-cover-backdrop.png",
            "usage": "Title slide or opening slide background behind editable HTML text.",
            "prompt": (
                "Create a restrained academic presentation cover background. "
                f"{style_sentence} Use a quiet white field, sparse thin-line geometry, subtle accent "
                "linework, and large calm negative space for title text. The result should feel like a "
                "serious research talk, not a poster or marketing hero. "
                f"{negative}"
            ),
            "integration_hint": "Use as an absolutely positioned background image at low opacity behind the title block.",
        },
        {
            "id": "section-divider-motif",
            "asset_type": "raster motif",
            "safe_to_generate": True,
            "output_filename": f"{generated}/style-section-divider.png",
            "usage": "Section divider, agenda transition, or chapter-opening slide.",
            "prompt": (
                "Create a minimal section-divider visual motif for an academic HTML slide deck. "
                f"{style_sentence} Use thin horizontal rules, sparse abstract research geometry, and "
                "one restrained accent band or motif that can coexist with an extracted real logo/header rule. "
                "Keep most of the canvas empty and presentation-safe. "
                f"{negative}"
            ),
            "integration_hint": "Place near the top or right edge of divider slides; keep official PPTX-derived logos separate.",
        },
        {
            "id": "figure-backplate-texture",
            "asset_type": "raster texture",
            "safe_to_generate": True,
            "output_filename": f"{generated}/style-figure-backplate.png",
            "usage": "Subtle backplate behind paper figures, method diagrams, or summary panels.",
            "prompt": (
                "Create a subtle academic figure backplate texture for HTML slides. "
                f"{style_sentence} Use near-white paper-like texture, very light grid or ruled-line structure, "
                "and low-contrast accent details. It must remain quiet enough that real figures, equations, "
                "and captions stay readable above it. "
                f"{negative}"
            ),
            "integration_hint": "Use behind content panels or figure frames with opacity below 0.18.",
        },
    ]

    roles = set(fingerprint.get("reusable_roles", []))
    if roles.intersection({"flow_arrow", "flow_connector", "text_module", "module_box"}):
        shape_sentence = _shape_grammar_sentence(fingerprint)
        briefs.extend(
            [
                {
                    "id": "vector-like-flow-elements",
                    "asset_type": "raster element sheet",
                    "safe_to_generate": True,
                    "output_filename": f"{generated}/style-flow-elements.png",
                    "usage": "Reusable non-factual arrows, chevrons, connectors, and process-flow accents for HTML slides.",
                    "prompt": (
                        "Create a transparent-background sheet of reusable vector-like academic presentation elements. "
                        f"{style_sentence} {shape_sentence} Include several chevrons, home-plate arrows, slim "
                        "connector arrows, small corner markers, and quiet accent rules. Use clean flat fills, crisp "
                        "edges, restrained shadows only if needed, and generous spacing between elements for later "
                        "cropping. "
                        f"{negative}"
                    ),
                    "integration_hint": "Crop individual elements or use the full sheet as a source image; decorative only, never for factual diagrams.",
                },
                {
                    "id": "vector-like-module-panels",
                    "asset_type": "raster element sheet",
                    "safe_to_generate": True,
                    "output_filename": f"{generated}/style-module-panels.png",
                    "usage": "Reusable non-factual module boxes, timeline cards, and emphasis panels for HTML slides.",
                    "prompt": (
                        "Create a transparent-background sheet of reusable vector-like module panels for serious "
                        f"academic HTML slides. {style_sentence} {shape_sentence} Include rounded rectangles, "
                        "outlined content panels, small label tabs, low-contrast callout blocks, and timeline nodes. "
                        "Leave every module blank with no embedded text, icons, numbers, or labels so HTML text can "
                        "be placed above it. "
                        f"{negative}"
                    ),
                    "integration_hint": "Use as decorative blank panels beneath editable HTML text; keep opacity and contrast conservative.",
                },
            ]
        )

    return {
        "schema_version": "1.0",
        "generation_mode": "optional-agent-imagegen",
        "style_fingerprint": fingerprint,
        "generation_policy": {
            "execute_only_when_image_generation_tool_is_available": True,
            "save_generated_assets_under": generated,
            "must_record_provenance": True,
            "allowed": [
                "cover backgrounds",
                "section divider motifs",
                "subtle research textures",
                "non-factual visual backplates",
            ],
            "forbidden": [
                "logos or official institution marks",
                "published paper figures",
                "charts or plotted data",
                "factual method diagrams with labels",
                "synthetic experimental evidence",
            ],
        },
        "asset_briefs": briefs,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepare safe image-generation briefs from a PPTX style profile.")
    parser.add_argument("style_profile", type=Path)
    parser.add_argument("--registry", type=Path, help="Optional asset-registry.json from mine_reusable_visuals.py.")
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument(
        "--generated-dir",
        default="assets/generated",
        help="Relative directory where generated images should be saved in the final HTML project.",
    )
    parser.add_argument(
        "--include-title-cues",
        action="store_true",
        help="Allow short slide title cues in prompts. Leave off by default to avoid leaking factual/private text.",
    )
    args = parser.parse_args()

    profile = json.loads(args.style_profile.read_text(encoding="utf-8"))
    registry = json.loads(args.registry.read_text(encoding="utf-8")) if args.registry else None
    briefs = build_imagegen_briefs(
        profile,
        registry,
        generated_dir=args.generated_dir,
        include_title_cues=args.include_title_cues,
    )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(briefs, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"imagegen_briefs": str(args.output.resolve())}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
