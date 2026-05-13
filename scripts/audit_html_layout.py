#!/usr/bin/env python3
"""Static checks for generated academic HTML presentations."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


def _linked_css_text(html_path: Path, html: str) -> str:
    css_parts: list[str] = []
    for href in re.findall(r"<link[^>]+href=[\"']([^\"']+\.css(?:\?[^\"']*)?)[\"'][^>]*>", html, re.I):
        clean_href = href.split("?", 1)[0]
        if re.match(r"https?://", clean_href, re.I):
            continue
        css_path = (html_path.parent / clean_href).resolve()
        try:
            if css_path.is_file():
                css_parts.append(css_path.read_text(encoding="utf-8"))
        except UnicodeDecodeError:
            continue
    return "\n".join(css_parts)


def audit_html_file(path: str | Path) -> dict[str, Any]:
    html_path = Path(path).expanduser().resolve()
    text = html_path.read_text(encoding="utf-8")
    css_text = _linked_css_text(html_path, text)
    combined_text = text + "\n" + css_text
    checks = {
        "has_doctype": bool(re.search(r"<!doctype html>", text, re.I)),
        "has_viewport": bool(re.search(r"<meta[^>]+name=[\"']viewport[\"']", text, re.I)),
        "has_academic_deck": "academic-deck" in text,
        "has_academic_slide": "academic-slide" in text,
        "has_print_css": bool(re.search(r"@media\s+print", combined_text, re.I)),
        "has_title": bool(re.search(r"<title>[^<]+</title>", text, re.I)),
    }
    errors: list[str] = []
    warnings: list[str] = []

    required = {
        "has_doctype": "Missing <!doctype html>.",
        "has_viewport": "Missing responsive viewport meta tag.",
        "has_academic_deck": "Missing .academic-deck root container.",
        "has_academic_slide": "Missing .academic-slide sections.",
        "has_print_css": "Missing @media print rules for PDF/export use.",
    }
    for key, message in required.items():
        if not checks[key]:
            errors.append(message)
    if not checks["has_title"]:
        warnings.append("Missing document title.")
    inherited_pptx_visuals = "pptx-slide" in text or "pptx-layout" in text
    if re.search(r"linear-gradient|radial-gradient|blur\(", combined_text, re.I) and not inherited_pptx_visuals:
        warnings.append("Decorative gradients or blur detected; keep academic templates restrained.")
    has_motion_css = bool(re.search(r"\b(animation|transition)\s*:", combined_text, re.I))
    has_reduced_motion_guard = bool(re.search(r"prefers-reduced-motion\s*:\s*reduce", combined_text, re.I))
    if has_motion_css and not has_reduced_motion_guard:
        warnings.append("Animation CSS is missing a prefers-reduced-motion guard.")
    if re.search(r"\banimation(?:-[^:]+)?\s*:[^;{}]*\binfinite\b", combined_text, re.I):
        warnings.append("Infinite animation detected; avoid looping motion in academic slides.")
    motion_values = re.findall(r"data-motion=[\"']([^\"']+)[\"']", text, re.I)
    allowed_motion = {"none", "subtle", "recording", "demo"}
    for value in motion_values:
        if value not in allowed_motion:
            warnings.append(f"Unknown data-motion preset: {value}.")
    if len(re.findall(r"class=[\"'][^\"']*card", text, re.I)) > 8:
        warnings.append("Many card-like components detected; academic slides should prioritize figures and claims.")

    return {"path": str(html_path), "checks": checks, "errors": errors, "warnings": warnings}


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit an academic HTML presentation for baseline structure.")
    parser.add_argument("html", type=Path)
    parser.add_argument("--fail-on-warning", action="store_true")
    args = parser.parse_args()

    report = audit_html_file(args.html)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if report["errors"] or (args.fail_on_warning and report["warnings"]):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
