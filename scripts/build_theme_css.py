#!/usr/bin/env python3
"""Build CSS design tokens from an academic PPTX style profile."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def _value(profile: dict[str, Any], path: tuple[str, ...], fallback: str) -> str:
    node: Any = profile
    for key in path:
        if not isinstance(node, dict) or key not in node:
            return fallback
        node = node[key]
    return str(node or fallback)


def _font_stack(font: str) -> str:
    font = font.replace('"', '\\"')
    fallback = '"Microsoft YaHei", "Noto Sans SC", "PingFang SC", Arial, sans-serif'
    if not font or font == "system-ui":
        return f"system-ui, {fallback}"
    return f'"{font}", {fallback}'


def css_from_profile(profile: dict[str, Any]) -> str:
    background = _value(profile, ("colors", "background"), "#FFFFFF")
    text = _value(profile, ("colors", "text"), "#202124")
    accent = _value(profile, ("colors", "accent"), "#315FA8")
    muted = _value(profile, ("colors", "muted"), "#666666")
    font = _value(profile, ("typography", "dominant_font"), "system-ui")
    font_stack = _font_stack(font)
    aspect = _value(profile, ("canvas", "aspect_ratio_label"), "16:9")
    ratio = "16 / 9" if aspect == "16:9" else "4 / 3" if aspect == "4:3" else "16 / 9"

    return f"""\
:root {{
  --academic-bg: {background};
  --academic-text: {text};
  --academic-accent: {accent};
  --academic-muted: {muted};
  --academic-rule: rgba(32, 33, 36, 0.14);
  --academic-rule: color-mix(in srgb, var(--academic-text) 14%, transparent);
  --academic-font: {font_stack};
  --academic-aspect: {ratio};
  --academic-slide-radius: 0px;
}}

html,
body {{
  margin: 0;
  min-height: 100%;
  background: var(--academic-bg);
  color: var(--academic-text);
  font-family: {font_stack};
}}

.academic-deck {{
  background: var(--academic-bg);
}}

.academic-slide {{
  position: relative;
  aspect-ratio: var(--academic-aspect);
  background: var(--academic-bg);
  color: var(--academic-text);
  overflow: hidden;
}}

.academic-title {{
  color: var(--academic-text);
  font-weight: 650;
  letter-spacing: 0;
}}

.academic-accent {{
  color: var(--academic-accent);
}}

.academic-citation {{
  color: var(--academic-muted);
  font-size: clamp(12px, 1.1vw, 14px);
}}

.academic-rule {{
  border-color: var(--academic-rule);
}}
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Build academic CSS tokens from style-profile.json.")
    parser.add_argument("profile", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    args = parser.parse_args()

    profile = json.loads(args.profile.read_text(encoding="utf-8"))
    css = css_from_profile(profile)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(css, encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
