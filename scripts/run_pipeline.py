#!/usr/bin/env python3
"""Run the full personal-style PPTX to HTML extraction pipeline."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from build_theme_css import css_from_profile
from extract_pptx_style import analyze_pptx
from make_asset_manifest import build_asset_manifest
from mine_reusable_visuals import build_reusable_visual_registry
from pptx_to_academic_html import ALLOWED_MOTION, convert_pptx_to_html
from prepare_imagegen_briefs import build_imagegen_briefs


def run_pipeline(pptx: str | Path, output_dir: str | Path, motion: str = "none") -> dict[str, str]:
    pptx_path = Path(pptx).expanduser().resolve()
    out = Path(output_dir).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    profile_path = out / "style-profile.json"
    asset_dir = out / "reference-assets"
    registry_path = out / "asset-registry.json"
    imagegen_briefs_path = out / "imagegen-briefs.json"
    theme_path = out / "theme.generated.css"
    html_dir = out / "html-preview"

    profile = analyze_pptx(pptx_path)
    profile_path.write_text(json.dumps(profile, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    manifest = build_asset_manifest(pptx_path, asset_dir)
    (asset_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    registry = build_reusable_visual_registry(pptx_path)
    registry_path.write_text(json.dumps(registry, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    imagegen_briefs = build_imagegen_briefs(profile, registry)
    imagegen_briefs_path.write_text(
        json.dumps(imagegen_briefs, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    theme_path.write_text(css_from_profile(profile), encoding="utf-8")
    index_path = convert_pptx_to_html(pptx_path, html_dir, profile_path, motion=motion)

    return {
        "style_profile": str(profile_path),
        "asset_manifest": str(asset_dir / "manifest.json"),
        "asset_registry": str(registry_path),
        "imagegen_briefs": str(imagegen_briefs_path),
        "theme_css": str(theme_path),
        "html_index": str(index_path),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the full PPTX style extraction and HTML preview pipeline.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("-o", "--output-dir", type=Path, required=True)
    parser.add_argument(
        "--motion",
        choices=sorted(ALLOWED_MOTION),
        default="none",
        help="Motion preset for generated HTML preview.",
    )
    args = parser.parse_args()

    outputs = run_pipeline(args.pptx, args.output_dir, motion=args.motion)
    print(json.dumps(outputs, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
