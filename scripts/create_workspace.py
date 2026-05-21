#!/usr/bin/env python3
"""Create a clean local workspace for a personal-style PPTX-to-HTML task."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


PROFILES = [
    "academic-report",
    "group-meeting",
    "paper-reading",
    "thesis-defense",
    "course-presentation",
    "research-notes",
    "recording",
]

LANGUAGES = ["zh", "en", "bilingual"]

PLANNING_FILES = {
    "claim-spine.md": "# Claim Spine\n\nWrite one slide-level claim per line before building final slides.\n",
    "proof-objects.md": (
        "# Proof Objects\n\nMap each slide claim to one proof object: figure, table, matrix, "
        "mechanism flow, timeline, equation, quote, or visual comparison.\n"
    ),
    "revision-log.md": (
        "# Revision Log\n\nRecord every update here: base file, new file, user request, "
        "preserved user edits, locked slides, and validation evidence.\n"
    ),
    "image-preferences.md": (
        "# Image Preferences\n\nRecord user image habits: source type, crop, placement, "
        "caption style, opacity, border treatment, and repeated choices.\n"
    ),
    "figure-plan.md": (
        "# Figure Plan\n\nRecord chart claims, data sources, chart type, code path, output path, "
        "and whether the figure is editable HTML/SVG or externally rendered.\n"
    ),
    "speaker-notes.md": "# Speaker Notes\n\nMirror generated or user-provided per-slide notes here when notes mode is used.\n",
}

JSON_PLANNING_FILES: dict[str, Any] = {
    "locked-slides.json": {"slides": {}},
    "image-inventory.json": {"images": []},
}


def write_text_once(path: Path, content: str) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def write_json_once(path: Path, payload: Any) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def create_workspace(
    workspace: Path,
    *,
    profile: str = "academic-report",
    language: str = "zh",
    slides: int = 12,
    motion: str = "none",
    fidelity: str = "master-faithful",
) -> dict[str, Any]:
    """Create directories and planning files for a deck task without overwriting user work."""
    if profile not in PROFILES:
        raise ValueError(f"Unsupported profile: {profile}")
    if language not in LANGUAGES:
        raise ValueError(f"Unsupported language: {language}")

    workspace = workspace.resolve()
    dirs = [
        "source",
        "source/reference-pptx",
        "source/notes",
        "source/data",
        "assets/reference",
        "assets/generated",
        "assets/user-images",
        "figures/_shared",
        "planning",
        "work/style-extract",
        "work/html-preview",
        "output/current",
        "output/versions",
        "output/previews",
        "validation/screenshots",
        "validation/contact-sheets",
    ]
    for folder in dirs:
        (workspace / folder).mkdir(parents=True, exist_ok=True)

    for name, content in PLANNING_FILES.items():
        write_text_once(workspace / "planning" / name, content)
    for name, payload in JSON_PLANNING_FILES.items():
        write_json_once(workspace / "planning" / name, payload)

    shared_readme = workspace / "figures" / "_shared" / "README.md"
    write_text_once(
        shared_readme,
        "# Figure Workflow\n\nCreate one folder per important figure with code/, data/, outputs/, sources/, and README.md when external chart generation is used.\n",
    )

    manifest = {
        "profile": profile,
        "language": language,
        "planned_slides": slides,
        "motion": motion,
        "fidelity": fidelity,
        "source_dir": "source",
        "reference_pptx_dir": "source/reference-pptx",
        "style_work_dir": "work/style-extract",
        "html_preview_dir": "work/html-preview",
        "current_output_dir": "output/current",
        "version_output_dir": "output/versions",
        "preview_output_dir": "output/previews",
        "validation_dir": "validation",
        "claim_spine": "planning/claim-spine.md",
        "proof_objects": "planning/proof-objects.md",
        "revision_log": "planning/revision-log.md",
        "locked_slides": "planning/locked-slides.json",
        "image_preferences": "planning/image-preferences.md",
        "image_inventory": "planning/image-inventory.json",
        "figure_plan": "planning/figure-plan.md",
        "speaker_notes": "planning/speaker-notes.md",
        "notes": [
            "Keep private source decks and notes inside this workspace, not inside the skill folder.",
            "Do not overwrite user-provided PPTX files, generated HTML, or manually edited revisions.",
            "Use output/versions for timestamped revisions when continuing a previously delivered deck.",
            "For generated charts, keep code, data, source notes, and outputs under figures/.",
            "Use locked-slides.json to preserve user-corrected slides across later revisions.",
        ],
    }
    manifest_path = workspace / "personal_style_ppt2html_task.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workspace", type=Path)
    parser.add_argument("--profile", choices=PROFILES, default="academic-report")
    parser.add_argument("--language", choices=LANGUAGES, default="zh")
    parser.add_argument("--slides", type=int, default=12)
    parser.add_argument("--motion", choices=["none", "subtle", "recording", "demo"], default="none")
    parser.add_argument("--fidelity", choices=["master-faithful", "academic-inspired", "notes", "recording"], default="master-faithful")
    args = parser.parse_args()

    manifest = create_workspace(
        args.workspace,
        profile=args.profile,
        language=args.language,
        slides=args.slides,
        motion=args.motion,
        fidelity=args.fidelity,
    )
    print(json.dumps({"workspace": str(args.workspace.resolve()), **manifest}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
