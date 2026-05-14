#!/usr/bin/env python3
"""Extract PPTX media files and write a reusable asset manifest."""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from typing import Any


def _image_size(path: Path) -> dict[str, int] | None:
    try:
        from PIL import Image

        with Image.open(path) as image:
            return {"width": image.width, "height": image.height}
    except Exception:
        return None


def build_asset_manifest(pptx: str | Path, asset_dir: str | Path) -> dict[str, Any]:
    pptx_path = Path(pptx).expanduser().resolve()
    output_dir = Path(asset_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    assets: list[dict[str, Any]] = []
    first_seen_by_hash: dict[str, str] = {}
    with zipfile.ZipFile(pptx_path) as zf:
        media_names = sorted(name for name in zf.namelist() if name.startswith("ppt/media/") and not name.endswith("/"))
        for name in media_names:
            data = zf.read(name)
            sha1 = hashlib.sha1(data).hexdigest()
            target = output_dir / Path(name).name
            target.write_bytes(data)
            item: dict[str, Any] = {
                "file": target.name,
                "source_path": name,
                "extension": target.suffix.lower().lstrip("."),
                "bytes": target.stat().st_size,
                "sha1": sha1,
            }
            if sha1 in first_seen_by_hash:
                item["duplicate_of"] = first_seen_by_hash[sha1]
            else:
                first_seen_by_hash[sha1] = target.name
            size = _image_size(target)
            if size:
                item.update(size)
            assets.append(item)

    return {"source": str(pptx_path), "asset_dir": str(output_dir), "assets": assets}


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract PPTX media assets and write manifest.json.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("-o", "--output-dir", type=Path, required=True)
    args = parser.parse_args()

    manifest = build_asset_manifest(args.pptx, args.output_dir)
    manifest_path = args.output_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(manifest_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
