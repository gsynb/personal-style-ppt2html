#!/usr/bin/env python3
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from extract_pptx_style import analyze_pptx
from build_theme_css import css_from_profile
from make_asset_manifest import build_asset_manifest
from audit_html_layout import audit_html_file
from pptx_to_academic_html import convert_pptx_to_html
from mine_reusable_visuals import build_reusable_visual_registry
from prepare_imagegen_briefs import build_imagegen_briefs
from create_workspace import create_workspace


PRESENTATION_XML = """\
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>
"""

THEME_XML = """\
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
    </a:clrScheme>
  </a:themeElements>
</a:theme>
"""

SLIDE_1_XML = """\
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="3657600" y="274320"/><a:ext cx="4876800" cy="640080"/></a:xfrm>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="3200"><a:latin typeface="微软雅黑"/><a:solidFill><a:srgbClr val="262625"/></a:solidFill></a:rPr>
        <a:t>Graph Transformer</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="914400" y="5486400"/><a:ext cx="10058400" cy="365760"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="1200"><a:latin typeface="微软雅黑"/><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr>
        <a:t>Vaswani et al., Attention is all you need, 2017.</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""

SLIDE_2_XML = """\
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:pic/>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="640080"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="2800"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="0D0D0D"/></a:solidFill></a:rPr>
        <a:t>Method Overview</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""

SLIDE_1_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout12.xml"/>
</Relationships>
"""

LAYOUT_12_XML = """\
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="0" y="1006679"/><a:ext cx="12192000" cy="83890"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:gradFill>
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="0E419C"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="B8C7E8"/></a:gs>
          </a:gsLst>
          <a:lin ang="0"/>
        </a:gradFill>
      </p:spPr>
    </p:sp>
    <p:pic>
      <p:spPr><a:xfrm><a:off x="106019" y="63907"/><a:ext cx="917438" cy="893906"/></a:xfrm></p:spPr>
      <p:blipFill><a:blip r:embed="rId2"/></p:blipFill>
    </p:pic>
  </p:spTree></p:cSld>
</p:sldLayout>
"""

LAYOUT_12_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""

SLIDE_2_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout12.xml"/>
</Relationships>
"""

LAYOUT_12_WITH_MASTER_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"""

SLIDE_MASTER_1_XML = """\
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="0" y="6200000"/><a:ext cx="12192000" cy="90000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="0E419C"/></a:solidFill>
      </p:spPr>
    </p:sp>
    <p:pic>
      <p:spPr><a:xfrm><a:off x="106019" y="63907"/><a:ext cx="917438" cy="893906"/></a:xfrm></p:spPr>
      <p:blipFill><a:blip r:embed="rId4"/></p:blipFill>
    </p:pic>
  </p:spTree></p:cSld>
</p:sldMaster>
"""

SLIDE_MASTER_1_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>
"""

SLIDE_MASTER_DUPLICATE_XML = """\
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:pic>
      <p:spPr><a:xfrm><a:off x="106019" y="63907"/><a:ext cx="917438" cy="893906"/></a:xfrm></p:spPr>
      <p:blipFill><a:blip r:embed="rId4"/></p:blipFill>
    </p:pic>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="914400" y="274320"/><a:ext cx="5000000" cy="400000"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr>
        <a:t>Click to edit Master title style</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldMaster>
"""

SLIDE_MASTER_DUPLICATE_RELS_XML = """\
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""

CHINESE_CITATION_SLIDE_XML = """\
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="7315200" cy="640080"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="3000"><a:latin typeface="微软雅黑"/></a:rPr>
        <a:t>中文参考文献页</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:spPr><a:xfrm><a:off x="914400" y="5486400"/><a:ext cx="10058400" cy="365760"/></a:xfrm></p:spPr>
      <p:txBody><a:p><a:r>
        <a:rPr sz="1200"><a:latin typeface="微软雅黑"/></a:rPr>
        <a:t>王明等，《物理学报》</a:t>
      </a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""

MANUAL_FOOTER_SHAPE_XML = """\
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="0" y="6700000"/><a:ext cx="12192000" cy="65000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="788AC6"/></a:solidFill>
      </p:spPr>
    </p:sp>
"""

FLOW_MODULES_XML = """\
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="1828800" y="1828800"/><a:ext cx="1524000" cy="457200"/></a:xfrm>
        <a:prstGeom prst="chevron"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
      </p:spPr>
    </p:sp>
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="3657600" y="1828800"/><a:ext cx="2133600" cy="731520"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="F7F7F7"/></a:solidFill>
        <a:ln w="9525"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>
      </p:spPr>
      <p:txBody><a:p><a:r><a:rPr sz="1800"><a:latin typeface="微软雅黑"/></a:rPr><a:t>关键任务</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:cxnSp>
      <p:spPr>
        <a:xfrm><a:off x="6096000" y="1981200"/><a:ext cx="1219200" cy="0"/></a:xfrm>
        <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
        <a:ln w="19050">
          <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
          <a:headEnd type="triangle"/>
        </a:ln>
      </p:spPr>
    </p:cxnSp>
"""


def write_minimal_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_1_XML)
        zf.writestr("ppt/slides/slide2.xml", SLIDE_2_XML)
        zf.writestr("ppt/media/image1.png", b"not-a-real-image")


def write_layout_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_1_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", SLIDE_1_RELS_XML)
        zf.writestr("ppt/slideLayouts/slideLayout12.xml", LAYOUT_12_XML)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout12.xml.rels", LAYOUT_12_RELS_XML)
        zf.writestr("ppt/media/image1.png", b"not-a-real-image")


def add_manual_footer(xml: str) -> str:
    return xml.replace("  </p:spTree></p:cSld>", MANUAL_FOOTER_SHAPE_XML + "  </p:spTree></p:cSld>")


def add_flow_modules(xml: str) -> str:
    return xml.replace("  </p:spTree></p:cSld>", FLOW_MODULES_XML + "  </p:spTree></p:cSld>")


def write_master_and_repeat_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", add_manual_footer(SLIDE_1_XML))
        zf.writestr("ppt/slides/slide2.xml", add_manual_footer(SLIDE_2_XML))
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", SLIDE_1_RELS_XML)
        zf.writestr("ppt/slides/_rels/slide2.xml.rels", SLIDE_2_RELS_XML)
        zf.writestr("ppt/slideLayouts/slideLayout12.xml", LAYOUT_12_XML)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout12.xml.rels", LAYOUT_12_WITH_MASTER_RELS_XML)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_1_XML)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_1_RELS_XML)
        zf.writestr("ppt/media/image1.png", b"layout-logo")
        zf.writestr("ppt/media/image2.png", b"master-logo")


def write_duplicate_master_layout_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_1_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", SLIDE_1_RELS_XML)
        zf.writestr("ppt/slideLayouts/slideLayout12.xml", LAYOUT_12_XML)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout12.xml.rels", LAYOUT_12_WITH_MASTER_RELS_XML)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER_DUPLICATE_XML)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_DUPLICATE_RELS_XML)
        zf.writestr("ppt/media/image1.png", b"same-logo")


def write_chinese_citation_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", CHINESE_CITATION_SLIDE_XML)


def write_duplicate_media_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_1_XML)
        zf.writestr("ppt/media/image1.png", b"same-image")
        zf.writestr("ppt/media/image2.png", b"same-image")


def write_flow_modules_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", add_flow_modules(SLIDE_1_XML))
        zf.writestr("ppt/slides/slide2.xml", add_flow_modules(SLIDE_2_XML))


def write_single_flow_modules_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML)
        zf.writestr("ppt/theme/theme1.xml", THEME_XML)
        zf.writestr("ppt/slides/slide1.xml", add_flow_modules(SLIDE_1_XML))


class AcademicHtmlToolTests(unittest.TestCase):
    def test_create_workspace_scaffolds_revision_planning_and_outputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            workspace = Path(tmp) / "mlip-html-deck"
            manifest = create_workspace(workspace, profile="group-meeting", language="zh", slides=15)

            self.assertEqual(manifest["profile"], "group-meeting")
            self.assertEqual(manifest["language"], "zh")
            self.assertEqual(manifest["planned_slides"], 15)
            self.assertTrue((workspace / "source").is_dir())
            self.assertTrue((workspace / "work" / "style-extract").is_dir())
            self.assertTrue((workspace / "output" / "versions").is_dir())
            self.assertTrue((workspace / "validation" / "screenshots").is_dir())
            self.assertTrue((workspace / "planning" / "claim-spine.md").exists())
            self.assertTrue((workspace / "planning" / "proof-objects.md").exists())
            self.assertTrue((workspace / "planning" / "revision-log.md").exists())
            self.assertTrue((workspace / "planning" / "locked-slides.json").exists())
            self.assertTrue((workspace / "personal_style_ppt2html_task.json").exists())

    def test_create_workspace_does_not_overwrite_existing_planning_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            workspace = Path(tmp) / "deck"
            existing = workspace / "planning" / "revision-log.md"
            existing.parent.mkdir(parents=True)
            existing.write_text("keep this revision history", encoding="utf-8")

            create_workspace(workspace)

            self.assertEqual(existing.read_text(encoding="utf-8"), "keep this revision history")

    def test_analyze_pptx_extracts_style_signals(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "sample.pptx"
            write_minimal_pptx(pptx)

            profile = analyze_pptx(pptx)

        self.assertEqual(profile["source"]["slide_count"], 2)
        self.assertEqual(profile["canvas"]["aspect_ratio_label"], "16:9")
        self.assertEqual(profile["theme"]["colors"]["accent1"], "#4472C4")
        self.assertEqual(profile["typography"]["dominant_font"], "微软雅黑")
        self.assertEqual(profile["slides"][0]["title_candidate"]["text"], "Graph Transformer")
        self.assertEqual(profile["slides"][0]["layout_class"], "centered-title-with-footer-citation")
        self.assertEqual(profile["slides"][1]["image_count"], 1)

    def test_css_from_profile_emits_academic_tokens(self):
        profile = {
            "typography": {"dominant_font": "微软雅黑"},
            "colors": {
                "background": "#FFFFFF",
                "text": "#262625",
                "accent": "#4472C4",
                "muted": "#666666",
            },
            "canvas": {"aspect_ratio_label": "16:9"},
        }

        css = css_from_profile(profile)

        self.assertIn("--academic-bg: #FFFFFF;", css)
        self.assertIn("--academic-text: #262625;", css)
        self.assertIn("--academic-accent: #4472C4;", css)
        self.assertIn('font-family: "微软雅黑"', css)

    def test_asset_manifest_extracts_pptx_media(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "sample.pptx"
            asset_dir = tmp_path / "assets"
            write_minimal_pptx(pptx)

            manifest = build_asset_manifest(pptx, asset_dir)

            self.assertEqual(manifest["source"], str(pptx.resolve()))
            self.assertEqual(len(manifest["assets"]), 1)
            self.assertTrue((asset_dir / "image1.png").exists())

    def test_asset_manifest_hashes_and_marks_duplicate_media(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "duplicates.pptx"
            asset_dir = tmp_path / "assets"
            write_duplicate_media_pptx(pptx)

            manifest = build_asset_manifest(pptx, asset_dir)

        self.assertEqual(len(manifest["assets"]), 2)
        self.assertEqual(manifest["assets"][0]["sha1"], manifest["assets"][1]["sha1"])
        self.assertEqual(manifest["assets"][1]["duplicate_of"], "image1.png")

    def test_audit_html_file_checks_academic_basics(self):
        with tempfile.TemporaryDirectory() as tmp:
            html = Path(tmp) / "index.html"
            html.write_text(
                """<!doctype html>
<html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>@media print { .academic-slide { break-after: page; } }</style>
</head><body><main class="academic-deck"><section class="academic-slide"></section></main></body></html>
""",
                encoding="utf-8",
            )

            report = audit_html_file(html)

        self.assertEqual(report["errors"], [])
        self.assertEqual(report["checks"]["has_viewport"], True)
        self.assertEqual(report["checks"]["has_print_css"], True)

    def test_audit_html_file_reads_linked_css(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            (tmp_path / "theme.css").write_text("@media print { .academic-slide { break-after: page; } }", encoding="utf-8")
            html = tmp_path / "index.html"
            html.write_text(
                """<!doctype html>
<html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Deck</title>
<link rel="stylesheet" href="./theme.css">
</head><body><main class="academic-deck"><section class="academic-slide"></section></main></body></html>
""",
                encoding="utf-8",
            )

            report = audit_html_file(html)

        self.assertEqual(report["errors"], [])
        self.assertEqual(report["checks"]["has_print_css"], True)

    def test_audit_warns_for_missing_img_alt_and_tiny_font(self):
        with tempfile.TemporaryDirectory() as tmp:
            html = Path(tmp) / "index.html"
            html.write_text(
                """<!doctype html>
<html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Deck</title>
<style>@media print { .academic-slide { break-after: page; } } .tiny { font-size: 8px; }</style>
</head><body><main class="academic-deck"><section class="academic-slide"><img src="figure.png"><p class="tiny">too small</p></section></main></body></html>
""",
                encoding="utf-8",
            )

            report = audit_html_file(html)

        self.assertEqual(report["errors"], [])
        self.assertIn("Image tag is missing an alt attribute.", report["warnings"])
        self.assertIn("Tiny font-size detected below 12px.", report["warnings"])

    def test_pptx_to_html_includes_reusable_layout_logo_and_bar(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "layout.pptx"
            out = tmp_path / "html"
            write_layout_pptx(pptx)

            index = convert_pptx_to_html(pptx, out)
            html_text = index.read_text(encoding="utf-8")
            css_text = (out / "style.css").read_text(encoding="utf-8")

            self.assertIn('src="assets/image1.png"', html_text)
            self.assertIn("linear-gradient", html_text + css_text)
            self.assertIn("pptx-layout", html_text)

    def test_pptx_to_html_includes_slide_master_reusable_elements(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "master.pptx"
            out = tmp_path / "html"
            write_master_and_repeat_pptx(pptx)

            index = convert_pptx_to_html(pptx, out)
            html_text = index.read_text(encoding="utf-8")

        self.assertIn("pptx-master", html_text)
        self.assertIn('src="assets/image2.png"', html_text)

    def test_pptx_to_html_deduplicates_master_layout_and_filters_english_placeholders(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "duplicate.pptx"
            out = tmp_path / "html"
            write_duplicate_master_layout_pptx(pptx)

            index = convert_pptx_to_html(pptx, out)
            html_text = index.read_text(encoding="utf-8")

        self.assertEqual(html_text.count('src="assets/image1.png"'), 1)
        self.assertNotIn("Click to edit Master title style", html_text)

    def test_extract_pptx_style_detects_chinese_citation_footer_without_year(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "chinese-citation.pptx"
            write_chinese_citation_pptx(pptx)

            profile = analyze_pptx(pptx)

        self.assertEqual(profile["style_summary"]["citation_footer_slides"], 1)
        self.assertIn("王明等", profile["slides"][0]["citation_candidates"][0])

    def test_reusable_visual_registry_finds_master_layout_and_manual_repeats(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "master.pptx"
            write_master_and_repeat_pptx(pptx)

            registry = build_reusable_visual_registry(pptx, min_occurrences=2)

        elements = registry["reusable_elements"]
        roles = {element["role"] for element in elements}
        source_levels = {element["source_level"] for element in elements}
        self.assertIn("institution_logo", roles)
        self.assertIn("header_rule", roles)
        self.assertIn("slide_master", source_levels)
        self.assertTrue(
            any(
                element["reuse_level"] == "manual_repeat"
                and element["role"] == "footer_rule"
                and element["occurrence_count"] == 2
                for element in elements
            )
        )

    def test_reusable_visual_registry_labels_flow_arrows_connectors_and_modules(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "flow-modules.pptx"
            write_flow_modules_pptx(pptx)

            registry = build_reusable_visual_registry(pptx, min_occurrences=2)

        elements = registry["reusable_elements"]
        by_role = {element["role"]: element for element in elements}
        self.assertEqual(by_role["flow_arrow"]["geometry"], "chevron")
        self.assertEqual(by_role["flow_connector"]["element_type"], "connector")
        self.assertEqual(by_role["flow_connector"]["arrow_head"], "triangle")
        self.assertEqual(by_role["text_module"]["geometry"], "roundRect")

    def test_reusable_visual_registry_keeps_single_slide_flow_grammar(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "single-flow-modules.pptx"
            write_single_flow_modules_pptx(pptx)

            registry = build_reusable_visual_registry(pptx, min_occurrences=2)

        roles = {element["role"] for element in registry["reusable_elements"]}
        self.assertIn("flow_arrow", roles)
        self.assertIn("flow_connector", roles)
        self.assertIn("text_module", roles)

    def test_prepare_imagegen_briefs_uses_style_evidence_and_safety_boundary(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "master.pptx"
            write_master_and_repeat_pptx(pptx)

            profile = analyze_pptx(pptx)
            registry = build_reusable_visual_registry(pptx, min_occurrences=2)
            briefs = build_imagegen_briefs(profile, registry)

        self.assertEqual(briefs["generation_mode"], "optional-agent-imagegen")
        self.assertEqual(len(briefs["asset_briefs"]), 3)
        self.assertIn("institution_logo", briefs["style_fingerprint"]["reusable_roles"])
        self.assertIn("#4472C4", briefs["asset_briefs"][0]["prompt"])
        self.assertIn("No text, no captions, no logos", briefs["asset_briefs"][0]["prompt"])
        self.assertIn("raw-slide-titles-excluded", briefs["style_fingerprint"]["content_cue_policy"])
        self.assertNotIn("Graph Transformer", briefs["asset_briefs"][0]["prompt"])
        self.assertEqual(briefs["asset_briefs"][0]["output_filename"], "assets/generated/style-cover-backdrop.png")

    def test_prepare_imagegen_briefs_adds_shape_asset_sheets_for_flow_grammar(self):
        with tempfile.TemporaryDirectory() as tmp:
            pptx = Path(tmp) / "flow-modules.pptx"
            write_flow_modules_pptx(pptx)

            profile = analyze_pptx(pptx)
            registry = build_reusable_visual_registry(pptx, min_occurrences=2)
            briefs = build_imagegen_briefs(profile, registry)

        ids = {brief["id"] for brief in briefs["asset_briefs"]}
        prompts = "\n".join(brief["prompt"] for brief in briefs["asset_briefs"])
        self.assertIn("vector-like-flow-elements", ids)
        self.assertIn("vector-like-module-panels", ids)
        self.assertIn("flow_arrow:chevron", briefs["style_fingerprint"]["reusable_geometry"])
        self.assertIn("transparent-background sheet", prompts)
        self.assertIn("No text, no captions, no logos", prompts)

    def test_pptx_to_html_supports_motion_and_image_optimization(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "layout.pptx"
            out = tmp_path / "html"
            write_layout_pptx(pptx)

            index = convert_pptx_to_html(pptx, out, motion="recording")
            html_text = index.read_text(encoding="utf-8")
            css_text = (out / "style.css").read_text(encoding="utf-8")

            self.assertIn('data-motion="recording"', html_text)
            self.assertIn('loading="lazy"', html_text)
            self.assertIn('decoding="async"', html_text)
            self.assertIn("@media (prefers-reduced-motion: reduce)", css_text)
            self.assertIn("pptx-enter-recording", css_text)

    def test_audit_warns_when_animation_lacks_reduced_motion_guard(self):
        with tempfile.TemporaryDirectory() as tmp:
            html = Path(tmp) / "index.html"
            html.write_text(
                """<!doctype html>
<html><head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
@media print { .academic-slide { break-after: page; } }
.academic-slide { animation: fadeIn 300ms ease both; }
</style>
</head><body><main class="academic-deck"><section class="academic-slide"></section></main></body></html>
""",
                encoding="utf-8",
            )

            report = audit_html_file(html)

        self.assertEqual(report["errors"], [])
        self.assertIn("Animation CSS is missing a prefers-reduced-motion guard.", report["warnings"])

    def test_run_pipeline_creates_full_output_set(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pptx = tmp_path / "sample.pptx"
            out = tmp_path / "pipeline"
            write_layout_pptx(pptx)

            result = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT_DIR / "run_pipeline.py"),
                    str(pptx),
                    "-o",
                    str(out),
                    "--motion",
                    "none",
                ],
                check=False,
                text=True,
                capture_output=True,
            )

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertTrue((out / "style-profile.json").exists())
            self.assertTrue((out / "reference-assets" / "manifest.json").exists())
            self.assertTrue((out / "asset-registry.json").exists())
            self.assertTrue((out / "imagegen-briefs.json").exists())
            self.assertTrue((out / "theme.generated.css").exists())
            self.assertTrue((out / "html-preview" / "index.html").exists())

    def test_template_defaults_to_no_motion_for_conservative_academic_output(self):
        template = (SCRIPT_DIR.parent / "assets" / "academic-html-template" / "index.html").read_text(
            encoding="utf-8"
        )

        self.assertIn('data-motion="none"', template)


if __name__ == "__main__":
    unittest.main()
