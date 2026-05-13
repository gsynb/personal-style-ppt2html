#!/usr/bin/env python3
import json
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

MANUAL_FOOTER_SHAPE_XML = """\
    <p:sp>
      <p:spPr>
        <a:xfrm><a:off x="0" y="6700000"/><a:ext cx="12192000" cy="65000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="788AC6"/></a:solidFill>
      </p:spPr>
    </p:sp>
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


class AcademicHtmlToolTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
