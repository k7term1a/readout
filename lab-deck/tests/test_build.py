"""Smoke tests: every example builds in every nav x density combination without layout warnings."""
import json
import sys
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.shapes.connector import Connector

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "skills/2slide/scripts"))
import build_deck as bd  # noqa: E402

EXAMPLES = [ROOT / "skills/2slide/assets/example-deck.json", ROOT / "examples/progress-report/deck.json"]


@pytest.mark.parametrize("spec_path", EXAMPLES, ids=lambda p: p.parent.name)
@pytest.mark.parametrize("nav", ["single", "double"])
@pytest.mark.parametrize("density", list(bd.DENSITIES))
def test_builds_without_layout_warnings(tmp_path, spec_path, nav, density):
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    out = tmp_path / "out.pptx"
    bd.render(spec, spec_path.parent, density, nav, str(out))
    layout_warnings = [w for w in bd.WARN if "image missing" not in w]
    assert not layout_warnings, layout_warnings
    assert len(Presentation(str(out)).slides) == 1 + len(spec["slides"])


def test_active_tab_matches_chapter(tmp_path):
    spec = {"chapters": ["甲", "乙"], "cover": {"title": "t"},
            "slides": [{"type": "content", "chapter": "乙", "tag": "x", "one_line": "y"}]}
    out = tmp_path / "out.pptx"
    bd.render(spec, tmp_path, "visual", "single", str(out))
    slide = Presentation(str(out)).slides[1]
    fills = {sh.text_frame.text: str(sh.fill.fore_color.rgb) for sh in slide.shapes
             if sh.has_text_frame and sh.text_frame.text in ("甲", "乙")}
    assert fills == {"乙": bd.DEFAULT_THEME["colors"]["active"], "甲": bd.DEFAULT_THEME["colors"]["inactive"]}


C = bd.DEFAULT_THEME["colors"]
DOUBLE_SPEC = {
    "chapters": [{"name": "介紹", "sections": ["背景", "動機"]}, {"name": "結論", "sections": []}],
    "cover": {"title": "t"},
    "slides": [{"type": "content", "chapter": "介紹", "section": "動機", "tag": "一、標籤", "kind": "problem",
                "title": "黑字標題", "subtitle": "灰色說明", "one_line": "y"},
               {"type": "content", "chapter": "結論", "tag": "x", "one_line": "z"}],
}


def build(tmp_path, spec, nav, density="visual"):
    out = tmp_path / "out.pptx"
    bd.render(spec, tmp_path, density, nav, str(out))
    return list(Presentation(str(out)).slides)


def texts(slide):
    return {sh.text_frame.text: sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text}


def test_double_nav_tabs_and_sections(tmp_path):
    t = texts(build(tmp_path, DOUBLE_SPEC, "double")[1])
    fill = {k: str(t[k].fill.fore_color.rgb) for k in ("介紹", "結論", "背景", "動機")}
    assert fill == {"介紹": C["active"], "結論": C["inactive_double"], "動機": C["active"], "背景": C["inactive_double"]}
    assert all(str(t[k].text_frame.paragraphs[0].runs[0].font.color.rgb) == "FFFFFF" for k in fill)


def test_double_rule_breaks_around_sections(tmp_path):
    slides = build(tmp_path, DOUBLE_SPEC, "double")
    rules = lambda s: sorted((c.begin_x, c.end_x) for c in s.shapes if isinstance(c, Connector))
    with_secs, without = rules(slides[1]), rules(slides[2])
    pills = [sh for sh in slides[1].shapes if sh.has_text_frame and sh.text_frame.text in ("背景", "動機")]
    assert len(with_secs) == 2 and len(without) == 1
    (l0, l1), (r0, r1) = with_secs
    assert l1 < min(p.left for p in pills) and r0 > max(p.left + p.width for p in pills)
    # the section pills straddle the rule
    rule_y = next(c for c in slides[1].shapes if isinstance(c, Connector)).begin_y
    assert all(p.top < rule_y < p.top + p.height for p in pills)


def test_double_uses_black_title_instead_of_tag(tmp_path):
    t = texts(build(tmp_path, DOUBLE_SPEC, "double")[1])
    assert "一、標籤" not in t and "/" not in t
    run = t["黑字標題"].text_frame.paragraphs[0].runs[0]
    assert run.font.bold and str(run.font.color.rgb) == C["dark"]
    assert "灰色說明" in t
    single = texts(build(tmp_path, DOUBLE_SPEC, "single")[1])
    assert "一、標籤" in single and "黑字標題" not in single


def slidenum_fields(slide):
    return slide._element.findall(".//" + bd.qn("a:fld"))


def test_page_numbers_double_only_and_not_on_cover(tmp_path):
    slides = build(tmp_path, DOUBLE_SPEC, "double")
    assert not slidenum_fields(slides[0])
    assert [slidenum_fields(s)[0].get("type") for s in slides[1:]] == ["slidenum", "slidenum"]
    assert all(not slidenum_fields(s) for s in build(tmp_path, DOUBLE_SPEC, "single"))


def test_divider_is_rejected(tmp_path):
    spec = {**DOUBLE_SPEC, "slides": [{"type": "divider", "chapter": "介紹"}] + DOUBLE_SPEC["slides"]}
    assert len(build(tmp_path, spec, "double")) == 3
    assert any("divider" in w for w in bd.WARN)


def test_no_shadow_on_rules(tmp_path):
    for c in (sh for sh in build(tmp_path, DOUBLE_SPEC, "double")[1].shapes if isinstance(sh, Connector)):
        assert c._element.spPr.find(bd.qn("a:effectLst")) is not None
