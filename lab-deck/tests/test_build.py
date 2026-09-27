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
    # left-aligned but not flush: a short stub of rule sits before the first pill
    first = min(p.left for p in pills)
    assert l0 < l1 < first < bd.Inches(bd.W / 4)
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


COVER = {"title": "中文標題", "byline": "報告人．用途", "venue": "NeurIPS 2022", "paper_title": "Paper Title",
         "authors": ["A. Author", "B. Author"], "affiliations": ["Some University"], "presenter": "王小明"}


def cover_of(tmp_path, nav, **cover):
    spec = {"chapters": ["甲"], "cover": {**COVER, **cover}, "slides": []}
    return texts(build(tmp_path, spec, nav)[0])


def test_cover_default_follows_nav(tmp_path):
    single = cover_of(tmp_path, "single")
    assert "中文標題" in single and "報告人．用途" in single and "NeurIPS 2022" not in single
    double = cover_of(tmp_path, "double")
    assert {"NeurIPS 2022", "Paper Title", "報告者：王小明"} <= set(double) and "中文標題" not in double


def test_cover_style_overrides_nav(tmp_path):
    assert "NeurIPS 2022" in cover_of(tmp_path, "single", style="paper")
    assert "中文標題" in cover_of(tmp_path, "double", style="centered")


def test_paper_cover_layout(tmp_path):
    t = cover_of(tmp_path, "double")
    badge, title = t["NeurIPS 2022"], t["Paper Title"]
    assert str(badge.fill.fore_color.rgb) == C["active"] and badge.top < title.top
    assert title.text_frame.paragraphs[0].runs[0].font.name == bd.DEFAULT_THEME["fonts"]["ref_latin"]
    authors, affil = t["A. Author\nB. Author"], t["Some University"]
    assert authors.top == affil.top and authors.left < affil.left  # two columns
    pres = t["報告者：王小明"]
    assert pres.left > bd.Inches(bd.W / 2) and pres.top > authors.top  # bottom-right
    assert "中文標題" in cover_of(tmp_path, "double", paper_title=None)


def test_unknown_cover_style_warns(tmp_path):
    cover_of(tmp_path, "double", style="fancy")
    assert any("cover.style" in w for w in bd.WARN)


def test_paper_cover_falls_back_to_byline(tmp_path):
    t = cover_of(tmp_path, "double", authors=None, paper_title=None)
    assert "中文標題" in t and "報告人．用途" in t


def bar_spec(conclusion, **extra):
    return {"chapters": ["甲"], "cover": {"title": "t"},
            "slides": [{"type": "content", "chapter": "甲", "tag": "x", "title": "標題", "full": ["段落"],
                        "short": ["條列"], "one_line": "一句話", "figs": [{"id": "Figure 1"}],
                        "conclusion": conclusion, **extra}]}


@pytest.mark.parametrize("nav", ["single", "double"])
@pytest.mark.parametrize("density", list(bd.DENSITIES))
def test_conclusion_bar_bottom_reserves_space(tmp_path, nav, density):
    slide = build(tmp_path, bar_spec("既有方法算不動", conclusion_pos="bottom"), nav, density)[1]
    bar = texts(slide)["既有方法算不動"]
    assert bar.name == "結論橫條" and str(bar.fill.fore_color.rgb) == C["problem"]
    assert str(bar.text_frame.paragraphs[0].runs[0].font.color.rgb) == "FFFFFF"
    assert bar.left == bd.Inches(0.6) and bar.width == bd.Inches(bd.W - 1.2)  # full width, same as the rule
    # everything else (except the page number) ends above the bar
    others = [sh for sh in slide.shapes if sh.shape_id != bar.shape_id and not isinstance(sh, Connector)
              and not slidenum_fields_in(sh) and sh.top > bd.Inches(1.1)]
    assert others and all(sh.top + sh.height <= bar.top for sh in others)


def slidenum_fields_in(shape):
    return shape._element.findall(".//" + bd.qn("a:fld"))


def test_no_bar_without_conclusion(tmp_path):
    assert all(sh.name != "結論橫條" for sh in build(tmp_path, bar_spec(None), "double")[1].shapes)


def test_long_conclusion_warns(tmp_path):
    build(tmp_path, bar_spec("太長" * 60), "double")
    assert any("conclusion bar" in w for w in bd.WARN)


def named(slide, name):
    return [sh for sh in slide.shapes if sh.name == name]


def timing_targets(slide):
    """(spid, nodeType) of every entrance effect in the slide's timing tree."""
    out = []
    for ctn in slide._element.iter(bd.qn("p:cTn")):
        if ctn.get("presetClass") == "entr":
            spid = next(ctn.iter(bd.qn("p:spTgt"))).get("spid")
            out.append((int(spid), ctn.get("nodeType"), ctn.get("presetID")))
    return out


def test_conclusion_fades_in_on_one_click_with_veil(tmp_path):
    slide = build(tmp_path, bar_spec("結論"), "double", "balanced")[1]
    (veil,), (bar,) = named(slide, "結論遮罩"), named(slide, "結論橫條")
    # fade (presetID 10): veil on click, bar with it
    assert timing_targets(slide) == [(veil.shape_id, "clickEffect", "10"), (bar.shape_id, "withEffect", "10")]
    # translucent white veil from under the heading down to the bar, drawn above the content
    alpha = veil.fill._xPr.find(bd.qn("a:solidFill"))[0].find(bd.qn("a:alpha"))
    assert str(veil.fill.fore_color.rgb) == "FFFFFF" and int(alpha.get("val")) == bd.VEIL_ALPHA * 1000
    assert veil.left == bar.left and veil.width == bar.width
    title = texts(slide)["標題"]
    assert veil.top >= title.top + title.height
    # z-order: veil above every content shape, bar above the veil (only the page number may come later)
    order = [sh.shape_id for sh in slide.shapes if not slidenum_fields_in(sh)]
    assert order[-2:] == [veil.shape_id, bar.shape_id]


@pytest.mark.parametrize("kind", ["problem", "solution", "neutral"])
def test_conclusion_kind_colours(tmp_path, kind):
    (bar,) = named(build(tmp_path, bar_spec("結論", conclusion_kind=kind), "double")[1], "結論橫條")
    assert str(bar.fill.fore_color.rgb) == C[kind]


def test_keyframes_split_into_before_and_after(tmp_path):
    out = tmp_path / "kf.pptx"
    bd.render(bar_spec("結論"), tmp_path, "balanced", "double", str(out), keyframes=True)
    cover, before, after = Presentation(str(out)).slides
    assert not named(before, "結論橫條") and not named(before, "結論遮罩")
    assert named(after, "結論橫條") and named(after, "結論遮罩")
    assert not timing_targets(before) and not timing_targets(after)  # no animation in keyframe mode
    # the content does not move between frames
    geom = lambda sl: [(sh.left, sh.top, sh.width, sh.height) for sh in sl.shapes
                       if sh.name not in ("結論橫條", "結論遮罩") and not slidenum_fields_in(sh)]
    assert geom(before) == geom(after)


def near(a, b):
    return abs(a - b) <= 2  # EMU rounding


def content_geometry(slide):
    return [(sh.left, sh.top, sh.width, sh.height) for sh in slide.shapes
            if sh.name not in ("結論橫條", "結論遮罩") and not slidenum_fields_in(sh)]


@pytest.mark.parametrize("nav", ["single", "double"])
@pytest.mark.parametrize("density", list(bd.DENSITIES))
def test_conclusion_centered_in_veil_by_default(tmp_path, nav, density):
    slide = build(tmp_path, bar_spec("結論"), nav, density)[1]
    (veil,), (bar,) = named(slide, "結論遮罩"), named(slide, "結論橫條")
    assert abs((veil.top + veil.height / 2) - (bar.top + bar.height / 2)) <= 1  # vertically centred
    assert near(veil.top + veil.height, bd.Inches(bd.H - 0.5))  # veil covers the whole body
    # no space reserved: content is laid out exactly as without a conclusion
    plain = build(tmp_path, bar_spec(None), nav, density)[1]
    assert content_geometry(slide) == content_geometry(plain)


def test_conclusion_pos_bottom_veil_stops_at_bar(tmp_path):
    slide = build(tmp_path, bar_spec("結論", conclusion_pos="bottom"), "double")[1]
    (veil,), (bar,) = named(slide, "結論遮罩"), named(slide, "結論橫條")
    assert veil.top + veil.height == bar.top and bar.top + bar.height == bd.Inches(bd.H - 0.5)


def test_conclusion_pos_deck_default_and_override(tmp_path):
    spec = bar_spec("結論")
    spec["style"] = {"conclusion_pos": "bottom"}
    (bar,) = named(build(tmp_path, spec, "double")[1], "結論橫條")
    assert near(bar.top + bar.height, bd.Inches(bd.H - 0.5))
    spec["slides"][0]["conclusion_pos"] = "center"
    (veil,) = named(build(tmp_path, spec, "double")[1], "結論遮罩")
    assert near(veil.top + veil.height, bd.Inches(bd.H - 0.5))


def test_unknown_conclusion_pos_warns(tmp_path):
    build(tmp_path, bar_spec("結論", conclusion_pos="top"), "double")
    assert any("conclusion_pos" in w for w in bd.WARN)


CONCEPTS = {"圖像層級": 1, "區域層級": 2}


def concept_spec(slides, concepts=CONCEPTS):
    return {"chapters": ["甲"], "concepts": concepts, "cover": {"title": "t"}, "slides": slides}


def runs_of(slide):
    return [r for sh in slide.shapes if sh.has_text_frame
            for p in sh.text_frame.paragraphs for r in p.runs]


def test_parse_marks():
    assert bd.parse_marks("a ==b== {{c}} {{c|d}}") == [
        ("a ", False, None), ("b", True, None), (" ", False, None), ("c", False, "c"), (" ", False, None),
        ("d", False, "c")]
    assert bd.plain_text("{{圖像層級|整張圖}}的特徵") == "整張圖的特徵"


def test_inline_concept_text_is_coloured(tmp_path):
    spec = concept_spec([{"type": "content", "chapter": "甲", "tag": "x", "title": "{{區域層級}}標題",
                          "short": ["比較{{圖像層級}}與{{區域層級|區域}}"]}])
    slide = build(tmp_path, spec, "double", "balanced")[1]
    runs = {r.text: r for r in runs_of(slide)}
    assert "{{" not in "".join(runs)  # marks never leak into the slide
    assert str(runs["圖像層級"].font.color.rgb) == C["concept1_fg"] and runs["圖像層級"].font.bold
    assert str(runs["區域"].font.color.rgb) == C["concept2_fg"]
    assert str(runs["區域層級"].font.color.rgb) == C["concept2_fg"]  # in the title too


def test_figure_and_table_row_concept_blocks(tmp_path):
    spec = concept_spec([
        {"type": "content", "chapter": "甲", "one_line": "y", "figs": [{"id": "Figure 1", "concept": "區域層級"}]},
        {"type": "content", "chapter": "甲", "short": ["t"],
         "table": [["方法", "分數"], ["{{圖像層級}} A", "1"], ["B", "2"]]}])
    slides = build(tmp_path, spec, "double", "visual")
    ph = texts(slides[1])["Figure 1\n從論文 PDF 截圖後替換"]
    assert str(ph.fill.fore_color.rgb) == C["concept2_bg"] and str(ph.line.color.rgb) == C["concept2_fg"]
    slides = build(tmp_path, spec, "double", "balanced")
    tbl = next(sh for sh in slides[2].shapes if sh.has_table).table
    assert [str(tbl.cell(1, c).fill.fore_color.rgb) for c in range(2)] == [C["concept1_bg"]] * 2
    assert str(tbl.cell(2, 0).fill.fore_color.rgb) != C["concept1_bg"]
    assert tbl.cell(1, 0).text == "圖像層級 A"


def test_colour_table_lists_used_colours_and_pages(tmp_path):
    spec = concept_spec([
        {"type": "content", "chapter": "甲", "tag": "一、問題", "kind": "problem", "short": ["{{圖像層級}}"]},
        {"type": "mapping", "chapter": "甲", "pairs": [["問題 A", "解法 A"]]},
        {"type": "content", "chapter": "甲", "tag": "x", "short": ["{{圖像層級}} 再次出現"]},
        {"type": "content", "chapter": "甲", "tag": "x", "short": ["無"], "conclusion": "結論"}],
        concepts={**CONCEPTS, "沒用到": 3})
    rows = bd.render(spec, tmp_path, "balanced", "single", str(tmp_path / "o.pptx"))
    by = {r["color"]: r for r in rows}
    assert set(by) == {"紅", "綠", "概念色 1（藍）"}  # unused concept 2/3 not listed
    assert by["概念色 1（藍）"]["pages"] == [2, 4] and by["概念色 1（藍）"]["meaning"] == "圖像層級"
    assert by["概念色 1（藍）"]["hex"] == [C["concept1_bg"], C["concept1_fg"]]
    assert by["紅"]["meaning"] == "問題：問題 A" and by["紅"]["pages"] == [2, 3, 5]  # mapping names win
    md = bd.color_table_md(rows)
    assert "| 概念色 1（藍） | `B6CFF5`／`3C78D8` | 圖像層級 | 第 2、4 頁 |" in md
    assert bd.page_ranges([2, 3, 4, 7]) == "第 2–4、7 頁"


@pytest.mark.parametrize("concepts, needle", [({"a": 5}, "slot 5"), ({"a": 1, "b": 1}, "used twice")])
def test_bad_concept_slots_warn(tmp_path, concepts, needle):
    build(tmp_path, concept_spec([], concepts), "double")
    assert any(needle in w for w in bd.WARN)


def test_undeclared_concept_warns_and_renders_plain(tmp_path):
    slide = build(tmp_path, concept_spec([{"type": "content", "chapter": "甲", "short": ["{{未宣告}}"]}]),
                  "double", "balanced")[1]
    assert any("未宣告" in w for w in bd.WARN)
    assert "未宣告" in [r.text for r in runs_of(slide)]


def test_bullet_keeps_body_colour_before_concept(tmp_path):
    spec = concept_spec([{"type": "content", "chapter": "甲", "short": ["{{圖像層級}}：說明"]}])
    slide = build(tmp_path, spec, "double", "balanced")[1]
    clr = slide._element.find(".//" + bd.qn("a:buClr"))
    assert clr is not None and clr[0].get("val") == C["text"]


def test_concept_frame_around_real_image(tmp_path):
    from PIL import Image
    Image.new("RGB", (400, 300), "white").save(tmp_path / "f.png")
    spec = concept_spec([{"type": "content", "chapter": "甲", "one_line": "y",
                          "figs": [{"id": "Figure 1", "path": "f.png", "concept": "圖像層級"}]}])
    slide = build(tmp_path, spec, "double", "visual")[1]
    (frame,) = named(slide, "概念框/圖像層級")
    pic = next(sh for sh in slide.shapes if sh.shape_type == 13)  # picture
    assert str(frame.line.color.rgb) == C["concept1_fg"]
    assert frame.left < pic.left and frame.left + frame.width > pic.left + pic.width
