# pyright: reportPrivateUsage=false

"""Round-trip integration tests for python-pptx.

These tests create presentations, save to BytesIO, reload, and verify
that content survives the round-trip through serialization/deserialization.
"""

from __future__ import annotations

import subprocess
import sys
from io import BytesIO
from textwrap import dedent

import pytest


def _run_roundtrip_test(test_code: str) -> None:
    """Run a round-trip test in a subprocess to avoid mock pollution from other tests."""
    preamble = "from io import BytesIO\nfrom pptx import Presentation\nfrom pptx.util import Inches, Pt\n"
    script = preamble + dedent(test_code)
    result = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True)
    if result.returncode != 0:
        pytest.fail(f"Round-trip test failed:\\n{result.stderr}")


class DescribeRoundTrip:
    """Integration tests verifying save/load round-trip fidelity."""

    def it_round_trips_a_blank_presentation(self):
        _run_roundtrip_test("""\
prs = Presentation()
stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)
assert len(prs2.slides) == 0
assert len(prs2.slide_layouts) > 0
""")

    def it_round_trips_a_slide_with_shapes(self):
        _run_roundtrip_test("""\
prs = Presentation()
layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(layout)
txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
txBox.text_frame.text = "Hello, World!"

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

assert len(prs2.slides) == 1
slide2 = prs2.slides[0]
textboxes = [s for s in slide2.shapes if s.has_text_frame and not s.is_placeholder]
assert len(textboxes) == 1
assert textboxes[0].text_frame.text == "Hello, World!"
""")

    def it_round_trips_shape_accessibility_properties(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
shape.alternative_text = "A test textbox"
shape.title = "Test Shape"
shape.decorative = True

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

shape2 = [s for s in prs2.slides[0].shapes if not s.is_placeholder][0]
assert shape2.alternative_text == "A test textbox"
assert shape2.title == "Test Shape"
assert shape2.decorative is True
""")

    def it_round_trips_font_properties(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
tf = shape.text_frame
tf.text = "Styled text"
run = tf.paragraphs[0].runs[0]
run.font.bold = True
run.font.italic = True
run.font.size = Pt(24)
run.font.name = "Arial"

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

shape2 = [s for s in prs2.slides[0].shapes if not s.is_placeholder][0]
run2 = shape2.text_frame.paragraphs[0].runs[0]
assert run2.font.bold is True
assert run2.font.italic is True
assert run2.font.size == Pt(24)
assert run2.font.name == "Arial"
""")

    def it_round_trips_slide_operations(self):
        _run_roundtrip_test("""\
prs = Presentation()
layout = prs.slide_layouts[5]

slide1 = prs.slides.add_slide(layout)
slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text_frame.text = "Slide 1"
slide2 = prs.slides.add_slide(layout)
slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text_frame.text = "Slide 2"
slide3 = prs.slides.add_slide(layout)
slide3.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text_frame.text = "Slide 3"

prs.slides.delete(slide2)
assert len(prs.slides) == 2

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

assert len(prs2.slides) == 2
texts = []
for s in prs2.slides:
    for shape in s.shapes:
        if shape.has_text_frame and not shape.is_placeholder:
            texts.append(shape.text_frame.text)
assert texts == ["Slide 1", "Slide 3"]
""")

    def it_round_trips_a_duplicated_slide(self):
        _run_roundtrip_test("""\
prs = Presentation()
layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(layout)
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
shape.text_frame.text = "Original"

new_slide = prs.slides.duplicate(slide)
assert len(prs.slides) == 2

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

assert len(prs2.slides) == 2
for s in prs2.slides:
    textboxes = [sh for sh in s.shapes if sh.has_text_frame and not sh.is_placeholder]
    assert len(textboxes) == 1
    assert textboxes[0].text_frame.text == "Original"
""")

    def it_gives_a_duplicated_slide_its_own_chart(self):
        _run_roundtrip_test("""\
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
chart_data = CategoryChartData()
chart_data.categories = ["a", "b"]
chart_data.add_series("S1", (1, 2))
slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1), Inches(4), Inches(3), chart_data
)
dup = prs.slides.duplicate(slide)
assert slide.shapes[0].chart.part is not dup.shapes[0].chart.part

new_data = CategoryChartData()
new_data.categories = ["a", "b"]
new_data.add_series("S1", (7, 8))
dup.shapes[0].chart.replace_data(new_data)

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

values = [tuple(s.shapes[0].chart.plots[0].series[0].values) for s in prs2.slides]
assert values == [(1.0, 2.0), (7.0, 8.0)], values
partnames = {s.shapes[0].chart.part.partname for s in prs2.slides}
assert len(partnames) == 2, partnames
""")

    def it_keeps_the_workbook_of_an_imported_chart(self):
        _run_roundtrip_test("""\
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

src = Presentation()
src_slide = src.slides.add_slide(src.slide_layouts[6])
for top in (Inches(0), Inches(3)):
    chart_data = CategoryChartData()
    chart_data.categories = ["a", "b"]
    chart_data.add_series("S1", (1, 2) if top == 0 else (3, 4))
    src_slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), top, Inches(4), Inches(2), chart_data
    )

dst = Presentation()
dst.slides.import_slide(src.slides[0])

stream = BytesIO()
dst.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

charts = [shape.chart for shape in prs2.slides[0].shapes]
assert [tuple(c.plots[0].series[0].values) for c in charts] == [(1.0, 2.0), (3.0, 4.0)]
workbooks = [c.part.chart_workbook.xlsx_part for c in charts]
assert all(wb is not None for wb in workbooks)
assert len({wb.partname for wb in workbooks}) == 2
assert len({c.part.partname for c in charts}) == 2
""")

    def it_round_trips_hidden_shapes(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
shape.hidden = True

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

shape2 = [s for s in prs2.slides[0].shapes if not s.is_placeholder][0]
assert shape2.hidden is True
""")

    def it_loads_modern_comment_parts_typed_and_saves_them_untouched(self):
        _run_roundtrip_test("""\
import zipfile
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.comments import AuthorsPart, ModernCommentsPart
from tests.unitutil.modern_comments import AUTHORS_XML, COMMENTS_PARTNAME, COMMENTS_XML
from tests.unitutil.modern_comments import modern_comments_pptx

def canonical(xml):
    from lxml import etree
    return etree.tostring(etree.fromstring(xml.encode() if isinstance(xml, str) else xml),
                          method="c14n")

prs = Presentation(modern_comments_pptx())
comments_part = prs.slides[0].part.part_related_by(RT.MODERN_COMMENTS)
authors_part = prs.part.part_related_by(RT.AUTHORS)
assert isinstance(comments_part, ModernCommentsPart)
assert isinstance(authors_part, AuthorsPart)
assert len(comments_part) == 2
assert len(authors_part) == 2

stream = BytesIO()
prs.save(stream)
zf = zipfile.ZipFile(stream)
xml_decl, body = COMMENTS_XML.split("\\n", 1)
assert canonical(zf.read(COMMENTS_PARTNAME)) == canonical(body)
assert canonical(zf.read("ppt/authors.xml")) == canonical(AUTHORS_XML.split("\\n", 1)[1])
slide_xml = zf.read("ppt/slides/slide1.xml").decode()
assert "p188:commentRel" in slide_xml
""")

    def it_reads_modern_comment_threads_with_their_replies(self):
        _run_roundtrip_test("""\
from tests.unitutil.modern_comments import modern_comments_pptx

slide = Presentation(modern_comments_pptx()).slides[0]
threads = slide.threaded_comments

assert len(threads) == 2
first, second = threads
assert (first.author, first.text, first.anchor, first.resolved) == (
    "Alice Adams", "Tighten this title", "slide", False
)
assert [(r.author, r.text) for r in first.replies] == [
    ("Bob Brown", "First reply"), ("Alice Adams", "Second reply")
]
assert (second.author, second.anchor, second.resolved, second.replies) == (
    "Bob Brown", "shape", True, []
)
assert threads[1].id == second.id
# -- legacy comments are a separate collection and stay empty --
assert len(slide.comments) == 0
""")

    def it_round_trips_added_modern_comment_threads(self):
        _run_roundtrip_test("""\
import re
import zipfile

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
thread = slide.threaded_comments.add("Tighten this title", "Alice Adams", "AA")
thread.reply("Agreed", "Bob Brown", "BB")
thread.reply("Done", "Alice Adams")
other = slide.threaded_comments.add("Check the date", "Bob Brown")
other.resolved = True

stream = BytesIO()
prs.save(stream)

# -- package shape PowerPoint expects --
zf = zipfile.ZipFile(stream)
names = zf.namelist()
creation_id = int(re.search(
    r'<p14:creationId [^>]*val="(\\d+)"', zf.read("ppt/slides/slide1.xml").decode()
).group(1))
partname = "ppt/comments/modernComment_%X_%X.xml" % (slide.slide_id, creation_id)
assert partname in names, names
assert "ppt/authors.xml" in names
content_types = zf.read("[Content_Types].xml").decode()
assert 'PartName="/%s" ContentType="application/vnd.ms-powerpoint.comments+xml"' % partname in (
    content_types
)
assert 'PartName="/ppt/authors.xml" ContentType="application/vnd.ms-powerpoint.authors+xml"' in (
    content_types
)
assert "relationships/authors" in zf.read("ppt/_rels/presentation.xml.rels").decode()
assert "relationships/comments" in zf.read("ppt/slides/_rels/slide1.xml.rels").decode()
slide_xml = zf.read("ppt/slides/slide1.xml").decode()
assert "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}" in slide_xml
assert 'sldId="%d"' % slide.slide_id in zf.read(partname).decode()

# -- and it reads back --
stream.seek(0)
threads = Presentation(stream).slides[0].threaded_comments
assert len(threads) == 2
first, second = threads
assert (first.author, first.author_initials, first.text, first.anchor, first.resolved) == (
    "Alice Adams", "AA", "Tighten this title", "slide", False
)
assert [(r.author, r.text) for r in first.replies] == [
    ("Bob Brown", "Agreed"), ("Alice Adams", "Done")
]
assert (second.author, second.text, second.resolved) == ("Bob Brown", "Check the date", True)
assert first.created is not None and first.created.tzinfo is None
""")

    def it_adds_to_the_modern_comments_already_in_a_file(self):
        _run_roundtrip_test("""\
from tests.unitutil.modern_comments import COMMENTS_PARTNAME, modern_comments_pptx

prs = Presentation(modern_comments_pptx())
slide = prs.slides[0]
slide.threaded_comments.add("Third thread", "Alice Adams")
slide.threaded_comments[0].reply("Third reply", "Carol", "CC")
slide.threaded_comments[1].resolved = False

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)
threads = prs2.slides[0].threaded_comments
assert [t.text for t in threads] == ["Tighten this title", "Logo is off-brand", "Third thread"]
assert threads[2].author == "Alice Adams"
assert [r.author for r in threads[0].replies] == ["Bob Brown", "Alice Adams", "Carol"]
assert threads[1].resolved is False
assert str(prs2.slides[0].part.modern_comments_part.partname) == "/" + COMMENTS_PARTNAME
assert len(prs2.part.authors_part) == 3
""")

    def it_leaves_modern_comments_behind_when_duplicating_a_slide(self):
        _run_roundtrip_test("""\
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from tests.unitutil.modern_comments import modern_comments_pptx

prs = Presentation(modern_comments_pptx())
dup = prs.slides.duplicate(prs.slides[0])
assert RT.MODERN_COMMENTS not in [r.reltype for r in dup.part.rels.values()]
assert dup.part._element.comment_rel_rId is None
assert prs.slides[0].part._element.comment_rel_rId is not None
# -- the copy is a new slide, so modern comments added to it must not name the original --
original_cid = prs.slides[0].part._element.cSld.creation_id
assert dup.part._element.cSld.creation_id not in (None, original_cid)
dup.threaded_comments.add("On the copy", "Alice Adams")
assert dup.part.modern_comments_part.partname != prs.slides[0].part.modern_comments_part.partname

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)
assert [t.text for t in prs2.slides[1].threaded_comments] == ["On the copy"]
assert len(prs2.slides[0].threaded_comments) == 2
""")

    def it_leaves_modern_comments_behind_when_importing_a_slide(self):
        _run_roundtrip_test("""\
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from tests.unitutil.modern_comments import modern_comments_pptx

src = Presentation(modern_comments_pptx())
dst = Presentation()
imported = dst.slides.import_slide(src.slides[0])
assert RT.MODERN_COMMENTS not in [r.reltype for r in imported.part.rels.values()]
assert imported.part._element.comment_rel_rId is None
assert imported.part._element.cSld.creation_id not in (
    None, src.slides[0].part._element.cSld.creation_id
)

stream = BytesIO()
dst.save(stream)
stream.seek(0)
assert len(Presentation(stream).slides) == 1
""")

    def it_round_trips_a_duplicated_shape(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
shape.text_frame.text = "Original"

new_shape = slide.shapes.duplicate_shape(shape)
assert len([s for s in slide.shapes if not s.is_placeholder]) == 2
assert new_shape.text_frame.text == "Original"
assert new_shape.shape_id != shape.shape_id

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

slide2 = prs2.slides[0]
textboxes = [s for s in slide2.shapes if s.has_text_frame and not s.is_placeholder]
assert len(textboxes) == 2
assert all(t.text_frame.text == "Original" for t in textboxes)
# shape IDs should be unique
ids = [t.shape_id for t in textboxes]
assert len(set(ids)) == 2
""")

    def it_round_trips_slide_numbers(self):
        _run_roundtrip_test("""\
prs = Presentation()
prs.first_slide_number = 5

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

assert prs2.first_slide_number == 5
""")

    def it_round_trips_transparency_color(self):
        _run_roundtrip_test("""\
from pptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

import os, tempfile
# create a tiny valid PNG (1x1 white pixel)
png_bytes = (
    b'\\x89PNG\\r\\n\\x1a\\n\\x00\\x00\\x00\\rIHDR\\x00\\x00\\x00\\x01'
    b'\\x00\\x00\\x00\\x01\\x08\\x02\\x00\\x00\\x00\\x90wS\\xde\\x00'
    b'\\x00\\x00\\x0cIDATx\\x9cc\\xf8\\x0f\\x00\\x00\\x01\\x01\\x00'
    b'\\x05\\x18\\xd8N\\x00\\x00\\x00\\x00IEND\\xaeB`\\x82'
)
with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
    f.write(png_bytes)
    img_path = f.name

try:
    pic = slide.shapes.add_picture(img_path, Inches(1), Inches(1))
    pic.transparency_color = RGBColor(0xFF, 0xFF, 0xFF)
    assert pic.transparency_color == RGBColor(0xFF, 0xFF, 0xFF)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    prs2 = Presentation(stream)

    slide2 = prs2.slides[0]
    pics = [s for s in slide2.shapes if s.shape_type is not None]
    from pptx.shapes.picture import Picture
    pic2 = [s for s in slide2.shapes if isinstance(s, Picture)][0]
    assert pic2.transparency_color == RGBColor(0xFF, 0xFF, 0xFF)

    # can remove it
    pic2.transparency_color = None
    assert pic2.transparency_color is None
finally:
    os.unlink(img_path)
""")

    def it_round_trips_line_cap_and_join_style(self):
        _run_roundtrip_test("""\
from pptx.enum.dml import MSO_LINE_CAP_STYLE, MSO_LINE_JOIN_STYLE

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))

shape.line.cap_style = MSO_LINE_CAP_STYLE.ROUND
shape.line.join_style = MSO_LINE_JOIN_STYLE.MITER
shape.line.miter_limit = 4.0

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

shape2 = prs2.slides[0].shapes[0]
assert shape2.line.cap_style == MSO_LINE_CAP_STYLE.ROUND
assert shape2.line.join_style == MSO_LINE_JOIN_STYLE.MITER
assert shape2.line.miter_limit == 4.0
""")

    def it_round_trips_a_picture_filled_rectangle(self):
        _run_roundtrip_test("""\
import os, tempfile
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL, MSO_GRADIENT_TYPE, MSO_RECT_ALIGNMENT

# tiny valid PNG (1x1 white pixel)
png_bytes = (
    b'\\x89PNG\\r\\n\\x1a\\n\\x00\\x00\\x00\\rIHDR\\x00\\x00\\x00\\x01'
    b'\\x00\\x00\\x00\\x01\\x08\\x02\\x00\\x00\\x00\\x90wS\\xde\\x00'
    b'\\x00\\x00\\x0cIDATx\\x9cc\\xf8\\x0f\\x00\\x00\\x01\\x01\\x00'
    b'\\x05\\x18\\xd8N\\x00\\x00\\x00\\x00IEND\\xaeB`\\x82'
)
with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
    f.write(png_bytes)
    img_path = f.name

try:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    stretched = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(2)
    )
    stretched.fill.picture(img_path)

    tiled = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(4), Inches(1), Inches(2), Inches(2)
    )
    tiled.fill.picture(img_path)
    tiled.fill.tile(sx=0.5, sy=0.5, algn=MSO_RECT_ALIGNMENT.CENTER)

    radial = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1), Inches(4), Inches(2), Inches(2)
    )
    radial.fill.gradient()
    radial.fill.gradient_type = MSO_GRADIENT_TYPE.RADIAL
    radial.fill.gradient_fill_to_rect = (0.2, 0.2, 0.8, 0.8)

    stream = BytesIO()
    prs.save(stream)
    stream.seek(0)
    prs2 = Presentation(stream)

    shapes2 = [s for s in prs2.slides[0].shapes if not s.is_placeholder]
    stretched2, tiled2, radial2 = shapes2

    assert stretched2.fill.type == MSO_FILL.PICTURE
    assert stretched2.fill._fill._blipFill.stretch is not None

    assert tiled2.fill.type == MSO_FILL.PICTURE
    tile_elm = tiled2.fill._fill._blipFill.tile
    assert tile_elm is not None
    assert tile_elm.sx == 0.5
    assert tile_elm.sy == 0.5
    assert tile_elm.algn == MSO_RECT_ALIGNMENT.CENTER

    assert radial2.fill.type == MSO_FILL.GRADIENT
    assert radial2.fill.gradient_type == MSO_GRADIENT_TYPE.RADIAL
    assert radial2.fill.gradient_fill_to_rect == (0.2, 0.2, 0.8, 0.8)
finally:
    os.unlink(img_path)
""")

    def it_round_trips_theme_effect_scheme(self):
        _run_roundtrip_test("""\
from pptx.theme import EffectScheme

prs = Presentation()

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

master = prs2.slide_masters[0]
theme = master.theme
effect_scheme = theme.effect_scheme
assert effect_scheme is not None
assert len(effect_scheme) == 3
assert effect_scheme.subtle.has_effect_list is True
assert effect_scheme.intense.has_3d_scene is True
assert effect_scheme.intense.has_3d_shape is True
assert effect_scheme.name == "Office"
""")

    def it_round_trips_slide_import(self):
        _run_roundtrip_test("""\
# --- create source presentation with a text box ---
src_prs = Presentation()
src_slide = src_prs.slides.add_slide(src_prs.slide_layouts[6])  # blank layout
txBox = src_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
txBox.text_frame.text = "Imported Slide"

# --- save and reload source ---
src_stream = BytesIO()
src_prs.save(src_stream)
src_stream.seek(0)
src_prs2 = Presentation(src_stream)

# --- create target presentation and import slide ---
tgt_prs = Presentation()
imported = tgt_prs.slides.import_slide(src_prs2.slides[0])

assert imported is not None
# target should now have 1 slide (the imported one)
assert len(tgt_prs.slides) == 1

# verify content survived
shapes = list(tgt_prs.slides[0].shapes)
textboxes = [s for s in shapes if s.has_text_frame and not s.is_placeholder]
assert len(textboxes) == 1
assert textboxes[0].text_frame.text == "Imported Slide"

# --- save and reload to verify round-trip ---
tgt_stream = BytesIO()
tgt_prs.save(tgt_stream)
tgt_stream.seek(0)
tgt_prs2 = Presentation(tgt_stream)

assert len(tgt_prs2.slides) == 1
shapes2 = list(tgt_prs2.slides[0].shapes)
textboxes2 = [s for s in shapes2 if s.has_text_frame and not s.is_placeholder]
assert len(textboxes2) == 1
assert textboxes2[0].text_frame.text == "Imported Slide"
""")

    def it_round_trips_merge_presentations(self):
        _run_roundtrip_test("""\
# --- create two source presentations ---
prs_a = Presentation()
slide_a = prs_a.slides.add_slide(prs_a.slide_layouts[6])
txBox_a = slide_a.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
txBox_a.text_frame.text = "Slide A"

prs_b = Presentation()
slide_b1 = prs_b.slides.add_slide(prs_b.slide_layouts[6])
txBox_b1 = slide_b1.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
txBox_b1.text_frame.text = "Slide B1"
slide_b2 = prs_b.slides.add_slide(prs_b.slide_layouts[6])
txBox_b2 = slide_b2.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
txBox_b2.text_frame.text = "Slide B2"

# --- save and reload ---
stream_a = BytesIO()
prs_a.save(stream_a)
stream_a.seek(0)
prs_a2 = Presentation(stream_a)

stream_b = BytesIO()
prs_b.save(stream_b)
stream_b.seek(0)
prs_b2 = Presentation(stream_b)

# --- merge B into A ---
new_slides = prs_a2.slides.merge(prs_b2)
assert len(new_slides) == 2
assert len(prs_a2.slides) == 3  # 1 original + 2 merged

# --- save and reload ---
out_stream = BytesIO()
prs_a2.save(out_stream)
out_stream.seek(0)
merged = Presentation(out_stream)

assert len(merged.slides) == 3
texts = []
for slide in merged.slides:
    for shape in slide.shapes:
        if shape.has_text_frame and not shape.is_placeholder:
            texts.append(shape.text_frame.text)
assert texts == ["Slide A", "Slide B1", "Slide B2"]
""")

    def it_preserves_potx_content_type_on_round_trip(self):
        _run_roundtrip_test("""\
from pptx.opc.constants import CONTENT_TYPE as CT

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Template Slide"

# Override the main part content type to template
del prs.part.__dict__['content_type']  # clear lazyproperty cache
prs.part._content_type = CT.PML_TEMPLATE_MAIN

stream = BytesIO()
prs.save(stream)
stream.seek(0)

# Reload and verify content type is preserved
prs2 = Presentation(stream)
assert prs2.part.content_type == CT.PML_TEMPLATE_MAIN
assert prs2.slides[0].shapes.title.text == "Template Slide"
""")

    def it_round_trips_shape_range_alignment(self):
        _run_roundtrip_test("""\
from pptx.shapes.range import ShapeRange

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

s1 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
s2 = slide.shapes.add_textbox(Inches(3), Inches(2), Inches(2), Inches(1))
s3 = slide.shapes.add_textbox(Inches(5), Inches(3), Inches(2), Inches(1))

# use shape_range convenience method
sr = slide.shapes.shape_range()
assert len(sr) == 3

# align left and distribute vertical
sr.align_left()
sr.distribute_vertical()

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

slide2 = prs2.slides[0]
shapes = [s for s in slide2.shapes if not s.is_placeholder]
assert len(shapes) == 3

# all shapes should have the same left edge after align_left
lefts = [s.left for s in shapes]
assert len(set(lefts)) == 1
""")

    def it_imports_slide_with_image(self):
        _run_roundtrip_test("""\
import os, tempfile
from pptx.shapes.picture import Picture

# tiny valid PNG (1x1 white pixel)
png_bytes = (
    b'\\x89PNG\\r\\n\\x1a\\n\\x00\\x00\\x00\\rIHDR\\x00\\x00\\x00\\x01'
    b'\\x00\\x00\\x00\\x01\\x08\\x02\\x00\\x00\\x00\\x90wS\\xde\\x00'
    b'\\x00\\x00\\x0cIDATx\\x9cc\\xf8\\x0f\\x00\\x00\\x01\\x01\\x00'
    b'\\x05\\x18\\xd8N\\x00\\x00\\x00\\x00IEND\\xaeB`\\x82'
)
with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
    f.write(png_bytes)
    img_path = f.name

try:
    # --- source with image ---
    src = Presentation()
    slide = src.slides.add_slide(src.slide_layouts[6])
    slide.shapes.add_picture(img_path, Inches(1), Inches(1))

    src_stream = BytesIO()
    src.save(src_stream)
    src_stream.seek(0)
    src2 = Presentation(src_stream)

    # --- import into target ---
    tgt = Presentation()
    tgt.slides.import_slide(src2.slides[0])

    tgt_stream = BytesIO()
    tgt.save(tgt_stream)
    tgt_stream.seek(0)
    tgt2 = Presentation(tgt_stream)

    assert len(tgt2.slides) == 1
    pics = [s for s in tgt2.slides[0].shapes if isinstance(s, Picture)]
    assert len(pics) == 1
finally:
    os.unlink(img_path)
""")


    def it_round_trips_master_text_styles_and_the_default_text_style(self):
        _run_roundtrip_test("""\
from pptx.enum.text import PP_ALIGN

prs = Presentation()
styles = prs.slide_master.text_styles
styles.body[0].font.size = Pt(20)
styles.body[1].margin_left = Inches(1)
styles.title[0].alignment = PP_ALIGN.LEFT
styles.other[0].font.bold = True
prs.default_text_style[0].font.italic = True
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.placeholders[1].text_frame.text = "inherits body level 1"

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

styles2 = prs2.slide_master.text_styles
assert styles2.body[0].font.size == Pt(20)
assert styles2.body[1].margin_left == Inches(1)
assert styles2.title[0].alignment == PP_ALIGN.LEFT
assert styles2.other[0].font.bold is True
assert prs2.default_text_style[0].font.italic is True
sldMaster = prs2.slide_master._element
assert sldMaster.xpath("p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz") == ["2000"]
# -- the slide sets no size of its own, so the master's level-1 body style applies --
p = prs2.slides[0].placeholders[1].text_frame.paragraphs[0]
assert p.level == 0
assert p.font.size is None and all(r.font.size is None for r in p.runs)
""")

    def it_round_trips_a_theme_applied_from_another_presentation(self):
        _run_roundtrip_test("""\
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

source = Presentation()
source.theme.color_scheme._clrScheme.accent1[0].set("val", "C00000")
source.theme.font_scheme.major_font.latin = "Georgia"

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])
shape = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
shape.fill.solid()
shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
prs.slide_master.apply_theme(source)

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

theme = prs2.slide_master.theme
assert theme.color_scheme.accent_1 == "C00000"
assert theme.font_scheme.major_font.latin == "Georgia"
assert dict(theme.color_scheme) == dict(source.theme.color_scheme)
theme_rels = [r for r in prs2.slide_master.part.rels.values() if r.reltype == RT.THEME]
assert len(theme_rels) == 1
theme_parts = [p for p in prs2.part.package.iter_parts() if p.partname.startswith("/ppt/theme/")]
assert theme_parts == [prs2.slide_master.part.theme_part]

# -- every scheme-color reference on the slide, layouts, and master resolves in the new theme --
clrMap = prs2.slide_master._element.find(
    "{http://schemas.openxmlformats.org/presentationml/2006/main}clrMap"
)
elements = [prs2.slide_master._element, prs2.slides[0]._element]
elements += [layout._element for layout in prs2.slide_layouts]
for elm in elements:
    for val in elm.xpath(".//a:schemeClr/@val"):
        if val == "phClr":
            continue
        slot = clrMap.get(val, val)
        assert theme.color_scheme[slot] is not None, (val, slot)
assert prs2.slides[0].shapes[-1].fill.fore_color.theme_color == MSO_THEME_COLOR.ACCENT_1
""")

    def it_round_trips_a_freeform_shape_with_curve_and_arc_segments(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

builder = slide.shapes.build_freeform(start_x=0, start_y=0, scale=1.0)
builder.add_line_segments([(100, 0)], close=False)
builder.add_cubic_bezier((120, 20), (140, 40), (150, 50))
builder.add_quadratic_bezier((170, 30), (190, 50))
builder.add_arc(x_radius=20, y_radius=10, start_angle=0, swing_angle=90)
builder.add_line_segments([(0, 0)], close=True)
shape = builder.convert_to_shape(origin_x=Inches(1), origin_y=Inches(1))

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

shape2 = [s for s in prs2.slides[0].shapes if not s.is_placeholder][0]
path = shape2._element.spPr.custGeom.pathLst.path_lst[0]

assert len(path.xpath("a:lnTo")) == 2
assert len(path.xpath("a:cubicBezTo")) == 1
assert len(path.xpath("a:cubicBezTo/a:pt")) == 3
assert len(path.xpath("a:quadBezTo")) == 1
assert len(path.xpath("a:quadBezTo/a:pt")) == 2
assert len(path.xpath("a:arcTo")) == 1
arcTo = path.xpath("a:arcTo")[0]
assert arcTo.get("wR") == "20"
assert arcTo.get("hR") == "10"
assert arcTo.get("stAng") == "0"
assert arcTo.get("swAng") == "5400000"
assert len(path.xpath("a:close")) == 1
""")

    def it_round_trips_a_table_after_row_and_column_edits(self):
        _run_roundtrip_test("""\
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
gf = slide.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(6), Inches(3))
table = gf.table

for r in range(3):
    for c in range(3):
        table.cell(r, c).text = f"{r},{c}"

# -- merge a 2x1 vertical range, then insert a row inside it --
table.cell(0, 0).merge(table.cell(1, 0))
table.rows.add(1)

# -- append a column, then remove the second row --
table.columns.add()
table.rows.remove(table.rows[1])

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

table2 = prs2.slides[0].shapes[0].table
assert len(table2.rows) == 3
assert len(table2.columns) == 4
# -- every row has one cell per grid column --
for row in table2.rows:
    assert len(row.cells) == 4
""")

    def it_round_trips_a_theme_applied_from_a_thmx_file(self):
        _run_roundtrip_test("""\
import os
import tempfile
import zipfile

import pptx

template = os.path.join(os.path.dirname(pptx.__file__), "templates", "default.pptx")
with zipfile.ZipFile(template) as z:
    theme_xml = z.read("ppt/theme/theme1.xml").decode("utf-8")
theme_xml = theme_xml.replace(
    '<a:accent2><a:srgbClr val="C0504D"/>', '<a:accent2><a:srgbClr val="00B050"/>'
)
assert "00B050" in theme_xml
thmx = os.path.join(tempfile.mkdtemp(), "green.thmx")
with zipfile.ZipFile(thmx, "w") as z:
    z.writestr(
        "[Content_Types].xml",
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/theme/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    )
    z.writestr(
        "_rels/.rels",
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Target="theme/theme/theme1.xml" Type="http://schemas.'
        'openxmlformats.org/officeDocument/2006/relationships/officeDocument"/></Relationships>',
    )
    z.writestr("theme/theme/theme1.xml", theme_xml)

prs = Presentation()
prs.slides.add_slide(prs.slide_layouts[0])
prs.slide_master.apply_theme(thmx)

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)
assert prs2.slide_master.theme.color_scheme.accent_2 == "00B050"
assert len(prs2.slides) == 1
""")

    def it_round_trips_custom_shows_and_notes_size(self):
        _run_roundtrip_test("""\
prs = Presentation()
layout = prs.slide_layouts[5]
slide1 = prs.slides.add_slide(layout)
slide2 = prs.slides.add_slide(layout)
slide3 = prs.slides.add_slide(layout)

prs.notes_height = Inches(11)
prs.notes_width = Inches(8.5)

show = prs.custom_shows.add("Intro Only", [slide1, slide3])

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

assert prs2.notes_height == Inches(11)
assert prs2.notes_width == Inches(8.5)

assert len(prs2.custom_shows) == 1
custom_show = prs2.custom_shows[0]
assert custom_show.name == "Intro Only"
assert [s.slide_id for s in custom_show.slides] == [
    prs2.slides[0].slide_id,
    prs2.slides[2].slide_id,
]

# -- deleting a referenced slide prunes it from the custom show --
prs2.slides.delete(prs2.slides[2])
stream2 = BytesIO()
prs2.save(stream2)
stream2.seek(0)
prs3 = Presentation(stream2)

assert len(prs3.slides) == 2
custom_show3 = prs3.custom_shows[0]
assert len(custom_show3.slides) == 1
assert custom_show3.slides[0].slide_id == prs3.slides[0].slide_id
""")

    def it_round_trips_embedded_fonts(self):
        _run_roundtrip_test("""\
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

prs = Presentation()
prs_part = prs.part

font_part = Part(
    PackURI("/ppt/fonts/font1.fntdata"), CT.X_FONTDATA, prs_part.package, b"dummy-font-bytes"
)
rId = prs_part.relate_to(font_part, RT.FONT)
entry = prs._element.get_or_add_embeddedFontLst().add_embeddedFont("Calibri")
entry.add_style("regular", rId)

stream = BytesIO()
prs.save(stream)
stream.seek(0)
prs2 = Presentation(stream)

embedded_fonts = prs2.embedded_fonts
assert len(embedded_fonts) == 1
assert embedded_fonts[0].typeface == "Calibri"
assert embedded_fonts[0].styles == ("regular",)

embedded_fonts[0].remove()
assert len(prs2.embedded_fonts) == 0

stream2 = BytesIO()
prs2.save(stream2)
stream2.seek(0)
prs3 = Presentation(stream2)
assert len(prs3.embedded_fonts) == 0
""")
