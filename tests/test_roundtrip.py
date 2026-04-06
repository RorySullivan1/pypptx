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

