"""Round-trip and fidelity tests over the real-world corpus in `tests/test_files/real_world/`.

The corpus is PowerPoint-authored test data from Apache POI (see `SOURCES.md` there). Every
file must read cleanly through the public API, survive an edit, save and reload, and lose
nothing on a plain open-and-save.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from pptx import Presentation
from pptx.util import Inches

from .unitutil.deckwalk import walk

CORPUS_DIR = Path(__file__).parent / "test_files" / "real_world"
CORPUS = sorted(CORPUS_DIR.glob("*.pptx"), key=lambda p: p.name.lower())
MARKER_TEXT = "pypptx round-trip marker"

_blankless = etree.XMLParser(remove_blank_text=True)


def _canonical(xml: bytes) -> bytes:
    """`xml` in canonical form, ignoring whitespace-only text between elements."""
    return etree.tostring(etree.fromstring(xml, _blankless), method="c14n")


def _content_types(zipf: zipfile.ZipFile) -> dict[str, str | None]:
    """Map each member's partname to the content type `[Content_Types].xml` gives it."""
    types = etree.fromstring(zipf.read("[Content_Types].xml"))
    ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
    defaults = {
        e.get("Extension").lower(): e.get("ContentType") for e in types.xpath("ct:Default", namespaces=ns)
    }
    overrides = {
        e.get("PartName").lower(): e.get("ContentType") for e in types.xpath("ct:Override", namespaces=ns)
    }
    mapping: dict[str, str | None] = {}
    for name in zipf.namelist():
        if name.endswith("/") or name == "[Content_Types].xml":
            continue
        partname = "/" + name
        ext = posixpath.basename(name).rpartition(".")[2].lower()
        mapping[partname] = overrides.get(partname.lower(), defaults.get(ext))
    return mapping


def it_has_a_corpus_to_test():
    assert len(CORPUS) >= 20, "expected the real-world corpus under %s" % CORPUS_DIR


@pytest.mark.realworld
@pytest.mark.parametrize("path", CORPUS, ids=[p.name for p in CORPUS])
class DescribeRealWorldDeck:
    """Each PowerPoint-authored deck in the corpus."""

    def it_reads_through_the_public_api_without_raising(self, path: Path):
        prs = Presentation(str(path))

        failures = walk(prs)

        assert failures == [], "\n".join("%s: %s" % f for f in failures[:20])

    def it_survives_an_edit_a_save_and_a_reload(self, path: Path):
        prs = Presentation(str(path))
        slide_count = len(prs.slides)
        if slide_count:
            textbox = prs.slides[0].shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
            textbox.text_frame.text = MARKER_TEXT
        stream = io.BytesIO()
        prs.save(stream)
        stream.seek(0)

        reloaded = Presentation(stream)

        assert len(reloaded.slides) == slide_count
        if slide_count:
            texts = [s.text_frame.text for s in reloaded.slides[0].shapes if s.has_text_frame]
            assert MARKER_TEXT in texts
        assert walk(reloaded) == []
        with zipfile.ZipFile(path) as original, zipfile.ZipFile(stream) as saved:
            missing = set(original.namelist()) - set(saved.namelist())
        assert missing == set()

    def it_changes_no_xml_on_a_plain_open_and_save(self, path: Path):
        stream = io.BytesIO()
        Presentation(str(path)).save(stream)

        with zipfile.ZipFile(path) as original, zipfile.ZipFile(stream) as saved:
            changed = [
                name
                for name in original.namelist()
                if name.endswith(".xml")
                and name != "[Content_Types].xml"
                and _canonical(original.read(name)) != _canonical(saved.read(name))
            ]
            original_types = _content_types(original)
            saved_types = _content_types(saved)

        assert changed == []
        # -- every part keeps its content type, or gains one the original failed to declare --
        lost = {
            partname: (content_type, saved_types.get(partname))
            for partname, content_type in original_types.items()
            if content_type is not None and saved_types.get(partname) != content_type
        }
        assert lost == {}


def it_keeps_the_corpus_test_only():
    """The corpus is Apache POI test data, licensed for use here as test data only."""
    sources = (CORPUS_DIR / "SOURCES.md").read_text(encoding="utf-8")
    listed = set(re.findall(r"^\| `([^`]+\.pptx)` \| \d+ \| [^|]+ \| [^|]+ \|", sources, re.M))
    assert listed == {p.name for p in CORPUS}
    assert (CORPUS_DIR / "LICENSE").exists() and (CORPUS_DIR / "NOTICE").exists()
