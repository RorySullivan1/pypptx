"""Sphinx configuration for the pypptx API reference.

Build locally with ``sphinx-build -W -b html docs docs/_build/html`` after
``pip install -e ".[docs]"``. The pages are generated from the in-tree docstrings; there is no
hosted site yet.
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))

import pptx  # noqa: E402

project = "pypptx"
author = "pypptx contributors"
release = pptx.__version__
version = ".".join(release.split(".")[:2])

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx.ext.intersphinx",
    "sphinx_autodoc_typehints",
]

exclude_patterns = ["_build"]
# -- docstrings write XML names as `a:ln`; render them as code, not the default italics --
default_role = "literal"
html_theme = "alabaster"

autodoc_default_options = {
    "members": True,
    "undoc-members": False,
    "show-inheritance": True,
}
autodoc_member_order = "bysource"
autoclass_content = "class"

intersphinx_mapping = {"python": ("https://docs.python.org/3", None)}

# `ShapeElement` is a string alias over element classes imported only under TYPE_CHECKING;
# render it by name instead of trying to evaluate it.
autodoc_type_aliases = {"ShapeElement": "pptx.oxml.shapes.ShapeElement"}


def _substitutions() -> str:
    """Return an `rst_epilog` defining the `|Name|` substitutions the docstrings use.

    The docstrings follow the python-pptx convention of writing `|Chart|` for a link to the
    `Chart` class and `|None|` for the literal. Every class defined in a `pptx` module gets a
    substitution linking to it, so a new class needs no entry here. A `|Name|` that matches
    nothing stays undefined and fails the `-W` build, which is how a typo is caught.
    """
    import importlib
    import inspect
    import pkgutil

    literals = ("True", "False", "None")
    builtins = ("AttributeError", "IndexError", "KeyError", "NotImplementedError", "ValueError")
    builtins += ("TypeError", "float", "int", "str", "bool")
    lines = [".. |%s| replace:: ``%s``" % (name, name) for name in literals]
    lines.append(".. |pp| replace:: pypptx")
    lines += [".. |%s| replace:: :class:`%s`" % (name, name) for name in builtins]

    classes: dict[str, str] = {}
    modules = [pptx.__name__] + [
        info.name for info in pkgutil.walk_packages(pptx.__path__, prefix="pptx.")
    ]
    for module_name in modules:
        module = importlib.import_module(module_name)
        for name, obj in vars(module).items():
            if inspect.isclass(obj) and obj.__module__ == module_name:
                # -- first definition wins; `pptx.chart.data.Category` and
                # -- `pptx.chart.category.Category` are also reachable module-qualified --
                classes.setdefault(name, "%s.%s" % (module_name, name))
                short_module = module_name.rsplit(".", 1)[-1]
                classes.setdefault("%s.%s" % (short_module, name), "%s.%s" % (module_name, name))
    for name, target in sorted(classes.items()):
        lines.append(".. |%s| replace:: :class:`~%s`" % (name, target))

    # -- conceptual names the docstrings use for a family of classes whose shared base is private --
    concepts = {
        "Axis": "pptx.chart.axis._BaseAxis",
        "CoreProperties": "pptx.parts.coreprops.CorePropertiesPart",
        "GradientStops": "pptx.dml.fill._GradientStops",
        "Plot": "pptx.chart.plot._BasePlot",
        "Series": "pptx.chart.series._BaseSeries",
    }
    for name, target in concepts.items():
        lines.append(".. |%s| replace:: :class:`%s <%s>`" % (name, name, target))

    # -- `|Class.member|` links to a member of a class --
    members = {
        "Chart.has_data_table": ":attr:`~pptx.chart.chart.Chart.has_data_table`",
        "ThreadedComments.add": ":meth:`~pptx.slide.ThreadedComments.add`",
    }
    lines += [".. |%s| replace:: %s" % item for item in members.items()]
    return "\n".join(lines) + "\n"


rst_epilog = _substitutions()


def _shield_attribute_summary(app, what, name, obj, options, lines):
    """Stop napoleon reading an attribute's first line as a Google-style `type: description`.

    Property summaries here routinely name an XML element, as in "The `a:ln` element ...", and
    napoleon splits that line at the colon inside the single backticks, turning the start of the
    sentence into a bogus `:type:`. Napoleon does not split inside a double-backtick literal, and
    with `default_role = "literal"` both spellings render the same, so the first line's
    single-backtick spans are rewritten as double-backtick ones.
    """
    if what in ("attribute", "property", "data") and lines and ":" in lines[0]:
        lines[0] = re.sub(r"(?<!`)`([^`]+)`(?!`)", r"``\1``", lines[0])


def setup(app):
    # -- napoleon's handler runs at the default priority (500); this must run first --
    app.connect("autodoc-process-docstring", _shield_attribute_summary, priority=400)
