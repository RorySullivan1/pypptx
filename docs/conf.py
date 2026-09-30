"""Sphinx configuration for the pypptx API reference.

Build locally with ``sphinx-build -W -b html docs docs/_build/html`` after
``pip install -e ".[docs]"``. The pages are generated from the in-tree docstrings; there is no
hosted site yet.
"""

from __future__ import annotations

import os
import sys
import warnings

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
    # -- much of the API is defined on private bases (`_BaseGroupShapes.add_shape`,
    # -- `_BaseSeries.name`, `_BaseAxis.has_major_gridlines`); show it on the public
    # -- subclasses, but not what those bases inherit from the standard library --
    "inherited-members": "object,int,str,tuple,list,dict,Enum,IntEnum,Sequence,Mapping",
}
autodoc_member_order = "bysource"
autoclass_content = "class"

intersphinx_mapping = {"python": ("https://docs.python.org/3", None)}

# `ShapeElement` is a string alias over element classes imported only under TYPE_CHECKING;
# render it by name instead of trying to evaluate it.
autodoc_type_aliases = {"ShapeElement": "pptx.oxml.shapes.ShapeElement"}


def _documented_modules() -> list[str]:
    """Names of the modules the `docs/api/*.rst` pages document with `automodule`."""
    api_dir = os.path.join(os.path.dirname(__file__), "api")
    names: list[str] = []
    for filename in sorted(os.listdir(api_dir)):
        with open(os.path.join(api_dir, filename), encoding="utf-8") as f:
            names += [
                line.split("::", 1)[1].strip()
                for line in f
                if line.startswith(".. automodule::")
            ]
    return names


def _substitutions() -> str:
    """Return an `rst_epilog` defining the `|Name|` substitutions the docstrings use.

    The docstrings follow the python-pptx convention of writing `|Chart|` for a link to the
    `Chart` class and `|None|` for the literal. Every class defined in a `pptx` module gets a
    substitution linking to it, so a new class needs no entry here. A `|Name|` that matches
    nothing stays undefined and fails the `-W` build, which is how a typo is caught.

    Each name is defined once. Precedence: the fixed entries below, then classes in the
    documented modules (in page order), then classes anywhere else in `pptx`. Where two
    modules define the same class name, the module-qualified form (`|data.Category|`) reaches
    either one.
    """
    import importlib
    import inspect
    import pkgutil

    subs: dict[str, str] = {name: "``%s``" % name for name in ("True", "False", "None")}
    subs["pp"] = "pypptx"
    builtins = ("AttributeError", "IndexError", "KeyError", "NotImplementedError", "ValueError")
    builtins += ("TypeError", "float", "int", "str", "bool")
    subs.update({name: ":class:`%s`" % name for name in builtins})

    # -- conceptual names the docstrings use for a family of classes whose shared base is private --
    concepts = {
        "Axis": "pptx.chart.axis._BaseAxis",
        "CoreProperties": "pptx.parts.coreprops.CorePropertiesPart",
        "GradientStops": "pptx.dml.fill._GradientStops",
        "GradientStop": "pptx.dml.fill._GradientStop",
        "Plot": "pptx.chart.plot._BasePlot",
        "Series": "pptx.chart.series._BaseSeries",
    }
    subs.update({name: ":class:`%s <%s>`" % (name, target) for name, target in concepts.items()})

    # -- `|Class.member|` links to a member of a class --
    subs["Chart.has_data_table"] = ":attr:`~pptx.chart.chart.Chart.has_data_table`"
    subs["ThreadedComments.add"] = ":meth:`~pptx.slide.ThreadedComments.add`"

    documented = _documented_modules()
    others = [pptx.__name__] + [
        info.name
        for info in pkgutil.walk_packages(pptx.__path__, prefix="pptx.")
        if info.name not in documented
    ]
    for module_name in documented + others:
        try:
            module = importlib.import_module(module_name)
        except Exception as e:  # noqa: BLE001 -- report and carry on; autodoc reports it too
            warnings.warn("docs/conf.py: cannot import %s: %s" % (module_name, e))
            continue
        short_module = module_name.rsplit(".", 1)[-1]
        for name, obj in vars(module).items():
            if inspect.isclass(obj) and obj.__module__ == module_name:
                target = ":class:`~%s.%s`" % (module_name, name)
                subs.setdefault(name, target)
                subs.setdefault("%s.%s" % (short_module, name), target)

    return "\n".join(".. |%s| replace:: %s" % item for item in sorted(subs.items())) + "\n"


rst_epilog = _substitutions()


def _shield_attribute_summary(app, what, name, obj, options, lines):
    """Stop napoleon reading an attribute's first line as a Google-style `type: description`.

    Napoleon splits the first line of a property or attribute docstring at its first colon and
    renders the part before it as the type. The docstrings here never use that form, but their
    summaries often contain a colon -- an XML name such as `a:ln`, a `:ref:` role, or prose like
    "The thread's status: ...". Leading the docstring with an empty RST comment (`..` and a
    blank line) gives napoleon a first line with no colon; the comment renders as nothing.
    """
    if what in ("attribute", "property", "data") and lines and lines[0].strip():
        lines[0:0] = ["..", ""]


def setup(app):
    # -- napoleon's handler runs at the default priority (500); this must run first --
    app.connect("autodoc-process-docstring", _shield_attribute_summary, priority=400)
