"""Guards the BDD test-name convention, which pytest's collection depends on."""

from __future__ import annotations

import ast
import re
from pathlib import Path

TESTS_DIR = Path(__file__).parent


class DescribeTestNames:
    """pytest collects `it_*`, `and_it_*` and `but_it_*` methods (pyproject.toml)."""

    def it_finds_no_and_or_but_test_that_pytest_would_skip(self):
        """A method named e.g. `and_its_*` or `but_a_*` reads like a test but is never run."""
        skipped = []
        for path in sorted(TESTS_DIR.rglob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for cls in ast.walk(tree):
                if not (isinstance(cls, ast.ClassDef) and cls.name.startswith("Describe")):
                    continue
                for fn in cls.body:
                    if (
                        isinstance(fn, ast.FunctionDef)
                        and re.match(r"(and|but)_", fn.name)
                        and not re.match(r"(and|but)_it_", fn.name)
                    ):
                        skipped.append(f"{path.relative_to(TESTS_DIR)}:{fn.lineno} {fn.name}")

        assert skipped == []
