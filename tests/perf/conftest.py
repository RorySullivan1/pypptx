"""Pytest fixtures for the performance harness.

Generates synthetic decks once per session under a tmp_path_factory directory.
The fixtures yield `Path` objects so individual benchmarks can choose to open
the deck fresh each iteration (timing cold opens) rather than share a single
loaded `Presentation` instance.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.perf import corpus


@pytest.fixture(scope="session")
def corpus_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return tmp_path_factory.mktemp("perf_corpus")


@pytest.fixture(scope="session")
def synth_small(corpus_dir: Path) -> Path:
    return corpus.build_deck(corpus.SMALL, corpus_dir / "small.pptx")


@pytest.fixture(scope="session")
def synth_medium(corpus_dir: Path) -> Path:
    return corpus.build_deck(corpus.MEDIUM, corpus_dir / "medium.pptx")


@pytest.fixture(scope="session")
def synth_large(corpus_dir: Path) -> Path:
    return corpus.build_deck(corpus.LARGE, corpus_dir / "large.pptx")
