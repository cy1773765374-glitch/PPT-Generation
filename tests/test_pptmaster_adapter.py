#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import os
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from generate_catalog_ppt import build_deck_plan, prepare_images  # noqa: E402
from openclaw_pptmaster_adapter import write_pptmaster_project  # noqa: E402


def test_write_pptmaster_project_outputs_spec_and_svg():
    prompt = "Generate a 2-page luxury stationery product catalog PPT. Product categories: Notebooks, Pens. All text in English."
    with tempfile.TemporaryDirectory() as td:
        os.environ["PPT_REQUIRE_IMAGES"] = "1"
        plan = build_deck_plan(prompt)
        assets = prepare_images(plan, Path(td), "placeholder")
        project = write_pptmaster_project(plan, assets, Path(td))
        assert (project / "design_spec.md").exists()
        assert (project / "spec_lock.md").exists()
        assert (project / "deck_plan.json").exists()
        svg_files = sorted((project / "svg_output").glob("*.svg"))
        assert len(svg_files) == 2
        assert "viewBox=\"0 0 1280 720\"" in svg_files[0].read_text(encoding="utf-8")
        assert (project / "images" / "slide_01.png").exists()
        assert "../images/slide_01.png" in svg_files[0].read_text(encoding="utf-8")


def test_discover_pptmaster_root_supports_nested_skill_scripts(monkeypatch):
    from openclaw_pptmaster_adapter import discover_pptmaster_root, _script
    with tempfile.TemporaryDirectory() as td:
        root = Path(td) / "vendor" / "ppt-master"
        scripts_dir = root / "skills" / "ppt-master" / "scripts"
        scripts_dir.mkdir(parents=True)
        (scripts_dir / "svg_to_pptx.py").write_text("# dummy", encoding="utf-8")
        monkeypatch.setenv("PPT_MASTER_ROOT", str(root))
        discovered = discover_pptmaster_root(Path(td))
        assert discovered == root.resolve()
        assert _script(discovered, "svg_to_pptx.py") == (scripts_dir / "svg_to_pptx.py").resolve()
