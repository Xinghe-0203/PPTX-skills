"""Tests for the v6.0 ``pptx_skill.cli`` module.

Covers ``build_parser`` and ``main`` exit codes for the info, capability,
template, generate, from-markdown and edit subcommands.
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import tempfile
import unittest
from argparse import ArgumentParser
from unittest.mock import patch

from pptx_skill.cli import _ensure_utf8, build_parser, main


def _run_cli(argv: list[str]) -> int:
    """Run ``main`` with stdout/stderr captured; return the exit code."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
        return main(argv)


class CliParserTests(unittest.TestCase):
    # 1. build_parser returns ArgumentParser ------------------------------
    def test_build_parser_returns_argument_parser(self):
        parser = build_parser()
        self.assertIsInstance(parser, ArgumentParser)

    # 2. No-args invocation prints help and returns 0 ---------------------
    def test_no_args_returns_zero(self):
        rc = _run_cli([])
        self.assertEqual(rc, 0)

    def test_utf8_setup_reconfigures_existing_streams(self):
        with patch("pptx_skill.cli.configure_utf8_console") as configure:
            _ensure_utf8()
        configure.assert_called_once_with()


class CliSubcommandTests(unittest.TestCase):
    # 3. capability -------------------------------------------------------
    def test_capability_returns_zero(self):
        rc = _run_cli(["capability"])
        self.assertEqual(rc, 0)

    # 4. template ---------------------------------------------------------
    def test_template_returns_zero(self):
        rc = _run_cli(["template"])
        self.assertEqual(rc, 0)

    # 5. info on a non-existent file -------------------------------------
    def test_info_nonexistent_file_returns_one(self):
        rc = _run_cli(["info", "definitely_nonexistent_file.pptx"])
        self.assertEqual(rc, 1)

    # 6. generate ---------------------------------------------------------
    def test_generate_from_sections_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            sections = [{"title": "第一节", "bullets": ["要点 A", "要点 B"]}]
            sections_path = os.path.join(tmp, "sections.json")
            with open(sections_path, "w", encoding="utf-8") as f:
                json.dump(sections, f, ensure_ascii=False)
            out_path = os.path.join(tmp, "out.pptx")
            rc = _run_cli([
                "generate",
                "--title", "测试",
                "--sections", sections_path,
                "--output", out_path,
            ])
            self.assertEqual(rc, 0)
            self.assertTrue(os.path.exists(out_path))

    # 7. from-markdown ----------------------------------------------------
    def test_from_markdown_subcommand(self):
        with tempfile.TemporaryDirectory() as tmp:
            md_path = os.path.join(tmp, "input.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write("# 标题\n## 节\n- 要点\n")
            out_path = os.path.join(tmp, "out.pptx")
            rc = _run_cli(["from-markdown", md_path, "--output", out_path])
            self.assertEqual(rc, 0)
            self.assertTrue(os.path.exists(out_path))

    # 8. edit -------------------------------------------------------------
    def test_edit_find_replace(self):
        from pptx_skill import auto_generate_ppt

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "deck.pptx")
            auto_generate_ppt(
                title="Old Title",
                sections=[{"title": "Old Section", "bullets": ["Old content"]}],
                output_path=path,
                auto_search_images=False,
            )
            rc = _run_cli(["edit", path, "--find-replace", "Old:New"])
            self.assertEqual(rc, 0)


if __name__ == "__main__":
    unittest.main()
