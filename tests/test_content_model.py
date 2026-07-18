"""Tests for the PR1 core content model and stable identity rules."""
from __future__ import annotations

import unittest

from pptx_skill.content_model import (
    BBox,
    CanvasSpec,
    ContentSpec,
    ElementSpec,
    GeometrySpec,
    SafeInsets,
    SlideSpec,
    StableIdGenerator,
    canvas_from_name,
    make_namespace,
)


class CanvasAndGeometryTests(unittest.TestCase):
    def test_default_16_9_canvas(self):
        canvas = canvas_from_name("16:9")
        self.assertAlmostEqual(canvas.width_pt, 959.976, places=3)
        self.assertAlmostEqual(canvas.height_pt, 540.0, places=3)
        self.assertAlmostEqual(canvas.safe_width, 959.976 - 48 - 48, places=3)

    def test_custom_canvas(self):
        canvas = CanvasSpec(width_pt=720, height_pt=540, safe=SafeInsets(36, 36, 36, 36))
        self.assertEqual(canvas.safe_width, 720 - 72)
        self.assertEqual(canvas.safe_height, 540 - 72)

    def test_bbox_intersection(self):
        a = BBox(0, 0, 100, 100)
        b = BBox(50, 50, 100, 100)
        inter = a.intersection(b)
        self.assertIsNotNone(inter)
        self.assertEqual(inter, BBox(50, 50, 50, 50))

    def test_bbox_no_intersection(self):
        a = BBox(0, 0, 10, 10)
        b = BBox(20, 20, 10, 10)
        self.assertIsNone(a.intersection(b))


class StableIdTests(unittest.TestCase):
    def test_namespace_is_deterministic(self):
        self.assertEqual(make_namespace("hello"), make_namespace("hello"))
        self.assertNotEqual(make_namespace("hello"), make_namespace("world"))

    def test_element_ids_are_stable(self):
        gen = StableIdGenerator.from_seed("deck")
        sid = gen.source_section_id(0)
        e1 = gen.element_id(sid, 0, "title", "title")
        e2 = gen.element_id(sid, 0, "title", "title")
        self.assertEqual(e1, e2)

    def test_different_roles_get_different_ids(self):
        gen = StableIdGenerator.from_seed("deck")
        sid = gen.source_section_id(0)
        title = gen.element_id(sid, 0, "title", "title")
        body = gen.element_id(sid, 0, "body", "body")
        self.assertNotEqual(title, body)

    def test_unique_enforcement_for_collisions(self):
        gen = StableIdGenerator.from_seed("deck")
        sid = gen.source_section_id(0)
        # Two distinct items with explicit item keys should not collide.
        a = gen.element_id(sid, 0, "image", "image-0")
        b = gen.element_id(sid, 0, "image", "image-1")
        self.assertNotEqual(a, b)

    def test_slide_id_uniqueness(self):
        gen = StableIdGenerator.from_seed("deck")
        sid = gen.source_section_id(0)
        s1 = gen.slide_id(sid, 0)
        s2 = gen.slide_id(sid, 1)
        self.assertNotEqual(s1, s2)


class ContentModelTests(unittest.TestCase):
    def test_content_spec_round_trip(self):
        slide = SlideSpec(
            id="s1",
            role="bullets",
            communication_goal="intro",
            elements=[
                ElementSpec(
                    id="e1",
                    kind="text",
                    role="title",
                    content={"text": "Title"},
                    style_ref="component.title",
                )
            ],
        )
        content = ContentSpec(
            id="deck",
            title="Deck",
            subtitle="Sub",
            slides=[slide],
            locale="zh-CN",
        )
        self.assertEqual(content.slides[0].elements[0].content["text"], "Title")

    def test_geometry_spec_defaults(self):
        geom = GeometrySpec(bbox=BBox(0, 0, 100, 100))
        self.assertEqual(geom.rotation_deg, 0.0)
        self.assertIsNone(geom.polygon)


if __name__ == "__main__":
    unittest.main()
