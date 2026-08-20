import importlib
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch

from PIL import Image

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

pptx_skill = importlib.import_module("pptx_skill")
Section = pptx_skill.Section
pixabay_search = importlib.import_module("pixabay_search")
pptx_helper = importlib.import_module("pptx_helper")


class _Response:
    def __init__(self, payload=b"image-bytes"):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return False

    def read(self):
        return self.payload


class PixabayManifestTests(unittest.TestCase):
    def test_download_writes_and_merges_credential_free_asset_manifest(self):
        hit = {
            "id": 42,
            "tags": "steel, factory",
            "largeImageURL": "https://cdn.example.test/42.jpg",
            "webformatURL": "https://cdn.example.test/42-small.jpg",
            "imageWidth": 2400,
            "imageHeight": 1600,
            "user": "Example Author",
            "user_id": 7,
            "pageURL": "https://pixabay.com/photos/example-42/",
        }
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(pixabay_search.urllib.request, "urlopen", return_value=_Response()):
                paths = pixabay_search.download_images(
                    [hit], output_dir=tmp, size="large", query="steel mill"
                )
            self.assertEqual(len(paths), 1)

            # The second query reuses the cached file but must still be merged
            # into the attribution trail.
            with patch.object(pixabay_search.urllib.request, "urlopen") as urlopen:
                pixabay_search.download_images(
                    [hit], output_dir=tmp, size="large", query="Anshan industry"
                )
                urlopen.assert_not_called()

            manifest = json.loads(
                (Path(tmp) / pixabay_search.ASSET_MANIFEST_FILENAME).read_text(encoding="utf-8")
            )
            self.assertEqual(manifest["provider"], "Pixabay")
            self.assertEqual(manifest["license_url"], pixabay_search.PIXABAY_LICENSE_URL)
            self.assertEqual(len(manifest["assets"]), 1)
            asset = manifest["assets"][0]
            self.assertEqual(asset["queries"], ["steel mill", "Anshan industry"])
            self.assertEqual(asset["creator"], "Example Author")
            self.assertEqual(asset["page_url"], hit["pageURL"])
            self.assertNotIn("api_key", json.dumps(manifest))


class AutoSearchRoutingTests(unittest.TestCase):
    def test_non_image_layouts_do_not_consume_search_quota(self):
        sections = [
            Section(title="指标", layout="dashboard", metrics=[{"value": "1"}]),
            Section(title="路线", layout="timeline", events=[{"date": "D1", "title": "到达"}]),
            Section(title="故事", layout="text_image", image_query="city story"),
        ]
        calls = []

        def fake_search(**kwargs):
            calls.append(kwargs["query"])
            return ["candidate.jpg"]

        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(pptx_helper, "_load_pixabay_search", return_value=fake_search):
                pptx_helper._auto_search_images(sections, tmp, "zh")

        self.assertEqual(calls, ["city story"])
        self.assertEqual(sections[0].images, [])
        self.assertEqual(sections[1].images, [])
        self.assertEqual(sections[2].images, ["candidate.jpg"])

    def test_conflicting_image_query_is_skipped_with_warning(self):
        section = Section(
            title="指标",
            layout="dashboard",
            metrics=[{"value": "1"}],
            image_query="unused photo",
        )
        stderr = io.StringIO()
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(pptx_helper, "_load_pixabay_search", return_value=lambda **kwargs: []):
                with redirect_stderr(stderr):
                    pptx_helper._auto_search_images([section], tmp, "zh")
        self.assertIn("不消费图片", stderr.getvalue())

    def test_full_image_warning_recognizes_images_as_consumed(self):
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            pptx_helper._warn_unconsumed("full_image", {"images": ["candidate.jpg"]})
        self.assertEqual(stderr.getvalue(), "")

    def test_explicit_cover_image_skips_redundant_cover_search(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "cover.jpg"
            Image.new("RGB", (1600, 900), "navy").save(image_path)
            output_path = Path(tmp) / "deck.pptx"
            with patch.object(pptx_helper, "_auto_search_images"):
                with patch.object(pptx_helper, "_search_cover_image") as cover_search:
                    pptx_helper.auto_generate_ppt(
                        title="显式封面",
                        sections=[
                            {"title": "显式封面", "layout": "cover", "images": [str(image_path)]},
                            {"title": "内容", "layout": "bullets", "bullets": ["一"]},
                        ],
                        output_path=str(output_path),
                        auto_search_images=True,
                    )
            cover_search.assert_not_called()
            self.assertTrue(output_path.exists())


if __name__ == "__main__":
    unittest.main()
