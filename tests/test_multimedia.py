"""Tests for the multimedia module (video/audio embedding + playback)."""
from __future__ import annotations

import struct
import tempfile
import unittest
import wave
from pathlib import Path

from pptx import Presentation

from pptx_skill.multimedia import (
    add_audio,
    add_video,
    extract_audio,
    extract_video,
    list_media,
    remove_media,
    set_audio_playback,
    set_video_playback,
)


def _make_blank_pptx(path: str | Path) -> str:
    prs = Presentation()
    layout = prs.slide_layouts[6]
    prs.slides.add_slide(layout)
    prs.save(str(path))
    return str(path)


def _make_wav(path: str | Path) -> str:
    with wave.open(str(path), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(8000)
        w.writeframes(struct.pack("<4000h", *([0] * 4000)))
    return str(path)


def _make_fake_mp4(path: str | Path) -> str:
    Path(path).write_bytes(b"\x00" * 1024)
    return str(path)


class MultimediaTests(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.deck = _make_blank_pptx(self.root / "deck.pptx")
        self.wav = _make_wav(self.root / "tone.wav")
        self.mp4 = _make_fake_mp4(self.root / "clip.mp4")

    def tearDown(self):
        self.tmp.cleanup()

    def test_add_video_creates_media_shape(self):
        name = add_video(self.deck, 1, video_path=self.mp4, auto_play=True, volume=70)
        media = list_media(self.deck)
        self.assertEqual(len(media), 1)
        self.assertEqual(media[0]["name"], name)
        self.assertEqual(media[0]["type"], "video")
        self.assertEqual(media[0]["width"], 8.0)
        # shape must be visible to python-pptx as a media-type shape
        prs = Presentation(self.deck)
        self.assertEqual(len(prs.slides), 1)

    def test_add_audio_creates_audio_shape(self):
        name = add_audio(self.deck, 1, audio_path=self.wav, volume=60, loop=True)
        media = list_media(self.deck)
        self.assertEqual(len(media), 1)
        self.assertEqual(media[0]["name"], name)
        self.assertEqual(media[0]["type"], "audio")
        # the WAV must be embedded in the package
        import zipfile
        with zipfile.ZipFile(self.deck) as zf:
            media_names = [n for n in zf.namelist() if n.startswith("ppt/media/")]
        self.assertTrue(any(n.endswith(".wav") for n in media_names))

    def test_set_playback_updates_settings(self):
        aname = add_audio(self.deck, 1, audio_path=self.wav, volume=60)
        self.assertTrue(set_audio_playback(self.deck, 1, aname, volume=80, loop=True))
        vname = add_video(self.deck, 1, video_path=self.mp4)
        self.assertTrue(set_video_playback(self.deck, 1, vname, loop=True, volume=50))
        self.assertEqual(len(list_media(self.deck)), 2)

    def test_extract_media_roundtrip(self):
        aname = add_audio(self.deck, 1, audio_path=self.wav)
        vname = add_video(self.deck, 1, video_path=self.mp4)
        out_wav = self.root / "out.wav"
        out_mp4 = self.root / "out.mp4"
        self.assertTrue(extract_audio(self.deck, 1, aname, str(out_wav)))
        self.assertTrue(extract_video(self.deck, 1, vname, str(out_mp4)))
        self.assertEqual(out_wav.read_bytes(), Path(self.wav).read_bytes())
        self.assertEqual(out_mp4.read_bytes(), Path(self.mp4).read_bytes())

    def test_remove_media(self):
        aname = add_audio(self.deck, 1, audio_path=self.wav)
        self.assertTrue(remove_media(self.deck, 1, aname))
        self.assertEqual(list_media(self.deck), [])
        # removing a missing shape returns False
        self.assertFalse(remove_media(self.deck, 1, "Nope"))

    def test_list_media_slide_scoping(self):
        add_audio(self.deck, 1, audio_path=self.wav)
        self.assertEqual(len(list_media(self.deck)), 1)
        self.assertEqual(len(list_media(self.deck, slide_index=1)), 1)
        with self.assertRaises(IndexError):
            list_media(self.deck, slide_index=2)

    def test_add_media_on_presentation_object(self):
        prs = Presentation()
        layout = prs.slide_layouts[6]
        prs.slides.add_slide(layout)
        name = add_audio(prs, 1, audio_path=self.wav)
        media = list_media(prs)
        self.assertEqual(len(media), 1)
        self.assertEqual(media[0]["name"], name)

    def test_slide_index_out_of_range(self):
        with self.assertRaises(IndexError):
            add_audio(self.deck, 99, audio_path=self.wav)


if __name__ == "__main__":
    unittest.main()
