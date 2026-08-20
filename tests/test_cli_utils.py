import sys
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from cli_utils import configure_utf8_console  # noqa: E402


class _ReconfigurableStream:
    def __init__(self):
        self.calls = []

    def reconfigure(self, **kwargs):
        self.calls.append(kwargs)


class ConsoleEncodingTests(unittest.TestCase):
    def test_configures_both_streams_as_utf8(self):
        stdout = _ReconfigurableStream()
        stderr = _ReconfigurableStream()

        configure_utf8_console(stdout, stderr)

        self.assertEqual(stdout.calls, [{"encoding": "utf-8", "errors": "replace"}])
        self.assertEqual(stderr.calls, [{"encoding": "utf-8", "errors": "replace"}])

    def test_wrapped_stream_without_reconfigure_is_supported(self):
        configure_utf8_console(object(), object())


if __name__ == "__main__":
    unittest.main()
