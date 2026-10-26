import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

try:
    import fontTools
except ImportError:
    fontTools = None


@unittest.skipUnless(fontTools, 'fonttools required for CI font tests')
class FontCacheTests(unittest.TestCase):
    def test_only_pinned_bytes_are_accepted_and_install_is_offline(self):
        spec = importlib.util.spec_from_file_location('ci_font', Path(__file__).parents[1] / 'install-ci-fonts.py')
        font = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(font)
        import hashlib
        with tempfile.TemporaryDirectory() as temporary:
            cache = Path(temporary) / 'cache'
            cache.mkdir()
            target = Path(temporary) / 'installed'
            with patch.object(font, 'OUTPUT_HASHES', {'test.ttf': hashlib.sha256(b'fixed').hexdigest()}):
                self.assertFalse(font.valid_cache(cache))
                (cache / 'test.ttf').write_bytes(b'fixed')
                self.assertTrue(font.valid_cache(cache))
                with patch.object(font, 'urlopen', side_effect=AssertionError('cache hit must be offline')):
                    font.install_cached(cache, target, False)
                self.assertEqual(b'fixed', (target / 'test.ttf').read_bytes())
                (cache / 'test.ttf').write_bytes(b'corrupt')
                self.assertFalse(font.valid_cache(cache))
