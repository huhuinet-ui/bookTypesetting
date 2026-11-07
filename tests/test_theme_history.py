import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import local_server


class ThemeHistoryTests(unittest.TestCase):
    def setUp(self):
        work = Path(__file__).resolve().parents[1] / 'tmp' / 'theme-history-tests'
        work.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=str(work))
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'colors.json'
        override = patch.object(local_server, 'THEME_HISTORY_FILE', self.path)
        override.start()
        self.addCleanup(override.stop)

    def test_colors_are_validated_normalized_and_deduplicated(self):
        self.assertEqual(local_server.normalize_theme_colors(
            ['#a1dbed', ' #A1DBED ', '#336699', None, 12, '#fff', 'red', '<script>']
        ), ['#A1DBED', '#336699'])
        self.assertEqual(local_server.normalize_theme_colors({'color': '#336699'}), [])

    def test_recent_color_moves_to_front_and_survives_reload(self):
        local_server.remember_theme_color('#a1dbed')
        local_server.remember_theme_color('#336699')
        local_server.remember_theme_color('#A1DBED')
        self.assertEqual(local_server.read_theme_colors(), ['#A1DBED', '#336699'])
        self.assertEqual(json.loads(self.path.read_text()), ['#A1DBED', '#336699'])
        self.assertFalse(self.path.with_suffix('.json.tmp').exists())

    def test_only_24_recent_colors_are_kept(self):
        for number in range(30):
            local_server.remember_theme_color('#{:06X}'.format(number))
        self.assertEqual(local_server.read_theme_colors(),
                         ['#{:06X}'.format(n) for n in range(29, 5, -1)])

    def test_missing_and_corrupt_history_recover(self):
        self.assertEqual(local_server.read_theme_colors(), [])
        self.path.write_text('not JSON', encoding='utf-8')
        self.assertEqual(local_server.read_theme_colors(), [])
        self.assertEqual(local_server.remember_theme_color('#DEB887'), ['#DEB887'])

    def test_invalid_color_does_not_overwrite_history(self):
        local_server.remember_theme_color('#336699')
        with self.assertRaisesRegex(ValueError, '无效的主题色'):
            local_server.remember_theme_color('not-a-color')
        self.assertEqual(local_server.read_theme_colors(), ['#336699'])


if __name__ == '__main__':
    unittest.main()
