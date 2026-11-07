from pathlib import Path
import tempfile
import unittest

from local_server import safe_path
from word_worker import validate_image_path


class ExportPathTests(unittest.TestCase):
    def setUp(self):
        work = Path(__file__).resolve().parents[1] / "tmp" / "export-path-tests"
        work.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(dir=str(work))
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        self.root = self.directory / "uploads"
        self.root.mkdir()
        self.image = self.root / "图" / "image.png"
        self.image.parent.mkdir()
        self.image.write_bytes(b"image path validation fixture")
        self.outside = self.directory / "uploads-other" / "image.png"
        self.outside.parent.mkdir()
        self.outside.write_bytes(b"outside import directory")

    def test_upload_path_with_windows_separators(self):
        self.assertEqual(safe_path(self.root, "图\\image.png"), self.image.resolve())
        self.assertEqual(safe_path(self.root, "./图/../图/image.png"), self.image.resolve())

    def test_upload_paths_reject_escape_and_root(self):
        for name in (".", "图/..", "../uploads-other/image.png", str(self.outside)):
            with self.subTest(name=name):
                with self.assertRaisesRegex(ValueError, "文件路径超出导入目录"):
                    safe_path(self.root, name)

    def test_image_within_import_directory(self):
        self.assertEqual(validate_image_path(self.image, self.root), self.image.resolve())

    def test_image_rejects_common_prefix_sibling(self):
        with self.assertRaisesRegex(ValueError, "图片必须位于导入文件夹内"):
            validate_image_path(self.outside, self.root)

    def test_image_rejects_parent_traversal(self):
        with self.assertRaisesRegex(ValueError, "图片必须位于导入文件夹内"):
            validate_image_path(self.root / ".." / "uploads-other" / "image.png", self.root)

    def test_missing_image_retains_clear_error(self):
        with self.assertRaisesRegex(ValueError, "找不到图片"):
            validate_image_path(self.root / "missing.png", self.root)
        with self.assertRaisesRegex(ValueError, "找不到图片"):
            validate_image_path(self.root, self.root)


if __name__ == "__main__":
    unittest.main()
