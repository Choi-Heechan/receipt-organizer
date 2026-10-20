import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import SCRIPTS  # noqa: E402

from scan_folder import scan  # noqa: E402


class ScanFolderTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)

    def touch(self, rel):
        p = self.root / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"x")

    def test_multi_mode_lists_people_sorted_recursive(self):
        for rel in ["B/1.JPG", "A/sub/2.pdf", "A/1.png"]:
            self.touch(rel)
        r = scan(self.root)
        self.assertEqual(r["mode"], "multi")
        self.assertEqual([f["name"] for f in r["folders"]], ["A", "B"])
        self.assertEqual(r["folders"][0]["files"], ["1.png", "sub/2.pdf"])
        self.assertEqual(r["folders"][1]["files"], ["1.JPG"])
        self.assertEqual(r["folders"][0]["path"], str((self.root / "A").resolve()))

    def test_ignores_system_and_lock_files(self):
        for rel in ["A/.DS_Store", "A/Thumbs.db", "A/~$x.xlsx", "A/.~lock.y#",
                    "A/.hidden.jpg", "A/.hid/z.jpg", "A/desktop.ini"]:
            self.touch(rel)
        r = scan(self.root)
        self.assertEqual(r["folders"][0]["files"], [])
        self.assertEqual(r["skipped"], [])
        self.assertEqual(r["empty_folders"], ["A"])

    def test_unsupported_reported_as_skipped(self):
        self.touch("A/x.heic")
        self.touch("A/1.jpg")
        r = scan(self.root)
        self.assertEqual(r["skipped"], [{"path": "A/x.heic", "reason": "unsupported"}])

    def test_loose_receipts_in_root_reported_but_outputs_ignored(self):
        for rel in ["c.jpg", "list_receipt_20261020.xlsx", "A/1.jpg"]:
            self.touch(rel)
        r = scan(self.root)
        self.assertEqual(r["loose_files"], ["c.jpg"])

    def test_single_mode_when_no_subfolders(self):
        self.touch("1.jpg")
        r = scan(self.root)
        self.assertEqual(r["mode"], "single")
        self.assertEqual(len(r["folders"]), 1)
        self.assertEqual(r["folders"][0]["name"], self.root.name)
        self.assertEqual(r["folders"][0]["files"], ["1.jpg"])
        self.assertEqual(r["loose_files"], [])

    def test_japanese_and_space_names(self):
        self.touch("株式会社 A/領収書 1.jpg")
        r = scan(self.root)
        self.assertEqual(r["folders"][0]["name"], "株式会社 A")
        self.assertEqual(r["folders"][0]["files"], ["領収書 1.jpg"])

    def test_cli_outputs_json_and_errors_on_missing_root(self):
        self.touch("株式会社 A/1.jpg")
        script = str(SCRIPTS / "scan_folder.py")
        ok = subprocess.run([sys.executable, script, str(self.root)], capture_output=True)
        self.assertEqual(ok.returncode, 0)
        data = json.loads(ok.stdout.decode("utf-8"))
        self.assertEqual(data["folders"][0]["name"], "株式会社 A")
        bad = subprocess.run([sys.executable, script, str(self.root / "nope")], capture_output=True)
        self.assertEqual(bad.returncode, 1)
        self.assertTrue(bad.stderr.decode("utf-8").strip())


if __name__ == "__main__":
    unittest.main()
