import hashlib
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


CHECKER = Path(__file__).resolve().parents[1] / "check-repository-storage.py"


class RepositoryStorageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        subprocess.run(["git", "init", "-q", str(self.root)], check=True)

    def stage(self, name, data):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        # Stage exact bytes so tests can simulate broken LFS configuration.
        oid = subprocess.check_output(["git", "hash-object", "-w", "--stdin"],
                                      cwd=self.root, input=data).decode().strip()
        subprocess.run(["git", "update-index", "--add", "--cacheinfo", "100644", oid, name],
                       cwd=self.root, check=True)

    def check(self, *args):
        return subprocess.run([sys.executable, str(CHECKER), *args], cwd=self.root,
                              text=True, capture_output=True)

    def test_large_blob_and_generated_artifacts_are_rejected(self):
        self.stage("artifacts/run/result.log", b"run result")
        self.stage("oversized.dat", b"x" * (5 * 1024 * 1024 + 1))
        result = self.check()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Generated artifacts", result.stderr)
        self.assertIn("exceeds 5 MiB", result.stderr)

    def test_raw_lfs_blob_is_rejected_even_if_worktree_attributes_differ(self):
        self.stage(".gitattributes", b"*.zip filter=lfs -text\n")
        self.stage("fixture.zip", b"raw payload")
        (self.root / ".gitattributes").write_text("*.zip -filter\n")
        result = self.check()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("raw blob or invalid pointer", result.stderr)

    def test_optional_evidence_and_required_fixture_hydration(self):
        payload = b"example fixture bytes"
        pointer = ("version https://git-lfs.github.com/spec/v1\n"
                   f"oid sha256:{hashlib.sha256(payload).hexdigest()}\n"
                   f"size {len(payload)}\n").encode()
        self.stage(".gitattributes", b"*.zip filter=lfs -text\n")
        self.stage("docs/evidence/archive.zip", pointer)
        self.stage("tests/fixture.zip", pointer)
        self.assertEqual(self.check().returncode, 0)
        self.assertNotEqual(self.check("--require-materialized", "tests").returncode, 0)
        (self.root / "tests/fixture.zip").write_bytes(payload)
        self.assertEqual(self.check("--require-materialized", "tests").returncode, 0)
        (self.root / "tests/fixture.zip").write_bytes(b"x" * len(payload))
        result = self.check("--require-materialized", "tests")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("hash mismatch", result.stderr)


if __name__ == "__main__":
    unittest.main()
