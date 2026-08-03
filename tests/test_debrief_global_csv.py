import ast
import csv
import hashlib
import io
import json
import os
from pathlib import Path
import tempfile
import time
import unittest
import uuid
from datetime import datetime


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
FUNCTIONS = {
    "_sha256_file",
    "_atomic_write_debrief_bytes",
    "_atomic_write_debrief_json",
    "_read_csv_rows_with_dialect",
    "build_debrief_global_csv",
    "update_infos_projet_debrief_global_csv",
}


def _load_functions(pcfixe_root: Path):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    nodes = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS]

    def load_json(path, default):
        try:
            return json.loads(Path(path).read_text(encoding="utf-8-sig"))
        except Exception:
            return default

    def save_json(path, data):
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    namespace = {
        "Path": Path,
        "hashlib": hashlib,
        "uuid": uuid,
        "os": os,
        "json": json,
        "csv": csv,
        "io": io,
        "datetime": datetime,
        "time": time,
        "load_json": load_json,
        "save_json": save_json,
        "PCFIXE_AFFAIRES_ROOT": pcfixe_root,
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


def _write_source(path: Path, rows: list[tuple[str, str, str, str]]):
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(["start", "end", "speaker", "text"])
        writer.writerows(rows)


class DebriefGlobalCsvTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.ns = _load_functions(self.root / "pcfixe")
        self.debrief_dir = self.root / "debrief"
        self.debrief_dir.mkdir()
        self.source_a = self.debrief_dir / "historique(wav).csv"
        self.source_b = self.debrief_dir / "debrief_dictees.csv"
        _write_source(self.source_a, [("0", "1", "A", "ancien")])
        _write_source(self.source_b, [("2", "3", "B", "dictée")])
        self.canonical = self.debrief_dir / "debrief_global.csv"

    def tearDown(self):
        self.temp.cleanup()

    def build(self, **kwargs):
        return self.ns["build_debrief_global_csv"](
            sources=[self.source_a, self.source_b],
            output_csv=kwargs.pop("output_csv", self.canonical),
            id_affaire="2025-J47",
            id_captation="cap",
            **kwargs,
        )

    def test_canonical_absent_is_created(self):
        manifest = self.build()
        self.assertTrue(self.canonical.is_file())
        self.assertEqual(manifest["total_rows"], 2)
        self.assertEqual(manifest["output_path"], str(self.canonical))
        self.assertTrue(manifest["content_changed"])

    def test_different_content_replaces_canonical_and_updates_mtime(self):
        self.canonical.write_text("old", encoding="utf-8")
        os.utime(self.canonical, (1000, 1000))
        before_sha = hashlib.sha256(self.canonical.read_bytes()).hexdigest()
        manifest = self.build(allow_overwrite=True)
        self.assertNotEqual(before_sha, manifest["sha256_after"])
        self.assertGreater(self.canonical.stat().st_mtime_ns, 1_000_000_000_000)
        self.assertGreater(manifest["mtime_after_ns"], manifest["mtime_before_ns"])

    def test_identical_content_keeps_mtime(self):
        self.build()
        before_mtime = self.canonical.stat().st_mtime_ns
        before_sha = hashlib.sha256(self.canonical.read_bytes()).hexdigest()
        manifest = self.build(allow_overwrite=True)
        self.assertFalse(manifest["content_changed"])
        self.assertEqual(manifest["message"], "contenu inchangé")
        self.assertEqual(self.canonical.stat().st_mtime_ns, before_mtime)
        self.assertEqual(manifest["sha256_after"], before_sha)

    def test_timestamped_option_archives_and_updates_canonical(self):
        self.canonical.write_text("old", encoding="utf-8")
        manifest = self.build(create_timestamped_archive=True)
        archive = Path(manifest["archive_path"])
        self.assertTrue(archive.is_file())
        self.assertRegex(archive.name, r"^debrief_global_\d{8}_\d{6}(?:_\d{2})?\.csv$")
        self.assertEqual(archive.read_bytes(), self.canonical.read_bytes())
        self.assertEqual(manifest["output_path"], str(self.canonical))
        self.assertTrue(self.canonical.read_text(encoding="utf-8").startswith("start;end;speaker;text"))

    def test_requested_timestamp_name_still_publishes_canonical(self):
        requested = self.debrief_dir / "debrief_global_20260803_120000.csv"
        manifest = self.build(output_csv=requested)
        self.assertEqual(manifest["output_path"], str(self.canonical))
        self.assertFalse(requested.exists())

    def test_infos_projet_always_points_to_canonical(self):
        self.build()
        infos = self.root / "infos_projet.json"
        infos.write_text("{}", encoding="utf-8")
        fake_archive = self.debrief_dir / "debrief_global_20260803_120000.csv"
        updated = self.ns["update_infos_projet_debrief_global_csv"](
            infos,
            global_csv=fake_archive,
            csv_sources=[self.source_a, self.source_b],
            id_affaire="2025-J47",
            id_captation="cap",
            nas_trans_dir=self.debrief_dir,
        )
        self.assertEqual(Path(updated["debrief"]["global_csv"]).name, "debrief_global.csv")
        self.assertEqual(Path(updated["debrief"]["csv"]).name, "debrief_global.csv")
        self.assertEqual(Path(updated["fichier_debrief"]).name, "debrief_global.csv")
        self.assertEqual(Path(updated["pcfixe"]["fichier_debrief"]).name, "debrief_global.csv")

    def test_no_metadata_preserving_copy_is_used(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        tree = ast.parse(source)
        selected = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS]
        rendered = "\n".join(ast.unparse(node) for node in selected)
        self.assertNotIn("copy2", rendered)
        self.assertNotIn("copystat", rendered)
        self.assertIn("os.replace", rendered)
        self.assertIn("os.fsync", rendered)


if __name__ == "__main__":
    unittest.main()
