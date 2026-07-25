from __future__ import annotations

import ast
import os
import shutil
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    names = {
        "_valid_file_path",
        "_infos_declared_path",
        "_annotation_resource_paths",
        "_annotation_laptop_mirror_photos_dir",
        "_annotation_photos_source_set",
        "_path_accessible_quick",
        "_annotation_file_profile",
        "_annotation_compare_state",
        "_annotation_resource_audit",
        "_file_audit_row",
        "_annotation_report_resource_rows",
        "_csv_schema_ok",
        "_csv_row_count",
    }
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)
    ns = {
        "Path": Path,
        "csv": __import__("csv"),
        "datetime": datetime,
        "time": __import__("time"),
        "_ANNOTATION_FILE_PROFILE_CACHE": {},
        "ANNOTATION_FILE_PROFILE_TTL_SECONDS": 45.0,
        "AFFAIRES_ROOT": r"C:\Affaires",
        "_annotation_cache_prune": lambda cache, ttl: None,
        "_sha256_file": lambda path: "sha",
        "_test_path_with_timeout": lambda path: (Path(path).exists(), ""),
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class AnnotationResourcePathsTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        if not temp_root.parent.exists():
            temp_root = APP_PATH.parent / "temp"
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"resource_case_{uuid.uuid4().hex}"
        self.base.mkdir(parents=True, exist_ok=False)
        self.ns = _load_functions()
        self.infos_path = self.base / "infos_projet.json"
        self.infos_path.write_text("{}", encoding="utf-8")
        self.paths = {
            "nas_photos": self.base / "nas" / "photos.csv",
            "nas_photos_batch": self.base / "nas" / "photos_batch.csv",
            "pcfixe_photos": self.base / "pc" / "photos.csv",
            "pcfixe_photos_batch": self.base / "pc" / "photos_batch.csv",
            "pcfixe_unc_photos": self.base / "pc_unc" / "photos.csv",
            "pcfixe_unc_photos_batch": self.base / "pc_unc" / "photos_batch.csv",
            "pcfixe_unc_photos_dir": self.base / "pc_unc",
            "pcfixe_photos_dir": self.base / "pc",
            "nas_infos": self.infos_path,
            "pcfixe_infos": self.base / "pc" / "infos_projet.json",
            "pcfixe_unc_infos": self.base / "pc_unc" / "infos_projet.json",
        }
        for path in (
            self.paths["nas_photos"],
            self.paths["nas_photos_batch"],
            self.paths["pcfixe_photos"],
            self.paths["pcfixe_photos_batch"],
            self.paths["pcfixe_unc_photos"],
            self.paths["pcfixe_unc_photos_batch"],
        ):
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("photo_rel_native,nom_fichier_image\np1,P1.JPG\n", encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def test_valid_file_path_rejects_empty_variants(self):
        fn = self.ns["_valid_file_path"]
        self.assertIsNone(fn(None))
        self.assertIsNone(fn(""))
        self.assertIsNone(fn("   "))
        self.assertIsNone(fn("."))
        self.assertIsNone(fn(Path("")))
        self.assertIsNone(fn(Path(".")))

    def test_valid_file_path_accepts_missing_but_named_path(self):
        fn = self.ns["_valid_file_path"]
        candidate = fn(self.base / "missing" / "photos.csv")
        self.assertEqual(candidate, self.base / "missing" / "photos.csv")

    def test_annotation_resource_paths_handles_empty_inputs(self):
        fn = self.ns["_annotation_resource_paths"]
        resources = fn(self.infos_path, {"fichier_photos": "", "fichier_photos_batch": ""})
        self.assertIsNone(resources["photos.csv"])
        self.assertIsNone(resources["photos_batch.csv"])

    def test_annotation_resource_paths_derives_neighbor_batch(self):
        fn = self.ns["_annotation_resource_paths"]
        photos = self.base / "work" / "photos.csv"
        resources = fn(self.infos_path, {"fichier_photos": str(photos)})
        self.assertEqual(resources["photos.csv"], photos)
        self.assertEqual(resources["photos_batch.csv"], photos.with_name("photos_batch.csv"))

    def test_annotation_resource_paths_handles_path_dot_without_exception(self):
        fn = self.ns["_annotation_resource_paths"]
        resources = fn(self.infos_path, {"fichier_photos": Path("."), "fichier_photos_batch": Path("")})
        self.assertIsNone(resources["photos.csv"])
        self.assertIsNone(resources["photos_batch.csv"])

    def test_annotation_photos_source_set_falls_back_to_nas_when_missing(self):
        fn = self.ns["_annotation_photos_source_set"]
        source_set = fn(self.paths, {"fichier_photos": "", "fichier_photos_batch": ""}, "2022-J01", "accedit-2025-04-07")
        self.assertEqual(source_set["photos.csv"]["work"], self.paths["nas_photos"])
        self.assertEqual(source_set["photos_batch.csv"]["work"], self.paths["nas_photos_batch"])

    def test_report_resource_rows_handles_missing_declared_paths(self):
        fn = self.ns["_annotation_report_resource_rows"]
        rows = fn(
            self.paths,
            {"pcfixe": {"fichier_photos": "", "fichier_photos_batch": Path(".")}},
        )
        by_name = {row["ressource"]: row for row in rows}
        self.assertTrue(by_name["photos.csv"]["chemin UI"].endswith("photos.csv"))
        self.assertTrue(by_name["photos_batch.csv"]["chemin UI"].endswith("photos_batch.csv"))

    def test_resource_audit_handles_2022_j01_case_without_traceback(self):
        fn = self.ns["_annotation_resource_audit"]
        audit = fn(self.paths, {"fichier_photos": "", "fichier_photos_batch": ""}, "2022-J01", "accedit-2025-04-07")
        self.assertIn("rows", audit)
        self.assertTrue(any(row["ressource"] == "photos.csv" for row in audit["rows"]))

    def test_resource_audit_keeps_2025_j47_valid_case(self):
        fn = self.ns["_annotation_resource_audit"]
        photos = self.base / "j47" / "photos.csv"
        batch = self.base / "j47" / "photos_batch.csv"
        photos.parent.mkdir(parents=True, exist_ok=True)
        photos.write_text("photo_rel_native,nom_fichier_image\np1,P1.JPG\n", encoding="utf-8")
        batch.write_text("photo_rel_native,nom_fichier_image\np1,P1.JPG\n", encoding="utf-8")
        audit = fn(
            self.paths,
            {"fichier_photos": str(photos), "fichier_photos_batch": str(batch)},
            "2025-J47",
            "accedit-2025-11-13",
        )
        rows = [row for row in audit["rows"] if row["ressource"] == "photos.csv" and row["niveau"] == "fichier de travail UI"]
        self.assertEqual(rows[0]["présent"], "oui")


if __name__ == "__main__":
    unittest.main()
