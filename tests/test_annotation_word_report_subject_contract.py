"""Tests du contrat optionnel photo -> sujet pour le job annotation_photos_word_report.

Le generateur Word etant execute sur le PC fixe, ce depot laptop ne transmet
que des chemins canoniques quand l affectation photo -> sujet existe. Sans
affectation, le job historique doit rester strictement inchange.
"""

from __future__ import annotations

import ast
import shutil
import tempfile
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
FIELDS_FN = "_annotation_word_report_photo_subject_fields"


def _find_function(name: str):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"fonction {name} introuvable dans app.py")


def _extract_fields_function():
    node = _find_function(FIELDS_FN)
    module = ast.Module(body=[node], type_ignores=[])
    ast.fix_missing_locations(module)
    ns = {"Path": Path}
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns[FIELDS_FN]


class PhotoSubjectFieldsFunctionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="annotation_word_contract_", dir=str(Path(__file__).resolve().parents[1] / "tests")))
        self.fields_fn = _extract_fields_function()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _paths(self):
        return {
            "nas_photos_dir": self.tmp / "nas" / "AE_Expert_captations" / "cap" / "photos",
            "nas_trans_dir": self.tmp / "nas" / "AF_Expert_ASR" / "transcriptions" / "cap",
            "pcfixe_photos_dir": self.tmp / "pc" / "AE_Expert_captations" / "cap" / "photos",
            "pcfixe_trans_dir": self.tmp / "pc" / "AF_Expert_ASR" / "transcriptions" / "cap",
        }

    def test_no_assignment_file_keeps_historical_contract(self):
        # Aucune affectation -> aucun champ ajoute (job historique inchange).
        self.assertEqual(self.fields_fn(self._paths()), {})

    def test_non_dict_input_returns_empty(self):
        self.assertEqual(self.fields_fn(None), {})
        self.assertEqual(self.fields_fn({}), {})

    def test_assignment_exists_adds_nas_and_mirror_paths(self):
        paths = self._paths()
        assignment = paths["nas_photos_dir"] / "photo_subject_assignments.json"
        assignment.parent.mkdir(parents=True, exist_ok=True)
        assignment.write_text("{}", encoding="utf-8")

        fields = self.fields_fn(paths)

        self.assertEqual(
            fields,
            {
                "photo_subject_assignments_nas": str(assignment),
                "sujets_xlsx_nas": str(paths["nas_trans_dir"] / "Sujets.xlsx"),
                "mirror_pc_photo_subject_assignments": str(
                    paths["pcfixe_photos_dir"] / "photo_subject_assignments.json"
                ),
                "mirror_pc_sujets_xlsx": str(paths["pcfixe_trans_dir"] / "Sujets.xlsx"),
            },
        )

    def test_assignment_exists_without_mirror_dirs_adds_nas_only(self):
        assignment = self.tmp / "nas" / "AE_Expert_captations" / "cap" / "photos" / "photo_subject_assignments.json"
        assignment.parent.mkdir(parents=True, exist_ok=True)
        assignment.write_text("{}", encoding="utf-8")
        paths = {
            "nas_photos_dir": assignment.parent,
            "nas_trans_dir": self.tmp / "nas" / "AF_Expert_ASR" / "transcriptions" / "cap",
        }

        fields = self.fields_fn(paths)

        self.assertEqual(
            fields,
            {
                "photo_subject_assignments_nas": str(assignment),
                "sujets_xlsx_nas": str(paths["nas_trans_dir"] / "Sujets.xlsx"),
            },
        )

    def test_assignment_file_missing_on_nas_ignores_mirror_copy(self):
        # Le miroir PC peut contenir un vieux fichier : sans la source NAS,
        # le job reste historique (le NAS est la source de verite).
        paths = self._paths()
        mirror = paths["pcfixe_photos_dir"] / "photo_subject_assignments.json"
        mirror.parent.mkdir(parents=True, exist_ok=True)
        mirror.write_text("{}", encoding="utf-8")

        self.assertEqual(self.fields_fn(paths), {})


class PhotoReportJobBuilderContractTests(unittest.TestCase):
    def test_builder_injects_optional_fields_after_job_dict_and_before_deposit(self):
        node = _find_function("_photo_report_job_preview")
        source = APP_PATH.read_text(encoding="utf-8-sig")
        segments = []
        for stmt in node.body:
            seg = ast.get_source_segment(source, stmt)
            if seg is None:
                continue
            segments.append(seg)

        text = "\n".join(segments)
        pos_job = text.find("job = {")
        pos_update = text.find("job.update(_annotation_word_report_photo_subject_fields(paths))")
        pos_queued = text.find("queued_path = get_pcfixe_jobs_queued_dir()")
        self.assertNotEqual(pos_job, -1)
        self.assertNotEqual(pos_update, -1)
        self.assertNotEqual(pos_queued, -1)
        self.assertTrue(pos_job < pos_update < pos_queued, "l injection doit preceder le depot du job")

        job_dict_literal = text[pos_job:pos_update]
        self.assertNotIn("photo_subject_assignments_nas", job_dict_literal)
        self.assertNotIn("sujets_xlsx_nas", job_dict_literal)

    def test_fields_are_never_required_by_builder(self):
        # Les champs n apparaissent qu une seule fois dans le builder (l update
        # optionnel) : aucune validation/raise ne les rend obligatoires.
        node = _find_function("_photo_report_job_preview")
        source = APP_PATH.read_text(encoding="utf-8-sig")
        body_text = "\n".join(
            seg for stmt in node.body if (seg := ast.get_source_segment(source, stmt)) is not None
        )
        self.assertEqual(body_text.count(FIELDS_FN), 1)
        self.assertNotIn("raise", body_text.replace("raise ValueError", "").replace("raise RuntimeError", ""))


class IncludeExcludedPhotosContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = APP_PATH.read_text(encoding="utf-8-sig")

    def test_builder_param_defaults_to_false_and_key_is_explicit(self):
        node = _find_function("_photo_report_job_preview")
        arg_names = [a.arg for a in node.args.kwonlyargs]
        self.assertIn("include_excluded_photos", arg_names)
        default = node.args.kw_defaults[arg_names.index("include_excluded_photos")]
        self.assertIsInstance(default, ast.Constant)
        self.assertIs(default.value, False)
        body = "\n".join(
            seg for stmt in node.body if (seg := ast.get_source_segment(self.source, stmt)) is not None
        )
        self.assertIn('"include_excluded_photos": bool(include_excluded_photos)', body)

    def test_ui_checkbox_is_disabled_by_default_and_wired_to_both_calls(self):
        self.assertIn("Inclure les photos exclues", self.source)
        self.assertIn('key="ann_photos_report_include_excluded"', self.source)
        self.assertEqual(self.source.count("include_excluded_photos=ann_report_include_excluded"), 2)


if __name__ == "__main__":
    unittest.main()
