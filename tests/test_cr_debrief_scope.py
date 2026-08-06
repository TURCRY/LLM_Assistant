import ast
import csv
import io
import json
import os
from pathlib import Path
import tempfile
import unittest


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
FUNCTIONS = {
    "sync_cr_debrief_scope_state",
    "default_debrief_csv_selection",
    "assert_debrief_global_destination_matches_captation",
    "_is_debrief_global_csv",
    "_load_debrief_global_manifests",
    "_debrief_manifest_source_names",
    "_is_debrief_csv_excluded",
    "_debrief_csv_keyword_score",
    "_debrief_wav_match_score",
    "_debrief_csv_asr_text_info",
    "_read_csv_rows_with_dialect",
    "list_debrief_csv_candidates",
}


def _load_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    nodes = [
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS
    ]

    def load_json(path, default):
        try:
            return json.loads(Path(path).read_text(encoding="utf-8-sig"))
        except Exception:
            return default

    namespace = {
        "Path": Path,
        "csv": csv,
        "io": io,
        "json": json,
        "os": os,
        "load_json": load_json,
        "DEBRIEF_CSV_KEYWORDS": (
            "debrief", "debref", "amendement", "correction", "complément", "complement",
        ),
        "DEBRIEF_ASR_TEXT_COLUMNS": ("text", "texte", "transcript", "transcription"),
        "CR_DEBRIEF_SCOPE_STATE_KEYS": (
            "cr_debrief_csv_dir",
            "cr_debrief_csv_sources",
            "cr_debrief_csv_manual_paths",
            "cr_debrief_csv_order",
            "cr_debrief_global_output_display",
        ),
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class CrDebriefScopeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name) / "2026-A60" / "AF_Expert_ASR" / "transcriptions"
        self.old_captation = "accedit-2026-05-21"
        self.new_captation = "accedit-2026-06-17"
        self.old_dir = self.root / self.old_captation / "debrief"
        self.new_dir = self.root / self.new_captation / "debrief"
        self.new_dir.mkdir(parents=True)
        self.dictees = self.new_dir / "debrief_dictees.csv"
        self.dictees.write_text(
            "start;end;speaker;text\n0;1;LOCUTEUR_00;Dictée active\n",
            encoding="utf-8-sig",
        )
        self.ns = _load_functions()

    def tearDown(self):
        self.temp.cleanup()

    def test_captation_change_rebuilds_scan_selection_and_destination(self):
        state = {}
        self.ns["sync_cr_debrief_scope_state"](
            state,
            id_affaire="2026-A60",
            id_captation=self.old_captation,
        )
        old_global = self.old_dir / "debrief_global.csv"
        state.update({
            "cr_debrief_csv_dir": str(self.old_dir),
            "cr_debrief_csv_sources": ["ancien.csv"],
            "cr_debrief_csv_manual_paths": str(self.old_dir / "ancien.csv"),
            "cr_debrief_csv_order": "ancien.csv",
            "cr_debrief_global_output_display": str(old_global),
            "cr_debrief_csv_canonical_dir_old": str(self.old_dir),
            "cr_debrief_csv_advanced": False,
            "cr_debrief_global_timestamped": True,
        })

        scope, changed = self.ns["sync_cr_debrief_scope_state"](
            state,
            id_affaire="2026-A60",
            id_captation=self.new_captation,
        )
        self.assertTrue(changed)
        self.assertEqual(scope, f"2026-A60|{self.new_captation}")
        for key in (
            "cr_debrief_csv_dir",
            "cr_debrief_csv_sources",
            "cr_debrief_csv_manual_paths",
            "cr_debrief_csv_order",
            "cr_debrief_global_output_display",
            "cr_debrief_csv_canonical_dir_old",
        ):
            self.assertNotIn(key, state)
        self.assertFalse(state["cr_debrief_csv_advanced"])
        self.assertTrue(state["cr_debrief_global_timestamped"])
        self.assertFalse(self.old_dir.exists())

        candidates = self.ns["list_debrief_csv_candidates"](self.new_dir)
        visible = [item for item in candidates if not item["excluded"]]
        self.assertEqual([item["name"] for item in visible], ["debrief_dictees.csv"])
        options = [item["name"] for item in visible]
        lookup = {item["name"]: item["path"] for item in visible}
        selected = self.ns["default_debrief_csv_selection"](options, lookup, None)
        self.assertEqual(selected, ["debrief_dictees.csv"])

        displayed_destination = self.new_dir / "debrief_global.csv"
        effective_destination = self.new_dir / "debrief_global.csv"
        self.assertEqual(displayed_destination, effective_destination)
        self.ns["assert_debrief_global_destination_matches_captation"](
            effective_destination,
            canonical_debrief_dir=self.new_dir,
            id_captation=self.new_captation,
        )
        rendered = json.dumps(
            {
                "state": state,
                "candidates": [str(item["path"]) for item in visible],
                "selected": selected,
                "displayed": str(displayed_destination),
                "effective": str(effective_destination),
            },
            ensure_ascii=False,
        )
        self.assertNotIn(self.old_captation, rendered)

    def test_destination_guard_rejects_another_captation(self):
        with self.assertRaisesRegex(ValueError, self.new_captation):
            self.ns["assert_debrief_global_destination_matches_captation"](
                self.old_dir / "debrief_global.csv",
                canonical_debrief_dir=self.new_dir,
                id_captation=self.new_captation,
            )
        with self.assertRaisesRegex(ValueError, self.new_captation):
            self.ns["assert_debrief_global_destination_matches_captation"](
                self.old_dir / "debrief_global.csv",
                canonical_debrief_dir=self.old_dir,
                id_captation=self.new_captation,
            )


if __name__ == "__main__":
    unittest.main()
