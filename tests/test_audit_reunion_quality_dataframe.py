from __future__ import annotations

import ast
import unittest
from pathlib import Path

import pandas as pd
import pyarrow as pa


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_dataframe_preparer():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    function = next(
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name == "prepare_audit_reunion_quality_dataframe"
    )
    module = ast.Module(body=[function], type_ignores=[])
    ast.fix_missing_locations(module)
    namespace = {"pd": pd}
    exec(compile(module, str(APP_PATH), "exec"), namespace)
    return namespace["prepare_audit_reunion_quality_dataframe"]


class AuditReunionQualityDataframeTests(unittest.TestCase):
    def test_taille_is_nullable_integer_and_arrow_compatible(self):
        prepare = _load_dataframe_preparer()
        rows = [
            {"artefact": "integer", "taille": 12},
            {"artefact": "numeric_string", "taille": "34"},
            {"artefact": "empty", "taille": ""},
            {"artefact": "none", "taille": None},
        ]

        dataframe = prepare(rows)
        arrow_table = pa.Table.from_pandas(dataframe, preserve_index=False)

        self.assertEqual(str(dataframe["taille"].dtype), "Int64")
        self.assertEqual(dataframe["taille"].tolist()[:2], [12, 34])
        self.assertTrue(pd.isna(dataframe.loc[2, "taille"]))
        self.assertTrue(pd.isna(dataframe.loc[3, "taille"]))
        self.assertEqual(arrow_table.column("taille").to_pylist(), [12, 34, None, None])
        self.assertEqual(dataframe["artefact"].tolist(), [row["artefact"] for row in rows])


if __name__ == "__main__":
    unittest.main()
