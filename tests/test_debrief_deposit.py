"""Tests deterministes du depot de WAV de debrief (laptop / NAS / PC fixe).

Aucun acces reseau reel : les trois racines (laptop, NAS, PC fixe) sont injectees
dans un dossier temporaire. Aucun job ASR n'est soumis.
"""

import ast
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest
import uuid


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
FUNCTIONS = {
    "sha256_file",
    "_glob_existing_files",
    "debrief_affaire_root",
    "debrief_audio_dest_dir",
    "validate_debrief_audio_filename",
    "resolve_debrief_audio_destinations",
    "debrief_audio_scan_dirs",
    "list_debrief_audio_candidates",
    "probe_debrief_audio_destinations",
    "debrief_audio_source_fingerprint",
    "deposit_debrief_audio_to_dir",
    "deposit_debrief_audio",
    "debrief_deposit_rows",
}
CONSTANTS = {"DEBRIEF_AUDIO_UPLOAD_EXTENSIONS", "DEBRIEF_AUDIO_SCAN_PATTERNS"}


def _load_namespace():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    body = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS:
            body.append(node)
        elif isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id in CONSTANTS
            for target in node.targets
        ):
            body.append(node)

    namespace = {
        "Path": Path,
        "os": os,
        "hashlib": hashlib,
        "uuid": uuid,
        "shutil": shutil,
        "AFFAIRES_ROOT": r"C:\Affaires",
        "effective_nas_affaire_root": lambda cfg, aff: rf"\\192.168.1.20\Affaires\{aff}",
        "pcfixe_unc_root_for_laptop": lambda cfg, aff: rf"\\192.168.0.155\Affaires\{aff}",
        "_test_path_with_timeout": lambda raw: (Path(raw).exists(), ""),
    }
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class DebriefDepositTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name)
        self.roots = {
            "laptop": self.base / "laptop" / "Affaires",
            "nas": self.base / "nas" / "Affaires",
            "pcfixe": self.base / "pcfixe" / "Affaires",
        }
        self.ns = _load_namespace()
        self.affaire = "2025-J48"
        self.captation = "accedit-2025-11-20"
        self.wav_name = "J48 Enghien_Debrief_2025-11-20_mono16_16000Hz.wav"
        self.wav_bytes = b"RIFF" + bytes(range(256)) * 8
        self.wav_sha = hashlib.sha256(self.wav_bytes).hexdigest()
        self.relative_parts = (
            self.affaire,
            "AE_Expert_captations",
            self.captation,
            "debrief",
        )

    def tearDown(self):
        self.temp.cleanup()

    # ------------------------------------------------------------------ helpers
    def _destinations(self, filename=""):
        return self.ns["resolve_debrief_audio_destinations"](
            project_config={},
            affaire_id=self.affaire,
            id_captation=self.captation,
            filename=filename,
            laptop_root=self.roots["laptop"],
            nas_root=self.roots["nas"],
            pcfixe_root=self.roots["pcfixe"],
        )

    def _deposit(self, **overrides):
        kwargs = {
            "project_config": {},
            "affaire_id": self.affaire,
            "id_captation": self.captation,
            "payload": self.wav_bytes,
            "filename": self.wav_name,
            "laptop_root": self.roots["laptop"],
            "nas_root": self.roots["nas"],
            "pcfixe_root": self.roots["pcfixe"],
        }
        kwargs.update(overrides)
        return self.ns["deposit_debrief_audio"](**kwargs)

    # ------------------------------------------------------- destinations
    def test_destination_canonique_laptop_nas_pcfixe(self):
        destinations = self._destinations(self.wav_name)
        self.assertTrue(destinations["ready"])
        self.assertEqual(destinations["relative_dir"], os.path.join(*self.relative_parts))
        for role, root in self.roots.items():
            expected_dir = Path(root).joinpath(*self.relative_parts)
            self.assertEqual(Path(destinations["targets"][role]["dir"]), expected_dir)
            self.assertEqual(Path(destinations["targets"][role]["path"]), expected_dir / self.wav_name)
        self.assertEqual(destinations["filename"], self.wav_name)

    def test_destination_non_resolue_sans_captation(self):
        destinations = self.ns["resolve_debrief_audio_destinations"](
            project_config={},
            affaire_id=self.affaire,
            id_captation="",
            filename=self.wav_name,
            laptop_root=self.roots["laptop"],
            nas_root=self.roots["nas"],
            pcfixe_root=self.roots["pcfixe"],
        )
        self.assertFalse(destinations["ready"])
        self.assertEqual(destinations["relative_dir"], "")
        self.assertEqual(destinations["targets"]["laptop"]["dir"], "")
        self.assertEqual(destinations["targets"]["laptop"]["path"], "")

    def test_racines_par_defaut_utilisent_les_helpers_de_l_application(self):
        destinations = self.ns["resolve_debrief_audio_destinations"](
            project_config={},
            affaire_id=self.affaire,
            id_captation=self.captation,
            filename=self.wav_name,
        )
        self.assertEqual(
            destinations["targets"]["laptop"]["root"],
            os.path.join(r"C:\Affaires", self.affaire),
        )
        self.assertEqual(
            destinations["targets"]["nas"]["root"],
            rf"\\192.168.1.20\Affaires\{self.affaire}",
        )
        self.assertEqual(
            destinations["targets"]["pcfixe"]["root"],
            rf"\\192.168.0.155\Affaires\{self.affaire}",
        )

    def test_racine_deja_cadree_sur_l_affaire_non_dupliquee(self):
        """Cas reel : roots.nas / roots.pcfixe pointent deja sur <...>/Affaires/<affaire>."""
        affaire_root = self.base / "pcfixe" / "Affaires" / self.affaire
        destinations = self.ns["resolve_debrief_audio_destinations"](
            project_config={},
            affaire_id=self.affaire,
            id_captation=self.captation,
            filename=self.wav_name,
            laptop_root=self.roots["laptop"],
            nas_root=self.roots["nas"],
            pcfixe_root=affaire_root,
        )
        pcfixe_dir = Path(destinations["targets"]["pcfixe"]["dir"])
        self.assertEqual(
            pcfixe_dir,
            affaire_root / "AE_Expert_captations" / self.captation / "debrief",
        )
        self.assertNotIn(self.affaire + os.sep + self.affaire, str(pcfixe_dir))

    # ------------------------------------------------------------- validation
    def test_extensions_audio_admises(self):
        self.assertIn(".wav", self.ns["DEBRIEF_AUDIO_UPLOAD_EXTENSIONS"])
        self.assertIn("*.wav", self.ns["DEBRIEF_AUDIO_SCAN_PATTERNS"])
        self.assertIn("*.WAV", self.ns["DEBRIEF_AUDIO_SCAN_PATTERNS"])
        self.assertEqual(
            self.ns["validate_debrief_audio_filename"](self.wav_name),
            self.wav_name,
        )
        for rejected in ("rapport.txt", "note.pdf", ""):
            with self.assertRaises(ValueError):
                self.ns["validate_debrief_audio_filename"](rejected)

    def test_nom_de_fichier_reduit_au_basename(self):
        validated = self.ns["validate_debrief_audio_filename"](
            r"C:\tmp\..\ailleurs" + os.sep + self.wav_name
        )
        self.assertEqual(validated, self.wav_name)

    def test_fingerprint_payload_et_fichier(self):
        payload_fp = self.ns["debrief_audio_source_fingerprint"](payload=self.wav_bytes)
        self.assertEqual(payload_fp["size"], len(self.wav_bytes))
        self.assertEqual(payload_fp["sha256"], self.wav_sha)

        source = self.base / self.wav_name
        source.write_bytes(self.wav_bytes)
        file_fp = self.ns["debrief_audio_source_fingerprint"](source_path=source)
        self.assertEqual(file_fp["size"], len(self.wav_bytes))
        self.assertEqual(file_fp["sha256"], self.wav_sha)
        self.assertEqual(file_fp["filename"], self.wav_name)

    # ---------------------------------------------------------------- depot
    def test_depot_laptop_cree_le_dossier_debrief(self):
        result = self._deposit(targets=("laptop",))
        self.assertTrue(result["ok"])
        self.assertEqual(result["targets"]["laptop"]["state"], "copied")
        deposited = Path(result["targets"]["laptop"]["path"])
        self.assertTrue(deposited.is_file())
        self.assertEqual(deposited.parent.name, "debrief")
        self.assertEqual(deposited.read_bytes(), self.wav_bytes)
        self.assertEqual(result["sha256"], self.wav_sha)
        self.assertEqual(result["targets"]["laptop"]["sha256"], self.wav_sha)
        self.assertEqual(result["requested"], ["laptop"])
        self.assertEqual(result["targets"]["pcfixe"]["state"], "skipped")

    def test_depot_sur_les_trois_cibles(self):
        result = self._deposit()
        self.assertTrue(result["ok"])
        self.assertEqual(result["requested"], ["laptop", "nas", "pcfixe"])
        self.assertEqual(result["failed"], [])
        self.assertEqual(result["conflicts"], [])
        for role in ("laptop", "nas", "pcfixe"):
            self.assertEqual(result["targets"][role]["state"], "copied")
            self.assertEqual(result["targets"][role]["sha256"], self.wav_sha)
            self.assertTrue(Path(result["targets"][role]["path"]).is_file())

    def test_fichier_identique_deja_present_non_recopie(self):
        first = self._deposit(targets=("laptop",))
        deposited = Path(first["targets"]["laptop"]["path"])
        mtime = deposited.stat().st_mtime_ns

        second = self._deposit(targets=("laptop",))
        self.assertTrue(second["ok"])
        self.assertEqual(second["targets"]["laptop"]["state"], "already_present")
        self.assertEqual(second["targets"]["laptop"]["sha256"], self.wav_sha)
        self.assertEqual(deposited.stat().st_mtime_ns, mtime)

    def test_homonyme_contenu_different_bloque_sans_ecraser(self):
        first = self._deposit(targets=("laptop",))
        deposited = Path(first["targets"]["laptop"]["path"])
        deposited.write_bytes(b"contenu different")

        result = self._deposit(targets=("laptop",))
        self.assertFalse(result["ok"])
        self.assertEqual(result["targets"]["laptop"]["state"], "conflict")
        self.assertEqual(result["conflicts"], ["laptop"])
        self.assertIn("contenu different", result["targets"]["laptop"]["error"])
        self.assertEqual(deposited.read_bytes(), b"contenu different")
        self.assertEqual(
            sorted(path.name for path in deposited.parent.iterdir() if path.is_file()),
            [self.wav_name],
        )

    def test_pc_fixe_inaccessible_est_signale(self):
        blocked_root = self.base / "pcfixe_hors_ligne"
        blocked_root.write_bytes(b"pas un dossier")
        result = self._deposit(pcfixe_root=blocked_root, targets=("pcfixe",))
        self.assertFalse(result["ok"])
        self.assertEqual(result["targets"]["pcfixe"]["state"], "failed")
        self.assertEqual(result["failed"], ["pcfixe"])
        self.assertTrue(result["targets"]["pcfixe"]["error"])
        self.assertEqual(result["targets"]["laptop"]["state"], "skipped")

    # ------------------------------------------------------------------ scan
    def test_wav_depose_immediatement_detecte(self):
        destinations = self._destinations(self.wav_name)
        self.assertEqual(self.ns["list_debrief_audio_candidates"](destinations), [])

        self._deposit()
        found = self.ns["list_debrief_audio_candidates"](destinations)
        self.assertEqual([path.name for path in found], [self.wav_name])
        self.assertEqual(Path(found[0]).parent, Path(destinations["targets"]["laptop"]["dir"]))

    def test_scan_laptop_prioritaire_et_dedoublonne(self):
        destinations = self._destinations()
        laptop_dir = Path(destinations["targets"]["laptop"]["dir"])
        nas_dir = Path(destinations["targets"]["nas"]["dir"])
        laptop_dir.mkdir(parents=True)
        nas_dir.mkdir(parents=True)
        (laptop_dir / self.wav_name).write_bytes(self.wav_bytes)
        (nas_dir / self.wav_name).write_bytes(self.wav_bytes)
        (nas_dir / "Autre debrief.WAV").write_bytes(self.wav_bytes)

        found = self.ns["list_debrief_audio_candidates"](destinations)
        self.assertEqual(
            [path.name for path in found],
            [self.wav_name, "Autre debrief.WAV"],
        )
        self.assertEqual(Path(found[0]).parent, laptop_dir)
        self.assertEqual(Path(found[1]).parent, nas_dir)
        self.assertEqual(
            self.ns["debrief_audio_scan_dirs"](destinations),
            [laptop_dir, nas_dir],
        )

    # ----------------------------------------------------------------- probe
    def test_probe_des_cibles(self):
        destinations = self._destinations(self.wav_name)
        laptop_affaire_root = os.path.join(str(self.roots["laptop"]), self.affaire)
        rows = self.ns["probe_debrief_audio_destinations"](
            destinations,
            probe_fn=lambda raw: (str(raw) == laptop_affaire_root, "simule"),
        )
        self.assertEqual([row["cible"] for row in rows], ["Laptop", "NAS", "PC fixe"])
        self.assertEqual(rows[0]["racine accessible"], "OK")
        self.assertEqual(rows[1]["racine accessible"], "inaccessible")
        self.assertTrue(rows[0]["destination"].endswith(self.wav_name))

    def test_lignes_de_resultat_de_depot(self):
        result = self._deposit()
        rows = self.ns["debrief_deposit_rows"](result)
        self.assertEqual([row["label"] for row in rows], ["Laptop", "NAS", "PC fixe"])
        self.assertEqual({row["state_label"] for row in rows}, {"copie"})
        self.assertTrue(all(row["sha256"] == self.wav_sha for row in rows))

# ---------------------------------------------------------------------------
# Non-regression du pipeline debrief existant (job asr_voxtral purpose=debrief)
# ---------------------------------------------------------------------------
class DebriefPipelineNonRegressionTests(unittest.TestCase):
    def test_job_asr_debrief_toujours_present(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        tree = ast.parse(source)
        names = {node.name for node in ast.walk(tree) if isinstance(node, ast.FunctionDef)}
        for expected in (
            "build_debrief_audio_block",
            "submit_asr_v2_job",
            "prepare_voxtral_audio",
            "prepare_selected_debrief_audio",
            "submit_debrief_asr_job",
        ):
            self.assertIn(expected, names)
        self.assertIn('"purpose": "debrief"', source)
        self.assertIn('"type": "asr_voxtral"', source)

    def test_ui_depot_debrief_branchee(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        # le dossier debrief ne doit plus dependre d'une cle de widget figee
        self.assertNotIn('key="voxtral_debrief_dir_path"', source)
        self.assertIn(
            'key=f"voxtral_debrief_dir_path_{affaire_id}_{selected_captation}"',
            source,
        )
        self.assertIn("st.file_uploader(", source)
        self.assertIn("deposit_debrief_audio(", source)
        self.assertIn("list_debrief_audio_candidates(debrief_destinations)", source)

    def test_copie_csv_debrief_gardee_par_is_file(self):
        """Lot WinError 3 : la copie du CSV debrief doit rester conditionnee a is_file()."""
        source = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertIn(
            "if pc_debrief and debrief_source and debrief_source.is_file():",
            source,
        )
        self.assertNotIn("if pc_debrief and debrief_source:\n", source)


# ---------------------------------------------------------------------------
# Garde de copie du CSV debrief dans submit_asr_v2_job (bug WinError 3)
# ---------------------------------------------------------------------------
ASR_FUNCTIONS = {"submit_asr_v2_job", "load_json"}


class _ShutilSpy:
    """Enregistre les appels copy2 tout en executant le vrai shutil.copy2."""

    def __init__(self, record):
        self._record = record

    def copy2(self, src, dst, **kwargs):
        self._record.append((str(src), str(dst)))
        return shutil.copy2(src, dst, **kwargs)

    def __getattr__(self, name):
        return getattr(shutil, name)


def _load_asr_namespace(overrides):
    """Charge submit_asr_v2_job avec des dependances injectees (aucun acces reseau)."""
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    body = [
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in ASR_FUNCTIONS
    ]
    namespace = {
        "Path": Path,
        "os": os,
        "json": json,
        "datetime": datetime.datetime,
        "uuid": uuid,
        "PCFIXE_AFFAIRES_ROOT": Path(r"C:\Affaires"),
        "PCFIXE_BOOST_FILE": Path(r"D:\GPT4All_Local\config\boost_vocab.txt"),
    }
    namespace.update(overrides)
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class SubmitAsrV2DebriefCopyGuardTests(unittest.TestCase):
    """Reproduit le bug WinError 3 : copie d'un CSV debrief pas encore produit."""

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.base = Path(self.temp.name)
        self.affaire = "2025-J48"
        self.captation = "accedit-2025-11-20"
        self.wav_name = "J48 Enghien_Debrief_2025-11-20_mono16_16000Hz.wav"
        self.csv_name = f"{Path(self.wav_name).stem}(wav).csv"

        self.pcfixe_root = self.base / "pcfixe" / "Affaires"
        self.queued = self.pcfixe_root / "_jobs" / "queued"
        self.unc_trans_dir = (
            self.pcfixe_root / self.affaire / "AF_Expert_ASR" / "transcriptions" / self.captation
        )
        self.unc_audio_dir = (
            self.pcfixe_root / self.affaire / "AE_Expert_captations" / self.captation / "debrief"
        )
        self.queued.mkdir(parents=True)
        self.unc_trans_dir.mkdir(parents=True)
        self.unc_audio_dir.mkdir(parents=True)

        self.laptop_trans_dir = (
            self.base / "laptop" / "AF_Expert_ASR" / "transcriptions" / self.captation
        )
        self.laptop_trans_dir.mkdir(parents=True)
        self.infos_laptop = self.laptop_trans_dir / "infos_projet.json"
        self.infos_laptop.write_text(
            json.dumps({"id_affaire": self.affaire, "id_captation": self.captation}),
            encoding="utf-8",
        )
        self.wav = (
            self.base / "laptop" / "AE_Expert_captations" / self.captation / "debrief" / self.wav_name
        )
        self.wav.parent.mkdir(parents=True)
        self.wav.write_bytes(b"RIFF-debrief")

        self.copies = []
        self.ns = _load_asr_namespace({
            "shutil": _ShutilSpy(self.copies),
            "get_pcfixe_affaires_root": lambda: Path(self.pcfixe_root),
            "get_pcfixe_jobs_queued_dir": lambda: Path(self.queued),
            "preflight_pcfixe_target_dir": lambda target: {
                "target_dir": str(target),
                "target_ready": True,
                "probes": [],
            },
        })

    def tearDown(self):
        self.temp.cleanup()

    def _submit(self, debrief_csv):
        return self.ns["submit_asr_v2_job"](
            infos_path=str(self.infos_laptop),
            audio_prepared=str(self.wav),
            proper_names_path=None,
            model="Voxtral_Mini_3B_Transformers",
            diarize=False,
            purpose="debrief",
            debrief_csv=str(debrief_csv),
            audio_target_dir_pcfixe=str(self.wav.parent),
            audio_target_dir_unc=str(self.unc_audio_dir),
            output_dir_pcfixe=str(self.laptop_trans_dir / "debrief"),
            output_dir_unc=str(self.unc_trans_dir / "debrief"),
        )

    def test_csv_debrief_absent_aucune_copie_tentee(self):
        """Le dossier parent du CSV n'existe pas : plus aucun copy2 ne doit partir."""
        csv_path = self.laptop_trans_dir / "debrief" / self.csv_name
        self.assertFalse(csv_path.parent.exists())

        result = self._submit(csv_path)

        self.assertEqual(
            self.copies,
            [(str(self.wav), str(self.unc_audio_dir / self.wav_name))],
        )
        self.assertEqual(result["job"]["debrief_csv"], str(csv_path))
        self.assertEqual(result["job"]["purpose"], "debrief")
        self.assertEqual(result["job"]["type"], "asr_voxtral")

    def test_csv_debrief_present_copie_effectuee(self):
        csv_dir = self.laptop_trans_dir / "debrief"
        csv_dir.mkdir(parents=True)
        csv_path = csv_dir / self.csv_name
        csv_path.write_text("row_ref,text\n1,debrief\n", encoding="utf-8")

        result = self._submit(csv_path)

        self.assertIn(
            (str(csv_path.resolve()), str(self.unc_trans_dir / self.csv_name)),
            self.copies,
        )
        self.assertEqual(len(self.copies), 2)
        self.assertTrue((self.unc_trans_dir / self.csv_name).is_file())
        self.assertEqual(result["job"]["debrief_csv"], str(csv_path))


if __name__ == "__main__":
    unittest.main()
