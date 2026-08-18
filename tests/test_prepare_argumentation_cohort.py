from __future__ import annotations

import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.documents.prepare_argumentation_cohort import (
    build_manifest,
    detect_target_collisions,
    dwg_reference_pdf_lines,
    email_html,
    email_text,
    export_html_to_pdf_with_word,
    export_text_to_pdf_with_word,
    find_xlsx_pdf_equivalent,
    MsgMetadata,
    normalize_subject,
    parse_piece_token,
    source_item,
    target_name_for_dwg_reference,
    target_name_for_msg,
    target_name_for_pdf,
    target_name_for_png,
    convert_msg_to_pdf_atomic,
)


class FakeWordDoc:
    def __init__(self, export_error: Exception | None = None):
        self.export_error = export_error
        self.close_calls = []
        self.range_obj = types.SimpleNamespace(Text="")
        self.saved_as = []

    def Range(self):
        return self.range_obj

    def ExportAsFixedFormat(self, *args, **kwargs):
        if self.export_error:
            raise self.export_error

    def SaveAs2(self, *args, **kwargs):
        self.saved_as.append((args, kwargs))

    def Close(self, value):
        self.close_calls.append(value)


class FakeWordDocuments:
    def __init__(self, doc: FakeWordDoc | None = None, add_error: Exception | None = None):
        self.doc = doc or FakeWordDoc()
        self.add_error = add_error

    def Add(self):
        if self.add_error:
            raise self.add_error
        return self.doc

    def Open(self, **_kwargs):
        return self.doc


class FakeWordApp:
    def __init__(self, doc: FakeWordDoc | None = None, add_error: Exception | None = None):
        self.Documents = FakeWordDocuments(doc, add_error)
        self.quit_calls = 0
        self.Visible = True
        self.DisplayAlerts = None

    def Quit(self):
        self.quit_calls += 1


class FakePythonCom:
    def __init__(self):
        self.initialize_calls = 0
        self.uninitialize_calls = 0

    def CoInitialize(self):
        self.initialize_calls += 1

    def CoUninitialize(self):
        self.uninitialize_calls += 1


class PrepareArgumentationCohortTests(unittest.TestCase):
    def install_fake_word_modules(self, word_app: FakeWordApp, pythoncom: FakePythonCom | None = None):
        fake_client = types.SimpleNamespace(DispatchEx=lambda name: word_app)
        fake_win32com = types.ModuleType("win32com")
        fake_win32com.client = fake_client
        modules = {
            "win32com": fake_win32com,
            "win32com.client": fake_client,
        }
        if pythoncom is not None:
            modules["pythoncom"] = pythoncom
        return patch.dict(sys.modules, modules)

    def test_parse_piece_number_variants(self):
        self.assertEqual(parse_piece_token("5_SEBIA - Suivi.msg")[:3], ("5", "", "5"))
        self.assertEqual(parse_piece_token("5a_Re_ SEBIA.msg")[:3], ("5", "a", "5.a"))
        self.assertEqual(parse_piece_token("05a_SEBIA LIS2.msg")[:3], ("05", "a", "05.a"))
        self.assertEqual(parse_piece_token("10_SEBIA LIS2.pdf")[:3], ("10", "", "10"))

    def test_normalize_subject(self):
        self.assertEqual(normalize_subject("Réserve 2521"), "Reserve 2521")
        self.assertEqual(normalize_subject("  Introduction  "), "Introduction")

    def test_target_names_for_msg_and_annex_pdf(self):
        root = Path("root")
        msg = source_item(root / "Introduction" / "5_SEBIA - Suivi des écarts PCA _ EXP _ Cuisine.msg", root)
        pdf = source_item(root / "Introduction" / "5_24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA.pdf", root)
        self.assertEqual(
            target_name_for_msg(msg),
            "Piece 5 - Introduction - SEBIA - Suivi des écarts PCA EXP Cuisine.pdf",
        )
        self.assertEqual(
            target_name_for_pdf(pdf, has_msg_master=True),
            "Piece 5 - Introduction - Annexe - 24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA.pdf",
        )

    def test_email_html_uses_plain_text_body(self):
        metadata = MsgMetadata(
            subject="Sujet",
            body="Ligne 1\nLigne 2",
            html_body="<img src='cid:image001'><script>ignored</script>",
            attachments=["annexe.pdf"],
        )
        payload = email_html(metadata, Path("source.msg"))
        self.assertIn("<pre>Ligne 1\nLigne 2</pre>", payload)
        self.assertIn("annexe.pdf", payload)
        self.assertNotIn("cid:image001", payload)
        self.assertNotIn("<script>", payload)

    def test_email_text_contains_message_metadata_and_attachments(self):
        metadata = MsgMetadata(
            subject="Objet",
            sender="Expediteur",
            to="Destinataire",
            sent_on="2026-08-17",
            body="Corps complet",
            attachments=["a.pdf", "b.xlsx"],
        )
        payload = email_text(metadata, Path("source.msg"))
        self.assertIn("Objet", payload)
        self.assertIn("Expediteur", payload)
        self.assertIn("Corps complet", payload)
        self.assertIn("- a.pdf", payload)
        self.assertIn("- b.xlsx", payload)

    def test_word_export_uses_dedicated_dispatch_ex_and_quits_after_success(self):
        doc = FakeWordDoc()
        word = FakeWordApp(doc)
        pythoncom = FakePythonCom()
        with self.install_fake_word_modules(word, pythoncom):
            export_text_to_pdf_with_word("contenu", Path("unused.pdf"))
        self.assertEqual(doc.close_calls, [False])
        self.assertEqual(word.quit_calls, 1)
        self.assertEqual(pythoncom.initialize_calls, 1)
        self.assertEqual(pythoncom.uninitialize_calls, 1)

    def test_word_export_quits_after_export_exception(self):
        doc = FakeWordDoc(export_error=RuntimeError("export failed"))
        word = FakeWordApp(doc)
        with self.install_fake_word_modules(word, FakePythonCom()):
            with self.assertRaises(RuntimeError):
                export_text_to_pdf_with_word("contenu", Path("unused.pdf"))
        self.assertEqual(doc.close_calls, [False])
        self.assertEqual(word.quit_calls, 1)

    def test_word_export_quits_when_document_creation_fails(self):
        word = FakeWordApp(add_error=RuntimeError("add failed"))
        with self.install_fake_word_modules(word, FakePythonCom()):
            with self.assertRaises(RuntimeError):
                export_text_to_pdf_with_word("contenu", Path("unused.pdf"))
        self.assertEqual(word.quit_calls, 1)

    def test_word_html_export_closes_opened_document(self):
        doc = FakeWordDoc()
        word = FakeWordApp(doc)
        with self.install_fake_word_modules(word, FakePythonCom()):
            export_html_to_pdf_with_word(Path("source.html"), Path("unused.pdf"))
        self.assertEqual(doc.close_calls, [False])
        self.assertEqual(word.quit_calls, 1)

    def test_msg_conversion_uses_word_cleanup_on_export_failure(self):
        doc = FakeWordDoc(export_error=RuntimeError("export failed"))
        word = FakeWordApp(doc)
        metadata = MsgMetadata(subject="Sujet", body="Message")
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "message.msg"
            source.write_bytes(b"msg")
            with patch("scripts.documents.prepare_argumentation_cohort.read_msg_metadata_outlook", return_value=metadata):
                with self.install_fake_word_modules(word, FakePythonCom()):
                    with self.assertRaises(RuntimeError):
                        convert_msg_to_pdf_atomic(source, Path(tmp) / "message.pdf")
        self.assertEqual(doc.close_calls, [False])
        self.assertEqual(word.quit_calls, 1)

    def test_msg_pdf_same_subject_number_marks_pdf_as_annex(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Introduction"
            subject.mkdir(parents=True)
            (subject / "5_SEBIA - Suivi.msg").write_bytes(b"msg")
            (subject / "5_Annexe PCA.pdf").write_bytes(b"%PDF-1.4\n")
            rows, _ = build_manifest(source, target, read_msg_metadata=False)
            by_name = {row.source_name: row for row in rows}
            self.assertEqual(by_name["5_SEBIA - Suivi.msg"].role, "msg_maitre")
            self.assertEqual(by_name["5_Annexe PCA.pdf"].role, "annexe_pdf")
            self.assertIn("Annexe", by_name["5_Annexe PCA.pdf"].target_name)

    def test_pdf_without_msg_is_principal(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Reserve 4908"
            subject.mkdir(parents=True)
            (subject / "10_Document principal.pdf").write_bytes(b"%PDF-1.4\n")
            rows, _ = build_manifest(source, target, read_msg_metadata=False)
            self.assertEqual(rows[0].role, "pdf")
            self.assertNotIn("Annexe", rows[0].target_name)

    def test_xlsx_pdf_equivalence(self):
        root = Path("root")
        xlsx = source_item(root / "Introduction" / "7_2025 09 02 - SEBIA - Suivi Des Ecarts.xlsx", root)
        pdf = source_item(root / "Introduction" / "7_2025 09 02 - SEBIA - Suivi Des Ecarts.pdf", root)
        decision, reason, related_pdf = find_xlsx_pdf_equivalent(xlsx, [pdf])
        self.assertEqual(decision, "ignorer_ingestion_pdf_equivalent")
        self.assertIn(pdf.path.name, reason)
        self.assertEqual(related_pdf, str(pdf.path))

    def test_xlsx_pdf_equivalence_with_added_pdf_variant(self):
        root = Path("root")
        xlsx = source_item(root / "Introduction" / "9_Kairnial-Commentaires reserves.xlsx", root)
        pdf = source_item(root / "Introduction" / "9_2026-06-10_Kairnial-Commentaires reserves.pdf", root)
        decision, _, related_pdf = find_xlsx_pdf_equivalent(xlsx, [pdf])
        self.assertEqual(decision, "ignorer_ingestion_pdf_equivalent")
        self.assertEqual(related_pdf, str(pdf.path))

    def test_xlsx_pdf_equivalence_prefers_exact_pdf_when_multiple_match(self):
        root = Path("root")
        xlsx = source_item(root / "Reserve 5951" / "9_Kairnial-Commentaires reserves.xlsx", root)
        exact_pdf = source_item(root / "Reserve 5951" / "9_Kairnial-Commentaires reserves.pdf", root)
        dated_pdf = source_item(root / "Reserve 5951" / "9_2026-06-10_Kairnial-Commentaires reserves.pdf", root)
        decision, _, related_pdf = find_xlsx_pdf_equivalent(xlsx, [dated_pdf, exact_pdf])
        self.assertEqual(decision, "ignorer_ingestion_pdf_equivalent")
        self.assertEqual(related_pdf, str(exact_pdf.path))

    def test_office_temp_xlsx_is_ignored(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Réserve 2521"
            subject.mkdir(parents=True)
            temp_xlsx = subject / "~$1_SEBIA_Tableau synthèse.xlsx"
            temp_xlsx.write_bytes(b"temp")
            rows, summary = build_manifest(source, target, read_msg_metadata=False)
            self.assertEqual(rows[0].role, "office_temp")
            self.assertEqual(rows[0].decision, "ignorer_fichier_temporaire_office")
            self.assertEqual(summary["office_temp_ignored"], 1)
            self.assertTrue(temp_xlsx.exists())

    def test_png_decision_converts_to_pdf(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Introduction"
            subject.mkdir(parents=True)
            png = subject / "2_Capture d'écran carnet des écarts.png"
            png.write_bytes(b"not-an-image-for-dry-run")
            rows, summary = build_manifest(source, target, read_msg_metadata=False)
            row = rows[0]
            self.assertEqual(row.role, "image_png")
            self.assertEqual(row.decision, "convertir_png_pdf")
            self.assertEqual(row.conversion_status, "dry_run_non_converti")
            self.assertEqual(summary["png_to_convert"], 1)
            self.assertTrue(row.target_name.endswith(".pdf"))
            self.assertTrue(png.exists())

    def test_png_target_name(self):
        root = Path("root")
        png = source_item(root / "Réserve 5584" / "1_dimension porte restaurant avant travaux .png", root)
        self.assertEqual(
            target_name_for_png(png),
            "Piece 1 - Reserve 5584 - dimension porte restaurant avant travaux.pdf",
        )

    def test_dwg_reference_pdf_metadata_and_target(self):
        root = Path("root")
        dwg = source_item(root / "Réserve 5584" / "1_RESTAURANT ind. I du 13_10_2018.dwg", root)
        lines = dwg_reference_pdf_lines(dwg)
        self.assertIn("Document source non visualisé dans le présent PDF", lines)
        self.assertIn("Fichier original : 1_RESTAURANT ind. I du 13_10_2018.dwg", lines)
        self.assertIn("Format original : DWG", lines)
        self.assertIn("Sujet : Réserve 5584", lines)
        self.assertIn("Numéro de pièce de la partie : 1", lines)
        self.assertTrue(any("conserve intact" in line for line in lines))
        self.assertTrue(any("artefact de référencement" in line for line in lines))
        self.assertTrue(any("ne constitue pas une representation graphique" in line for line in lines))
        self.assertEqual(
            target_name_for_dwg_reference(dwg),
            "Piece 1 - Reserve 5584 - 1 RESTAURANT ind. I du 13 10 2018.dwg.pdf",
        )

    def test_dwg_is_not_ingested_directly(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Réserve 5584"
            subject.mkdir(parents=True)
            dwg = subject / "1_RESTAURANT ind. I du 13_10_2018.dwg"
            dwg.write_bytes(b"dwg")
            rows, summary = build_manifest(source, target, read_msg_metadata=False)
            row = rows[0]
            self.assertEqual(row.role, "dwg_reference")
            self.assertEqual(row.decision, "creer_pdf_reference_dwg")
            self.assertEqual(row.conversion_status, "dwg_reference_pdf")
            self.assertTrue(row.target_name.endswith(".dwg.pdf"))
            self.assertEqual(summary["dwg_reference_pdf"], 1)
            self.assertTrue(dwg.exists())

    def test_png_dwg_possible_relation_is_only_a_note(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            subject = source / "Réserve 5584"
            subject.mkdir(parents=True)
            (subject / "1_dimension porte restaurant avant travaux .png").write_bytes(b"png")
            (subject / "1_RESTAURANT ind. I du 13_10_2018.dwg").write_bytes(b"dwg")
            rows, summary = build_manifest(source, target, read_msg_metadata=False)
            by_ext = {row.source_extension: row for row in rows}
            self.assertEqual(by_ext[".png"].note, "possible_relation_with_dwg")
            self.assertNotEqual(by_ext[".png"].target_name, by_ext[".dwg"].target_name)
            self.assertEqual(summary["pdf_presented_to_app"], 2)

    def test_collision_detection(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            (source / "Introduction").mkdir(parents=True)
            (source / "Introduction" / "5_Test.msg").write_bytes(b"msg")
            (source / "Introduction" / "5_Test_.msg").write_bytes(b"msg")
            rows, _ = build_manifest(source, target, read_msg_metadata=False)
            collisions = detect_target_collisions(rows)
            self.assertEqual(len(collisions), 1)
            self.assertTrue(all(row.decision == "collision_a_arbitrer" for row in rows))

    def test_script_does_not_generate_expert_doc_id(self):
        with tempfile.TemporaryDirectory(dir=r"C:\CodexWorkspace") as tmp:
            source = Path(tmp) / "source"
            target = Path(tmp) / "target"
            (source / "Introduction").mkdir(parents=True)
            (source / "Introduction" / "5_SEBIA - Suivi.msg").write_bytes(b"msg")
            rows, _ = build_manifest(source, target, read_msg_metadata=False)
            payload = "\n".join(str(row.__dict__) for row in rows)
            self.assertNotIn("expert_doc_id", payload)


if __name__ == "__main__":
    unittest.main()
