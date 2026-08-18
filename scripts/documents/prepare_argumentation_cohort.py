from __future__ import annotations

import argparse
import csv
import gc
import hashlib
import html
import json
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path


PDF_SIGNATURE = b"%PDF-"
MANIFEST_JSON = "prepare_argumentation_cohort.json"
MANIFEST_CSV = "prepare_argumentation_cohort.csv"
MSG_CONVERSION_METHOD = "outlook_com_html_word_export_pdf"
MSG_DRY_RUN_METHOD = "outlook_com_saveas_docx_word_export_pdf_preconise"
PNG_CONVERSION_METHOD = "word_com_image_to_pdf"
DWG_CONVERSION_METHOD = "word_com_reference_pdf"


@dataclass
class SourceItem:
    path: Path
    subject_folder: str
    numero_piece_partie: str
    piece_suffix: str
    piece_ref: str
    label: str
    extension: str


@dataclass
class ManifestRow:
    subject_folder: str
    numero_piece_partie: str
    piece_suffix: str
    source_path: str
    source_name: str
    source_extension: str
    role: str
    target_name: str = ""
    target_path: str = ""
    decision: str = ""
    reason: str = ""
    source_msg: str = ""
    attachments_count: str = ""
    attachments_names: str = ""
    attachment_match_status: str = ""
    conversion_method: str = ""
    conversion_status: str = ""
    related_pdf: str = ""
    note: str = ""
    sha256_source: str = ""
    sha256_target: str = ""
    attachment_matches: list[dict] = field(default_factory=list)


@dataclass
class MsgMetadata:
    sender: str = ""
    to: str = ""
    cc: str = ""
    sent_on: str = ""
    subject: str = ""
    body: str = ""
    html_body: str = ""
    attachments: list[str] = field(default_factory=list)
    error: str = ""


def coinitialize_com():
    try:
        import pythoncom  # type: ignore
    except Exception:
        return None
    try:
        pythoncom.CoInitialize()
    except Exception:
        return None
    return pythoncom


def couninitialize_com(pythoncom_module) -> None:
    if pythoncom_module is None:
        return
    try:
        pythoncom_module.CoUninitialize()
    except Exception:
        pass


def close_word_document(doc) -> None:
    if doc is None:
        return
    try:
        doc.Close(False)
    except Exception:
        pass


def quit_word_application(word) -> None:
    if word is None:
        return
    try:
        word.Quit()
    except Exception:
        pass


def release_word_com(word, doc=None, *extra_refs, pythoncom_module=None) -> None:
    close_word_document(doc)
    quit_word_application(word)
    for ref in extra_refs:
        del ref
    del doc
    del word
    gc.collect()
    couninitialize_com(pythoncom_module)


def compact_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def strip_diacritics(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_subject(value: str) -> str:
    return compact_spaces(strip_diacritics(value))


def sanitize_filename_part(value: str, max_len: int = 140) -> str:
    text = compact_spaces(str(value or "").replace("_", " "))
    text = re.sub(r'[<>:"/\\|?*]+', " ", text)
    text = compact_spaces(text).strip(" .-_")
    if len(text) > max_len:
        text = text[:max_len].rstrip(" .-_")
    return text or "document"


def parse_piece_token(filename: str) -> tuple[str, str, str, str]:
    stem = Path(filename or "").stem
    match = re.match(r"^\s*(?P<num>\d+)(?P<suffix>[A-Za-z]?)\s*[_\-\s]+(?P<label>.+)$", stem)
    if not match:
        return "", "", "", compact_spaces(stem)
    numero = match.group("num")
    suffix = match.group("suffix").lower()
    piece_ref = f"{numero}.{suffix}" if suffix else numero
    return numero, suffix, piece_ref, compact_spaces(match.group("label"))


def normalized_piece_key(numero: str, suffix: str = "") -> str:
    number = str(numero or "").lstrip("0") or "0"
    suffix = str(suffix or "").lower()
    return f"{number}.{suffix}" if suffix else number


def source_item(path: Path, source_root: Path) -> SourceItem:
    subject = path.parent.name if path.parent != source_root else ""
    numero, suffix, piece_ref, label = parse_piece_token(path.name)
    return SourceItem(
        path=path,
        subject_folder=subject,
        numero_piece_partie=numero,
        piece_suffix=suffix,
        piece_ref=piece_ref,
        label=label,
        extension=path.suffix.lower(),
    )


def target_name_for_msg(item: SourceItem) -> str:
    subject = normalize_subject(item.subject_folder)
    ref = item.piece_ref or "non numerotee"
    return f"Piece {ref} - {subject} - {sanitize_filename_part(item.label)}.pdf"


def target_name_for_pdf(item: SourceItem, has_msg_master: bool) -> str:
    subject = normalize_subject(item.subject_folder)
    ref = item.piece_ref or "non numerotee"
    label = sanitize_filename_part(item.label)
    if has_msg_master:
        return f"Piece {ref} - {subject} - Annexe - {label}.pdf"
    return f"Piece {ref} - {subject} - {label}.pdf"


def target_name_for_png(item: SourceItem) -> str:
    subject = normalize_subject(item.subject_folder)
    ref = item.piece_ref or "non numerotee"
    return f"Piece {ref} - {subject} - {sanitize_filename_part(item.label)}.pdf"


def target_name_for_dwg_reference(item: SourceItem) -> str:
    subject = normalize_subject(item.subject_folder)
    ref = item.piece_ref or "non numerotee"
    label = sanitize_filename_part(item.path.name)
    return f"Piece {ref} - {subject} - {label}.pdf"


def match_normalized_key(name: str) -> str:
    stem = Path(name or "").stem
    text = strip_diacritics(stem).lower()
    text = text.replace(".xlsx -", " ").replace(".xls -", " ")
    text = re.sub(r"\b(xlsx|xls|pdf|msg)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return compact_spaces(text)


def equivalence_tokens(name: str) -> set[str]:
    tokens = set(match_normalized_key(name).split())
    return {token for token in tokens if len(token) >= 3 and not token.isdigit()}


def find_xlsx_pdf_equivalent(xlsx: SourceItem, pdf_items: list[SourceItem]) -> tuple[str, str, str]:
    xkey = match_normalized_key(xlsx.path.name)
    xpiece = normalized_piece_key(xlsx.numero_piece_partie, xlsx.piece_suffix)
    xtokens = equivalence_tokens(xlsx.path.name)
    same_subject = [
        item for item in pdf_items
        if item.subject_folder == xlsx.subject_folder
        and normalized_piece_key(item.numero_piece_partie, item.piece_suffix) == xpiece
    ]
    exact_candidates = []
    prefix_candidates = []
    token_candidates = []
    for pdf in same_subject:
        pkey = match_normalized_key(pdf.path.name)
        ptokens = equivalence_tokens(pdf.path.name)
        overlap = xtokens & ptokens
        token_match = bool(xtokens and ptokens and (
            xtokens <= ptokens
            or ptokens <= xtokens
            or len(overlap) / max(len(xtokens), len(ptokens)) >= 0.65
        ))
        if pkey == xkey:
            exact_candidates.append(pdf)
        elif pkey.startswith(xkey) or xkey.startswith(pkey):
            prefix_candidates.append(pdf)
        elif token_match:
            token_candidates.append(pdf)
    candidates = exact_candidates or prefix_candidates or token_candidates
    if len(candidates) == 1:
        pdf = candidates[0]
        return "ignorer_ingestion_pdf_equivalent", f"PDF equivalent probable : {pdf.path.name}", str(pdf.path)
    if len(candidates) > 1:
        names = ", ".join(item.path.name for item in candidates[:5])
        return "a_arbitrer", f"plusieurs PDF equivalents possibles : {names}", ""
    return "a_arbitrer", "aucun PDF equivalent fiable dans le meme sujet", ""


def is_possible_png_dwg_relation(item: SourceItem) -> bool:
    subject = normalize_subject(item.subject_folder).casefold()
    label = strip_diacritics(item.label).casefold()
    return item.extension == ".png" and subject == "reserve 5584" and "dimension porte restaurant" in label


def is_office_temp_file(path: Path) -> bool:
    return path.name.startswith("~$")


def dwg_reference_pdf_lines(item: SourceItem) -> list[str]:
    return [
        "Document source non visualisé dans le présent PDF",
        f"Fichier original : {item.path.name}",
        "Format original : DWG",
        f"Sujet : {item.subject_folder}",
        f"Numéro de pièce de la partie : {item.numero_piece_partie or 'non numérotée'}",
        "Le fichier DWG original est conserve intact dans le depot de la partie.",
        "Ce PDF est uniquement un artefact de référencement destiné à l'ingestion documentaire.",
        "Il ne constitue pas une representation graphique du fichier DWG.",
    ]


def write_dwg_reference_pdf_atomic(item: SourceItem, target_pdf: Path) -> None:
    target_pdf.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=target_pdf.name + ".", suffix=".tmp.pdf", dir=str(target_pdf.parent), delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        export_text_to_pdf_with_word("\n".join(dwg_reference_pdf_lines(item)), tmp_path)
        verify_pdf(tmp_path)
        os.replace(str(tmp_path), str(target_pdf))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    verify_pdf(target_pdf)


def convert_png_to_pdf_atomic(source_png: Path, target_pdf: Path) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"pywin32_indisponible: {exc}") from exc

    target_pdf.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=target_pdf.name + ".", suffix=".tmp.pdf", dir=str(target_pdf.parent), delete=False) as tmp:
        tmp_path = Path(tmp.name)
    pythoncom_module = coinitialize_com()
    word = None
    doc = None
    paragraph = None
    shape = None
    page_setup = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Add()
        paragraph = doc.Paragraphs(1)
        paragraph.Alignment = 1
        shape = doc.InlineShapes.AddPicture(FileName=str(source_png), LinkToFile=False, SaveWithDocument=True)
        page_setup = doc.PageSetup
        usable_width = float(page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin)
        usable_height = float(page_setup.PageHeight - page_setup.TopMargin - page_setup.BottomMargin)
        width = float(shape.Width)
        height = float(shape.Height)
        if width > 0 and height > 0:
            scale = min(usable_width / width, usable_height / height, 1.0)
            shape.Width = width * scale
            shape.Height = height * scale
        doc.ExportAsFixedFormat(str(tmp_path), 17)
        verify_pdf(tmp_path)
        os.replace(str(tmp_path), str(target_pdf))
    finally:
        release_word_com(word, doc, paragraph, shape, page_setup, pythoncom_module=pythoncom_module)
        paragraph = None
        shape = None
        page_setup = None
        doc = None
        word = None
        gc.collect()
        if tmp_path.exists():
            tmp_path.unlink()
    verify_pdf(target_pdf)


def detect_target_collisions(rows: list[ManifestRow]) -> dict[str, list[ManifestRow]]:
    groups: dict[str, list[ManifestRow]] = defaultdict(list)
    for row in rows:
        if row.target_name:
            groups[row.target_name.casefold()].append(row)
    return {key: value for key, value in groups.items() if len(value) > 1}


def match_attachments_to_pdfs(attachments: list[str], pdf_items: list[SourceItem]) -> tuple[str, list[dict]]:
    if not attachments:
        return "attachments_none", []
    pdf_by_exact = defaultdict(list)
    pdf_by_norm = defaultdict(list)
    for item in pdf_items:
        pdf_by_exact[item.path.name.casefold()].append(item)
        pdf_by_norm[match_normalized_key(item.path.name)].append(item)
    results = []
    statuses = []
    for attachment in attachments:
        exact = pdf_by_exact.get(str(attachment).casefold(), [])
        normalized = pdf_by_norm.get(match_normalized_key(attachment), [])
        if len(exact) == 1:
            status = "attachment_match_exact"
            matched = [str(exact[0].path)]
        elif len(exact) > 1:
            status = "attachment_ambiguous"
            matched = [str(item.path) for item in exact]
        elif len(normalized) == 1:
            status = "attachment_match_normalized"
            matched = [str(normalized[0].path)]
        elif len(normalized) > 1:
            status = "attachment_ambiguous"
            matched = [str(item.path) for item in normalized]
        else:
            status = "attachment_not_found"
            matched = []
        statuses.append(status)
        results.append({"attachment": attachment, "status": status, "matched_paths": matched})
    if "attachment_ambiguous" in statuses:
        return "attachment_ambiguous", results
    if "attachment_not_found" in statuses:
        return "attachment_not_found", results
    if "attachment_match_normalized" in statuses:
        return "attachment_match_normalized", results
    return "attachment_match_exact", results


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_msg_metadata_outlook(path: Path) -> MsgMetadata:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        return MsgMetadata(error=f"pywin32_indisponible: {exc}")
    outlook = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        try:
            return read_msg_metadata_from_outlook_session(outlook, path)
        except Exception as direct_exc:
            with tempfile.TemporaryDirectory(prefix="msg_open_") as tmp:
                tmp_msg = Path(tmp) / "source.msg"
                shutil.copy2(path, tmp_msg)
                try:
                    return read_msg_metadata_from_outlook_session(outlook, tmp_msg)
                except Exception as tmp_exc:
                    return MsgMetadata(error=f"lecture_msg_echec: {direct_exc}; fallback_temp_echec: {tmp_exc}")
    finally:
        del outlook


def read_msg_metadata_from_outlook_session(outlook, path: Path) -> MsgMetadata:
    msg = None
    try:
        msg = outlook.Session.OpenSharedItem(str(path))
        attachments = []
        try:
            for idx in range(1, int(msg.Attachments.Count) + 1):
                attachments.append(str(msg.Attachments.Item(idx).FileName or ""))
        except Exception:
            attachments = []
        return MsgMetadata(
            sender=str(getattr(msg, "SenderName", "") or getattr(msg, "SenderEmailAddress", "") or ""),
            to=str(getattr(msg, "To", "") or ""),
            cc=str(getattr(msg, "CC", "") or ""),
            sent_on=str(getattr(msg, "SentOn", "") or ""),
            subject=str(getattr(msg, "Subject", "") or ""),
            body=str(getattr(msg, "Body", "") or ""),
            html_body=str(getattr(msg, "HTMLBody", "") or ""),
            attachments=attachments,
        )
    finally:
        try:
            if msg is not None:
                msg.Close(1)
        except Exception:
            pass
        del msg


def email_html(metadata: MsgMetadata, source_msg: Path) -> str:
    attachments = metadata.attachments or []
    attachment_html = "".join(f"<li>{html.escape(name)}</li>" for name in attachments) or "<li>(aucune)</li>"
    body = f"<pre>{html.escape(metadata.body or '')}</pre>"
    return f"""<!doctype html>
<html>
<head><meta charset="utf-8"><style>
body {{ font-family: Arial, sans-serif; font-size: 10.5pt; }}
.meta {{ border-collapse: collapse; margin-bottom: 1em; }}
.meta th {{ text-align: left; padding: 3px 8px 3px 0; white-space: nowrap; }}
.meta td {{ padding: 3px 0; }}
h1 {{ font-size: 15pt; }}
</style></head>
<body>
<h1>{html.escape(metadata.subject or source_msg.stem)}</h1>
<table class="meta">
<tr><th>Source MSG</th><td>{html.escape(str(source_msg))}</td></tr>
<tr><th>Expéditeur</th><td>{html.escape(metadata.sender)}</td></tr>
<tr><th>Destinataires</th><td>{html.escape(metadata.to)}</td></tr>
<tr><th>CC</th><td>{html.escape(metadata.cc)}</td></tr>
<tr><th>Date/heure</th><td>{html.escape(metadata.sent_on)}</td></tr>
<tr><th>Objet</th><td>{html.escape(metadata.subject)}</td></tr>
</table>
<h2>Message</h2>
{body}
<h2>Pièces jointes</h2>
<ul>{attachment_html}</ul>
</body></html>"""


def email_text(metadata: MsgMetadata, source_msg: Path) -> str:
    attachments = metadata.attachments or []
    attachment_lines = "\n".join(f"- {name}" for name in attachments) or "- (aucune)"
    return "\n".join(
        [
            metadata.subject or source_msg.stem,
            "",
            f"Source MSG : {source_msg}",
            f"Expediteur : {metadata.sender}",
            f"Destinataires : {metadata.to}",
            f"CC : {metadata.cc}",
            f"Date/heure : {metadata.sent_on}",
            f"Objet : {metadata.subject}",
            "",
            "Message",
            "",
            metadata.body or "",
            "",
            "Pieces jointes",
            attachment_lines,
        ]
    )


def export_html_to_pdf_with_word(html_path: Path, pdf_path: Path) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"pywin32_indisponible: {exc}") from exc
    pythoncom_module = coinitialize_com()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(FileName=str(html_path), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(OutputFileName=str(pdf_path), ExportFormat=17)
        except AttributeError:
            doc.SaveAs2(FileName=str(pdf_path), FileFormat=17)
    finally:
        release_word_com(word, doc, pythoncom_module=pythoncom_module)
        doc = None
        word = None
        gc.collect()


def export_text_to_pdf_with_word(text: str, pdf_path: Path) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError(f"pywin32_indisponible: {exc}") from exc
    pythoncom_module = coinitialize_com()
    word = None
    doc = None
    doc_range = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Add()
        doc_range = doc.Range()
        doc_range.Text = text
        try:
            doc.ExportAsFixedFormat(OutputFileName=str(pdf_path), ExportFormat=17)
        except AttributeError:
            doc.SaveAs2(FileName=str(pdf_path), FileFormat=17)
    finally:
        release_word_com(word, doc, doc_range, pythoncom_module=pythoncom_module)
        doc_range = None
        doc = None
        word = None
        gc.collect()


def convert_msg_to_pdf_atomic(source_msg: Path, target_pdf: Path) -> tuple[MsgMetadata, str]:
    metadata = read_msg_metadata_outlook(source_msg)
    if metadata.error:
        return metadata, "conversion_msg_echec"
    target_pdf.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="msg_to_pdf_") as tmp:
        tmp_dir = Path(tmp)
        pdf_tmp = tmp_dir / "message.pdf"
        export_text_to_pdf_with_word(email_text(metadata, source_msg), pdf_tmp)
        verify_pdf(pdf_tmp)
        os.replace(str(pdf_tmp), str(target_pdf))
    verify_pdf(target_pdf)
    return metadata, "conversion_msg_ok"


def verify_pdf(path: Path) -> None:
    if not path.exists() or path.stat().st_size <= 0:
        raise RuntimeError(f"PDF absent ou vide : {path}")
    with path.open("rb") as handle:
        if handle.read(5) != PDF_SIGNATURE:
            raise RuntimeError(f"Signature PDF invalide : {path}")


def copy_file_atomic(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=target.name + ".", suffix=".tmp", dir=str(target.parent), delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        shutil.copy2(source, tmp_path)
        if tmp_path.stat().st_size <= 0:
            raise RuntimeError(f"copie vide : {source}")
        os.replace(str(tmp_path), str(target))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def build_manifest(
    source: Path,
    target: Path,
    *,
    read_msg_metadata: bool = True,
    compute_hashes: bool = False,
) -> tuple[list[ManifestRow], dict]:
    files = sorted([p for p in source.rglob("*") if p.is_file()], key=lambda p: str(p).casefold())
    items = [source_item(path, source) for path in files]
    pdf_items = [item for item in items if item.extension == ".pdf"]
    msg_items = [item for item in items if item.extension == ".msg"]
    msg_keys = {
        (item.subject_folder, normalized_piece_key(item.numero_piece_partie, item.piece_suffix))
        for item in msg_items
    }
    msg_meta_by_path: dict[Path, MsgMetadata] = {}
    if read_msg_metadata:
        for item in msg_items:
            msg_meta_by_path[item.path] = read_msg_metadata_outlook(item.path)

    rows: list[ManifestRow] = []
    for item in items:
        target_name = ""
        decision = ""
        reason = ""
        role = ""
        source_msg = ""
        conversion_method = ""
        conversion_status = ""
        related_pdf = ""
        note = ""
        attachment_status = ""
        attachment_matches: list[dict] = []
        attachments: list[str] = []

        if is_office_temp_file(item.path):
            role = "office_temp"
            decision = "ignorer_fichier_temporaire_office"
            reason = "fichier temporaire Office, non destine a l'ingestion documentaire"
        elif item.extension == ".msg":
            role = "msg_maitre"
            decision = "convertir_msg_pdf"
            target_name = target_name_for_msg(item)
            conversion_method = MSG_DRY_RUN_METHOD
            conversion_status = "dry_run_non_converti"
            metadata = msg_meta_by_path.get(item.path, MsgMetadata(error="metadata_non_lue_dry_run"))
            attachments = metadata.attachments or []
            if metadata.error:
                attachment_status = "attachments_non_lus"
                reason = metadata.error
            else:
                same_subject_pdfs = [pdf for pdf in pdf_items if pdf.subject_folder == item.subject_folder]
                attachment_status, attachment_matches = match_attachments_to_pdfs(attachments, same_subject_pdfs)
                reason = "message MSG converti en PDF de consultation en mode execute"
        elif item.extension == ".pdf":
            key = (item.subject_folder, normalized_piece_key(item.numero_piece_partie, item.piece_suffix))
            has_msg_master = key in msg_keys
            role = "annexe_pdf" if has_msg_master else "pdf"
            decision = "copier"
            target_name = target_name_for_pdf(item, has_msg_master)
            source_msg = "msg_maitre_meme_sujet_numero" if has_msg_master else ""
            reason = "PDF annexe car MSG maitre de meme sujet/numero" if has_msg_master else "PDF principal sans MSG maitre correspondant"
        elif item.extension in {".xlsx", ".xls"}:
            role = "xlsx"
            decision, reason, related_pdf = find_xlsx_pdf_equivalent(item, pdf_items)
        elif item.extension == ".png":
            role = "image_png"
            decision = "convertir_png_pdf"
            target_name = target_name_for_png(item)
            conversion_method = PNG_CONVERSION_METHOD
            conversion_status = "dry_run_non_converti"
            reason = "image PNG convertie en PDF en mode execute"
            if is_possible_png_dwg_relation(item):
                note = "possible_relation_with_dwg"
        elif item.extension == ".dwg":
            role = "dwg_reference"
            decision = "creer_pdf_reference_dwg"
            target_name = target_name_for_dwg_reference(item)
            conversion_method = DWG_CONVERSION_METHOD
            conversion_status = "dwg_reference_pdf"
            reason = "PDF de substitution documentaire cree en mode execute, sans conversion graphique du DWG"
        else:
            role = item.extension.lstrip(".") or "autre"
            decision = "a_arbitrer"
            reason = "extension non prise en charge automatiquement dans cette version"

        target_path = str(target / target_name) if target_name else ""
        source_hash = ""
        if compute_hashes and decision != "ignorer_fichier_temporaire_office":
            source_hash = sha256_file(item.path)
        row = ManifestRow(
            subject_folder=item.subject_folder,
            numero_piece_partie=item.numero_piece_partie,
            piece_suffix=item.piece_suffix,
            source_path=str(item.path),
            source_name=item.path.name,
            source_extension=item.extension,
            role=role,
            target_name=target_name,
            target_path=target_path,
            decision=decision,
            reason=reason,
            source_msg=source_msg,
            attachments_count=str(len(attachments)) if item.extension == ".msg" else "",
            attachments_names=" | ".join(attachments) if item.extension == ".msg" else "",
            attachment_match_status=attachment_status,
            conversion_method=conversion_method,
            conversion_status=conversion_status,
            related_pdf=related_pdf,
            note=note,
            sha256_source=source_hash,
            attachment_matches=attachment_matches,
        )
        rows.append(row)

    collisions = detect_target_collisions(rows)
    for collision_rows in collisions.values():
        for row in collision_rows:
            row.decision = "collision_a_arbitrer"
            row.reason = f"collision de nom cible : {row.target_name}"

    summary = summarize(rows, items, collisions)
    return rows, summary


def summarize(rows: list[ManifestRow], items: list[SourceItem], collisions: dict[str, list[ManifestRow]]) -> dict:
    ext_counts = Counter(item.extension or "<sans_ext>" for item in items)
    attachment_counts = Counter(row.attachment_match_status for row in rows if row.attachment_match_status)
    pdf_outputs = [
        row for row in rows
        if row.target_name and row.decision in {"copier", "convertir_msg_pdf", "convertir_png_pdf", "creer_pdf_reference_dwg"}
    ]
    return {
        "subjects_count": len({item.subject_folder for item in items}),
        "files_total": len(items),
        "extensions": dict(sorted(ext_counts.items())),
        "msg_to_convert": sum(1 for row in rows if row.role == "msg_maitre" and row.decision == "convertir_msg_pdf"),
        "pdf_to_copy": sum(1 for row in rows if row.source_extension == ".pdf" and row.decision == "copier"),
        "xlsx_ignored_pdf_equivalent": sum(1 for row in rows if row.decision == "ignorer_ingestion_pdf_equivalent"),
        "office_temp_ignored": sum(1 for row in rows if row.decision == "ignorer_fichier_temporaire_office"),
        "png_to_convert": sum(1 for row in rows if row.decision == "convertir_png_pdf"),
        "dwg_reference_pdf": sum(1 for row in rows if row.decision == "creer_pdf_reference_dwg"),
        "to_arbitrate": sum(1 for row in rows if row.decision in {"a_arbitrer", "collision_a_arbitrer"}),
        "attachment_statuses": dict(sorted(attachment_counts.items())),
        "collisions": len(collisions),
        "pdf_presented_to_app": len(pdf_outputs),
    }


def row_to_dict(row: ManifestRow) -> dict:
    data = asdict(row)
    data["attachment_matches"] = json.dumps(data["attachment_matches"], ensure_ascii=False)
    return data


def write_manifest_atomic(rows: list[ManifestRow], target: Path) -> tuple[Path, Path]:
    target.mkdir(parents=True, exist_ok=True)
    json_path = target / MANIFEST_JSON
    csv_path = target / MANIFEST_CSV
    payload = [row_to_dict(row) for row in rows]
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".json.tmp", dir=str(target), delete=False) as tmp:
        tmp_json = Path(tmp.name)
        json.dump(payload, tmp, ensure_ascii=False, indent=2)
    os.replace(str(tmp_json), str(json_path))

    fieldnames = list(payload[0].keys()) if payload else [field.name for field in ManifestRow.__dataclass_fields__.values()]
    with tempfile.NamedTemporaryFile("w", encoding="utf-8-sig", newline="", suffix=".csv.tmp", dir=str(target), delete=False) as tmp:
        tmp_csv = Path(tmp.name)
        writer = csv.DictWriter(tmp, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        for row in payload:
            writer.writerow(row)
    os.replace(str(tmp_csv), str(csv_path))
    return json_path, csv_path


def execute(rows: list[ManifestRow], target: Path) -> list[ManifestRow]:
    unresolved = [row for row in rows if row.decision == "collision_a_arbitrer"]
    existing = [row for row in rows if row.target_path and Path(row.target_path).exists()]
    if unresolved:
        raise RuntimeError(f"collisions non resolues : {len(unresolved)}")
    if existing:
        names = ", ".join(Path(row.target_path).name for row in existing[:5])
        raise RuntimeError(f"fichiers cibles existants, arrêt avant copie : {names}")

    for row in rows:
        source = Path(row.source_path)
        target_path = Path(row.target_path) if row.target_path else None
        if row.decision == "ignorer_fichier_temporaire_office":
            continue
        row.sha256_source = sha256_file(source)
        if row.decision == "copier" and target_path is not None:
            copy_file_atomic(source, target_path)
            if row.source_extension == ".pdf":
                verify_pdf(target_path)
            row.sha256_target = sha256_file(target_path)
        elif row.decision == "convertir_msg_pdf" and target_path is not None:
            try:
                metadata, status = convert_msg_to_pdf_atomic(source, target_path)
            except Exception as exc:
                row.decision = "conversion_msg_echec"
                row.conversion_status = "conversion_msg_echec"
                row.reason = str(exc)
                continue
            row.conversion_method = MSG_CONVERSION_METHOD
            row.conversion_status = status
            row.attachments_count = str(len(metadata.attachments))
            row.attachments_names = " | ".join(metadata.attachments)
            if status != "conversion_msg_ok":
                row.decision = "conversion_msg_echec"
                row.reason = metadata.error or "conversion MSG echouee"
            else:
                row.sha256_target = sha256_file(target_path)
        elif row.decision == "convertir_png_pdf" and target_path is not None:
            try:
                convert_png_to_pdf_atomic(source, target_path)
            except Exception as exc:
                row.decision = "conversion_png_echec"
                row.conversion_status = "conversion_png_echec"
                row.reason = str(exc)
            else:
                row.conversion_status = "conversion_png_ok"
                row.sha256_target = sha256_file(target_path)
        elif row.decision == "creer_pdf_reference_dwg" and target_path is not None:
            item = SourceItem(
                path=source,
                subject_folder=row.subject_folder,
                numero_piece_partie=row.numero_piece_partie,
                piece_suffix=row.piece_suffix,
                piece_ref=row.numero_piece_partie + (f".{row.piece_suffix}" if row.piece_suffix else ""),
                label=parse_piece_token(source.name)[3],
                extension=row.source_extension,
            )
            write_dwg_reference_pdf_atomic(item, target_path)
            row.conversion_status = "dwg_reference_pdf"
            row.sha256_target = sha256_file(target_path)
    write_manifest_atomic(rows, target)
    return rows


def print_summary(summary: dict, rows: list[ManifestRow], *, execute_mode: bool) -> None:
    mode = "EXECUTE" if execute_mode else "DRY-RUN"
    print(f"Mode : {mode}")
    print(f"Sujets : {summary['subjects_count']}")
    print(f"Fichiers totaux : {summary['files_total']}")
    print("Extensions :")
    for ext, count in summary["extensions"].items():
        print(f"  {ext}: {count}")
    print(f"MSG a convertir : {summary['msg_to_convert']}")
    print(f"PDF existants a copier : {summary['pdf_to_copy']}")
    print(f"XLSX ignores avec PDF equivalent : {summary['xlsx_ignored_pdf_equivalent']}")
    print(f"Fichiers temporaires Office ignores : {summary['office_temp_ignored']}")
    print(f"PNG a convertir en PDF : {summary['png_to_convert']}")
    print(f"DWG donnant lieu a PDF de substitution : {summary['dwg_reference_pdf']}")
    print(f"Fichiers a arbitrer : {summary['to_arbitrate']}")
    print(f"Attachments status : {summary['attachment_statuses']}")
    print(f"Collisions : {summary['collisions']}")
    print(f"PDF presentes ensuite a app.py : {summary['pdf_presented_to_app']}")
    print("")
    print("Cas a arbitrer :")
    arbitrate = [row for row in rows if row.decision in {"a_arbitrer", "collision_a_arbitrer"}]
    if not arbitrate:
        print("  aucun")
    for row in arbitrate:
        print(f"  - {row.subject_folder} | {row.source_name} | {row.decision} | {row.reason}")
    print("")
    print("Attachments problematiques :")
    problematic = [
        row for row in rows
        if row.attachment_match_status in {"attachment_not_found", "attachment_ambiguous", "attachments_non_lus"}
    ]
    if not problematic:
        print("  aucun")
    for row in problematic:
        print(f"  - {row.subject_folder} | {row.source_name} | {row.attachment_match_status} | {row.reason}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Prepare une cohorte plate PDF pour l'Argumentation SEBIA.")
    parser.add_argument("--source", required=True, help="Dossier source contenant les sous-dossiers sujets.")
    parser.add_argument("--target", required=True, help="Dossier cible de cohorte preparee.")
    parser.add_argument("--execute", action="store_true", help="Effectue les copies/conversions et ecrit les manifestes.")
    parser.add_argument(
        "--skip-msg-metadata",
        action="store_true",
        help="Ne lit pas les metadonnees MSG via Outlook pendant le dry-run.",
    )
    parser.add_argument(
        "--hash-dry-run",
        action="store_true",
        help="Calcule sha256_source meme en dry-run.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    source = Path(args.source)
    target = Path(args.target)
    if not source.exists() or not source.is_dir():
        print(f"Source introuvable ou non dossier : {source}", file=sys.stderr)
        return 2
    rows, summary = build_manifest(
        source,
        target,
        read_msg_metadata=not args.skip_msg_metadata,
        compute_hashes=bool(args.execute or args.hash_dry_run),
    )
    if args.execute:
        rows = execute(rows, target)
        collisions = detect_target_collisions(rows)
        source_items = [source_item(Path(row.source_path), source) for row in rows]
        summary = summarize(rows, source_items, collisions)
        print(f"Manifestes ecrits : {target / MANIFEST_JSON} ; {target / MANIFEST_CSV}")
    print_summary(summary, rows, execute_mode=bool(args.execute))
    if not args.execute:
        print("")
        print("Dry-run uniquement : aucun fichier cree dans la cible.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
