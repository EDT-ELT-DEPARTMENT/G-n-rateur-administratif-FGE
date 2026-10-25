# -*- coding: utf-8 -*-
"""Workflow de demandes et de génération d'attestations administratives UDL-SBA."""
from __future__ import annotations

from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from zoneinfo import ZoneInfo
import hashlib
import hmac
import os
import re
import smtplib
import ssl
import unicodedata

import pandas as pd
import streamlit as st

from attestation_generator import DEPARTEMENTS, generate_attestation
from storage import (
    backend_mode, create_request, get_request, get_setting, issue_request, list_requests,
    save_generation_file, list_generation_files, get_generation_file, delete_all_request_history,
)
from edt_individuel import (
    export_timetable_html, export_timetable_pdf, export_timetable_xlsx,
    load_edt_sessions, match_teacher_sessions, timetable_grid_records,
    unique_subjects, unique_teacher_names,
)

APP_DIR = Path(__file__).resolve().parent
TZ_ALGERIA = ZoneInfo("Africa/Algiers")
LANGUES = ["Bilingue (arabe puis français)", "Arabe (conforme au modèle PDF)", "Français"]
REFERENCE_PAR_DEFAUT = f"UDL-GEL-ER-120-{datetime.now(TZ_ALGERIA).year}"
SOURCE_PERSONNEL_FILE = APP_DIR / "Vos coordonnées (réponses).xlsx"
EDT_SOURCE_FILE = APP_DIR / "EDT_Edition.xlsx"
ALL_SOURCE_REQUESTERS = "__TOUS_LES_DEMANDEURS__"

# Valeurs de repli demandées. Pour les autres départements, les noms peuvent être
# définis dans Streamlit Secrets (CHEF_NAME_FR_... et CHEF_NAME_AR_...).
CHEFS_PAR_DEFAUT = {
    "Electrotechnique": {"fr": "MILOUA FARID", "ar": "ميلوى فريد"},
    "Automatique": {"fr": "", "ar": ""},
    "Electronique": {"fr": "", "ar": ""},
    "Télécommunications": {"fr": "", "ar": ""},
}

DEPT_SECRET_SUFFIX = {
    "Electrotechnique": "ELECTROTECHNIQUE",
    "Automatique": "AUTOMATIQUE",
    "Electronique": "ELECTRONIQUE",
    "Télécommunications": "TELECOMMUNICATIONS",
}

st.set_page_config(
    page_title="Demandes d'attestations administratives — UDL-SBA",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {max-width: 1250px; padding-top: 1.2rem; padding-bottom: 2rem;}
.hero {padding: 1.4rem 1.6rem; border-radius: 16px; color: white; background: linear-gradient(115deg,#0f172a,#1d4ed8 70%,#0e7490); margin-bottom: 1rem;}
.hero h1 {color: white; font-size: 1.8rem; margin: 0 0 .3rem 0;}
.hero p {color: #dbeafe; margin: 0;}
div[data-testid="stMetric"] {background: #f8fafc; border: 1px solid #dbeafe; border-radius: 12px; padding: .6rem .8rem;}
</style>
<div class="hero"><h1>📄 Demandes d’attestations administratives</h1><p>Faculté de Génie Électrique — Université Djillali Liabès de Sidi Bel Abbès</p></div>
""", unsafe_allow_html=True)


def clean_date_input(value: str) -> str | None:
    value = (value or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return None


def _norm_source_header(value: object) -> str:
    text = "" if value is None else str(value).strip().casefold()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(re.sub(r"[^a-z0-9]+", " ", text).split())


def _source_value(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "none", "nat", "<na>"} else text


def _normalize_source_date(value: object) -> str:
    raw = _source_value(value)
    if not raw:
        return ""
    parsed = clean_date_input(raw)
    if parsed:
        return parsed
    # Les réponses du formulaire contiennent aussi des dates écrites en français.
    month_map = {
        "janvier": "January", "fevrier": "February", "février": "February",
        "mars": "March", "avril": "April", "mai": "May", "juin": "June",
        "juillet": "July", "aout": "August", "août": "August",
        "septembre": "September", "octobre": "October", "novembre": "November",
        "decembre": "December", "décembre": "December",
    }
    translated = raw
    for fr, en in month_map.items():
        translated = re.sub(fr, en, translated, flags=re.IGNORECASE)
    for fmt in ("%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(translated.strip(), fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return raw


def _split_full_name(full_name: str) -> tuple[str, str]:
    """Préremplit Nom/Prénom selon l'intitulé source « Nom et prénom ».

    Les champs restent modifiables, car les réponses du formulaire ne suivent pas
    toutes le même ordre. L'utilisateur doit vérifier le découpage avant l'envoi.
    """
    words = re.sub(r"\s+", " ", _source_value(full_name)).split(" ")
    if not words:
        return "", ""
    if len(words) == 1:
        return words[0], ""
    return words[0], " ".join(words[1:])


def _find_source_column(columns: list[object], aliases: list[str]) -> str | None:
    normalized = {_norm_source_header(col): str(col) for col in columns}
    for alias in aliases:
        key = _norm_source_header(alias)
        if key in normalized:
            return normalized[key]
    # Match stable fragments for year-specific subject-column labels.
    for alias in aliases:
        key = _norm_source_header(alias)
        if len(key) >= 8:
            for col_norm, original in normalized.items():
                if key in col_norm or col_norm in key:
                    return original
    return None


@st.cache_data(show_spinner=False)
def _load_source_people(file_bytes: bytes) -> list[dict]:
    """Lit le fichier de réponses du personnel sans modifier le classeur source."""
    frame = pd.read_excel(BytesIO(file_bytes), dtype=str).fillna("")
    columns = list(frame.columns)
    mapping = {
        "full_name_fr": ["Nom et prénom en francais", "Nom et prénom en français", "Nom prénom français"],
        "full_name_ar": ["Nom et prénom en arabe", "Nom prénom arabe"],
        "dob": ["Date de naissance", "Date naissance"],
        "birthplace": ["Lieu de naissance", "Lieu naissance"],
        "email_primary": ["Adresse e-mail", "Email", "Adresse email"],
        "email_secondary": ["Adresse e-mail 2", "Adresse email 2", "Email personnel"],
        "phone": ["Numéro de téléphone", "Numero de telephone", "Téléphone", "Telephone"],
        "address": ["Adresse", "Adresse personnelle"],
        "subjects_s1": ["Matière (s) enseignée (s) en semestre 01_2026-2027", "Matières enseignées semestre 1", "Matiere s1"],
        "teaching_type": ["Type d'enseignement", "Type enseignement"],
        "degree": ["Dernier diplôme", "Dernier diplome", "Diplôme"],
        "department": ["Département", "Departement", "Département d'affiliation", "Département d'affectation"],
    }
    cols = {key: _find_source_column(columns, aliases) for key, aliases in mapping.items()}
    if not cols["full_name_fr"]:
        raise ValueError("Colonne 'Nom et prénom en français' introuvable dans le fichier source.")

    result: list[dict] = []
    for row_idx, (_, row) in enumerate(frame.iterrows(), start=2):
        full_fr = _source_value(row.get(cols["full_name_fr"], ""))
        if not full_fr:
            continue
        full_ar = _source_value(row.get(cols["full_name_ar"], "")) if cols["full_name_ar"] else ""
        nom, prenom = _split_full_name(full_fr)
        nom_ar, prenom_ar = _split_full_name(full_ar)
        mail = (_source_value(row.get(cols["email_secondary"], "")) if cols["email_secondary"] else "")
        if not mail and cols["email_primary"]:
            mail = _source_value(row.get(cols["email_primary"], ""))
        dept = _source_value(row.get(cols["department"], "")) if cols["department"] else ""
        result.append({
            "source_row": row_idx,
            "full_name_fr": full_fr,
            "full_name_ar": full_ar,
            "nom": nom,
            "prenom": prenom,
            "nom_ar": nom_ar,
            "prenom_ar": prenom_ar,
            "date_naissance": _normalize_source_date(row.get(cols["dob"], "")) if cols["dob"] else "",
            "lieu_naissance": _source_value(row.get(cols["birthplace"], "")) if cols["birthplace"] else "",
            "email": mail,
            "telephone": _source_value(row.get(cols["phone"], "")) if cols["phone"] else "",
            "adresse_source": _source_value(row.get(cols["address"], "")) if cols["address"] else "",
            "matiere_s1": _source_value(row.get(cols["subjects_s1"], "")) if cols["subjects_s1"] else "",
            "type_enseignement": _source_value(row.get(cols["teaching_type"], "")) if cols["teaching_type"] else "",
            "dernier_diplome": _source_value(row.get(cols["degree"], "")) if cols["degree"] else "",
            "source_department": dept,
        })
    return result


@st.cache_data(show_spinner=False)
def _load_edt_cached(file_bytes: bytes) -> list[dict]:
    """Charge le classeur EDT versionné dans le dépôt GitHub."""
    return load_edt_sessions(file_bytes)


def _safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(text or "")).strip("_") or "Personnel"


def _archive_docx_record(record: dict, language_mode: str | None = None, content: bytes | None = None) -> dict:
    """Sauvegarde une version Word générée dans archives_generations/ et dans le backend."""
    values = values_from_record(record)
    language_mode = language_mode or store_language_mode(record)
    if content is None:
        content = generate_attestation(values, language_mode)
    number = str(record.get("attestation_number") or values.get("numero") or "SansNumero")
    year = int(record.get("issued_year") or datetime.now(TZ_ALGERIA).year)
    request_id = str(record.get("id", ""))
    lang_tag = {LANGUES[0]: "Bilingue", LANGUES[1]: "Arabe", LANGUES[2]: "Francais"}.get(language_mode, "Document")
    safe_name = _safe_filename(f"{values.get('nom','')}_{values.get('prenom','')}")
    department = str(values.get("department", record.get("department", "")))
    filename = f"Attestation_{safe_name}_{_safe_filename(department)}_{number}_{year}_{lang_tag}.docx"
    archive_key = f"docx:{request_id}:{year}:{number}:{lang_tag}"
    try:
        result = save_generation_file(
            archive_key, filename, content,
            department=department,
            person_name=f"{values.get('prenom','')} {values.get('nom','')}".strip(),
            attestation_number=number,
            issued_year=year,
            language_mode=language_mode,
            file_kind="DOCX",
        )
        if not result.get("persistent", False):
            st.session_state["archive_save_notice"] = (
                "Le document Word a été généré, mais la sauvegarde persistante n'est pas disponible. "
                "Vérifiez le dossier local ou le schéma SQL des archives dans Supabase."
            )
        return result
    except Exception as exc:
        # Une erreur d'archivage ne doit pas invalider une attestation déjà émise.
        st.session_state["archive_save_notice"] = f"Attestation générée, mais erreur de sauvegarde des archives : {exc}"
        return {"persistent": False, "error": str(exc)}


def _archive_zip_file(content: bytes, file_name: str, records: list[dict], language_modes: list[str], file_kind: str = "ZIP") -> dict:
    departments = sorted({str(r.get("department", "")) for r in records if r.get("department")})
    department = departments[0] if len(departments) == 1 else "Tous les départements"
    record_ids = sorted(str(r.get("id", "")) for r in records)
    fingerprint_source = "|".join(record_ids) + "::" + "|".join(language_modes) + "::" + file_name
    fingerprint = hashlib.sha256(fingerprint_source.encode("utf-8")).hexdigest()[:24]
    years = [int(r.get("issued_year")) for r in records if str(r.get("issued_year", "")).isdigit()]
    issue_year = max(years) if years else datetime.now(TZ_ALGERIA).year
    try:
        result = save_generation_file(
            f"zip:{fingerprint}", file_name, content,
            department=department,
            person_name=f"Archive groupée ({len(records)} attestations)",
            attestation_number="",
            issued_year=issue_year,
            language_mode=", ".join(language_modes),
            file_kind=file_kind,
        )
        if not result.get("persistent", False):
            st.session_state["archive_save_notice"] = "Le ZIP est prêt, mais sa sauvegarde cloud n'a pas abouti. Vérifiez le schéma SQL des archives."
        return result
    except Exception as exc:
        st.session_state["archive_save_notice"] = f"ZIP prêt, mais erreur de sauvegarde : {exc}"
        return {"persistent": False, "error": str(exc)}


def _edt_subjects_grouped_by_type(sessions: list[dict]) -> tuple[list[str], str]:
    """Retourne les matières uniques groupées par type d'enseignement tiré de l'EDT.

    Exemple : ``Cours : Réseaux électriques, Machines électriques`` puis une ligne TD
    et/ou TP. Le type vient de la colonne Code de l'EDT, pas du formulaire personnel.
    """
    order = ["Cours", "TD", "TP", "Autre"]
    grouped: dict[str, list[str]] = {kind: [] for kind in order}
    seen: set[tuple[str, str]] = set()
    for session in sessions:
        subject = _source_value(session.get("Matiere") or session.get("Enseignements", ""))
        if not subject:
            continue
        kind = _source_value(session.get("Type", ""))
        if kind not in grouped:
            kind = "Autre"
        key = (kind, _norm_source_header(subject))
        if key in seen:
            continue
        seen.add(key)
        grouped[kind].append(subject)
    lines = [f"{kind} : {', '.join(grouped[kind])}" for kind in order if grouped[kind]]
    return [f"{kind} — {subject}" for kind in order for subject in grouped[kind]], "\n".join(lines)


def _edt_promotions(sessions: list[dict]) -> list[str]:
    """Promotions présentes dans les séances effectivement associées à l'enseignant."""
    seen: dict[str, str] = {}
    for session in sessions:
        value = _source_value(session.get("Promotion", ""))
        if value:
            seen.setdefault(_norm_source_header(value), value)
    return sorted(seen.values(), key=_norm_source_header)


def _build_source_snapshot(person: dict, source_name: str, edt_sessions: list[dict],
                           matched_teachers: list[str], match_mode: str, edt_source_name: str) -> dict:
    """Snapshot utile au chef et au générateur pour une personne issue du classeur source."""
    subjects = unique_subjects(edt_sessions) if edt_sessions else []
    subjects_with_type, subjects_text = _edt_subjects_grouped_by_type(edt_sessions)
    promotions = _edt_promotions(edt_sessions)
    if not edt_sessions:
        subjects_text = _source_value(person.get("matiere_s1", ""))
    session_rows = [
        {k: str(session.get(k, "") or "") for k in
         ("Enseignements", "Matiere", "Code", "Type", "Enseignants", "Horaire", "Jours", "Lieu", "Promotion", "Département")}
        for session in edt_sessions
    ]
    return {
        "source_row": person.get("source_row"),
        "source_file_name": source_name,
        "source_nom_complet_fr": person.get("full_name_fr", ""),
        "source_nom_complet_ar": person.get("full_name_ar", ""),
        "nom_ar": person.get("nom_ar", ""),
        "prenom_ar": person.get("prenom_ar", ""),
        "matiere_s1": subjects_text,
        "edt_subjects_text": subjects_text,
        "edt_subjects": subjects,
        "edt_subjects_with_type": subjects_with_type,
        "edt_promotions": promotions,
        "edt_promotions_text": ", ".join(promotions),
        "edt_teacher_names": matched_teachers,
        "edt_match_mode": match_mode,
        "edt_source_file": edt_source_name,
        "edt_session_count": len(edt_sessions),
        "edt_sessions": session_rows,
        "type_enseignement": person.get("type_enseignement", ""),
        "dernier_diplome": person.get("dernier_diplome", ""),
    }


def _attestations_zip(records: list[dict], language_modes: list[str] | None = None, *, save_to_archive: bool = False) -> tuple[bytes, int]:
    """Regroupe les attestations déjà émises dans un ou plusieurs formats de langue.

    Chaque copie reprend le même numéro, la même date et la même référence officielle.
    Le téléchargement ne crée jamais de nouvelle émission.
    """
    selected_modes = [mode for mode in (language_modes or [LANGUES[0]]) if mode in LANGUES]
    if not selected_modes:
        selected_modes = [LANGUES[0]]
    language_suffix = {
        LANGUES[0]: "Bilingue",
        LANGUES[1]: "Arabe",
        LANGUES[2]: "Francais",
    }
    out = BytesIO()
    count = 0
    used_names: set[str] = set()
    with ZipFile(out, "w", compression=ZIP_DEFLATED) as archive:
        for record in records:
            if str(record.get("status", "")) != "Générée" or record.get("attestation_number") in (None, ""):
                continue
            values = values_from_record(record)
            if not values.get("numero") or not values.get("date_attestation"):
                continue
            base = _safe_filename(
                f"Attestation_{values.get('nom','')}_{values.get('prenom','')}_"
                f"{values.get('department','')}_{values['numero']}_{record.get('issued_year','')}"
            )
            for language_mode in selected_modes:
                docx = generate_attestation(values, language_mode)
                filename = f"{base}_{language_suffix.get(language_mode, 'Document')}.docx"
                if filename in used_names:
                    filename = f"{base}_{str(record.get('id',''))[:8]}_{language_suffix.get(language_mode, 'Document')}.docx"
                used_names.add(filename)
                archive.writestr(filename, docx)
                # L'archivage persistant est déclenché explicitement ou lors d'une génération officielle en lot.
                if save_to_archive:
                    archived_doc = _archive_docx_record(record, language_mode, content=docx)
                    if not archived_doc.get("persistent", False):
                        st.session_state["archive_save_notice"] = "Une copie Word a été générée, mais n'a pas pu être conservée dans le stockage persistant. Vérifiez le schéma SQL des archives."
                count += 1
        archive.writestr(
            "LISEZ-MOI.txt",
            "Archive des attestations administratives déjà émises.\n"
            "Les demandes en attente sont exclues : elles doivent être vérifiées et émises par le chef.\n"
            f"Formats sélectionnés : {', '.join(selected_modes)}\n"
            f"Nombre de fichiers Word : {count}\n"
            "Les versions linguistiques d'une même attestation conservent le numéro, la date et la référence déjà attribués.\n",
        )
    zip_content = out.getvalue()
    if count and save_to_archive:
        dept_name = "Tous_Departements" if len({r.get("department") for r in records if r.get("department")}) != 1 else _safe_filename(next(iter({r.get("department") for r in records if r.get("department")})))
        language_label = "_".join({LANGUES[0]: "Bilingue", LANGUES[1]: "Arabe", LANGUES[2]: "Francais"}[item] for item in selected_modes)
        archive_name = f"Attestations_{dept_name}_{language_label}.zip"
        try:
            saved_zip = _archive_zip_file(zip_content, archive_name, records, selected_modes, file_kind="ZIP attestations")
            if not saved_zip.get("persistent", False):
                st.session_state["archive_save_notice"] = "Le ZIP a été préparé mais sa sauvegarde cloud a échoué. Vérifiez la migration SQL des archives."
        except Exception as exc:
            # Le téléchargement principal ne doit pas échouer à cause du stockage d'archives.
            st.session_state["archive_save_notice"] = f"ZIP généré, mais sa sauvegarde a rencontré une erreur : {exc}"
    return zip_content, count


def _department_from_source(value: str) -> str | None:
    key = _norm_source_header(value)
    aliases = {
        "Electrotechnique": ["electrotechnique", "electrotechnique elt", "elt"],
        "Automatique": ["automatique", "auto"],
        "Electronique": ["electronique", "eln"],
        "Télécommunications": ["telecommunications", "telecommunication", "telecom", "tlc"],
    }
    for department, candidates in aliases.items():
        if key in {_norm_source_header(candidate) for candidate in candidates}:
            return department
    return None


def _chef_par_defaut(department: str) -> tuple[str, str]:
    suffix = DEPT_SECRET_SUFFIX.get(department, "")
    fr = get_setting(f"CHEF_NAME_FR_{suffix}") if suffix else ""
    ar = get_setting(f"CHEF_NAME_AR_{suffix}") if suffix else ""
    fallback = CHEFS_PAR_DEFAUT.get(department, {})
    return fr or fallback.get("fr", ""), ar or fallback.get("ar", "")


def dept_password_key(department: str) -> str:
    return "CHEF_PASSWORD_" + DEPT_SECRET_SUFFIX[department]


def expected_password(department: str) -> tuple[str, bool]:
    """Retourne mot de passe, et indique s'il s'agit du super-administrateur."""
    master = get_setting("ADMIN_PASSWORD")
    dept_pwd = get_setting(dept_password_key(department))
    return master, bool(master)


def configured_for_department(department: str) -> str:
    return get_setting(dept_password_key(department))


def send_head_notification(record: dict) -> tuple[bool, str]:
    """Envoie une notification sans inclure les détails personnels du demandeur."""
    suffix = DEPT_SECRET_SUFFIX.get(record.get("department", ""), "")
    recipient = get_setting("CHEF_EMAIL_" + suffix) if suffix else ""
    host = get_setting("SMTP_HOST")
    sender = get_setting("SMTP_FROM")
    if not (recipient and host and sender):
        return False, "Notification e-mail non configurée : la demande est enregistrée et visible dans l'espace du chef."
    try:
        port_txt = get_setting("SMTP_PORT", "587") or "587"
        port = int(port_txt)
        msg = EmailMessage()
        msg["Subject"] = f"Nouvelle demande d'attestation — {record['department']}"
        msg["From"] = sender
        msg["To"] = recipient
        app_url = get_setting("PUBLIC_APP_URL")
        app_line = f"\nOuvrir l'application : {app_url}\n" if app_url else ""
        msg.set_content(
            "Une nouvelle demande d'attestation administrative a été déposée dans l'application UDL-SBA.\n\n"
            f"Département : {record['department']}\n"
            f"Référence de suivi : {record['id']}\n"
            "Pour protéger les données personnelles, le courriel ne contient pas l'identité du demandeur. "
            "Connectez-vous à l'espace du chef de département pour consulter et traiter le dossier.\n"
            + app_line
        )
        username = get_setting("SMTP_USER")
        password = get_setting("SMTP_PASSWORD")
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context()) as server:
                if username:
                    server.login(username, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=20) as server:
                server.ehlo()
                if port in (587, 25):
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                if username:
                    server.login(username, password)
                server.send_message(msg)
        return True, f"Notification envoyée au chef du département {record['department']}."
    except Exception as exc:
        return False, f"La demande est enregistrée, mais la notification e-mail a échoué : {exc}"


def send_head_batch_notification(department: str, count: int) -> tuple[bool, str]:
    """Envoie un seul avis au chef par département après un dépôt en lot."""
    suffix = DEPT_SECRET_SUFFIX.get(department, "")
    recipient = get_setting("CHEF_EMAIL_" + suffix) if suffix else ""
    host = get_setting("SMTP_HOST")
    sender = get_setting("SMTP_FROM")
    if not (recipient and host and sender):
        return False, "Notification e-mail non configurée."
    try:
        port = int(get_setting("SMTP_PORT", "587") or "587")
        msg = EmailMessage()
        msg["Subject"] = f"Demandes groupées d'attestation — {department}"
        msg["From"] = sender
        msg["To"] = recipient
        app_url = get_setting("PUBLIC_APP_URL")
        app_line = f"\nOuvrir l'application : {app_url}\n" if app_url else ""
        msg.set_content(
            f"{count} demande(s) d'attestation administrative ont été ajoutées à la file du département {department}.\n\n"
            "Aucune identité n'est incluse dans ce courriel. Connectez-vous à l'espace sécurisé pour vérifier et traiter les dossiers.\n"
            + app_line
        )
        username = get_setting("SMTP_USER")
        password = get_setting("SMTP_PASSWORD")
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context()) as server:
                if username:
                    server.login(username, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(host, port, timeout=20) as server:
                server.ehlo()
                if port in (587, 25):
                    server.starttls(context=ssl.create_default_context())
                    server.ehlo()
                if username:
                    server.login(username, password)
                server.send_message(msg)
        return True, "Notification groupée envoyée."
    except Exception as exc:
        return False, f"Notification groupée non envoyée : {exc}"


def label_request(row: dict) -> str:
    created = str(row.get("created_at", ""))[:16].replace("T", " ")
    return f"{row.get('prenom', '')} {row.get('nom', '')} — {row.get('department', '')} — déposée le {created} — {str(row.get('id', ''))[:8]}"


def values_from_record(record: dict) -> dict:
    values = dict(record.get("final_data") or {})
    source_snapshot = record.get("source_snapshot") or {}
    # Repli utile pour les demandes historiques et les champs non encore copiés dans final_data.
    if not values.get("matiere_s1"):
        values["matiere_s1"] = source_snapshot.get("edt_subjects_text") or source_snapshot.get("matiere_s1", "")
    if not values.get("nom_ar"):
        values["nom_ar"] = source_snapshot.get("nom_ar", "")
    if not values.get("prenom_ar"):
        values["prenom_ar"] = source_snapshot.get("prenom_ar", "")
    if not values.get("promotions_s1"):
        values["promotions_s1"] = values.get("section_s1") or source_snapshot.get("edt_promotions_text", "")
    if not values.get("promotions_s2"):
        values["promotions_s2"] = values.get("section_s2", "")
    # Les informations identitaires déposées par le demandeur restent la source de vérité.
    values.update({
        "department": record.get("department", "Electrotechnique"),
        "nom": record.get("nom", ""),
        "prenom": record.get("prenom", ""),
        "date_naissance": record.get("date_naissance", ""),
        "lieu_naissance": record.get("lieu_naissance", ""),
        "numero": str(record.get("attestation_number") or ""),
        "date_attestation": record.get("issued_date") or "",
    })
    return values


def store_language_mode(record: dict) -> str:
    final_data = record.get("final_data") or {}
    return final_data.get("language_mode") or record.get("language_mode") or LANGUES[0]


def render_download(record: dict, button_key: str) -> None:
    try:
        values = values_from_record(record)
        if not values.get("numero") or not values.get("date_attestation"):
            st.warning("Le numéro et la date ne sont pas encore attribués. Générez officiellement cette attestation depuis les demandes en attente.")
            return
        content = generate_attestation(values, store_language_mode(record))
        safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{values.get('nom','')}_{values.get('prenom','')}").strip("_") or "Personnel"
        filename = f"Attestation_{safe_name}_{values.get('department','')}_{values['numero']}_{record.get('issued_year','')}.docx"
        st.download_button(
            "📥 Télécharger l'attestation Word (.docx)",
            data=content,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
            key=button_key,
        )
        st.caption(f"Numéro officiel : {values['numero']}/{record.get('issued_year', '')} · Date de délivrance : {values['date_attestation']}")
    except Exception as exc:
        st.error(f"Impossible de reconstruire le document Word : {exc}")


# Navigation publique / chef de département.
role_options = ["Déposer une demande", "Espace chef de département"]
role = st.radio("Accès à l'application", role_options, horizontal=True, key="workflow_role")

mode = backend_mode()
if mode == "sqlite":
    st.warning(
        ""
        ""
    )

if role == "Déposer une demande":
    st.subheader("Déposer une demande d'attestation")
    st.write(
        "Sélectionnez votre nom dans le fichier source : les informations disponibles seront préremplies. "
        "Vérifiez le nom, le prénom et la date de naissance avant d'envoyer. Le chef complétera les données pédagogiques et administratives."
    )

    confirmation = st.session_state.get("latest_request_confirmation")
    if confirmation:
        st.success(f"Demande déposée avec succès. Votre référence de suivi est : {confirmation['id']}")
        st.info(confirmation.get("notice", "La demande a été enregistrée."))
        if st.button("Masquer la confirmation", key="hide_request_confirmation"):
            st.session_state.pop("latest_request_confirmation", None)
            st.rerun()

    if SOURCE_PERSONNEL_FILE.exists():
        source_bytes = SOURCE_PERSONNEL_FILE.read_bytes()
        source_name = SOURCE_PERSONNEL_FILE.name
    else:
        st.warning("Le fichier source 'Vos coordonnées (réponses).xlsx' est introuvable à côté de app.py. Téléversez-le pour charger la liste des demandeurs.")
        uploaded_source = st.file_uploader("Téléverser le fichier source du personnel (.xlsx)", type=["xlsx"], key="source_personnel_upload")
        source_bytes = uploaded_source.getvalue() if uploaded_source is not None else b""
        source_name = uploaded_source.name if uploaded_source is not None else ""

    source_people: list[dict] = []
    if source_bytes:
        try:
            source_people = _load_source_people(source_bytes)
            st.caption(f"Source chargée : {source_name} · {len(source_people)} personne(s) trouvée(s).")
        except Exception as exc:
            st.error(f"Impossible de lire le fichier source : {exc}")

    if source_people:
        def _format_source_person(choice):
            if choice == ALL_SOURCE_REQUESTERS:
                return "🌐 Tous les demandeurs — préparer les demandes en lot"
            person = source_people[int(choice)]
            return person["full_name_fr"] + (f" — {person['full_name_ar']}" if person.get("full_name_ar") else "")

        selected_source_idx = st.selectbox(
            "👤 Sélectionner le demandeur dans le fichier source *",
            options=[ALL_SOURCE_REQUESTERS] + [str(idx) for idx in range(len(source_people))],
            index=1,
            format_func=_format_source_person,
            key="source_person_selection",
            help="Choisissez une personne pour déposer une demande individuelle, ou « Tous les demandeurs » pour préparer les demandes du fichier en une seule opération. La génération officielle reste soumise à la vérification et aux compléments du chef.",
        )

        if selected_source_idx == ALL_SOURCE_REQUESTERS:
            st.markdown("### 📚 Préparer les demandes de tous les demandeurs")
            st.warning(
                "Le fichier source ne contient pas de colonne Département. Pour éviter d'envoyer une demande au mauvais chef, "
                "attribuez un département à chaque personne dans le tableau. Cette opération crée des demandes **en attente** ; "
                "le chef devra compléter et valider les informations administratives avant l'émission officielle."
            )
            admin_pwd_configured = get_setting("ADMIN_PASSWORD")
            if not admin_pwd_configured:
                st.error("L'envoi groupé est désactivé tant que le secret ADMIN_PASSWORD n'est pas configuré dans Streamlit Secrets.")
            batch_df = pd.DataFrame({
                "Ligne source": [int(person.get("source_row", i + 2)) for i, person in enumerate(source_people)],
                "Demandeur": [person.get("full_name_fr", "") for person in source_people],
                "Département": [
                    _department_from_source(person.get("source_department", "")) or "— Choisir —"
                    for person in source_people
                ],
            })
            department_choices = ["— Choisir —"] + list(DEPARTEMENTS.keys())
            with st.form("form_demandes_groupes_source", clear_on_submit=False):
                edited_batch = st.data_editor(
                    batch_df,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="fixed",
                    column_config={
                        "Ligne source": st.column_config.NumberColumn("Ligne source", disabled=True, width="small"),
                        "Demandeur": st.column_config.TextColumn("Demandeur", disabled=True, width="large"),
                        "Département": st.column_config.SelectboxColumn("Département à associer", options=department_choices, required=True, width="medium"),
                    },
                    key="bulk_source_department_editor",
                )
                bulk_default_language = st.selectbox("Langue par défaut des demandes", LANGUES, index=0, key="bulk_source_language")
                bulk_admin_password = st.text_input("Mot de passe administrateur pour autoriser l'envoi groupé", type="password", key="bulk_source_admin_password")
                bulk_confirm = st.checkbox("Je confirme vouloir créer une demande en attente pour chaque personne de ce tableau.", key="bulk_source_confirm")
                submit_bulk_source = st.form_submit_button("📨 Préparer toutes les demandes", type="primary", use_container_width=True)
            if submit_bulk_source:
                if not admin_pwd_configured or not hmac.compare_digest(bulk_admin_password or "", admin_pwd_configured):
                    st.error("Mot de passe administrateur incorrect ou non configuré. Aucune demande n'a été créée.")
                elif not bulk_confirm:
                    st.error("Cochez la confirmation avant de créer les demandes groupées.")
                else:
                    missing_depts = edited_batch[edited_batch["Département"].isin(["— Choisir —", "", None])]
                    if not missing_depts.empty:
                        st.error("Choisissez un département pour chaque personne avant de continuer : " + ", ".join(missing_depts["Demandeur"].astype(str).tolist()))
                    else:
                        # Charge l'EDT une seule fois pour alimenter chaque dossier de façon cohérente.
                        edt_all_for_batch = []
                        edt_batch_name = ""
                        if EDT_SOURCE_FILE.exists():
                            try:
                                edt_batch_name = EDT_SOURCE_FILE.name
                                edt_all_for_batch = _load_edt_cached(EDT_SOURCE_FILE.read_bytes())
                            except Exception as exc:
                                st.warning(f"Le classeur EDT n'a pas pu être utilisé pour le lot : {exc}")
                        try:
                            existing_requests = list_requests(limit=10000)
                        except Exception:
                            existing_requests = []
                        existing_sources = []
                        for old_request in existing_requests:
                            old_snapshot = old_request.get("source_snapshot") or {}
                            existing_sources.append((
                                str(old_snapshot.get("source_row", "")),
                                _norm_source_header(old_snapshot.get("source_nom_complet_fr", "")),
                                str(old_snapshot.get("source_file_name", "")),
                            ))
                        created_count, duplicate_count, failed = 0, 0, []
                        created_by_department: dict[str, int] = {}
                        dept_by_source_row = {int(row["Ligne source"]): str(row["Département"]) for _, row in edited_batch.iterrows()}
                        for person in source_people:
                            source_row_no = int(person.get("source_row", 0))
                            dept_value = dept_by_source_row.get(source_row_no, "— Choisir —")
                            identity_key = _norm_source_header(person.get("full_name_fr", ""))
                            is_duplicate = any(
                                old_row == str(source_row_no)
                                and old_name == identity_key
                                and (not old_file or old_file == source_name)
                                for old_row, old_name, old_file in existing_sources
                            )
                            if is_duplicate:
                                duplicate_count += 1
                                continue
                            dob = clean_date_input(person.get("date_naissance", ""))
                            birthplace = _source_value(person.get("lieu_naissance", ""))
                            if not (person.get("nom") and person.get("prenom") and dob and birthplace):
                                failed.append(f"{person.get('full_name_fr','')} (identité ou date/lieu de naissance incomplet)")
                                continue
                            person_sessions, matched_names, match_mode = ([], [], "aucun")
                            if edt_all_for_batch:
                                try:
                                    person_sessions, matched_names, match_mode = match_teacher_sessions(person.get("full_name_fr", ""), edt_all_for_batch)
                                except Exception:
                                    person_sessions, matched_names, match_mode = ([], [], "aucun")
                            snapshot = _build_source_snapshot(person, source_name, person_sessions, matched_names, match_mode, edt_batch_name)
                            try:
                                record = create_request({
                                    "department": dept_value,
                                    "nom": person["nom"],
                                    "prenom": person["prenom"],
                                    "date_naissance": dob,
                                    "lieu_naissance": birthplace,
                                    "email": person.get("email", ""),
                                    "telephone": person.get("telephone", ""),
                                    "purpose": "Demande préparée en lot depuis le fichier source du personnel.",
                                    "language_mode": bulk_default_language,
                                    "source_snapshot": snapshot,
                                })
                                created_count += 1
                                created_by_department[dept_value] = created_by_department.get(dept_value, 0) + 1
                            except Exception as exc:
                                failed.append(f"{person.get('full_name_fr','')}: {exc}")
                        notifications = []
                        for dept_name, dept_count in created_by_department.items():
                            ok, notice = send_head_batch_notification(dept_name, dept_count)
                            if notice and ok:
                                notifications.append(f"{dept_name} : notification envoyée")
                        st.session_state["bulk_source_request_result"] = {
                            "created": created_count, "duplicates": duplicate_count, "failed": failed,
                            "notifications": notifications,
                        }
                        st.rerun()
            result = st.session_state.get("bulk_source_request_result")
            if result:
                st.success(f"Lot terminé : {result['created']} demande(s) créée(s), {result['duplicates']} déjà existante(s) ignorée(s), {len(result['failed'])} erreur(s).")
                if result["failed"]:
                    with st.expander("Voir les lignes non créées"):
                        st.write("\n".join(result["failed"]))
                if result.get("notifications"):
                    st.caption(" · ".join(result["notifications"]))
                if st.button("Effacer le bilan du lot", key="clear_bulk_source_result"):
                    st.session_state.pop("bulk_source_request_result", None)
                    st.rerun()
            st.info("Étape suivante : connectez-vous dans « Espace chef de département » pour compléter chaque demande. Quand les attestations sont émises, utilisez l'archive ZIP et sélectionnez un ou plusieurs formats de langue.")
            st.stop()

        selected_source_idx = int(selected_source_idx)
        selected_person = source_people[selected_source_idx]

        # L'EDT est lu depuis le fichier placé dans le dépôt GitHub, nommé exactement EDT_Edition.xlsx.
        if EDT_SOURCE_FILE.exists():
            edt_bytes = EDT_SOURCE_FILE.read_bytes()
            edt_source_name = EDT_SOURCE_FILE.name
        else:
            st.warning("Le fichier EDT_Edition.xlsx est introuvable à côté de app.py. Téléversez-le pour récupérer automatiquement les matières et les séances.")
            uploaded_edt = st.file_uploader("Téléverser le fichier de l'emploi du temps (.xlsx)", type=["xlsx"], key="edt_source_upload")
            edt_bytes = uploaded_edt.getvalue() if uploaded_edt is not None else b""
            edt_source_name = uploaded_edt.name if uploaded_edt is not None else "EDT_Edition.xlsx"
        edt_all_sessions: list[dict] = []
        edt_sessions: list[dict] = []
        edt_matched_teachers: list[str] = []
        edt_match_mode = "aucun"
        if edt_bytes:
            try:
                edt_all_sessions = _load_edt_cached(edt_bytes)
                edt_sessions, edt_matched_teachers, edt_match_mode = match_teacher_sessions(selected_person.get("full_name_fr", ""), edt_all_sessions)
            except Exception as exc:
                st.error(f"Impossible de lire le fichier EDT : {exc}")

        # Si le nom source ne correspond pas au nom inscrit dans l'EDT, le demandeur peut
        # sélectionner une correspondance explicite. Aucun cours n'est inventé.
        if edt_all_sessions and not edt_sessions:
            with st.expander("🔎 Associer manuellement le nom à l'enseignant présent dans l'EDT", expanded=False):
                available_edt_teachers = unique_teacher_names(edt_all_sessions)
                manual_names = st.multiselect(
                    "Nom(s) à rechercher dans EDT_Edition.xlsx",
                    options=available_edt_teachers,
                    key=f"edt_manual_teacher_{selected_source_idx}",
                    help="Utilisez cette option uniquement lorsque le nom de la source est écrit différemment dans l'EDT.",
                )
            if manual_names:
                edt_sessions, edt_matched_teachers, edt_match_mode = match_teacher_sessions(
                    selected_person.get("full_name_fr", ""), edt_all_sessions, explicit_names=manual_names
                )

        edt_subjects = unique_subjects(edt_sessions)
        edt_subjects_with_type, edt_subjects_text = _edt_subjects_grouped_by_type(edt_sessions)
        edt_promotions = _edt_promotions(edt_sessions)
        edt_promotions_text = ", ".join(edt_promotions)
        if not edt_sessions:
            # La matière auto-déclarée reste visible comme solution de repli, mais aucun type
            # n'est inventé lorsque le nom ne correspond pas à l'EDT.
            edt_subjects_text = _source_value(selected_person.get("matiere_s1", ""))
        st.info(
            f"**Nom en français :** {selected_person['full_name_fr']}  \n"
            f"**Nom en arabe :** {selected_person.get('full_name_ar') or 'Non renseigné'}  \n"
            f"**Matière(s) source S1 :** {selected_person.get('matiere_s1') or 'Non renseignée'}  \n"
            f"**Type d'enseignement :** {selected_person.get('type_enseignement') or 'Non renseigné'}  ·  "
            f"**Dernier diplôme :** {selected_person.get('dernier_diplome') or 'Non renseigné'}"
        )
        st.markdown("### 📅 Emploi du temps individuel extrait du fichier EDT_Edition.xlsx")
        if edt_sessions:
            match_note = "correspondance automatique" if edt_match_mode == "exact" else ("correspondance approximative à vérifier" if edt_match_mode == "approximatif" else "association manuelle")
            st.success(f"{len(edt_sessions)} séance(s) trouvée(s) · Enseignant(s) EDT : {', '.join(edt_matched_teachers)} · {match_note}.")
            if edt_match_mode == "approximatif":
                st.warning("Le nom a été rapproché par similarité. Vérifiez que les séances appartiennent bien au demandeur avant de transmettre la demande.")
            grid_days, grid_times, grid_matrix = timetable_grid_records(edt_sessions)
            grid_df = pd.DataFrame(grid_matrix, columns=["Jour"] + grid_times)
            st.dataframe(grid_df, use_container_width=True, hide_index=True)
            st.caption("Les jours sont affichés verticalement et les créneaux horaires horizontalement. Les matières distinctes ci-dessous sont reprises automatiquement dans le champ du semestre 1 de l'attestation.")
            safe_teacher = _safe_filename(selected_person.get("full_name_fr", "Enseignant"))
            html_bytes = export_timetable_html(selected_person.get("full_name_fr", "Enseignant"), edt_sessions)
            xlsx_bytes = export_timetable_xlsx(selected_person.get("full_name_fr", "Enseignant"), edt_sessions)
            pdf_bytes = export_timetable_pdf(selected_person.get("full_name_fr", "Enseignant"), edt_sessions)
            exp1, exp2, exp3 = st.columns(3)
            with exp1:
                st.download_button("🌐 Télécharger l'EDT en HTML", data=html_bytes, file_name=f"EDT_{safe_teacher}.html", mime="text/html", use_container_width=True, key=f"edt_html_{selected_source_idx}")
            with exp2:
                st.download_button("📊 Télécharger l'EDT en Excel", data=xlsx_bytes, file_name=f"EDT_{safe_teacher}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"edt_excel_{selected_source_idx}")
            with exp3:
                st.download_button("📄 Télécharger l'EDT en PDF", data=pdf_bytes, file_name=f"EDT_{safe_teacher}.pdf", mime="application/pdf", use_container_width=True, key=f"edt_pdf_{selected_source_idx}")
            with st.expander(f"📚 Matières distinctes détectées ({len(edt_subjects)})", expanded=False):
                st.write("\n".join(f"- {subject}" for subject in edt_subjects_with_type) or "Aucune matière détectée.")
                st.write("**Promotion(s) trouvée(s) :** " + (edt_promotions_text or "Aucune promotion détectée."))
        else:
            st.warning("Aucune séance n'a été trouvée dans EDT_Edition.xlsx pour ce nom. Les matières du formulaire source restent disponibles comme solution de repli ; vérifiez le nom ou utilisez l'association manuelle ci-dessus.")

        source_department = _department_from_source(selected_person.get("source_department", ""))
        department_options = list(DEPARTEMENTS.keys())
        if source_department:
            department_options_ui = department_options
            department_default_index = department_options.index(source_department)
        else:
            department_options_ui = ["— Choisir le département —"] + department_options
            department_default_index = 0
            st.caption("Le classeur source ne contient pas de colonne Département : choisissez manuellement le département concerné. Aucun département n'est présélectionné.")

        with st.form("form_depot_demande_attestation", clear_on_submit=False):
            department_req = st.selectbox("Département concerné *", department_options_ui, index=department_default_index)
            language_req = st.selectbox("Langue souhaitée pour l'attestation *", LANGUES, index=0)
            c1, c2 = st.columns(2)
            with c1:
                nom_req = st.text_input("Nom *", value=selected_person.get("nom", ""), max_chars=100, key=f"req_nom_{selected_source_idx}")
                prenom_req = st.text_input("Prénom *", value=selected_person.get("prenom", ""), max_chars=100, key=f"req_prenom_{selected_source_idx}")
                date_naissance_req = st.text_input("Date de naissance * (JJ/MM/AAAA)", value=selected_person.get("date_naissance", ""), placeholder="JJ/MM/AAAA", max_chars=40, key=f"req_dob_{selected_source_idx}")
            with c2:
                lieu_naissance_req = st.text_input("Lieu de naissance *", value=selected_person.get("lieu_naissance", ""), max_chars=160, key=f"req_birthplace_{selected_source_idx}")
                email_req = st.text_input("Adresse e-mail de contact", value=selected_person.get("email", ""), max_chars=180, key=f"req_email_{selected_source_idx}")
                telephone_req = st.text_input("Téléphone de contact", value=selected_person.get("telephone", ""), max_chars=50, key=f"req_phone_{selected_source_idx}")
            st.markdown("#### Identité en arabe (préremplie depuis la source, modifiable)")
            ar_c1, ar_c2 = st.columns(2)
            with ar_c1:
                nom_ar_req = st.text_input("Nom en arabe", value=selected_person.get("nom_ar", ""), max_chars=100, key=f"req_nom_ar_{selected_source_idx}")
            with ar_c2:
                prenom_ar_req = st.text_input("Prénom en arabe", value=selected_person.get("prenom_ar", ""), max_chars=120, key=f"req_prenom_ar_{selected_source_idx}")
            purpose_req = st.text_area("Précisions facultatives à transmettre au chef", max_chars=1500, placeholder="Ex. : précision utile pour identifier ou traiter la demande…")
            accepted_privacy = st.checkbox("Je confirme être la personne sélectionnée et que les informations affichées sont exactes.")
            submit_req = st.form_submit_button("📨 Envoyer la demande au chef de département", type="primary", use_container_width=True)

        if submit_req:
            normalized_dob = clean_date_input(date_naissance_req) or _normalize_source_date(date_naissance_req)
            if not clean_date_input(normalized_dob):
                normalized_dob = None
            problems = []
            if department_req not in DEPARTEMENTS: problems.append("le département concerné")
            if not nom_req.strip(): problems.append("le nom")
            if not prenom_req.strip(): problems.append("le prénom")
            if not normalized_dob: problems.append("une date de naissance valide au format JJ/MM/AAAA")
            if not lieu_naissance_req.strip(): problems.append("le lieu de naissance")
            if email_req.strip() and not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email_req.strip()):
                problems.append("une adresse e-mail valide ou laissez ce champ vide")
            if not accepted_privacy: problems.append("la confirmation d'identité et de l'exactitude des informations")
            if problems:
                st.error("Veuillez corriger : " + "; ".join(problems) + ".")
            else:
                try:
                    source_snapshot = {
                        "source_row": selected_person.get("source_row"),
                        "source_file_name": source_name,
                        "source_nom_complet_fr": selected_person.get("full_name_fr", ""),
                        "source_nom_complet_ar": " ".join(part for part in (nom_ar_req.strip(), prenom_ar_req.strip()) if part),
                        "nom_ar": nom_ar_req.strip(),
                        "prenom_ar": prenom_ar_req.strip(),
                        "matiere_s1": edt_subjects_text or selected_person.get("matiere_s1", ""),
                        "edt_subjects_text": edt_subjects_text,
                        "edt_subjects": edt_subjects,
                        "edt_subjects_with_type": edt_subjects_with_type,
                        "edt_promotions": edt_promotions,
                        "edt_promotions_text": edt_promotions_text,
                        "edt_teacher_names": edt_matched_teachers,
                        "edt_match_mode": edt_match_mode,
                        "edt_source_file": edt_source_name if edt_bytes else "",
                        "edt_session_count": len(edt_sessions),
                        "edt_sessions": [
                            {k: str(session.get(k, "") or "") for k in ("Enseignements", "Matiere", "Code", "Type", "Enseignants", "Horaire", "Jours", "Lieu", "Promotion", "Département")}
                            for session in edt_sessions
                        ],
                        "type_enseignement": selected_person.get("type_enseignement", ""),
                        "dernier_diplome": selected_person.get("dernier_diplome", ""),
                    }
                    record = create_request({
                        "department": department_req,
                        "nom": nom_req.strip(),
                        "prenom": prenom_req.strip(),
                        "date_naissance": normalized_dob,
                        "lieu_naissance": lieu_naissance_req.strip(),
                        "email": email_req.strip(),
                        "telephone": telephone_req.strip(),
                        "purpose": purpose_req.strip(),
                        "language_mode": language_req,
                        "source_snapshot": source_snapshot,
                    })
                    mail_ok, mail_notice = send_head_notification(record)
                    notice = mail_notice if mail_ok else (mail_notice + " Le chef peut consulter la demande dans l'espace protégé de l'application.")
                    st.session_state["latest_request_confirmation"] = {"id": record["id"], "notice": notice}
                    st.rerun()
                except Exception as exc:
                    st.error(f"La demande n'a pas pu être enregistrée : {exc}")
    else:
        st.info("Chargez le fichier source pour pouvoir sélectionner automatiquement un demandeur.")

    with st.expander("🔎 À propos des données préremplies"):
        st.markdown(
            "Les coordonnées sont reprises du classeur source. Le champ département est choisi manuellement, car il n'existe pas dans le fichier transmis. "
            "Les matières du semestre 1, le type d'enseignement et le dernier diplôme seront transmis au chef comme informations de référence. "
            "Le dépôt de la demande ne génère pas encore de numéro officiel."
        )

else:
    st.subheader("Espace sécurisé du chef de département")
    if not st.session_state.get("chief_authenticated", False):
        with st.form("form_login_chef", clear_on_submit=False):
            dept_login = st.selectbox("Votre département", list(DEPARTEMENTS.keys()))
            pwd_login = st.text_input("Mot de passe", type="password")
            login = st.form_submit_button("🔐 Ouvrir l'espace de traitement", type="primary", use_container_width=True)
        if login:
            master = get_setting("ADMIN_PASSWORD")
            dept_pwd = configured_for_department(dept_login)
            if master and hmac.compare_digest(pwd_login, master):
                st.session_state["chief_authenticated"] = True
                st.session_state["chief_department"] = None
                st.session_state["chief_is_master"] = True
                st.rerun()
            elif dept_pwd and hmac.compare_digest(pwd_login, dept_pwd):
                st.session_state["chief_authenticated"] = True
                st.session_state["chief_department"] = dept_login
                st.session_state["chief_is_master"] = False
                st.rerun()
            else:
                st.error("Mot de passe incorrect ou non configuré pour ce département.")
        st.caption("Chaque chef peut disposer d'un mot de passe propre à son département. L'administrateur global peut traiter les quatre départements.")
        if not get_setting("ADMIN_PASSWORD") and not any(configured_for_department(d) for d in DEPARTEMENTS):
            st.warning("Aucun mot de passe chef n'est configuré. Ajoutez ADMIN_PASSWORD ou les mots de passe CHEF_PASSWORD_* dans les Secrets.")
    else:
        if st.session_state.get("chief_is_master"):
            allowed_dept = None
            st.success("Session administrateur global : accès aux quatre départements.")
        else:
            allowed_dept = st.session_state.get("chief_department")
            st.success(f"Session chef de département : {allowed_dept}.")
        if st.button("Déconnexion de l'espace chef", key="chief_logout"):
            for key in ("chief_authenticated", "chief_department", "chief_is_master", "last_issued_request_id"):
                st.session_state.pop(key, None)
            st.rerun()

        dept_filter = allowed_dept
        if allowed_dept is None:
            dept_filter = st.selectbox("Département à traiter", list(DEPARTEMENTS.keys()), key="master_department_filter")
        try:
            pending = list_requests(department=dept_filter, status="En attente")
            issued = list_requests(department=dept_filter, status="Générée")
        except Exception as exc:
            st.error(f"Impossible de charger les demandes. Vérifiez la configuration de la base de données : {exc}")
            pending, issued = [], []

        m1, m2, m3 = st.columns(3)
        m1.metric("Demandes en attente", len(pending))
        m2.metric("Attestations générées", len(issued))
        m3.metric("Département", dept_filter or "Tous")

        # Émission groupée administrative depuis les fichiers sources, même sans demande individuelle.
        with st.expander("⚡ Génération directe en lot depuis le fichier source — sans demande préalable", expanded=False):
            st.warning(
                "Cette fonction émet des attestations officielles au nom du chef de département, même si l'enseignant "
                "n'a pas déposé de demande. Vérifiez les identités, le département, les données EDT et les informations "
                "communes avant de confirmer. Les champs inconnus resteront vides dans le document plutôt que d'être inventés."
            )
            if allowed_dept is None:
                generation_scope = st.selectbox(
                    "Périmètre de génération",
                    ["Tous les départements"] + list(DEPARTEMENTS.keys()),
                    index=0,
                    key="direct_source_generation_scope",
                    help="Pour tous les départements, le département est déduit des séances EDT quand cela est possible ; les cas ambigus doivent être attribués manuellement."
                )
            else:
                generation_scope = allowed_dept
                st.info(f"Votre compte de chef est limité au département **{allowed_dept}**. Seuls les enseignants qui peuvent être associés à ce département seront proposés.")

            source_people_batch = []
            batch_source_name = SOURCE_PERSONNEL_FILE.name
            batch_edt_name = EDT_SOURCE_FILE.name
            try:
                if SOURCE_PERSONNEL_FILE.exists():
                    source_people_batch = _load_source_people(SOURCE_PERSONNEL_FILE.read_bytes())
                else:
                    uploaded_batch_source = st.file_uploader(
                        "Téléverser le fichier source du personnel (.xlsx)", type=["xlsx"], key="direct_batch_source_upload"
                    )
                    if uploaded_batch_source is not None:
                        batch_source_name = uploaded_batch_source.name
                        source_people_batch = _load_source_people(uploaded_batch_source.getvalue())
                if EDT_SOURCE_FILE.exists():
                    batch_edt_sessions = _load_edt_cached(EDT_SOURCE_FILE.read_bytes())
                else:
                    uploaded_batch_edt = st.file_uploader(
                        "Téléverser le fichier EDT_Edition.xlsx", type=["xlsx"], key="direct_batch_edt_upload"
                    )
                    batch_edt_sessions = _load_edt_cached(uploaded_batch_edt.getvalue()) if uploaded_batch_edt is not None else []
                    if uploaded_batch_edt is not None:
                        batch_edt_name = uploaded_batch_edt.name
            except Exception as exc:
                source_people_batch, batch_edt_sessions = [], []
                st.error(f"Impossible de charger les fichiers sources : {exc}")

            if not source_people_batch:
                st.info("Le fichier source des coordonnées est absent ou ne contient aucune personne. Placez 'Vos coordonnées (réponses).xlsx' à côté de app.py.")
            elif not batch_edt_sessions:
                st.error("Le fichier EDT_Edition.xlsx est nécessaire pour associer les enseignants et proposer leur département. Aucun document n'a été généré.")
            else:
                try:
                    existing_source_records = list_requests(limit=10000)
                except Exception as exc:
                    st.error(f"Impossible de lire le registre des attestations : {exc}. La génération directe est bloquée pour éviter de créer des numéros ou dossiers en double.")
                    st.stop()

                # Réunir les demandes existantes par ligne/source afin de ne pas créer de doublons.
                existing_for_source: dict[int, list[dict]] = {}
                for old_record in existing_source_records:
                    old_snap = old_record.get("source_snapshot") or {}
                    old_row = old_snap.get("source_row")
                    try:
                        old_row_int = int(old_row)
                    except (TypeError, ValueError):
                        continue
                    old_file = _source_value(old_snap.get("source_file_name", ""))
                    if old_file and _norm_source_header(old_file) != _norm_source_header(batch_source_name):
                        continue
                    existing_for_source.setdefault(old_row_int, []).append(old_record)

                candidate_by_row: dict[int, dict] = {}
                candidate_rows: list[dict] = []
                ambiguous_count = 0
                for person in source_people_batch:
                    try:
                        source_row_no = int(person.get("source_row", 0))
                    except (TypeError, ValueError):
                        continue
                    sessions, matched_names, match_mode = match_teacher_sessions(person.get("full_name_fr", ""), batch_edt_sessions)
                    session_departments = sorted({
                        mapped for mapped in (_department_from_source(s.get("Département", "")) for s in sessions) if mapped
                    })
                    declared_department = _department_from_source(person.get("source_department", ""))
                    if declared_department:
                        suggested_department = declared_department
                        dept_method = "Source personnel"
                    elif len(session_departments) == 1:
                        suggested_department = session_departments[0]
                        dept_method = "Département EDT unique"
                    else:
                        suggested_department = ""
                        dept_method = "À vérifier manuellement" if len(session_departments) > 1 else "Aucune correspondance EDT"

                    if allowed_dept is not None:
                        # Un chef de département ne peut émettre que pour des personnes affectées
                        # à son département dans la source ou ayant des séances dans ce département.
                        if declared_department and declared_department != allowed_dept:
                            continue
                        if allowed_dept not in session_departments and declared_department != allowed_dept:
                            continue
                        assigned_for_table = allowed_dept
                    elif generation_scope != "Tous les départements":
                        target_dept = generation_scope
                        if declared_department and declared_department != target_dept:
                            continue
                        if not declared_department and target_dept not in session_departments:
                            continue
                        assigned_for_table = target_dept
                    else:
                        assigned_for_table = suggested_department
                        if not assigned_for_table:
                            ambiguous_count += 1

                    matching_existing = existing_for_source.get(source_row_no, [])
                    generated_same_source = [r for r in matching_existing if r.get("status") == "Générée" and r.get("attestation_number") not in (None, "")]
                    if generated_same_source:
                        status_label = "Déjà émise — numéro conservé"
                    elif matching_existing:
                        status_label = "Demande existante — peut être finalisée"
                    else:
                        status_label = "À générer"
                    date_ok = bool(clean_date_input(person.get("date_naissance", "")) or clean_date_input(_normalize_source_date(person.get("date_naissance", ""))))
                    identity_ok = bool(person.get("nom") and person.get("prenom") and date_ok and person.get("lieu_naissance"))
                    match_label = f"{match_mode} — {len(sessions)} séance(s)"
                    if match_mode == "approximatif":
                        match_label += " (à vérifier)"
                    candidate_by_row[source_row_no] = {
                        "person": person,
                        "sessions": sessions,
                        "matched_names": matched_names,
                        "match_mode": match_mode,
                        "session_departments": session_departments,
                        "suggested_department": suggested_department,
                        "department_method": dept_method,
                        "existing_records": matching_existing,
                        "identity_ok": identity_ok,
                    }
                    candidate_rows.append({
                        "Ligne source": source_row_no,
                        "Enseignant (fichier source)": person.get("full_name_fr", ""),
                        "Département à attribuer": assigned_for_table or "— Choisir —",
                        "Association EDT": match_label,
                        "Identité complète": "Oui" if identity_ok else "Non — à vérifier",
                        "État du registre": status_label,
                    })

                if allowed_dept is not None:
                    st.caption("La liste est limitée aux enseignants pour lesquels le département courant est indiqué dans la source ou retrouvé dans les séances EDT. Les enseignants sans rattachement fiable ne sont pas générés par ce compte.")
                elif generation_scope != "Tous les départements":
                    st.caption(f"La liste ci-dessous concerne les enseignants associés au département {generation_scope}. Les personnes sans rattachement EDT à ce département sont exclues.")
                else:
                    st.caption("Le fichier de coordonnées ne comporte pas de colonne Département. Les départements uniques détectés dans l'EDT sont proposés automatiquement ; pour les cas ambigus ou sans séance, attribuez le département dans la colonne avant de générer.")

                if candidate_rows:
                    batch_table = pd.DataFrame(candidate_rows)
                    scope_key = re.sub(r"[^A-Za-z0-9]+", "_", generation_scope).strip("_")
                    if allowed_dept is None and generation_scope == "Tous les départements":
                        edited_batch_table = st.data_editor(
                            batch_table,
                            use_container_width=True,
                            hide_index=True,
                            num_rows="fixed",
                            key=f"direct_batch_assignment_{scope_key}",
                            column_config={
                                "Ligne source": st.column_config.NumberColumn("Ligne source", disabled=True, width="small"),
                                "Enseignant (fichier source)": st.column_config.TextColumn(disabled=True, width="large"),
                                "Département à attribuer": st.column_config.SelectboxColumn(
                                    options=["— Choisir —"] + list(DEPARTEMENTS.keys()), required=True, width="medium"
                                ),
                                "Association EDT": st.column_config.TextColumn(disabled=True, width="medium"),
                                "Identité complète": st.column_config.TextColumn(disabled=True, width="medium"),
                                "État du registre": st.column_config.TextColumn(disabled=True, width="medium"),
                            },
                        )
                    else:
                        st.dataframe(batch_table, use_container_width=True, hide_index=True)
                        edited_batch_table = batch_table
                    st.caption(f"{len(edited_batch_table)} personne(s) proposées · {ambiguous_count} attribution(s) initialement ambiguë(s) dans la portée globale.")

                    # Champs administratifs communs du lot. Les matières/promotions S1 sont
                    # individualisées depuis l'EDT ; les valeurs S2 ci-dessous sont communes si renseignées.
                    batch_departments = list(DEPARTEMENTS.keys()) if generation_scope == "Tous les départements" and allowed_dept is None else [generation_scope if allowed_dept is None else allowed_dept]
                    batch_chef_defaults = {dept: _chef_par_defaut(dept) for dept in batch_departments}
                    with st.form(f"form_direct_generation_{scope_key}", clear_on_submit=False):
                        st.markdown("#### Formats de l'attestation à inclure")
                        bulk_languages = st.multiselect(
                            "Type(s) d'attestation à générer",
                            LANGUES,
                            default=[LANGUES[0]],
                            key=f"direct_batch_languages_{scope_key}",
                            help="Une seule attestation officielle et un seul numéro sont attribués par personne ; chaque langue sélectionnée produit une copie dans le ZIP avec le même numéro et la même date."
                        )
                        bulk_reference = st.text_input("Référence administrative à inscrire", value=REFERENCE_PAR_DEFAUT, key=f"direct_batch_ref_{scope_key}")
                        st.markdown("#### Informations de période communes (facultatives)")
                        st.caption("Les matières et promotions du semestre 1 restent propres à chaque enseignant et sont importées depuis EDT_Edition.xlsx. Les valeurs S2 saisies ici sont répétées dans chaque document du lot ; laissez-les vides si elles ne sont pas identiques pour tous.")
                        b1, b2 = st.columns(2)
                        with b1:
                            batch_s1_start = st.text_input("Début S1 (JJ/MM/AAAA)", key=f"direct_batch_s1_start_{scope_key}")
                            batch_s1_end = st.text_input("Fin S1 (JJ/MM/AAAA)", key=f"direct_batch_s1_end_{scope_key}")
                            batch_s1_year = st.text_input("Année universitaire S1", value=f"{date.today().year}/{date.today().year + 1}", key=f"direct_batch_s1_year_{scope_key}")
                        with b2:
                            batch_s2_subjects = st.text_area("Matière(s) S2 commune(s) — facultatif", key=f"direct_batch_s2_subjects_{scope_key}", height=68)
                            batch_s2_start = st.text_input("Début S2 (JJ/MM/AAAA)", key=f"direct_batch_s2_start_{scope_key}")
                            batch_s2_end = st.text_input("Fin S2 (JJ/MM/AAAA)", key=f"direct_batch_s2_end_{scope_key}")
                            batch_s2_promotions = st.text_input("Promotion(s) S2 commune(s) — facultatif", key=f"direct_batch_s2_promotions_{scope_key}")
                            batch_s2_year = st.text_input("Année universitaire S2", value=f"{date.today().year}/{date.today().year + 1}", key=f"direct_batch_s2_year_{scope_key}")
                        chef_fields: dict[str, tuple[str, str]] = {}
                        st.markdown("#### Signataires des départements concernés")
                        for dept in batch_departments:
                            default_fr, default_ar = batch_chef_defaults[dept]
                            if len(batch_departments) > 1:
                                st.markdown(f"**{dept}**")
                            cc1, cc2 = st.columns(2)
                            with cc1:
                                chef_fr_value = st.text_input("Nom du chef (français)" + (" *" if len(batch_departments) == 1 else " * pour les attestations de ce département"), value=default_fr, key=f"direct_batch_chef_fr_{scope_key}_{DEPT_SECRET_SUFFIX[dept]}")
                            with cc2:
                                chef_ar_value = st.text_input("Nom du chef (arabe)" + (" *" if len(batch_departments) == 1 else " * pour les attestations de ce département"), value=default_ar, key=f"direct_batch_chef_ar_{scope_key}_{DEPT_SECRET_SUFFIX[dept]}")
                            chef_fields[dept] = (chef_fr_value.strip(), chef_ar_value.strip())
                        confirm_direct_batch = st.checkbox(
                            "Je confirme que le chef de département autorise cette émission groupée sans demande individuelle et que les identités, les rapprochements EDT et les départements ont été vérifiés.",
                            key=f"direct_batch_confirm_{scope_key}"
                        )
                        direct_batch_submit = st.form_submit_button("📦 Générer les attestations et télécharger le ZIP", type="primary", use_container_width=True)

                    if direct_batch_submit:
                        errors: list[str] = []
                        if not bulk_languages:
                            errors.append("Choisissez au moins une langue : arabe, français ou bilingue.")
                        if not confirm_direct_batch:
                            errors.append("Cochez la confirmation d'émission administrative en lot.")
                        rows_to_process = []
                        for _, row in edited_batch_table.iterrows():
                            try:
                                row_no = int(row["Ligne source"])
                            except (TypeError, ValueError):
                                continue
                            dept_value = _source_value(row.get("Département à attribuer", ""))
                            if dept_value not in DEPARTEMENTS:
                                errors.append(f"{row.get('Enseignant (fichier source)', '')} : département non attribué, ligne {row_no}.")
                                continue
                            if allowed_dept is not None and dept_value != allowed_dept:
                                errors.append(f"{row.get('Enseignant (fichier source)', '')} : votre compte ne peut traiter que {allowed_dept}.")
                                continue
                            if generation_scope != "Tous les départements" and dept_value != generation_scope:
                                errors.append(f"{row.get('Enseignant (fichier source)', '')} : le département attribué ne correspond pas à la portée choisie.")
                                continue
                            detail = candidate_by_row.get(row_no)
                            if not detail:
                                continue
                            person = detail["person"]
                            dob = clean_date_input(person.get("date_naissance", "")) or clean_date_input(_normalize_source_date(person.get("date_naissance", "")))
                            birthplace = _source_value(person.get("lieu_naissance", ""))
                            if not (person.get("nom") and person.get("prenom") and dob and birthplace):
                                errors.append(f"{person.get('full_name_fr','')} : nom, prénom, date ou lieu de naissance incomplet.")
                                continue
                            existing_records_for_person = detail["existing_records"]
                            other_dept_record = next((r for r in existing_records_for_person if r.get("department") != dept_value), None)
                            same_dept_generated = next((r for r in existing_records_for_person if r.get("department") == dept_value and r.get("status") == "Générée" and r.get("attestation_number") not in (None, "")), None)
                            same_dept_pending = next((r for r in existing_records_for_person if r.get("department") == dept_value and r.get("status") != "Générée"), None)
                            if other_dept_record and not same_dept_generated and not same_dept_pending:
                                errors.append(f"{person.get('full_name_fr','')} : une demande/attestation existe déjà dans un autre département ; vérification nécessaire.")
                                continue
                            if same_dept_generated:
                                rows_to_process.append((row_no, dept_value, detail, same_dept_generated, None))
                            else:
                                rows_to_process.append((row_no, dept_value, detail, same_dept_pending, None))

                        departments_needing_signer = sorted({dept for _, dept, _, _, _ in rows_to_process}, key=lambda d: list(DEPARTEMENTS.keys()).index(d))
                        for dept in departments_needing_signer:
                            chef_fr_value, chef_ar_value = chef_fields.get(dept, ("", ""))
                            if not (chef_fr_value or chef_ar_value):
                                errors.append(f"Nom du chef de {dept} manquant en français et en arabe.")
                        if errors:
                            st.error("Certains éléments nécessitent une correction ou seront ignorés. " + " | ".join(errors[:12]) + (f" | … {len(errors)-12} autre(s) élément(s)." if len(errors) > 12 else ""))
                        fatal_errors = [e for e in errors if e.startswith("Nom du chef de ")]
                        blocked = (not bulk_languages) or (not confirm_direct_batch) or bool(fatal_errors)
                        if rows_to_process and not blocked:
                            newly_issued = 0
                            reused_issued = 0
                            failed = []
                            records_for_zip: dict[str, dict] = {}
                            progress = st.progress(0, text="Préparation des attestations…")
                            for index, (row_no, dept_value, detail, prior_record, _) in enumerate(rows_to_process, start=1):
                                person = detail["person"]
                                sessions = detail["sessions"]
                                matched_names = detail["matched_names"]
                                match_mode = detail["match_mode"]
                                snapshot = _build_source_snapshot(person, batch_source_name, sessions, matched_names, match_mode, batch_edt_name)
                                dob = clean_date_input(person.get("date_naissance", "")) or clean_date_input(_normalize_source_date(person.get("date_naissance", "")))
                                birthplace = _source_value(person.get("lieu_naissance", ""))
                                chef_fr_value, chef_ar_value = chef_fields[dept_value]
                                if prior_record and prior_record.get("status") == "Générée" and prior_record.get("attestation_number") not in (None, ""):
                                    issued_record = prior_record
                                    reused_issued += 1
                                else:
                                    try:
                                        if prior_record:
                                            request_record = prior_record
                                            # Garder l'identité/source existante, compléter le snapshot pour les futurs exports.
                                            if request_record.get("source_snapshot"):
                                                snapshot = {**request_record["source_snapshot"], **snapshot}
                                        else:
                                            request_record = create_request({
                                                "department": dept_value,
                                                "nom": person["nom"],
                                                "prenom": person["prenom"],
                                                "date_naissance": dob,
                                                "lieu_naissance": birthplace,
                                                "email": person.get("email", ""),
                                                "telephone": person.get("telephone", ""),
                                                "purpose": "Émission administrative groupée par le chef de département, sans demande préalable de l'intéressé.",
                                                "language_mode": bulk_languages[0],
                                                "source_snapshot": snapshot,
                                            })
                                        final_data = {
                                            "department": dept_value,
                                            "nom": person["nom"], "prenom": person["prenom"],
                                            "date_naissance": dob, "lieu_naissance": birthplace,
                                            "nom_ar": person.get("nom_ar", ""), "prenom_ar": person.get("prenom_ar", ""),
                                            "chef_fr": chef_fr_value, "chef_ar": chef_ar_value,
                                            "reference": bulk_reference.strip() or REFERENCE_PAR_DEFAUT,
                                            "matiere_s1": snapshot.get("edt_subjects_text") or person.get("matiere_s1", ""),
                                            "debut_s1": batch_s1_start.strip(), "fin_s1": batch_s1_end.strip(),
                                            "promotions_s1": snapshot.get("edt_promotions_text", ""),
                                            "section_s1": snapshot.get("edt_promotions_text", ""),
                                            "annee_s1": batch_s1_year.strip(),
                                            "matiere_s2": batch_s2_subjects.strip(),
                                            "debut_s2": batch_s2_start.strip(), "fin_s2": batch_s2_end.strip(),
                                            "promotions_s2": batch_s2_promotions.strip(), "section_s2": batch_s2_promotions.strip(),
                                            "annee_s2": batch_s2_year.strip(),
                                            "language_mode": bulk_languages[0],
                                            "emission_sans_demande": True,
                                            "source_nom_complet_fr": person.get("full_name_fr", ""),
                                            "source_nom_complet_ar": person.get("full_name_ar", ""),
                                            "type_enseignement_source": person.get("type_enseignement", ""),
                                            "dernier_diplome_source": person.get("dernier_diplome", ""),
                                            "edt_subjects": snapshot.get("edt_subjects", []),
                                            "edt_subjects_with_type": snapshot.get("edt_subjects_with_type", []),
                                            "edt_subjects_text": snapshot.get("edt_subjects_text", ""),
                                            "edt_promotions": snapshot.get("edt_promotions", []),
                                            "edt_promotions_text": snapshot.get("edt_promotions_text", ""),
                                            "edt_teacher_names": snapshot.get("edt_teacher_names", []),
                                            "edt_match_mode": snapshot.get("edt_match_mode", "aucun"),
                                            "edt_session_count": snapshot.get("edt_session_count", 0),
                                            "edt_source_file": batch_edt_name,
                                            "source_row": person.get("source_row"),
                                        }
                                        issued_record = issue_request(request_record["id"], final_data)
                                        if not issued_record or issued_record.get("status") != "Générée":
                                            raise RuntimeError("Le registre n'a pas confirmé l'émission officielle.")
                                        newly_issued += 1
                                    except Exception as exc:
                                        failed.append(f"{person.get('full_name_fr','')}: {exc}")
                                        continue
                                records_for_zip[str(issued_record.get("id"))] = issued_record
                                progress.progress(index / max(1, len(rows_to_process)), text=f"Traitement {index}/{len(rows_to_process)} : {person.get('full_name_fr','')}")
                            progress.empty()
                            if records_for_zip:
                                zip_data, files_count = _attestations_zip(list(records_for_zip.values()), language_modes=bulk_languages, save_to_archive=True)
                                st.success(f"Lot terminé : {newly_issued} nouvelle(s) émission(s), {reused_issued} attestation(s) déjà émises réutilisée(s), {len(failed)} échec(s). {files_count} document(s) Word dans le ZIP.")
                                if failed:
                                    with st.expander("Détails des échecs", expanded=False):
                                        st.write("\n".join(failed))
                                stamp = datetime.now(TZ_ALGERIA).strftime("%Y%m%d_%H%M")
                                dept_zip_name = "Tous_Departements" if generation_scope == "Tous les départements" else _safe_filename(generation_scope)
                                language_zip_name = "_".join({LANGUES[0]: "Bilingue", LANGUES[1]: "Arabe", LANGUES[2]: "Francais"}[x] for x in bulk_languages)
                                st.download_button(
                                    f"📥 Télécharger le ZIP des attestations ({files_count} document(s))",
                                    data=zip_data,
                                    file_name=f"Attestations_Source_{dept_zip_name}_{language_zip_name}_{stamp}.zip",
                                    mime="application/zip",
                                    type="primary",
                                    use_container_width=True,
                                    key=f"direct_batch_download_{scope_key}_{stamp}",
                                )
                                st.caption("Un seul numéro officiel est attribué par enseignant ; les différentes versions linguistiques d'un même document conservent le même numéro, la même date et la même référence. Les dossiers déjà générés ne reçoivent pas un nouveau numéro.")
                            else:
                                st.error("Aucune attestation n'a pu être incluse dans l'archive.")
                else:
                    st.info("Aucun enseignant ne correspond à la portée sélectionnée, ou aucun fichier source disponible.")

        tab_pending, tab_issued = st.tabs(["📥 Demandes à traiter", "✅ Attestations générées"])
        with tab_pending:
            if not pending:
                st.info("Aucune demande en attente pour ce département.")
            else:
                selected_id = st.selectbox(
                    "Choisir la demande à traiter",
                    [r["id"] for r in pending],
                    format_func=lambda rid: label_request(next(x for x in pending if x["id"] == rid)),
                    key=f"pending_request_{dept_filter or 'all'}",
                )
                request_record = get_request(selected_id)
                if request_record:
                    st.markdown("### Informations saisies par le demandeur")
                    info1, info2, info3 = st.columns(3)
                    info1.write(f"**Nom :** {request_record.get('nom','')}")
                    info1.write(f"**Prénom :** {request_record.get('prenom','')}")
                    info1.write(f"**Date de naissance :** {request_record.get('date_naissance','')}")
                    info2.write(f"**Lieu de naissance :** {request_record.get('lieu_naissance','')}")
                    info2.write(f"**E-mail :** {request_record.get('email','') or 'Non renseigné'}")
                    info2.write(f"**Téléphone :** {request_record.get('telephone','') or 'Non renseigné'}")
                    info3.write(f"**Langue demandée :** {request_record.get('language_mode', LANGUES[0])}")
                    info3.write(f"**Demande déposée :** {request_record.get('created_at','')}")
                    st.write("**Précisions du demandeur :** " + (request_record.get("purpose") or "Aucune"))
                    source_snapshot = request_record.get("source_snapshot") or {}
                    if source_snapshot:
                        st.markdown("#### Données récupérées du fichier source")
                        src1, src2, src3 = st.columns(3)
                        src1.write("**Nom complet (français) :** " + (source_snapshot.get("source_nom_complet_fr") or "Non renseigné"))
                        src1.write("**Nom complet (arabe) :** " + (source_snapshot.get("source_nom_complet_ar") or "Non renseigné"))
                        src2.write("**Type d'enseignement :** " + (source_snapshot.get("type_enseignement") or "Non renseigné"))
                        src2.write("**Dernier diplôme :** " + (source_snapshot.get("dernier_diplome") or "Non renseigné"))
                        src3.write("**Matière(s) S1 avec type :** " + (source_snapshot.get("edt_subjects_text") or source_snapshot.get("matiere_s1") or "Non renseignée"))
                        src3.write("**Promotion(s) EDT :** " + (source_snapshot.get("edt_promotions_text") or "Non renseignée"))
                        if source_snapshot.get("edt_session_count"):
                            st.success(f"EDT automatiquement associé : {source_snapshot.get('edt_session_count')} séance(s), {len(source_snapshot.get('edt_subjects') or [])} matière(s) distincte(s).")
                            if source_snapshot.get("edt_match_mode") == "approximatif":
                                st.warning("La correspondance du nom EDT était approximative : vérifiez la liste des matières avant de générer l'attestation.")
                            with st.expander("Voir les séances importées depuis EDT_Edition.xlsx"):
                                st.dataframe(pd.DataFrame(source_snapshot.get("edt_sessions") or []), use_container_width=True, hide_index=True)
                    st.divider()
                    st.markdown("### Compléter les informations administratives et pédagogiques")
                    st.caption("Le numéro et la date officiels sont attribués automatiquement au moment de la génération. Ils ne sont pas modifiables manuellement.")
                    chef_fr_default, chef_ar_default = _chef_par_defaut(request_record.get("department", ""))
                    source_snapshot = request_record.get("source_snapshot") or {}
                    with st.form(f"form_finalize_{selected_id}", clear_on_submit=False):
                        language_final = st.selectbox("Langue du document à émettre", LANGUES, index=LANGUES.index(request_record.get("language_mode", LANGUES[0])) if request_record.get("language_mode", LANGUES[0]) in LANGUES else 0, key=f"final_lang_{selected_id}")
                        c1, c2 = st.columns(2)
                        with c1:
                            chef_fr = st.text_input("Nom du chef de département (français) *", value=chef_fr_default, key=f"final_chef_fr_{selected_id}")
                            chef_ar = st.text_input("Nom du chef de département (arabe) *", value=chef_ar_default, key=f"final_chef_ar_{selected_id}")
                            reference = st.text_input("Référence administrative", value=REFERENCE_PAR_DEFAUT, key=f"final_ref_{selected_id}")
                        with c2:
                            st.text_input("Numéro de l'attestation (automatique)", value="Attribué à la génération", disabled=True, key=f"final_num_preview_{selected_id}")
                            st.text_input("Date de délivrance (automatique)", value=datetime.now(TZ_ALGERIA).strftime("%d/%m/%Y"), disabled=True, key=f"final_date_preview_{selected_id}")
                        st.markdown("#### Identité en arabe (préremplie depuis la source, modifiable)")
                        ca1, ca2 = st.columns(2)
                        with ca1:
                            nom_ar_final = st.text_input("Nom en arabe", value=source_snapshot.get("nom_ar", ""), key=f"final_nom_ar_{selected_id}")
                        with ca2:
                            prenom_ar_final = st.text_input("Prénom en arabe", value=source_snapshot.get("prenom_ar", ""), key=f"final_prenom_ar_{selected_id}")
                        st.markdown("#### Semestre 1")
                        if source_snapshot.get("edt_subjects_text"):
                            st.caption("Les matières S1 sont regroupées par type (Cours, TD, TP) d'après la colonne Code de EDT_Edition.xlsx. Vérifiez et complétez si nécessaire.")
                        elif source_snapshot.get("matiere_s1"):
                            st.warning("Aucune correspondance fiable avec EDT_Edition.xlsx : la matière déclarée dans le formulaire est affichée sans type confirmé. Associez le nom à l'EDT ou vérifiez les types manuellement.")
                        s11, s12, s13 = st.columns(3)
                        with s11:
                            s1_default = source_snapshot.get("edt_subjects_text") or source_snapshot.get("matiere_s1", "")
                            matiere_s1 = st.text_area("Matière(s) du semestre 1 — types Cours / TD / TP importés depuis EDT_Edition.xlsx", value=s1_default, height=120, key=f"final_s1_mat_{selected_id}")
                        with s12:
                            debut_s1 = st.text_input("Début de la période — S1 (JJ/MM/AAAA)", key=f"final_s1_start_{selected_id}")
                            fin_s1 = st.text_input("Fin de la période — S1 (JJ/MM/AAAA)", key=f"final_s1_end_{selected_id}")
                        with s13:
                            promotion_s1 = st.text_input("Promotion(s) — S1 (importées depuis EDT_Edition.xlsx)", value=source_snapshot.get("edt_promotions_text", ""), key=f"final_s1_promotions_{selected_id}")
                            annee_s1 = st.text_input("Année universitaire — S1", value=f"{date.today().year}/{date.today().year + 1}", key=f"final_s1_year_{selected_id}")
                        st.markdown("#### Semestre 2")
                        s21, s22, s23 = st.columns(3)
                        with s21:
                            matiere_s2 = st.text_area("Matière(s) du semestre 2", height=90, key=f"final_s2_mat_{selected_id}")
                        with s22:
                            debut_s2 = st.text_input("Début de la période — S2 (JJ/MM/AAAA)", key=f"final_s2_start_{selected_id}")
                            fin_s2 = st.text_input("Fin de la période — S2 (JJ/MM/AAAA)", key=f"final_s2_end_{selected_id}")
                        with s23:
                            promotion_s2 = st.text_input("Promotion(s) — S2", key=f"final_s2_promotions_{selected_id}", help="À compléter si les promotions du semestre 2 diffèrent de celles de l'EDT source disponible.")
                            annee_s2 = st.text_input("Année universitaire — S2", value=f"{date.today().year}/{date.today().year + 1}", key=f"final_s2_year_{selected_id}")
                        issue_button = st.form_submit_button("🖨️ Générer l'attestation officielle et attribuer le numéro", type="primary", use_container_width=True)
                    if issue_button:
                        if not (chef_fr.strip() or chef_ar.strip()):
                            st.error("Renseignez le nom du chef en français ou en arabe.")
                        else:
                            final_data = {
                                "department": request_record["department"],
                                "nom": request_record["nom"],
                                "prenom": request_record["prenom"],
                                "date_naissance": request_record["date_naissance"],
                                "lieu_naissance": request_record["lieu_naissance"],
                                "chef_fr": chef_fr.strip(), "chef_ar": chef_ar.strip(),
                                "reference": reference.strip() or REFERENCE_PAR_DEFAUT,
                                "matiere_s1": matiere_s1.strip(), "debut_s1": debut_s1.strip(), "fin_s1": fin_s1.strip(),
                                "promotions_s1": promotion_s1.strip(), "section_s1": promotion_s1.strip(), "annee_s1": annee_s1.strip(),
                                "matiere_s2": matiere_s2.strip(), "debut_s2": debut_s2.strip(), "fin_s2": fin_s2.strip(),
                                "promotions_s2": promotion_s2.strip(), "section_s2": promotion_s2.strip(), "annee_s2": annee_s2.strip(),
                                "language_mode": language_final,
                                "nom_ar": nom_ar_final.strip(),
                                "prenom_ar": prenom_ar_final.strip(),
                                "source_nom_complet_fr": source_snapshot.get("source_nom_complet_fr", ""),
                                "source_nom_complet_ar": source_snapshot.get("source_nom_complet_ar", ""),
                                "type_enseignement_source": source_snapshot.get("type_enseignement", ""),
                                "dernier_diplome_source": source_snapshot.get("dernier_diplome", ""),
                                "edt_subjects": source_snapshot.get("edt_subjects", []),
                                "edt_subjects_with_type": source_snapshot.get("edt_subjects_with_type", []),
                                "edt_subjects_text": source_snapshot.get("edt_subjects_text", ""),
                                "edt_promotions": source_snapshot.get("edt_promotions", []),
                                "edt_promotions_text": source_snapshot.get("edt_promotions_text", ""),
                                "edt_teacher_names": source_snapshot.get("edt_teacher_names", []),
                                "edt_session_count": source_snapshot.get("edt_session_count", 0),
                                "edt_source_file": source_snapshot.get("edt_source_file", "EDT_Edition.xlsx"),
                            }
                            try:
                                issued_record = issue_request(selected_id, final_data)
                                if not issued_record:
                                    raise RuntimeError("Le registre n'a pas retourné le dossier généré.")
                                values = values_from_record(issued_record)
                                docx_bytes = generate_attestation(values, store_language_mode(issued_record))
                                safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{values['nom']}_{values['prenom']}").strip("_") or "Personnel"
                                filename = f"Attestation_{safe_name}_{values['department']}_{values['numero']}_{issued_record.get('issued_year','')}.docx"
                                archive_result = _archive_docx_record(issued_record, store_language_mode(issued_record), content=docx_bytes)
                                if not archive_result.get("persistent", False):
                                    st.session_state["archive_save_notice"] = "Le document est généré et sauvegardé localement, mais l'archive cloud n'a pas pu être enregistrée. Vérifiez la migration SQL des archives."
                                st.session_state["last_issued_document"] = {
                                    "request_id": selected_id,
                                    "bytes": docx_bytes,
                                    "filename": filename,
                                    "number": values["numero"],
                                    "year": issued_record.get("issued_year", ""),
                                    "date": values["date_attestation"],
                                    "department": values["department"],
                                }
                                st.session_state["last_issued_request_id"] = selected_id
                                st.success(f"Attestation générée. Numéro attribué : {values['numero']}/{issued_record.get('issued_year','')} — Date : {values['date_attestation']}.")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"Erreur pendant la finalisation : {exc}")

        with tab_issued:
            zip_scope_department = allowed_dept if allowed_dept is not None else None
            try:
                zip_issued_records = list_requests(department=zip_scope_department, status="Générée", limit=10000)
            except Exception as exc:
                zip_issued_records = []
                st.error(f"Impossible de préparer l'archive des attestations : {exc}")
            if zip_issued_records:
                bulk_languages = st.multiselect(
                    "Formats à inclure dans le ZIP (vous pouvez en choisir un, deux ou les trois)",
                    LANGUES,
                    default=list(LANGUES),
                    key=f"bulk_attestation_languages_{allowed_dept or 'all'}",
                    help="Chaque attestation sélectionnée sera générée dans chacun des formats choisis : arabe, français et/ou bilingue. Les numéros, dates et références existants restent identiques.",
                )
                if not bulk_languages:
                    st.info("Sélectionnez au moins un format de langue pour préparer l'archive.")
                else:
                    zip_data, zip_count = _attestations_zip(zip_issued_records, language_modes=bulk_languages)
                    zip_label = "tous les départements" if allowed_dept is None else f"département {allowed_dept}"
                    language_label = "_".join({LANGUES[0]: "Bilingue", LANGUES[1]: "Arabe", LANGUES[2]: "Francais"}[item] for item in bulk_languages)
                    st.download_button(
                        f"📦 Télécharger les attestations Word en ZIP ({zip_count} fichier(s) — {len(zip_issued_records)} attestation(s) — {zip_label})",
                        data=zip_data,
                        file_name=f"Attestations_{'Tous_Departements' if allowed_dept is None else _safe_filename(allowed_dept)}_{language_label}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True,
                        key=f"zip_attestations_{allowed_dept or 'all'}",
                    )
                    st.caption("Pour chaque attestation déjà émise, l'archive contient une copie pour chaque langue choisie. Les numéros, dates et références ne sont pas réattribués. Les demandes encore en attente doivent d'abord être complétées et validées par le chef.")
                    if st.button("💾 Sauvegarder ces versions et ce ZIP dans le dossier d'archives", key=f"save_multilang_archive_{allowed_dept or 'all'}"):
                        try:
                            _saved_zip_data, _saved_zip_count = _attestations_zip(
                                zip_issued_records, language_modes=bulk_languages, save_to_archive=True
                            )
                            st.session_state["archive_saved_notice"] = f"{_saved_zip_count} document(s) Word et le ZIP ont été enregistrés dans les archives."
                            st.rerun()
                        except Exception as exc:
                            st.error(f"La sauvegarde a échoué : {exc}")
            else:
                st.info("Aucune attestation finalisée à inclure dans l'archive pour ce périmètre.")
            if not issued:
                st.info("Aucune attestation générée dans le département actuellement sélectionné.")
            else:
                generated_id = st.selectbox(
                    "Choisir une attestation finalisée",
                    [r["id"] for r in issued],
                    format_func=lambda rid: label_request(next(x for x in issued if x["id"] == rid)) + f" — N° {next(x for x in issued if x['id'] == rid).get('attestation_number','')}/{next(x for x in issued if x['id'] == rid).get('issued_year','')}",
                    key=f"issued_request_{dept_filter or 'all'}",
                )
                gen_record = get_request(generated_id)
                if gen_record:
                    gv = values_from_record(gen_record)
                    st.write(f"**Numéro :** {gv.get('numero')} / {gen_record.get('issued_year')} · **Date :** {gv.get('date_attestation')}")
                    render_download(gen_record, f"redownload_issued_{generated_id}")
                    if st.button("💾 Sauvegarder cette attestation dans le dossier d'archives", key=f"archive_single_issued_{generated_id}"):
                        try:
                            chosen_language = store_language_mode(gen_record)
                            chosen_values = values_from_record(gen_record)
                            chosen_docx = generate_attestation(chosen_values, chosen_language)
                            saved_one = _archive_docx_record(gen_record, chosen_language, content=chosen_docx)
                            if saved_one.get("persistent", False):
                                st.success("Attestation enregistrée dans le dossier des générations sauvegardées.")
                            else:
                                st.warning("Le fichier est sauvegardé localement, mais pas dans la base persistante. Vérifiez le schéma SQL des archives.")
                        except Exception as exc:
                            st.error(f"Impossible d'archiver cette attestation : {exc}")

        # Dossier d'archives indépendant de l'historique des demandes : les Word et ZIP générés y restent sauvegardés.
        with st.expander("💾 Dossier des générations sauvegardées", expanded=False):
            st.write("Chaque attestation Word générée et chaque ZIP préparé sont archivés ici. Effacer l'historique des demandes ne supprime pas ce dossier et ne remet pas les numéros à zéro.")
            if allowed_dept is None:
                archive_scope_label = st.selectbox(
                    "Filtrer les archives par département",
                    ["Tous les départements"] + list(DEPARTEMENTS.keys()),
                    key="generation_archive_scope_filter",
                )
                archive_department_filter = None if archive_scope_label == "Tous les départements" else archive_scope_label
            else:
                archive_department_filter = allowed_dept
                st.caption(f"Archives accessibles pour votre département : {allowed_dept}.")
            try:
                saved_files = list_generation_files(department=archive_department_filter, limit=5000)
            except Exception as exc:
                saved_files = []
                st.error(
                    "Impossible de lire les archives persistantes. Si l'application est sur Streamlit Cloud, "
                    "exécutez le script supabase_schema.sql mis à jour dans Supabase > SQL Editor. "
                    f"Détail technique : {exc}"
                )
            if saved_files:
                archive_rows = [{
                    "Date de sauvegarde": str(item.get("created_at", ""))[:19].replace("T", " "),
                    "Type": item.get("file_kind", ""),
                    "Département": item.get("department", ""),
                    "Personne / lot": item.get("person_name", ""),
                    "N° officiel": f"{item.get('attestation_number','')}/{item.get('issued_year','')}" if item.get("attestation_number") else "—",
                    "Langue(s)": item.get("language_mode", ""),
                    "Fichier": item.get("file_name", ""),
                    "Taille (Ko)": round((int(item.get("byte_size") or 0) / 1024), 1),
                } for item in saved_files]
                st.metric("Fichiers sauvegardés", len(saved_files))
                st.dataframe(pd.DataFrame(archive_rows), use_container_width=True, hide_index=True)
                archive_by_id = {str(item["id"]): item for item in saved_files}
                chosen_archive_id = st.selectbox(
                    "Choisir un fichier archivé à télécharger",
                    list(archive_by_id),
                    format_func=lambda file_id: (
                        f"{archive_by_id[file_id].get('file_name','')} — "
                        f"{archive_by_id[file_id].get('person_name','')} — "
                        f"{archive_by_id[file_id].get('created_at','')[:10]}"
                    ),
                    key=f"chosen_saved_generation_{archive_department_filter or 'all'}",
                )
                chosen_archive = get_generation_file(chosen_archive_id) if chosen_archive_id else None
                if chosen_archive and chosen_archive.get("content"):
                    st.download_button(
                        "📥 Télécharger le fichier sélectionné",
                        data=chosen_archive["content"],
                        file_name=chosen_archive.get("file_name") or "generation.bin",
                        mime="application/zip" if str(chosen_archive.get("file_kind", "")).upper().startswith("ZIP") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key=f"download_saved_generation_{chosen_archive_id}",
                    )
                if st.button("📦 Préparer une sauvegarde ZIP du dossier d'archives", type="secondary", key=f"prepare_generation_archive_zip_{archive_department_filter or 'all'}"):
                    try:
                        archive_buf = BytesIO()
                        with ZipFile(archive_buf, "w", compression=ZIP_DEFLATED) as archive:
                            added_count = 0
                            for item in saved_files:
                                full_item = get_generation_file(str(item["id"]))
                                if not full_item or not full_item.get("content"):
                                    continue
                                dept_folder = _safe_filename(str(item.get("department") or "Sans_Departement"))
                                year_folder = str(item.get("issued_year") or datetime.now(TZ_ALGERIA).year)
                                arcname = f"{year_folder}/{dept_folder}/{item.get('file_name') or 'generation.bin'}"
                                archive.writestr(arcname, full_item["content"])
                                added_count += 1
                            archive.writestr(
                                "LISEZ-MOI.txt",
                                "Sauvegarde des fichiers générés par l'application d'attestations UDL-SBA.\n"
                                f"Fichiers inclus : {added_count}.\n"
                                "Cette archive est indépendante de l'historique des demandes.\n"
                            )
                        st.session_state["generation_archive_download"] = {
                            "content": archive_buf.getvalue(),
                            "file_name": f"Dossier_Generations_{_safe_filename(archive_scope_label if allowed_dept is None else allowed_dept)}_{datetime.now(TZ_ALGERIA).strftime('%Y%m%d_%H%M%S')}.zip",
                            "count": added_count,
                            "scope_key": archive_department_filter or "all",
                        }
                    except Exception as exc:
                        st.error(f"Impossible de préparer la sauvegarde ZIP : {exc}")
                archive_package = st.session_state.get("generation_archive_download")
                if archive_package and archive_package.get("scope_key") == (archive_department_filter or "all"):
                    st.download_button(
                        f"⬇️ Télécharger la sauvegarde ZIP ({archive_package['count']} fichier(s))",
                        data=archive_package["content"],
                        file_name=archive_package["file_name"],
                        mime="application/zip",
                        type="primary",
                        use_container_width=True,
                        key=f"download_generation_archive_zip_{archive_package['file_name']}",
                    )
            else:
                st.info("Aucun fichier généré n'est encore archivé pour ce périmètre.")
            archive_dir_path = APP_DIR / "archives_generations"
            st.caption(f"Dossier local de sauvegarde : `{archive_dir_path}`. Sur Streamlit Cloud, les fichiers sont également conservés dans la base Supabase après l'installation du schéma SQL mis à jour.")

        # Suppression des demandes : réservée à l'administrateur global, séparée du dossier d'archives.
        if st.session_state.get("chief_is_master"):
            with st.expander("⚠️ Zone sensible — Effacer tout l'historique des demandes", expanded=False):
                st.error(
                    "Cette action supprime définitivement toutes les demandes (en attente et générées) du registre. "
                    "Elle NE supprime PAS le dossier des générations sauvegardées et NE remet PAS les numéros officiels à zéro."
                )
                confirm_history_clear = st.checkbox(
                    "J'ai compris que toutes les demandes du registre, dans les quatre départements, seront supprimées.",
                    key="confirm_clear_all_attestation_history",
                )
                confirm_history_text = st.text_input(
                    "Pour confirmer, saisissez exactement : EFFACER TOUT L'HISTORIQUE",
                    key="confirm_clear_all_attestation_history_text",
                )
                if st.button(
                    "🗑️ Effacer tout l'historique des demandes",
                    type="primary",
                    disabled=not (confirm_history_clear and confirm_history_text.strip() == "EFFACER TOUT L'HISTORIQUE"),
                    key="clear_all_attestation_history_button",
                ):
                    try:
                        deleted_count = delete_all_request_history()
                        for session_key in ("last_issued_document", "last_issued_request_id", "generation_archive_download"):
                            st.session_state.pop(session_key, None)
                        st.session_state["history_cleared_count_notice"] = deleted_count
                        st.session_state["confirm_clear_all_attestation_history"] = False
                        st.session_state["confirm_clear_all_attestation_history_text"] = ""
                        st.rerun()
                    except Exception as exc:
                        st.error(f"L'historique n'a pas pu être effacé : {exc}")
        if st.session_state.pop("history_cleared_count_notice", None) is not None:
            st.success("Historique des demandes effacé. Les archives générées et les compteurs officiels sont conservés.")
        archive_notice = st.session_state.pop("archive_save_notice", None)
        if archive_notice:
            st.warning(archive_notice)
        archive_saved_notice = st.session_state.pop("archive_saved_notice", None)
        if archive_saved_notice:
            st.success(archive_saved_notice)

        last_doc = st.session_state.get("last_issued_document")
        if last_doc:
            st.divider()
            st.subheader("📄 Dernière attestation générée dans cette session")
            st.write(f"**{last_doc['department']}** · N° **{last_doc['number']}/{last_doc['year']}** · Date **{last_doc['date']}**")
            st.download_button(
                "📥 Télécharger le document Word généré",
                data=last_doc["bytes"],
                file_name=last_doc["filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key=f"download_latest_{last_doc['request_id']}",
            )

st.divider()
