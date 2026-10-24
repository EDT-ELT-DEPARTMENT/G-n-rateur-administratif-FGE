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
from storage import backend_mode, create_request, get_request, get_setting, issue_request, list_requests
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


def _attestations_zip(records: list[dict], language_mode: str | None = None) -> tuple[bytes, int]:
    """Regroupe les attestations déjà émises dans la langue choisie, sans réattribuer de numéro."""
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
            # La langue choisie pour le lot modifie seulement la copie dans le ZIP.
            # Le numéro, la date et la référence de l'attestation déjà émise restent inchangés.
            docx = generate_attestation(values, language_mode or store_language_mode(record))
            base = _safe_filename(
                f"Attestation_{values.get('nom','')}_{values.get('prenom','')}_"
                f"{values.get('department','')}_{values['numero']}_{record.get('issued_year','')}"
            )
            filename = base + ".docx"
            if filename in used_names:
                filename = f"{base}_{str(record.get('id',''))[:8]}.docx"
            used_names.add(filename)
            archive.writestr(filename, docx)
            count += 1
        archive.writestr(
            "LISEZ-MOI.txt",
            "Archive des attestations administratives déjà émises.\n"
            "Les demandes encore en attente sont exclues : elles n'ont pas de numéro officiel ni de document final.\n"
            f"Nombre de documents : {count}\n",
        )
    return out.getvalue(), count


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
        selected_source_idx = st.selectbox(
            "👤 Sélectionner le demandeur dans le fichier source *",
            options=list(range(len(source_people))),
            format_func=lambda idx: (
                source_people[idx]["full_name_fr"] +
                (f" — {source_people[idx]['full_name_ar']}" if source_people[idx].get("full_name_ar") else "")
            ),
            key="source_person_selection",
            help="La sélection remplit automatiquement les champs disponibles ci-dessous. Les champs restent modifiables pour corriger une éventuelle erreur de saisie source.",
        )
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
                bulk_language = st.selectbox(
                    "Langue pour toutes les attestations de l'archive ZIP",
                    LANGUES,
                    index=0,
                    key=f"bulk_attestation_language_{allowed_dept or 'all'}",
                    help="La langue choisie s'applique aux copies Word placées dans le ZIP. Les numéros, dates et références déjà attribués sont conservés.",
                )
                zip_data, zip_count = _attestations_zip(zip_issued_records, language_mode=bulk_language)
                zip_label = "tous les départements" if allowed_dept is None else f"département {allowed_dept}"
                st.download_button(
                    f"📦 Télécharger toutes les attestations Word en ZIP ({zip_count} document(s) — {zip_label})",
                    data=zip_data,
                    file_name=("Attestations_Tous_Departements.zip" if allowed_dept is None else f"Attestations_{_safe_filename(allowed_dept)}.zip"),
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                    key=f"zip_attestations_{allowed_dept or 'all'}",
                )
                st.caption("L'archive contient toutes les attestations déjà émises pour le périmètre autorisé, dans la langue sélectionnée. Chaque document conserve son numéro officiel, sa date de délivrance et sa référence ; le téléchargement groupé ne consomme pas de nouveaux numéros. Les demandes en attente doivent d'abord être complétées et émises.")
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
with st.expander("ℹ️ Configuration, stockage durable et déploiement"):
    st.markdown("""
### Déploiement
- Téléversez `app.py`, `attestation_generator.py`, `edt_individuel.py`, `storage.py`, `requirements.txt`, `assets/logo_udl.png`, `EDT_Edition.xlsx`, `Vos coordonnées (réponses).xlsx` (dépôt privé uniquement), `Modele_Source_Donnees_Attestations.xlsx`, `.streamlit/config.toml`, `.streamlit/secrets.toml.example` et `supabase_schema.sql` dans votre dépôt GitHub.
- Dans Streamlit Community Cloud, sélectionnez ce dépôt, la branche `main` et `app.py` comme fichier principal.
- Dans **Settings → Secrets**, configurez les mots de passe des chefs. Pour le stockage permanent Cloud, ajoutez `SUPABASE_URL` et `SUPABASE_SERVICE_ROLE_KEY`, puis exécutez `supabase_schema.sql` dans le SQL Editor de Supabase.
- Pour les notifications par e-mail, renseignez `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `SMTP_FROM` et `CHEF_EMAIL_*`. Sans SMTP, la demande reste enregistrée et visible dans l'espace du chef, mais aucun e-mail automatique ne sera envoyé.

### Numérotation
Le numéro est attribué au moment de la génération officielle et suit un compteur annuel partagé entre les quatre départements. Le numéro/date sont conservés quand on retélécharge un document déjà émis; un simple téléchargement ne consomme pas un nouveau numéro.

### Données personnelles
N'ajoutez aucun fichier de personnel confidentiel au dépôt GitHub. La clé Supabase service-role doit rester dans les Secrets et ne jamais être intégrée au code ou au dépôt. Utilisez un dépôt privé et donnez l'accès seulement aux personnes autorisées.
""")

st.caption("Le document Word reprend les matières identifiées dans EDT_Edition.xlsx lorsqu'une correspondance enseignant est trouvée. Les autres informations pédagogiques et la date/numérotation officielles sont vérifiées ou attribuées par le chef lors de l'émission.")
