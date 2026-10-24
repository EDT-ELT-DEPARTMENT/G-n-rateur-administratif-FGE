# -*- coding: utf-8 -*-
"""Workflow de demandes et de génération d'attestations administratives UDL-SBA."""
from __future__ import annotations

from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
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

APP_DIR = Path(__file__).resolve().parent
TZ_ALGERIA = ZoneInfo("Africa/Algiers")
LANGUES = ["Bilingue (arabe puis français)", "Arabe (conforme au modèle PDF)", "Français"]
REFERENCE_PAR_DEFAUT = f"UDL-GEL-ER-120-{datetime.now(TZ_ALGERIA).year}"
SOURCE_PERSONNEL_FILE = APP_DIR / "Vos coordonnées (réponses).xlsx"

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
        st.info(
            f"**Nom en français :** {selected_person['full_name_fr']}  \n"
            f"**Nom en arabe :** {selected_person.get('full_name_ar') or 'Non renseigné'}  \n"
            f"**Matière(s) source S1 :** {selected_person.get('matiere_s1') or 'Non renseignée'}  \n"
            f"**Type d'enseignement :** {selected_person.get('type_enseignement') or 'Non renseigné'}  ·  "
            f"**Dernier diplôme :** {selected_person.get('dernier_diplome') or 'Non renseigné'}"
        )

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
                        "source_nom_complet_ar": selected_person.get("full_name_ar", ""),
                        "nom_ar": selected_person.get("nom_ar", ""),
                        "prenom_ar": selected_person.get("prenom_ar", ""),
                        "matiere_s1": selected_person.get("matiere_s1", ""),
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
                        src3.write("**Matière(s) S1 dans la source :** " + (source_snapshot.get("matiere_s1") or "Non renseignée"))
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
                        if source_snapshot.get("matiere_s1"):
                            st.caption("La matière S1 est préremplie depuis le fichier source. Vérifiez et complétez si nécessaire.")
                        s11, s12, s13 = st.columns(3)
                        with s11:
                            matiere_s1 = st.text_area("Matière(s) du semestre 1", value=source_snapshot.get("matiere_s1", ""), height=90, key=f"final_s1_mat_{selected_id}")
                        with s12:
                            debut_s1 = st.text_input("Début de la période — S1 (JJ/MM/AAAA)", key=f"final_s1_start_{selected_id}")
                            fin_s1 = st.text_input("Fin de la période — S1 (JJ/MM/AAAA)", key=f"final_s1_end_{selected_id}")
                        with s13:
                            section_s1 = st.text_input("Section / niveau / groupe — S1", key=f"final_s1_section_{selected_id}")
                            annee_s1 = st.text_input("Année universitaire — S1", value=f"{date.today().year}/{date.today().year + 1}", key=f"final_s1_year_{selected_id}")
                        st.markdown("#### Semestre 2")
                        s21, s22, s23 = st.columns(3)
                        with s21:
                            matiere_s2 = st.text_area("Matière(s) du semestre 2", height=90, key=f"final_s2_mat_{selected_id}")
                        with s22:
                            debut_s2 = st.text_input("Début de la période — S2 (JJ/MM/AAAA)", key=f"final_s2_start_{selected_id}")
                            fin_s2 = st.text_input("Fin de la période — S2 (JJ/MM/AAAA)", key=f"final_s2_end_{selected_id}")
                        with s23:
                            section_s2 = st.text_input("Section / niveau / groupe — S2", key=f"final_s2_section_{selected_id}")
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
                                "section_s1": section_s1.strip(), "annee_s1": annee_s1.strip(),
                                "matiere_s2": matiere_s2.strip(), "debut_s2": debut_s2.strip(), "fin_s2": fin_s2.strip(),
                                "section_s2": section_s2.strip(), "annee_s2": annee_s2.strip(),
                                "language_mode": language_final,
                                "nom_ar": nom_ar_final.strip(),
                                "prenom_ar": prenom_ar_final.strip(),
                                "source_nom_complet_fr": source_snapshot.get("source_nom_complet_fr", ""),
                                "source_nom_complet_ar": source_snapshot.get("source_nom_complet_ar", ""),
                                "type_enseignement_source": source_snapshot.get("type_enseignement", ""),
                                "dernier_diplome_source": source_snapshot.get("dernier_diplome", ""),
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
            if not issued:
                st.info("Aucune attestation générée pour le moment.")
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
- Téléversez `app.py`, `attestation_generator.py`, `storage.py`, `requirements.txt`, `assets/logo_udl.png`, `Modele_Source_Donnees_Attestations.xlsx`, `.streamlit/config.toml`, `.streamlit/secrets.toml.example` et `supabase_schema.sql` dans votre dépôt GitHub.
- Dans Streamlit Community Cloud, sélectionnez ce dépôt, la branche `main` et `app.py` comme fichier principal.
- Dans **Settings → Secrets**, configurez les mots de passe des chefs. Pour le stockage permanent Cloud, ajoutez `SUPABASE_URL` et `SUPABASE_SERVICE_ROLE_KEY`, puis exécutez `supabase_schema.sql` dans le SQL Editor de Supabase.
- Pour les notifications par e-mail, renseignez `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`, `SMTP_FROM` et `CHEF_EMAIL_*`. Sans SMTP, la demande reste enregistrée et visible dans l'espace du chef, mais aucun e-mail automatique ne sera envoyé.

### Numérotation
Le numéro est attribué au moment de la génération officielle et suit un compteur annuel partagé entre les quatre départements. Le numéro/date sont conservés quand on retélécharge un document déjà émis; un simple téléchargement ne consomme pas un nouveau numéro.

### Données personnelles
N'ajoutez aucun fichier de personnel confidentiel au dépôt GitHub. La clé Supabase service-role doit rester dans les Secrets et ne jamais être intégrée au code ou au dépôt. Utilisez un dépôt privé et donnez l'accès seulement aux personnes autorisées.
""")

st.caption("Le document Word est produit à partir du modèle administratif transmis. Les informations pédagogiques et la date/numérotation officielles sont renseignées ou attribuées par le chef lors de l'émission.")
