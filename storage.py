# -*- coding: utf-8 -*-
"""Stockage des demandes : Supabase pour le cloud, SQLite pour le mode local."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import json
import os
import sqlite3
import uuid
from typing import Any

import streamlit as st

APP_DIR = Path(__file__).resolve().parent
SQLITE_PATH = APP_DIR / "attestation_requests.sqlite3"
TZ_ALGERIA = ZoneInfo("Africa/Algiers")


def now_local() -> datetime:
    return datetime.now(TZ_ALGERIA)


def _is_placeholder(value: str) -> bool:
    upper = str(value or "").upper()
    return any(marker in upper for marker in ("REMPLACER", "VOTRE-PROJET", "VOTRE_PROJET", "VOTRE_CLE", "VOTRE-APPLICATION", "VOTRE-ETABLISSEMENT", "VOTRE_MOT_DE_PASSE", "MOT_DE_PASSE_", "EXEMPLE.DZ", "SMTP.VOTRE-SERVEUR", "COMPTE-NOTIFICATION@"))


def get_setting(name: str, default: str = "") -> str:
    value = os.environ.get(name, "")
    try:
        secret_value = str(st.secrets.get(name, "") or "").strip()
        if secret_value and not _is_placeholder(secret_value):
            value = secret_value
    except Exception:
        pass
    cleaned = str(value or default).strip()
    return "" if _is_placeholder(cleaned) else cleaned


def backend_mode() -> str:
    if get_setting("SUPABASE_URL") and (get_setting("SUPABASE_SERVICE_ROLE_KEY") or get_setting("SUPABASE_KEY")):
        return "supabase"
    return "sqlite"


@st.cache_resource(show_spinner=False)
def _supabase_client(url: str, key: str):
    from supabase import create_client
    return create_client(url, key)


def _client():
    url = get_setting("SUPABASE_URL")
    key = get_setting("SUPABASE_SERVICE_ROLE_KEY") or get_setting("SUPABASE_KEY")
    if not url or not key:
        return None
    return _supabase_client(url, key)


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(SQLITE_PATH, timeout=30, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS attestation_requests (
            id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            department TEXT NOT NULL,
            nom TEXT NOT NULL,
            prenom TEXT NOT NULL,
            date_naissance TEXT NOT NULL,
            lieu_naissance TEXT NOT NULL,
            email TEXT NOT NULL DEFAULT '',
            telephone TEXT NOT NULL DEFAULT '',
            purpose TEXT NOT NULL DEFAULT '',
            language_mode TEXT NOT NULL DEFAULT 'Bilingue (arabe puis français)',
            status TEXT NOT NULL DEFAULT 'En attente',
            final_data TEXT,
            attestation_number INTEGER,
            issued_date TEXT,
            issued_year INTEGER,
            UNIQUE(issued_year, attestation_number)
        );
        CREATE TABLE IF NOT EXISTS attestation_sequences (
            issue_year INTEGER PRIMARY KEY,
            last_number INTEGER NOT NULL
        );
    """)
    return conn


def _sqlite_row(row: sqlite3.Row | None) -> dict[str, Any] | None:
    if row is None:
        return None
    data = dict(row)
    if data.get("final_data"):
        try:
            data["final_data"] = json.loads(data["final_data"])
        except (TypeError, json.JSONDecodeError):
            data["final_data"] = {}
    else:
        data["final_data"] = {}
    return data


def create_request(payload: dict[str, Any]) -> dict[str, Any]:
    request_id = str(uuid.uuid4())
    now = now_local().isoformat(timespec="seconds")
    data = {
        "id": request_id,
        "created_at": now,
        "updated_at": now,
        "department": str(payload["department"]),
        "nom": str(payload["nom"]).strip(),
        "prenom": str(payload["prenom"]).strip(),
        "date_naissance": str(payload["date_naissance"]).strip(),
        "lieu_naissance": str(payload["lieu_naissance"]).strip(),
        "email": str(payload.get("email", "")).strip(),
        "telephone": str(payload.get("telephone", "")).strip(),
        "purpose": str(payload.get("purpose", "")).strip(),
        "language_mode": str(payload.get("language_mode", "Bilingue (arabe puis français)")),
        "status": "En attente",
        "final_data": {},
        "attestation_number": None,
        "issued_date": None,
        "issued_year": None,
    }
    client = _client()
    if client:
        row = {k: v for k, v in data.items() if k != "final_data"}
        row["final_data"] = None
        result = client.table("attestation_requests").insert(row).execute()
        if result.data:
            return _normalize_supabase_row(result.data[0])
        raise RuntimeError("La demande n'a pas pu être enregistrée dans Supabase.")

    conn = _connect()
    try:
        conn.execute("""INSERT INTO attestation_requests
            (id, created_at, updated_at, department, nom, prenom, date_naissance, lieu_naissance,
             email, telephone, purpose, language_mode, status, final_data, attestation_number, issued_date, issued_year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["id"], data["created_at"], data["updated_at"], data["department"], data["nom"], data["prenom"],
             data["date_naissance"], data["lieu_naissance"], data["email"], data["telephone"], data["purpose"],
             data["language_mode"], data["status"], "{}", None, None, None))
        row = conn.execute("SELECT * FROM attestation_requests WHERE id = ?", (request_id,)).fetchone()
        return _sqlite_row(row) or data
    finally:
        conn.close()


def _normalize_supabase_row(row: dict[str, Any]) -> dict[str, Any]:
    result = dict(row)
    fd = result.get("final_data")
    if isinstance(fd, str):
        try:
            fd = json.loads(fd)
        except json.JSONDecodeError:
            fd = {}
    result["final_data"] = fd if isinstance(fd, dict) else {}
    return result


def list_requests(department: str | None = None, status: str | None = None, limit: int = 500) -> list[dict[str, Any]]:
    client = _client()
    if client:
        query = client.table("attestation_requests").select("*").order("created_at", desc=True).limit(limit)
        if department:
            query = query.eq("department", department)
        if status:
            query = query.eq("status", status)
        rows = query.execute().data or []
        return [_normalize_supabase_row(r) for r in rows]

    conn = _connect()
    try:
        sql = "SELECT * FROM attestation_requests"
        where, params = [], []
        if department:
            where.append("department = ?")
            params.append(department)
        if status:
            where.append("status = ?")
            params.append(status)
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY created_at DESC LIMIT ?"
        params.append(int(limit))
        rows = conn.execute(sql, params).fetchall()
        return [_sqlite_row(r) for r in rows if r]
    finally:
        conn.close()


def get_request(request_id: str) -> dict[str, Any] | None:
    client = _client()
    if client:
        rows = client.table("attestation_requests").select("*").eq("id", request_id).limit(1).execute().data or []
        return _normalize_supabase_row(rows[0]) if rows else None
    conn = _connect()
    try:
        return _sqlite_row(conn.execute("SELECT * FROM attestation_requests WHERE id = ?", (request_id,)).fetchone())
    finally:
        conn.close()


def issue_request(request_id: str, final_data: dict[str, Any]) -> dict[str, Any]:
    """Finalise un dossier; attribue atomiquement un numéro unique pour l'année.

    Un dossier déjà émis conserve son numéro et sa date si le document est téléchargé à nouveau.
    """
    client = _client()
    if client:
        result = client.rpc("issue_attestation", {
            "p_request_id": request_id,
            "p_final_data": final_data,
        }).execute()
        returned = result.data
        if isinstance(returned, list):
            returned = returned[0] if returned else None
        if isinstance(returned, dict):
            return _normalize_supabase_row(returned)
        record = get_request(request_id)
        if not record:
            raise RuntimeError("Demande introuvable après génération.")
        return record

    conn = _connect()
    try:
        conn.execute("BEGIN IMMEDIATE")
        row = conn.execute("SELECT * FROM attestation_requests WHERE id = ?", (request_id,)).fetchone()
        record = _sqlite_row(row)
        if record is None:
            raise ValueError("Cette demande n'existe plus.")
        if record.get("status") == "Générée" and record.get("attestation_number") is not None:
            conn.execute("COMMIT")
            return record

        issued = now_local()
        year = issued.year
        seq = conn.execute("SELECT last_number FROM attestation_sequences WHERE issue_year = ?", (year,)).fetchone()
        next_number = (int(seq[0]) + 1) if seq else 1
        conn.execute("""INSERT INTO attestation_sequences(issue_year, last_number) VALUES (?, ?)
            ON CONFLICT(issue_year) DO UPDATE SET last_number = excluded.last_number""", (year, next_number))
        conn.execute("""UPDATE attestation_requests SET updated_at=?, status='Générée', final_data=?,
            attestation_number=?, issued_date=?, issued_year=? WHERE id=?""",
            (issued.isoformat(timespec="seconds"), json.dumps(final_data, ensure_ascii=False), next_number,
             issued.strftime("%d/%m/%Y"), year, request_id))
        conn.execute("COMMIT")
        return get_request(request_id) or {}
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except sqlite3.Error:
            pass
        raise
    finally:
        conn.close()
