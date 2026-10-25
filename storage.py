# -*- coding: utf-8 -*-
"""Stockage des demandes : Supabase pour le cloud, SQLite pour le mode local."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import base64
import hashlib
import json
import os
import re
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
            source_snapshot TEXT NOT NULL DEFAULT '{}',
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
        CREATE TABLE IF NOT EXISTS attestation_generation_files (
            id TEXT PRIMARY KEY,
            archive_key TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL,
            department TEXT NOT NULL DEFAULT '',
            person_name TEXT NOT NULL DEFAULT '',
            attestation_number TEXT NOT NULL DEFAULT '',
            issued_year INTEGER,
            language_mode TEXT NOT NULL DEFAULT '',
            file_kind TEXT NOT NULL DEFAULT 'DOCX',
            file_name TEXT NOT NULL,
            byte_size INTEGER NOT NULL DEFAULT 0,
            sha256 TEXT NOT NULL DEFAULT '',
            content BLOB NOT NULL
        );
        CREATE INDEX IF NOT EXISTS attestation_generation_files_dept_created_idx
            ON attestation_generation_files(department, created_at DESC);
    """)
    # Migration sans perte pour les bases locales créées avec l'ancienne version.
    columns = {row[1] for row in conn.execute("PRAGMA table_info(attestation_requests)").fetchall()}
    if "source_snapshot" not in columns:
        conn.execute("ALTER TABLE attestation_requests ADD COLUMN source_snapshot TEXT NOT NULL DEFAULT '{}' ")
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
    if data.get("source_snapshot"):
        try:
            data["source_snapshot"] = json.loads(data["source_snapshot"])
        except (TypeError, json.JSONDecodeError):
            data["source_snapshot"] = {}
    else:
        data["source_snapshot"] = {}
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
        "source_snapshot": payload.get("source_snapshot", {}) if isinstance(payload.get("source_snapshot", {}), dict) else {},
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
        row["source_snapshot"] = data["source_snapshot"]
        result = client.table("attestation_requests").insert(row).execute()
        if result.data:
            return _normalize_supabase_row(result.data[0])
        raise RuntimeError("La demande n'a pas pu être enregistrée dans Supabase.")

    conn = _connect()
    try:
        conn.execute("""INSERT INTO attestation_requests
            (id, created_at, updated_at, department, nom, prenom, date_naissance, lieu_naissance,
             email, telephone, purpose, source_snapshot, language_mode, status, final_data, attestation_number, issued_date, issued_year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["id"], data["created_at"], data["updated_at"], data["department"], data["nom"], data["prenom"],
             data["date_naissance"], data["lieu_naissance"], data["email"], data["telephone"], data["purpose"],
             json.dumps(data["source_snapshot"], ensure_ascii=False), data["language_mode"], data["status"], "{}", None, None, None))
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
    snapshot = result.get("source_snapshot")
    if isinstance(snapshot, str):
        try:
            snapshot = json.loads(snapshot)
        except json.JSONDecodeError:
            snapshot = {}
    result["source_snapshot"] = snapshot if isinstance(snapshot, dict) else {}
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


# -------------------- Archives des générations --------------------
_ARCHIVE_DIR = APP_DIR / "archives_generations"


def _safe_file_component(value: str, fallback: str = "Document") -> str:
    text = str(value or "").strip()
    text = re.sub(r"[\\/:*?\"<>|\x00-\x1f]+", "_", text)
    text = re.sub(r"\s+", "_", text).strip(" ._")
    return (text[:150] or fallback)


def _local_archive_path(archive_key: str, file_name: str, department: str = "", issued_year: int | None = None) -> Path:
    year = str(issued_year or now_local().year)
    dept = _safe_file_component(department or "Sans_departement", "Sans_departement")
    key_hash = hashlib.sha256(str(archive_key).encode("utf-8")).hexdigest()[:12]
    filename = _safe_file_component(file_name, "generation.bin")
    target_dir = _ARCHIVE_DIR / year / dept
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"{key_hash}_{filename}"


def save_generation_file(
    archive_key: str,
    file_name: str,
    content: bytes,
    *,
    department: str = "",
    person_name: str = "",
    attestation_number: str = "",
    issued_year: int | None = None,
    language_mode: str = "",
    file_kind: str = "DOCX",
) -> dict[str, Any]:
    """Archive un document en local et dans le backend actif.

    En mode Supabase, le contenu est conservé en base pour survivre aux redémarrages
    de Streamlit Cloud. Le dossier local reste pratique en installation sur serveur/PC.
    Un archive_key stable met à jour le même document au lieu de multiplier les copies.
    """
    if not archive_key or not file_name or not isinstance(content, (bytes, bytearray)):
        raise ValueError("Clé d'archive, nom du fichier et contenu binaire obligatoires.")
    payload_bytes = bytes(content)
    now = now_local().isoformat(timespec="seconds")
    issued_year = int(issued_year or now_local().year)
    file_id = str(uuid.uuid5(uuid.NAMESPACE_URL, "udl-attestation-archive:" + archive_key))
    digest = hashlib.sha256(payload_bytes).hexdigest()
    local_path: str | None = None
    try:
        path = _local_archive_path(archive_key, file_name, department, issued_year)
        path.write_bytes(payload_bytes)
        local_path = str(path)
    except OSError:
        # Le stockage base peut réussir même si le système de fichiers local est en lecture seule.
        pass

    try:
        client = _client()
    except Exception as exc:
        return {
            "id": file_id, "archive_key": str(archive_key), "created_at": now,
            "department": str(department or ""), "person_name": str(person_name or ""),
            "attestation_number": str(attestation_number or ""), "issued_year": issued_year,
            "language_mode": str(language_mode or ""), "file_kind": str(file_kind or "DOCX"),
            "file_name": _safe_file_component(file_name, "generation.bin"),
            "byte_size": len(payload_bytes), "sha256": digest, "local_path": local_path,
            "persistent": False, "error": str(exc),
        }
    if client:
        row = {
            "id": file_id,
            "archive_key": str(archive_key),
            "created_at": now,
            "department": str(department or ""),
            "person_name": str(person_name or ""),
            "attestation_number": str(attestation_number or ""),
            "issued_year": issued_year,
            "language_mode": str(language_mode or ""),
            "file_kind": str(file_kind or "DOCX"),
            "file_name": _safe_file_component(file_name, "generation.bin"),
            "byte_size": len(payload_bytes),
            "sha256": digest,
            "content_base64": base64.b64encode(payload_bytes).decode("ascii"),
        }
        try:
            client.table("attestation_generation_files").upsert(row, on_conflict="archive_key").execute()
            return {k: v for k, v in row.items() if k != "content_base64"} | {"local_path": local_path, "persistent": True}
        except Exception as exc:
            # Le document existe au minimum dans le dossier local de l'instance.
            # L'administrateur sera invité à appliquer la migration SQL pour la persistance cloud.
            return {k: v for k, v in row.items() if k != "content_base64"} | {
                "local_path": local_path, "persistent": False, "error": str(exc)
            }

    conn = _connect()
    try:
        conn.execute("""INSERT INTO attestation_generation_files
            (id, archive_key, created_at, department, person_name, attestation_number, issued_year,
             language_mode, file_kind, file_name, byte_size, sha256, content)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(archive_key) DO UPDATE SET
             created_at=excluded.created_at, department=excluded.department,
             person_name=excluded.person_name, attestation_number=excluded.attestation_number,
             issued_year=excluded.issued_year, language_mode=excluded.language_mode,
             file_kind=excluded.file_kind, file_name=excluded.file_name,
             byte_size=excluded.byte_size, sha256=excluded.sha256, content=excluded.content""",
            (file_id, str(archive_key), now, str(department or ""), str(person_name or ""),
             str(attestation_number or ""), issued_year, str(language_mode or ""),
             str(file_kind or "DOCX"), _safe_file_component(file_name, "generation.bin"),
             len(payload_bytes), digest, sqlite3.Binary(payload_bytes)))
    finally:
        conn.close()
    return {
        "id": file_id, "archive_key": str(archive_key), "created_at": now,
        "department": str(department or ""), "person_name": str(person_name or ""),
        "attestation_number": str(attestation_number or ""), "issued_year": issued_year,
        "language_mode": str(language_mode or ""), "file_kind": str(file_kind or "DOCX"),
        "file_name": _safe_file_component(file_name, "generation.bin"),
        "byte_size": len(payload_bytes), "sha256": digest, "local_path": local_path,
        "persistent": True,
    }


def list_generation_files(department: str | None = None, limit: int = 1000) -> list[dict[str, Any]]:
    """Renvoie les métadonnées des générations archivées, sans charger les contenus."""
    client = _client()
    if client:
        query = (client.table("attestation_generation_files")
                 .select("id,archive_key,created_at,department,person_name,attestation_number,issued_year,language_mode,file_kind,file_name,byte_size,sha256")
                 .order("created_at", desc=True).limit(int(limit)))
        if department:
            query = query.eq("department", department)
        return query.execute().data or []

    conn = _connect()
    try:
        sql = "SELECT id, archive_key, created_at, department, person_name, attestation_number, issued_year, language_mode, file_kind, file_name, byte_size, sha256 FROM attestation_generation_files"
        params: list[Any] = []
        if department:
            sql += " WHERE department = ?"
            params.append(department)
        sql += " ORDER BY created_at DESC LIMIT ?"
        params.append(int(limit))
        return [dict(row) for row in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()


def get_generation_file(file_id: str) -> dict[str, Any] | None:
    """Récupère les métadonnées et le contenu binaire d'une génération archivée."""
    client = _client()
    if client:
        rows = (client.table("attestation_generation_files").select("*")
                .eq("id", str(file_id)).limit(1).execute().data or [])
        if not rows:
            return None
        row = dict(rows[0])
        encoded = row.pop("content_base64", "")
        try:
            content = base64.b64decode(encoded) if encoded else b""
        except (ValueError, TypeError):
            content = b""
        row["content"] = content
        return row

    conn = _connect()
    try:
        row = conn.execute("SELECT * FROM attestation_generation_files WHERE id = ?", (str(file_id),)).fetchone()
        if not row:
            return None
        data = dict(row)
        data["content"] = bytes(data.get("content") or b"")
        return data
    finally:
        conn.close()


def delete_all_request_history() -> int:
    """Supprime toutes les demandes, mais conserve les compteurs et les archives générées."""
    client = _client()
    if client:
        try:
            result = client.rpc("clear_attestation_history").execute()
            value = result.data
            if isinstance(value, list) and value:
                value = value[0]
            if isinstance(value, dict):
                value = value.get("clear_attestation_history", value.get("deleted_count", 0))
            return int(value or 0)
        except Exception:
            # Compatibilité avec un schéma Supabase antérieur à l'ajout de la fonction RPC.
            count_result = client.table("attestation_requests").select("id", count="exact").execute()
            count = int(getattr(count_result, "count", 0) or len(count_result.data or []))
            client.table("attestation_requests").delete().neq("id", "00000000-0000-0000-0000-000000000000").execute()
            return count

    conn = _connect()
    try:
        conn.execute("BEGIN IMMEDIATE")
        count = int(conn.execute("SELECT COUNT(*) FROM attestation_requests").fetchone()[0])
        conn.execute("DELETE FROM attestation_requests")
        # Ne pas supprimer attestation_sequences : les numéros officiels ne doivent jamais être réutilisés.
        conn.execute("COMMIT")
        return count
    except Exception:
        conn.execute("ROLLBACK")
        raise
    finally:
        conn.close()
