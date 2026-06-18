"""
API Key management and protected REST endpoints.

Mount this router in main.py with:
    from api import api_router
    app.include_router(api_router)
"""

import os, hashlib, secrets, json
from datetime import datetime
from fastapi import APIRouter, Depends, HTTPException, Header, Query, Request
from typing import Optional

# ── Import DB helpers from main ────────────────────────────────────────────────
from main import get_db, db_push, normalize_status, FIELDS

api_router = APIRouter()

# ── DB table ───────────────────────────────────────────────────────────────────

def init_api_keys_table():
    """Create api_keys table if it doesn't exist."""
    with get_db() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS api_keys (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            key_hash TEXT UNIQUE NOT NULL,
            key_prefix TEXT NOT NULL,
            label TEXT NOT NULL DEFAULT 'Untitled',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_used_at TIMESTAMP,
            is_active INTEGER DEFAULT 1
        )""")
    db_push()

init_api_keys_table()

# ── Auth dependency ────────────────────────────────────────────────────────────

def _hash_key(raw_key: str) -> str:
    return hashlib.sha256(raw_key.encode()).hexdigest()

async def verify_api_key(
    request: Request,
    x_api_key: Optional[str] = Header(None),
):
    """Validate API key from header or query param."""
    raw_key = x_api_key or request.query_params.get("api_key")
    if not raw_key:
        raise HTTPException(
            status_code=401,
            detail="Missing API key. Provide via X-API-Key header or ?api_key= query param."
        )
    key_hash = _hash_key(raw_key)
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, is_active FROM api_keys WHERE key_hash = ?", (key_hash,)
        ).fetchone()
    if not row:
        raise HTTPException(status_code=401, detail="Invalid API key.")
    if not row["is_active"]:
        raise HTTPException(status_code=403, detail="API key has been revoked.")
    # Update last_used_at
    with get_db() as conn:
        conn.execute(
            "UPDATE api_keys SET last_used_at = ? WHERE id = ?",
            (datetime.now().isoformat(), row["id"])
        )
    return raw_key

# ── Sanitize status for API output ─────────────────────────────────────────────

_BLANK_STATUSES = {"due", "paid", None}

def _sanitize_row(row: dict) -> dict:
    """Return a clean dict with status='due'/'paid' blanked out."""
    out = {}
    for f in FIELDS:
        val = row.get(f) or ""
        if f == "status":
            normalized = normalize_status(val)
            if normalized in _BLANK_STATUSES or val.strip() == "":
                val = ""
        out[f] = val
    return out

# ── API Key CRUD (admin — no auth) ────────────────────────────────────────────

@api_router.post("/api-keys", tags=["API Keys"])
def create_api_key(label: str = Query("Untitled", max_length=100)):
    """Generate a new API key. The full key is returned ONLY ONCE."""
    raw_key = f"ipm_{secrets.token_hex(24)}"
    key_hash = _hash_key(raw_key)
    key_prefix = raw_key[:12] + "…"
    with get_db() as conn:
        conn.execute(
            "INSERT INTO api_keys (key_hash, key_prefix, label, created_at) VALUES (?,?,?,?)",
            (key_hash, key_prefix, label.strip() or "Untitled", datetime.now().isoformat())
        )
    db_push()
    return {
        "api_key": raw_key,
        "prefix": key_prefix,
        "label": label.strip() or "Untitled",
        "message": "Save this key now — it cannot be shown again."
    }

@api_router.get("/api-keys", tags=["API Keys"])
def list_api_keys():
    """List all API keys (prefix only, not the full key)."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, key_prefix, label, created_at, last_used_at, is_active "
            "FROM api_keys ORDER BY created_at DESC"
        ).fetchall()
    return [dict(r) for r in rows]

@api_router.delete("/api-keys/{key_id}", tags=["API Keys"])
def revoke_api_key(key_id: int):
    """Revoke (deactivate) an API key."""
    with get_db() as conn:
        row = conn.execute("SELECT id FROM api_keys WHERE id = ?", (key_id,)).fetchone()
        if not row:
            raise HTTPException(404, "API key not found.")
        conn.execute("UPDATE api_keys SET is_active = 0 WHERE id = ?", (key_id,))
    db_push()
    return {"message": "API key revoked.", "id": key_id}

# ── Protected API endpoints ────────────────────────────────────────────────────

@api_router.get("/api/v1/policies", tags=["Policies API"], dependencies=[Depends(verify_api_key)])
def api_list_policies(
    limit: int = Query(50, ge=1, le=1000),
    offset: int = Query(0, ge=0),
):
    """Paginated list of all policies."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM policies ORDER BY id DESC LIMIT ? OFFSET ?", (limit, offset)
        ).fetchall()
        total = conn.execute("SELECT COUNT(*) AS cnt FROM policies").fetchone()["cnt"]
    return {
        "total": total,
        "limit": limit,
        "offset": offset,
        "data": [_sanitize_row(dict(r)) for r in rows],
    }

@api_router.get("/api/v1/policies/count", tags=["Policies API"], dependencies=[Depends(verify_api_key)])
def api_policy_count():
    """Get total policy count."""
    with get_db() as conn:
        count = conn.execute("SELECT COUNT(*) AS cnt FROM policies").fetchone()["cnt"]
        last = conn.execute("SELECT MAX(updated_at) AS lu FROM policies").fetchone()["lu"]
    return {"count": count, "last_updated": last}

@api_router.get("/api/v1/policies/search", tags=["Policies API"], dependencies=[Depends(verify_api_key)])
def api_search_policies(
    q: str = Query(..., min_length=1),
    field: str = Query("policyno"),
    limit: int = Query(100, ge=1, le=1000),
):
    """Search policies by policy number, name, or both."""
    query_val = q.strip().upper()
    with get_db() as conn:
        if field == "name":
            rows = conn.execute(
                "SELECT * FROM policies WHERE UPPER(COALESCE(name,'')) LIKE ? ORDER BY name LIMIT ?",
                (f"%{query_val}%", limit)
            ).fetchall()
        elif field == "both":
            rows = conn.execute(
                "SELECT * FROM policies WHERE UPPER(TRIM(policyno)) LIKE ? OR UPPER(COALESCE(name,'')) LIKE ? ORDER BY policyno LIMIT ?",
                (f"%{query_val}%", f"%{query_val}%", limit)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM policies WHERE UPPER(TRIM(policyno)) LIKE ? ORDER BY policyno LIMIT ?",
                (f"%{query_val}%", limit)
            ).fetchall()
    return {
        "query": q,
        "field": field,
        "count": len(rows),
        "data": [_sanitize_row(dict(r)) for r in rows],
    }

@api_router.get("/api/v1/policies/{policyno}", tags=["Policies API"], dependencies=[Depends(verify_api_key)])
def api_get_policy(policyno: str):
    """Get a single policy by exact policy number."""
    pno = policyno.strip().upper()
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM policies WHERE UPPER(TRIM(policyno)) = ?", (pno,)
        ).fetchone()
    if not row:
        raise HTTPException(404, f"Policy '{policyno}' not found.")
    return _sanitize_row(dict(row))
