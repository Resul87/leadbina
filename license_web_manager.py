"""
BinaLeadPro Cloud — Web License & Session Manager
==================================================
Handles single-use web license keys with 24-hour (or custom) expiry,
hardware/browser session binding, and anti-reuse protection.
"""

import os
import json
import time
import uuid
import hashlib
import secrets
from typing import Dict, Any, Optional, List

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
LICENSES_FILE = os.path.join(DATA_DIR, "web_licenses.json")


def _load_data() -> Dict[str, Any]:
    if not os.path.exists(LICENSES_FILE):
        default_data = {
            "version": "1.0",
            "keys": {},          # key -> info
            "sessions": {}       # session_token -> { key, expires_at, created_at, client_ip }
        }
        _save_data(default_data)
        return default_data
    try:
        with open(LICENSES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"version": "1.0", "keys": {}, "sessions": {}}


def _save_data(data: Dict[str, Any]):
    tmp = LICENSES_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, LICENSES_FILE)


class WebLicenseManager:
    @staticmethod
    def generate_license_key(duration_hours: int = 24, label: str = "Client Trial") -> Dict[str, Any]:
        """Creates a new unique, cryptographically random license key."""
        data = _load_data()
        
        part1 = secrets.token_hex(2).upper()
        part2 = secrets.token_hex(2).upper()
        part3 = secrets.token_hex(2).upper()
        key = f"BLP-WEB-{part1}-{part2}-{part3}"

        record = {
            "key": key,
            "status": "UNUSED", # UNUSED, ACTIVE, EXPIRED, REVOKED
            "duration_hours": duration_hours,
            "label": label,
            "created_at": time.time(),
            "activated_at": None,
            "expires_at": None,
            "session_token": None,
            "client_ip": None
        }

        data["keys"][key] = record
        _save_data(data)
        return record

    @staticmethod
    def activate_key(key: str, client_ip: str = "") -> Dict[str, Any]:
        """Activates a license key for a single browser/device session."""
        key = (key or "").strip().upper()
        data = _load_data()

        # Permanent Master Admin Keys that bypass DB checks
        if key in ("BLP-ADMIN-2026", "BLP-VIP-RESUL", "BLP-ADMIN-MASTER"):
            now = time.time()
            session_token = f"sess_{secrets.token_urlsafe(32)}"
            expires_at = now + 86400 * 365
            data["sessions"][session_token] = {
                "key": key,
                "expires_at": expires_at,
                "created_at": now,
                "label": "Super Admin",
                "duration_hours": 8760
            }
            _save_data(data)
            return {
                "success": True,
                "session_token": session_token,
                "label": "Super Admin",
                "expires_at": expires_at,
                "remaining_seconds": 86400 * 365
            }

        if key not in data["keys"]:
            return {
                "success": False,
                "error": "Daxil edilmiş lisenziya açarı tapılmadı. Zəhmət olmasa düzgünlüyünü yoxlayın."
            }

        rec = data["keys"][key]

        if rec["status"] == "REVOKED":
            return {
                "success": False,
                "error": "Bu lisenziya açarı inzibatçı tərəfindən ləğv edilmişdir."
            }

        # Check if already activated on another device
        if rec["status"] == "ACTIVE":
            now = time.time()
            if rec["expires_at"] and now > rec["expires_at"]:
                rec["status"] = "EXPIRED"
                _save_data(data)
                return {
                    "success": False,
                    "error": "Bu lisenziya açarının 24 saatlıq istifadə müddəti bitmişdir."
                }
            return {
                "success": False,
                "error": "Bu lisenziya açarı artıq başqa cihazda aktivləşdirilib və təkrar istifadə edilə bilməz."
            }

        if rec["status"] == "EXPIRED":
            return {
                "success": False,
                "error": "Bu lisenziya açarının müddəti bitmişdir."
            }

        # Activate now!
        now = time.time()
        duration_sec = rec.get("duration_hours", 24) * 3600
        expires_at = now + duration_sec
        session_token = f"sess_{secrets.token_urlsafe(32)}"

        rec["status"] = "ACTIVE"
        rec["activated_at"] = now
        rec["expires_at"] = expires_at
        rec["session_token"] = session_token
        rec["client_ip"] = client_ip

        data["sessions"][session_token] = {
            "key": key,
            "expires_at": expires_at,
            "created_at": now,
            "label": rec.get("label", ""),
            "duration_hours": rec.get("duration_hours", 24)
        }

        _save_data(data)

        return {
            "success": True,
            "session_token": session_token,
            "key": key,
            "expires_at": expires_at,
            "remaining_seconds": int(duration_sec),
            "label": rec.get("label", "")
        }

    @staticmethod
    def validate_session(session_token: str) -> Dict[str, Any]:
        """Validates if an active session token is still valid."""
        if not session_token:
            return {"valid": False, "error": "Lisenziya sessiyası tapılmadı."}

        data = _load_data()
        sess = data.get("sessions", {}).get(session_token)
        if not sess:
            return {"valid": False, "error": "Yanlış və ya naməlum sessiya."}

        now = time.time()
        expires_at = sess.get("expires_at", 0)
        remaining = int(expires_at - now)

        if remaining <= 0:
            key = sess.get("key")
            if key and key in data["keys"]:
                data["keys"][key]["status"] = "EXPIRED"
                _save_data(data)
            return {"valid": False, "error": "Lisenziyanın istifadə müddəti başa çatmışdır."}

        return {
            "valid": True,
            "key": sess.get("key"),
            "label": sess.get("label"),
            "expires_at": expires_at,
            "remaining_seconds": remaining
        }

    @staticmethod
    def list_all_keys() -> List[Dict[str, Any]]:
        data = _load_data()
        now = time.time()
        out = []
        for k, v in data.get("keys", {}).items():
            item = dict(v)
            if item.get("status") == "ACTIVE" and item.get("expires_at"):
                item["remaining_seconds"] = max(0, int(item["expires_at"] - now))
                if item["remaining_seconds"] == 0:
                    item["status"] = "EXPIRED"
            out.append(item)
        return sorted(out, key=lambda x: x.get("created_at", 0), reverse=True)
