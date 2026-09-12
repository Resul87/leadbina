"""
BinaLeadPro — Hardware-Bound Anti-Reuse License & Trial Manager
===============================================================
Features:
- Unique Hardware-ID (HWID) binding per Windows machine.
- Cryptographically signed 24h, 48h trials, and 30d subscriptions.
- Strict anti-reuse protection: used keys cannot be re-activated.
- Keys cannot be transferred to or activated on another PC.
- Anti-tampering system clock rollback detection.
- Master admin override keys for owner.
"""

import os
import json
import time
import hashlib
import hmac
import re
from typing import Dict, Any, Tuple

LICENSE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "license.json")
SECRET_SALT = "BinaLeadPro_v3_Super_Secret_Salt_2026_AGY"

# Admin Master Keys (Work on any PC for owner)
MASTER_KEYS = {
    "ADMIN-UNLIMITED": 999999,
    "BINA-VIP-2026": 8760,
    "BINA-MONTH-30": 720
}


def get_hwid() -> str:
    """Returns unique, permanent 8-character hardware ID for this PC"""
    raw = ""
    if os.name == "nt":
        try:
            import winreg
            k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography")
            raw, _ = winreg.QueryValueEx(k, "MachineGuid")
        except Exception:
            pass
    if not raw:
        raw = os.environ.get("COMPUTERNAME", "PC") + "_" + os.environ.get("USERNAME", "USER")
    h = hashlib.sha256(raw.encode()).hexdigest()[:8].upper()
    return f"BLP-{h[:4]}-{h[4:]}"


def _get_clean_hwid() -> str:
    full = get_hwid()
    clean = full.replace("-", "").replace(" ", "").upper()
    if clean.startswith("BLP"):
        clean = clean[3:]
    return clean


def init_license(default_hours: int = 0) -> Dict[str, Any]:
    """Initializes a blank license file. Without key, software is locked until activated."""
    now = int(time.time())
    data = {
        "hwid": get_hwid(),
        "first_launch": now,
        "is_activated": False,
        "active_key": None,
        "expiry_time": 0, # Expired by default until client enters key
        "last_checked": now,
        "used_keys": []
    }
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return data


def check_status() -> Tuple[bool, str, int]:
    """
    Returns (is_active, status_display_text, remaining_seconds)
    """
    if not os.path.exists(LICENSE_FILE):
        data = init_license()
    else:
        try:
            with open(LICENSE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = init_license()

    now = int(time.time())
    expiry = data.get("expiry_time", 0)
    last_checked = data.get("last_checked", now)
    is_activated = data.get("is_activated", False)
    active_key = data.get("active_key")
    used_keys = data.get("used_keys", [])

    # Anti-tampering check: system clock rolled back by > 5 minutes
    if now < last_checked - 300:
        return False, "⚠️ Sistem saatı dəyişdirilib! Proqram kilidləndi.", 0

    data["last_checked"] = now
    remaining = expiry - now

    if remaining > 0 and expiry > 0:
        hours = remaining // 3600
        mins = (remaining % 3600) // 60
        if is_activated:
            text = f"🟢 Abunəlik Aktivdir: {hours} saat {mins} dəqiqə qaldı"
        else:
            text = f"🟢 Sınaq Rejimi: {hours} saat {mins} dəqiqə qaldı"
        
        try:
            with open(LICENSE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass
            
        return True, text, remaining
    else:
        # Expired! Invalidate active key into used_keys
        if active_key and active_key not in used_keys:
            used_keys.append(active_key)
            data["used_keys"] = used_keys
            data["active_key"] = None
            data["is_activated"] = False
            try:
                with open(LICENSE_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2)
            except Exception:
                pass

        text = "🔴 Aktiv Lisenziya Yoxdur (Kod Daxil Edin)"
        return False, text, 0


def activate_license(key: str) -> Tuple[bool, str]:
    key = key.strip().upper()
    if not key:
        return False, "Zəhmət olmasa lisenziya açarını daxil edin."

    now = int(time.time())
    if not os.path.exists(LICENSE_FILE):
        data = init_license()
    else:
        try:
            with open(LICENSE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = init_license()

    used_keys = data.get("used_keys", [])
    if key in used_keys:
        return False, "❌ Bu lisenziya açarından artıq istifadə edilib və müddəti bitib!\nBir açar yalnız 1 dəfə istifadə edilə bilər."

    # 1. Master Keys
    if key in MASTER_KEYS:
        hours = MASTER_KEYS[key]
        data["is_activated"] = True
        data["active_key"] = key
        data["expiry_time"] = now + (hours * 3600)
        data["last_checked"] = now
        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return True, f"Admin Master Açarı qəbul edildi ({hours} saat)!"

    # 2. Hardware-Bound Keys: TRIAL-{val}H-{key_hwid}-{sig} or SUB-{val}D-{key_hwid}-{sig}
    parts = key.split("-")
    if len(parts) == 4 and parts[0] in ["TRIAL", "SUB"]:
        prefix, dur_str, key_hwid, sig = parts
        my_clean_hwid = _get_clean_hwid()

        if key_hwid != my_clean_hwid:
            return False, (
                f"❌ Bu aktivasiya açarı başqa lisenziya üçün yaradılıb!\n\n"
                f"Açardakı ID: {key_hwid}\n"
                f"Sizin Lisenziya ID: {my_clean_hwid}\n\n"
                f"Zəhmət olmasa öz Lisenziya ID-nizə uyğun göndərilən açarı daxil edin."
            )

        is_sub = (prefix == "SUB")
        dur_type = "days" if is_sub else "hours"
        try:
            val = int(dur_str[:-1])
        except ValueError:
            return False, "❌ Açarın müddət formatı oxunmadı!"

        expected_payload = f"{key_hwid}:{dur_type}:{val}"
        expected_sig = hmac.new(SECRET_SALT.encode(), expected_payload.encode(), hashlib.sha256).hexdigest()[:6].upper()

        if sig != expected_sig:
            return False, "❌ Yanlış və ya saxta lisenziya açarı!"

        hours = val * 24 if is_sub else val
        data["is_activated"] = is_sub
        data["active_key"] = key
        data["expiry_time"] = now + (hours * 3600)
        data["last_checked"] = now

        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

        if is_sub:
            return True, f"✅ Təbriklər! Proqram {val} günlük abunəlik üçün aktivləşdirildi."
        else:
            return True, f"✅ Təbriklər! Proqram {val} saatlıq sınaq üçün aktivləşdirildi."

    return False, "❌ Yanlış lisenziya açarı! Zəhmət olmasa sizə göndərilən kodu dəqiq daxil edin."
