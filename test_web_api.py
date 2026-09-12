"""
Live HTTP Test for BinaLeadPro Cloud using requests
"""

import sys
import os
import json
import time
import threading
import requests
import uvicorn

# Console UTF-8
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from license_web_manager import WebLicenseManager

BASE_URL = "http://127.0.0.1:8765"


def run_server():
    from web_server import app
    uvicorn.run(app, host="127.0.0.1", port=8765, log_level="warning")


# Start server in background thread
server_thread = threading.Thread(target=run_server, daemon=True)
server_thread.start()
time.sleep(1.5)

print("═"*65)
print("🧪 BİNALAUNCH PRO CLOUD — CANLI HTTP TESTLƏRİ BAŞLAYIR")
print("═"*65)

# 1. Test Index Page
resp_index = requests.get(f"{BASE_URL}/")
assert resp_index.status_code == 200
assert "BinaLeadPro" in resp_index.text
print("✅ [1/6] GET / (SaaS Frontend HTML) Uğurla yükləndi.")

# 2. Test Locations API
resp_locs = requests.get(f"{BASE_URL}/api/locations")
assert resp_locs.status_code == 200
locs = resp_locs.json()
assert "Xətai r." in locs
print(f"✅ [2/6] GET /api/locations Uğurlu ({len(locs)} rayon strukturu yükləndi).")

# 3. Test Invalid License Key
resp_bad_key = requests.post(f"{BASE_URL}/api/license/activate", json={"key": "INVALID-KEY-123"})
assert resp_bad_key.status_code == 400
assert "tapılmadı" in resp_bad_key.json()["error"]
print("✅ [3/6] Yanlış lisenziya açarı bloklandı (Düzgün error qaytarıldı).")

# 4. Test Valid License Key Generation & Activation
rec = WebLicenseManager.generate_license_key(duration_hours=24, label="Test Müştəri")
valid_key = rec["key"]
print(f"  🔑 Sınaq Açarı Yaradıldı: {valid_key}")

resp_act = requests.post(f"{BASE_URL}/api/license/activate", json={"key": valid_key})
assert resp_act.status_code == 200
act_data = resp_act.json()
assert act_data["success"] is True
session_token = act_data["session_token"]
print(f"✅ [4/6] Lisenziya aktivləşdirildi. Qalan vaxt: {round(act_data['remaining_seconds']/3600, 1)} saat")

# 5. Test Anti-Reuse Protection (Activating same key again should fail!)
resp_reuse = requests.post(f"{BASE_URL}/api/license/activate", json={"key": valid_key})
assert resp_reuse.status_code == 400
assert "artıq başqa cihazda aktivləşdirilib" in resp_reuse.json()["error"]
print("✅ [5/6] Təkrar istifadə cəhdi bloklandı (Anti-reuse müdafiəsi tam işləkdir).")

# 6. Test Excel Export API with active session
sample_leads = [
    {
        "№": 1,
        "Mənbə": "bina.az",
        "Ad / Mülkiyyətçi": "Əli Həsənov",
        "Əlaqə Nömrəsi": "(055) 123-45-67",
        "Status": "Mülkiyyətçi",
        "Qiymət (AZN)": 185000,
        "Otaq": 2,
        "Sahə (m²)": 68.5,
        "Yerləşmə": "Xətai r., Həzi Aslanov m.",
        "Kupça": "Var",
        "İpoteka": "Var",
        "Təmir": "Təmirli",
        "Elan ID": "6543210",
        "Elan Keçidi": "https://bina.az/items/6543210"
    }
]

resp_excel = requests.post(
    f"{BASE_URL}/api/export/excel",
    json={"leads": sample_leads},
    headers={"X-Session-Token": session_token}
)
assert resp_excel.status_code == 200
assert len(resp_excel.content) > 1000
print(f"✅ [6/6] Excel ixrac API-si uğurlu ({len(resp_excel.content)} bayt .xlsx generatsiya olundu).")

print("═"*65)
print("🎉 BÜTÜN WEB TESTLƏR 100% UĞURLA TAMAMLANDI!")
print("═"*65 + "\n")
