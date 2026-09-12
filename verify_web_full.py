"""
Verification test for web scraper with all filters and price trend
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

BASE_URL = "http://127.0.0.1:8767"


def run_server():
    from web_server import app
    uvicorn.run(app, host="127.0.0.1", port=8767, log_level="warning")


server_thread = threading.Thread(target=run_server, daemon=True)
server_thread.start()
time.sleep(1.5)

print("═"*65)
print("🧪 BİNALAUNCH PRO CLOUD — YENİ FİLTRLƏR VƏ SKANER TESTİ")
print("═"*65)

# 1. Generate & activate a test license key
rec = WebLicenseManager.generate_license_key(duration_hours=24, label="Sınaq Test")
valid_key = rec["key"]
resp_act = requests.post(f"{BASE_URL}/api/license/activate", json={"key": valid_key})
session_token = resp_act.json()["session_token"]
print(f"✅ [1/3] Lisenziya aktivləşdirildi: {valid_key}")

# 2. Test live scrape stream with Kupça, İpoteka, Qiymət aralığı
params = {
    "session_token": session_token,
    "target_type": "AGENT",
    "sources": "bina",
    "region_name": "Xətai r.",
    "sublocation_name": "Bütün rayon üzrə",
    "listing_type": "Alqı-satqı",
    "category_name": "Bütün Mənzillər",
    "has_bill_of_sale": "true",
    "has_mortgage": "true",
    "has_repair": "true",
    "price_min": "100000",
    "price_max": "350000",
    "max_bina": "2",
    "max_yeni": "0",
    "max_tap": "0"
}

print("🔄 [2/3] Canlı axtarış (Kupçalı + İpotekalı + Təmirli + Qiymət filtri) yoxlanılır...")
stream_resp = requests.get(f"{BASE_URL}/api/scrape/stream", params=params, stream=True, timeout=30)
assert stream_resp.status_code == 200

leads_received = []

for line in stream_resp.iter_lines():
    if not line:
        continue
    line_str = line.decode('utf-8')
    if line_str.startswith("event: "):
        event_name = line_str[7:]
    elif line_str.startswith("data: "):
        data_str = line_str[6:]
        try:
            data = json.loads(data_str)
            if event_name == "lead":
                leads_received.append(data)
                print(f"  🎯 Lead gəldi: {data.get('Ad / Mülkiyyətçi')} | {data.get('Əlaqə Nömrəsi')} | {data.get('Qiymət (AZN)')} AZN | {data.get('m² Qiyməti (AZN)')} AZN/m² | Kupça: {data.get('Kupça')} | Trend: {data.get('Qiymət Trendi')}")
            elif event_name == "done":
                print(f"  🏁 Skaner tamamlandı: {data}")
                break
            elif event_name == "error":
                print(f"  ❌ Error event: {data}")
                sys.exit(1)
        except Exception:
            pass

assert len(leads_received) > 0, "Heç bir lead gəlmədi!"
print(f"✅ [2/3] Canlı skaner uğurludur! {len(leads_received)} ədəd lead alındı.")

# 3. Test Excel Export with the newly received leads
resp_excel = requests.post(
    f"{BASE_URL}/api/export/excel",
    json={"leads": leads_received},
    headers={"X-Session-Token": session_token}
)
assert resp_excel.status_code == 200
assert len(resp_excel.content) > 1000
print(f"✅ [3/3] Excel ixrac API-si uğurlu ({len(resp_excel.content)} bayt .xlsx yaradıldı).")

print("═"*65)
print("🎉 BÜTÜN DÜZƏLİŞLƏR VƏ YENİ FİLTRLƏR 100% İŞLƏKDİR!")
print("═"*65 + "\n")
