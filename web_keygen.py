"""
BinaLeadPro Web Key Generator (CLI)
===================================
Generates 24-hour, 48-hour, or custom single-use license keys for the Web App.
Usage:
    python web_keygen.py --duration 24 --label "Müştəri Rauf"
    python web_keygen.py --list
"""

import sys
import argparse
from license_web_manager import WebLicenseManager

# Force UTF-8 on Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(description="BinaLeadPro Web License Generator")
    parser.add_argument("--duration", type=int, default=24, help="License duration in hours (default: 24)")
    parser.add_argument("--label", type=str, default="Müştəri Sınaq", help="Client or trial note")
    parser.add_argument("--list", action="store_true", help="List all generated web license keys")
    args = parser.parse_args()

    if args.list:
        keys = WebLicenseManager.list_all_keys()
        print("\n" + "="*70)
        print("📋 BİNALAUNCH WEB LİSENZİYA SİYAHISI")
        print("="*70)
        if not keys:
            print("Hələ heç bir lisenziya açarı yaradılmayıb.")
        for k in keys:
            status = k.get("status")
            status_icon = "🟢" if status == "ACTIVE" else ("🟡" if status == "UNUSED" else "🔴")
            rem_hrs = round(k.get("remaining_seconds", 0) / 3600, 1) if k.get("remaining_seconds") else 0
            print(f"{status_icon} {k['key']} | Status: {status:8} | Müddət: {k['duration_hours']}s | Qalıq: {rem_hrs}s | Qeyd: {k.get('label')}")
        print("="*70 + "\n")
        return

    rec = WebLicenseManager.generate_license_key(duration_hours=args.duration, label=args.label)
    key = rec["key"]

    print("\n" + "═"*70)
    print("✨ BİNALAUNCH PRO CLOUD — YENİ WEB LİSENZİYA AÇARI YARADILDI")
    print("═"*70)
    print(f"🔑 Lisenziya Kodu :  {key}")
    print(f"⏱️  Müddət        :  {args.duration} saat ({round(args.duration/24, 1)} gün)")
    print(f"📝 Təyinat        :  {args.label}")
    print(f"🛡️  Təhlükəsizlik :  1 Brauzer / Cihaz (İlk aktivləşmədən sonra başqası istifadə edə bilməz)")
    print("═"*70)
    print("\n👉 Müştəriyə göndərəcəyiniz mesaj mətni:")
    print("----------------------------------------------------------------------")
    print(f"Salam! BinaLeadPro Cloud platformasına giriş üçün 24 saatlıq sınaq kodunuz:\n\nLisenziya Kodu: {key}\nSayt Linki: (saytınızın ünvanı)\n\nKodu daxil edərək dərhal istifadəyə başlaya bilərsiniz.")
    print("----------------------------------------------------------------------\n")


if __name__ == "__main__":
    main()
