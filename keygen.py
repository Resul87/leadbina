"""
BinaLeadPro — Admin Key Generator Tool (HWID-Bound)
===================================================
Generates cryptographically signed, one-time activation keys
tied to the client's unique Hardware ID (Cihaz Kodu).

Usage Examples:
    python keygen.py --hwid BLP-1ED4-746C --hours 24
    python keygen.py --hwid BLP-1ED4-746C --hours 48
    python keygen.py --hwid BLP-1ED4-746C --days 30
    python keygen.py --my-hwid
"""

import sys
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass
import argparse
import hashlib
import hmac
import license_manager

SECRET_SALT = "BinaLeadPro_v3_Super_Secret_Salt_2026_AGY"

def generate_key(hwid: str, duration_type: str, value: int) -> str:
    clean_hwid = hwid.replace("-", "").replace(" ", "").upper()
    if clean_hwid.startswith("BLP"):
        clean_hwid = clean_hwid[3:]

    payload = f"{clean_hwid}:{duration_type}:{value}"
    sig = hmac.new(SECRET_SALT.encode(), payload.encode(), hashlib.sha256).hexdigest()[:6].upper()

    if duration_type == "hours":
        return f"TRIAL-{value}H-{clean_hwid}-{sig}"
    else:
        return f"SUB-{value}D-{clean_hwid}-{sig}"

def main():
    parser = argparse.ArgumentParser(description="BinaLeadPro License ID-Bound Key Generator")
    parser.add_argument("--id", "--hwid", dest="hwid", type=str, help="Client's License ID (e.g. BLP-1ED4-746C or 1ED4746C)")
    parser.add_argument("--hours", type=int, help="Trial hours (e.g. 24 or 48)")
    parser.add_argument("--days", type=int, help="Subscription days (e.g. 30)")
    parser.add_argument("--client", type=str, default="", help="Client name (optional note)")
    parser.add_argument("--my-id", "--my-hwid", dest="my_hwid", action="store_true", help="Print this PC's License ID")
    args = parser.parse_args()

    if args.my_hwid:
        print(f"\n🔑 Bu Kompüterin Lisenziya ID-si: {license_manager.get_hwid()}\n")
        return

    if not args.hwid:
        print("\n❌ XƏTA: Müştərinin Lisenziya ID-sini (--id) daxil edin!")
        print("\nİstifadə qaydası:")
        print("  1 Günlük Test Açarı:")
        print("    python keygen.py --id BLP-1ED4-746C --hours 24")
        print("\n  2 Günlük Test Açarı:")
        print("    python keygen.py --id BLP-1ED4-746C --hours 48")
        print("\n  1 Aylıq Abunəlik Açarı:")
        print("    python keygen.py --id BLP-1ED4-746C --days 30\n")
        return

    if args.hours:
        key = generate_key(args.hwid, "hours", args.hours)
        print("\n=======================================================")
        print(f"🔑 MÜŞTƏRİYƏ GÖNDƏRİLƏCƏK {args.hours} SAATLIQ AKTİVASİYA AÇARI:")
        print(f"   {key}")
        print("=======================================================")
        print(f"• Lisenziya ID: {args.hwid}")
        print(f"• Müddət: {args.hours} saat (aktivləşəndən sonra)")
        print("• Qoruma: Yalnız bu Lisenziya ID üçün işləyəcək, başqasına ötürülə bilməz.")
        print("• Təkrar İstifadə: 24 saat bitdikdən sonra bu açar kilitlənir.\n")
    elif args.days:
        key = generate_key(args.hwid, "days", args.days)
        print("\n=======================================================")
        print(f"🔑 MÜŞTƏRİYƏ GÖNDƏRİLƏCƏK {args.days} GÜNLÜK ABUNƏLİK AÇARI:")
        print(f"   {key}")
        print("=======================================================")
        print(f"• Lisenziya ID: {args.hwid}")
        print(f"• Müddət: {args.days} gün")
        print("• Qoruma: Yalnız bu Lisenziya ID üçün işləyəcək, başqasına ötürülə bilməz.\n")
    else:
        print("\n❌ Zəhmət olmasa --hours (məs: 24) və ya --days (məs: 30) qeyd edin.")

if __name__ == "__main__":
    main()
