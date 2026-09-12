# BinaLeadPro Cloud — Veb SaaS Təlimatı və İstifadə Qaydaları

BinaLeadPro Cloud daşınmaz əmlak elanlarını (Bina.az, YeniEmlak.az, Tap.az) real-vaxt rejimində skan edən, mülkiyyətçi və makler əlaqələrini çıxaran və müştərilərə **heç bir fayl yüklətdirmədən** birbaşa brauzerdən təqdim edə biləcəyiniz müasir SaaS veb platformasıdır.

---

## 🚀 1. Saytı Öz Kompüterinizdə İşə Salmaq (1 Kliklə)

Qovluqdakı **`RUN_WEB.bat`** faylına iki dəfə klikləyin:
- Server avtomatik olaraq `http://127.0.0.1:8000` ünvanında işə düşəcək.
- Brauzeriniz avtomatik açılacaq və sistemin interfeysi görünəcək.

*(Əl ilə terminalda başlatmaq istəsəniz: `python web_server.py`)*

---

## 🔑 2. Müştəriyə 24 Saatlıq Lisenziya Kodu Yaratmaq

Müştərinizə göndərmək üçün istədiyiniz müddətə (məsələn, 24 saatlıq) açar yaratmaq üçün terminalda yazın:

```bash
python web_keygen.py --duration 24 --label "Müştəri Rauf"
```

**Nəticə nümunəsi:**
```
══════════════════════════════════════════════════════════════════════
✨ BİNALAUNCH PRO CLOUD — YENİ WEB LİSENZİYA AÇARI YARADILDI
══════════════════════════════════════════════════════════════════════
🔑 Lisenziya Kodu :  BLP-WEB-9DF7-ADC5-F446
⏱️  Müddət        :  24 saat (1.0 gün)
📝 Təyinat        :  Müştəri Rauf
🛡️  Təhlükəsizlik :  1 Cihaz (İlk aktivləşmədən sonra başqası istifadə edə bilməz)
══════════════════════════════════════════════════════════════════════
```

Bütün mövcud və istifadə edilmiş açarların siyahısına baxmaq üçün:
```bash
python web_keygen.py --list
```

---

## 🌐 3. Saytı Müştəriyə Necə Göndərmək Olar? (Pulsuz Yollar)

Müştəriyə göndərmək üçün kompüterinizdə işləyən veb serveri internetə açmağın **ən asan və 100% pulsuz** 2 yolu:

### Variant A: Cloudflare Tunnel (Tövsiyə olunur — 0 AZN, Dərhal Link)
1. Pulsuz `cloudflared` proqramını yükləyin və ya terminalda yazın:
   ```bash
   cloudflared tunnel --url http://localhost:8000
   ```
2. Cloudflare sizə dərhal rəsmi HTTPS linki verəcək:
   👉 `https://adiniz-realestate.trycloudflare.com`
3. Bu linki və 24 saatlıq kodu müştəriyə göndərirsiniz. Müştəri telefonundan və ya kompüterindən daxil olur, heç bir fayl yükləmir, kodu yazır və sistemi test edir!

### Variant B: Render.com və ya Railway.app (7/24 Buludda Qalma — Pulsuz)
Kompüterinizi açıq saxlamaq istəmirsinizsə:
1. Bu qovluğu GitHub-a yükləyin (Private repo).
2. [Render.com](https://render.com)-a daxil olub **"New Web Service"** seçin.
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `uvicorn web_server:app --host 0.0.0.0 --port $PORT`
5. Render sizə `https://binaleadpro.onrender.com` kimi daimi pulsuz domen verir.

---

## 🛡️ Təhlükəsizlik və Anti-Reuse Qaydaları

1. **Tək Cihaza Bağlanma:** Müştəri kodu brauzerə daxil etdiyi an həmin brauzerin sessiyasına bağlanır. Həmin kodu başqa adama ötürsə belə, digər adam *"Bu lisenziya artıq başqa cihazda aktivləşdirilib"* xətası alacaq.
2. **24 Saatlıq Geri Sayım:** Kod daxil edildiyi andan etibarən 86 400 saniyə sonra lisenziya avtomatik bitir və ekran bloklanır.
3. **Məlumat Təhlükəsizliyi:** Müştəri yalnız öz axtarış nəticələrini görür və birbaşa brauzerdən WhatsApp-a keçid və ya Excel yükləmə imkanına malik olur.
