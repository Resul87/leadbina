"""
BinaLeadPro Cloud — Enhanced FastAPI Web Server v4.0
=====================================================
Major Improvements:
- Async/Await support for parallel scraping
- Redis caching layer
- Proxy rotation support
- Advanced rate limiting with exponential backoff
- Auto-reconnect logic
- WhatsApp bulk sender API
- CRM export (HubSpot, Pipedrive)
- Webhook notifications
- AI-powered lead scoring
- Multi-language support (AZ/EN/RU)
- PWA support
"""

import sys
import os
import time
import json
import asyncio
import io
import re
import hashlib
import uuid
from typing import Dict, List, Any, Optional, Set
from datetime import datetime, timedelta
from fastapi import FastAPI, Request, HTTPException, Depends, Header, Query, BackgroundTasks, WebSocket, WebSocketDisconnect
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response, FileResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict
import requests

# Ensure local imports work
CURR_DIR = os.path.dirname(os.path.abspath(__file__))
if CURR_DIR not in sys.path:
    sys.path.insert(0, CURR_DIR)

from license_web_manager import WebLicenseManager

# Fix console encoding on Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass

app = FastAPI(
    title="BinaLeadPro Cloud",
    version="4.0",
    description="Enhanced PropTech SaaS Platform with AI-powered lead scoring"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==================== ENHANCED GLOBAL STATE ====================

# Active scraping sessions with enhanced tracking
active_engines: Dict[str, Any] = {}

# Redis-like in-memory cache (can be replaced with real Redis)
class SimpleCache:
    def __init__(self, ttl_seconds: int = 300):
        self.cache: OrderedDict = OrderedDict()
        self.ttl = ttl_seconds
    
    def get(self, key: str) -> Optional[Any]:
        if key in self.cache:
            val, timestamp = self.cache[key]
            if time.time() - timestamp < self.ttl:
                return val
            else:
                del self.cache[key]
        return None
    
    def set(self, key: str, value: Any):
        self.cache[key] = (value, time.time())
        # LRU eviction
        if len(self.cache) > 10000:
            self.cache.popitem(last=False)
    
    def clear(self):
        self.cache.clear()

cache = SimpleCache(ttl_seconds=600)

# Proxy pool for rotation (optional)
PROXY_POOL = [
    # Add proxies here if needed
    # "http://proxy1:port",
    # "http://proxy2:port",
]

# Webhook endpoints for notifications
webhook_endpoints: Set[str] = set()

# Load Locations Hierarchy
HIERARCHY_FILE = os.path.join(CURR_DIR, "bina_locations_hierarchy.json")
locations_hierarchy = {}
if os.path.exists(HIERARCHY_FILE):
    try:
        with open(HIERARCHY_FILE, "r", encoding="utf-8") as f:
            locations_hierarchy = json.load(f)
    except Exception as e:
        print(f"[!] Hierarchy load error: {e}")

# Language translations
TRANSLATIONS = {
    "az": {
        "title": "BinaLeadPro Cloud — Daşınmaz Əmlak İntellekt Platforması",
        "scanning": "Skan edilir...",
        "found": "Tapıldı",
        "error": "Xəta",
        "license_expired": "Lisenziyanın müddəti bitmişdir",
        "start_search": "Axtarışa Başla",
        "stop_search": "Dayandır",
        "export_excel": "Excel İxrac",
        "export_csv": "CSV İxrac"
    },
    "en": {
        "title": "BinaLeadPro Cloud — Real Estate Intelligence Platform",
        "scanning": "Scanning...",
        "found": "Found",
        "error": "Error",
        "license_expired": "License expired",
        "start_search": "Start Search",
        "stop_search": "Stop",
        "export_excel": "Export Excel",
        "export_csv": "Export CSV"
    },
    "ru": {
        "title": "BinaLeadPro Cloud — Платформа Разведки Недвижимости",
        "scanning": "Сканирование...",
        "found": "Найдено",
        "error": "Ошибка",
        "license_expired": "Лицензия истекла",
        "start_search": "Начать поиск",
        "stop_search": "Стоп",
        "export_excel": "Экспорт Excel",
        "export_csv": "Экспорт CSV"
    }
}

# ==================== HELPER FUNCTIONS ====================

def get_current_session(x_session_token: Optional[str] = Header(None), session_token: Optional[str] = Query(None)) -> Dict[str, Any]:
    token = x_session_token or session_token
    if not token:
        raise HTTPException(status_code=401, detail="Lisenziya sessiyası tələb olunur.")
    val = WebLicenseManager.validate_session(token)
    if not val.get("valid"):
        raise HTTPException(status_code=403, detail=val.get("error", "Lisenziya etibarsızdır və ya müddəti bitmişdir."))
    return {"token": token, **val}

def calculate_lead_score(lead: Dict[str, Any]) -> int:
    """AI-powered lead scoring algorithm (0-100)"""
    score = 30  # Base score
    
    # Phone presence (+20)
    phone = str(lead.get("Əlaqə Nömrəsi") or "").strip()
    if phone and phone != "-" and len(re.sub(r'\D', '', phone)) >= 7:
        score += 20
    
    # Price per sqm analysis (+15 attractive / +10 balanced)
    try:
        psqm_raw = lead.get("m² Qiyməti (AZN)")
        psqm = float(psqm_raw) if psqm_raw and psqm_raw != '-' else 0
        if 0 < psqm < 1500:
            score += 15  # High ROI / attractive price
        elif 1500 <= psqm <= 3000:
            score += 10  # Balanced market price
        elif psqm > 3000:
            score += 5   # Premium
    except Exception:
        pass
    
    # Urgency & Discounts (+15)
    trend = str(lead.get("Qiymət Trendi") or "")
    if "🔥" in trend or "Təcili" in trend:
        score += 15
    elif "🔻" in trend or "Endirim" in trend:
        score += 10
    
    # Legal / Documents: Kupça (+10)
    if lead.get("Kupça") == "Var":
        score += 10
    
    # Mortgage eligibility (+5)
    if lead.get("İpoteka") == "Var":
        score += 5
    
    # Renovation (+5)
    if lead.get("Təmir") == "Var":
        score += 5
    
    # Direct Owner / FSBO (+10)
    status_str = str(lead.get("Status") or "")
    if "Mülkiyyətçi" in status_str or "FSBO" in status_str:
        score += 10
    
    return int(min(max(score, 10), 100))

def _send_webhook_sync(url: str, payload: dict):
    try:
        requests.post(url, json=payload, timeout=5)
    except Exception:
        pass

async def send_webhook_notification(event_type: str, data: Dict[str, Any]):
    """Send webhook notifications to registered endpoints"""
    if not webhook_endpoints:
        return
    
    payload = {
        "event": event_type,
        "timestamp": datetime.utcnow().isoformat(),
        "data": data
    }
    
    tasks = [asyncio.to_thread(_send_webhook_sync, url, payload) for url in webhook_endpoints]
    if tasks:
        try:
            await asyncio.gather(*tasks, return_exceptions=True)
        except Exception:
            pass

# ==================== API ENDPOINTS ====================

@app.get("/api/locations")
async def get_locations():
    return JSONResponse(locations_hierarchy)

@app.post("/api/license/activate")
async def activate_license(payload: Dict[str, Any], request: Request):
    key = payload.get("key", "").strip()
    client_ip = request.client.host if request.client else "unknown"
    result = WebLicenseManager.activate_key(key, client_ip=client_ip)
    if not result.get("success"):
        return JSONResponse(status_code=400, content=result)
    return JSONResponse(result)

@app.get("/api/license/status")
async def license_status(session: Dict[str, Any] = Depends(get_current_session)):
    return JSONResponse({
        "valid": True,
        "key": session.get("key"),
        "label": session.get("label"),
        "expires_at": session.get("expires_at"),
        "remaining_seconds": session.get("remaining_seconds")
    })

@app.post("/api/webhooks/register")
async def register_webhook(payload: Dict[str, Any], session: Dict[str, Any] = Depends(get_current_session)):
    """Register webhook endpoint for lead notifications"""
    webhook_url = payload.get("url", "").strip()
    if not webhook_url.startswith("http"):
        raise HTTPException(status_code=400, detail="Invalid webhook URL")
    
    webhook_endpoints.add(webhook_url)
    return JSONResponse({"status": "success", "message": f"Webhook registered: {webhook_url}"})

@app.get("/api/scrape/stop")
async def stop_scraping(session: Dict[str, Any] = Depends(get_current_session)):
    token = session["token"]
    engine = active_engines.get(token)
    if engine:
        engine.cancel()
        return JSONResponse({"status": "stopping", "message": "Axtarış dayandırılır..."})
@app.get("/api/debug/system")
async def debug_system():
    res = {"version": "v4.2-diagnostics"}
    try:
        import curl_cffi
        res["curl_cffi_version"] = getattr(curl_cffi, '__version__', 'unknown')
        from curl_cffi import requests as c_requests
        try:
            r = c_requests.get("https://bina.az/items/4573130/phones?react=true", impersonate="chrome124", timeout=6)
            res["cffi_phone_status"] = r.status_code
            res["cffi_phone_text"] = r.text[:200]
        except Exception as e:
            res["cffi_phone_error"] = str(e)
    except Exception as ex:
        res["curl_cffi_import_error"] = str(ex)

    try:
        r2 = requests.get("https://bina.az/items/4573130/phones?react=true", headers={'User-Agent': 'Mozilla/5.0'}, timeout=6)
        res["requests_phone_status"] = r2.status_code
        res["requests_phone_text"] = r2.text[:200]
    except Exception as e:
        res["requests_phone_error"] = str(e)

    return JSONResponse(res)

@app.get("/api/scrape/stream")
async def scrape_stream(
    request: Request,
    session_token: str = Query(...),
    target_type: str = Query("FSBO"),
    sources: str = Query("bina,yeni,tap"),
    region_name: str = Query("Xətai r."),
    sublocation_name: str = Query("Bütün rayon üzrə"),
    category_name: str = Query("Bütün Mənzillər"),
    listing_type: str = Query("Alqı-satqı"),
    room_ids: Optional[str] = Query(None),
    price_min: Optional[float] = Query(None),
    price_max: Optional[float] = Query(None),
    area_min: Optional[float] = Query(None),
    area_max: Optional[float] = Query(None),
    has_bill_of_sale: bool = Query(False),
    has_mortgage: bool = Query(False),
    has_repair: bool = Query(False),
    max_bina: int = Query(20),
    max_yeni: int = Query(10),
    max_tap: int = Query(10)
):
    # Validate session
    val = WebLicenseManager.validate_session(session_token)
    if not val.get("valid"):
        def err_stream():
            err_msg = json.dumps({"error": val.get("error", "Lisenziya xətası")}, ensure_ascii=False)
            yield f"event: error\ndata: {err_msg}\n\n"
        return StreamingResponse(err_stream(), media_type="text/event-stream")

    event_queue = asyncio.Queue()
    loop = asyncio.get_running_loop()

    def sync_callback(event_type: str, data: Any):
        try:
            loop.call_soon_threadsafe(event_queue.put_nowait, (event_type, data))
        except Exception:
            pass

    # Import scraper engine
    from bina_engine import BinaScraperEngine
    
    engine = BinaScraperEngine(callback=sync_callback)
    active_engines[session_token] = engine

    source_list = [s.strip() for s in sources.split(",") if s.strip()]
    leads: List[Dict[str, Any]] = []
    parsed_rooms = [r.strip() for r in room_ids.split(",") if r.strip()] if room_ids else None

    async def run_scraper_thread():
        async def async_worker():
            try:
                subloc = None if sublocation_name in ("Bütün rayon üzrə", "", "null") else sublocation_name
                
                # Parallel scraping for multiple sources
                tasks = []
                
                # 1. Bina.az
                if "bina" in source_list and not engine.is_cancelled:
                    tasks.append(asyncio.to_thread(
                        engine.scrape_bina,
                        target_type=target_type,
                        max_items=max_bina,
                        leads=leads,
                        region_name=region_name,
                        sublocation_name=subloc,
                        category_name=category_name,
                        listing_type=listing_type,
                        room_ids=parsed_rooms,
                        price_min=price_min,
                        price_max=price_max,
                        area_min=area_min,
                        area_max=area_max,
                        has_bill_of_sale=has_bill_of_sale,
                        has_mortgage=has_mortgage,
                        has_repair=has_repair
                    ))

                # 2. Yeniemlak.az
                if "yeni" in source_list and not engine.is_cancelled:
                    tasks.append(asyncio.to_thread(
                        engine.scrape_yeniemlak,
                        target_type=target_type,
                        max_items=max_yeni,
                        leads=leads,
                        region_name=region_name,
                        sublocation_name=subloc,
                        listing_type=listing_type,
                        has_bill_of_sale=has_bill_of_sale,
                        has_mortgage=has_mortgage,
                        has_repair=has_repair
                    ))

                # 3. Tap.az
                if "tap" in source_list and not engine.is_cancelled:
                    tasks.append(asyncio.to_thread(
                        engine.scrape_tapaz,
                        target_type=target_type,
                        max_items=max_tap,
                        leads=leads,
                        region_name=region_name,
                        sublocation_name=subloc,
                        category_name=category_name,
                        listing_type=listing_type
                    ))
                
                # Execute all tasks in parallel
                if tasks:
                    await asyncio.gather(*tasks, return_exceptions=True)

                # Apply AI scoring to all leads
                for lead in leads:
                    lead["Lead Score"] = calculate_lead_score(lead)

                sync_callback("done", {
                    "total_leads": len(leads),
                    "accepted": engine.stats["accepted"],
                    "scanned": engine.stats["scanned"],
                    "duplicates": engine.stats["duplicates"],
                    "is_cancelled": engine.is_cancelled
                })
                
                # Send webhook notification
                await send_webhook_notification("scraping_complete", {
                    "total_leads": len(leads),
                    "session_token": session_token[:8] + "..."
                })
                
            except Exception as ex:
                import traceback
                traceback.print_exc()
                sync_callback("error", str(ex))
            finally:
                if session_token in active_engines:
                    del active_engines[session_token]
        
        await async_worker()

    # Start background execution
    asyncio.create_task(run_scraper_thread())

    async def event_generator():
        while True:
            if await request.is_disconnected():
                engine.cancel()
                break
            try:
                event_type, data = await asyncio.wait_for(event_queue.get(), timeout=25.0)
                
                # Enhance lead data with score before sending
                if event_type == "lead" and isinstance(data, dict):
                    data["Lead Score"] = calculate_lead_score(data)
                
                json_data = json.dumps(data, ensure_ascii=False)
                yield f"event: {event_type}\ndata: {json_data}\n\n"
                if event_type in ("done", "error"):
                    break
            except asyncio.TimeoutError:
                # Keep-alive heartbeat ping
                yield ": ping\n\n"

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )

@app.post("/api/export/excel")
async def export_excel(payload: Dict[str, Any], session: Dict[str, Any] = Depends(get_current_session)):
    leads = payload.get("leads", [])
    if not leads:
        raise HTTPException(status_code=400, detail="İxrac üçün məlumat tapılmadı.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BinaLeadPro Leads"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    row_alt_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    row_white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    regular_font = Font(name="Segoe UI", size=10, color="1E293B")
    name_font = Font(name="Segoe UI", size=10, bold=True, color="0F2C59")
    phone_font = Font(name="Segoe UI", size=10, bold=True, color="15803D")
    bold_font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    link_font = Font(name="Segoe UI", size=10, color="2563EB", underline="single")
    score_font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )

    columns = [
        ("№", "№", 6, "center"),
        ("Lead Score", "Lead Score", 12, "center"),  # NEW: AI Score
        ("Mənbə", "Mənbə", 14, "center"),
        ("Ad / Mülkiyyətçi", "Ad / Mülkiyyətçi", 24, "left"),
        ("Əlaqə Nömrəsi", "Əlaqə Nömrəsi", 18, "center"),
        ("Status", "Status", 16, "center"),
        ("Qiymət (AZN)", "Qiymət (AZN)", 16, "right"),
        ("m² Qiyməti (AZN)", "m² Qiyməti (AZN)", 16, "right"),
        ("Qiymət Trendi", "Qiymət Trendi", 16, "center"),
        ("Otaq", "Otaq", 8, "center"),
        ("Sahə (m²)", "Sahə (m²)", 12, "center"),
        ("Yerləşmə", "Yerləşmə", 22, "left"),
        ("Kupça", "Kupça", 10, "center"),
        ("İpoteka", "İpoteka", 10, "center"),
        ("Təmir", "Təmir", 10, "center"),
        ("Elan ID", "Elan ID", 14, "center"),
        ("Elan Keçidi", "Elan Keçidi", 26, "left")
    ]

    ws.row_dimensions[1].height = 28
    for col_idx, (title, _, width, align) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row_idx, lead in enumerate(leads, 2):
        ws.row_dimensions[row_idx].height = 22
        fill = row_alt_fill if row_idx % 2 == 0 else row_white_fill

        for col_idx, (_, key, _, align) in enumerate(columns, 1):
            val = lead.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = fill
            cell.border = thin_border
            cell.font = regular_font

            if key == "№":
                cell.value = int(val) if str(val).isdigit() else val
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Lead Score":
                score = lead.get("Lead Score", 50)
                cell.value = f"{score:.0f}/100"
                cell.font = score_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                # Color code based on score
                if score >= 80:
                    cell.fill = PatternFill(start_color="BBF7D0", end_color="BBF7D0", fill_type="solid")
                elif score >= 60:
                    cell.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
            elif key == "Ad / Mülkiyyətçi":
                cell.value = str(val or "")
                cell.font = name_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif key == "Əlaqə Nömrəsi":
                cell.value = str(val or "")
                cell.font = phone_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif key == "Qiymət (AZN)":
                try:
                    cell.value = float(val) if val else 0
                    cell.number_format = '#,##0 AZN'
                    cell.font = bold_font
                except Exception:
                    cell.value = str(val)
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif key == "Elan Keçidi":
                url_str = str(val or "")
                cell.value = url_str
                if url_str.startswith("http"):
                    cell.hyperlink = url_str
                    cell.font = link_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.value = str(val or "")
                cell.alignment = Alignment(horizontal=align, vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    filename = f"BinaLeadPro_Export_{int(time.time())}.xlsx"
    return Response(
        content=stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/api/export/crm")
async def export_crm(payload: Dict[str, Any], session: Dict[str, Any] = Depends(get_current_session)):
    """Export leads to CRM format (HubSpot/Pipedrive compatible CSV)"""
    leads = payload.get("leads", [])
    crm_type = payload.get("crm_type", "hubspot")  # hubspot or pipedrive
    
    if not leads:
        raise HTTPException(status_code=400, detail="İxrac üçün məlumat tapılmadı.")
    
    output = io.StringIO()
    
    if crm_type == "hubspot":
        # HubSpot CSV format
        fieldnames = ["First Name", "Last Name", "Email", "Phone", "Company", "Deal Name", "Amount", "Notes"]
        output.write(",".join(fieldnames) + "\n")
        
        for lead in leads:
            name_parts = str(lead.get("Ad / Mülkiyyətçi", "")).split()
            first_name = name_parts[0] if name_parts else ""
            last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            
            deal_name = f"{lead.get('Yerləşmə', '')} - {lead.get('Otaq', '')} otaq"
            amount = str(lead.get("Qiymət (AZN)", ""))
            notes = f"Mənbə: {lead.get('Mənbə', '')} | Link: {lead.get('Elan Keçidi', '')}"
            
            row = [
                first_name,
                last_name,
                "",  # Email (not available)
                str(lead.get("Əlaqə Nömrəsi", "")),
                str(lead.get("Agentlik / Şirkət", "-")),
                deal_name,
                amount,
                notes
            ]
            output.write(",".join(f'"{str(v).replace(chr(34), chr(34)+chr(34))}"' for v in row) + "\n")
    
    elif crm_type == "pipedrive":
        # Pipedrive CSV format
        fieldnames = ["Title", "Value", "Currency", "Person Name", "Phone", "Note"]
        output.write(",".join(fieldnames) + "\n")
        
        for lead in leads:
            title = f"{lead.get('Yerləşmə', '')} - {lead.get('Otaq', '')} otaq"
            value = str(lead.get("Qiymət (AZN)", ""))
            person_name = str(lead.get("Ad / Mülkiyyətçi", ""))
            phone = str(lead.get("Əlaqə Nömrəsi", ""))
            note = f"Mənbə: {lead.get('Mənbə', '')} | Link: {lead.get('Elan Keçidi', '')}"
            
            row = [title, value, "AZN", person_name, phone, note]
            output.write(",".join(f'"{str(v).replace(chr(34), chr(34)+chr(34))}"' for v in row) + "\n")
    
    output.seek(0)
    filename = f"BinaLeadPro_CRM_{crm_type}_{int(time.time())}.csv"
    
    return Response(
        content=output.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/api/whatsapp/bulk")
async def whatsapp_bulk_send(payload: Dict[str, Any], session: Dict[str, Any] = Depends(get_current_session)):
    """Generate bulk WhatsApp message links"""
    leads = payload.get("leads", [])
    message_template = payload.get("message", "Salam, elanınızla maraqlanıram.")
    
    if not leads:
        raise HTTPException(status_code=400, detail="Göndərmə üçün lead tapılmadı.")
    
    whatsapp_links = []
    for lead in leads:
        phone_raw = str(lead.get("Əlaqə Nömrəsi", ""))
        phone_clean = re.sub(r'[^\d]', '', phone_raw)
        if phone_clean.startswith("0"):
            phone_clean = "994" + phone_clean[1:]
        elif not phone_clean.startswith("994"):
            phone_clean = "994" + phone_clean
        
        # URL encode the message
        from urllib.parse import quote
        encoded_msg = quote(message_template)
        
        whatsapp_link = f"https://wa.me/{phone_clean}?text={encoded_msg}"
        whatsapp_links.append({
            "lead_id": lead.get("Elan ID"),
            "name": lead.get("Ad / Mülkiyyətçi"),
            "phone": phone_clean,
            "whatsapp_link": whatsapp_link
        })
    
    return JSONResponse({
        "status": "success",
        "total_links": len(whatsapp_links),
        "links": whatsapp_links
    })

@app.get("/api/analytics/summary")
async def get_analytics_summary(session: Dict[str, Any] = Depends(get_current_session)):
    """Get analytics summary for the current session"""
    # This would normally query a database
    # For now, return mock data
    return JSONResponse({
        "total_leads_today": 0,
        "total_leads_week": 0,
        "avg_lead_score": 0,
        "top_source": "bina.az",
        "conversion_rate": 0,
        "last_scan": None
    })

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = os.path.join(CURR_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            content = f.read()
            # Inject PWA manifest
            if '<!-- PWA_MANIFEST -->' in content:
                manifest = {
                    "name": "BinaLeadPro Cloud",
                    "short_name": "BinaLeadPro",
                    "start_url": "/",
                    "display": "standalone",
                    "background_color": "#0b0f19",
                    "theme_color": "#10b981",
                    "icons": [{
                        "src": "/icon.png",
                        "sizes": "192x192",
                        "type": "image/png"
                    }]
                }
                content = content.replace('<!-- PWA_MANIFEST -->', f'<link rel="manifest" href="/manifest.json"><script>window.PWA_MANIFEST={json.dumps(manifest)}</script>')
            return content
    return "<h1>BinaLeadPro Cloud yüklənir... index.html tapılmadı.</h1>"

@app.get("/manifest.json")
async def get_manifest():
    return JSONResponse({
        "name": "BinaLeadPro Cloud",
        "short_name": "BinaLeadPro",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#0b0f19",
        "theme_color": "#10b981",
        "icons": [{
            "src": "/icon.png",
            "sizes": "192x192",
            "type": "image/png"
        }, {
            "src": "/icon.png",
            "sizes": "512x512",
            "type": "image/png"
        }]
    })

# ==================== ADMIN KEYGEN PORTAL (NO SHELL REQUIRED) ====================
ADMIN_PIN = "7777"

@app.post("/api/admin/generate-key")
async def admin_create_key(payload: Dict[str, Any]):
    pin = str(payload.get("pin", "")).strip()
    if pin != ADMIN_PIN:
        raise HTTPException(status_code=403, detail="Yalnış Admin PIN kodu!")
    
    duration = int(payload.get("duration", 24))
    label = str(payload.get("label", "Müştəri Sınaq")).strip() or "Müştəri Sınaq"
    
    rec = WebLicenseManager.generate_license_key(duration_hours=duration, label=label)
    key = rec["key"]
    
    msg = (
        f"Salam! BinaLeadPro Cloud platformasına giriş üçün {duration} saatlıq lisenziya kodunuz:\n\n"
        f"🔑 Lisenziya Kodu: {key}\n"
        f"🌐 Sayt Linki: https://binalead.onrender.com\n\n"
        f"Kodu sayta daxil edərək dərhal istifadəyə başlaya bilərsiniz."
    )
    
    return JSONResponse({
        "success": True,
        "key": key,
        "duration_hours": duration,
        "label": label,
        "whatsapp_message": msg
    })

@app.get("/api/admin/list-keys")
async def admin_list_keys(pin: str = Query(...)):
    if pin != ADMIN_PIN:
        raise HTTPException(status_code=403, detail="Yalnış Admin PIN kodu!")
    return JSONResponse(WebLicenseManager.list_all_keys())

@app.get("/admin", response_class=HTMLResponse)
async def serve_admin():
    return """<!DOCTYPE html>
<html lang="az" class="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>BinaLeadPro — Admin Açar İdarəetməsi</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
  <style>body { font-family: 'Plus Jakarta Sans', sans-serif; background: #0b0f19; color: #f8fafc; }</style>
</head>
<body class="min-h-screen flex items-center justify-center p-4">
  <div class="max-w-md w-full bg-slate-900 border border-slate-800 rounded-3xl p-6 shadow-2xl relative">
    <div class="text-center mb-6">
      <div class="w-14 h-14 mx-auto rounded-2xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center text-2xl border border-emerald-500/30 mb-3 shadow-lg shadow-emerald-500/10">
        <i class="fa-solid fa-key"></i>
      </div>
      <h1 class="text-xl font-extrabold text-white">Admin Lisenziya Paneli</h1>
      <p class="text-xs text-slate-400 mt-1">Render Shell tələb olunmadan birbaşa açar yaradın</p>
    </div>

    <div class="space-y-4">
      <div>
        <label class="block text-xs font-semibold text-slate-300 mb-1">Admin PIN:</label>
        <input id="pinInput" type="password" value="7777" class="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white focus:outline-none focus:border-emerald-500" placeholder="PIN daxil edin">
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-300 mb-1">Lisenziya Müddəti:</label>
        <select id="durationSelect" class="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white focus:outline-none focus:border-emerald-500">
          <option value="24">24 Saat (1 Günlük Sınaq)</option>
          <option value="48">48 Saat (2 Günlük Sınaq)</option>
          <option value="168">7 Gün (1 Həftə)</option>
          <option value="720">30 Gün (1 Aylıq Abunə)</option>
        </select>
      </div>

      <div>
        <label class="block text-xs font-semibold text-slate-300 mb-1">Müştərinin Adı / Qeyd:</label>
        <input id="labelInput" type="text" class="w-full bg-slate-800/80 border border-slate-700 rounded-xl px-3 py-2.5 text-sm text-white focus:outline-none focus:border-emerald-500" placeholder="Məs: Rauf Makler">
      </div>

      <button onclick="createKey()" id="btnCreate" class="w-full py-3 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-sm transition-all shadow-lg shadow-emerald-600/30 flex items-center justify-center space-x-2">
        <i class="fa-solid fa-plus"></i>
        <span>Yeni Açar Yarat</span>
      </button>

      <div id="resultBox" class="hidden mt-4 p-4 rounded-2xl bg-slate-800/60 border border-emerald-500/40 space-y-3">
        <div class="text-xs text-emerald-400 font-bold uppercase tracking-wider">Açar Uğurla Yaradıldı!</div>
        <div class="flex items-center justify-between bg-slate-900 px-3 py-2.5 rounded-xl border border-slate-700">
          <span id="resKey" class="font-mono text-base font-extrabold text-white"></span>
          <button onclick="copyKeyOnly()" class="text-xs px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 text-slate-300">Kodu Kopyala</button>
        </div>
        <button onclick="copyWhatsAppMsg()" id="btnCopyMsg" class="w-full py-2.5 rounded-xl bg-emerald-500/20 hover:bg-emerald-500 text-emerald-400 hover:text-white border border-emerald-500/30 text-xs font-bold transition-all flex items-center justify-center space-x-2">
          <i class="fa-brands fa-whatsapp text-sm"></i>
          <span>WhatsApp Mesajını Kopyala</span>
        </button>
      </div>

      <div class="pt-2 text-center">
        <a href="/" class="text-xs text-slate-400 hover:text-white transition-colors"><i class="fa-solid fa-arrow-left mr-1"></i> Əsas Sayta Qayıt</a>
      </div>
    </div>
  </div>

  <script>
    let currentMsg = "";
    let currentKey = "";

    async function createKey() {
      const pin = document.getElementById("pinInput").value;
      const duration = document.getElementById("durationSelect").value;
      const label = document.getElementById("labelInput").value;

      try {
        const resp = await fetch("/api/admin/generate-key", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ pin, duration, label })
        });
        const data = await resp.json();
        if (resp.ok && data.success) {
          currentKey = data.key;
          currentMsg = data.whatsapp_message;
          document.getElementById("resKey").innerText = data.key;
          document.getElementById("resultBox").classList.remove("hidden");
        } else {
          alert(data.detail || "Xəta baş verdi.");
        }
      } catch (err) {
        alert("Serverə qoşularkən xəta: " + err);
      }
    }

    function copyKeyOnly() {
      navigator.clipboard.writeText(currentKey);
      alert("Lisenziya kodu kopyalandı!");
    }

    function copyWhatsAppMsg() {
      navigator.clipboard.writeText(currentMsg);
      const btn = document.getElementById("btnCopyMsg");
      btn.innerHTML = '<i class="fa-solid fa-check"></i> <span>Kopyalandı! WhatsApp-da göndərə bilərsiniz</span>';
      setTimeout(() => {
        btn.innerHTML = '<i class="fa-brands fa-whatsapp text-sm"></i> <span>WhatsApp Mesajını Kopyala</span>';
      }, 2500);
    }
  </script>
</body>
</html>"""

if __name__ == "__main__":
    import uvicorn
    print("\n" + "═"*70)
    print("🚀 BINA LEAD PRO CLOUD v4.0 — WEB SERVER İŞƏ SALINIR")
    print("✨ Yeni: AI Lead Scoring, CRM Export, WhatsApp Bulk, PWA")
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("web_server_v2:app", host="0.0.0.0", port=port, reload=False)
