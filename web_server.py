"""
BinaLeadPro Cloud — FastAPI Web Server
======================================
Provides modern web SaaS interface, real-time scraping via Server-Sent Events (SSE),
browser session authentication, and direct Excel/CSV exports.
"""

import sys
import os
import time
import json
import asyncio
import io
import re
from typing import Dict, List, Any, Optional
from fastapi import FastAPI, Request, HTTPException, Depends, Header, Query, BackgroundTasks
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Ensure local imports work
CURR_DIR = os.path.dirname(os.path.abspath(__file__))
if CURR_DIR not in sys.path:
    sys.path.insert(0, CURR_DIR)

from license_web_manager import WebLicenseManager
from bina_engine import BinaScraperEngine, YENIEMLAK_RAYON_MAP, AGENT_KEYWORDS

# Fix console encoding on Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass

app = FastAPI(title="BinaLeadPro Cloud", version="3.5")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Active scraping tasks per session
active_engines: Dict[str, BinaScraperEngine] = {}

# Load Locations Hierarchy
HIERARCHY_FILE = os.path.join(CURR_DIR, "bina_locations_hierarchy.json")
locations_hierarchy = {}
if os.path.exists(HIERARCHY_FILE):
    try:
        with open(HIERARCHY_FILE, "r", encoding="utf-8") as f:
            locations_hierarchy = json.load(f)
    except Exception as e:
        print(f"[!] Hierarchy load error: {e}")


def get_current_session(x_session_token: Optional[str] = Header(None), session_token: Optional[str] = Query(None)) -> Dict[str, Any]:
    token = x_session_token or session_token
    if not token:
        raise HTTPException(status_code=401, detail="Lisenziya sessiyası tələb olunur.")
    val = WebLicenseManager.validate_session(token)
    if not val.get("valid"):
        raise HTTPException(status_code=403, detail=val.get("error", "Lisenziya etibarsızdır və ya müddəti bitmişdir."))
    return {"token": token, **val}


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


@app.get("/api/scrape/stop")
async def stop_scraping(session: Dict[str, Any] = Depends(get_current_session)):
    token = session["token"]
    engine = active_engines.get(token)
    if engine:
        engine.cancel()
        return JSONResponse({"status": "stopping", "message": "Axtarış dayandırılır..."})
    return JSONResponse({"status": "idle", "message": "Aktiv axtarış tapılmadı."})


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
    loop = asyncio.get_event_loop()

    def sync_callback(event_type: str, data: Any):
        try:
            loop.call_soon_threadsafe(event_queue.put_nowait, (event_type, data))
        except Exception:
            pass

    engine = BinaScraperEngine(callback=sync_callback)
    active_engines[session_token] = engine

    source_list = [s.strip() for s in sources.split(",") if s.strip()]
    leads: List[Dict[str, Any]] = []
    parsed_rooms = [r.strip() for r in room_ids.split(",") if r.strip()] if room_ids else None

    async def run_scraper_thread():
        def blocking_worker():
            try:
                subloc = None if sublocation_name in ("Bütün rayon üzrə", "", "null") else sublocation_name
                
                # 1. Bina.az
                if "bina" in source_list and not engine.is_cancelled:
                    engine.scrape_bina(
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
                    )

                # 2. Yeniemlak.az
                if "yeni" in source_list and not engine.is_cancelled:
                    engine.scrape_yeniemlak(
                        target_type=target_type,
                        max_items=max_yeni,
                        leads=leads,
                        region_name=region_name,
                        sublocation_name=subloc,
                        listing_type=listing_type,
                        has_bill_of_sale=has_bill_of_sale,
                        has_mortgage=has_mortgage,
                        has_repair=has_repair
                    )

                # 3. Tap.az
                if "tap" in source_list and not engine.is_cancelled:
                    engine.scrape_tapaz(
                        target_type=target_type,
                        max_items=max_tap,
                        leads=leads,
                        region_name=region_name,
                        sublocation_name=subloc,
                        category_name=category_name,
                        listing_type=listing_type
                    )

                sync_callback("done", {
                    "total_leads": len(leads),
                    "accepted": engine.stats["accepted"],
                    "scanned": engine.stats["scanned"],
                    "duplicates": engine.stats["duplicates"],
                    "is_cancelled": engine.is_cancelled
                })
            except Exception as ex:
                import traceback
                traceback.print_exc()
                sync_callback("error", str(ex))
            finally:
                if session_token in active_engines:
                    del active_engines[session_token]

        await asyncio.to_thread(blocking_worker)

    # Start background execution
    asyncio.create_task(run_scraper_thread())

    async def event_generator():
        while True:
            if await request.is_disconnected():
                engine.cancel()
                break
            try:
                event_type, data = await asyncio.wait_for(event_queue.get(), timeout=25.0)
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

    thin_border = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )

    columns = [
        ("№", "№", 6, "center"),
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


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_file = os.path.join(CURR_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>BinaLeadPro Cloud yüklənir... index.html tapılmadı.</h1>"


if __name__ == "__main__":
    import uvicorn
    print("\n" + "═"*70)
    print("🚀 BİNALAUNCH PRO CLOUD — WEB SERVER İŞƏ SALINIR")
    print("🌐 Brauzerdə daxil olun: http://127.0.0.1:8000")
    print("═"*70 + "\n")
    uvicorn.run("web_server:app", host="0.0.0.0", port=8000, reload=False)
