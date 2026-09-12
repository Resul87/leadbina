"""
BinaLeadPro Multi v3.0 — Modern Desktop Application & Database Viewer
=====================================================================
Built with CustomTkinter & ttk for high-performance, beautiful UI on Windows.
Features:
- In-App Interactive Database Table (Treeview) with live row updates & double-click to open in browser.
- Per-platform count settings: bina.az, yeniemlak.az, tap.az.
- Sub-location hierarchy (Rayon -> Metro / Settlement dropdown).
- Direct clickable links in UI and Excel.
- 24h/48h trial and subscription manager.
"""

import sys
import os
import threading
import time
import webbrowser
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Any, Optional, Dict, List
import customtkinter as ctk

# Ensure correct working directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

from bina_engine import BinaScraperEngine, HIERARCHY_MAP, CATEGORIES_MAP
import license_manager

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class BinaLeadProApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("BinaLeadPro Multi v3.0 — B2B Əmlak Baza Çıxarıcı & Daxili Baza")
        self.geometry("1260x860")
        self.minsize(1100, 750)

        self.engine: BinaScraperEngine = None
        self.worker_thread: threading.Thread = None
        self.last_excel_path = None
        self.all_leads: List[Dict[str, Any]] = []

        self._init_layout()
        self._check_trial_timer()

    def _init_layout(self):
        # 2-column layout: Left Filters (400px), Right Main Tabs (Table & Log)
        self.grid_columnconfigure(0, weight=0, minsize=400)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # -------------------------------------------------------------
        # LEFT SIDEBAR: FILTERS & TARGET CONTROLS
        # -------------------------------------------------------------
        self.left_frame = ctk.CTkScrollableFrame(self, width=400, corner_radius=0)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)

        # Brand Title
        self.logo_label = ctk.CTkLabel(
            self.left_frame,
            text="🏢 BinaLeadPro Multi v3",
            font=ctk.CTkFont(size=23, weight="bold"),
            text_color="#38BDF8"
        )
        self.logo_label.pack(anchor="w", padx=20, pady=(16, 2))

        self.sublogo_label = ctk.CTkLabel(
            self.left_frame,
            text="B2B Mülkiyyətçi və Agent Baza Sistemi",
            font=ctk.CTkFont(size=12),
            text_color="#94A3B8"
        )
        self.sublogo_label.pack(anchor="w", padx=20, pady=(0, 12))

        # Trial / License Box
        self.license_card = ctk.CTkFrame(self.left_frame, fg_color="#1E293B", corner_radius=10)
        self.license_card.pack(fill="x", padx=15, pady=(0, 14))

        self.license_title = ctk.CTkLabel(
            self.license_card,
            text="🟢 Sınaq Rejimi: Yoxlanılır...",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#4ADE80"
        )
        self.license_title.pack(anchor="w", padx=12, pady=(10, 2))

        # License ID display
        self.hwid_str = license_manager.get_hwid()
        self.hwid_frame = ctk.CTkFrame(self.license_card, fg_color="transparent")
        self.hwid_frame.pack(fill="x", padx=12, pady=(2, 6))

        self.hwid_label = ctk.CTkLabel(
            self.hwid_frame,
            text=f"🔑 Lisenziya ID: {self.hwid_str}",
            font=ctk.CTkFont(size=11, family="Consolas", weight="bold"),
            text_color="#38BDF8"
        )
        self.hwid_label.pack(side="left")

        self.copy_hwid_btn = ctk.CTkButton(
            self.hwid_frame,
            text="📋 Kopyala",
            width=70,
            height=22,
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color="#0284C7",
            hover_color="#0369A1",
            command=self._copy_hwid_to_clipboard
        )
        self.copy_hwid_btn.pack(side="right")

        self.activate_btn = ctk.CTkButton(
            self.license_card,
            text="⚡ Aktivasiya Açarı Daxil Et",
            height=28,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#334155",
            hover_color="#475569",
            command=self._open_activation_dialog
        )
        self.activate_btn.pack(fill="x", padx=12, pady=(0, 10))

        # Target Type (FSBO vs AGENT)
        self.target_label = ctk.CTkLabel(self.left_frame, text="👥 Kimi Çıxarmaq İstəyirsiniz?", font=ctk.CTkFont(size=13, weight="bold"))
        self.target_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.target_segment = ctk.CTkSegmentedButton(
            self.left_frame,
            values=["Mülkiyyətçi (FSBO)", "Əmlak Agenti (Rieltor)"],
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8"
        )
        self.target_segment.set("Mülkiyyətçi (FSBO)")
        self.target_segment.pack(fill="x", padx=15, pady=(0, 14))

        # Platform Selection & Individual Counts Box
        self.source_box = ctk.CTkFrame(self.left_frame, fg_color="#1E293B", corner_radius=10)
        self.source_box.pack(fill="x", padx=15, pady=(0, 14))

        self.source_title = ctk.CTkLabel(
            self.source_box,
            text="🌐 Mənbələr və Çıxarılacaq Saylar:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#38BDF8"
        )
        self.source_title.pack(anchor="w", padx=12, pady=(8, 6))

        # 1. bina.az row
        self.bina_row = ctk.CTkFrame(self.source_box, fg_color="transparent")
        self.bina_row.pack(fill="x", padx=12, pady=3)
        self.src_bina_var = ctk.BooleanVar(value=True)
        self.src_bina_cb = ctk.CTkCheckBox(self.bina_row, text="bina.az", variable=self.src_bina_var, width=120)
        self.src_bina_cb.pack(side="left")
        self.src_bina_count = ctk.CTkEntry(self.bina_row, width=65, height=28)
        self.src_bina_count.insert(0, "20")
        self.src_bina_count.pack(side="right")
        self.bina_lbl = ctk.CTkLabel(self.bina_row, text="ədəd", font=ctk.CTkFont(size=11), text_color="#94A3B8")
        self.bina_lbl.pack(side="right", padx=5)

        # 2. yeniemlak.az row
        self.yeni_row = ctk.CTkFrame(self.source_box, fg_color="transparent")
        self.yeni_row.pack(fill="x", padx=12, pady=3)
        self.src_yeni_var = ctk.BooleanVar(value=True)
        self.src_yeni_cb = ctk.CTkCheckBox(self.yeni_row, text="yeniemlak.az", variable=self.src_yeni_var, width=120)
        self.src_yeni_cb.pack(side="left")
        self.src_yeni_count = ctk.CTkEntry(self.yeni_row, width=65, height=28)
        self.src_yeni_count.insert(0, "10")
        self.src_yeni_count.pack(side="right")
        self.yeni_lbl = ctk.CTkLabel(self.yeni_row, text="ədəd", font=ctk.CTkFont(size=11), text_color="#94A3B8")
        self.yeni_lbl.pack(side="right", padx=5)

        # 3. tap.az row
        self.tap_row = ctk.CTkFrame(self.source_box, fg_color="transparent")
        self.tap_row.pack(fill="x", padx=12, pady=(3, 10))
        self.src_tap_var = ctk.BooleanVar(value=True)
        self.src_tap_cb = ctk.CTkCheckBox(self.tap_row, text="tap.az", variable=self.src_tap_var, width=120)
        self.src_tap_cb.pack(side="left")
        self.src_tap_count = ctk.CTkEntry(self.tap_row, width=65, height=28)
        self.src_tap_count.insert(0, "5")
        self.src_tap_count.pack(side="right")
        self.tap_lbl = ctk.CTkLabel(self.tap_row, text="ədəd", font=ctk.CTkFont(size=11), text_color="#94A3B8")
        self.tap_lbl.pack(side="right", padx=5)

        # Rayon Dropdown
        self.region_label = ctk.CTkLabel(self.left_frame, text="📍 Rayon Seçimi:", font=ctk.CTkFont(size=13, weight="bold"))
        self.region_label.pack(anchor="w", padx=20, pady=(5, 4))

        regions_list = ["Bütün Bakı"] + list(HIERARCHY_MAP.keys())
        self.region_dropdown = ctk.CTkComboBox(
            self.left_frame,
            values=regions_list,
            height=36,
            command=self._on_region_changed
        )
        self.region_dropdown.set("Yasamal r.")
        self.region_dropdown.pack(fill="x", padx=15, pady=(0, 10))

        # Sub-Location Dropdown (Metro / Qəsəbə)
        self.subloc_label = ctk.CTkLabel(
            self.left_frame,
            text="🚇 Dəqiq Ərazi / Metro (Sub-kateqoriya):",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.subloc_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.subloc_dropdown = ctk.CTkComboBox(
            self.left_frame,
            values=["Bütün rayon üzrə"],
            height=36
        )
        self.subloc_dropdown.set("Bütün rayon üzrə")
        self.subloc_dropdown.pack(fill="x", padx=15, pady=(0, 14))
        self._update_sublocations_for_region("Yasamal r.")

        # Listing Type (Sale vs Rent)
        self.type_label = ctk.CTkLabel(self.left_frame, text="📋 Əməliyyat Növü:", font=ctk.CTkFont(size=13, weight="bold"))
        self.type_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.type_segment = ctk.CTkSegmentedButton(
            self.left_frame,
            values=["Alqı-satqı", "Kirayə"],
            selected_color="#059669",
            selected_hover_color="#047857"
        )
        self.type_segment.set("Alqı-satqı")
        self.type_segment.pack(fill="x", padx=15, pady=(0, 14))

        # Category Selector
        self.cat_label = ctk.CTkLabel(self.left_frame, text="🏠 Kateqoriya:", font=ctk.CTkFont(size=13, weight="bold"))
        self.cat_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.cat_dropdown = ctk.CTkComboBox(
            self.left_frame,
            values=list(CATEGORIES_MAP.keys()),
            height=36
        )
        self.cat_dropdown.set("Bütün Mənzillər")
        self.cat_dropdown.pack(fill="x", padx=15, pady=(0, 14))

        # Room Count Checkboxes
        self.room_label = ctk.CTkLabel(self.left_frame, text="🔢 Otaq Sayı:", font=ctk.CTkFont(size=13, weight="bold"))
        self.room_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.room_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.room_frame.pack(fill="x", padx=15, pady=(0, 14))

        self.room_vars = {}
        for r_num in ["1", "2", "3", "4", "5+"]:
            var = ctk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(self.room_frame, text=r_num, variable=var, width=45, checkbox_width=18, checkbox_height=18)
            cb.pack(side="left", padx=5)
            self.room_vars[r_num] = var

        # Price Range
        self.price_label = ctk.CTkLabel(self.left_frame, text="💰 Qiymət Aralığı (AZN):", font=ctk.CTkFont(size=13, weight="bold"))
        self.price_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.price_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.price_frame.pack(fill="x", padx=15, pady=(0, 14))

        self.price_min_entry = ctk.CTkEntry(self.price_frame, placeholder_text="Min AZN", width=170, height=34)
        self.price_min_entry.pack(side="left", padx=(0, 10))

        self.price_max_entry = ctk.CTkEntry(self.price_frame, placeholder_text="Max AZN", width=170, height=34)
        self.price_max_entry.pack(side="left")

        # Area Range
        self.area_label = ctk.CTkLabel(self.left_frame, text="📐 Sahə Aralığı (m²):", font=ctk.CTkFont(size=13, weight="bold"))
        self.area_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.area_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.area_frame.pack(fill="x", padx=15, pady=(0, 14))

        self.area_min_entry = ctk.CTkEntry(self.area_frame, placeholder_text="Min m²", width=170, height=34)
        self.area_min_entry.pack(side="left", padx=(0, 10))

        self.area_max_entry = ctk.CTkEntry(self.area_frame, placeholder_text="Max m²", width=170, height=34)
        self.area_max_entry.pack(side="left")

        # Checkboxes (Kupça, İpoteka, Təmir)
        self.attr_label = ctk.CTkLabel(self.left_frame, text="📄 Xüsusi Şərtlər:", font=ctk.CTkFont(size=13, weight="bold"))
        self.attr_label.pack(anchor="w", padx=20, pady=(5, 4))

        self.kupca_var = ctk.BooleanVar(value=False)
        self.kupca_cb = ctk.CTkCheckBox(self.left_frame, text="Yalnız Kupçalı (Çıxarışlı)", variable=self.kupca_var)
        self.kupca_cb.pack(anchor="w", padx=25, pady=3)

        self.ipoteka_var = ctk.BooleanVar(value=False)
        self.ipoteka_cb = ctk.CTkCheckBox(self.left_frame, text="İpotekaya Yararlı", variable=self.ipoteka_var)
        self.ipoteka_cb.pack(anchor="w", padx=25, pady=3)

        self.temir_var = ctk.BooleanVar(value=False)
        self.temir_cb = ctk.CTkCheckBox(self.left_frame, text="Yalnız Təmirli", variable=self.temir_var)
        self.temir_cb.pack(anchor="w", padx=25, pady=(3, 20))

        # ACTION BUTTONS
        self.start_btn = ctk.CTkButton(
            self.left_frame,
            text="🚀 AXTARIŞI BAŞLAT",
            height=46,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#10B981",
            hover_color="#059669",
            command=self.start_scraping
        )
        self.start_btn.pack(fill="x", padx=15, pady=(0, 8))

        self.stop_btn = ctk.CTkButton(
            self.left_frame,
            text="⏹️ DAYANDIR",
            height=36,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#EF4444",
            hover_color="#DC2626",
            state="disabled",
            command=self.stop_scraping
        )
        self.stop_btn.pack(fill="x", padx=15, pady=(0, 8))

        self.excel_btn = ctk.CTkButton(
            self.left_frame,
            text="📂 EXCEL-DƏ AÇ",
            height=38,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            state="disabled",
            command=self.open_excel_file
        )
        self.excel_btn.pack(fill="x", padx=15, pady=(0, 8))

        self.folder_btn = ctk.CTkButton(
            self.left_frame,
            text="📁 Bütün Fayllar Qovluğu",
            height=32,
            font=ctk.CTkFont(size=11),
            fg_color="#334155",
            hover_color="#475569",
            command=self.open_output_folder
        )
        self.folder_btn.pack(fill="x", padx=15, pady=(0, 25))

        # -------------------------------------------------------------
        # RIGHT MAIN PANEL: STATS & TABS (TABLE VIEWER + LIVE CONSOLE)
        # -------------------------------------------------------------
        self.right_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#0F172A")
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)

        # Top Status Bar
        self.top_bar = ctk.CTkFrame(self.right_frame, height=55, fg_color="#1E293B", corner_radius=0)
        self.top_bar.pack(fill="x", side="top")

        self.status_title = ctk.CTkLabel(
            self.top_bar,
            text="Sistem Hazırdır. Parametrləri seçib 'Axtarışı Başlat' düyməsinə klikləyin.",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#E2E8F0"
        )
        self.status_title.pack(side="left", padx=20, pady=15)

        # KPI Stats Grid (4 Cards)
        self.kpi_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.kpi_frame.pack(fill="x", padx=20, pady=(16, 8))
        for i in range(4):
            self.kpi_frame.grid_columnconfigure(i, weight=1)

        self.card_scanned = self._create_kpi_card(self.kpi_frame, 0, "Yoxlanılan Elanlar", "0", "#38BDF8")
        self.card_accepted = self._create_kpi_card(self.kpi_frame, 1, "Təsdiqlənmiş Lid", "0", "#4ADE80")
        self.card_skipped = self._create_kpi_card(self.kpi_frame, 2, "Kənarlaşdırılan", "0", "#F87171")
        self.card_dups = self._create_kpi_card(self.kpi_frame, 3, "Təkrar Nömrə", "0", "#FBBF24")

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.right_frame, height=10, progress_color="#10B981")
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=20, pady=(6, 12))

        # Tabview: 1. Daxili Baza Cədvəli (Table) | 2. Canlı Konsol (Log)
        self.tabview = ctk.CTkTabview(self.right_frame, corner_radius=10, fg_color="#1E293B")
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.tab_table = self.tabview.add("📊 DAXİLİ BAZA CƏDVƏLİ")
        self.tab_log = self.tabview.add("📝 CANLI KONSOL LOQU")

        # -------------------------------------------------------------
        # TAB 1: IN-APP DATABASE TABLE (Treeview)
        # -------------------------------------------------------------
        self.table_ctrl_frame = ctk.CTkFrame(self.tab_table, fg_color="transparent", height=40)
        self.table_ctrl_frame.pack(fill="x", padx=10, pady=(5, 8))

        self.table_info_lbl = ctk.CTkLabel(
            self.table_ctrl_frame,
            text="Bazadakı elanlar (Birbaşa brauzerdə açmaq üçün sətirə iki dəfə klikləyin):",
            font=ctk.CTkFont(size=12),
            text_color="#94A3B8"
        )
        self.table_info_lbl.pack(side="left")

        self.open_selected_btn = ctk.CTkButton(
            self.table_ctrl_frame,
            text="🔗 Seçilmiş Elana Bax",
            height=30,
            width=160,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#0284C7",
            hover_color="#0369A1",
            command=self._open_selected_table_lead
        )
        self.open_selected_btn.pack(side="right")

        # Treeview styling & setup
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background="#090D16",
                        foreground="#F8FAFC",
                        fieldbackground="#090D16",
                        rowheight=28,
                        font=("Segoe UI", 10))
        style.configure("Treeview.Heading",
                        background="#0F172A",
                        foreground="#38BDF8",
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[('selected', '#2563EB')])

        tree_frame = ctk.CTkFrame(self.tab_table, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("no", "source", "name", "phone", "status", "price", "rooms", "loc", "url")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

        self.tree.heading("no", text="№")
        self.tree.heading("source", text="Mənbə")
        self.tree.heading("name", text="Ad / Şəxs")
        self.tree.heading("phone", text="Telefon")
        self.tree.heading("status", text="Status")
        self.tree.heading("price", text="Qiymət")
        self.tree.heading("rooms", text="Otaq")
        self.tree.heading("loc", text="Yerləşmə")
        self.tree.heading("url", text="Elan Linki")

        self.tree.column("no", width=40, anchor="center")
        self.tree.column("source", width=95, anchor="center")
        self.tree.column("name", width=140, anchor="w")
        self.tree.column("phone", width=130, anchor="center")
        self.tree.column("status", width=110, anchor="center")
        self.tree.column("price", width=105, anchor="e")
        self.tree.column("rooms", width=60, anchor="center")
        self.tree.column("loc", width=180, anchor="w")
        self.tree.column("url", width=220, anchor="w")

        tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", lambda event: self._open_selected_table_lead())

        # -------------------------------------------------------------
        # TAB 2: LIVE CONSOLE LOG
        # -------------------------------------------------------------
        self.log_textbox = ctk.CTkTextbox(
            self.tab_log,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color="#090D16",
            text_color="#E2E8F0",
            wrap="word"
        )
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=10)

    def _create_kpi_card(self, parent, col, title, value, color):
        card = ctk.CTkFrame(parent, fg_color="#1E293B", corner_radius=10)
        card.grid(row=0, column=col, padx=6, sticky="ew")

        lbl_val = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=24, weight="bold"), text_color=color)
        lbl_val.pack(pady=(10, 0))

        lbl_title = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=11), text_color="#94A3B8")
        lbl_title.pack(pady=(0, 10))
        return lbl_val

    def _on_region_changed(self, selected_region):
        self._update_sublocations_for_region(selected_region)

    def _update_sublocations_for_region(self, region_name):
        options = ["Bütün rayon üzrə"]
        if region_name in HIERARCHY_MAP:
            sublocs = list(HIERARCHY_MAP[region_name].get("sublocations", {}).keys())
            if sublocs:
                options.extend(sublocs)
        self.subloc_dropdown.configure(values=options)
        self.subloc_dropdown.set("Bütün rayon üzrə")

    def log(self, text: str):
        self.log_textbox.insert("end", text + "\n")
        self.log_textbox.see("end")

    def _open_selected_table_lead(self):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            url = item["values"][8]
            if url:
                webbrowser.open(url)
        else:
            messagebox.showinfo("Məlumat", "Zəhmət olmasa cədvəldən bir elan seçin.")

    def _check_trial_timer(self):
        is_active, text, rem = license_manager.check_status()
        self.license_title.configure(text=text)
        if not is_active:
            self.license_title.configure(text_color="#EF4444")
            self.status_title.configure(text="Sınaq müddəti bitmişdir. Lisenziya kodunu daxil edin.", text_color="#EF4444")
        else:
            self.license_title.configure(text_color="#4ADE80")

        self.after(60000, self._check_trial_timer)

    def _copy_hwid_to_clipboard(self):
        self.clipboard_clear()
        self.clipboard_append(self.hwid_str)
        self.update()
        messagebox.showinfo(
            "Kopyalandı",
            f"Lisenziya ID-niz kopyalandı:\n\n{self.hwid_str}\n\n"
            "Bu kodu WhatsApp və ya Telegram ilə adminə göndərib aktivasiya açarınızı əldə edin."
        )

    def _open_activation_dialog(self):
        dialog = ctk.CTkInputDialog(
            text=f"Sizin Lisenziya ID: {self.hwid_str}\n(Bu kodu adminə göndərib açar alın)\n\nAktivasiya Açarını daxil edin:",
            title="BinaLeadPro Aktivasiya"
        )
        key = dialog.get_input()
        if key:
            success, msg = license_manager.activate_license(key)
            if success:
                messagebox.showinfo("Uğurlu Aktivasiya", msg)
                self._check_trial_timer()
            else:
                messagebox.showerror("Aktivasiya Xətası", msg)

    def start_scraping(self):
        is_active, text, _ = license_manager.check_status()
        if not is_active:
            messagebox.showwarning(
                "Sınaq Müddəti Bitdi",
                "Sınaq müddətiniz başa çatıb!\n\n"
                "Proqramı aktivləşdirmək üçün lisenziya kodunu daxil edin."
            )
            return

        target_mode = "FSBO" if "Mülkiyyətçi" in self.target_segment.get() else "AGENT"
        region = self.region_dropdown.get()
        subloc = self.subloc_dropdown.get()
        category = self.cat_dropdown.get()
        listing_type = self.type_segment.get()

        # Parse per-source target counts
        bina_cnt = 0
        if self.src_bina_var.get():
            try:
                bina_cnt = max(0, int(self.src_bina_count.get().strip()))
            except ValueError:
                bina_cnt = 20

        yeni_cnt = 0
        if self.src_yeni_var.get():
            try:
                yeni_cnt = max(0, int(self.src_yeni_count.get().strip()))
            except ValueError:
                yeni_cnt = 10

        tap_cnt = 0
        if self.src_tap_var.get():
            try:
                tap_cnt = max(0, int(self.src_tap_count.get().strip()))
            except ValueError:
                tap_cnt = 5

        total_target = bina_cnt + yeni_cnt + tap_cnt
        if total_target <= 0:
            messagebox.showwarning("Xəta", "Zəhmət olmasa ən azı bir mənbə və hədəf sayı seçin.")
            return

        selected_rooms = [r for r, v in self.room_vars.items() if v.get()]
        
        price_min = None
        try:
            if self.price_min_entry.get().strip():
                price_min = int(self.price_min_entry.get().strip())
        except ValueError:
            pass

        price_max = None
        try:
            if self.price_max_entry.get().strip():
                price_max = int(self.price_max_entry.get().strip())
        except ValueError:
            pass

        area_min = None
        try:
            if self.area_min_entry.get().strip():
                area_min = float(self.area_min_entry.get().strip())
        except ValueError:
            pass

        area_max = None
        try:
            if self.area_max_entry.get().strip():
                area_max = float(self.area_max_entry.get().strip())
        except ValueError:
            pass

        has_bill_of_sale = self.kupca_var.get()
        has_mortgage = self.ipoteka_var.get()
        has_repair = self.temir_var.get()

        # Reset UI
        self.all_leads = []
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.log_textbox.delete("1.0", "end")
        self.progress_bar.set(0)
        self.card_scanned.configure(text="0")
        self.card_accepted.configure(text="0")
        self.card_skipped.configure(text="0")
        self.card_dups.configure(text="0")
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.excel_btn.configure(state="disabled")
        self.status_title.configure(text=f"⏳ Axtarış aktivdir... Toplam {total_target} ədəd lid toplanır.", text_color="#38BDF8")

        def engine_callback(event_type: str, data: Any):
            self.after(0, lambda: self._handle_engine_event(event_type, data))

        self.engine = BinaScraperEngine(callback=engine_callback)

        def run_thread():
            self.engine.run(
                target_type=target_mode,
                region_name=region,
                sublocation_name=subloc,
                category_name=category,
                listing_type=listing_type,
                room_ids=selected_rooms,
                price_min=price_min,
                price_max=price_max,
                area_min=area_min,
                area_max=area_max,
                has_bill_of_sale=has_bill_of_sale,
                has_mortgage=has_mortgage,
                has_repair=has_repair,
                bina_target=bina_cnt,
                yeni_target=yeni_cnt,
                tap_target=tap_cnt,
                output_dir=BASE_DIR
            )

        self.worker_thread = threading.Thread(target=run_thread, daemon=True)
        self.worker_thread.start()

    def stop_scraping(self):
        if self.engine:
            self.engine.cancel()
            self.stop_btn.configure(state="disabled")
            self.status_title.configure(text="⏹️ Axtarış dayandırıldı.", text_color="#F87171")

    def _handle_engine_event(self, event_type: str, data: Any):
        if event_type == "log":
            self.log(str(data))
        elif event_type == "stats":
            self.card_scanned.configure(text=str(data.get("scanned", 0)))
            self.card_accepted.configure(text=str(data.get("accepted", 0)))
            self.card_skipped.configure(text=str(data.get("skipped", 0)))
            self.card_dups.configure(text=str(data.get("duplicates", 0)))
        elif event_type == "lead":
            self.all_leads.append(data)
            try:
                p_val = float(data.get('Qiymət (AZN)', 0) or 0)
                price_str = f"{int(p_val):,}".replace(',', ' ') + " AZN" if p_val > 0 else "-"
            except Exception:
                price_str = str(data.get('Qiymət (AZN)', "-"))
            self.tree.insert("", "end", values=(
                data.get("№", len(self.all_leads)),
                data.get("Mənbə", "bina.az"),
                data.get("Ad / Mülkiyyətçi", "-"),
                data.get("Əlaqə Nömrəsi", "-"),
                data.get("Status", "-"),
                price_str,
                data.get("Otaq", "-"),
                data.get("Yerləşmə", "-"),
                data.get("Elan Keçidi", "")
            ))
            # Auto-scroll table to bottom
            children = self.tree.get_children()
            if children:
                self.tree.see(children[-1])
        elif event_type == "progress":
            self.progress_bar.set(float(data))
        elif event_type == "finished":
            self.last_excel_path = data.get("excel_path")
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            self.excel_btn.configure(state="normal")
            self.status_title.configure(text=f"✅ Uğurla tamamlandı! {data.get('count')} ədəd lid bazaya oturdu.", text_color="#4ADE80")
            messagebox.showinfo(
                "Tamamlandı",
                f"{data.get('count')} ədəd əlaqə nömrəsi və məlumat daxili bazaya oturdu!\n\n"
                "• Birbaşa proqramın içindəki cədvəldən baxa bilərsiniz.\n"
                "• Elana baxmaq üçün sətirə iki dəfə klikləyin.\n"
                "• İstəsəniz 'Excel-də Aç' düyməsi ilə cədvəli aça bilərsiniz."
            )

    def open_excel_file(self):
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            os.startfile(self.last_excel_path)
        else:
            messagebox.showwarning("Fayl Tapılmadı", "Excel faylı tapılmadı.")

    def open_output_folder(self):
        os.startfile(BASE_DIR)


if __name__ == "__main__":
    app = BinaLeadProApp()
    app.mainloop()
