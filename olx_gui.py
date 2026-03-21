"""
GUI do scrapowania ofert wynajmu mieszkan w Gdansku.
Nowoczesny interfejs oparty na CustomTkinter.
"""

import os
import re
import subprocess
import sys
import threading
import time
import urllib.request

import customtkinter as ctk
from tkinter import filedialog

# Import scrapera
sys.path.insert(0, os.path.dirname(__file__))
from olx_scraper import (
    collect_olx_links, collect_otodom_links, collect_facebook_links,
    extract_listing_data, extract_facebook_data_from_google,
    save_csv, save_xlsx, dismiss_cookies, parse_price_value,
    IMAGES_DIR,
)
from playwright.sync_api import sync_playwright

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CDP_PORT = 9222

# ── Konfiguracja wygladu ──────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Kolory
ACCENT = "#6C63FF"
ACCENT_HOVER = "#5A52D5"
GREEN = "#4CAF50"
RED = "#EF5350"
YELLOW = "#FFB74D"
CARD_BG = "#1E1E2E"
SURFACE = "#2A2A3C"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Skaner Mieszkan - Gdansk")
        self.geometry("960x720")
        self.minsize(800, 600)

        # Zmienne filtrow
        self.var_min_area = ctk.IntVar(value=30)
        self.var_max_price = ctk.IntVar(value=3500)
        self.var_pets = ctk.BooleanVar(value=False)
        self.var_balcony = ctk.BooleanVar(value=False)
        self.var_gym = ctk.BooleanVar(value=False)
        self.var_gated = ctk.BooleanVar(value=False)
        self.var_wifi = ctk.BooleanVar(value=False)
        self.var_olx = ctk.BooleanVar(value=True)
        self.var_otodom = ctk.BooleanVar(value=True)
        self.var_facebook = ctk.BooleanVar(value=True)
        self.var_olx_pages = ctk.IntVar(value=5)
        self.var_otodom_pages = ctk.IntVar(value=3)
        self.var_fb_pages = ctk.IntVar(value=2)
        self.var_output_dir = ctk.StringVar(value=os.path.expanduser("~/Desktop"))

        self.results = []
        self.csv_path = ""
        self.xlsx_path = ""

        self.show_start_panel()

    # ══════════════════════════════════════════════════════════════════════════
    #  PANEL STARTOWY
    # ══════════════════════════════════════════════════════════════════════════
    def show_start_panel(self):
        self.clear_window()

        # Kontener glowny z scrollem
        container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # ── Header ──
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(header, text="Skaner Mieszkan",
                     font=ctk.CTkFont(size=32, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(header, text="Przeszukaj OLX i Otodom w poszukiwaniu idealnego mieszkania w Gdansku",
                     font=ctk.CTkFont(size=14),
                     text_color=("gray50", "gray70")).pack(anchor="w", pady=(4, 0))

        # ── Sekcja: Platformy ──
        self.card_label(container, "Platformy")
        plat_card = ctk.CTkFrame(container, corner_radius=12)
        plat_card.pack(fill="x", pady=(0, 16))

        plat_grid = ctk.CTkFrame(plat_card, fg_color="transparent")
        plat_grid.pack(fill="x", padx=20, pady=16)
        plat_grid.grid_columnconfigure((0, 1, 2), weight=1)

        # OLX
        olx_frame = ctk.CTkFrame(plat_grid, corner_radius=10)
        olx_frame.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        olx_inner = ctk.CTkFrame(olx_frame, fg_color="transparent")
        olx_inner.pack(padx=16, pady=12)

        ctk.CTkSwitch(olx_inner, text="OLX.pl", variable=self.var_olx,
                       font=ctk.CTkFont(size=14, weight="bold"),
                       progress_color=ACCENT).pack(anchor="w")
        pages_row = ctk.CTkFrame(olx_inner, fg_color="transparent")
        pages_row.pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(pages_row, text="Stron do przeszukania:",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray50", "gray70")).pack(side="left")
        ctk.CTkOptionMenu(pages_row, values=[str(i) for i in range(1, 11)],
                          variable=self.var_olx_pages, width=60,
                          fg_color=ACCENT, button_color=ACCENT_HOVER,
                          command=lambda v: self.var_olx_pages.set(int(v))
                          ).pack(side="left", padx=(8, 0))

        # Otodom
        oto_frame = ctk.CTkFrame(plat_grid, corner_radius=10)
        oto_frame.grid(row=0, column=1, padx=(8, 0), sticky="nsew")
        oto_inner = ctk.CTkFrame(oto_frame, fg_color="transparent")
        oto_inner.pack(padx=16, pady=12)

        ctk.CTkSwitch(oto_inner, text="Otodom.pl", variable=self.var_otodom,
                       font=ctk.CTkFont(size=14, weight="bold"),
                       progress_color=ACCENT).pack(anchor="w")
        pages_row2 = ctk.CTkFrame(oto_inner, fg_color="transparent")
        pages_row2.pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(pages_row2, text="Stron do przeszukania:",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray50", "gray70")).pack(side="left")
        ctk.CTkOptionMenu(pages_row2, values=[str(i) for i in range(1, 11)],
                          variable=self.var_otodom_pages, width=60,
                          fg_color=ACCENT, button_color=ACCENT_HOVER,
                          command=lambda v: self.var_otodom_pages.set(int(v))
                          ).pack(side="left", padx=(8, 0))

        # Facebook
        fb_frame = ctk.CTkFrame(plat_grid, corner_radius=10)
        fb_frame.grid(row=0, column=2, padx=(8, 0), sticky="nsew")
        fb_inner = ctk.CTkFrame(fb_frame, fg_color="transparent")
        fb_inner.pack(padx=16, pady=12)

        ctk.CTkSwitch(fb_inner, text="Facebook", variable=self.var_facebook,
                       font=ctk.CTkFont(size=14, weight="bold"),
                       progress_color="#1877F2").pack(anchor="w")
        fb_note = ctk.CTkLabel(fb_inner, text="via Google (bez logowania)",
                               font=ctk.CTkFont(size=10),
                               text_color=("gray50", "gray60"))
        fb_note.pack(anchor="w", pady=(2, 0))
        pages_row3 = ctk.CTkFrame(fb_inner, fg_color="transparent")
        pages_row3.pack(anchor="w", pady=(6, 0))
        ctk.CTkLabel(pages_row3, text="Stron Google:",
                     font=ctk.CTkFont(size=12),
                     text_color=("gray50", "gray70")).pack(side="left")
        ctk.CTkOptionMenu(pages_row3, values=[str(i) for i in range(1, 6)],
                          variable=self.var_fb_pages, width=60,
                          fg_color="#1877F2", button_color="#1565C0",
                          command=lambda v: self.var_fb_pages.set(int(v))
                          ).pack(side="left", padx=(8, 0))

        # ── Sekcja: Kryteria ──
        self.card_label(container, "Kryteria wyszukiwania")
        crit_card = ctk.CTkFrame(container, corner_radius=12)
        crit_card.pack(fill="x", pady=(0, 16))
        crit_inner = ctk.CTkFrame(crit_card, fg_color="transparent")
        crit_inner.pack(fill="x", padx=20, pady=16)

        # Powierzchnia
        area_row = ctk.CTkFrame(crit_inner, fg_color="transparent")
        area_row.pack(fill="x", pady=(0, 12))
        self.area_label = ctk.CTkLabel(area_row,
                                        text=f"Min. powierzchnia: {self.var_min_area.get()} m2",
                                        font=ctk.CTkFont(size=13))
        self.area_label.pack(anchor="w")
        area_slider = ctk.CTkSlider(area_row, from_=15, to=100,
                                     variable=self.var_min_area,
                                     number_of_steps=85,
                                     progress_color=ACCENT, button_color=ACCENT,
                                     button_hover_color=ACCENT_HOVER,
                                     command=lambda v: self.area_label.configure(
                                         text=f"Min. powierzchnia: {int(v)} m2"))
        area_slider.pack(fill="x", pady=(4, 0))

        # Cena
        price_row = ctk.CTkFrame(crit_inner, fg_color="transparent")
        price_row.pack(fill="x", pady=(0, 8))
        self.price_label = ctk.CTkLabel(price_row,
                                         text=f"Max. cena calkowita: {self.var_max_price.get()} zl",
                                         font=ctk.CTkFont(size=13))
        self.price_label.pack(anchor="w")
        price_slider = ctk.CTkSlider(price_row, from_=1000, to=8000,
                                      variable=self.var_max_price,
                                      number_of_steps=70,
                                      progress_color=ACCENT, button_color=ACCENT,
                                      button_hover_color=ACCENT_HOVER,
                                      command=lambda v: self.price_label.configure(
                                          text=f"Max. cena calkowita: {int(v)} zl"))
        price_slider.pack(fill="x", pady=(4, 0))

        # ── Sekcja: Udogodnienia ──
        self.card_label(container, "Wymagane udogodnienia")
        feat_card = ctk.CTkFrame(container, corner_radius=12)
        feat_card.pack(fill="x", pady=(0, 16))
        feat_inner = ctk.CTkFrame(feat_card, fg_color="transparent")
        feat_inner.pack(fill="x", padx=20, pady=16)

        feat_grid = ctk.CTkFrame(feat_inner, fg_color="transparent")
        feat_grid.pack(fill="x")
        feat_grid.grid_columnconfigure((0, 1, 2), weight=1)

        features = [
            ("Zwierzeta akceptowane", self.var_pets, 0, 0),
            ("Balkon / taras", self.var_balcony, 0, 1),
            ("Silownia / fitness", self.var_gym, 0, 2),
            ("Osiedle strzezone", self.var_gated, 1, 0),
            ("WiFi / internet", self.var_wifi, 1, 1),
        ]
        for text, var, row, col in features:
            ctk.CTkCheckBox(feat_grid, text=text, variable=var,
                            font=ctk.CTkFont(size=13),
                            fg_color=ACCENT, hover_color=ACCENT_HOVER,
                            corner_radius=6).grid(row=row, column=col,
                                                   padx=8, pady=8, sticky="w")

        # ── Sekcja: Zapis ──
        self.card_label(container, "Folder zapisu wynikow")
        dir_card = ctk.CTkFrame(container, corner_radius=12)
        dir_card.pack(fill="x", pady=(0, 20))
        dir_inner = ctk.CTkFrame(dir_card, fg_color="transparent")
        dir_inner.pack(fill="x", padx=20, pady=16)

        dir_row = ctk.CTkFrame(dir_inner, fg_color="transparent")
        dir_row.pack(fill="x")

        ctk.CTkEntry(dir_row, textvariable=self.var_output_dir,
                     font=ctk.CTkFont(size=12), height=36).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(dir_row, text="Zmien...", width=80, height=36,
                      fg_color="transparent", border_width=1,
                      border_color=ACCENT, text_color=ACCENT,
                      hover_color=("gray90", "gray25"),
                      command=self.pick_dir).pack(side="left", padx=(10, 0))

        # ── Przycisk START ──
        ctk.CTkButton(container, text="Zacznij skanowanie",
                      font=ctk.CTkFont(size=16, weight="bold"),
                      height=50, corner_radius=12,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=self.start_scraping).pack(fill="x", pady=(10, 0))

    # ══════════════════════════════════════════════════════════════════════════
    #  PANEL POSTEPU
    # ══════════════════════════════════════════════════════════════════════════
    def show_progress_panel(self):
        self.clear_window()

        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=30, pady=30)

        # Header
        ctk.CTkLabel(container, text="Skanowanie w toku",
                     font=ctk.CTkFont(size=28, weight="bold")).pack(pady=(0, 5))

        self.lbl_status = ctk.CTkLabel(container, text="Uruchamiam Chrome...",
                                        font=ctk.CTkFont(size=14),
                                        text_color=("gray50", "gray70"))
        self.lbl_status.pack(pady=(0, 20))

        # Progress bar
        self.progress = ctk.CTkProgressBar(container, width=600, height=8,
                                            progress_color=ACCENT, corner_radius=4)
        self.progress.pack(pady=(0, 8))
        self.progress.set(0)

        self.lbl_counter = ctk.CTkLabel(container, text="0 / 0",
                                         font=ctk.CTkFont(size=13),
                                         text_color=("gray50", "gray70"))
        self.lbl_counter.pack(pady=(0, 16))

        # Statystyki w czasie rzeczywistym
        stats_frame = ctk.CTkFrame(container, corner_radius=12)
        stats_frame.pack(fill="x", pady=(0, 16))
        stats_inner = ctk.CTkFrame(stats_frame, fg_color="transparent")
        stats_inner.pack(fill="x", padx=20, pady=12)
        stats_inner.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.stat_labels = {}
        stat_items = [
            ("Znalezione", "0", ACCENT),
            ("W budzecie", "0", GREEN),
            ("Poza budzetem", "0", RED),
            ("Ze zwierzetami", "0", YELLOW),
        ]
        for i, (label, val, color) in enumerate(stat_items):
            frame = ctk.CTkFrame(stats_inner, fg_color="transparent")
            frame.grid(row=0, column=i, padx=8)
            vlbl = ctk.CTkLabel(frame, text=val, font=ctk.CTkFont(size=24, weight="bold"),
                                text_color=color)
            vlbl.pack()
            ctk.CTkLabel(frame, text=label, font=ctk.CTkFont(size=11),
                         text_color=("gray50", "gray70")).pack()
            self.stat_labels[label] = vlbl

        # Log
        self.log_text = ctk.CTkTextbox(container, font=ctk.CTkFont(family="Consolas", size=11),
                                        corner_radius=10, height=300)
        self.log_text.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════════
    #  PANEL WYNIKOW
    # ══════════════════════════════════════════════════════════════════════════
    def show_results_panel(self):
        self.clear_window()

        container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(header, text="Wyniki skanowania",
                     font=ctk.CTkFont(size=28, weight="bold")).pack(anchor="w")

        # ── Podsumowanie - karty ──
        summary_frame = ctk.CTkFrame(container, fg_color="transparent")
        summary_frame.pack(fill="x", pady=(0, 20))
        summary_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        total = len(self.results)
        olx_c = sum(1 for r in self.results if r.get("Platforma") == "OLX")
        oto_c = sum(1 for r in self.results if r.get("Platforma") == "Otodom")
        fb_c = sum(1 for r in self.results if r.get("Platforma") == "Facebook")
        budget_c = sum(1 for r in self.results if "POZA BUDZETEM" not in r.get("Uwagi", ""))
        pets_c = sum(1 for r in self.results if r.get("Zwierzeta") == "Tak")

        cards_data = [
            ("Wszystkie", str(total), ACCENT),
            ("OLX", str(olx_c), GREEN),
            ("Otodom", str(oto_c), "#2196F3"),
            ("Facebook", str(fb_c), "#1877F2"),
            ("W budzecie", str(budget_c), GREEN),
            ("Ze zwierzetami", str(pets_c), YELLOW),
        ]
        for i, (label, val, color) in enumerate(cards_data):
            card = ctk.CTkFrame(summary_frame, corner_radius=12)
            card.grid(row=0, column=i, padx=6, sticky="nsew")
            card_inner = ctk.CTkFrame(card, fg_color="transparent")
            card_inner.pack(padx=16, pady=14)
            ctk.CTkLabel(card_inner, text=val,
                         font=ctk.CTkFont(size=28, weight="bold"),
                         text_color=color).pack()
            ctk.CTkLabel(card_inner, text=label,
                         font=ctk.CTkFont(size=11),
                         text_color=("gray50", "gray70")).pack()

        # ── TOP 10 ──
        self.card_label(container, "TOP 10 ofert wg Twoich kryteriow")

        top10 = self.get_top_n(10)

        for i, r in enumerate(top10, 1):
            self.render_listing_card(container, i, r)

        # ── Pliki ──
        self.card_label(container, "Zapisane pliki")
        files_card = ctk.CTkFrame(container, corner_radius=12)
        files_card.pack(fill="x", pady=(0, 16))
        files_inner = ctk.CTkFrame(files_card, fg_color="transparent")
        files_inner.pack(fill="x", padx=20, pady=14)

        for f in [self.csv_path, self.xlsx_path]:
            row = ctk.CTkFrame(files_inner, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=f, font=ctk.CTkFont(family="Consolas", size=11),
                         text_color=ACCENT).pack(side="left")

        # Przyciski
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkButton(btn_frame, text="Otworz XLSX",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      height=44, corner_radius=10,
                      fg_color=ACCENT, hover_color=ACCENT_HOVER,
                      command=lambda: os.startfile(self.xlsx_path)).pack(side="left", padx=(0, 10))

        ctk.CTkButton(btn_frame, text="Otworz folder",
                      font=ctk.CTkFont(size=14),
                      height=44, corner_radius=10,
                      fg_color="transparent", border_width=1,
                      border_color=ACCENT, text_color=ACCENT,
                      hover_color=("gray90", "gray25"),
                      command=lambda: os.startfile(self.var_output_dir.get())).pack(side="left", padx=(0, 10))

        ctk.CTkButton(btn_frame, text="Skanuj ponownie",
                      font=ctk.CTkFont(size=14),
                      height=44, corner_radius=10,
                      fg_color="transparent", border_width=1,
                      border_color="gray50", text_color="gray70",
                      hover_color=("gray90", "gray25"),
                      command=self.show_start_panel).pack(side="right")

    def render_listing_card(self, parent, rank, r):
        """Renderuje karte pojedynczej oferty w TOP 10."""
        card = ctk.CTkFrame(parent, corner_radius=12)
        card.pack(fill="x", pady=(0, 8))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=12)
        inner.grid_columnconfigure(1, weight=1)

        # Rank badge
        rank_color = ACCENT if rank <= 3 else ("gray50", "gray60")
        rank_label = ctk.CTkLabel(inner, text=f"#{rank}",
                                   font=ctk.CTkFont(size=18, weight="bold"),
                                   text_color=rank_color, width=40)
        rank_label.grid(row=0, column=0, rowspan=2, padx=(0, 12))

        # Nazwa + platforma
        name = r.get("Nazwa", "?")[:55]
        platform = r.get("Platforma", "?")
        plat_color = GREEN if platform == "OLX" else "#1877F2" if platform == "Facebook" else "#2196F3"

        name_frame = ctk.CTkFrame(inner, fg_color="transparent")
        name_frame.grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(name_frame, text=name,
                     font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        ctk.CTkLabel(name_frame, text=f"  {platform}",
                     font=ctk.CTkFont(size=11),
                     text_color=plat_color).pack(side="left")

        # Detale
        details_frame = ctk.CTkFrame(inner, fg_color="transparent")
        details_frame.grid(row=1, column=1, sticky="w", pady=(4, 0))

        price = r.get("Cena_Suma", "?")
        area = r.get("Powierzchnia", "?")
        pets = r.get("Zwierzeta", "?")
        balcony = r.get("Balkon", "?")
        date = r.get("Data_Wystawienia", "?")

        tags = [
            (f"{price}", ACCENT),
            (f"{area}", ("gray50", "gray70")),
        ]
        if pets == "Tak":
            tags.append(("Zwierzeta OK", GREEN))
        elif pets == "Nie":
            tags.append(("Bez zwierzat", RED))
        if balcony == "Tak":
            tags.append(("Balkon", YELLOW))
        if date and date != "Brak danych" and date != "?":
            tags.append((date[:20], ("gray50", "gray70")))

        for text, color in tags:
            ctk.CTkLabel(details_frame, text=text,
                         font=ctk.CTkFont(size=12),
                         text_color=color,
                         padx=4).pack(side="left", padx=(0, 12))

        # Przycisk "Otworz"
        url = r.get("URL", "")
        if url:
            ctk.CTkButton(inner, text="Otworz", width=70, height=30,
                          corner_radius=8, font=ctk.CTkFont(size=11),
                          fg_color="transparent", border_width=1,
                          border_color=ACCENT, text_color=ACCENT,
                          hover_color=("gray90", "gray25"),
                          command=lambda u=url: os.startfile(u)
                          ).grid(row=0, column=2, rowspan=2, padx=(12, 0))

    # ══════════════════════════════════════════════════════════════════════════
    #  LOGIKA SKANOWANIA
    # ══════════════════════════════════════════════════════════════════════════
    def start_scraping(self):
        if not self.var_olx.get() and not self.var_otodom.get() and not self.var_facebook.get():
            dialog = ctk.CTkInputDialog(text="Wybierz przynajmniej jedna platforme!", title="Uwaga")
            return

        self.show_progress_panel()
        thread = threading.Thread(target=self.scrape_worker, daemon=True)
        thread.start()

    def scrape_worker(self):
        try:
            self.log("Zamykam Chrome...")
            os.system("taskkill /F /IM chrome.exe >nul 2>&1")
            time.sleep(2)

            self.log("Uruchamiam Chrome...")
            subprocess.Popen([
                CHROME_PATH,
                f"--remote-debugging-port={CDP_PORT}",
                "--user-data-dir=C:\\Users\\gabri\\chrome-olx-debug",
                "--no-first-run", "--no-default-browser-check",
                "about:blank",
            ])
            time.sleep(5)

            try:
                urllib.request.urlopen(f"http://127.0.0.1:{CDP_PORT}/json/version", timeout=5)
                self.log("Chrome uruchomiony!")
            except Exception:
                self.log("[BLAD] Chrome nie nasluchuje. Zamknij Chrome i sprobuj ponownie.")
                return

            with sync_playwright() as p:
                browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
                context = browser.contexts[0] if browser.contexts else browser.new_context()
                page = context.pages[0] if context.pages else context.new_page()
                time.sleep(2)

                all_links = []

                if self.var_olx.get():
                    self.update_status("Zbieram linki z OLX...")
                    olx_links = collect_olx_links(page, max_pages=self.var_olx_pages.get())
                    self.log(f"OLX: {len(olx_links)} ofert")
                    all_links += [(link, "OLX") for link in olx_links]

                if self.var_otodom.get():
                    self.update_status("Zbieram linki z Otodom...")
                    otodom_links = collect_otodom_links(page, max_pages=self.var_otodom_pages.get())
                    self.log(f"Otodom: {len(otodom_links)} ofert")
                    all_links += [(link, "Otodom") for link in otodom_links]

                fb_links = []
                if self.var_facebook.get():
                    self.update_status("Szukam ofert z Facebooka (via Google)...")
                    fb_links = collect_facebook_links(page, max_pages=self.var_fb_pages.get())
                    self.log(f"Facebook: {len(fb_links)} ofert")

                total = len(all_links) + len(fb_links)
                self.log(f"\nLacznie: {total} ofert\n")
                self.results = []

                budget_ok = 0
                budget_over = 0
                pets_count = 0
                counter = 0

                # OLX + Otodom
                for i, (link, platform) in enumerate(all_links, 1):
                    counter += 1
                    self.update_status(f"[{platform}] {counter}/{total}")
                    self.update_progress(counter, total)

                    data = extract_listing_data(page, link, counter, platform)
                    self.results.append(data)

                    if "POZA BUDZETEM" in data.get("Uwagi", ""):
                        budget_over += 1
                    else:
                        budget_ok += 1
                    if data.get("Zwierzeta") == "Tak":
                        pets_count += 1

                    self.update_live_stats(counter, budget_ok, budget_over, pets_count)

                    name = data["Nazwa"][:35]
                    price = data["Cena_Suma"]
                    self.log(f"[{counter}/{total}] [{platform}] {name} | {price}")

                # Facebook
                for i, link in enumerate(fb_links, 1):
                    counter += 1
                    self.update_status(f"[Facebook] {counter}/{total}")
                    self.update_progress(counter, total)

                    data = extract_facebook_data_from_google(page, link, counter)
                    self.results.append(data)

                    if "POZA BUDZETEM" in data.get("Uwagi", ""):
                        budget_over += 1
                    else:
                        budget_ok += 1
                    if data.get("Zwierzeta") == "Tak":
                        pets_count += 1

                    self.update_live_stats(counter, budget_ok, budget_over, pets_count)

                    name = data["Nazwa"][:35]
                    price = data["Cena_Suma"]
                    self.log(f"[{counter}/{total}] [Facebook] {name} | {price}")

                self.log("\nZapisuje pliki...")
                out_dir = self.var_output_dir.get()
                self.csv_path = os.path.join(out_dir, "mieszkania_gdansk.csv")
                self.xlsx_path = os.path.join(out_dir, "mieszkania_gdansk.xlsx")

                save_csv(self.results, self.csv_path)
                save_xlsx(self.results, self.xlsx_path)

                self.log(f"\nGotowe! {len(self.results)} ofert.")
                self.update_status("Zakonczone!")

                self.after(2000, self.show_results_panel)

        except Exception as e:
            self.log(f"\n[BLAD] {e}")
            self.update_status(f"Blad: {e}")

    def get_top_n(self, n=10):
        scored = []
        for r in self.results:
            area = parse_price_value(r.get("Powierzchnia", "0")) or 0
            total = parse_price_value(r.get("Cena_Suma", "99999")) or 99999
            animals = r.get("Zwierzeta", "Brak danych").lower()

            score = 0
            if area >= self.var_min_area.get():
                score += 4
            elif area > 0:
                score -= 2
            if total <= self.var_max_price.get():
                score += 4
            elif total < 99999:
                score -= 2
            if self.var_pets.get() and animals == "tak":
                score += 3
            elif self.var_pets.get() and animals == "nie":
                score -= 3
            if self.var_balcony.get() and r.get("Balkon") == "Tak":
                score += 2
            if self.var_gym.get() and r.get("Silownia") == "Tak":
                score += 2
            if self.var_gated.get() and r.get("Osiedle_Strzezone") == "Tak":
                score += 2
            if self.var_wifi.get() and r.get("WiFi") == "Tak":
                score += 1
            if area > 0 and total < 99999:
                score += max(0, 3 - (total / area / 30))

            scored.append((score, r))

        scored.sort(key=lambda x: -x[0])
        return [r for _, r in scored[:n]]

    # ══════════════════════════════════════════════════════════════════════════
    #  HELPERY UI
    # ══════════════════════════════════════════════════════════════════════════
    def clear_window(self):
        for w in self.winfo_children():
            w.destroy()

    def card_label(self, parent, text):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=ACCENT).pack(anchor="w", pady=(12, 6))

    def pick_dir(self):
        d = filedialog.askdirectory(initialdir=self.var_output_dir.get())
        if d:
            self.var_output_dir.set(d)

    def log(self, msg):
        def _log():
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
        self.after(0, _log)

    def update_status(self, text):
        self.after(0, lambda: self.lbl_status.configure(text=text))

    def update_progress(self, current, total):
        def _update():
            self.progress.set(current / total if total > 0 else 0)
            self.lbl_counter.configure(text=f"{current} / {total}")
        self.after(0, _update)

    def update_live_stats(self, found, budget_ok, budget_over, pets):
        def _update():
            self.stat_labels["Znalezione"].configure(text=str(found))
            self.stat_labels["W budzecie"].configure(text=str(budget_ok))
            self.stat_labels["Poza budzetem"].configure(text=str(budget_over))
            self.stat_labels["Ze zwierzetami"].configure(text=str(pets))
        self.after(0, _update)


if __name__ == "__main__":
    app = App()
    app.mainloop()
