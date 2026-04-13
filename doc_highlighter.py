"""
╔══════════════════════════════════════════════════════╗
║         Google Docs Auto-Highlighter                 ║
║         Navy & Gold Edition  —  v2                   ║
╚══════════════════════════════════════════════════════╝

Requirements:
    pip install customtkinter pyautogui pillow

Usage:
    1. Run this script
    2. Add words/sentences to the Highlight Queue and pick a color
    3. Open your Google Doc in a browser
    4. Go to Settings → Calibrate each color swatch position
    5. Click  ▶ START  — the script takes over mouse & keyboard
    6. Move mouse to top-left corner at any time to abort (failsafe)
"""

import customtkinter as ctk
from PIL import Image, ImageDraw
import math
import pyautogui
import threading
import time
import tkinter.messagebox as messagebox
import os
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# ── pyautogui ─────────────────────────────────────────────────────────────────
pyautogui.PAUSE = 0.05

# ── App appearance ─────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ══════════════════════════════════════════════════════════════════════════════
#  NAVY / GOLD THEME  (mirrors human_typer structure, recolored)
# ══════════════════════════════════════════════════════════════════════════════
GOLD         = ("#c9a84c", "#e8c56d")      # primary accent — gold
GOLD_HOVER   = ("#a8872e", "#f5d98b")
CARD_BG      = ("#162135", "#162135")      # card background  — dark navy
CARD_BDR     = ("#2a3f5e", "#2a3f5e")      # card border
SIDEBAR_BG   = ("#0a1520", "#0a1520")      # sidebar — deepest navy
APP_BG       = ("#0d1b2a", "#0d1b2a")      # main background
NAV_ACTIVE   = ("#1a3050", "#1a3050")
NAV_HOVER    = ("#152840", "#152840")
MUTED        = ("#5a7a9a", "#5a7a9a")
HDR_TEXT     = ("#e8eaf0", "#e8eaf0")
RED_ERR      = ("#e05c5c", "#e05c5c")
GOLD_RGB     = (232, 197, 109)
MUTED_RGB    = (90, 122, 154)
NAV_W        = 68

HTTP_PORT = 7798
_app      = None   # set at entry point; used by the HTTP handler

HIGHLIGHT_COLORS = {
    "Yellow": "#FFFF00",
    "Green":  "#00FF00",
    "Blue":   "#6fa8dc",
    "Red":    "#FF6666",
    "Purple": "#c27ba0",
    "Orange": "#FFB347",
    "Teal":   "#76d7c4",
    "Pink":   "#f48fb1",
}

# ══════════════════════════════════════════════════════════════════════════════
#  LOCAL HTTP SERVER  (Claude Code integration)
# ══════════════════════════════════════════════════════════════════════════════
class _HighlightHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # suppress console noise

    def _body(self):
        length = int(self.headers.get("Content-Length", 0))
        return self.rfile.read(length).decode("utf-8", errors="replace").strip()

    def _ok(self, msg="OK"):
        self.send_response(200); self.end_headers()
        self.wfile.write(msg.encode())

    def _err(self, code, msg):
        self.send_response(code); self.end_headers()
        self.wfile.write(msg.encode())

    def do_POST(self):
        if _app is None:
            self._err(503, "App not ready"); return
        body = self._body()

        if self.path in ("/add", "/add-and-start"):
            if not body:
                self._err(400, "No text provided"); return
            if "::" in body:
                text, color = body.split("::", 1)
                text  = text.strip()
                color = color.strip().capitalize()
                if color not in HIGHLIGHT_COLORS:
                    color = "Yellow"
            else:
                text  = body
                color = "Yellow"
            auto_start = self.path == "/add-and-start"
            def _inject_add(t=text, c=color, a=auto_start):
                _app.entries.append(HighlightEntry(t, c))
                _app._refresh_listbox()
                _app.show_page("main")
                _app._set_status(f'Added via Claude: "{t}"  [{c}]')
                if a:
                    _app._start_automation()
            _app.after(0, _inject_add)
            self._ok(f"Added '{text}' [{color}]" + (" — starting" if auto_start else ""))

        elif self.path == "/batch":
            if not body:
                self._err(400, "No text provided"); return
            parsed = _parse_highlight_lines(body)
            def _inject_batch(ps=parsed):
                for t, c in ps:
                    _app.entries.append(HighlightEntry(t, c))
                _app._refresh_listbox()
                _app.show_page("main")
                _app._set_status(f"Imported {len(ps)} entries via Claude.")
            _app.after(0, _inject_batch)
            self._ok(f"Batch imported ({len(parsed)} entries)")

        elif self.path == "/clear":
            def _inject_clear():
                _app.entries.clear()
                _app._refresh_listbox()
                _app._set_status("Queue cleared via Claude.")
            _app.after(0, _inject_clear)
            self._ok("Queue cleared")

        elif self.path == "/start":
            _app.after(0, _app._start_automation)
            self._ok("Starting automation")

        elif self.path == "/stop":
            _app.after(0, _app._stop_automation)
            self._ok("Stopped")

        else:
            self._err(404, "Not found")

    def do_GET(self):
        if _app is not None and self.path == "/status":
            data = json.dumps({
                "running": _app.running,
                "queue":   len(_app.entries),
            })
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(data.encode())
        else:
            self._err(404, "Not found")


def _parse_highlight_lines(raw: str, default_color: str = "Yellow") -> list:
    """Parse batch text into [(text, color), ...]. Supports 'text::Color' syntax."""
    entries = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if "::" in line:
            text, color = line.split("::", 1)
            text  = text.strip()
            color = color.strip().capitalize()
            if color not in HIGHLIGHT_COLORS:
                color = "Yellow"
        else:
            text  = line
            color = default_color
        entries.append((text, color))
    return entries


def _start_http_server():
    try:
        server = ThreadingHTTPServer(("127.0.0.1", HTTP_PORT), _HighlightHandler)
        server.serve_forever()
    except OSError:
        pass  # port already in use — fail silently

threading.Thread(target=_start_http_server, daemon=True).start()

# ══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════
class HighlightEntry:
    def __init__(self, text, color_name):
        self.text = text
        self.color_name = color_name

class CalibrationPoint:
    def __init__(self, label):
        self.label = label
        self.x = None
        self.y = None

    @property
    def is_set(self):
        return self.x is not None and self.y is not None

# ══════════════════════════════════════════════════════════════════════════════
#  ICON BUILDERS
# ══════════════════════════════════════════════════════════════════════════════
def make_grid_icon(size=22, color=GOLD_RGB):
    s = size * 4
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    gap = s // 7
    half = (s - gap) // 2
    r = s // 9
    for row in range(2):
        for col in range(2):
            x0 = col * (half + gap)
            y0 = row * (half + gap)
            d.rounded_rectangle([x0, y0, x0 + half, y0 + half], radius=r, fill=(*color, 255))
    return img.resize((size, size), Image.LANCZOS)

def make_gear_icon(size=22, color=GOLD_RGB):
    s = size * 4
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    cx, cy = s // 2, s // 2
    teeth = 8
    outer_r, tooth_r, inner_r, hole_r = int(s*.36), int(s*.46), int(s*.22), int(s*.13)
    th = math.pi / (teeth * 2.2)
    pts = []
    for i in range(teeth):
        base = 2 * math.pi * i / teeth
        for angle, radius in [
            (base - th * 1.6, outer_r), (base - th, tooth_r),
            (base + th, tooth_r),       (base + th * 1.6, outer_r),
        ]:
            pts.append((cx + radius * math.cos(angle), cy + radius * math.sin(angle)))
    d.polygon(pts, fill=(*color, 255))
    d.ellipse([cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r], fill=(*color, 255))
    d.ellipse([cx - hole_r,  cy - hole_r,  cx + hole_r,  cy + hole_r],  fill=(0, 0, 0, 0))
    return img.resize((size, size), Image.LANCZOS)

def make_highlight_icon(size=32, color=GOLD_RGB):
    s = size * 4
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    lh = s // 8
    yw = int(s * 0.55)
    d.rounded_rectangle([s//6, s//2 - lh, s//6 + yw, s//2 + lh],
                        radius=lh//2, fill=(*color, 255))
    for row_y in [s//3, int(s*0.70)]:
        d.rounded_rectangle([s//6, row_y - lh//2, int(s*0.78), row_y + lh//2],
                            radius=lh//3, fill=(*MUTED_RGB, 180))
    pen_pts = [
        (int(s*0.72), int(s*0.22)),
        (int(s*0.88), int(s*0.38)),
        (int(s*0.60), int(s*0.50)),
    ]
    d.polygon(pen_pts, fill=(*color, 240))
    return img.resize((size, size), Image.LANCZOS)

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ══════════════════════════════════════════════════════════════════════════════
class DocHighlighterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Doc Highlighter")
        self.geometry("820x700")
        self.resizable(True, True)
        self.configure(fg_color=APP_BG)

        # State
        self.entries: list[HighlightEntry] = []
        self.calibration: dict[str, CalibrationPoint] = {
            "highlight_btn": CalibrationPoint("Highlight Button")
        }
        for c in HIGHLIGHT_COLORS:
            self.calibration[f"color_{c}"] = CalibrationPoint(f"Color — {c}")

        self.delay_var       = ctk.DoubleVar(value=0.8)
        self.start_delay_var = ctk.IntVar(value=5)
        self.running         = False
        self._stop_flag      = False

        # Icons
        self._ico_home_active = ctk.CTkImage(make_grid_icon(22, GOLD_RGB),  size=(22, 22))
        self._ico_home_idle   = ctk.CTkImage(make_grid_icon(22, MUTED_RGB), size=(22, 22))
        self._ico_gear_active = ctk.CTkImage(make_gear_icon(22, GOLD_RGB),  size=(22, 22))
        self._ico_gear_idle   = ctk.CTkImage(make_gear_icon(22, MUTED_RGB), size=(22, 22))
        self._ico_logo        = ctk.CTkImage(make_highlight_icon(32, GOLD_RGB), size=(32, 32))

        self._build_layout()
        self.show_page("main")

    # ─── LAYOUT SKELETON ──────────────────────────────────────────────────────
    def _build_layout(self):
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=NAV_W, corner_radius=0, fg_color=SIDEBAR_BG)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        right = ctk.CTkFrame(self, corner_radius=0, fg_color=APP_BG)
        right.pack(side="left", fill="both", expand=True)

        # Sidebar contents
        ctk.CTkLabel(self.sidebar, image=self._ico_logo, text="").pack(pady=(16, 6))
        ctk.CTkFrame(self.sidebar, height=1, fg_color=CARD_BDR).pack(fill="x", padx=10, pady=6)

        self.nav_home_btn = ctk.CTkButton(
            self.sidebar, image=self._ico_home_active, text="",
            width=46, height=46, fg_color=NAV_ACTIVE,
            hover_color=NAV_HOVER, corner_radius=12,
            command=lambda: self.show_page("main"))
        self.nav_home_btn.pack(pady=(4, 2))

        self.nav_gear_btn = ctk.CTkButton(
            self.sidebar, image=self._ico_gear_idle, text="",
            width=46, height=46, fg_color="transparent",
            hover_color=NAV_HOVER, corner_radius=12,
            command=lambda: self.show_page("settings"))
        self.nav_gear_btn.pack(pady=2)

        # Page header
        page_header = ctk.CTkFrame(right, height=56, corner_radius=0, fg_color="transparent")
        page_header.pack(fill="x")
        page_header.pack_propagate(False)

        self.page_title_lbl = ctk.CTkLabel(
            page_header, text="Doc Highlighter",
            font=ctk.CTkFont(size=22, weight="bold"), text_color=HDR_TEXT)
        self.page_title_lbl.pack(side="left", padx=20)

        self.status_lbl = ctk.CTkLabel(
            page_header, text="Ready",
            font=ctk.CTkFont(size=11), text_color=MUTED)
        self.status_lbl.pack(side="right", padx=20)

        ctk.CTkFrame(right, height=1, fg_color=CARD_BDR).pack(fill="x")

        # Page container
        self.page_container = ctk.CTkFrame(right, corner_radius=0, fg_color="transparent")
        self.page_container.pack(fill="both", expand=True)
        self.page_container.grid_rowconfigure(0, weight=1)
        self.page_container.grid_columnconfigure(0, weight=1)

        self._build_main_page()
        self._build_settings_page()

    # ─── WIDGET HELPERS ───────────────────────────────────────────────────────
    def _card(self, parent, **kw):
        return ctk.CTkFrame(parent, corner_radius=12, fg_color=CARD_BG,
                            border_width=1, border_color=CARD_BDR, **kw)

    def _section_lbl(self, parent, text):
        ctk.CTkLabel(parent, text=text,
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=MUTED).pack(anchor="w", padx=4, pady=(14, 4))

    def _divider(self, parent):
        ctk.CTkFrame(parent, height=1, fg_color=CARD_BDR).pack(fill="x", padx=12, pady=4)

    def _gold_btn(self, parent, text, cmd, height=36, width=None, font_size=13):
        kw = {"height": height,
              "font": ctk.CTkFont(size=font_size, weight="bold"),
              "fg_color": GOLD, "hover_color": GOLD_HOVER,
              "text_color": ("#0d1b2a", "#0d1b2a"),
              "corner_radius": 8, "command": cmd, "text": text}
        if width:
            kw["width"] = width
        return ctk.CTkButton(parent, **kw)

    def _muted_btn(self, parent, text, cmd, height=30, width=None):
        kw = {"height": height,
              "font": ctk.CTkFont(size=11),
              "fg_color": "transparent", "hover_color": NAV_HOVER,
              "border_width": 1, "border_color": CARD_BDR,
              "text_color": MUTED, "corner_radius": 8,
              "command": cmd, "text": text}
        if width:
            kw["width"] = width
        return ctk.CTkButton(parent, **kw)

    # ─── MAIN PAGE ────────────────────────────────────────────────────────────
    def _build_main_page(self):
        self.main_wrap = ctk.CTkFrame(self.page_container, fg_color="transparent", corner_radius=0)
        self.main_wrap.grid(row=0, column=0, sticky="nsew")
        self.main_wrap.grid_rowconfigure(0, weight=1)
        self.main_wrap.grid_columnconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self.main_wrap, fg_color="transparent",
                                        corner_radius=0, scrollbar_button_color=CARD_BDR)
        scroll.grid(row=0, column=0, sticky="nsew", padx=16, pady=(10, 0))

        # ── Queue input ──────────────────────────────────────────────────────
        self._section_lbl(scroll, "HIGHLIGHT QUEUE")
        queue_card = self._card(scroll)
        queue_card.pack(fill="x", pady=(0, 2))

        entry_row = ctk.CTkFrame(queue_card, fg_color="transparent")
        entry_row.pack(fill="x", padx=12, pady=(12, 6))

        self.word_entry = ctk.CTkEntry(
            entry_row, placeholder_text="Word or sentence to highlight…",
            height=36, corner_radius=8, font=ctk.CTkFont(size=13),
            border_color=CARD_BDR, text_color=HDR_TEXT)
        self.word_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.word_entry.bind("<Return>", lambda e: self._add_entry())

        self.color_var = ctk.StringVar(value="Yellow")
        ctk.CTkOptionMenu(
            entry_row, variable=self.color_var,
            values=list(HIGHLIGHT_COLORS.keys()),
            width=110, height=36, corner_radius=8,
            fg_color=CARD_BG, button_color=CARD_BDR,
            button_hover_color=NAV_HOVER,
            text_color=HDR_TEXT, dropdown_fg_color=CARD_BG,
            dropdown_text_color=HDR_TEXT, dropdown_hover_color=NAV_HOVER,
            font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 8))

        self._gold_btn(entry_row, "+ Add", self._add_entry,
                       height=36, width=80).pack(side="left")

        # Queue list
        self.listbox_frame = ctk.CTkScrollableFrame(
            queue_card, height=180, fg_color=APP_BG, corner_radius=8,
            scrollbar_button_color=CARD_BDR)
        self.listbox_frame.pack(fill="x", padx=12, pady=(0, 8))
        self.listbox_rows: list = []

        # Queue actions
        qbtn_row = ctk.CTkFrame(queue_card, fg_color="transparent")
        qbtn_row.pack(fill="x", padx=12, pady=(0, 12))
        self._muted_btn(qbtn_row, "Remove Last", self._remove_last, height=28).pack(side="left", padx=(0, 6))
        self._muted_btn(qbtn_row, "Clear All",   self._clear_all,  height=28).pack(side="left")
        self.queue_count_lbl = ctk.CTkLabel(
            qbtn_row, text="0 entries",
            font=ctk.CTkFont(size=11), text_color=MUTED)
        self.queue_count_lbl.pack(side="right")

        # ── Batch input ──────────────────────────────────────────────────────
        self._section_lbl(scroll, "BATCH INPUT")
        batch_card = self._card(scroll)
        batch_card.pack(fill="x", pady=(0, 2))

        ctk.CTkLabel(batch_card,
                     text="One entry per line.  Assign a color with  word::Yellow",
                     font=ctk.CTkFont(size=11), text_color=MUTED, anchor="w"
                     ).pack(anchor="w", padx=16, pady=(10, 4))

        self.batch_text = ctk.CTkTextbox(
            batch_card, height=80, corner_radius=8,
            fg_color=APP_BG, border_width=0,
            font=ctk.CTkFont(size=12), text_color=HDR_TEXT)
        self.batch_text.pack(fill="x", padx=12, pady=(0, 8))

        bbtn_row = ctk.CTkFrame(batch_card, fg_color="transparent")
        bbtn_row.pack(fill="x", padx=12, pady=(0, 12))
        self._gold_btn(bbtn_row, "Import Batch", self._import_batch, height=32).pack(side="left")

        # ── Automation controls ───────────────────────────────────────────────
        self._section_lbl(scroll, "AUTOMATION")
        run_card = self._card(scroll)
        run_card.pack(fill="x", pady=(0, 4))

        # Countdown
        cd_row = ctk.CTkFrame(run_card, fg_color="transparent")
        cd_row.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(cd_row, text="Countdown Before Start",
                     font=ctk.CTkFont(size=13), text_color=HDR_TEXT).pack(side="left")
        for secs in (3, 5, 10, 15):
            ctk.CTkRadioButton(
                cd_row, text=f"{secs}s", variable=self.start_delay_var, value=secs,
                font=ctk.CTkFont(size=12), text_color=HDR_TEXT,
                fg_color=GOLD, hover_color=GOLD_HOVER).pack(side="right", padx=6)

        self._divider(run_card)

        # Action delay
        delay_top = ctk.CTkFrame(run_card, fg_color="transparent")
        delay_top.pack(fill="x", padx=14, pady=(8, 0))
        ctk.CTkLabel(delay_top, text="Delay Between Actions",
                     font=ctk.CTkFont(size=13), text_color=HDR_TEXT).pack(side="left")
        self.delay_val_lbl = ctk.CTkLabel(
            delay_top, text="0.8s",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=GOLD[1])
        self.delay_val_lbl.pack(side="right")

        ctk.CTkSlider(
            run_card, from_=0.2, to=2.0, number_of_steps=18,
            variable=self.delay_var,
            progress_color=GOLD[1], button_color=GOLD[1],
            button_hover_color=GOLD_HOVER[1],
            command=lambda v: self.delay_val_lbl.configure(text=f"{v:.1f}s")
        ).pack(fill="x", padx=14, pady=(6, 0))

        ctk.CTkLabel(run_card,
                     text="Increase if Google Docs misses highlights on a slow connection.",
                     font=ctk.CTkFont(size=11), text_color=MUTED, anchor="w"
                     ).pack(anchor="w", padx=16, pady=(2, 8))

        self._divider(run_card)

        ctk.CTkLabel(run_card,
                     text="⚠   Move mouse to top-left corner at any time to abort (failsafe).",
                     font=ctk.CTkFont(size=11), text_color=RED_ERR, anchor="w"
                     ).pack(anchor="w", padx=16, pady=(8, 6))

        self._divider(run_card)

        # Progress
        self.progress_var = ctk.DoubleVar(value=0)
        self.progress_bar = ctk.CTkProgressBar(
            run_card, variable=self.progress_var,
            progress_color=GOLD[1], fg_color=APP_BG,
            height=8, corner_radius=4)
        self.progress_bar.pack(fill="x", padx=14, pady=(10, 4))
        self.progress_bar.set(0)

        self.progress_lbl = ctk.CTkLabel(
            run_card, text="", font=ctk.CTkFont(size=11), text_color=MUTED)
        self.progress_lbl.pack(anchor="w", padx=16, pady=(0, 8))

        # Start / Stop
        btn_row = ctk.CTkFrame(run_card, fg_color="transparent")
        btn_row.pack(fill="x", padx=14, pady=(4, 14))

        self.start_btn = self._gold_btn(
            btn_row, "▶  Start Highlighting", self._start_automation,
            height=44, font_size=15)
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.stop_btn = ctk.CTkButton(
            btn_row, text="■  Stop", height=44, width=110,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent", border_width=1, border_color=RED_ERR,
            text_color=RED_ERR, hover_color=NAV_HOVER,
            corner_radius=8, state="disabled",
            command=self._stop_automation)
        self.stop_btn.pack(side="left")

        ctk.CTkFrame(scroll, height=20, fg_color="transparent").pack()

    # ─── SETTINGS PAGE ────────────────────────────────────────────────────────
    def _build_settings_page(self):
        self.settings_wrap = ctk.CTkFrame(self.page_container, fg_color="transparent", corner_radius=0)
        self.settings_wrap.grid(row=0, column=0, sticky="nsew")
        self.settings_wrap.grid_rowconfigure(0, weight=1)
        self.settings_wrap.grid_columnconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(self.settings_wrap, fg_color="transparent",
                                        corner_radius=0, scrollbar_button_color=CARD_BDR)
        scroll.grid(row=0, column=0, sticky="nsew", padx=16, pady=(10, 0))

        # ── How to ──────────────────────────────────────────────────────────
        self._section_lbl(scroll, "HOW TO CALIBRATE")
        how_card = self._card(scroll)
        how_card.pack(fill="x", pady=(0, 4))
        steps = [
            "1.  Open Google Docs in your browser at the zoom level you normally use.",
            "2.  Click the Highlight colour button in the Docs toolbar (the A with underline).",
            "3.  The colour palette will open — leave it open.",
            "4.  Click CAPTURE next to each colour.  You have 5 seconds to hover over\n     that swatch before the position is recorded.",
            "5.  Also capture the Highlight Button itself (close the palette first).",
            "6.  Return to the main page and press  ▶  Start Highlighting.",
        ]
        for s in steps:
            ctk.CTkLabel(how_card, text=s, font=ctk.CTkFont(size=12),
                         text_color=HDR_TEXT, anchor="w", justify="left",
                         wraplength=580).pack(anchor="w", padx=16, pady=(6, 0))
        ctk.CTkFrame(how_card, height=10, fg_color="transparent").pack()

        # ── Highlight button ─────────────────────────────────────────────────
        self._section_lbl(scroll, "TOOLBAR BUTTON")
        hbtn_card = self._card(scroll)
        hbtn_card.pack(fill="x", pady=(0, 4))

        row = ctk.CTkFrame(hbtn_card, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=12)
        ctk.CTkLabel(row, text="Highlight Button  (A icon in toolbar)",
                     font=ctk.CTkFont(size=13), text_color=HDR_TEXT).pack(side="left")
        self.hbtn_status = ctk.CTkLabel(
            row, text="not set",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=RED_ERR)
        self.hbtn_status.pack(side="right", padx=(0, 8))
        self._muted_btn(
            row, "CAPTURE  (5s)",
            lambda: self._capture("highlight_btn", self.hbtn_status),
            height=30, width=130).pack(side="right")

        # ── Color swatches ───────────────────────────────────────────────────
        self._section_lbl(scroll, "COLOUR SWATCH POSITIONS")
        colors_card = self._card(scroll)
        colors_card.pack(fill="x", pady=(0, 4))

        self.color_status_labels: dict[str, ctk.CTkLabel] = {}
        for i, c in enumerate(HIGHLIGHT_COLORS):
            if i > 0:
                self._divider(colors_card)
            row = ctk.CTkFrame(colors_card, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=6)

            ctk.CTkLabel(row, text="⬤",
                         font=ctk.CTkFont(size=16),
                         text_color=HIGHLIGHT_COLORS[c]).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(row, text=c,
                         font=ctk.CTkFont(size=13),
                         text_color=HDR_TEXT, width=70, anchor="w").pack(side="left")

            key = f"color_{c}"
            status_lbl = ctk.CTkLabel(
                row, text="not set",
                font=ctk.CTkFont(size=12, weight="bold"), text_color=RED_ERR)
            status_lbl.pack(side="right", padx=(0, 8))
            self.color_status_labels[key] = status_lbl

            self._muted_btn(
                row, "CAP",
                lambda k=key, l=status_lbl: self._capture(k, l),
                height=28, width=56).pack(side="right")

        # ── Claude Code integration ──────────────────────────────────────────
        self._section_lbl(scroll, "CLAUDE CODE INTEGRATION")
        claude_card = self._card(scroll)
        claude_card.pack(fill="x", pady=(0, 4))

        top_row = ctk.CTkFrame(claude_card, fg_color="transparent")
        top_row.pack(fill="x", padx=16, pady=(14, 4))
        ctk.CTkLabel(top_row, text="Local HTTP Server",
                     font=ctk.CTkFont(size=13), text_color=HDR_TEXT).pack(side="left")
        ctk.CTkLabel(top_row, text=f"port {HTTP_PORT}  ·  running",
                     font=ctk.CTkFont(size=12, weight="bold"), text_color=GOLD[1]).pack(side="right")

        ctk.CTkLabel(claude_card,
                     text="Ask Claude Code to load words and start highlighting. Four endpoints:",
                     font=ctk.CTkFont(size=11), text_color=MUTED, anchor="w",
                     ).pack(anchor="w", padx=16, pady=(0, 6))

        for label, cmd in (
            ("Add one entry (text::Color optional)",
             f'curl -X POST http://localhost:{HTTP_PORT}/add -d "keyword::Yellow"'),
            ("Add + start immediately",
             f'curl -X POST http://localhost:{HTTP_PORT}/add-and-start -d "keyword::Green"'),
            ("Batch import (one per line, text::Color)",
             f'curl -X POST http://localhost:{HTTP_PORT}/batch -d "word1::Yellow\\nword2::Blue"'),
            ("Clear queue",
             f'curl -X POST http://localhost:{HTTP_PORT}/clear'),
        ):
            blk = ctk.CTkFrame(claude_card, fg_color=APP_BG, corner_radius=8)
            blk.pack(fill="x", padx=16, pady=(0, 6))
            ctk.CTkLabel(blk, text=label,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color=MUTED).pack(anchor="w", padx=10, pady=(6, 0))
            ctk.CTkLabel(blk, text=cmd,
                         font=ctk.CTkFont(family="Courier", size=11),
                         text_color=HDR_TEXT, anchor="w", wraplength=520, justify="left",
                         ).pack(anchor="w", padx=10, pady=(2, 6))

        ctk.CTkLabel(claude_card,
                     text='  Tip: tell Claude "send to Doc Highlighter" and it will run the curl command for you.',
                     font=ctk.CTkFont(size=11), text_color=GOLD[1], anchor="w", wraplength=520,
                     ).pack(anchor="w", padx=16, pady=(0, 12))

        ctk.CTkFrame(scroll, height=20, fg_color="transparent").pack()

    # ─── PAGE SWITCHING ───────────────────────────────────────────────────────
    def show_page(self, name):
        if name == "main":
            self.main_wrap.tkraise()
            self.page_title_lbl.configure(text="Doc Highlighter")
            self.nav_home_btn.configure(image=self._ico_home_active, fg_color=NAV_ACTIVE)
            self.nav_gear_btn.configure(image=self._ico_gear_idle,   fg_color="transparent")
        else:
            self.settings_wrap.tkraise()
            self.page_title_lbl.configure(text="Settings  ·  Calibration")
            self.nav_home_btn.configure(image=self._ico_home_idle,   fg_color="transparent")
            self.nav_gear_btn.configure(image=self._ico_gear_active, fg_color=NAV_ACTIVE)

    # ─── QUEUE MANAGEMENT ─────────────────────────────────────────────────────
    def _refresh_listbox(self):
        for w in self.listbox_frame.winfo_children():
            w.destroy()
        self.listbox_rows.clear()
        for e in self.entries:
            row = ctk.CTkFrame(self.listbox_frame, corner_radius=6,
                               fg_color=CARD_BG, height=34)
            row.pack(fill="x", pady=2)
            row.pack_propagate(False)

            ctk.CTkLabel(row, text="⬤",
                         font=ctk.CTkFont(size=12),
                         text_color=HIGHLIGHT_COLORS.get(e.color_name, "#888")
                         ).pack(side="left", padx=(8, 4))
            ctk.CTkLabel(row, text=f"[{e.color_name}]",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=GOLD[1], width=72, anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=e.text,
                         font=ctk.CTkFont(size=12),
                         text_color=HDR_TEXT, anchor="w"
                         ).pack(side="left", fill="x", expand=True)
            self.listbox_rows.append(row)

        n = len(self.entries)
        self.queue_count_lbl.configure(
            text=f"{n} {'entry' if n == 1 else 'entries'}")

    def _add_entry(self):
        text  = self.word_entry.get().strip()
        color = self.color_var.get()
        if not text:
            return
        self.entries.append(HighlightEntry(text, color))
        self.word_entry.delete(0, "end")
        self._refresh_listbox()
        self._set_status(f"Added: \"{text}\"  [{color}]")

    def _remove_last(self):
        if self.entries:
            self.entries.pop()
            self._refresh_listbox()

    def _clear_all(self):
        if not self.entries:
            return
        if messagebox.askyesno("Clear All", "Remove all entries from the queue?"):
            self.entries.clear()
            self._refresh_listbox()

    def _import_batch(self):
        raw = self.batch_text.get("1.0", "end").strip()
        if not raw:
            return
        parsed = _parse_highlight_lines(raw, default_color=self.color_var.get())
        for text, color in parsed:
            self.entries.append(HighlightEntry(text, color))
        self.batch_text.delete("1.0", "end")
        self._refresh_listbox()
        self._set_status(f"Imported {len(parsed)} entries.")

    # ─── CALIBRATION ──────────────────────────────────────────────────────────
    def _capture(self, key: str, label: ctk.CTkLabel):
        """5-second delay then capture mouse position."""
        self._set_status(f"Hover over the target — capturing in 5 seconds…")
        label.configure(text="capturing…", text_color=GOLD[1])
        self.after(5000, lambda: self._do_capture(key, label))

    def _do_capture(self, key: str, label: ctk.CTkLabel):
        x, y = pyautogui.position()
        self.calibration[key].x = x
        self.calibration[key].y = y
        label.configure(text=f"✓  ({x}, {y})", text_color=GOLD[1])
        self._set_status(f"Captured '{key}' at ({x}, {y})")

    # ─── AUTOMATION ───────────────────────────────────────────────────────────
    def _start_automation(self):
        if not self.entries:
            messagebox.showwarning("No Entries", "Add at least one word or sentence first.")
            return
        if not self.calibration["highlight_btn"].is_set:
            messagebox.showwarning("Calibration Missing",
                                   "Go to Settings and calibrate the Highlight Button first.")
            return
        colors_needed = {e.color_name for e in self.entries}
        missing = [c for c in colors_needed
                   if not self.calibration.get(f"color_{c}", CalibrationPoint("")).is_set]
        if missing:
            ok = messagebox.askyesno(
                "Missing Calibration",
                f"These colors haven't been calibrated: {', '.join(missing)}\n\n"
                "Continue anyway? (those entries will be skipped)")
            if not ok:
                return

        self._stop_flag = False
        self.running    = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress_bar.set(0)
        self.progress_lbl.configure(text="")

        threading.Thread(target=self._countdown_then_run, daemon=True).start()

    def _countdown_then_run(self):
        countdown = self.start_delay_var.get()
        for i in range(countdown, 0, -1):
            if self._stop_flag:
                self.after(0, self._on_done)
                return
            self._set_status(f"Starting in {i}s — switch to your Google Doc now…")
            time.sleep(1)
        self._run_automation()

    def _stop_automation(self):
        self._stop_flag = True
        self._set_status("Stopping after current action…")

    def _run_automation(self):
        entries = list(self.entries)
        total   = len(entries)
        delay   = self.delay_var.get()
        hbtn    = self.calibration["highlight_btn"]

        for idx, entry in enumerate(entries):
            if self._stop_flag:
                break

            text     = entry.text
            color    = entry.color_name
            color_pt = self.calibration.get(f"color_{color}")

            self.after(0, lambda f=idx/total: self.progress_bar.set(f))
            self.after(0, lambda i=idx, t=total: self.progress_lbl.configure(
                text=f"{i+1} / {t}"))
            self._set_status(f"[{idx+1}/{total}]  \"{text}\"  →  {color}")

            try:
                # 1. Open Find bar
                pyautogui.hotkey("ctrl", "f")
                time.sleep(delay)

                # 2. Select all in search box and type new text
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.15)
                pyautogui.typewrite(text, interval=0.05)
                time.sleep(delay)

                # 3. Jump to first match (text gets selected)
                pyautogui.press("enter")
                time.sleep(delay * 0.6)

                # 4. Close Find bar — selection is preserved in Google Docs
                pyautogui.press("escape")
                time.sleep(delay * 0.6)

                # 5. Apply highlight via toolbar
                if color_pt and color_pt.is_set:
                    pyautogui.click(hbtn.x, hbtn.y)
                    time.sleep(delay)
                    pyautogui.click(color_pt.x, color_pt.y)
                    time.sleep(delay * 0.5)
                else:
                    self._set_status(f"⚠  Skipped '{color}' — not calibrated.")
                    time.sleep(delay * 0.3)

            except pyautogui.FailSafeException:
                self._set_status("Failsafe triggered — stopped.")
                break
            except Exception as exc:
                self._set_status(f"Error on \"{text}\": {exc}")
                time.sleep(delay)

        self.running = False
        self.after(0, self._on_done)

    def _on_done(self):
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.progress_bar.set(1)
        self.progress_lbl.configure(text="Done ✓")
        self._set_status("Automation complete — review your Google Doc.")

    def _set_status(self, msg):
        self.after(0, lambda: self.status_lbl.configure(text=msg))


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = DocHighlighterApp()
    _app = app
    app.mainloop()