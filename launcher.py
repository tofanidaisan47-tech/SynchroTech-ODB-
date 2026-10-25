"""
launcher.py — SyncroTech ODB
────────────────────────────────────────────────────────────────────────
Aircraft Flap Sensor & Synchro Motor Diagnostic Suite
GMF AeroAsia · Avionics & Component Services

Unified multiprocess workstation for:
  [01] API 8810A Data Logger (Flap Sensor Digital Protractor)
  [02] Synchro Motor Voltage Measurement System (Agilent 34401A)

Designed under anti-AI-slop principles (Hallmark framework):
  - Utilitarian / Aerospace Cockpit layout (no cookie-cutter cards or emojis)
  - Strict typographic hierarchy (Segoe UI + Consolas monospace telemetry)
  - Isolated multiprocess architecture (zero theme/mainloop/thread collision)
  - PBKDF2-SHA256 authenticated console access

Author : Integration Engineering – GMF AeroAsia
System : SyncroTech ODB v2.0
"""

# ── stdlib ────────────────────────────────────────────────────────────
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Dict, Optional

# ── third-party ───────────────────────────────────────────────────────
import customtkinter as ctk

# ═════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM — Hallmark Aerospace Dark Console
# ═════════════════════════════════════════════════════════════════════
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Palette tokens (High contrast, muted industrial dark)
C_BG            = "#080C14"   # Deep obsidian background
C_PANEL         = "#0F172A"   # Slate-900 surface
C_PANEL_ALT     = "#131D31"   # Elevated card surface
C_BORDER        = "#1E293B"   # Slate-800 hairline border
C_BORDER_FOCUS  = "#38BDF8"   # Focus / active border
C_TEXT_PRIMARY  = "#F8FAFC"   # Slate-50 high contrast text
C_TEXT_MUTED    = "#94A3B8"   # Slate-400 secondary text
C_TEXT_DIM      = "#64748B"   # Slate-500 technical labels
C_CYAN          = "#0284C7"   # Primary instrument accent (API 8810A)
C_CYAN_HOVER    = "#0369A1"
C_AMBER         = "#D97706"   # Secondary instrument accent (Synchro)
C_AMBER_HOVER   = "#B45309"
C_EMERALD       = "#10B981"   # Online / telemetry pass
C_ROSE          = "#E11D48"   # Error / alert
C_SLATE_BTN     = "#1E293B"
C_SLATE_HOVER   = "#334155"

FONT_SANS = "Segoe UI"
FONT_MONO = "Consolas"

# ═════════════════════════════════════════════════════════════════════
#  SECURITY & AUTHENTICATION
# ═════════════════════════════════════════════════════════════════════
_SALT = b"aims_gmf_aeroasia_2024"
_ITER = 100_000
_DEFAULT_PW = "000"
_CFG_FILE = "password_config.json"


def _hash_pw(password: str) -> str:
    """Derive hex digest using PBKDF2-SHA256."""
    return hashlib.pbkdf2_hmac("sha256", password.encode(), _SALT, _ITER).hex()


def _load_or_create_config(base_dir: Path) -> dict:
    """Load password config, checking root and _internal locations."""
    candidates = [
        base_dir / _CFG_FILE,
        base_dir / "_internal" / _CFG_FILE,
        Path(__file__).resolve().parent / _CFG_FILE,
    ]
    for p in candidates:
        if p.is_file():
            try:
                data = json.loads(p.read_text("utf-8"))
                if "password_hash" in data:
                    return data
            except Exception:
                pass

    # Default fallback config with '000'
    cfg = {
        "password_hash": _hash_pw(_DEFAULT_PW),
        "note": f"Default password is {_DEFAULT_PW}"
    }
    try:
        (base_dir / _CFG_FILE).write_text(json.dumps(cfg, indent=2), "utf-8")
    except OSError:
        pass
    return cfg


# ═════════════════════════════════════════════════════════════════════
#  APPLICATION CORE
# ═════════════════════════════════════════════════════════════════════
class SyncroTechODBApp(ctk.CTk):
    """SyncroTech ODB Aerospace Diagnostic Workstation Launcher."""

    def __init__(self):
        super().__init__()

        # Window geometry
        self.title("SyncroTech ODB — Aircraft Diagnostic Workstation")
        self.geometry("1060x680")
        self.minsize(980, 620)
        self.configure(fg_color=C_BG)
        self._center_window(1060, 680)

        # Icon
        self._apply_icon()

        # Paths
        self._root_dir = Path(__file__).resolve().parent
        if getattr(sys, "frozen", False):
            self._root_dir = Path(sys.executable).resolve().parent

        # Auth
        self._cfg = _load_or_create_config(self._root_dir)

        # Active processes: subdir -> subprocess.Popen
        self._procs: Dict[str, subprocess.Popen] = {}

        # Frame container
        self._frame: Optional[ctk.CTkFrame] = None
        self._show_login()
        self.protocol("WM_DELETE_WINDOW", self._exit)

    def _center_window(self, w: int, h: int):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    def _apply_icon(self):
        candidates = [
            Path(__file__).resolve().parent / "datalogger_repo" / "logo.ico",
            Path(getattr(sys, "_MEIPASS", "")) / "datalogger_repo" / "logo.ico",
            Path(getattr(sys, "_MEIPASS", "")) / "logo.ico",
        ]
        for p in candidates:
            if p.is_file():
                try:
                    self.iconbitmap(str(p))
                except Exception:
                    pass
                return

    def _swap_frame(self) -> ctk.CTkFrame:
        if self._frame:
            self._frame.destroy()
        self._frame = ctk.CTkFrame(self, fg_color=C_BG)
        self._frame.pack(fill="both", expand=True)
        return self._frame

    # ─────────────────────────────────────────────────────────────────
    #  LOGIN VIEW — Asymmetric Industrial Split
    # ─────────────────────────────────────────────────────────────────
    def _show_login(self):
        root = self._swap_frame()

        # Center stage container
        shell = ctk.CTkFrame(
            root,
            fg_color=C_PANEL,
            corner_radius=8,
            border_width=1,
            border_color=C_BORDER,
            width=880,
            height=500,
        )
        shell.place(relx=0.5, rely=0.5, anchor="center")
        shell.pack_propagate(False)

        # ── Left Column: Telemetry & Technical ID ─────────────────────
        left_col = ctk.CTkFrame(shell, fg_color="#0B1324", corner_radius=0, width=420)
        left_col.pack(side="left", fill="both", expand=False)
        left_col.pack_propagate(False)

        left_pad = ctk.CTkFrame(left_col, fg_color="transparent")
        left_pad.pack(fill="both", expand=True, padx=36, pady=36)

        # Organization Header
        ctk.CTkLabel(
            left_pad,
            text="GMF AEROASIA  ·  AVIONICS COMPONENT SERVICES",
            font=(FONT_MONO, 10, "bold"),
            text_color=C_CYAN,
            anchor="w",
        ).pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            left_pad,
            text="SyncroTech ODB",
            font=(FONT_SANS, 34, "bold"),
            text_color=C_TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")

        ctk.CTkLabel(
            left_pad,
            text="Flap Sensor & Synchro Diagnostic Suite",
            font=(FONT_SANS, 13),
            text_color=C_TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(2, 24))

        # Technical Telemetry Card
        spec_box = ctk.CTkFrame(
            left_pad,
            fg_color="#080D18",
            corner_radius=6,
            border_width=1,
            border_color=C_BORDER,
        )
        spec_box.pack(fill="x", pady=(0, 24))

        spec_lines = [
            ("SYSTEM ID", "DF-AVN-2026"),
            ("INSTRUMENT 01", "API 8810A PROTRACTOR"),
            ("INSTRUMENT 02", "34401A SYNCHRO MOTOR"),
            ("BUS ARCH", "MULTIPROCESS ISOLATED"),
            ("SHEET ENGINE", "EXCEL LIVE ADAPTER"),
        ]
        for idx, (label, val) in enumerate(spec_lines):
            row = ctk.CTkFrame(spec_box, fg_color="transparent")
            row.pack(fill="x", padx=14, pady=5 if idx > 0 else (10, 5))
            ctk.CTkLabel(
                row,
                text=label,
                font=(FONT_MONO, 10),
                text_color=C_TEXT_DIM,
                anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row,
                text=val,
                font=(FONT_MONO, 10, "bold"),
                text_color=C_TEXT_PRIMARY,
                anchor="e",
            ).pack(side="right")
        # Bottom spacer inside spec box
        ctk.CTkFrame(spec_box, fg_color="transparent", height=6).pack()

        # Status badge
        status_bar = ctk.CTkFrame(left_pad, fg_color="transparent")
        status_bar.pack(side="bottom", fill="x")
        ctk.CTkLabel(
            status_bar,
            text="● CONSOLE READY FOR AUTHENTICATION",
            font=(FONT_MONO, 10),
            text_color=C_EMERALD,
            anchor="w",
        ).pack(side="left")

        # ── Vertical Divider ──────────────────────────────────────────
        ctk.CTkFrame(shell, fg_color=C_BORDER, width=1).pack(side="left", fill="y")

        # ── Right Column: Access Authentication ───────────────────────
        right_col = ctk.CTkFrame(shell, fg_color="transparent")
        right_col.pack(side="right", fill="both", expand=True)

        right_pad = ctk.CTkFrame(right_col, fg_color="transparent")
        right_pad.pack(fill="both", expand=True, padx=40, pady=40)

        ctk.CTkLabel(
            right_pad,
            text="Console Access",
            font=(FONT_SANS, 20, "bold"),
            text_color=C_TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")

        ctk.CTkLabel(
            right_pad,
            text="Enter authorized operator PIN code to initialize workstation instruments.",
            font=(FONT_SANS, 12),
            text_color=C_TEXT_MUTED,
            anchor="w",
            wraplength=340,
            justify="left",
        ).pack(fill="x", pady=(4, 28))

        ctk.CTkLabel(
            right_pad,
            text="OPERATOR ACCESS CODE",
            font=(FONT_MONO, 10, "bold"),
            text_color=C_TEXT_DIM,
            anchor="w",
        ).pack(fill="x", pady=(0, 6))

        # PIN Entry Row
        pw_box = ctk.CTkFrame(
            right_pad,
            fg_color="#090E17",
            border_width=1,
            border_color=C_BORDER,
            corner_radius=6,
            height=44,
        )
        pw_box.pack(fill="x")
        pw_box.pack_propagate(False)

        self._pw_entry = ctk.CTkEntry(
            pw_box,
            placeholder_text="Enter PIN (e.g. 000)",
            show="•",
            font=(FONT_MONO, 14),
            fg_color="transparent",
            border_width=0,
            text_color=C_TEXT_PRIMARY,
        )
        self._pw_entry.pack(side="left", fill="both", expand=True, padx=12)

        self._pw_visible = False
        self._btn_eye = ctk.CTkButton(
            pw_box,
            text="SHOW",
            width=54,
            height=32,
            font=(FONT_MONO, 10, "bold"),
            fg_color="transparent",
            hover_color=C_BORDER,
            text_color=C_TEXT_MUTED,
            command=self._toggle_password_visibility,
        )
        self._btn_eye.pack(side="right", padx=6)

        # Login button
        self._btn_login = ctk.CTkButton(
            right_pad,
            text="AUTHENTICATE  ->",
            height=42,
            corner_radius=6,
            font=(FONT_MONO, 12, "bold"),
            fg_color=C_CYAN,
            hover_color=C_CYAN_HOVER,
            text_color="#FFFFFF",
            command=self._attempt_login,
        )
        self._btn_login.pack(fill="x", pady=(18, 0))

        # Error notification label
        self._lbl_error = ctk.CTkLabel(
            right_pad,
            text="",
            font=(FONT_MONO, 11),
            text_color=C_ROSE,
            anchor="w",
        )
        self._lbl_error.pack(fill="x", pady=(10, 0))

        # Developer credit at bottom
        credit_box = ctk.CTkFrame(right_pad, fg_color="transparent")
        credit_box.pack(side="bottom", fill="x")
        ctk.CTkLabel(
            credit_box,
            text="Developed by Daisan Tofani, Odayakana Upupha Konnikof, Bramantyo Al Fajri  18 Oktober 2026",
            font=(FONT_SANS, 9),
            text_color=C_TEXT_DIM,
            anchor="w",
            wraplength=380,
            justify="left",
        ).pack(side="left", fill="x")

        # Focus & bind Enter
        self._pw_entry.bind("<Return>", lambda _: self._attempt_login())
        self._pw_entry.focus_set()

    def _toggle_password_visibility(self):
        self._pw_visible = not self._pw_visible
        self._pw_entry.configure(show="" if self._pw_visible else "•")
        self._btn_eye.configure(text="HIDE" if self._pw_visible else "SHOW")

    def _attempt_login(self):
        entered = self._pw_entry.get()
        if _hash_pw(entered) == self._cfg.get("password_hash", ""):
            self._lbl_error.configure(text="")
            self._show_dashboard()
        else:
            self._lbl_error.configure(text="[!] ACCESS DENIED: INVALID PASSCODE")
            self._pw_entry.delete(0, "end")
            self._pw_entry.focus_set()

    # -----------------------------------------------------------------
    #  DASHBOARD VIEW - Aerospace Technical Instrument Deck
    # -----------------------------------------------------------------
    def _show_dashboard(self):
        root = self._swap_frame()

        # -- Top Nav Masthead -----------------------------------------
        top_bar = ctk.CTkFrame(root, fg_color=C_PANEL, height=56, corner_radius=0)
        top_bar.pack(fill="x")
        top_bar.pack_propagate(False)

        # Brand / Station identity
        nav_left = ctk.CTkFrame(top_bar, fg_color="transparent")
        nav_left.pack(side="left", padx=24)

        ctk.CTkLabel(
            nav_left,
            text="SyncroTech ODB",
            font=(FONT_SANS, 18, "bold"),
            text_color=C_TEXT_PRIMARY,
        ).pack(side="left")

        ctk.CTkLabel(
            nav_left,
            text="[ GMF-STATION-01 ]",
            font=(FONT_MONO, 11),
            text_color=C_CYAN,
        ).pack(side="left", padx=(12, 8))

        ctk.CTkLabel(
            nav_left,
            text="ONLINE",
            font=(FONT_MONO, 10, "bold"),
            text_color=C_EMERALD,
        ).pack(side="left")

        # Actions on right
        nav_right = ctk.CTkFrame(top_bar, fg_color="transparent")
        nav_right.pack(side="right", padx=20)

        ctk.CTkLabel(
            nav_right,
            text="OPERATOR: AUTHORIZED",
            font=(FONT_MONO, 10),
            text_color=C_TEXT_DIM,
        ).pack(side="left", padx=(0, 16))

        ctk.CTkButton(
            nav_right,
            text="Lock Session",
            width=94,
            height=30,
            font=(FONT_MONO, 11),
            fg_color="transparent",
            hover_color=C_SLATE_HOVER,
            border_width=1,
            border_color=C_BORDER,
            text_color=C_TEXT_MUTED,
            command=self._logout,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            nav_right,
            text="Exit [X]",
            width=68,
            height=30,
            font=(FONT_MONO, 11, "bold"),
            fg_color="#991B1B",
            hover_color=C_ROSE,
            text_color="#FFFFFF",
            command=self._exit,
        ).pack(side="left")

        # Border separator
        ctk.CTkFrame(root, fg_color=C_BORDER, height=1, corner_radius=0).pack(fill="x")

        # Telemetry Ribbon
        ribbon = ctk.CTkFrame(root, fg_color="#0A101D", height=32, corner_radius=0)
        ribbon.pack(fill="x")
        ribbon.pack_propagate(False)

        ctk.CTkLabel(
            ribbon,
            text="TARGET: B737/A320 FLAP RIGGING INSPECTION   |   I/O BUS: ISOLATED PROCESSES   |   EXCEL ENGINE: DIRECT SYNC",
            font=(FONT_MONO, 10),
            text_color=C_TEXT_DIM,
        ).pack(side="left", padx=24)

        # ── Main Content Area ────────────────────────────────────────
        deck = ctk.CTkFrame(root, fg_color="transparent")
        deck.pack(fill="both", expand=True, padx=28, pady=20)

        # Subtitle
        ctk.CTkLabel(
            deck,
            text="SELECT SUBSYSTEM TO LAUNCH",
            font=(FONT_MONO, 11, "bold"),
            text_color=C_TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(0, 12))

        # Instrument Grid (Two primary cards)
        cards_row = ctk.CTkFrame(deck, fg_color="transparent")
        cards_row.pack(fill="both", expand=True)
        cards_row.columnconfigure(0, weight=1)
        cards_row.columnconfigure(1, weight=1)
        cards_row.rowconfigure(0, weight=1)

        # ── CARD 01: API 8810A Data Logger ───────────────────────────
        self._build_instrument_bay(
            parent=cards_row,
            col=0,
            badge="SUBSYSTEM 01 - PROTRACTOR",
            title="API 8810A Data Logger",
            subtitle="Digital Protractor & Flap Rigging Inspector",
            desc="Automated angle acquisition, live stability detection, and direct Excel sheet logging for API 8810A digital protractor.",
            accent_color=C_CYAN,
            hover_color=C_CYAN_HOVER,
            specs=[
                ("HARDWARE INTERFACE", "RS-232 / TCP ETHERNET"),
                ("TARGET WORKBOOK", "MASTER_rev3.xlsx"),
                ("DEFAULT UI MODE", "ENTERPRISE LIGHT"),
                ("DATA PRECISION", "0.01 DEG / DUAL CW-CCW"),
            ],
            btn_text="LAUNCH API 8810A LOGGER  ->",
            command=lambda: self._launch(
                "datalogger_repo",
                "Flap_Data_Logger_GMF.exe",
                "API 8810A Data Logger",
            ),
        )

        # -- CARD 02: Synchro Voltage Measurement System --------------
        self._build_instrument_bay(
            parent=cards_row,
            col=1,
            badge="SUBSYSTEM 02 - VOLTMETER",
            title="Synchro Voltage Monitor",
            subtitle="Agilent 34401A Flap Position Sensor",
            desc="High-precision digital multimeter voltage measurement for synchro motor accuracy verification (Stall, Autopilot, Flap).",
            accent_color=C_AMBER,
            hover_color=C_AMBER_HOVER,
            specs=[
                ("HARDWARE INTERFACE", "GPIB / PyVISA PROTOCOL"),
                ("TARGET WORKBOOK", "2708 Accuracy Sheet.xlsx"),
                ("DEFAULT UI MODE", "DARK NEON VFD"),
                ("ACQUISITION RATE", "1.6s STABILITY THRESHOLD"),
            ],
            btn_text="LAUNCH SYNCHRO MONITOR  ->",
            command=lambda: self._launch(
                "synchro_app",
                "VFlap.exe",
                "Synchro Voltage Monitor",
            ),
        )

        # ── Bottom Cockpit Strip: Dual Inspection Mode ───────────────
        dual_bar = ctk.CTkFrame(
            deck,
            fg_color=C_PANEL,
            corner_radius=6,
            border_width=1,
            border_color=C_BORDER,
            height=64,
        )
        dual_bar.pack(fill="x", pady=(16, 0))
        dual_bar.pack_propagate(False)

        dual_info = ctk.CTkFrame(dual_bar, fg_color="transparent")
        dual_info.pack(side="left", padx=20, pady=10)

        ctk.CTkLabel(
            dual_info,
            text="DUAL SUBSYSTEM INSPECTION",
            font=(FONT_MONO, 11, "bold"),
            text_color=C_TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")

        ctk.CTkLabel(
            dual_info,
            text="Initialize both API 8810A Logger and Synchro Voltage Monitor concurrently in separate isolated processes.",
            font=(FONT_SANS, 11),
            text_color=C_TEXT_MUTED,
            anchor="w",
        ).pack(fill="x")

        ctk.CTkButton(
            dual_bar,
            text="LAUNCH BOTH INSTRUMENTS  [1 + 2]",
            width=260,
            height=38,
            corner_radius=4,
            font=(FONT_MONO, 11, "bold"),
            fg_color="#4F46E5",
            hover_color="#4338CA",
            text_color="#FFFFFF",
            command=self._launch_both,
        ).pack(side="right", padx=16, pady=12)

    def _build_instrument_bay(
        self, parent, col, badge, title, subtitle, desc, accent_color, hover_color, specs, btn_text, command
    ):
        """Build an authentic aerospace instrument card."""
        card = ctk.CTkFrame(
            parent,
            fg_color=C_PANEL,
            corner_radius=8,
            border_width=1,
            border_color=C_BORDER,
        )
        card.grid(row=0, column=col, padx=8, pady=0, sticky="nsew")

        pad = ctk.CTkFrame(card, fg_color="transparent")
        pad.pack(fill="both", expand=True, padx=22, pady=20)

        # Header Badge
        top_tag = ctk.CTkFrame(pad, fg_color="transparent")
        top_tag.pack(fill="x", pady=(0, 6))

        ctk.CTkLabel(
            top_tag,
            text=badge,
            font=(FONT_MONO, 10, "bold"),
            text_color=accent_color,
            anchor="w",
        ).pack(side="left")

        # Title
        ctk.CTkLabel(
            pad,
            text=title,
            font=(FONT_SANS, 20, "bold"),
            text_color=C_TEXT_PRIMARY,
            anchor="w",
        ).pack(fill="x")

        ctk.CTkLabel(
            pad,
            text=subtitle,
            font=(FONT_SANS, 12),
            text_color=C_TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", pady=(2, 10))

        # Description
        ctk.CTkLabel(
            pad,
            text=desc,
            font=(FONT_SANS, 11),
            text_color=C_TEXT_DIM,
            anchor="w",
            justify="left",
            wraplength=420,
        ).pack(fill="x", pady=(0, 16))

        # Specs Box
        spec_box = ctk.CTkFrame(
            pad,
            fg_color="#090E1A",
            corner_radius=4,
            border_width=1,
            border_color=C_BORDER,
        )
        spec_box.pack(fill="x", expand=True, pady=(0, 18))

        for idx, (label, val) in enumerate(specs):
            row = ctk.CTkFrame(spec_box, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=4 if idx > 0 else (8, 4))
            ctk.CTkLabel(
                row,
                text=label,
                font=(FONT_MONO, 9),
                text_color=C_TEXT_DIM,
                anchor="w",
            ).pack(side="left")
            ctk.CTkLabel(
                row,
                text=val,
                font=(FONT_MONO, 9, "bold"),
                text_color=C_TEXT_PRIMARY,
                anchor="e",
            ).pack(side="right")
        ctk.CTkFrame(spec_box, fg_color="transparent", height=4).pack()

        # Primary Launch Button
        ctk.CTkButton(
            pad,
            text=btn_text,
            height=42,
            corner_radius=6,
            font=(FONT_MONO, 11, "bold"),
            fg_color=accent_color,
            hover_color=hover_color,
            text_color="#FFFFFF",
            command=command,
        ).pack(fill="x")

        # Subtle border hover feedback
        def _on_enter(_evt):
            card.configure(border_color=accent_color)

        def _on_leave(_evt):
            card.configure(border_color=C_BORDER)

        card.bind("<Enter>", _on_enter)
        card.bind("<Leave>", _on_leave)

    # ─────────────────────────────────────────────────────────────────
    #  MULTIPROCESS ISOLATION & LIFECYCLE
    # ─────────────────────────────────────────────────────────────────
    def _launch(self, subdir: str, exe_name: str, label: str):
        """Launch child program in an isolated OS process."""
        self._procs = {k: p for k, p in self._procs.items() if p.poll() is None}

        if subdir in self._procs:
            return  # Prevent duplicate instance port collision

        prog_dir = self._root_dir / subdir
        main_py = prog_dir / "main.py"

        names = [exe_name]
        if subdir == "synchro_app":
            names.extend(["VFlap.exe", "Synchro Voltage Measurement System.exe"])
        candidate_exes = []
        for name in dict.fromkeys(names):
            s = Path(name).stem
            candidate_exes.extend([
                self._root_dir / name,
                prog_dir / name,
                prog_dir / "dist" / name,
                prog_dir / "dist" / s / name,
                self._root_dir / "dist" / name,
                self._root_dir / "dist" / s / name,
            ])
        found_exe = next((p for p in candidate_exes if p.is_file()), None)

        cmd = None
        cwd = str(prog_dir) if prog_dir.is_dir() else str(self._root_dir)

        mode = "datalogger" if "datalogger" in subdir else "synchro"

        if getattr(sys, "frozen", False):
            if found_exe:
                cmd = [str(found_exe)]
            else:
                cmd = [sys.executable, "--mode", mode]
        elif main_py.exists():
            venv_py = prog_dir / ".venv" / "Scripts" / "python.exe"
            python_bin = sys.executable
            if venv_py.is_file():
                try:
                    res = subprocess.run([str(venv_py), "-c", "import sys"], capture_output=True, timeout=2)
                    if res.returncode == 0:
                        python_bin = str(venv_py)
                except Exception:
                    pass
            cmd = [python_bin, str(main_py)]
        elif found_exe:
            cmd = [str(found_exe)]
        else:
            cmd = [sys.executable, str(Path(__file__).resolve()), "--mode", mode]

        try:
            flags = 0
            if os.name == "nt":
                flags = subprocess.CREATE_NEW_PROCESS_GROUP
            proc = subprocess.Popen(cmd, cwd=cwd, creationflags=flags)
            self._procs[subdir] = proc
        except Exception as exc:
            from tkinter import messagebox
            messagebox.showerror("Execution Error", f"Failed to launch {label}:\n\n{exc}")

    def _launch_both(self):
        """Simultaneously trigger both instrument subsystems."""
        self._launch("datalogger_repo", "Flap_Data_Logger_GMF.exe", "API 8810A Data Logger")
        self._launch("synchro_app", "VFlap.exe", "Synchro Voltage Monitor")

    def _logout(self):
        self._terminate_children()
        self._show_login()

    def _terminate_children(self):
        for proc in self._procs.values():
            if proc.poll() is None:
                try:
                    proc.terminate()
                    proc.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    try:
                        proc.kill()
                    except OSError:
                        pass
                except OSError:
                    pass
        self._procs.clear()

    def _exit(self):
        self._terminate_children()
        self.destroy()


# ═════════════════════════════════════════════════════════════════════
#  CHILD PROCESS DISPATCH
# ═════════════════════════════════════════════════════════════════════
def _run_child_mode(mode: str, base_dir: Path):
    """Run child application in an isolated subprocess."""
    try:
        if mode == "datalogger":
            target_dir = base_dir / "datalogger_repo"
            if not target_dir.is_dir() and getattr(sys, "_MEIPASS", None):
                target_dir = Path(sys._MEIPASS) / "datalogger_repo"
            os.chdir(str(target_dir))
            sys.path.insert(0, str(target_dir))
            import main as datalogger_main
            app = datalogger_main.FlapSensorApp()
            app.mainloop()
        elif mode == "synchro":
            target_dir = base_dir / "synchro_app"
            if not target_dir.is_dir() and getattr(sys, "_MEIPASS", None):
                target_dir = Path(sys._MEIPASS) / "synchro_app"
            os.chdir(str(target_dir))
            sys.path.insert(0, str(target_dir))
            import main as synchro_main
            app = synchro_main.SynchroApp()
            app.mainloop()
    except Exception as exc:
        import traceback
        err_msg = traceback.format_exc()
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "SyncroTech ODB - Subsystem Error",
                f"Gagal menjalankan subsistem [{mode}]:\n\n{exc}\n\nDetail lengkap disimpan di crash log."
            )
            root.destroy()
        except Exception:
            pass
        try:
            log_path = base_dir / f"{mode}_crash.log"
            log_path.write_text(err_msg, encoding="utf-8")
        except Exception:
            pass
        raise


# ═════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════
if __name__ == "__main__":
    if "--mode" in sys.argv:
        idx = sys.argv.index("--mode")
        if idx + 1 < len(sys.argv):
            child_mode = sys.argv[idx + 1]
            root = (
                Path(sys.executable).resolve().parent
                if getattr(sys, "frozen", False)
                else Path(__file__).resolve().parent
            )
            _run_child_mode(child_mode, root)
            sys.exit(0)

    app = SyncroTechODBApp()
    app.mainloop()
