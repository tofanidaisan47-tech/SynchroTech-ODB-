import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
from PIL import Image
import pyvisa
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
import threading
import time
import os
import random
import math
import sys
import io
import wave
import struct
import socket
import configparser
import json
import urllib.request
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import queue
from decimal import Decimal, ROUND_HALF_UP

try:
    import winsound
except ImportError:
    winsound = None

# ==============================================================================
# GMF AEROASIA THEME & BRANDING COLORS
# ==============================================================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

GMF_BLUE = "#005a9c"
GMF_BLUE_HOVER = "#00467a"
VFD_BG = "#13171f"
VFD_FG = "#f8fafc"
VFD_MUTED = "#64748b"
VFD_ACCENT = "#38bdf8"
# ==============================================================================
# FLAPATE HMI & ROTARY TABLE TRIGGER CONFIGURATION (ISM-7412E @ 192.168.1.20)
# ==============================================================================
DEFAULT_FLAPATE_URL = "http://127.0.0.1:8765"
DEFAULT_FLAPATE_MOTOR_IP = "192.168.1.20"
DEFAULT_FLAPATE_PORT_TCP = 7776
DEFAULT_FLAPATE_PORT_UDP = 7775
DEFAULT_LOCAL_API_PORT = 8766
DEFAULT_OKE_PATH = r"C:\Users\USER\Downloads\OKE\OKE"
DEFAULT_HMI_INI_PATH = os.path.join(DEFAULT_OKE_PATH, "HMI_Param.ini")

TRIGGER_MODE_FLAPATE = "flapate"      # Trigger when FlapATE takes angle data or needle stops
TRIGGER_MODE_DYNAMIX = "flapate"      # Backward compatibility alias

DYN_STATE_READY_TO_SAVE = "ready_to_save"
DYN_STATE_AWAITING_MOVE = "awaiting_move"
DYN_STATE_AWAITING_STOP = "awaiting_stop"
FLAPATE_SETTLE_TIME_FAST = 0.2        # Fast settling (0.2s): captures voltage instantly before FlapATE auto-resumes
FLAPATE_SETTLE_TIME_NORMAL = 0.8      # Normal settling (0.8s)
DYNAMIX_SETTLE_TIME = FLAPATE_SETTLE_TIME_FAST   # Backward compatibility alias
DEFAULT_DYNAMIX_IP = DEFAULT_FLAPATE_MOTOR_IP
DEFAULT_DYNAMIX_PORT_TCP = DEFAULT_FLAPATE_PORT_TCP
DEFAULT_DYNAMIX_PORT_UDP = DEFAULT_FLAPATE_PORT_UDP

# Specific Sheet Mapping Configurations for GMF Excel Workbook
SHEET_CONFIGS = {
    "ACCURACY STALL": {
        "start_row": 13,
        "total_rows": 11,
        "tp_col": 1,
        "sp_col": 2,                                  # Col B
        "has_z": True,
        "test_cols": {"X": 6, "Y": 7, "Z": 8},        # Col F, G, H
        "ref_cols": {"X": 3, "Y": 4, "Z": 5},         # Col C, D, E
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "ACCURACY AUTOPILOT": {
        "start_row": 12,
        "total_rows": 11,
        "tp_col": 1,
        "sp_col": 3,                                  # Col C
        "has_z": True,
        "test_cols": {"X": 7, "Y": 8, "Z": 9},        # Col G, H, I
        "ref_cols": {"X": 4, "Y": 5, "Z": 6},         # Col D, E, F
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "ACCURACY AUTO PILOT": {
        "start_row": 12,
        "total_rows": 11,
        "tp_col": 1,
        "sp_col": 3,                                  # Col C
        "has_z": True,
        "test_cols": {"X": 7, "Y": 8, "Z": 9},        # Col G, H, I
        "ref_cols": {"X": 4, "Y": 5, "Z": 6},         # Col D, E, F
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "ACCURACY FLAP": {
        "start_row": 12,
        "total_rows": 11,
        "tp_col": 3,                                  # Col C
        "sp_col": 4,                                  # Col D
        "has_z": False,                               # Flap measures X and Y only!
        "test_cols": {"X": 7, "Y": 8},                # Col G, H
        "ref_cols": {"X": 5, "Y": 6},                 # Col E, F
        "channels": ["X to GND", "Y to GND"]
    },
    "CAM OUTPUT STALL": {
        "start_row": 11,
        "total_rows": 182,
        "tp_col": 1,                                  # Col A
        "sp_col": 3,                                  # Col C (Shaft Angle)
        "has_z": True,
        "test_cols": {"X": 7, "Y": 8, "Z": 9},        # Col G, H, I
        "ref_cols": {"X": 4, "Y": 5, "Z": 6},         # Col D, E, F
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "CAM OUTPUT AUTOPILOT": {
        "start_row": 11,
        "total_rows": 182,
        "tp_col": 1,                                  # Col A
        "sp_col": 3,                                  # Col C (Shaft Angle)
        "has_z": True,
        "test_cols": {"X": 7, "Y": 8, "Z": 9},        # Col G, H, I
        "ref_cols": {"X": 4, "Y": 5, "Z": 6},         # Col D, E, F
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "CAM OUTPUT AUTO PILOT": {
        "start_row": 11,
        "total_rows": 182,
        "tp_col": 1,                                  # Col A
        "sp_col": 3,                                  # Col C (Shaft Angle)
        "has_z": True,
        "test_cols": {"X": 7, "Y": 8, "Z": 9},        # Col G, H, I
        "ref_cols": {"X": 4, "Y": 5, "Z": 6},         # Col D, E, F
        "channels": ["X to GND", "Y to GND", "Z to GND"]
    },
    "CAM OUTPUT FLAP": {
        "start_row": 11,
        "total_rows": 182,
        "tp_col": 1,                                  # Col A
        "sp_col": 3,                                  # Col C (Shaft Angle)
        "has_z": False,
        "test_cols": {"X": 6, "Y": 7},                # Col F, G
        "ref_cols": {"X": 4, "Y": 5},                 # Col D, E
        "channels": ["X to GND", "Y to GND"]
    }
}

# 11 Default Test Points (Standard SP Degrees)
DEFAULT_SP_DEGREES = [277, 270, 240, 210, 180, 150, 113, 76, 37, 0, -6]


def generate_blink_sound_wav(sample_rate=44100):
    """
    Generate in-memory WAV audio for normal confirmation chirp (OK / Green / Valid).
    Two-tone cheerful high chime (A6: 1760 Hz -> E7: 2637 Hz) with smooth envelope as in SyncroTech_ODB.
    """
    duration = 0.16
    num_samples = int(sample_rate * duration)
    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        frames = bytearray()
        t1 = int(0.07 * sample_rate)
        t2 = num_samples - t1
        # Tone 1: 1760 Hz
        for i in range(t1):
            env = math.sin(math.pi * (i / t1)) ** 0.8
            s = int(0.80 * env * math.sin(2.0 * math.pi * 1760.0 * (i / sample_rate)) * 32767)
            frames.extend(struct.pack("<h", max(-32767, min(32767, s))))
        # Tone 2: 2637 Hz (higher pleasant resolve)
        for i in range(t2):
            env = math.sin(math.pi * (i / t2)) ** 0.8
            s = int(0.85 * env * math.sin(2.0 * math.pi * 2637.0 * (i / sample_rate)) * 32767)
            frames.extend(struct.pack("<h", max(-32767, min(32767, s))))
        wf.writeframes(frames)
    return buffer.getvalue()


def generate_warn_sound_wav(sample_rate=44100):
    """
    Generate in-memory WAV audio alert for invalid/deviation warning (Yellow / Caution / Invalid).
    High-visibility acoustic alert (780 Hz + harmonics -> 580 Hz + harmonics) as in SyncroTech_ODB.
    """
    duration = 0.32
    num_samples = int(sample_rate * duration)
    buffer = io.BytesIO()
    with wave.open(buffer, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        frames = bytearray()
        pulse_len = int(0.14 * sample_rate)
        gap_len = int(0.04 * sample_rate)

        # Pulse 1: 780 Hz (aviation warning horn tone) with rich 2nd & 3rd harmonics
        for i in range(pulse_len):
            t = i / sample_rate
            env = min(1.0, 15.0 * (i / pulse_len), 15.0 * ((pulse_len - i) / pulse_len))
            sig = (
                0.60 * math.sin(2.0 * math.pi * 780.0 * t) +
                0.28 * math.sin(2.0 * math.pi * 1560.0 * t) +
                0.12 * math.sin(2.0 * math.pi * 2340.0 * t)
            )
            val = int(0.98 * env * sig * 32767)
            frames.extend(struct.pack("<h", max(-32767, min(32767, val))))

        # Gap between pulses
        for _ in range(gap_len):
            frames.extend(struct.pack("<h", 0))

        # Pulse 2: 580 Hz (descending caution drop) with rich harmonics
        for i in range(pulse_len):
            t = i / sample_rate
            env = min(1.0, 15.0 * (i / pulse_len), 15.0 * ((pulse_len - i) / pulse_len))
            sig = (
                0.62 * math.sin(2.0 * math.pi * 580.0 * t) +
                0.26 * math.sin(2.0 * math.pi * 1160.0 * t) +
                0.12 * math.sin(2.0 * math.pi * 1740.0 * t)
            )
            val = int(0.98 * env * sig * 32767)
            frames.extend(struct.pack("<h", max(-32767, min(32767, val))))

        wf.writeframes(frames)
    return buffer.getvalue()


def get_3rd_decimal_digit(val):
    """Get the 3rd digit after the decimal point (0-9) from a numerical value."""
    try:
        val_abs = abs(float(val))
        s = f"{val_abs:.6f}"
        decimals = s.split(".")[1]
        return int(decimals[2]) if len(decimals) > 2 else 0
    except Exception:
        return -1


def extract_2_decimals(val):
    """
    Ambil 2 angka di belakang koma murni TANPA pembulatan (truncation).
    Contoh: 7.818 -> 7.81, 5.216 -> 5.21, 12.969 -> 12.96.
    """
    if val is None or val == "-":
        return "-"
    try:
        if isinstance(val, str):
            s_clean = val.replace("VAC", "").replace("mVAC", "").replace("V", "").strip().replace(",", ".")
            val_f = float(s_clean)
        else:
            val_f = float(val)

        is_neg = val_f < 0
        val_abs = abs(val_f)
        s = f"{val_abs:.6f}"
        parts = s.split(".")
        int_part = parts[0]
        decimals = parts[1]
        res = float(f"{int_part}.{decimals[:2]}")
        return -res if is_neg else res
    except (ValueError, TypeError):
        return val


def format_2_decimals(val):
    """Format string 2 angka di belakang koma murni TANPA pembulatan."""
    if val is None or val == "-":
        return "-"
    try:
        if isinstance(val, str):
            s_clean = val.replace("VAC", "").replace("mVAC", "").replace("V", "").strip().replace(",", ".")
            val_f = float(s_clean)
        else:
            val_f = float(val)

        is_neg = val_f < 0
        s = f"{abs(val_f):.6f}"
        parts = s.split(".")
        return f"{'-' if is_neg else ''}{parts[0]}.{parts[1][:2]}"
    except (ValueError, TypeError):
        return str(val)


def get_resource_path(relative_path):
    """Get absolute path to resource, compatible with PyInstaller and development mode."""
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    res_path = os.path.join(base_path, relative_path)
    if not os.path.exists(res_path) and getattr(sys, "frozen", False):
        exe_dir_path = os.path.join(os.path.dirname(sys.executable), relative_path)
        if os.path.exists(exe_dir_path):
            return exe_dir_path
    return res_path


# ==============================================================================
# LOCAL REST API / WEBHOOK SERVER FOR FLAPATE INTEGRATION
# ==============================================================================
class LocalAPIHandler(BaseHTTPRequestHandler):
    """
    HTTP Request Handler for FlapATE Webhook & External API integration.
    Endpoints:
      - POST /api/trigger or GET /api/trigger: Triggers voltage capture & row advance
      - GET /api/voltage: Returns current live multimeter voltage
      - GET /api/status: Returns full app state (channel, row, connection, etc.)
      - POST /api/step: Sets active row / test point
      - POST /api/channel: Sets active channel (X, Y, Z)
    """
    def log_message(self, format, *args):
        # Suppress noisy standard HTTP access logs
        pass

    def _send_json(self, status_code, data):
        try:
            self.send_response(status_code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.end_headers()
            self.wfile.write(json.dumps(data).encode("utf-8"))
        except Exception:
            pass

    def do_OPTIONS(self):
        self._send_json(200, {"status": "ok"})

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        app = getattr(self.server, "app", None)
        if not app:
            self._send_json(500, {"error": "App reference missing"})
            return

        trigger_paths = (
            "/api/trigger", "/trigger", "/api/measure", "/measure",
            "/api/record", "/record", "/api/result", "/result",
            "/api/data", "/data", "/api/fail", "/fail",
            "/api/pass", "/pass", "/api/save", "/save", "/api/sample"
        )

        if path in trigger_paths:
            # Query params can specify step or angle: e.g. /api/trigger?step=5&angle=30.0&status=FAIL
            qs = urllib.parse.parse_qs(parsed.query)
            q_step = qs.get("step", [None])[0] or qs.get("row", [None])[0] or qs.get("point", [None])[0]
            q_angle = qs.get("angle", [None])[0] or qs.get("sensing", [None])[0] or qs.get("sp", [None])[0]
            q_status = qs.get("status", [None])[0] or qs.get("result", [None])[0] or ""
            source_tag = f"HTTP GET {path}"
            if q_status:
                source_tag += f" [{q_status.upper()}]"
            res = app.trigger_from_api(source=source_tag, angle=q_angle, step=q_step)
            self._send_json(200, res)
        elif path in ("/api/voltage", "/voltage"):
            self._send_json(200, {
                "voltage": app.current_live_voltage,
                "channel": app.active_channel,
                "row": app.selected_row_idx,
                "status": "connected" if app.is_connected else "disconnected"
            })
        elif path in ("/api/status", "/status"):
            self._send_json(200, {
                "app": "VFlap - Synchro Voltage Measurement System",
                "integration": "FlapATE",
                "voltage": app.current_live_voltage,
                "channel": app.active_channel,
                "row": app.selected_row_idx,
                "total_rows": app.total_rows,
                "sheet": app.sheet_name,
                "connected": app.is_connected,
                "direction": getattr(app, "input_direction", "forward")
            })
        else:
            self._send_json(404, {"error": "Not Found", "available_endpoints": ["/api/trigger", "/api/voltage", "/api/status", "/api/direction", "/api/channel", "/api/step"]})

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        app = getattr(self.server, "app", None)
        if not app:
            self._send_json(500, {"error": "App reference missing"})
            return

        content_length = int(self.headers.get("Content-Length", 0))
        body = {}
        if content_length > 0:
            try:
                body_bytes = self.rfile.read(content_length)
                raw_str = body_bytes.decode("utf-8", errors="ignore").strip()
                if raw_str.startswith("{") or raw_str.startswith("["):
                    body = json.loads(raw_str)
                else:
                    # Support application/x-www-form-urlencoded payloads
                    qs = urllib.parse.parse_qs(raw_str)
                    body = {k: v[0] if len(v) == 1 else v for k, v in qs.items()}
            except Exception:
                body = {}

        trigger_paths = (
            "/api/trigger", "/trigger", "/api/measure", "/measure",
            "/api/record", "/record", "/api/result", "/result",
            "/api/data", "/data", "/api/fail", "/fail",
            "/api/pass", "/pass", "/api/save", "/save", "/api/sample"
        )

        if path in trigger_paths:
            target_row = (
                body.get("row") or body.get("step") or 
                body.get("step_index") or body.get("no") or body.get("point")
            )
            if target_row is not None:
                try:
                    target_row_idx = int(target_row)
                    if 1 <= target_row_idx <= app.total_rows and target_row_idx != app.selected_row_idx:
                        app.after(0, lambda r=target_row_idx: app.select_row(r, from_auto_advance=False))
                except Exception:
                    pass

            target_angle = (
                body.get("angle") or body.get("sensing") or 
                body.get("sp") or body.get("synchro") or body.get("actual") or body.get("value")
            )
            status_tag = body.get("status") or body.get("result") or body.get("judgment") or body.get("judge") or ""
            source_info = f"FlapATE Webhook ({path})"
            if status_tag:
                source_info += f" [{str(status_tag).upper()}]"
            if target_angle is not None:
                try:
                    source_info += f" (Angle {float(target_angle):.2f}°)"
                except Exception:
                    source_info += f" (Angle: {target_angle})"

            # Snapshot current state for immediate response
            val = app.current_live_voltage
            val_2dec = extract_2_decimals(val)
            row_idx = app.selected_row_idx
            ch = app.active_channel

            # Safely schedule measurement and GUI update on Tkinter main thread
            app.run_on_gui_thread(app.trigger_from_api, source=source_info, angle=target_angle, step=target_row)
            self._send_json(200, {
                "success": True,
                "status": "triggered",
                "row": row_idx,
                "channel": ch,
                "voltage": val_2dec,
                "angle": target_angle,
                "source": source_info
            })

        elif path in ("/api/step", "/api/row"):
            new_row = body.get("row") or body.get("step")
            if new_row is not None:
                try:
                    app.run_on_gui_thread(app.select_row, int(new_row), from_auto_advance=False)
                    self._send_json(200, {"success": True, "selected_row": int(new_row)})
                    return
                except Exception as e:
                    self._send_json(400, {"error": str(e)})
                    return
            self._send_json(400, {"error": "Missing 'row' in body"})

        elif path in ("/api/channel",):
            ch = body.get("channel", "").upper()
            if ch in ("X", "Y", "Z"):
                app.run_on_gui_thread(app.on_channel_focus_changed, f"{ch} to GND")
                self._send_json(200, {"success": True, "channel": ch})
            else:
                self._send_json(400, {"error": "Invalid channel (must be X, Y, or Z)"})

        elif path in ("/api/direction",):
            d = body.get("direction", "").lower()
            if d in ("forward", "reverse"):
                app.run_on_gui_thread(app.set_direction, d)
                self._send_json(200, {"success": True, "direction": d})
            else:
                self._send_json(400, {"error": "Invalid direction (must be 'forward' or 'reverse')"})
        else:
            self._send_json(404, {"error": "Endpoint not found"})


class LocalAPIServer:
    def __init__(self, app, port=DEFAULT_LOCAL_API_PORT):
        self.app = app
        self.port = port
        self.server = None
        self.thread = None
        self.is_running = False

    def start(self):
        for try_port in [self.port, 8767, 8088, 8089]:
            try:
                self.port = try_port
                self.server = HTTPServer(("0.0.0.0", self.port), LocalAPIHandler)
                self.server.app = self.app
                self.is_running = True
                self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
                self.thread.start()
                return True, f"http://127.0.0.1:{self.port}"
            except Exception:
                continue
        return False, "Failed to bind local API port"

    def stop(self):
        if self.server:
            try:
                self.server.shutdown()
                self.server.server_close()
            except Exception:
                pass
            self.is_running = False


# ==============================================================================
# FLAPATE HMI MONITOR & ROTARY TABLE SCL CLIENT (ISM-7412E @ 192.168.1.20)
# ==============================================================================
class FlapATEClient:
    """
    Client and monitor for FlapATE HMI (http://127.0.0.1:8765) and NI ISM-7412E Stepper Motor at IP 192.168.1.20.
    Detects when FlapATE takes angle data or when the needle stops:
      1. Web/REST API monitoring to FlapATE HMI at http://127.0.0.1:8765:
         - Probes FlapATE API endpoints for result records, sensing angles, and sequence steps.
         - Detects when FlapATE inputs a new angle data point.
      2. Direct hardware socket to motor at IP 192.168.1.20 (port 7776 TCP / 7775 UDP SCL):
         - Queries Status Code ('SC\\r') -> Bit 4 is Moving, Bit 3 is Target In Position.
         - Queries Immediate Velocity ('IV\\r') -> velocity > 0 is Moving.
         - Fast settling detection so voltage is captured instantly when needle stops,
           before FlapATE automatically resumes motor motion!
      3. OKE INI file monitoring ('HMI_Param.ini') if present.
      4. Integrated simulation mode for offline testing.
    """
    def __init__(self, url=DEFAULT_FLAPATE_URL, ip=DEFAULT_FLAPATE_MOTOR_IP, tcp_port=DEFAULT_FLAPATE_PORT_TCP,
                 udp_port=DEFAULT_FLAPATE_PORT_UDP, ini_path=DEFAULT_HMI_INI_PATH, on_trigger_callback=None):
        self.api_url = url
        self.ip = ip
        self.tcp_port = tcp_port
        self.udp_port = udp_port
        self.ini_path = ini_path
        self.on_trigger_callback = on_trigger_callback

        self.is_connected = True
        self.is_lan_connected = False
        self.is_moving = False
        self.current_angle = 0.0
        self.current_rpm = 0.0
        self.connection_source = "Auto-Tracking"
        self.is_flapate_api_online = False

        self.last_state_change = time.time()
        self.last_stop_time = time.time()
        self.last_move_time = 0.0

        self._sim_moving = False
        self._sim_move_end_time = 0.0
        self._lock = threading.Lock()
        self._prev_ini_angle = None
        self._prev_api_record_count = None
        self._prev_api_sensing_angle = None
        self._prev_api_step = None
        self._prev_api_status = None
        self._running = True

        # Start non-blocking background polling thread
        self._bg_thread = threading.Thread(target=self._bg_polling_loop, daemon=True)
        self._bg_thread.start()

    def _bg_polling_loop(self):
        """Asynchronously query FlapATE API, motor hardware socket, and INI in background."""
        while self._running:
            try:
                # 1. Query FlapATE HTTP API
                api_ok, api_trig, api_angle, api_step = self._query_flapate_api()
                if api_ok:
                    self.is_flapate_api_online = True
                    self.is_lan_connected = True
                    if api_trig and self.on_trigger_callback:
                        self.on_trigger_callback(source="FlapATE API", angle=api_angle, step=api_step)

                # 2. Query Motor Hardware Socket (UDP / TCP)
                hw_ok, hw_moving, hw_angle, hw_rpm = self._query_hardware_socket()
                if hw_ok:
                    with self._lock:
                        self.is_lan_connected = True
                        src = f"FlapATE Motor ({self.ip})"
                        if self.is_flapate_api_online:
                            src += " + API Online"
                        self._update_motion_state(hw_moving, hw_angle, hw_rpm, src)
                    time.sleep(0.08)
                    continue

                # 3. Try INI file query
                ini_ok, ini_moving, ini_angle, ini_rpm = self._query_ini_file()
                if ini_ok:
                    with self._lock:
                        self.is_lan_connected = True
                        self._update_motion_state(ini_moving, ini_angle, ini_rpm, "FlapATE OKE Monitor")
                    time.sleep(0.10)
                    continue

                # 4. Check ping / port connectivity
                ping_ok = self._quick_ping()
                if ping_ok:
                    with self._lock:
                        self.is_connected = True
                        self.is_lan_connected = True
                        src = f"Ethernet Host ({self.ip})"
                        if self.is_flapate_api_online:
                            src += " + FlapATE API"
                        self.connection_source = src
                else:
                    with self._lock:
                        self.is_connected = True
                        self.is_lan_connected = False
                        self.connection_source = "Auto-Tracking (GPIB Only)"
            except Exception:
                pass
            time.sleep(0.12)

    def _query_flapate_api(self):
        """
        Poll FlapATE web server / API at self.api_url (e.g. http://127.0.0.1:8765 or LAN IP).
        Returns (api_ok, is_new_trigger, sensing_angle, step_index).
        """
        base_urls = []
        if self.api_url:
            base_urls.append(self.api_url.rstrip("/"))
        if self.ip and self.ip not in ("127.0.0.1", "localhost"):
            lan_url = f"http://{self.ip}:8765"
            if lan_url not in base_urls:
                base_urls.append(lan_url)

        endpoints = []
        for base_url in base_urls:
            endpoints.extend([
                f"{base_url}/api/record",
                f"{base_url}/api/status",
                f"{base_url}/api/data",
                f"{base_url}/api/result",
                f"{base_url}/api/sequence",
                f"{base_url}/api/synchro",
                base_url
            ])

        for ep in endpoints:
            try:
                req = urllib.request.Request(ep, headers={"User-Agent": "SynchroMeasurement/2.0"})
                with urllib.request.urlopen(req, timeout=0.15) as resp:
                    raw_data = resp.read()
                    content_type = resp.headers.get("Content-Type", "")

                    if "json" in content_type.lower() or raw_data.strip().startswith(b"{") or raw_data.strip().startswith(b"["):
                        try:
                            parsed_json = json.loads(raw_data.decode("utf-8"))
                            is_trig, angle, step = self._parse_flapate_json(parsed_json)
                            return True, is_trig, angle, step
                        except Exception:
                            pass
                    else:
                        is_trig, angle, step = self._parse_flapate_html(raw_data)
                        return True, is_trig, angle, step
            except Exception:
                continue

        return False, False, None, None

    def _parse_flapate_json(self, data):
        """Analyze FlapATE JSON response for new records, step progression, or sensing angle changes."""
        is_trig = False
        angle = None
        step = None

        if isinstance(data, dict):
            records = data.get("records") or data.get("results") or data.get("data")
            if isinstance(records, list):
                rec_len = len(records)
                if self._prev_api_record_count is not None and rec_len > self._prev_api_record_count:
                    is_trig = True
                    if records:
                        last_rec = records[-1]
                        if isinstance(last_rec, dict):
                            angle = last_rec.get("angle") or last_rec.get("sensing") or last_rec.get("synchro")
                            step = last_rec.get("step") or last_rec.get("no")
                self._prev_api_record_count = rec_len

            sensing = data.get("synchro") or data.get("sensing") or data.get("angle") or data.get("actual") or data.get("value")
            if sensing is not None:
                try:
                    f_sensing = float(sensing)
                    if self._prev_api_sensing_angle is not None and abs(f_sensing - self._prev_api_sensing_angle) > 0.05:
                        is_trig = True
                        angle = f_sensing
                    self._prev_api_sensing_angle = f_sensing
                except Exception:
                    pass

            curr_step = data.get("current_step") or data.get("step") or data.get("step_index") or data.get("no") or data.get("point")
            if curr_step is not None:
                try:
                    i_step = int(curr_step)
                    if getattr(self, "_prev_api_step", None) is not None and i_step != self._prev_api_step:
                        is_trig = True
                        step = i_step
                    self._prev_api_step = i_step
                except Exception:
                    pass

            status_val = data.get("status") or data.get("result") or data.get("judgment") or data.get("judge")
            if status_val is not None:
                s_str = str(status_val).upper().strip()
                status_step_key = f"{s_str}_{curr_step}"
                if s_str in ("PASS", "FAIL", "OK", "NG", "DONE") and status_step_key != getattr(self, "_prev_status_step_key", None):
                    is_trig = True
                    self._prev_status_step_key = status_step_key
                self._prev_api_status = s_str

        elif isinstance(data, list):
            rec_len = len(data)
            if self._prev_api_record_count is not None and rec_len > self._prev_api_record_count:
                is_trig = True
                if data:
                    last_rec = data[-1]
                    if isinstance(last_rec, dict):
                        angle = last_rec.get("angle") or last_rec.get("sensing") or last_rec.get("synchro")
                        step = last_rec.get("step") or last_rec.get("no")
            self._prev_api_record_count = rec_len

        return is_trig, angle, step

    def _parse_flapate_html(self, raw_bytes):
        """Fallback parser if FlapATE returns HTML containing Result Record table or Synchro angle."""
        text = raw_bytes.decode("utf-8", errors="ignore")
        if "FlapATE" in text or "Rotary table" in text or "Synchro" in text:
            import re
            m = re.search(r"Synchro[:\s]+([\d\.\-]+)\s*Deg", text, re.IGNORECASE)
            if m:
                try:
                    angle_val = float(m.group(1))
                    if self._prev_api_sensing_angle is not None and abs(angle_val - self._prev_api_sensing_angle) > 0.05:
                        self._prev_api_sensing_angle = angle_val
                        return True, angle_val, None
                    self._prev_api_sensing_angle = angle_val
                except Exception:
                    pass
            return False, None, None
        return False, None, None

    def test_api_connection(self):
        """Quickly test if FlapATE API is reachable."""
        try:
            req = urllib.request.Request(self.api_url, headers={"User-Agent": "SynchroMeasurement/2.0"})
            with urllib.request.urlopen(req, timeout=0.6) as resp:
                self.is_flapate_api_online = (resp.status in (200, 204, 301, 302))
                return self.is_flapate_api_online
        except Exception:
            self.is_flapate_api_online = False
            return False

    def _quick_ping(self):
        for port in [self.tcp_port, 8765, 80]:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(0.08)
                res = s.connect_ex((self.ip, port))
                s.close()
                if res == 0:
                    return True
            except Exception:
                pass
        return False

    def poll(self, is_sim_mode=False):
        """
        Poll cached hardware socket, HMI_Param.ini, or simulation state.
        Instantaneous return (non-blocking).
        Returns (is_connected, is_moving, current_angle, current_rpm).
        """
        now = time.time()

        if is_sim_mode:
            with self._lock:
                if self._sim_moving:
                    if now >= self._sim_move_end_time:
                        self._sim_moving = False
                        self.is_moving = False
                        self.current_rpm = 0.0
                        self.last_stop_time = now
                    else:
                        self.is_moving = True
                        self.current_rpm = 45.0
                        self.last_move_time = now
                else:
                    self.is_moving = False
                    self.current_rpm = 0.0
                self.is_connected = True
                self.is_lan_connected = True
                self.connection_source = "Demo Simulation"
                return self.is_connected, self.is_moving, self.current_angle, self.current_rpm

        with self._lock:
            return self.is_connected, self.is_moving, self.current_angle, self.current_rpm

    def _update_motion_state(self, moving, angle, rpm, source):
        now = time.time()
        if angle is not None:
            self.current_angle = angle
        if rpm is not None:
            self.current_rpm = rpm
        self.is_connected = True
        self.connection_source = source

        if moving != self.is_moving:
            self.is_moving = moving
            self.last_state_change = now
            if moving:
                self.last_move_time = now
            else:
                self.last_stop_time = now

    def _parse_scl_response(self, data):
        if not data:
            return False, False, None, 0.0
        text = data.decode("ascii", errors="ignore").strip()
        clean_ascii = "".join([c for c in text if 32 <= ord(c) <= 126])

        # 1. Check for velocity response (IV=...)
        if "IV=" in clean_ascii:
            try:
                vel_str = clean_ascii.split("IV=")[-1].strip()
                rpm = abs(float(vel_str))
                is_moving = (rpm > 0.05)
                return True, is_moving, None, rpm
            except Exception:
                pass

        # 2. Check for position response (IP=...)
        if "IP=" in clean_ascii:
            try:
                pos_str = clean_ascii.split("IP=")[-1].strip()
                deg = float(pos_str)
                return True, False, deg, 0.0
            except Exception:
                pass

        # 3. Check for status code (SC=... or raw 4-digit hex)
        clean_hex = clean_ascii.split("=")[-1].strip()
        clean_hex = "".join([c for c in clean_hex if c in "0123456789abcdefABCDEF"])
        if len(clean_hex) >= 4:
            try:
                code = int(clean_hex[-4:], 16)
                moving = bool(code & 0x0010)   # Bit 4: Motor Moving
                in_pos = bool(code & 0x0008)   # Bit 3: Target Position Reached
                is_moving = moving and not in_pos
                return True, is_moving, None, 0.0
            except Exception:
                pass

        return False, False, None, 0.0

    def _query_hardware_socket(self):
        for cmd in [b"\x00\x07SC\r", b"\x07SC\r", b"SC\r", b"\x00\x07IV\r", b"IV\r"]:
            sock = None
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                sock.settimeout(0.08)
                sock.sendto(cmd, (self.ip, self.udp_port))
                data, _ = sock.recvfrom(512)
                ok, is_moving, angle, rpm = self._parse_scl_response(data)
                if ok:
                    return True, is_moving, angle, rpm
            except Exception:
                pass
            finally:
                if sock:
                    try:
                        sock.close()
                    except Exception:
                        pass

        for cmd in [b"SC\r", b"\x00\x07SC\r", b"IV\r"]:
            sock = None
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(0.08)
                sock.connect((self.ip, self.tcp_port))
                sock.sendall(cmd)
                data = sock.recv(512)
                ok, is_moving, angle, rpm = self._parse_scl_response(data)
                if ok:
                    return True, is_moving, angle, rpm
            except Exception:
                pass
            finally:
                if sock:
                    try:
                        sock.close()
                    except Exception:
                        pass

        return False, False, None, 0.0

    def _query_ini_file(self):
        search_paths = [
            self.ini_path,
            os.path.join(DEFAULT_OKE_PATH, "HMI_Param.ini"),
            os.path.join(DEFAULT_OKE_PATH, "Data", "HMI_Param.ini"),
            os.path.join(os.getcwd(), "HMI_Param.ini"),
        ]
        target_path = next((p for p in search_paths if os.path.exists(p)), None)
        if not target_path:
            return False, False, None, 0.0

        try:
            cp = configparser.ConfigParser()
            cp.read(target_path)
            if cp.has_section("MONITOR"):
                spin = cp.get("MONITOR", "SPIN", fallback="FALSE").strip().upper()
                speed_str = cp.get("MONITOR", "SPEED.SENSE", fallback="0.0").strip()
                pos_str = cp.get("MONITOR", "POS.SENSE", fallback="0.0").strip()
                rpm = abs(float(speed_str))
                angle = float(pos_str)
                pos_delta = abs(angle - self._prev_ini_angle) if self._prev_ini_angle is not None else 0.0
                self._prev_ini_angle = angle
                is_moving = (spin == "TRUE") or (rpm > 0.02) or (pos_delta > 0.05)
                return True, is_moving, angle, rpm
        except Exception:
            pass

        return False, False, None, 0.0

    def trigger_simulation_move(self, target_angle=None, duration=0.8):
        """Simulate rotary table needle moving to target angle and then stopping with FlapATE capture."""
        with self._lock:
            self._sim_moving = True
            self._sim_move_end_time = time.time() + duration
            if target_angle is not None:
                self.current_angle = float(target_angle)
            self.is_moving = True
            self.last_move_time = time.time()


# Backward compatibility alias
DynamixPositionClient = FlapATEClient


class SynchroApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Main Window Configuration
        self.title("VFlap - Synchro Voltage Measurement System | GMF AeroAsia")
        self.geometry("1420x850")
        self.minsize(1220, 740)

        # Event loop recursion guard
        self._ignore_tree_event = False

        # Hardware VISA / Multimeter Communication
        self.visa_rm = None
        self.multimeter = None
        self.is_connected = False
        self.sim_mode = False

        # In a PyInstaller onedir build, the executable is nested under dist.
        # Keep the editable workbook in the project folder instead of copying it
        # into the generated application directory.
        excel_filename = "2708 Voltase Accuracy STALL, AUTOPILOT, FLAP.xlsx"
        if getattr(sys, "frozen", False):
            executable_dir = os.path.dirname(sys.executable)
            excel_search_dirs = (
                executable_dir,
                os.path.dirname(executable_dir),
                os.path.dirname(os.path.dirname(executable_dir)),
            )
        else:
            excel_search_dirs = (os.path.dirname(os.path.abspath(__file__)),)

        self.excel_file_path = next(
            (
                os.path.join(directory, excel_filename)
                for directory in excel_search_dirs
                if os.path.isfile(os.path.join(directory, excel_filename))
            ),
            "",
        )

        # Default Active Sheet: ACCURACY STALL
        self.sheet_name = "ACCURACY STALL"
        self.cur_cfg = SHEET_CONFIGS["ACCURACY STALL"]
        self.start_row = self.cur_cfg["start_row"]
        self.has_z = self.cur_cfg["has_z"]
        self.test_cols = self.cur_cfg["test_cols"]

        # Channel & Row Focus Measurement Status
        self.active_channel = "X"
        self.selected_row_idx = 1
        self.total_rows = self.cur_cfg.get("total_rows", 11)

        # Reading & Stability Variables (Stability Monitor 2 Decimal Places)
        self.current_live_voltage = 0.0
        self.is_value_stable = False
        self.stable_duration = 0.0

        # Stability tracking for 2 decimal places over 1.6 seconds (Voltage Stability mode)
        self.history_2dec = []
        self.stable_candidate_val = None
        self.auto_saved_this_row = False
        self.auto_save_enabled = True  # Mode Otomatis aktif secara default
        self.is_channel_armed = True
        self.channel_transition_active = False
        self.awaiting_movement = False
        self._flapate_is_moving = False
        self._flapate_stable_start = 0.0
        self._flapate_stable_raw = None
        self.last_saved_val_2dec = None
        self.last_saved_raw = None
        self.row_select_time = time.time()
        self._cell_highlight_labels = []
        self.vfd_bg_color = VFD_BG

        # FlapATE HMI Client & Rotary Table State Configuration
        self.flapate_url = DEFAULT_FLAPATE_URL
        self.flapate_motor_ip = DEFAULT_FLAPATE_MOTOR_IP
        self.flapate_settle_time = FLAPATE_SETTLE_TIME_FAST
        self.flapate_client = FlapATEClient(
            url=self.flapate_url,
            ip=self.flapate_motor_ip,
            ini_path=DEFAULT_HMI_INI_PATH,
            on_trigger_callback=self._on_flapate_api_trigger
        )
        # Backward compatibility aliases
        self.dynamix_ip = self.flapate_motor_ip
        self.dynamix_client = self.flapate_client
        self.trigger_mode = TRIGGER_MODE_FLAPATE
        self.dynamix_state = DYN_STATE_READY_TO_SAVE
        self.dynamix_needle_stop_time = time.time()

        # Start Local REST API / Webhook Server (Port 8766)
        self.local_api_server = LocalAPIServer(app=self, port=DEFAULT_LOCAL_API_PORT)
        api_started, api_endpoint = self.local_api_server.start()
        self.local_api_url = api_endpoint if api_started else "http://127.0.0.1:8766"

        # Input Direction Mode: "forward" (Forward: TP 1 -> Last) or "reverse" (Reverse: TP Last -> 1)
        self.input_direction = "forward"

        # Audio Feedback Variables: 'Blink' (Normal OK) & 'Warning' (Yellow Deviation) on Data Input
        self.sound_enabled = True
        self._blink_wav = None
        self._warn_wav = None
        try:
            self._blink_wav = generate_blink_sound_wav()
            self._warn_wav = generate_warn_sound_wav()
        except Exception as e:
            print(f"Sound generation error: {e}")

        # Polling background thread
        self.polling_active = True
        self.poll_thread = None

        # Thread-safe Asynchronous Excel Writer Queue (guarantees zero UI lag/freezes)
        self.excel_queue = queue.Queue()
        self.excel_worker = threading.Thread(target=self._excel_writer_worker, daemon=True)
        self.excel_worker.start()

        # Thread-safe GUI task queue for background threads (FlapATE API, Webhook server, etc.)
        self.gui_queue = queue.Queue()
        self._check_gui_queue()

        # UI state caches for high performance / lightweight rendering
        self._last_vfd_digits = None
        self._last_vfd_chan = None
        self._last_flapate_ui_state = None

        # Measurement Table Data Structure
        self.table_data = {}
        self.init_table_data()

        # Build GUI
        self.create_header()
        self.create_main_layout()

        # Keyboard Shortcut Binding (ENTER to Save Reading)
        self.bind("<Return>", lambda event: self.manual_save_reading())
        self.bind("<KP_Enter>", lambda event: self.manual_save_reading())
        self.bind_all("<Up>", lambda event: self.move_test_point(-1))
        self.bind_all("<Down>", lambda event: self.move_test_point(1))

        # Initial sheet configuration and port detection
        self.apply_sheet_configuration(self.sheet_name, initial=True)
        self.refresh_com_ports()

        # Handle window close cleanly
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Start live polling loop
        self.start_background_polling()

    def _check_gui_queue(self):
        """Poll and execute GUI tasks dispatched from background threads."""
        if hasattr(self, "gui_queue"):
            while not self.gui_queue.empty():
                try:
                    fn, args, kwargs = self.gui_queue.get_nowait()
                    fn(*args, **kwargs)
                except Exception as e:
                    print(f"Error executing queued GUI action: {e}")
        try:
            self.after(30, self._check_gui_queue)
        except Exception:
            pass

    def run_on_gui_thread(self, fn, *args, **kwargs):
        """Safely queue a callable to be executed on the Tkinter main thread."""
        if threading.current_thread() is threading.main_thread():
            try:
                fn(*args, **kwargs)
            except Exception as e:
                print(f"Error in GUI action: {e}")
        else:
            if hasattr(self, "gui_queue"):
                self.gui_queue.put((fn, args, kwargs))
            else:
                try:
                    self.after(0, lambda: fn(*args, **kwargs))
                except Exception:
                    pass

    def on_closing(self):
        self.polling_active = False
        if hasattr(self, "excel_queue") and self.excel_queue:
            try:
                self.excel_queue.put_nowait(None)
                if hasattr(self, "excel_worker") and self.excel_worker and self.excel_worker.is_alive():
                    self.excel_worker.join(timeout=1.5)
            except Exception:
                pass
        if hasattr(self, "local_api_server") and self.local_api_server:
            try:
                self.local_api_server.stop()
            except Exception:
                pass
        if hasattr(self, "flapate_client") and self.flapate_client:
            self.flapate_client._running = False
        if self.multimeter:
            try:
                self.multimeter.close()
            except Exception:
                pass
        try:
            self.destroy()
        except Exception:
            pass
        try:
            self.quit()
        except Exception:
            pass

    def init_table_data(self):
        """Initialize test points corresponding to active sheet total_rows."""
        self.table_data = {}
        for i in range(1, self.total_rows + 1):
            if i <= len(DEFAULT_SP_DEGREES):
                sp = DEFAULT_SP_DEGREES[i - 1]
            else:
                sp = "-"

            # Specific standard for X Ref TP 132 on CAM OUTPUT sheets: FLAP is 12.73, STALL/AUTOPILOT is 6.25
            sheet_upper = getattr(self, "sheet_name", "").upper()
            if "CAM OUTPUT" in sheet_upper and i == 132:
                ref_x_init = 12.73 if "FLAP" in sheet_upper else 6.25
            else:
                ref_x_init = "-"

            self.table_data[i] = {
                "tp": f"TP {i}",
                "sp": sp,
                "ref_x": ref_x_init,
                "ref_y": "-",
                "ref_z": "-",
                "test_x": "-",
                "test_y": "-",
                "test_z": "-",
                "status": "Pending",
                "is_warn": False,
                "warn_x": False,
                "warn_y": False,
                "warn_z": False,
            }

    # ==========================================================================
    # HEADER (GMF AEROASIA LOGO & SYSTEM TITLE)
    # ==========================================================================
    def create_header(self):
        header_frame = ctk.CTkFrame(self, height=76, corner_radius=0, fg_color=("#ffffff", "#1e222b"))
        header_frame.pack(side="top", fill="x", padx=0, pady=0)

        # GMF AeroAsia Logo (Top Left)
        logo_path = get_resource_path("gmf_logo.png")
        if os.path.exists(logo_path):
            try:
                pil_logo = Image.open(logo_path)
                self.gmf_img = ctk.CTkImage(light_image=pil_logo, dark_image=pil_logo, size=(225, 47))
                lbl_logo = ctk.CTkLabel(header_frame, image=self.gmf_img, text="")
                lbl_logo.pack(side="left", padx=(18, 15), pady=12)
            except Exception:
                lbl_fallback = ctk.CTkLabel(header_frame, text="GMF AeroAsia", font=ctk.CTkFont(size=20, weight="bold"), text_color="#00a3e0")
                lbl_fallback.pack(side="left", padx=18, pady=12)
        else:
            lbl_fallback = ctk.CTkLabel(header_frame, text="GMF AeroAsia", font=ctk.CTkFont(size=20, weight="bold"), text_color="#00a3e0")
            lbl_fallback.pack(side="left", padx=18, pady=12)

        title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_box.pack(side="left", padx=10, pady=8)

        lbl_app_title = ctk.CTkLabel(
            title_box,
            text="VFlap - Synchro Voltage Measurement",
            font=ctk.CTkFont(size=17, weight="bold")
        )
        lbl_app_title.pack(anchor="w")

        lbl_subtitle = ctk.CTkLabel(
            title_box,
            text="Agilent 34401A Multimeter • Flap Position Sensor Test Bench",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray70")
        )
        lbl_subtitle.pack(anchor="w")

        status_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        status_box.pack(side="right", padx=18, pady=10)

        self.lbl_hw_badge = ctk.CTkLabel(
            status_box,
            text="DISCONNECTED",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#ef4444",
            fg_color=("#f1f5f9", "#242a35"),
            corner_radius=4,
            padx=10,
            pady=3
        )
        self.lbl_hw_badge.pack(side="left", padx=(0, 8))

        self.lbl_lan_badge = ctk.CTkLabel(
            status_box,
            text="FLAPATE: GPIB TRACKING",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#38bdf8",
            fg_color=("#f1f5f9", "#242a35"),
            corner_radius=4,
            padx=10,
            pady=3
        )
        self.lbl_lan_badge.pack(side="left", padx=(0, 10))

        self.appearance_switch = ctk.CTkOptionMenu(
            status_box,
            values=["Dark", "Light", "System"],
            command=self.on_change_appearance_mode,
            width=85,
            height=26
        )
        self.appearance_switch.set("Dark")
        self.appearance_switch.pack(side="right")

    # ==========================================================================
    # MAIN 3-PANEL LAYOUT (LEFT, CENTER, RIGHT)
    # ==========================================================================
    def create_main_layout(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(side="top", fill="both", expand=True, padx=12, pady=10)

        container.grid_columnconfigure(0, weight=0, minsize=305)
        container.grid_columnconfigure(1, weight=1, minsize=540)
        container.grid_columnconfigure(2, weight=0, minsize=315)
        container.grid_rowconfigure(0, weight=1)

        self.create_left_panel(container)
        self.create_center_panel(container)
        self.create_right_panel(container)
        self.apply_theme_styling("Dark")

    # ==========================================================================
    # 1. LEFT PANEL: STEP-BY-STEP MEASUREMENT CONTROL
    # ==========================================================================
    def create_left_panel(self, parent):
        left_frame = ctk.CTkScrollableFrame(parent, corner_radius=8, width=305)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        ctk.CTkLabel(left_frame, text="Control Panel", font=ctk.CTkFont(size=14, weight="bold")).pack(padx=10, pady=(8, 8), anchor="w")

        # ----------------------------------------------------------------------
        # A. MEASUREMENT CHANNEL FOCUS
        # ----------------------------------------------------------------------
        chan_box = ctk.CTkFrame(left_frame, corner_radius=6)
        chan_box.pack(padx=6, pady=4, fill="x")

        ctk.CTkLabel(chan_box, text="1. Select Active Channel", font=ctk.CTkFont(size=12, weight="bold")).pack(padx=10, pady=(8, 2), anchor="w")

        self.lbl_channel_sub = ctk.CTkLabel(chan_box, text="Measure channel X, then Y (and Z if available):", font=ctk.CTkFont(size=11), text_color="gray70")
        self.lbl_channel_sub.pack(padx=10, pady=(0, 4), anchor="w")

        self.seg_channel = ctk.CTkSegmentedButton(
            chan_box,
            values=["X to GND", "Y to GND", "Z to GND"],
            command=self.on_channel_focus_changed,
            font=ctk.CTkFont(size=11, weight="bold")
        )
        self.seg_channel.set("X to GND")
        self.seg_channel.pack(padx=10, pady=(2, 6), fill="x")

        self.btn_confirm_probe = ctk.CTkButton(
            chan_box,
            text="✓ Probe Connected: Pin X-GND (Active)",
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=28,
            command=self.confirm_probe_connected
        )
        self.btn_confirm_probe.pack(padx=10, pady=(2, 6), fill="x")

        info_row = ctk.CTkFrame(chan_box, fg_color="transparent")
        info_row.pack(padx=10, pady=(2, 6), fill="x")

        self.lbl_target_info = ctk.CTkLabel(
            info_row,
            text="Focus: Test X (Col Q) | Row: TP 1 (277°) [FORWARD]",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_target_info.pack(side="left")

        # ----------------------------------------------------------------------
        # B. SAVE READING TO EXCEL (MANUAL / STABLE AUTO-SAVE)
        # ----------------------------------------------------------------------
        action_box = ctk.CTkFrame(left_frame, corner_radius=6)
        action_box.pack(padx=6, pady=4, fill="x")

        ctk.CTkLabel(action_box, text="2. Measurement & Auto-Input", font=ctk.CTkFont(size=12, weight="bold")).pack(padx=10, pady=(8, 4), anchor="w")

        self.btn_save_reading = ctk.CTkButton(
            action_box,
            text="Save Reading (Enter)",
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            height=34,
            command=self.manual_save_reading
        )
        self.btn_save_reading.pack(padx=10, pady=(4, 4), fill="x")

        ctk.CTkLabel(action_box, text="Press ENTER or click button to save reading.", font=ctk.CTkFont(size=10), text_color="gray70").pack(padx=10, pady=(0, 4))

        # Input Direction Selection (Forward / Reverse)
        dir_frame = ctk.CTkFrame(action_box, fg_color="transparent")
        dir_frame.pack(padx=10, pady=(2, 4), fill="x")

        ctk.CTkLabel(
            dir_frame,
            text="Direction:",
            font=ctk.CTkFont(size=11, weight="bold")
        ).pack(side="left", padx=(0, 6))

        self.seg_direction = ctk.CTkSegmentedButton(
            dir_frame,
            values=["Forward (1 -> End)", "Reverse (End -> 1)"],
            font=ctk.CTkFont(size=10, weight="bold"),
            command=self.on_direction_changed
        )
        self.seg_direction.set("Forward (1 -> End)")
        self.seg_direction.pack(side="left", fill="x", expand=True)

        # Auto-Save Control (Single Checkbox in English)
        auto_box = ctk.CTkFrame(action_box, fg_color=("#f1f5f9", "#1e242d"), corner_radius=8)
        auto_box.pack(padx=10, pady=(6, 6), fill="x")

        self.chk_auto_save = ctk.CTkCheckBox(
            auto_box,
            text="Enable Auto-Save",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self.on_toggle_autosave
        )
        self.chk_auto_save.pack(padx=10, pady=(8, 4), anchor="w")
        self.chk_auto_save.select()

        # FlapATE Status Indicator (English)
        self.lbl_needle_motion = ctk.CTkLabel(
            auto_box,
            text="FlapATE: Ready (Awaiting FlapATE stop)",
            font=ctk.CTkFont(size=10),
            text_color="#10b981"
        )
        self.lbl_needle_motion.pack(padx=10, pady=(0, 8), anchor="w")

        # Backward compatibility aliases
        self.chk_mode_auto = self.chk_auto_save
        self.chk_mode_manual = None
        self.chk_sound_alert = None
        self.sound_enabled = True
        self.lbl_flapate_badge = self.lbl_needle_motion
        self.lbl_local_api_badge = self.lbl_needle_motion

        self.lbl_stability_status = ctk.CTkLabel(
            action_box,
            text="Trigger: Waiting for FlapATE data capture / needle stop...",
            font=ctk.CTkFont(size=10),
            text_color="gray70"
        )
        self.lbl_stability_status.pack(padx=10, pady=(2, 6), anchor="w")

        # ----------------------------------------------------------------------
        # C. MULTIMETER CONNECTION & DEMO MODE
        # ----------------------------------------------------------------------
        conn_box = ctk.CTkFrame(left_frame, corner_radius=6)
        conn_box.pack(padx=6, pady=4, fill="x")

        ctk.CTkLabel(conn_box, text="3. Multimeter Connection", font=ctk.CTkFont(size=12, weight="bold")).pack(padx=10, pady=(8, 4), anchor="w")

        com_layout = ctk.CTkFrame(conn_box, fg_color="transparent")
        com_layout.pack(padx=10, pady=2, fill="x")

        self.com_combo = ctk.CTkComboBox(com_layout, values=["Scanning..."])
        self.com_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))

        btn_refresh = ctk.CTkButton(com_layout, text="Refresh", width=55, height=28, font=ctk.CTkFont(size=11), command=self.refresh_com_ports)
        btn_refresh.pack(side="right")

        self.btn_connect = ctk.CTkButton(
            conn_box,
            text="Connect Multimeter",
            font=ctk.CTkFont(weight="bold"),
            fg_color="#1d4ed8",
            hover_color="#1e40af",
            command=self.toggle_connection
        )
        self.btn_connect.pack(padx=10, pady=(4, 6), fill="x")

        self.chk_sim_mode = ctk.CTkCheckBox(
            conn_box,
            text="Demo / Simulation Mode",
            font=ctk.CTkFont(size=11),
            command=self.toggle_sim_mode
        )
        self.chk_sim_mode.pack(padx=10, pady=(2, 6), anchor="w")

        # ----------------------------------------------------------------------
        # D. EXCEL TARGET FILE SETUP
        # ----------------------------------------------------------------------
        excel_box = ctk.CTkFrame(left_frame, corner_radius=6)
        excel_box.pack(padx=6, pady=4, fill="x")

        ctk.CTkLabel(excel_box, text="4. Excel Target File", font=ctk.CTkFont(size=12, weight="bold")).pack(padx=10, pady=(8, 4), anchor="w")

        display_name = os.path.basename(self.excel_file_path) if self.excel_file_path else "No File Selected"
        self.lbl_excel_file = ctk.CTkLabel(
            excel_box,
            text=display_name,
            text_color="gray70" if not self.excel_file_path else "#38bdf8",
            wraplength=270,
            font=ctk.CTkFont(size=11)
        )
        self.lbl_excel_file.pack(padx=10, pady=2)

        btn_browse = ctk.CTkButton(excel_box, text="Browse...", font=ctk.CTkFont(size=11), command=self.browse_excel)
        btn_browse.pack(padx=10, pady=(2, 6), fill="x")

        ctk.CTkLabel(excel_box, text="Active Sheet:", font=ctk.CTkFont(size=11, weight="bold")).pack(padx=10, pady=(2, 2), anchor="w")
        self.sheet_combo = ctk.CTkComboBox(
            excel_box,
            values=[
                "ACCURACY STALL",
                "CAM OUTPUT STALL",
                "ACCURACY AUTOPILOT",
                "CAM OUTPUT AUTOPILOT",
                "ACCURACY FLAP",
                "CAM OUTPUT FLAP",
            ],
            command=self.on_sheet_changed
        )
        self.sheet_combo.set(self.sheet_name)
        self.sheet_combo.pack(padx=10, pady=(0, 4), fill="x")

        self.lbl_sheet_desc = ctk.CTkLabel(
            excel_box,
            text="Target: Synchro Under Test",
            font=ctk.CTkFont(size=10),
            text_color="#94a3b8"
        )
        self.lbl_sheet_desc.pack(padx=10, pady=(0, 6), anchor="w")

    # ==========================================================================
    # 2. CENTER PANEL: RESULTS TABLE X, Y, Z (EXCEL FORMAT)
    # ==========================================================================
    def create_center_panel(self, parent):
        center_frame = ctk.CTkFrame(parent, corner_radius=8, fg_color=("#ffffff", "#161920"), border_width=1, border_color=("#e2e8f0", "#242a35"))
        self.center_frame = center_frame
        center_frame.grid(row=0, column=1, sticky="nsew", padx=6)
        center_frame.grid_rowconfigure(2, weight=1)
        center_frame.grid_columnconfigure(0, weight=1)

        title_box = ctk.CTkFrame(center_frame, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=14, pady=(10, 4), sticky="ew")

        ctk.CTkLabel(
            title_box,
            text="Measurement Results (Excel Format)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")

        self.lbl_table_info = ctk.CTkLabel(
            title_box,
            text="Active Measurement: X to GND",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#38bdf8"
        )
        self.lbl_table_info.pack(side="right")

        self.super_header = ctk.CTkFrame(center_frame, fg_color=("#e2e8f0", "#1e242d"), height=26, corner_radius=4)
        self.super_header.grid(row=1, column=0, padx=14, pady=(4, 2), sticky="ew")
        self.super_header.grid_columnconfigure(0, weight=18)
        self.super_header.grid_columnconfigure(1, weight=31)
        self.super_header.grid_columnconfigure(2, weight=33)
        self.super_header.grid_columnconfigure(3, weight=18)
        self.super_header.grid_columnconfigure(4, weight=0, minsize=16)

        hdr_color = ("#475569", "#cbd5e1")
        ctk.CTkLabel(self.super_header, text="TEST POINT & SP", font=ctk.CTkFont(size=10, weight="bold"), text_color=hdr_color).grid(row=0, column=0, padx=5, pady=3)
        ctk.CTkLabel(self.super_header, text="SYNCHRO REFERENCE", font=ctk.CTkFont(size=10, weight="bold"), text_color=hdr_color).grid(row=0, column=1, padx=5, pady=3)
        self.lbl_hdr_test = ctk.CTkLabel(self.super_header, text="SYNCHRO UNDER TEST", font=ctk.CTkFont(size=10, weight="bold"), text_color=hdr_color)
        self.lbl_hdr_test.grid(row=0, column=2, padx=5, pady=3)
        ctk.CTkLabel(self.super_header, text="DEVIATION STATUS", font=ctk.CTkFont(size=10, weight="bold"), text_color=hdr_color).grid(row=0, column=3, padx=5, pady=3)

        table_container = ctk.CTkFrame(center_frame, fg_color="transparent")
        table_container.grid(row=2, column=0, padx=14, pady=4, sticky="nsew")
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_rowconfigure(1, weight=0)
        table_container.grid_columnconfigure(0, weight=1)
        table_container.grid_columnconfigure(1, weight=0)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Excel.Treeview",
            background="#161920",
            foreground="#e2e8f0",
            fieldbackground="#161920",
            rowheight=26,
            font=('Segoe UI', 9)
        )
        style.configure(
            "Excel.Treeview.Heading",
            background="#20242e",
            foreground="#cbd5e1",
            font=('Segoe UI', 9, 'bold'),
            relief="flat",
            padding=(2, 4)
        )
        style.map("Excel.Treeview", background=[('selected', '#1d4ed8')], foreground=[('selected', '#ffffff')])

        cols = ("TEST", "SP_DEG", "REF_X", "REF_Y", "REF_Z", "TEST_X", "TEST_Y", "TEST_Z", "STATUS")
        self.tree = ttk.Treeview(table_container, columns=cols, show="headings", style="Excel.Treeview", selectmode="browse")

        self.tree.heading("TEST", text="Test")
        self.tree.heading("SP_DEG", text="SP °")
        self.tree.heading("REF_X", text="Ref X")
        self.tree.heading("REF_Y", text="Ref Y")
        self.tree.heading("REF_Z", text="Ref Z")
        self.tree.heading("TEST_X", text="Test X")
        self.tree.heading("TEST_Y", text="Test Y")
        self.tree.heading("TEST_Z", text="Test Z")
        self.tree.heading("STATUS", text="Status")

        self.tree.column("TEST", width=58, minwidth=46, anchor="center", stretch=False)
        self.tree.column("SP_DEG", width=54, minwidth=44, anchor="center", stretch=False)
        self.tree.column("REF_X", width=62, minwidth=48, anchor="center", stretch=True)
        self.tree.column("REF_Y", width=62, minwidth=48, anchor="center", stretch=True)
        self.tree.column("REF_Z", width=62, minwidth=48, anchor="center", stretch=True)
        self.tree.column("TEST_X", width=66, minwidth=52, anchor="center", stretch=True)
        self.tree.column("TEST_Y", width=66, minwidth=52, anchor="center", stretch=True)
        self.tree.column("TEST_Z", width=66, minwidth=52, anchor="center", stretch=True)
        self.tree.column("STATUS", width=72, minwidth=55, anchor="center", stretch=True)

        self.tree.grid(row=0, column=0, sticky="nsew")

        # Tag styling for deviation status (initialized dynamically by apply_theme_styling)
        self.tree.tag_configure("warn_row", background="#2e210a", foreground="#fef08a")
        self.tree.tag_configure("ok_row", background="#161920", foreground="#e2e8f0")
        self.tree.tag_configure("even_ok", background="#191d26", foreground="#e2e8f0")
        self.tree.tag_configure("odd_ok", background="#151820", foreground="#e2e8f0")
        self.tree.tag_configure("even_all_warn", background="#451a03", foreground="#fef08a")
        self.tree.tag_configure("odd_all_warn", background="#3b1702", foreground="#fef08a")

        sb_y = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        sb_x = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)

        def _on_tree_yscroll(*args):
            sb_y.set(*args)
            self.after_idle(self._refresh_cell_highlights)

        def _on_tree_xscroll(*args):
            sb_x.set(*args)
            self.after_idle(self._refresh_cell_highlights)

        self.tree.configure(yscrollcommand=_on_tree_yscroll, xscrollcommand=_on_tree_xscroll)
        sb_y.grid(row=0, column=1, sticky="ns")
        sb_x.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self.on_tree_row_selected)
        self.tree.bind("<Configure>", lambda e: self.after_idle(self._refresh_cell_highlights), add="+")
        self.tree.bind("<MouseWheel>", lambda e: self.after_idle(self._refresh_cell_highlights), add="+")
        self.tree.bind("<Button-4>", lambda e: self.after_idle(self._refresh_cell_highlights), add="+")
        self.tree.bind("<Button-5>", lambda e: self.after_idle(self._refresh_cell_highlights), add="+")

        # Bottom Table Toolbar
        bottom_bar = ctk.CTkFrame(center_frame, fg_color="transparent")
        bottom_bar.grid(row=3, column=0, padx=14, pady=(6, 10), sticky="ew")

        btn_load = ctk.CTkButton(
            bottom_bar,
            text="Reload Excel",
            width=130,
            font=ctk.CTkFont(size=11),
            command=self.load_from_excel
        )
        btn_load.pack(side="left", padx=(0, 6))

        btn_save_all = ctk.CTkButton(
            bottom_bar,
            text="Save All to Excel",
            width=130,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color="#059669",
            hover_color="#047857",
            command=self.save_all_to_excel
        )
        btn_save_all.pack(side="left", padx=6)

        btn_clear = ctk.CTkButton(
            bottom_bar,
            text="Clear Results",
            width=130,
            font=ctk.CTkFont(size=11),
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=self.clear_test_results
        )
        btn_clear.pack(side="right", padx=(6, 0))

    # ==========================================================================
    # 3. RIGHT PANEL: AGILENT 34401A MONITOR & SCPI LOG
    # ==========================================================================
    def create_right_panel(self, parent):
        right_frame = ctk.CTkFrame(parent, corner_radius=8, width=315, fg_color=("#ffffff", "#161920"), border_width=1, border_color=("#e2e8f0", "#242a35"))
        self.right_frame = right_frame
        right_frame.grid(row=0, column=2, sticky="nsew", padx=(6, 0))
        right_frame.grid_rowconfigure(4, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        lbl_mon_title = ctk.CTkLabel(
            right_frame,
            text="Multimeter Monitor",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl_mon_title.grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        bezel = ctk.CTkFrame(
            right_frame,
            corner_radius=6,
            fg_color=("#f8fafc", "#1a1e27"),
            border_width=1,
            border_color=("#cbd5e1", "#2e3542")
        )
        self.bezel = bezel
        bezel.grid(row=1, column=0, padx=12, pady=4, sticky="ew")

        self.vfd_screen = ctk.CTkFrame(bezel, corner_radius=4, fg_color=("#ffffff", "#0b1219"), border_width=1, border_color=("#cbd5e1", "#1e293b"))
        self.vfd_screen.pack(padx=6, pady=6, fill="x")
        screen = self.vfd_screen

        annunc_box = ctk.CTkFrame(screen, fg_color="transparent")
        annunc_box.pack(padx=10, pady=(4, 0), fill="x")

        self.lbl_vfd_trig = ctk.CTkLabel(
            annunc_box,
            text="*",
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color=VFD_BG
        )
        self.lbl_vfd_trig.pack(side="left")

        self.lbl_vfd_auto = ctk.CTkLabel(
            annunc_box,
            text="AUTO",
            font=ctk.CTkFont(family="Consolas", size=10, weight="bold"),
            text_color=VFD_MUTED
        )
        self.lbl_vfd_auto.pack(side="left", padx=8)

        self.lbl_vfd_mode = ctk.CTkLabel(
            annunc_box,
            text="AC V",
            font=ctk.CTkFont(family="Consolas", size=10, weight="bold"),
            text_color=("#0284c7", "#38bdf8")
        )
        self.lbl_vfd_mode.pack(side="left", padx=4)

        self.lbl_vfd_rate = ctk.CTkLabel(
            annunc_box,
            text="6.5 DIGITS",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=VFD_MUTED
        )
        self.lbl_vfd_rate.pack(side="right")

        digits_box = ctk.CTkFrame(screen, fg_color="transparent")
        digits_box.pack(padx=10, pady=(2, 4), fill="x")

        self.lbl_vfd_digits = ctk.CTkLabel(
            digits_box,
            text="00.0000",
            font=ctk.CTkFont(family="Consolas", size=34, weight="bold"),
            text_color=("#0f172a", "#38bdf8")
        )
        self.lbl_vfd_digits.pack(side="left", expand=True, anchor="e")

        self.lbl_vfd_unit = ctk.CTkLabel(
            digits_box,
            text=" VAC",
            font=ctk.CTkFont(family="Consolas", size=20, weight="bold"),
            text_color=("#0f172a", "#38bdf8")
        )
        self.lbl_vfd_unit.pack(side="right", padx=(4, 8))

        sub_screen = ctk.CTkFrame(screen, fg_color="transparent")
        sub_screen.pack(padx=10, pady=(0, 4), fill="x")

        self.lbl_vfd_channel = ctk.CTkLabel(
            sub_screen,
            text="CHANNEL: X-GND",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=("#0284c7", "#38bdf8")
        )
        self.lbl_vfd_channel.pack(side="left")

        lbl_vfd_model = ctk.CTkLabel(
            sub_screen,
            text="Agilent 34401A",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=VFD_MUTED
        )
        lbl_vfd_model.pack(side="right")

        # Signal Stability Indicator
        stab_card = ctk.CTkFrame(right_frame, corner_radius=6, fg_color=("#f8fafc", "#1e242d"), border_width=1, border_color=("#e2e8f0", "#2e3542"))
        self.stab_card = stab_card
        stab_card.grid(row=2, column=0, padx=12, pady=4, sticky="ew")

        self.lbl_mon_stab = ctk.CTkLabel(
            stab_card,
            text="FlapATE Status: Ready",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="gray70"
        )
        self.lbl_mon_stab.pack(padx=10, pady=(6, 2), anchor="w")

        self.progress_stab = ctk.CTkProgressBar(stab_card, height=8)
        self.progress_stab.set(0.0)
        self.progress_stab.pack(padx=10, pady=(2, 6), fill="x")

        lbl_log = ctk.CTkLabel(right_frame, text="System Log", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_log.grid(row=3, column=0, padx=12, pady=(8, 2), sticky="w")

        self.txt_log = ctk.CTkTextbox(
            right_frame,
            font=ctk.CTkFont(family="Consolas", size=10),
            corner_radius=6,
            fg_color=("#f8fafc", "#12151c"),
            text_color=("#0f172a", "#e2e8f0"),
            border_width=1,
            border_color=("#e2e8f0", "#2a303c")
        )
        self.txt_log.grid(row=4, column=0, padx=12, pady=(2, 4), sticky="nsew")

        btn_clr_log = ctk.CTkButton(
            right_frame,
            text="Clear Log",
            height=22,
            width=70,
            font=ctk.CTkFont(size=10),
            command=lambda: self.txt_log.delete("1.0", "end")
        )
        btn_clr_log.grid(row=5, column=0, padx=12, pady=(0, 8), sticky="e")

        self.log_message("System initialized. Agilent 34401A interface ready.")

    def on_change_appearance_mode(self, mode):
        ctk.set_appearance_mode(mode)
        resolved = mode
        if mode == "System":
            resolved = ctk.get_appearance_mode()
        self.apply_theme_styling(resolved)
        self.update_table_view()

    def _get_channel_color(self, ch):
        is_light = getattr(self, "appearance_mode", "Dark") == "Light"
        if ch == "X":
            return "#0284c7" if is_light else "#38bdf8"
        elif ch == "Y":
            return "#0f766e" if is_light else "#2dd4bf"
        else:
            return "#d97706" if is_light else "#f59e0b"

    def apply_theme_styling(self, mode):
        """Apply full Light/Dark styling dynamically to Measurement Results, Treeview table, and Monitor."""
        mode_str = str(mode).lower()
        is_light = (mode_str == "light")
        self.appearance_mode = "Light" if is_light else "Dark"

        if is_light:
            tree_bg = "#ffffff"
            tree_fg = "#0f172a"
            even_bg = "#ffffff"
            odd_bg = "#f8fafc"
            warn_bg = "#fef08a"
            warn_fg = "#713f12"
            warn_odd_bg = "#fef9c3"
            heading_bg = "#f1f5f9"
            heading_fg = "#1e293b"
            sel_bg = "#2563eb"
            sel_fg = "#ffffff"
            self.vfd_bg_color = "#ffffff"
            vfd_digits_color = "#0f172a"
            vfd_mode_color = "#0284c7"
            vfd_muted_color = "#64748b"
        else:
            tree_bg = "#161920"
            tree_fg = "#e2e8f0"
            even_bg = "#191d26"
            odd_bg = "#151820"
            warn_bg = "#451a03"
            warn_fg = "#fef08a"
            warn_odd_bg = "#3b1702"
            heading_bg = "#20242e"
            heading_fg = "#cbd5e1"
            sel_bg = "#1d4ed8"
            sel_fg = "#ffffff"
            self.vfd_bg_color = "#0b1219"
            vfd_digits_color = "#38bdf8"
            vfd_mode_color = "#38bdf8"
            vfd_muted_color = "#64748b"

        # 1. Update Center Panel & Treeview Styling
        if hasattr(self, "center_frame") and self.center_frame.winfo_exists():
            self.center_frame.configure(
                fg_color="#ffffff" if is_light else "#161920",
                border_color="#e2e8f0" if is_light else "#242a35"
            )
        if hasattr(self, "super_header") and self.super_header.winfo_exists():
            self.super_header.configure(fg_color="#f1f5f9" if is_light else "#1e242d")

        style = ttk.Style()
        style.configure(
            "Excel.Treeview",
            background=tree_bg,
            foreground=tree_fg,
            fieldbackground=tree_bg,
            rowheight=26,
            font=('Segoe UI', 9)
        )
        style.configure(
            "Excel.Treeview.Heading",
            background=heading_bg,
            foreground=heading_fg,
            font=('Segoe UI', 9, 'bold'),
            relief="flat",
            padding=(2, 4)
        )
        style.map("Excel.Treeview", background=[('selected', sel_bg)], foreground=[('selected', sel_fg)])

        if hasattr(self, "tree") and self.tree.winfo_exists():
            self.tree.tag_configure("even_ok", background=even_bg, foreground=tree_fg)
            self.tree.tag_configure("odd_ok", background=odd_bg, foreground=tree_fg)
            self.tree.tag_configure("ok_row", background=even_bg, foreground=tree_fg)
            self.tree.tag_configure("even_all_warn", background=warn_bg, foreground=warn_fg)
            self.tree.tag_configure("odd_all_warn", background=warn_odd_bg, foreground=warn_fg)
            self.tree.tag_configure("warn_row", background=warn_bg, foreground=warn_fg)
            self.tree.tag_configure("even_warn", background=warn_bg, foreground=warn_fg)
            self.tree.tag_configure("odd_warn", background=warn_odd_bg, foreground=warn_fg)

        # 2. Update Right Panel & Multimeter Monitor Styling
        if hasattr(self, "right_frame") and self.right_frame.winfo_exists():
            self.right_frame.configure(
                fg_color="#ffffff" if is_light else "#161920",
                border_color="#e2e8f0" if is_light else "#242a35"
            )
        if hasattr(self, "bezel") and self.bezel.winfo_exists():
            self.bezel.configure(
                fg_color="#f8fafc" if is_light else "#1a1e27",
                border_color="#cbd5e1" if is_light else "#2e3542"
            )
        if hasattr(self, "vfd_screen") and self.vfd_screen.winfo_exists():
            self.vfd_screen.configure(
                fg_color=self.vfd_bg_color,
                border_color="#cbd5e1" if is_light else "#1e293b"
            )
        if hasattr(self, "lbl_vfd_trig") and self.lbl_vfd_trig.winfo_exists():
            self.lbl_vfd_trig.configure(text_color=self.vfd_bg_color)
        if hasattr(self, "lbl_vfd_digits") and self.lbl_vfd_digits.winfo_exists():
            self.lbl_vfd_digits.configure(text_color=vfd_digits_color)
        if hasattr(self, "lbl_vfd_unit") and self.lbl_vfd_unit.winfo_exists():
            self.lbl_vfd_unit.configure(text_color=vfd_digits_color)
        if hasattr(self, "lbl_vfd_mode") and self.lbl_vfd_mode.winfo_exists():
            self.lbl_vfd_mode.configure(text_color=vfd_mode_color)
        if hasattr(self, "lbl_vfd_auto") and self.lbl_vfd_auto.winfo_exists():
            self.lbl_vfd_auto.configure(text_color=vfd_muted_color)
        if hasattr(self, "lbl_vfd_rate") and self.lbl_vfd_rate.winfo_exists():
            self.lbl_vfd_rate.configure(text_color=vfd_muted_color)
        if hasattr(self, "lbl_vfd_channel") and self.lbl_vfd_channel.winfo_exists():
            self.lbl_vfd_channel.configure(text_color=self._get_channel_color(self.active_channel))
        if hasattr(self, "stab_card") and self.stab_card.winfo_exists():
            self.stab_card.configure(
                fg_color="#f8fafc" if is_light else "#1e242d",
                border_color="#e2e8f0" if is_light else "#2e3542"
            )
        if hasattr(self, "lbl_mon_stab") and self.lbl_mon_stab.winfo_exists():
            self.lbl_mon_stab.configure(text_color="#334155" if is_light else "gray70")
        if hasattr(self, "txt_log") and self.txt_log.winfo_exists():
            self.txt_log.configure(
                fg_color="#f8fafc" if is_light else "#12151c",
                text_color="#0f172a" if is_light else "#e2e8f0",
                border_color="#e2e8f0" if is_light else "#2a303c"
            )

        self.after_idle(self._refresh_cell_highlights)

    def _refresh_cell_highlights(self):
        """Place pooled highlight labels over individual error cells (TEST_X, TEST_Y, TEST_Z, STATUS)."""
        if not hasattr(self, "tree") or not self.tree.winfo_exists():
            return

        is_light = getattr(self, "appearance_mode", "Dark") == "Light"
        warn_bg = "#fef08a" if is_light else "#451a03"
        warn_fg = "#713f12" if is_light else "#fef08a"

        if not hasattr(self, "_cell_highlight_labels"):
            self._cell_highlight_labels = []

        pool_idx = 0
        children = self.tree.get_children()

        for iid in children:
            try:
                row_idx = int(iid)
            except ValueError:
                continue

            r = self.table_data.get(row_idx)
            if not r:
                continue

            warn_x = r.get("warn_x", False)
            warn_y = r.get("warn_y", False)
            warn_z = r.get("warn_z", False) if self.has_z else False
            is_all_warn = (warn_x and warn_y and warn_z) if self.has_z else (warn_x and warn_y)

            # If all channels error, entire row/block is colored yellow via even_all_warn/odd_all_warn tag
            if is_all_warn:
                continue

            warn_cols = []
            if warn_x:
                warn_cols.append(("TEST_X", format_2_decimals(r.get("test_x", "-"))))
            if warn_y:
                warn_cols.append(("TEST_Y", format_2_decimals(r.get("test_y", "-"))))
            if self.has_z and warn_z:
                warn_cols.append(("TEST_Z", format_2_decimals(r.get("test_z", "-"))))

            if warn_cols:
                # User requirement: status cell is also yellow when any channel errors
                warn_cols.append(("STATUS", str(r.get("status", "Deviation"))))

            for col_id, text_val in warn_cols:
                bbox = self.tree.bbox(iid, col_id)
                if not bbox or len(bbox) != 4:
                    continue
                bx, by, bw, bh = bbox
                if bw <= 0 or bh <= 0:
                    continue

                if pool_idx < len(self._cell_highlight_labels):
                    lbl = self._cell_highlight_labels[pool_idx]
                else:
                    lbl = tk.Label(self.tree, font=('Segoe UI', 9), anchor="center")
                    self._cell_highlight_labels.append(lbl)

                pool_idx += 1
                lbl.configure(text=text_val, bg=warn_bg, fg=warn_fg, cursor="hand2")
                lbl.place(x=bx, y=by, width=bw, height=bh)
                lbl.bind("<Button-1>", lambda e, r_idx=row_idx: self.select_row(r_idx, from_auto_advance=False))

        # Hide any excess pooled labels
        for idx in range(pool_idx, len(self._cell_highlight_labels)):
            self._cell_highlight_labels[idx].place_forget()

    # ==========================================================================
    # LOGIC: SHEET CONFIGURATION & SPECIFIC MAPPING
    # ==========================================================================
    def on_sheet_changed(self, choice):
        self.apply_sheet_configuration(choice, initial=False)

    def apply_sheet_configuration(self, choice, initial=False):
        """Apply sheet-specific configurations for STALL, AUTOPILOT, or FLAP."""
        self.sheet_name = choice

        # Search config: exact match first, then normalized equality, then substring
        cfg = SHEET_CONFIGS.get(choice)
        if cfg is None:
            choice_clean = choice.replace(" ", "").upper()
            for k, v in SHEET_CONFIGS.items():
                if k.replace(" ", "").upper() == choice_clean:
                    cfg = v
                    break
        if cfg is None:
            choice_clean = choice.replace(" ", "").upper()
            for k, v in SHEET_CONFIGS.items():
                k_clean = k.replace(" ", "").upper()
                if k_clean in choice_clean or choice_clean in k_clean:
                    cfg = v
                    break
        if cfg is None:
            cfg = SHEET_CONFIGS["ACCURACY STALL"]

        self.cur_cfg = cfg
        self.start_row = cfg["start_row"]
        self.total_rows = cfg.get("total_rows", 11)
        self.has_z = cfg["has_z"]
        self.test_cols = cfg["test_cols"]

        # Reset selected row based on active direction (TP 1 for forward, TP total_rows for reverse)
        if getattr(self, "input_direction", "forward") == "reverse":
            self.selected_row_idx = self.total_rows
        else:
            self.selected_row_idx = 1
        self.auto_saved_this_row = False
        self.last_saved_val_2dec = None
        self.awaiting_movement = False
        self.stable_candidate_val = None
        self.stable_candidate_raw = None
        self.stable_candidate_start_time = 0.0
        self.history_2dec = []

        # Re-inisialisasi struktur table_data sesuai total_rows sheet aktif
        self.init_table_data()

        # Perbarui Channel Selector
        channels = cfg["channels"]
        self.seg_channel.configure(values=channels)

        # Jika kanal Z tidak ada (seperti pada Flap), kembalikan ke X jika sedang di Z
        if not self.has_z and self.active_channel == "Z":
            self.active_channel = "X"
            self.seg_channel.set("X to GND")

        # Perbarui Heading Tabel & Tata Letak Kolom Presisi di UI
        ref_cols = cfg["ref_cols"]
        test_cols = cfg["test_cols"]
        tp_col_letter = openpyxl.utils.get_column_letter(cfg.get("tp_col", 1))
        sp_col_letter = openpyxl.utils.get_column_letter(cfg["sp_col"])
        self.tree.heading("TEST", text=f"Test ({tp_col_letter})")
        self.tree.heading("SP_DEG", text=f"SP ° ({sp_col_letter})")

        for channel in ("X", "Y"):
            self.tree.heading(
                f"REF_{channel}",
                text=f"Ref {channel} ({openpyxl.utils.get_column_letter(ref_cols[channel])})"
            )
            self.tree.heading(
                f"TEST_{channel}",
                text=f"Test {channel} ({openpyxl.utils.get_column_letter(test_cols[channel])})"
            )

        end_row = self.start_row + self.total_rows - 1
        test_col_letters = ", ".join(openpyxl.utils.get_column_letter(col) for col in test_cols.values())
        if not self.has_z:
            # Flap hanya mengukur X dan Y: sembunyikan kolom Z agar tabel presisi dan tidak terpotong
            self.tree.configure(displaycolumns=("TEST", "SP_DEG", "REF_X", "REF_Y", "TEST_X", "TEST_Y", "STATUS"))
            self.tree.column("TEST", width=62, minwidth=48, anchor="center", stretch=False)
            self.tree.column("SP_DEG", width=58, minwidth=46, anchor="center", stretch=False)
            self.tree.column("REF_X", width=75, minwidth=55, anchor="center", stretch=True)
            self.tree.column("REF_Y", width=75, minwidth=55, anchor="center", stretch=True)
            self.tree.column("TEST_X", width=80, minwidth=60, anchor="center", stretch=True)
            self.tree.column("TEST_Y", width=80, minwidth=60, anchor="center", stretch=True)
            self.tree.column("STATUS", width=78, minwidth=60, anchor="center", stretch=True)

            self.super_header.grid_columnconfigure(0, weight=23)
            self.super_header.grid_columnconfigure(1, weight=31)
            self.super_header.grid_columnconfigure(2, weight=33)
            self.super_header.grid_columnconfigure(3, weight=16)

            self.lbl_sheet_desc.configure(text=f"Target: Synchro Under Test (Cols {test_col_letters} | Rows {self.start_row}-{end_row})")
            self.lbl_channel_sub.configure(text="Flap measures X and Y only:")
        else:
            for channel in ("Z",):
                self.tree.heading(
                    f"REF_{channel}",
                    text=f"Ref {channel} ({openpyxl.utils.get_column_letter(ref_cols[channel])})"
                )
                self.tree.heading(
                    f"TEST_{channel}",
                    text=f"Test {channel} ({openpyxl.utils.get_column_letter(test_cols[channel])})"
                )
            self.tree.configure(displaycolumns=("TEST", "SP_DEG", "REF_X", "REF_Y", "REF_Z", "TEST_X", "TEST_Y", "TEST_Z", "STATUS"))
            self.tree.column("TEST", width=58, minwidth=46, anchor="center", stretch=False)
            self.tree.column("SP_DEG", width=54, minwidth=44, anchor="center", stretch=False)
            self.tree.column("REF_X", width=62, minwidth=48, anchor="center", stretch=True)
            self.tree.column("REF_Y", width=62, minwidth=48, anchor="center", stretch=True)
            self.tree.column("REF_Z", width=62, minwidth=48, anchor="center", stretch=True)
            self.tree.column("TEST_X", width=66, minwidth=52, anchor="center", stretch=True)
            self.tree.column("TEST_Y", width=66, minwidth=52, anchor="center", stretch=True)
            self.tree.column("TEST_Z", width=66, minwidth=52, anchor="center", stretch=True)
            self.tree.column("STATUS", width=72, minwidth=55, anchor="center", stretch=True)

            self.super_header.grid_columnconfigure(0, weight=18)
            self.super_header.grid_columnconfigure(1, weight=31)
            self.super_header.grid_columnconfigure(2, weight=33)
            self.super_header.grid_columnconfigure(3, weight=18)

            self.lbl_sheet_desc.configure(text=f"Target: Synchro Under Test (Cols {test_col_letters} | Rows {self.start_row}-{end_row})")
            self.lbl_channel_sub.configure(text="Measure channels X, Y, then Z:")

        self.update_target_info_label()

        # Baca data awal dari Excel jika file tersedia
        if self.excel_file_path and os.path.exists(self.excel_file_path):
            self.load_from_excel(quiet=True)
        else:
            self.update_table_view()

        self.log_message(f"Sheet set to '{choice}' | Data rows: {self.start_row}-{end_row} ({self.total_rows} rows) | Channel Z: {'Available' if self.has_z else 'Not available (Flap)'}")

    def on_channel_focus_changed(self, choice):
        prev_chan = getattr(self, "active_channel", "X")
        if "X" in choice:
            new_chan = "X"
        elif "Y" in choice:
            new_chan = "Y"
        elif "Z" in choice:
            if not self.has_z:
                messagebox.showinfo("Channel Z Unavailable", f"The '{self.sheet_name}' sheet measures Synchro Flap using channels X and Y only.")
                self.seg_channel.set("X to GND")
                self.active_channel = "X"
                return
            new_chan = "Z"
        else:
            new_chan = "X"

        is_switch = (new_chan != prev_chan)
        self.active_channel = new_chan

        self.last_saved_val_2dec = None
        self.last_saved_raw = None
        self.history_2dec = []
        self.stable_candidate_val = None
        self.stable_candidate_raw = None
        self.stable_candidate_start_time = 0.0
        self.auto_saved_this_row = False
        self._channel_settle_start_time = None
        self.update_target_info_label()
        col_letter = openpyxl.utils.get_column_letter(self.test_cols[self.active_channel])

        if is_switch:
            # STOP READING IMMEDIATELY while moving probe
            self.is_channel_armed = False
            self.channel_transition_active = True
            self.awaiting_movement = True
            self.dynamix_state = DYN_STATE_AWAITING_MOVE
            self.current_live_voltage = 0.0
            self.update_vfd_display(0.0, channel=f"{self.active_channel}-GND [PAUSED]")
            if hasattr(self, "btn_confirm_probe") and self.btn_confirm_probe is not None:
                self.btn_confirm_probe.configure(
                    text=f"👉 Connect Probe to Pin {self.active_channel}-GND (Click to Resume)",
                    fg_color="#d97706",
                    hover_color="#b45309"
                )
            self.log_message(f"Channel switched to {self.active_channel}-GND. Multimeter reading PAUSED. Move probe to Pin {self.active_channel}.")
        else:
            self.lbl_vfd_channel.configure(text=f"CHANNEL: {self.active_channel}-GND", text_color=self._get_channel_color(self.active_channel))
            self.log_message(f"Measurement focus on Channel {self.active_channel}-GND (Column {col_letter}).")

    def confirm_probe_connected(self):
        """Called when technician connects probe to the active channel pin and confirms."""
        self.channel_transition_active = False
        self.is_channel_armed = True
        self.awaiting_movement = True
        self.dynamix_state = DYN_STATE_AWAITING_MOVE
        self.auto_saved_this_row = False
        if hasattr(self, "btn_confirm_probe") and self.btn_confirm_probe is not None:
            self.btn_confirm_probe.configure(
                text=f"✓ Probe Connected: Pin {self.active_channel}-GND (Active)",
                fg_color="#059669",
                hover_color="#047857"
            )
        self.lbl_vfd_channel.configure(
            text=f"CHANNEL: {self.active_channel}-GND",
            text_color=self._get_channel_color(self.active_channel)
        )
        self.log_message(f"Probe confirmed on Pin {self.active_channel}-GND. Multimeter reading resumed.")

    def update_target_info_label(self):
        col_num = self.test_cols.get(self.active_channel, 17)
        col_letter = openpyxl.utils.get_column_letter(col_num)
        r_data = self.table_data.get(self.selected_row_idx, {})
        sp = r_data.get("sp", "-")
        tp = r_data.get("tp", f"TP {self.selected_row_idx}")
        sp_str = f"{sp}°" if sp != "-" and str(sp).replace(".", "").replace("-", "").isdigit() else str(sp)
        dir_badge = "REVERSE" if getattr(self, "input_direction", "forward") == "reverse" else "FORWARD"
        self.lbl_target_info.configure(
            text=f"Focus: Test {self.active_channel} (Col {col_letter}) | Row: {tp} ({sp_str}) [{dir_badge}]"
        )
        self.lbl_table_info.configure(
            text=f"Active Measurement: Test {self.active_channel} (Col {col_letter}) | Direction: {dir_badge}"
        )

    def set_direction(self, direction_str):
        """Programmatically switch input direction ('forward' or 'reverse') and sync UI segmented button."""
        if "reverse" in str(direction_str).lower():
            self.input_direction = "reverse"
            self.seg_direction.set("Reverse (End -> 1)")
            self.log_message(f"Input direction set to: REVERSE (TP {self.total_rows} -> TP 1)")
        else:
            self.input_direction = "forward"
            self.seg_direction.set("Forward (1 -> End)")
            self.log_message(f"Input direction set to: FORWARD (TP 1 -> TP {self.total_rows})")
        self.update_target_info_label()

    def on_direction_changed(self, choice):
        """Change row input advancement direction (Forward: 1->End or Reverse: End->1)."""
        if "Reverse" in choice:
            self.input_direction = "reverse"
            self.log_message(f"Input direction changed to: REVERSE (TP {self.total_rows} -> TP 1)")
            # If currently on row 1 and total_rows > 1, automatically move to bottom row
            if self.selected_row_idx == 1 and self.total_rows > 1:
                self.select_row(self.total_rows, from_auto_advance=False)
        else:
            self.input_direction = "forward"
            self.log_message(f"Input direction changed to: FORWARD (TP 1 -> TP {self.total_rows})")
            # If currently on bottom row and total_rows > 1, automatically move to row 1
            if self.selected_row_idx == self.total_rows and self.total_rows > 1:
                self.select_row(1, from_auto_advance=False)
        self.update_target_info_label()

    def _on_flapate_api_trigger(self, source="FlapATE API", angle=None, step=None):
        """Asynchronous callback from FlapATE client when FlapATE takes angle data."""
        self.run_on_gui_thread(self.trigger_from_api, source=source, angle=angle, step=step)

    def trigger_from_api(self, source="FlapATE API", angle=None, step=None):
        """
        Called when FlapATE or an external webhook signals that data has been taken.
        Inputs current voltage reading and advances to the next step.
        Thread-safe: can be called from main thread or background worker threads.
        """
        if threading.current_thread() is not threading.main_thread():
            self.run_on_gui_thread(self.trigger_from_api, source=source, angle=angle, step=step)
            return {
                "success": True,
                "status": "queued",
                "row": self.selected_row_idx,
                "channel": self.active_channel,
                "voltage": extract_2_decimals(self.current_live_voltage),
                "angle": angle,
                "source": source
            }

        if not getattr(self, "auto_save_enabled", True):
            return {"success": False, "message": "Auto-input is disabled in GUI"}

        if not self.is_connected and not self.sim_mode:
            self.log_message(f"[{source}] Trigger received, but multimeter is disconnected. Please connect multimeter.")
            return {"success": False, "error": "Multimeter not connected"}

        # If step is specified and valid, sync row index first
        if step is not None:
            try:
                target_step = int(step)
                if 1 <= target_step <= self.total_rows and target_step != self.selected_row_idx:
                    self.select_row(target_step, from_auto_advance=False)
            except Exception:
                pass

        val = self.current_live_voltage
        val_2dec = extract_2_decimals(val)
        row_idx = self.selected_row_idx
        ch = self.active_channel

        angle_info = ""
        if angle is not None:
            try:
                angle_info = f" | Angle: {float(angle):.2f}°"
            except Exception:
                angle_info = f" | Angle: {angle}"
        trig_src = f"{source}{angle_info}"

        self.record_and_advance(row_idx, ch, val_2dec, source=trig_src)
        return {
            "success": True,
            "row": row_idx,
            "channel": ch,
            "voltage": val_2dec,
            "angle": angle,
            "source": trig_src
        }

    def on_toggle_autosave(self):
        """Handler for Enable Auto-Save checkbox toggle."""
        is_auto = (self.chk_auto_save.get() == 1)
        self.auto_save_enabled = is_auto
        if is_auto:
            self.auto_saved_this_row = False
            self.awaiting_movement = True
            self.channel_transition_active = False
            self.dynamix_state = DYN_STATE_AWAITING_MOVE
            self.dynamix_needle_stop_time = time.time()
            self.log_message("Auto-Save ENABLED (System will auto-record when FlapATE stops at setpoint).")
            self.lbl_needle_motion.configure(
                text="FlapATE: Ready (Awaiting FlapATE stop)",
                text_color="#10b981"
            )
        else:
            self.log_message("Auto-Save DISABLED (Manual mode: Press [ENTER] or click Save Reading).")
            self.lbl_needle_motion.configure(
                text="Manual Mode: Press [ENTER] to save reading",
                text_color="#38bdf8"
            )

    on_toggle_mode_auto = on_toggle_autosave
    on_toggle_mode_manual = on_toggle_autosave

    def on_flapate_trigger_changed(self, choice=None):
        pass

    def on_reconnect_flapate(self):
        """Update FlapATE API URL and Motor IP address and reconnect."""
        url_str = self.entry_flapate_url.get().strip() if hasattr(self, "entry_flapate_url") else ""
        ip_str = self.entry_flapate_motor_ip.get().strip() if hasattr(self, "entry_flapate_motor_ip") else ""

        if url_str:
            self.flapate_url = url_str
            self.flapate_client.api_url = url_str
        if ip_str:
            self.flapate_motor_ip = ip_str
            self.flapate_client.ip = ip_str
            self.dynamix_ip = ip_str

        self.flapate_client.is_connected = True
        self.log_message(f"Connecting to FlapATE API at {self.flapate_url} & Rotary Table at {self.flapate_motor_ip}...")

        def _test():
            api_ok = self.flapate_client.test_api_connection()
            mot_ok = self.flapate_client._quick_ping()

            def _update():
                status_parts = []
                if api_ok:
                    status_parts.append(f"API Online ({self.flapate_url})")
                else:
                    status_parts.append(f"API Monitoring ({self.flapate_url})")
                if mot_ok:
                    status_parts.append(f"Motor Connected ({self.flapate_motor_ip})")
                else:
                    status_parts.append(f"Auto-Tracking ({self.flapate_motor_ip})")

                src = " | ".join(status_parts)
                self.flapate_client.connection_source = src
                self.lbl_flapate_badge.configure(
                    text=f"● FlapATE: CONNECTED [{src}]",
                    text_color="#22c55e"
                )
                self.log_message(f"FlapATE Sync Status: {src}")
            try:
                self.after(0, _update)
            except Exception:
                pass

        threading.Thread(target=_test, daemon=True).start()

    on_reconnect_dynamix = on_reconnect_flapate  # Backward compatibility alias

    def on_simulate_needle_step(self):
        """Simulate FlapATE rotary needle movement, stopping, and auto-capture (offline demo)."""
        row_data = self.table_data.get(self.selected_row_idx, {})
        sp_angle = row_data.get("sp", "-")
        try:
            target_deg = float(sp_angle)
        except Exception:
            target_deg = random.uniform(10.0, 350.0)

        self.log_message(f"[FlapATE Simulation] Rotating needle to SP: {target_deg:.2f}° (Testing Auto-Capture)...")
        self.flapate_client.trigger_simulation_move(target_angle=target_deg, duration=0.8)

    def _update_flapate_ui(self, is_conn, is_moving, angle, rpm, phase, elapsed=0.0, sp_str="", target_settle=0.2):
        """Update GUI badges and monitor card for FlapATE motion and API capture."""
        # Lightweight GUI optimization: skip redundant re-rendering when state is unchanged
        elapsed_key = round(elapsed, 1) if phase == "SETTLING" else 0
        rpm_key = round(rpm, 1) if rpm else 0
        state_key = (is_conn, is_moving, phase, sp_str, rpm_key, elapsed_key)
        if getattr(self, "_last_flapate_ui_state", None) == state_key:
            return
        self._last_flapate_ui_state = state_key

        angle_str = f"{angle:.2f}°" if (angle is not None and angle != 0.0) else (sp_str if sp_str else "--")
        src = getattr(self.flapate_client, "connection_source", "Auto-Tracking")

        # Synchronize LAN status badge in header
        if hasattr(self, "lbl_lan_badge"):
            is_lan = getattr(self.flapate_client, "is_lan_connected", False) or self.sim_mode
            if is_lan:
                src_ip = getattr(self.flapate_client, "ip", "192.168.1.20")
                self.lbl_lan_badge.configure(
                    text=f"FLAPATE LAN: ONLINE ({src_ip})",
                    text_color="#10b981"
                )
            else:
                self.lbl_lan_badge.configure(
                    text="FLAPATE: GPIB TRACKING",
                    text_color="#38bdf8"
                )

        if is_conn:
            self.lbl_flapate_badge.configure(
                text=f"FlapATE: CONNECTED [{src}]",
                text_color="#10b981"
            )
        else:
            self.lbl_flapate_badge.configure(
                text=f"FlapATE ({self.flapate_url}): CONNECTED [Auto-Tracking]",
                text_color="#10b981"
            )

        if phase == "PROBE_WAIT":
            self.lbl_needle_motion.configure(
                text=f"Probe Switch: Connect probe to Channel {self.active_channel}-GND",
                text_color="#f59e0b"
            )
            self.lbl_mon_stab.configure(
                text=f"Probe Repositioning: Connect to {self.active_channel}-GND",
                text_color="#f59e0b"
            )
            self.progress_stab.set(0.15)
            self.progress_stab.configure(progress_color="#f59e0b")
            self.lbl_stability_status.configure(
                text=f"Connect probe to Channel {self.active_channel}-GND. Auto-Save is paused.",
                text_color="#f59e0b"
            )
        elif phase == "MANUAL":
            self.lbl_needle_motion.configure(
                text="Manual Mode (Press [ENTER] to save reading)",
                text_color="#38bdf8"
            )
            self.lbl_mon_stab.configure(
                text=f"FlapATE ({sp_str}) - Ready for [ENTER]",
                text_color="#10b981"
            )
            self.progress_stab.set(1.0)
            self.progress_stab.configure(progress_color="#10b981")
            self.lbl_stability_status.configure(
                text="Monitoring voltage. Press [ENTER] to save reading",
                text_color="#38bdf8"
            )
        elif phase == "MOVING":
            self.lbl_needle_motion.configure(
                text=f"FlapATE Moving... Towards SP: {sp_str}",
                text_color="#f59e0b"
            )
            self.lbl_mon_stab.configure(
                text=f"FlapATE Rotating... Towards {sp_str}",
                text_color="#f59e0b"
            )
            self.progress_stab.set(0.35)
            self.progress_stab.configure(progress_color="#f59e0b")
            self.lbl_stability_status.configure(
                text=f"Awaiting FlapATE stop at {sp_str}...",
                text_color="#f59e0b"
            )
        elif phase == "AWAITING_MOVE":
            self.lbl_needle_motion.configure(
                text=f"Reading saved. Awaiting FlapATE motion to SP: {sp_str}",
                text_color="#38bdf8"
            )
            self.lbl_mon_stab.configure(
                text=f"Saved. Awaiting motion to {sp_str}",
                text_color="#38bdf8"
            )
            self.progress_stab.set(0.0)
            self.progress_stab.configure(progress_color="#38bdf8")
            self.lbl_stability_status.configure(
                text=f"Waiting for FlapATE to rotate to setpoint {sp_str}...",
                text_color="#38bdf8"
            )
        elif phase == "SETTLING":
            denom = max(0.01, target_settle)
            pct = min(1.0, max(0.05, elapsed / denom))
            self.lbl_needle_motion.configure(
                text=f"FlapATE Stopped at {sp_str} (Settling: {elapsed:.1f}s / {denom:.1f}s)",
                text_color="#38bdf8"
            )
            self.lbl_mon_stab.configure(
                text=f"FlapATE Stopped at {sp_str}. Acquiring ({elapsed:.1f}s)",
                text_color="#38bdf8"
            )
            self.progress_stab.set(pct)
            self.progress_stab.configure(progress_color="#38bdf8")
            self.lbl_stability_status.configure(
                text=f"FlapATE Stopped at {sp_str}. Settling voltage...",
                text_color="#38bdf8"
            )
        else:  # STOPPED
            self.lbl_needle_motion.configure(
                text=f"FlapATE Stopped at {sp_str} - Reading Captured",
                text_color="#10b981"
            )
            self.lbl_mon_stab.configure(
                text=f"FlapATE ({sp_str}) - Reading Captured",
                text_color="#10b981"
            )
            self.progress_stab.set(1.0)
            self.progress_stab.configure(progress_color="#10b981")
            self.lbl_stability_status.configure(
                text="Reading saved successfully.",
                text_color="#10b981"
            )

    _update_dynamix_ui = _update_flapate_ui  # Backward compatibility alias

    # ==========================================================================
    # AUDIO & VISUAL FEEDBACK: CONFIRMATION CHIME (OK) & CAUTION TONE (WARNING/DEVIATION)
    # ==========================================================================
    def on_toggle_sound_alert(self):
        """Enable or disable audio feedback alert upon data input."""
        if hasattr(self, "chk_sound_alert") and self.chk_sound_alert is not None:
            self.sound_enabled = (self.chk_sound_alert.get() == 1)
        else:
            self.sound_enabled = True
        if self.sound_enabled:
            self.log_message("Audio feedback: Sound alert enabled (Distinct sounds for OK vs Warning/Deviation).")
            self.play_blink_sound(is_warn=False, force=True)
        else:
            self.log_message("Audio feedback: Sound alert disabled (Muted).")

    def play_blink_sound(self, is_warn=False, force=False):
        """Play confirmation chime (green chirp) or deviation alert tone (yellow) asynchronously."""
        if not force and not getattr(self, "sound_enabled", True):
            return

        def _play():
            played = False
            target_wav = self._warn_wav if is_warn else self._blink_wav
            if not target_wav:
                try:
                    target_wav = generate_warn_sound_wav() if is_warn else generate_blink_sound_wav()
                    if is_warn:
                        self._warn_wav = target_wav
                    else:
                        self._blink_wav = target_wav
                except Exception:
                    pass

            # 1. Priority: Play in-memory WAV audio (clean sound, stereo/headphone support)
            if winsound and target_wav:
                try:
                    winsound.PlaySound(target_wav, winsound.SND_MEMORY | winsound.SND_NODEFAULT)
                    played = True
                except Exception:
                    played = False

            # 2. Fallback: Windows MessageBeep (plays through laptop speakers / headphones via system audio)
            if not played and winsound:
                try:
                    mb_type = winsound.MB_ICONHAND if is_warn else winsound.MB_OK
                    winsound.MessageBeep(mb_type)
                    played = True
                except Exception:
                    pass

            # 3. Fallback: Beep tone (distinct frequencies: low double buzz for invalid, high double chime for OK)
            if not played and winsound:
                try:
                    if is_warn:
                        winsound.Beep(480, 110)
                        time.sleep(0.03)
                        winsound.Beep(360, 130)
                    else:
                        winsound.Beep(1760, 80)
                        winsound.Beep(2637, 90)
                    played = True
                except Exception:
                    pass

            # 4. Fallback: Tkinter default bell
            if not played:
                try:
                    self.bell()
                except Exception:
                    pass

        threading.Thread(target=_play, daemon=True).start()

    def trigger_visual_blink(self, is_warn=False):
        """Visual blink effect on VFD display upon data saved (Green for OK, Amber for Deviation)."""
        try:
            if hasattr(self, "vfd_screen") and self.vfd_screen:
                is_light = getattr(self, "appearance_mode", "Dark") == "Light"
                orig_bg = getattr(self, "vfd_bg_color", ("#ffffff" if is_light else "#0b1219"))
                flash_bg = ("#fef08a" if is_light else "#2e210a") if is_warn else ("#dbeafe" if is_light else "#1e3a5f")
                trig_color = ("#b45309" if is_light else "#f59e0b") if is_warn else ("#0284c7" if is_light else "#38bdf8")
                self.vfd_screen.configure(fg_color=flash_bg)
                self.lbl_vfd_trig.configure(text_color=trig_color)
                def _restore():
                    try:
                        self.vfd_screen.configure(fg_color=orig_bg)
                        self.lbl_vfd_trig.configure(text_color=orig_bg)
                    except Exception:
                        pass
                self.after(140, _restore)
        except Exception:
            pass

    # ==========================================================================
    # LOGIC: SAVE READING PROCESS (MANUAL ENTER & AUTO-SAVE 1.6S)
    # ==========================================================================
    def manual_save_reading(self):
        """Called when technician presses SAVE button or ENTER key."""
        if not self.is_connected and not self.sim_mode:
            messagebox.showwarning("Warning", "The multimeter is not connected. Please connect the multimeter first.")
            return

        if getattr(self, "channel_transition_active", False) or not getattr(self, "is_channel_armed", True):
            self.confirm_probe_connected()

        val = extract_2_decimals(self.current_live_voltage)
        row_idx = self.selected_row_idx
        ch = self.active_channel

        self.record_and_advance(row_idx, ch, val, source="MANUAL [ENTER]")
        self.dynamix_state = DYN_STATE_AWAITING_MOVE

    def auto_save_reading(self, val, source=None):
        """Called automatically when triggered by FlapATE needle stop."""
        if getattr(self, "channel_transition_active", False) or not getattr(self, "is_channel_armed", True):
            return
        val_2dec = extract_2_decimals(val)
        row_idx = self.selected_row_idx
        ch = self.active_channel
        if not source:
            source = f"FLAPATE NEEDLE STOP ({self.flapate_motor_ip})"
        self.record_and_advance(row_idx, ch, val_2dec, source=source)

    def record_and_advance(self, row_idx, ch, val, source=""):
        """Save reading to memory, write ONLY to SYNCHRO UNDER TEST in Excel, and advance to next test point/channel."""
        val_2dec = extract_2_decimals(val)
        key = f"test_{ch.lower()}"
        self.table_data[row_idx][key] = val_2dec

        self.calculate_row_result(row_idx)
        self.update_single_table_row(row_idx)

        # Deteksi status deviasi HANYA untuk channel yang sedang aktif diuji (X, Y, atau Z)
        # Jika channel yang sedang diuji valid, suara yang berbunyi HARUS suara valid meskipun channel lain warn/invalid!
        row_data = self.table_data[row_idx]
        ch_warn_key = f"warn_{ch.lower()}"
        is_channel_warn = bool(row_data.get(ch_warn_key, False))

        # Bunyikan suara valid/invalid SESUAI CHANNEL YANG SEDANG DIUJI & kedipkan display konfirmasi
        self.play_blink_sound(is_warn=is_channel_warn)
        self.trigger_visual_blink(is_warn=is_channel_warn)

        # Save to Excel ONLY in SYNCHRO UNDER TEST columns (queued asynchronously)
        if self.excel_file_path:
            self.save_single_to_excel(row_idx, ch, val_2dec)

        tp_label = self.table_data[row_idx].get("tp", f"TP {row_idx}")
        val_str = format_2_decimals(val_2dec)
        status_tag = f" [WARN {ch}]" if is_channel_warn else " [OK]"
        self.log_message(f"[{source}] {tp_label} ({ch}-GND) = {val_str} VAC{status_tag} saved to Excel.")

        self.last_saved_val_2dec = val_2dec
        try:
            self.last_saved_raw = float(val)
        except Exception:
            self.last_saved_raw = self.current_live_voltage

        self.dynamix_state = DYN_STATE_AWAITING_MOVE
        self.auto_saved_this_row = False

        # Automatically advance to next row according to active input direction:
        # Zig-Zag / Boustrophedon mode:
        # - When Forward finishes at bottom (Row Max), next channel automatically switches to Reverse (Row Max -> 1)
        # - When Reverse finishes at top (Row 1), next channel automatically switches to Forward (Row 1 -> Max)
        if getattr(self, "input_direction", "forward") == "reverse":
            # Reverse Mode (Bottom TP to Top: e.g. TP 182 -> TP 181 -> ... -> TP 1)
            if row_idx > 1:
                next_row = row_idx - 1
                self.select_row(next_row, from_auto_advance=True)
            else:
                self.log_message(f"Channel {ch}-GND measurements are complete for all {self.total_rows} rows (Reverse mode).")
                # When finishing Reverse mode, next channel automatically switches to FORWARD mode starting at row 1
                if ch == "X":
                    self.set_direction("forward")
                    self.seg_channel.set("Y to GND")
                    self.on_channel_focus_changed("Y to GND")
                    self.select_row(1, from_auto_advance=False)
                    messagebox.showinfo(
                        "Channel X Complete -> Switching to Y (Forward)",
                        f"All {self.total_rows} rows for Channel X on '{self.sheet_name}' have been measured and saved (Reverse)!\n\n"
                        f"Direction has switched automatically to FORWARD mode (Row 1 / TP 1 -> {self.total_rows}).\n"
                        f"Channel has switched automatically to 'Y to GND'.\n\n"
                        "Please move multimeter probe to Channel Y.\n"
                        "Multimeter reading is PAUSED.\n\n"
                        "Click OK once the probe is connected to Pin Y to resume reading."
                    )
                    self.confirm_probe_connected()
                elif ch == "Y":
                    if self.has_z:
                        self.set_direction("forward")
                        self.seg_channel.set("Z to GND")
                        self.on_channel_focus_changed("Z to GND")
                        self.select_row(1, from_auto_advance=False)
                        messagebox.showinfo(
                            "Channel Y Complete -> Switching to Z (Forward)",
                            f"All {self.total_rows} rows for Channel Y on '{self.sheet_name}' have been measured and saved (Reverse)!\n\n"
                            f"Direction has switched automatically to FORWARD mode (Row 1 / TP 1 -> {self.total_rows}).\n"
                            f"Channel has switched automatically to 'Z to GND'.\n\n"
                            "Please move multimeter probe to Channel Z.\n"
                            "Multimeter reading is PAUSED.\n\n"
                            "Click OK once the probe is connected to Pin Z to resume reading."
                        )
                        self.confirm_probe_connected()
                    else:
                        messagebox.showinfo(
                            "VFlap - Test Complete",
                            f"All Synchro Flap measurements (Channels X and Y) on '{self.sheet_name}' are complete and saved to Excel!"
                        )
                elif ch == "Z":
                    messagebox.showinfo(
                        "VFlap - Test Complete",
                        f"All Synchro measurements (Channels X, Y, Z) on '{self.sheet_name}' are complete and saved to Excel!"
                    )
        else:
            # Forward Mode (Default: TP 1 -> TP 2 -> ... -> TP Max)
            if row_idx < self.total_rows:
                next_row = row_idx + 1
                self.select_row(next_row, from_auto_advance=True)
            else:
                self.log_message(f"Channel {ch}-GND measurements are complete for all {self.total_rows} rows (Forward mode).")
                # When finishing Forward mode, next channel automatically switches to REVERSE mode starting at bottom row
                if ch == "X":
                    self.set_direction("reverse")
                    self.seg_channel.set("Y to GND")
                    self.on_channel_focus_changed("Y to GND")
                    self.select_row(self.total_rows, from_auto_advance=False)
                    messagebox.showinfo(
                        "Channel X Complete -> Switching to Y (Reverse)",
                        f"All {self.total_rows} rows for Channel X on '{self.sheet_name}' have been measured and saved (Forward)!\n\n"
                        f"Direction has switched automatically to REVERSE mode (Row {self.total_rows} / TP {self.total_rows} -> 1).\n"
                        f"Channel has switched automatically to 'Y to GND'.\n\n"
                        "Please move multimeter probe to Channel Y.\n"
                        "Multimeter reading is PAUSED.\n\n"
                        "Click OK once the probe is connected to Pin Y to resume reading."
                    )
                    self.confirm_probe_connected()
                elif ch == "Y":
                    if self.has_z:
                        self.set_direction("reverse")
                        self.seg_channel.set("Z to GND")
                        self.on_channel_focus_changed("Z to GND")
                        self.select_row(self.total_rows, from_auto_advance=False)
                        messagebox.showinfo(
                            "Channel Y Complete -> Switching to Z (Reverse)",
                            f"All {self.total_rows} rows for Channel Y on '{self.sheet_name}' have been measured and saved (Forward)!\n\n"
                            f"Direction has switched automatically to REVERSE mode (Row {self.total_rows} / TP {self.total_rows} -> 1).\n"
                            f"Channel has switched automatically to 'Z to GND'.\n\n"
                            "Please move multimeter probe to Channel Z.\n"
                            "Multimeter reading is PAUSED.\n\n"
                            "Click OK once the probe is connected to Pin Z to resume reading."
                        )
                        self.confirm_probe_connected()
                    else:
                        messagebox.showinfo(
                            "VFlap - Test Complete",
                            f"All Synchro Flap measurements (Channels X and Y) on '{self.sheet_name}' are complete and saved to Excel!"
                        )
                elif ch == "Z":
                    messagebox.showinfo(
                        "VFlap - Test Complete",
                        f"All Synchro measurements (Channels X, Y, Z) on '{self.sheet_name}' are complete and saved to Excel!"
                    )

    # ==========================================================================
    # BACKGROUND POLLING & DETEKSI KESTABILAN (1.6 DETIK)
    # ==========================================================================
    def start_background_polling(self):
        self.polling_active = True
        self.poll_thread = threading.Thread(target=self._polling_loop, daemon=True)
        self.poll_thread.start()

    def _polling_loop(self):
        while self.polling_active:
            if not self.polling_active:
                break

            try:
                if self.is_connected:
                    is_transitioning = (
                        getattr(self, "channel_transition_active", False) 
                        or not getattr(self, "is_channel_armed", True)
                    )

                    if is_transitioning:
                        # Probe switch: DO NOT read multimeter while moving probe from pin X to Y
                        self.current_live_voltage = 0.0
                        try:
                            self.update_vfd_display(0.0, channel=f"{self.active_channel}-GND [PAUSED]")
                        except Exception:
                            pass
                        dyn_conn, dyn_moving, dyn_angle, dyn_rpm = self.flapate_client.poll(is_sim_mode=self.sim_mode)
                        self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, ch=self.active_channel: 
                                   self._update_flapate_ui(c, False, a, r, "PROBE_WAIT", sp_str=f"Pin {ch}"))
                        time.sleep(0.12)
                        continue

                    raw_val = self.read_multimeter_value(target_channel=self.active_channel)
                    if raw_val is not None:
                        self.current_live_voltage = raw_val
                        now = time.time()

                        # Ambil 2 angka di belakang koma secara eksak TANPA pembulatan
                        val_2dec = extract_2_decimals(raw_val)

                        try:
                            self.update_vfd_display(raw_val, channel=f"{self.active_channel}-GND")
                        except Exception:
                            pass

                        is_auto_save_active = getattr(self, "auto_save_enabled", True)

                        # Status FlapATE dari client (motor / simulation / API)
                        dyn_conn, dyn_moving, dyn_angle, dyn_rpm = self.flapate_client.poll(is_sim_mode=self.sim_mode)

                        # Perubahan tegangan antar sampel
                        prev_raw = getattr(self, "_prev_multimeter_raw", raw_val)
                        volt_diff = abs(raw_val - prev_raw)
                        self._prev_multimeter_raw = raw_val

                        target_sp = self.table_data.get(self.selected_row_idx, {}).get("sp", "-")
                        sp_str = f"{target_sp}°" if target_sp != "-" else f"TP {self.selected_row_idx}"

                        if not is_auto_save_active:
                            # Mode Manual: Multimeter reads continuously, saves strictly on [ENTER] / Save button
                            self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                       self._update_flapate_ui(c, False, a, r, "MANUAL", sp_str=s))
                        else:
                            # Mode Otomatis
                            is_lan_online = getattr(self.flapate_client, "is_lan_connected", False) or self.sim_mode

                            # 1. When connected via LAN cable: motor motion state is authoritative from FlapATE hardware
                            # 2. When only GPIB is connected: slew voltage tracking detects motion
                            if is_lan_online:
                                is_in_motion = dyn_moving
                            else:
                                has_moved_from_last = (
                                    hasattr(self, "last_saved_raw")
                                    and self.last_saved_raw is not None
                                    and abs(raw_val - self.last_saved_raw) >= 0.018
                                )
                                is_in_motion = (volt_diff > 0.015) or (getattr(self, "awaiting_movement", False) and has_moved_from_last)

                            if getattr(self, "awaiting_movement", False):
                                # Awaiting FlapATE to begin rotating towards next setpoint
                                if is_in_motion:
                                    self.awaiting_movement = False
                                    self.dynamix_state = DYN_STATE_AWAITING_STOP
                                    self.auto_saved_this_row = False
                                    self.dynamix_needle_stop_time = now
                                    self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                               self._update_flapate_ui(c, True, a, r, "MOVING", sp_str=s))
                                else:
                                    self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                               self._update_flapate_ui(c, False, a, r, "AWAITING_MOVE", sp_str=s))

                            elif self.dynamix_state == DYN_STATE_AWAITING_STOP:
                                # FlapATE is rotating, waiting for it to stop at setpoint
                                if is_in_motion:
                                    self.dynamix_needle_stop_time = now
                                    self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                               self._update_flapate_ui(c, True, a, r, "MOVING", sp_str=s))
                                else:
                                    # FlapATE stopped at setpoint! Start settle countdown
                                    self.dynamix_state = DYN_STATE_READY_TO_SAVE
                                    self.dynamix_needle_stop_time = now
                                    target_settle = 0.5 if is_lan_online else 0.9
                                    self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str, ts=target_settle: 
                                               self._update_flapate_ui(c, False, a, r, "SETTLING", elapsed=0.0, sp_str=s, target_settle=ts))

                            elif self.dynamix_state == DYN_STATE_READY_TO_SAVE:
                                # FlapATE stopped at setpoint: verify voltage is completely settled while FlapATE acquires reading
                                if is_in_motion:
                                    self.dynamix_state = DYN_STATE_AWAITING_STOP
                                    self.dynamix_needle_stop_time = now
                                    self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                               self._update_flapate_ui(c, True, a, r, "MOVING", sp_str=s))
                                else:
                                    target_settle = 0.5 if is_lan_online else 0.9
                                    settle_elapsed = now - self.dynamix_needle_stop_time
                                    if settle_elapsed < target_settle:
                                        self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, e=settle_elapsed, s=sp_str, ts=target_settle: 
                                                   self._update_flapate_ui(c, False, a, r, "SETTLING", elapsed=e, sp_str=s, target_settle=ts))
                                    else:
                                        # FlapATE has stopped and acquired reading: SAVE UNROUNDED 2-DECIMAL VALUE NOW
                                        self.after(0, lambda c=dyn_conn, a=dyn_angle, r=dyn_rpm, s=sp_str: 
                                                   self._update_flapate_ui(c, False, a, r, "STOPPED", sp_str=s))
                                        if not self.auto_saved_this_row:
                                            self.auto_saved_this_row = True
                                            self.awaiting_movement = True
                                            self.dynamix_state = DYN_STATE_AWAITING_MOVE
                                            val_to_save = val_2dec  # Exact 2-decimal truncation
                                            src_label = "FlapATE LAN" if is_lan_online else "FlapATE Stop"
                                            self.after(0, lambda v=val_to_save, s=sp_str, src=src_label: 
                                                       self.auto_save_reading(v, source=f"{src} ({s})"))
            except Exception:
                pass

            time.sleep(0.12)

    # ==========================================================================
    # LOGIC: HARDWARE VISA & SCPI READ
    # ==========================================================================
    def log_message(self, msg):
        def _append():
            ts = time.strftime("[%H:%M:%S] ")
            self.txt_log.insert("end", ts + msg + "\n")
            self.txt_log.see("end")
        try:
            self.after(0, _append)
        except Exception:
            pass

    def refresh_com_ports(self):
        self.com_combo.configure(values=["Scanning ports..."])
        self.com_combo.set("Scanning ports...")

        def _scan():
            try:
                if self.visa_rm is None:
                    self.visa_rm = pyvisa.ResourceManager()
                ports = self.visa_rm.list_resources()
                valid_ports = [p for p in ports if "ASRL" in p or "GPIB" in p or "COM" in p]
                # Prioritaskan port GPIB di urutan pertama jika terdeteksi
                valid_ports.sort(key=lambda x: (not x.startswith("GPIB"), x))
                def _update_ui():
                    if valid_ports:
                        self.com_combo.configure(values=valid_ports)
                        self.com_combo.set(valid_ports[0])
                        self.log_message(f"Detected VISA Ports: {', '.join(valid_ports)}")
                    else:
                        self.com_combo.configure(values=["No COM Found"])
                        self.com_combo.set("No COM Found")
                        self.log_message("No serial/GPIB ports detected.")
                try:
                    self.after(0, _update_ui)
                except Exception:
                    pass
            except Exception as e:
                def _err_ui():
                    self.com_combo.configure(values=["No COM Found"])
                    self.com_combo.set("No COM Found")
                    self.log_message(f"Error listing VISA ports: {e}")
                try:
                    self.after(0, _err_ui)
                except Exception:
                    pass

        threading.Thread(target=_scan, daemon=True).start()

    def toggle_connection(self):
        if not self.is_connected:
            selected_port = self.com_combo.get()

            if self.sim_mode:
                self.is_connected = True
                self.lbl_hw_badge.configure(text="SIMULATION ACTIVE", text_color="#38bdf8")
                self.btn_connect.configure(text="Disconnect Multimeter", fg_color="#dc2626", hover_color="#b91c1c")
                self.lbl_vfd_auto.configure(text_color=VFD_ACCENT)
                self.log_message("Connected in DEMO/SIMULATION Mode (Virtual Agilent 34401A).")
                return

            if "No COM" in selected_port or "Scanning" in selected_port or not selected_port:
                messagebox.showwarning("Warning", "Please select a valid COM port, or enable 'Demo / Simulation Mode'.")
                return

            try:
                self.log_message(f"Opening connection to {selected_port}...")
                self.multimeter = self.visa_rm.open_resource(selected_port)
                if selected_port.startswith("ASRL"):
                    self.multimeter.baud_rate = 9600
                    self.multimeter.data_bits = 8
                    self.multimeter.parity = pyvisa.constants.Parity.none
                    self.multimeter.stop_bits = pyvisa.constants.StopBits.two
                self.multimeter.timeout = 5000

                self.multimeter.write("*RST")
                time.sleep(0.3)
                self.multimeter.write("CONF:VOLT:AC")
                try:
                    self.multimeter.write("DET:BAND 20")
                except Exception:
                    pass

                idn = self.multimeter.query("*IDN?").strip()
                self.is_connected = True
                self.lbl_hw_badge.configure(text=f"CONNECTED ({selected_port})", text_color="#10b981")
                self.btn_connect.configure(text="Disconnect Multimeter", fg_color="#dc2626", hover_color="#b91c1c")
                self.lbl_vfd_auto.configure(text_color=VFD_ACCENT)
                self.log_message(f"Instrument Connected: {idn}")
            except Exception as e:
                self.multimeter = None
                self.is_connected = False
                self.log_message(f"Connection Failed: {e}")
                messagebox.showerror("Connection Error", f"Could not connect to instrument:\n{e}\n\nTip: You can check 'Demo / Simulation Mode' to test offline.")
        else:
            if self.multimeter:
                try:
                    self.multimeter.close()
                except:
                    pass
                self.multimeter = None

            self.is_connected = False
            self.current_live_voltage = 0.0
            self.lbl_hw_badge.configure(text="DISCONNECTED", text_color="#ef4444")
            self.btn_connect.configure(text="Connect Multimeter", fg_color="#1d4ed8", hover_color="#1e40af")
            self.lbl_vfd_auto.configure(text_color=VFD_MUTED)
            self.lbl_vfd_digits.configure(text="00.0000")
            self.lbl_stability_status.configure(text="Multimeter Disconnected", text_color="gray70")
            self.log_message("Multimeter disconnected.")

    def toggle_sim_mode(self):
        self.sim_mode = self.chk_sim_mode.get() == 1
        if self.sim_mode:
            self.is_connected = True
            self.lbl_hw_badge.configure(text="SIMULATION ACTIVE", text_color="#38bdf8")
            self.btn_connect.configure(text="Disconnect Multimeter", fg_color="#dc2626", hover_color="#b91c1c")
            self.lbl_vfd_auto.configure(text_color=VFD_ACCENT)
            self.auto_saved_this_row = False
            self.history_2dec = []
            self.stable_candidate_val = None
            self.stable_candidate_raw = None
            self.stable_candidate_start_time = 0.0
            self.awaiting_movement = False
            self.dynamix_state = DYN_STATE_READY_TO_SAVE
            self.dynamix_needle_stop_time = time.time()
            self.log_message("Simulation Mode Enabled (Virtual Agilent 34401A & FlapATE Active).")
        else:
            if self.multimeter is None:
                self.is_connected = False
                self.current_live_voltage = 0.0
                self.lbl_hw_badge.configure(text="DISCONNECTED", text_color="#ef4444")
                self.btn_connect.configure(text="Connect Multimeter", fg_color="#1d4ed8", hover_color="#1e40af")
                self.lbl_vfd_auto.configure(text_color=VFD_MUTED)
                self.lbl_vfd_digits.configure(text="00.0000")
                self.lbl_stability_status.configure(text="Multimeter Disconnected", text_color="gray70")
                self.lbl_mon_stab.configure(text="Signal Stability: Disconnected", text_color="gray70")
                self.progress_stab.set(0.0)
            self.log_message("Simulation Mode Disabled.")

    def read_multimeter_value(self, target_channel=None):
        if getattr(self, "channel_transition_active", False) or not getattr(self, "is_channel_armed", True):
            # Multimeter reading blocked during probe switch / repositioning
            return None

        if self.sim_mode or self.multimeter is None:
            row_data = self.table_data.get(self.selected_row_idx, {})
            ref_key = f"ref_{target_channel.lower()}" if target_channel else "ref_x"
            ref_val = row_data.get(ref_key, "-")
            nominal = 5.0
            if isinstance(ref_val, (int, float)):
                nominal = float(ref_val)
            elif isinstance(ref_val, str) and ref_val.replace(".", "", 1).replace("-", "", 1).isdigit():
                nominal = float(ref_val)

            # Ambil 2 desimal dasar dari nominal tanpa pembulatan
            base_2dec = extract_2_decimals(nominal)
            # Variasi halus hanya pada digit ke-3 dan ke-4 (sub-millivolt)
            # sehingga 2 angka di belakang koma tetap persis sama dan stabil
            sub_jitter = random.uniform(0.0012, 0.0035)
            sim_val = round(base_2dec + sub_jitter, 4)
            return sim_val

        try:
            raw = self.multimeter.query("READ?").strip()
            return float(raw)
        except Exception as e:
            self.log_message(f"SCPI Read Error: {e}")
            return None

    def update_vfd_display(self, value, channel=""):
        if getattr(self, "channel_transition_active", False) or not getattr(self, "is_channel_armed", True):
            formatted = "00.0000"
        else:
            formatted = f"{value:07.4f}"
        if getattr(self, "_last_vfd_digits", None) == formatted and getattr(self, "_last_vfd_chan", None) == channel:
            return
        self._last_vfd_digits = formatted
        self._last_vfd_chan = channel

        def _update():
            self.lbl_vfd_digits.configure(text=formatted)
            if channel:
                self.lbl_vfd_channel.configure(text=f"CHANNEL: {channel}")
        self.after(0, _update)

    # ==========================================================================
    # DEVIATION CALCULATION & STATUS (YELLOW if deviation > threshold)
    # ==========================================================================
    DEVIATION_THRESHOLD = 0.5  # VAC - deviation threshold before flagged

    def calculate_row_result(self, row_idx):
        """Calculate deviation status (WARN/OK) per row based on Test vs Ref difference."""
        row = self.table_data[row_idx]
        tx = row["test_x"]
        ty = row["test_y"]
        tz = row["test_z"]
        T = self.DEVIATION_THRESHOLD

        def _is_warn(test_val, ref_val):
            """True if |test - ref| > threshold, both must be numerical."""
            try:
                if test_val == "-" or ref_val == "-":
                    return False
                return abs(float(test_val) - float(ref_val)) > T
            except (ValueError, TypeError):
                return False

        warn_x = _is_warn(tx, row["ref_x"])
        warn_y = _is_warn(ty, row["ref_y"])
        warn_z = _is_warn(tz, row["ref_z"]) if self.has_z else False

        row["warn_x"] = warn_x
        row["warn_y"] = warn_y
        row["warn_z"] = warn_z

        has_any_measurement = (tx != "-" or ty != "-" or (self.has_z and tz != "-"))

        if not has_any_measurement:
            row["status"] = "Pending"
            row["is_warn"] = False
        elif warn_x or warn_y or warn_z:
            # Determine which channels deviated
            channels = []
            if warn_x: channels.append("X")
            if warn_y: channels.append("Y")
            if warn_z: channels.append("Z")
            row["status"] = "WARN (" + ",".join(channels) + ")"
            row["is_warn"] = True
        else:
            row["status"] = "OK"
            row["is_warn"] = False

    # ==========================================================================
    # EXCEL INTEGRATION (ONLY MODIFIES SYNCHRO UNDER TEST COLUMNS!)
    # ==========================================================================
    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
        if path:
            self.excel_file_path = path
            self.lbl_excel_file.configure(
                text=os.path.basename(path),
                text_color="#38bdf8"
            )
            self.log_message(f"Linked Excel file: {path}")

            try:
                wb = openpyxl.load_workbook(path, read_only=True)
                sheets = wb.sheetnames
                wb.close()
                filtered = [s for s in sheets if "ACCURACY" in s.upper() or "CAM OUTPUT" in s.upper()]
                display_sheets = filtered if filtered else sheets
                self.sheet_combo.configure(values=display_sheets)
                if self.sheet_name in display_sheets:
                    self.sheet_combo.set(self.sheet_name)
                else:
                    self.sheet_combo.set(display_sheets[0])
                    self.on_sheet_changed(display_sheets[0])
            except Exception as e:
                self.log_message(f"Warning reading sheets: {e}")

    def apply_excel_conditional_formatting(self, wb):
        """
        Apply YELLOW conditional formatting to SYNCHRO UNDER TEST columns
        across all ACCURACY and CAM OUTPUT sheets if |Test - Ref| > DEVIATION_THRESHOLD.
        Called every time the Excel file is saved by the application.
        """
        from openpyxl.styles.differential import DifferentialStyle
        from openpyxl.styles import PatternFill, Font
        from openpyxl.formatting.rule import Rule
        from openpyxl.utils import get_column_letter

        T = self.DEVIATION_THRESHOLD
        YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
        DARK_FONT = Font(color="000000", bold=True)
        sheet_configs = {}
        for sheet_name, cfg in SHEET_CONFIGS.items():
            sheet_configs.setdefault(
                sheet_name.replace(" ", "").upper(),
                cfg
            )

        for sheet_name in wb.sheetnames:
            cfg = sheet_configs.get(sheet_name.strip().replace(" ", "").upper())
            if cfg is None:
                continue
            ws = wb[sheet_name]
            s_row = cfg["start_row"]
            total_r = cfg.get("total_rows", self.total_rows)
            e_row = s_row + total_r - 1

            for ch, test_col in cfg["test_cols"].items():
                ref_col = cfg["ref_cols"].get(ch)
                if not ref_col:
                    continue
                test_let = get_column_letter(test_col)
                ref_let = get_column_letter(ref_col)
                cell_range = f"{test_let}{s_row}:{test_let}{e_row}"
                formula = f"AND(ISNUMBER({test_let}{s_row}),ISNUMBER({ref_let}{s_row}),ABS({test_let}{s_row}-{ref_let}{s_row})>{T})"

                # Remove previous rules with identical range to prevent stacking
                for existing_rule in list(ws.conditional_formatting):
                    if str(existing_rule.sqref) == cell_range:
                        del ws.conditional_formatting[existing_rule.sqref]

                dxf = DifferentialStyle(fill=YELLOW_FILL, font=DARK_FONT)
                rule = Rule(type="expression", dxf=dxf, formula=[formula])
                ws.conditional_formatting.add(cell_range, rule)

    def _excel_writer_worker(self):
        """Dedicated background worker to write rows to Excel without blocking GUI."""
        while getattr(self, "polling_active", True):
            try:
                task = self.excel_queue.get(timeout=0.5)
            except Exception:
                continue

            if task is None:
                break

            row_idx, channel, value = task
            try:
                self._do_save_single_to_excel(row_idx, channel, value)
            except Exception as e:
                self.log_message(f"Excel background write error: {e}")
            finally:
                self.excel_queue.task_done()

    def save_single_to_excel(self, row_idx, channel, value):
        """Enqueue single value to be saved to Excel asynchronously (guarantees zero GUI freeze)."""
        if hasattr(self, "excel_queue") and self.excel_queue is not None:
            self.excel_queue.put((row_idx, channel, value))
        else:
            self._do_save_single_to_excel(row_idx, channel, value)

    def _do_save_single_to_excel(self, row_idx, channel, value):
        """Save single value ONLY to SYNCHRO UNDER TEST column in the target row."""
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            return

        try:
            wb = openpyxl.load_workbook(self.excel_file_path)
            # Find sheet
            ws = None
            for s in wb.sheetnames:
                if s.replace(" ", "").upper() == self.sheet_name.replace(" ", "").upper():
                    ws = wb[s]
                    break
            if ws is None:
                ws = wb[self.sheet_name]

            target_row = self.start_row + (row_idx - 1)
            target_col = self.test_cols[channel]

            ws.cell(row=target_row, column=target_col, value=value)

            # Ensure D142 (TP 132 Ref X) is populated if missing: 12.73 for FLAP, 6.25 for STALL / AUTOPILOT
            for s in wb.sheetnames:
                su = s.replace(" ", "").upper()
                if su in ("CAMOUTPUTSTALL", "CAMOUTPUTAUTOPILOT", "CAMOUTPUTAUTOPILOT"):
                    if wb[s].cell(row=142, column=4).value in (None, "", "-"):
                        wb[s].cell(row=142, column=4, value=6.25)
                elif su == "CAMOUTPUTFLAP":
                    if wb[s].cell(row=142, column=4).value in (None, "", "-"):
                        wb[s].cell(row=142, column=4, value=12.73)

            # Apply conditional formatting so deviating cells highlight properly
            self.apply_excel_conditional_formatting(wb)

            # Atomic save to temporary file first, then replace (guarantees file never gets corrupted)
            tmp_path = self.excel_file_path + ".tmp"
            wb.save(tmp_path)
            if os.path.exists(tmp_path):
                os.replace(tmp_path, self.excel_file_path)

            col_letter = openpyxl.utils.get_column_letter(target_col)
            self.log_message(f"Excel Saved: [{self.sheet_name}] Row {target_row}, Col {col_letter} ({channel}) = {value} V")
        except PermissionError:
            self.log_message("Excel Locked: File is open in Microsoft Excel.")
            self.after(0, lambda: messagebox.showerror(
                "Excel Locked",
                f"File '{os.path.basename(self.excel_file_path)}' is open in Microsoft Excel.\n"
                "Save and close the workbook in Excel so the application can save data."
            ))
        except Exception as e:
            self.log_message(f"Excel Save Error: {e}")

    def save_all_to_excel(self):
        """Save all rows only to the SYNCHRO UNDER TEST columns."""
        if not self.excel_file_path:
            messagebox.showinfo("Select Excel File", "Please link an Excel file first.")
            self.browse_excel()
            if not self.excel_file_path:
                return

        try:
            wb = openpyxl.load_workbook(self.excel_file_path)
            ws = None
            for s in wb.sheetnames:
                if s.replace(" ", "").upper() == self.sheet_name.replace(" ", "").upper():
                    ws = wb[s]
                    break
            if ws is None:
                ws = wb[self.sheet_name]

            saved_count = 0
            for row_idx in range(1, self.total_rows + 1):
                target_row = self.start_row + (row_idx - 1)
                data = self.table_data[row_idx]

                for ch, col_num in self.test_cols.items():
                    val = data.get(f"test_{ch.lower()}")
                    if val != "-" and val is not None:
                        try:
                            ws.cell(row=target_row, column=col_num, value=float(extract_2_decimals(val)))
                            saved_count += 1
                        except ValueError:
                            pass

            # Ensure D142 (TP 132 Ref X) is populated if missing: 12.73 for FLAP, 6.25 for STALL / AUTOPILOT
            for s in wb.sheetnames:
                su = s.replace(" ", "").upper()
                if su in ("CAMOUTPUTSTALL", "CAMOUTPUTAUTOPILOT", "CAMOUTPUTAUTOPILOT"):
                    if wb[s].cell(row=142, column=4).value in (None, "", "-"):
                        wb[s].cell(row=142, column=4, value=6.25)
                elif su == "CAMOUTPUTFLAP":
                    if wb[s].cell(row=142, column=4).value in (None, "", "-"):
                        wb[s].cell(row=142, column=4, value=12.73)

            # Apply yellow conditional formatting to all sheets
            self.apply_excel_conditional_formatting(wb)

            # Atomic save to temporary file first, then replace (guarantees file never gets corrupted)
            tmp_path = self.excel_file_path + ".tmp"
            wb.save(tmp_path)
            if os.path.exists(tmp_path):
                os.replace(tmp_path, self.excel_file_path)
            self.log_message(f"Saved {saved_count} values to sheet '{self.sheet_name}'.")
            messagebox.showinfo("Success", f"All SYNCHRO UNDER TEST values were saved to:\n{self.sheet_name}")
        except PermissionError:
            messagebox.showerror("Error", "The file is open in Microsoft Excel. Close it and try again.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save to Excel:\n{e}")

    def load_from_excel(self, quiet=False):
        """Read reference data and Under Test data from active sheet in Excel."""
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            if not quiet:
                messagebox.showinfo("Select Excel File", "Please select an Excel file to load data from.")
                self.browse_excel()
            return

        try:
            wb = openpyxl.load_workbook(self.excel_file_path, data_only=True)
            ws = None
            for s in wb.sheetnames:
                if s.replace(" ", "").upper() == self.sheet_name.replace(" ", "").upper():
                    ws = wb[s]
                    break
            if ws is None:
                if not quiet:
                    messagebox.showwarning("Sheet Not Found", f"Sheet '{self.sheet_name}' was not found in workbook.")
                return

            # Auto-detect TEST header position in Column L if there is a row shift (ACCURACY sheets only)
            if "ACCURACY" in self.sheet_name.upper():
                for r in range(9, 16):
                    cell_val = ws.cell(row=r, column=12).value
                    if cell_val and str(cell_val).strip().upper() == "TEST":
                        self.start_row = r + 1
                        break
            else:
                self.start_row = self.cur_cfg["start_row"]

            ref_cols = self.cur_cfg["ref_cols"]
            test_cols = self.cur_cfg["test_cols"]
            tp_col = self.cur_cfg.get("tp_col", 1)
            sp_col = self.cur_cfg["sp_col"]

            for i in range(1, self.total_rows + 1):
                r = self.start_row + (i - 1)
                tp_val = ws.cell(row=r, column=tp_col).value
                sp_val = ws.cell(row=r, column=sp_col).value
                rx = ws.cell(row=r, column=ref_cols["X"]).value
                ry = ws.cell(row=r, column=ref_cols["Y"]).value
                rz = ws.cell(row=r, column=ref_cols["Z"]).value if self.has_z else "-"

                tx = ws.cell(row=r, column=test_cols["X"]).value
                ty = ws.cell(row=r, column=test_cols["Y"]).value
                tz = ws.cell(row=r, column=test_cols["Z"]).value if self.has_z else "-"

                if tp_val is not None and str(tp_val).strip() != "":
                    tp_str = str(tp_val).strip()
                    if tp_str.isdigit():
                        tp_str = f"TP {tp_str}"
                    self.table_data[i]["tp"] = tp_str

                if sp_val is not None:
                    self.table_data[i]["sp"] = self.clean_excel_num(sp_val)
                # Load X Ref directly from Excel; fallback to standard adjustment for TP 132 only if cell is empty
                if rx is not None:
                    self.table_data[i]["ref_x"] = self.clean_excel_num(rx)
                elif "CAM OUTPUT" in self.sheet_name.upper() and i == 132:
                    self.table_data[i]["ref_x"] = 12.73 if "FLAP" in self.sheet_name.upper() else 6.25
                if ry is not None:
                    self.table_data[i]["ref_y"] = self.clean_excel_num(ry)
                if self.has_z and rz is not None:
                    self.table_data[i]["ref_z"] = self.clean_excel_num(rz)
                elif not self.has_z:
                    self.table_data[i]["ref_z"] = "-"

                if tx is not None:
                    self.table_data[i]["test_x"] = self.clean_excel_num(tx)
                if ty is not None:
                    self.table_data[i]["test_y"] = self.clean_excel_num(ty)
                if self.has_z and tz is not None:
                    self.table_data[i]["test_z"] = self.clean_excel_num(tz)
                elif not self.has_z:
                    self.table_data[i]["test_z"] = "-"

                self.calculate_row_result(i)

            self.update_table_view()
            self.log_message(f"Data loaded from sheet '{self.sheet_name}' (Start Row: {self.start_row}, Total Rows: {self.total_rows}).")
            if not quiet:
                messagebox.showinfo("Data Loaded", f"Data from '{self.sheet_name}' was loaded from Excel.")
        except Exception as e:
            self.log_message(f"Error loading Excel data: {e}")
            if not quiet:
                messagebox.showerror("Error", f"Failed to read the Excel file:\n{e}")

    def clean_excel_num(self, val):
        """Clean string unit formats like '5.04 VAC' or '46.0 mVAC' and extract 2 decimals without rounding."""
        if val is None or val == "-":
            return "-"
        if isinstance(val, (int, float)):
            return extract_2_decimals(val)
        s = str(val).strip()
        # Clean unit string
        s = s.replace("VAC", "").replace("mVAC", "").replace("V", "").strip()
        s = s.replace(",", ".")
        try:
            return extract_2_decimals(float(s))
        except ValueError:
            return s

    def clear_test_results(self):
        confirm = messagebox.askyesno("Confirm Clear", f"Clear all SYNCHRO UNDER TEST results for the '{self.sheet_name}' sheet?")
        if confirm:
            for i in range(1, self.total_rows + 1):
                self.table_data[i]["test_x"] = "-"
                self.table_data[i]["test_y"] = "-"
                self.table_data[i]["test_z"] = "-"
                self.table_data[i]["status"] = "Pending"
                self.table_data[i]["is_warn"] = False
                self.table_data[i]["warn_x"] = False
                self.table_data[i]["warn_y"] = False
                self.table_data[i]["warn_z"] = False
            self.update_table_view()
            self.log_message(f"Measurement results for sheet '{self.sheet_name}' have been cleared.")

    # ==========================================================================
    # TABLE MANAGEMENT & ROW SELECTION (ANTI-RECURSION)
    # ==========================================================================
    def _get_row_tree_values(self, idx):
        """Format row values and status tag for Treeview display."""
        r = self.table_data.get(idx, {})
        warn_x = r.get("warn_x", False)
        warn_y = r.get("warn_y", False)
        warn_z = r.get("warn_z", False) if self.has_z else False
        is_all_warn = (warn_x and warn_y and warn_z) if self.has_z else (warn_x and warn_y)

        # Entire row yellow ONLY when all channels (X, Y, Z or X, Y) error
        if is_all_warn:
            tag = "even_all_warn" if idx % 2 == 0 else "odd_all_warn"
        else:
            tag = "even_ok" if idx % 2 == 0 else "odd_ok"

        def _fmt(val):
            return format_2_decimals(val)

        tp_val = r.get("tp", f"TP {idx}")
        sp_val = r.get("sp", "-")
        sp_str = f"{sp_val}°" if sp_val != "-" and str(sp_val).replace(".", "").replace("-", "").isdigit() else str(sp_val)

        vals = (
            tp_val,
            sp_str,
            _fmt(r.get('ref_x', "-")),
            _fmt(r.get('ref_y', "-")),
            _fmt(r.get('ref_z', "-")) if self.has_z else "-",
            _fmt(r.get('test_x', "-")),
            _fmt(r.get('test_y', "-")),
            _fmt(r.get('test_z', "-")) if self.has_z else "-",
            r.get("status", "Pending"),
        )
        return vals, tag

    def update_single_table_row(self, row_idx):
        """Ultra-fast in-place single row update without rebuilding the Treeview."""
        if not (1 <= row_idx <= self.total_rows):
            return
        str_idx = str(row_idx)
        if self.tree.exists(str_idx):
            vals, tag = self._get_row_tree_values(row_idx)
            self.tree.item(str_idx, values=vals, tags=(tag,))
            self.after_idle(self._refresh_cell_highlights)

    def update_table_view(self):
        """Update entire table view. In-place update if row count matches for maximum performance."""
        self._ignore_tree_event = True
        try:
            children = self.tree.get_children()
            if len(children) == self.total_rows:
                # Fast in-place update of all rows without deleting and recreating
                for idx in range(1, self.total_rows + 1):
                    str_idx = str(idx)
                    vals, tag = self._get_row_tree_values(idx)
                    self.tree.item(str_idx, values=vals, tags=(tag,))
            else:
                for item in children:
                    self.tree.delete(item)

                for idx in range(1, self.total_rows + 1):
                    vals, tag = self._get_row_tree_values(idx)
                    self.tree.insert(
                        "",
                        "end",
                        iid=str(idx),
                        values=vals,
                        tags=(tag,)
                    )

            if str(self.selected_row_idx) in self.tree.get_children():
                self.tree.selection_set(str(self.selected_row_idx))
                self.tree.focus(str(self.selected_row_idx))
            self.after_idle(self._refresh_cell_highlights)
        finally:
            self.after_idle(self._reset_ignore_tree_event)

    def _reset_ignore_tree_event(self):
        self._ignore_tree_event = False

    def on_tree_row_selected(self, event):
        if getattr(self, "_ignore_tree_event", False):
            return
        sel = self.tree.selection()
        if sel:
            try:
                idx = int(sel[0])
                if idx == self.selected_row_idx:
                    return
                self.select_row(idx, from_auto_advance=False)
            except Exception:
                pass

    def select_row(self, idx, from_auto_advance=False):
        if 1 <= idx <= self.total_rows:
            self.selected_row_idx = idx
            self.row_select_time = time.time()
            self.auto_saved_this_row = False
            self.history_2dec = []
            self.stable_candidate_val = None
            self.stable_candidate_raw = None
            self.stable_candidate_start_time = 0.0
            self.is_value_stable = False
            self.stable_duration = 0.0
            if from_auto_advance:
                self.awaiting_movement = True
                self.dynamix_state = DYN_STATE_AWAITING_MOVE
            else:
                self.awaiting_movement = True
                self.last_saved_val_2dec = None
                self.last_saved_raw = None
                self.dynamix_state = DYN_STATE_AWAITING_MOVE
                self.dynamix_needle_stop_time = time.time()

            self.update_target_info_label()
            str_idx = str(idx)
            current_sel = self.tree.selection()
            if current_sel != (str_idx,):
                self._ignore_tree_event = True
                try:
                    if str_idx in self.tree.get_children():
                        self.tree.selection_set(str_idx)
                        self.tree.focus(str_idx)
                        self.tree.see(str_idx)
                finally:
                    self.after_idle(self._reset_ignore_tree_event)

    def move_test_point(self, direction):
        """Move active test point with up/down arrow keys."""
        next_row = self.selected_row_idx + direction
        if 1 <= next_row <= self.total_rows:
            self.select_row(next_row, from_auto_advance=False)
        return "break"


if __name__ == "__main__":
    app = SynchroApp()
    app.mainloop()
