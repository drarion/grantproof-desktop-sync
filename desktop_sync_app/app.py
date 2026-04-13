from __future__ import annotations

import base64
import io
import json
import os
import socket
import subprocess
import sys
import threading
import uuid
from dataclasses import asdict, dataclass
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox
import tkinter as tk
from tkinter import ttk
from urllib.parse import urlencode

import qrcode
from PIL import Image, ImageOps, ImageTk

from reporting import GrantProofReportEngine

APP_NAME = "GrantProof Desktop Sync"
APP_VERSION = "9.3.0"
COMPANY_NAME = "NrD Studio"
COPYRIGHT_YEAR = "2026"
DEFAULT_PORT = 8765
SETTINGS_DIR = Path.home() / ".grantproof_desktop_sync"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"
DEFAULT_FOLDER = Path.home() / "Documents" / "GrantProof"

BG = "#f5f7fb"
CARD = "#ffffff"
CARD_ALT = "#fbfcff"
BORDER = "#d9e1ee"
TEXT = "#1f2937"
MUTED = "#667085"
BLUE = "#173b72"
BLUE_SOFT = "#2d5ca6"
ORANGE = "#f29b38"
ORANGE_SOFT = "#fff3e6"
GREEN = "#1fa463"
GREEN_SOFT = "#ecfdf3"
GREY_SOFT = "#eef2f8"


def resource_path(relative: str) -> Path:
    candidates = [
        Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)) / relative,
        Path(__file__).resolve().parent / relative,
        Path(__file__).resolve().parent / relative.replace('desktop_sync_app/', ''),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def local_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"


@dataclass
class AppSettings:
    workstation_name: str = "Poste GrantProof"
    base_folder: str = str(DEFAULT_FOLDER)
    port: int = DEFAULT_PORT
    pair_token: str = ""
    device_id: str = ""

    @classmethod
    def load(cls) -> "AppSettings":
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        if SETTINGS_FILE.exists():
            data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            settings = cls(**data)
        else:
            settings = cls()
        if not settings.pair_token:
            settings.pair_token = uuid.uuid4().hex
        if not settings.device_id:
            settings.device_id = uuid.uuid4().hex[:12]
        Path(settings.base_folder).mkdir(parents=True, exist_ok=True)
        settings.save()
        return settings

    def save(self) -> None:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")


class SyncRequestHandler(BaseHTTPRequestHandler):
    server_version = f"GrantProofDesktopSync/{APP_VERSION}"

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-GrantProof-Token")
        self.end_headers()

    def do_GET(self):
        if self.path == "/status":
            app = self.server.app  # type: ignore[attr-defined]
            self._send_json(
                {
                    "name": app.settings.workstation_name,
                    "device_id": app.settings.device_id,
                    "port": app.settings.port,
                    "ip": local_ip(),
                    "status": "ready",
                    "version": APP_VERSION,
                }
            )
            return
        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self):
        app = self.server.app  # type: ignore[attr-defined]
        if self.path == "/pair":
            token = self.headers.get("X-GrantProof-Token", "")
            if token != app.settings.pair_token:
                self._send_json({"error": "Invalid token"}, status=401)
                return
            self._send_json(
                {
                    "ok": True,
                    "name": app.settings.workstation_name,
                    "device_id": app.settings.device_id,
                    "base_folder": app.settings.base_folder,
                    "version": APP_VERSION,
                }
            )
            app.append_log("Nouveau jumelage validé.")
            return

        if self.path == "/upload":
            token = self.headers.get("X-GrantProof-Token", "")
            if token != app.settings.pair_token:
                self._send_json({"error": "Invalid token"}, status=401)
                return
            try:
                payload = self.rfile.read(int(self.headers.get("Content-Length", "0")))
                data = json.loads(payload.decode("utf-8"))
                rel = data.get("relative_path", "inbox/untitled.txt")
                content = data.get("content_base64")
                if not isinstance(rel, str) or not content:
                    raise ValueError("Missing relative_path or content_base64")
                raw = base64.b64decode(content)
                safe_rel = Path(rel)
                target = Path(app.settings.base_folder) / safe_rel
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(raw)
                app.append_log(f"Fichier reçu : {target}")
                app.handle_saved_file(target)
                self._send_json({"ok": True, "saved_to": str(target)})
                return
            except Exception as exc:  # noqa: BLE001
                app.append_log(f"Erreur upload : {exc}")
                self._send_json({"error": str(exc)}, status=400)
                return

        if self.path == "/reports/rebuild":
            token = self.headers.get("X-GrantProof-Token", "")
            if token != app.settings.pair_token:
                self._send_json({"error": "Invalid token"}, status=401)
                return
            try:
                payload = self.rfile.read(int(self.headers.get("Content-Length", "0")))
                data = json.loads(payload.decode("utf-8")) if payload else {}
                project_code = data.get("project_code") if isinstance(data, dict) else None
                report_language = data.get("report_language") if isinstance(data, dict) else None
                clean_project_code = project_code if isinstance(project_code, str) and project_code.strip() else None
                clean_report_language = report_language if isinstance(report_language, str) and report_language.strip() else None
                app.rebuild_reports(clean_project_code, clean_report_language)
                self._send_json({
                    "ok": True,
                    "project_code": clean_project_code or "all",
                    "report_language": clean_report_language or "auto",
                })
                return
            except Exception as exc:  # noqa: BLE001
                self._send_json({"error": str(exc)}, status=400)
                return

        self._send_json({"error": "Not found"}, status=404)

    def log_message(self, format: str, *args):
        pass


class GrantProofDesktopSyncApp:
    def __init__(self) -> None:
        self.settings = AppSettings.load()
        self.report_engine = GrantProofReportEngine(Path(self.settings.base_folder))
        self.root = Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1180x820")
        self.root.minsize(1080, 760)
        self.root.configure(bg=BG)

        self.workstation_name = StringVar(value=self.settings.workstation_name)
        self.folder_path = StringVar(value=self.settings.base_folder)
        self.port = StringVar(value=str(self.settings.port))
        self.running = BooleanVar(value=False)
        self.server: ThreadingHTTPServer | None = None
        self.server_thread: threading.Thread | None = None

        self.qr_photo: ImageTk.PhotoImage | None = None
        self.logo_horizontal: ImageTk.PhotoImage | None = None
        self.logo_square: ImageTk.PhotoImage | None = None
        self.log_text: tk.Text | None = None
        self.qr_label: ttk.Label | None = None
        self.address_value: ttk.Label | None = None
        self.status_value: ttk.Label | None = None
        self.manual_code_text: tk.Text | None = None
        self.short_status_value: ttk.Label | None = None
        self.next_step_value: ttk.Label | None = None
        self.workspace_value: ttk.Label | None = None
        self.footer_value: ttk.Label | None = None

        self._configure_styles()
        self._load_brand_assets()
        self._build_ui()
        self.report_engine.ensure_root_files()
        self._refresh_pairing_material()
        self._refresh_status_banner()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _configure_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("App.TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Header.TFrame", background=CARD)
        style.configure("Title.TLabel", background=CARD, foreground=BLUE, font=("Segoe UI", 25, "bold"))
        style.configure("Sub.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 11))
        style.configure("CardTitle.TLabel", background=CARD, foreground=BLUE, font=("Segoe UI", 16, "bold"))
        style.configure("Body.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Footer.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("BluePill.TLabel", background=GREY_SOFT, foreground=BLUE, font=("Segoe UI", 9, "bold"))
        style.configure("GreenPill.TLabel", background=GREEN_SOFT, foreground=GREEN, font=("Segoe UI", 9, "bold"))
        style.configure("OrangePill.TLabel", background=ORANGE_SOFT, foreground=ORANGE, font=("Segoe UI", 9, "bold"))
        style.configure("Primary.TButton", background=BLUE, foreground="white", borderwidth=0, focusthickness=3, focuscolor=BLUE)
        style.map("Primary.TButton", background=[("active", BLUE_SOFT), ("disabled", "#8da2c6")])
        style.configure("Secondary.TButton", background=ORANGE, foreground=TEXT, borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#f5ad57")])
        style.configure("Quiet.TButton", background=CARD_ALT, foreground=TEXT, borderwidth=1)
        style.map("Quiet.TButton", background=[("active", GREY_SOFT)])
        style.configure("Danger.TButton", background="#f2f4f7", foreground="#b42318", borderwidth=1)
        style.map("Danger.TButton", background=[("active", "#fef3f2")])
        style.configure("Modern.TEntry", fieldbackground=CARD_ALT, foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=8)

    def _load_brand_assets(self) -> None:
        try:
            logo_h = Image.open(resource_path("desktop_sync_app/assets/logo_horizontal.png"))
        except Exception:
            logo_h = Image.new("RGBA", (700, 146), "white")
        try:
            logo_s = Image.open(resource_path("desktop_sync_app/assets/logo_square.png"))
        except Exception:
            logo_s = Image.new("RGBA", (512, 512), "white")
        logo_h = ImageOps.contain(logo_h.convert("RGBA"), (300, 64))
        logo_s = ImageOps.contain(logo_s.convert("RGBA"), (64, 64))
        self.logo_horizontal = ImageTk.PhotoImage(logo_h)
        self.logo_square = ImageTk.PhotoImage(logo_s)
        try:
            self.root.iconphoto(True, self.logo_square)
        except Exception:
            pass

    def _card(self, parent, *, padding: int = 18, style: str = "Card.TFrame"):
        frame = ttk.Frame(parent, style=style, padding=padding)
        frame.configure(borderwidth=0)
        return frame

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=24)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header_card = self._card(outer, padding=22, style="Header.TFrame")
        header_card.grid(row=0, column=0, sticky="ew")
        header_card.columnconfigure(1, weight=1)
        if self.logo_horizontal is not None:
            ttk.Label(header_card, image=self.logo_horizontal, style="Sub.TLabel").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 18))
        ttk.Label(header_card, text="Compagnon PC local-first pour la réception, la synchronisation et la génération de rapports premium.", style="Sub.TLabel", wraplength=720).grid(row=0, column=1, sticky="w")
        self.workspace_value = ttk.Label(header_card, text="Configuration prête à finaliser", style="OrangePill.TLabel", padding=(12, 8))
        self.workspace_value.grid(row=0, column=2, sticky="e")
        self.next_step_value = ttk.Label(
            header_card,
            text="Ordre recommandé : 1. Enregistrer la configuration  •  2. Démarrer le compagnon  •  3. Scanner le QR dans GrantProof Mobile",
            style="Sub.TLabel",
            wraplength=760,
        )
        self.next_step_value.grid(row=1, column=1, columnspan=2, sticky="w", pady=(8, 0))

        body = ttk.Frame(outer, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew", pady=(18, 0))
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        config_card = self._card(body, padding=20)
        config_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        config_card.columnconfigure(0, weight=1)
        config_card.columnconfigure(1, weight=0)
        ttk.Label(config_card, text="Étape 1 — Configurez le compagnon", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(config_card, text="Choisissez le nom du poste, le dossier local GrantProof et le port, puis enregistrez avant de démarrer le serveur local.", style="Muted.TLabel", wraplength=500).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 16))

        ttk.Label(config_card, text="Nom du poste", style="Body.TLabel").grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Entry(config_card, textvariable=self.workstation_name, style="Modern.TEntry", width=42).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 14))

        ttk.Label(config_card, text="Dossier local GrantProof", style="Body.TLabel").grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Entry(config_card, textvariable=self.folder_path, style="Modern.TEntry", width=42).grid(row=5, column=0, sticky="ew", pady=(6, 14), padx=(0, 12))
        ttk.Button(config_card, text="Choisir…", command=self.choose_folder, style="Quiet.TButton").grid(row=5, column=1, sticky="ew", pady=(6, 14))

        ttk.Label(config_card, text="Port local", style="Body.TLabel").grid(row=6, column=0, columnspan=2, sticky="w")
        ttk.Entry(config_card, textvariable=self.port, style="Modern.TEntry", width=12).grid(row=7, column=0, sticky="w", pady=(6, 16))

        actions_row_1 = ttk.Frame(config_card, style="Card.TFrame")
        actions_row_1.grid(row=8, column=0, columnspan=2, sticky="ew")
        ttk.Button(actions_row_1, text="1. Enregistrer", command=self.save_settings, style="Secondary.TButton").pack(side="left")
        ttk.Button(actions_row_1, text="2. Démarrer", command=self.start_server, style="Primary.TButton").pack(side="left", padx=10)
        ttk.Button(actions_row_1, text="Arrêter", command=self.stop_server, style="Danger.TButton").pack(side="left")

        actions_row_2 = ttk.Frame(config_card, style="Card.TFrame")
        actions_row_2.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(actions_row_2, text="Nouveau code", command=self.regenerate_token, style="Quiet.TButton").pack(side="left")
        ttk.Button(actions_row_2, text="Régénérer rapports", command=self.rebuild_reports, style="Quiet.TButton").pack(side="left", padx=10)
        ttk.Button(actions_row_2, text="Créer le raccourci bureau", command=self.create_desktop_shortcut, style="Quiet.TButton").pack(side="left")

        self.short_status_value = ttk.Label(config_card, text="Enregistrez d’abord la configuration, puis démarrez le compagnon.", style="Muted.TLabel", wraplength=520)
        self.short_status_value.grid(row=10, column=0, columnspan=2, sticky="w", pady=(14, 0))

        pairing_card = self._card(body, padding=20)
        pairing_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        pairing_card.columnconfigure(0, weight=1)
        ttk.Label(pairing_card, text="Étape 2 — Démarrez puis scannez", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(pairing_card, text="Le QR code devient actif dès que le compagnon est démarré. Vous pouvez aussi copier le code manuel si besoin.", style="Muted.TLabel", wraplength=500).grid(row=1, column=0, sticky="w", pady=(6, 14))

        qr_frame = ttk.Frame(pairing_card, style="Card.TFrame")
        qr_frame.grid(row=2, column=0, sticky="nsew")
        self.qr_label = ttk.Label(qr_frame, text="QR en préparation…", style="Body.TLabel", anchor="center", justify="center")
        self.qr_label.pack(fill="both", expand=True)

        self.status_value = ttk.Label(pairing_card, text="Serveur arrêté", style="OrangePill.TLabel", padding=(12, 8))
        self.status_value.grid(row=3, column=0, sticky="w", pady=(14, 8))
        self.address_value = ttk.Label(pairing_card, text="", style="Body.TLabel", wraplength=520, justify="left")
        self.address_value.grid(row=4, column=0, sticky="w")

        ttk.Label(pairing_card, text="Code manuel", style="Body.TLabel").grid(row=5, column=0, sticky="w", pady=(14, 6))
        self.manual_code_text = tk.Text(pairing_card, height=3, wrap="word", bg=CARD_ALT, fg=TEXT, relief="flat", borderwidth=0, padx=10, pady=10, font=("Consolas", 9))
        self.manual_code_text.grid(row=6, column=0, sticky="ew")

        manual_actions = ttk.Frame(pairing_card, style="Card.TFrame")
        manual_actions.grid(row=7, column=0, sticky="w", pady=(12, 0))
        ttk.Button(manual_actions, text="Copier le code", command=self.copy_pairing_code, style="Quiet.TButton").pack(side="left")
        ttk.Button(manual_actions, text="Copier l’adresse", command=self.copy_server_address, style="Quiet.TButton").pack(side="left", padx=10)

        logs_card = self._card(body, padding=20)
        logs_card.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(18, 0))
        logs_card.columnconfigure(0, weight=1)
        logs_card.rowconfigure(1, weight=1)
        ttk.Label(logs_card, text="Journal du compagnon", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(logs_card, text="Réception des fichiers, régénération des rapports et messages utiles de configuration.", style="Muted.TLabel").grid(row=0, column=1, sticky="e")
        log_wrap = ttk.Frame(logs_card, style="Card.TFrame")
        log_wrap.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(12, 0))
        log_wrap.columnconfigure(0, weight=1)
        log_wrap.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_wrap, height=14, wrap="word", bg=CARD_ALT, fg=TEXT, relief="flat", borderwidth=0, padx=12, pady=12, font=("Segoe UI", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_wrap, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.append_log("Compagnon prêt.")

        self.footer_value = ttk.Label(outer, text=f"{COMPANY_NAME} © {COPYRIGHT_YEAR} • v{APP_VERSION}", style="Footer.TLabel")
        self.footer_value.grid(row=2, column=0, sticky="w", pady=(14, 0))

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        if self.log_text is not None:
            self.log_text.insert("end", f"[{timestamp}] {message}\n")
            self.log_text.see("end")

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.folder_path.get() or str(DEFAULT_FOLDER))
        if folder:
            self.folder_path.set(folder)

    def _pairing_payload(self) -> dict:
        return {
            "version": 3,
            "type": "grantproof_desktop_pairing",
            "transport": {
                "kind": "local_http",
                "host": local_ip(),
                "port": self.settings.port,
                "token": self.settings.pair_token,
            },
        }

    def _compact_pairing_code(self) -> str:
        transport = self._pairing_payload()["transport"]
        query = urlencode({
            "h": transport["host"],
            "p": transport["port"],
            "t": transport["token"],
        })
        return f"grantproof://pair?{query}"

    def _refresh_pairing_material(self) -> None:
        payload = self._compact_pairing_code()
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=5,
            border=3,
        )
        qr.add_data(payload)
        qr.make(fit=True)
        img = qr.make_image(fill_color=BLUE, back_color="white").convert("RGB")
        img = ImageOps.expand(img, border=12, fill="white")
        img = ImageOps.contain(img, (320, 320))
        canvas = Image.new("RGB", (340, 340), "white")
        x = (canvas.width - img.width) // 2
        y = (canvas.height - img.height) // 2
        canvas.paste(img, (x, y))
        self.qr_photo = ImageTk.PhotoImage(canvas)
        if self.qr_label is not None:
            self.qr_label.configure(image=self.qr_photo, text="")
        if self.address_value is not None:
            self.address_value.configure(text=f"Poste cible : {self.settings.workstation_name}  •  {local_ip()}:{self.settings.port}\nDossier local : {self.settings.base_folder}")
        if self.manual_code_text is not None:
            self.manual_code_text.configure(state="normal")
            self.manual_code_text.delete("1.0", "end")
            self.manual_code_text.insert("1.0", payload)
            self.manual_code_text.configure(state="disabled")

    def _refresh_status_banner(self) -> None:
        running = self.server is not None
        if self.status_value is not None:
            self.status_value.configure(
                text="Compagnon actif" if running else "Compagnon arrêté",
                style="GreenPill.TLabel" if running else "OrangePill.TLabel",
            )
        if self.workspace_value is not None:
            self.workspace_value.configure(
                text="Desktop prêt" if running else "Configuration prête à finaliser",
                style="GreenPill.TLabel" if running else "OrangePill.TLabel",
            )
        if self.short_status_value is not None:
            if running:
                self.short_status_value.configure(text="Le serveur local est démarré. Scannez maintenant le QR depuis GrantProof Mobile ou copiez le code manuel.")
            else:
                self.short_status_value.configure(text="Étape suivante : enregistrez la configuration puis cliquez sur 2. Démarrer pour activer le QR de jumelage.")

    def save_settings(self) -> None:
        try:
            port_value = int(self.port.get().strip())
            if not (1024 <= port_value <= 65535):
                raise ValueError
        except ValueError:
            messagebox.showerror(APP_NAME, "Le port doit être un nombre entre 1024 et 65535.")
            return

        folder = Path(self.folder_path.get().strip() or str(DEFAULT_FOLDER))
        folder.mkdir(parents=True, exist_ok=True)
        self.settings.workstation_name = self.workstation_name.get().strip() or "Poste GrantProof"
        self.settings.base_folder = str(folder)
        self.settings.port = port_value
        self.settings.save()
        self.report_engine = GrantProofReportEngine(Path(self.settings.base_folder))
        self.report_engine.ensure_root_files()
        self._refresh_pairing_material()
        self.append_log("Configuration enregistrée.")
        if self.short_status_value is not None:
            self.short_status_value.configure(text="Configuration enregistrée. Cliquez maintenant sur 2. Démarrer pour activer le jumelage et la synchronisation locale.")

    def copy_pairing_code(self) -> None:
        code = self._compact_pairing_code()
        self.root.clipboard_clear()
        self.root.clipboard_append(code)
        self.root.update()
        self.append_log("Code de jumelage copié dans le presse-papiers.")

    def copy_server_address(self) -> None:
        address = f"{local_ip()}:{self.settings.port}"
        self.root.clipboard_clear()
        self.root.clipboard_append(address)
        self.root.update()
        self.append_log("Adresse du serveur copiée dans le presse-papiers.")

    def regenerate_token(self) -> None:
        self.settings.pair_token = uuid.uuid4().hex
        self.settings.save()
        self._refresh_pairing_material()
        self.append_log("Nouveau code de jumelage généré.")

    def create_desktop_shortcut(self) -> None:
        if os.name != "nt":
            messagebox.showinfo(APP_NAME, "La création automatique de raccourci bureau est disponible sur Windows.")
            return
        desktop = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
        shortcut_path = desktop / "GrantProof Desktop Sync.lnk"
        icon_path = resource_path("desktop_sync_app/assets/grantproof_icon.ico")
        if getattr(sys, "frozen", False):
            target_path = Path(sys.executable)
            arguments = ""
            icon_location = str(target_path)
        else:
            target_path = Path(sys.executable)
            script_path = Path(__file__).resolve()
            arguments = f'"{script_path}"'
            icon_location = str(icon_path)
        escaped_shortcut_path = str(shortcut_path).replace("\\", "\\\\")
        escaped_target_path = str(target_path).replace("\\", "\\\\")
        escaped_arguments = arguments.replace("\\", "\\\\")
        escaped_working_directory = str(Path(self.settings.base_folder)).replace("\\", "\\\\")
        escaped_icon_location = icon_location.replace("\\", "\\\\")
        ps_script = f"""
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut('{escaped_shortcut_path}')
$Shortcut.TargetPath = '{escaped_target_path}'
$Shortcut.Arguments = '{escaped_arguments}'
$Shortcut.WorkingDirectory = '{escaped_working_directory}'
$Shortcut.IconLocation = '{escaped_icon_location},0'
$Shortcut.Description = 'GrantProof Desktop Sync'
$Shortcut.Save()
"""
        try:
            subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script], check=True, capture_output=True)
            self.append_log(f"Raccourci bureau créé : {shortcut_path}")
            messagebox.showinfo(APP_NAME, f"Raccourci bureau créé :\n{shortcut_path}")
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"Impossible de créer le raccourci : {exc}")
            messagebox.showerror(APP_NAME, f"Impossible de créer le raccourci bureau : {exc}")

    def start_server(self) -> None:
        self.save_settings()
        if self.server is not None:
            messagebox.showinfo(APP_NAME, "Le serveur est déjà démarré.")
            return
        try:
            self.server = ThreadingHTTPServer(("0.0.0.0", self.settings.port), SyncRequestHandler)
            self.server.app = self  # type: ignore[attr-defined]
            self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
            self.server_thread.start()
            self.running.set(True)
            self._refresh_status_banner()
            self.append_log(f"Compagnon démarré sur {local_ip()}:{self.settings.port}.")
        except OSError as exc:
            self.server = None
            self._refresh_status_banner()
            messagebox.showerror(APP_NAME, f"Impossible de démarrer le serveur : {exc}")
            self.append_log(f"Échec démarrage serveur : {exc}")

    def stop_server(self) -> None:
        if self.server is not None:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
            self.running.set(False)
            self._refresh_status_banner()
            self.append_log("Compagnon arrêté.")

    def on_close(self) -> None:
        self.stop_server()
        self.root.destroy()

    def handle_saved_file(self, target: Path) -> None:
        try:
            self.report_engine.ensure_root_files()
            project_code = self._project_code_from_target(target)
            if project_code:
                self.report_engine.rebuild_project(project_code)
                self.append_log(f"Rapports mis à jour pour {project_code}.")
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"Erreur génération rapports : {exc}")

    def rebuild_reports(self, project_code: str | None = None, report_language: str | None = None) -> None:
        try:
            self.report_engine.ensure_root_files()
            language_note = f" (langue: {report_language.upper()})" if isinstance(report_language, str) and report_language.strip() else ""
            if project_code:
                self.report_engine.rebuild_project(project_code, preferred_language=report_language)
                self.append_log(f"Rapports régénérés pour {project_code}{language_note}.")
            else:
                self.report_engine.rebuild_all(preferred_language=report_language)
                self.append_log(f"Rapports régénérés pour tous les projets{language_note}.")
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"Erreur régénération rapports : {exc}")
            raise

    def _project_code_from_target(self, target: Path) -> str | None:
        try:
            relative = target.relative_to(Path(self.settings.base_folder))
        except ValueError:
            return None
        parts = relative.parts
        if len(parts) >= 3 and parts[0] == "projects":
            return parts[1]
        return None

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    GrantProofDesktopSyncApp().run()
