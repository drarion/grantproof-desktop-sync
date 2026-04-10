import io
import json
import os
import socket
import threading
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tkinter import Tk, StringVar, BooleanVar, filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import qrcode
from PIL import Image, ImageTk

from reporting import GrantProofReportEngine

APP_NAME = "GrantProof Desktop Sync"
DEFAULT_PORT = 8765
SETTINGS_DIR = Path.home() / ".grantproof_desktop_sync"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"
DEFAULT_FOLDER = Path.home() / "Documents" / "GrantProof"


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
    server_version = "GrantProofDesktopSync/0.1"

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
            self._send_json({
                "name": app.settings.workstation_name,
                "device_id": app.settings.device_id,
                "port": app.settings.port,
                "ip": local_ip(),
                "status": "ready",
            })
            return
        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self):
        app = self.server.app  # type: ignore[attr-defined]
        if self.path == "/pair":
            token = self.headers.get("X-GrantProof-Token", "")
            if token != app.settings.pair_token:
                self._send_json({"error": "Invalid token"}, status=401)
                return
            self._send_json({
                "ok": True,
                "name": app.settings.workstation_name,
                "device_id": app.settings.device_id,
                "base_folder": app.settings.base_folder,
            })
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
                import base64
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
                app.rebuild_reports(project_code if isinstance(project_code, str) and project_code.strip() else None)
                self._send_json({"ok": True, "project_code": project_code or "all"})
                return
            except Exception as exc:  # noqa: BLE001
                self._send_json({"error": str(exc)}, status=400)
                return

        self._send_json({"error": "Not found"}, status=404)

    def log_message(self, format: str, *args):
        # Silence default HTTP console noise.
        pass


class GrantProofDesktopSyncApp:
    def __init__(self) -> None:
        self.settings = AppSettings.load()
        self.report_engine = GrantProofReportEngine(Path(self.settings.base_folder))
        self.root = Tk()
        self.root.title(APP_NAME)
        self.root.geometry("980x720")
        self.root.minsize(900, 650)
        self.root.configure(bg="#f5f7fb")

        self.workstation_name = StringVar(value=self.settings.workstation_name)
        self.folder_path = StringVar(value=self.settings.base_folder)
        self.port = StringVar(value=str(self.settings.port))
        self.running = BooleanVar(value=False)
        self.pair_code = StringVar(value="")
        self.server: ThreadingHTTPServer | None = None
        self.server_thread: threading.Thread | None = None
        self.qr_photo: ImageTk.PhotoImage | None = None

        self.log_text = None
        self.qr_label = None
        self.address_value = None
        self.status_value = None

        self._build_ui()
        self.report_engine.ensure_root_files()
        self._refresh_pairing_material()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#f5f7fb", foreground="#173b72", font=("Segoe UI", 22, "bold"))
        style.configure("Sub.TLabel", background="#f5f7fb", foreground="#5b6474", font=("Segoe UI", 11))
        style.configure("CardTitle.TLabel", background="#ffffff", foreground="#173b72", font=("Segoe UI", 14, "bold"))
        style.configure("Body.TLabel", background="#ffffff", foreground="#2d3748", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="GrantProof Desktop Sync", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Compagnon local-first. Reçoit les synchronisations mobile et les range dans votre dossier GrantProof.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        config_card = ttk.Frame(body, style="Card.TFrame", padding=18)
        config_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 10))
        ttk.Label(config_card, text="Configuration", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Label(config_card, text="Nom du poste", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(config_card, textvariable=self.workstation_name, width=42).grid(row=2, column=0, columnspan=3, sticky="ew")

        ttk.Label(config_card, text="Dossier local GrantProof", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(config_card, textvariable=self.folder_path, width=42).grid(row=4, column=0, columnspan=2, sticky="ew")
        ttk.Button(config_card, text="Choisir…", command=self.choose_folder).grid(row=4, column=2, padx=(8, 0))

        ttk.Label(config_card, text="Port local", style="Body.TLabel").grid(row=5, column=0, sticky="w", pady=(14, 6))
        ttk.Entry(config_card, textvariable=self.port, width=10).grid(row=6, column=0, sticky="w")

        actions = ttk.Frame(config_card, style="Card.TFrame")
        actions.grid(row=7, column=0, columnspan=3, sticky="w", pady=(18, 0))
        ttk.Button(actions, text="Enregistrer", command=self.save_settings).pack(side="left")
        ttk.Button(actions, text="Démarrer", command=self.start_server).pack(side="left", padx=8)
        ttk.Button(actions, text="Arrêter", command=self.stop_server).pack(side="left")
        ttk.Button(actions, text="Nouveau code", command=self.regenerate_token).pack(side="left", padx=8)

        pairing_card = ttk.Frame(body, style="Card.TFrame", padding=18)
        pairing_card.grid(row=0, column=1, sticky="nsew", pady=(0, 10))
        ttk.Label(pairing_card, text="Jumelage", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(pairing_card, text="Scannez ce QR code depuis GrantProof Mobile.", style="Body.TLabel").pack(anchor="w", pady=(10, 12))
        self.qr_label = ttk.Label(pairing_card, style="Body.TLabel")
        self.qr_label.pack(anchor="center", pady=6)
        self.address_value = ttk.Label(pairing_card, text="", style="Body.TLabel", wraplength=360)
        self.address_value.pack(anchor="center", pady=(8, 4))
        self.status_value = ttk.Label(pairing_card, text="Serveur arrêté", style="Body.TLabel")
        self.status_value.pack(anchor="center")

        logs_card = ttk.Frame(body, style="Card.TFrame", padding=18)
        logs_card.grid(row=1, column=0, columnspan=2, sticky="nsew")
        logs_card.rowconfigure(1, weight=1)
        logs_card.columnconfigure(0, weight=1)
        ttk.Label(logs_card, text="Journal", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.log_text = tk.Text(logs_card, height=14, wrap="word", bg="#fbfcff", fg="#1f2937", relief="flat")
        self.log_text.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.append_log("Application prête.")

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        if self.log_text is not None:
            self.log_text.insert("end", f"[{timestamp}] {message}\n")
            self.log_text.see("end")

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(initialdir=self.folder_path.get() or str(DEFAULT_FOLDER))
        if folder:
            self.folder_path.set(folder)

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

    def _pairing_payload(self) -> dict:
        return {
            "version": 1,
            "type": "grantproof_desktop_pairing",
            "name": self.settings.workstation_name,
            "device_id": self.settings.device_id,
            "transport": {
                "kind": "local_http",
                "host": local_ip(),
                "port": self.settings.port,
                "pair_path": "/pair",
                "upload_path": "/upload",
                "token": self.settings.pair_token,
            },
        }

    def _refresh_pairing_material(self) -> None:
        payload = json.dumps(self._pairing_payload(), separators=(",", ":"))
        qr = qrcode.QRCode(box_size=6, border=2)
        qr.add_data(payload)
        qr.make(fit=True)
        img = qr.make_image(fill_color="#173b72", back_color="white").convert("RGB")
        self.qr_photo = ImageTk.PhotoImage(img)
        if self.qr_label is not None:
            self.qr_label.configure(image=self.qr_photo)
        if self.address_value is not None:
            self.address_value.configure(text=f"{self.settings.workstation_name} • {local_ip()}:{self.settings.port}")

    def regenerate_token(self) -> None:
        self.settings.pair_token = uuid.uuid4().hex
        self.settings.save()
        self._refresh_pairing_material()
        self.append_log("Nouveau code de jumelage généré.")

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
            if self.status_value is not None:
                self.status_value.configure(text="Serveur local actif")
            self.append_log(f"Serveur démarré sur {local_ip()}:{self.settings.port}.")
        except OSError as exc:
            self.server = None
            messagebox.showerror(APP_NAME, f"Impossible de démarrer le serveur : {exc}")
            self.append_log(f"Échec démarrage serveur : {exc}")

    def stop_server(self) -> None:
        if self.server is not None:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
            self.running.set(False)
            if self.status_value is not None:
                self.status_value.configure(text="Serveur arrêté")
            self.append_log("Serveur arrêté.")

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

    def rebuild_reports(self, project_code: str | None = None) -> None:
        try:
            self.report_engine.ensure_root_files()
            if project_code:
                self.report_engine.rebuild_project(project_code)
                self.append_log(f"Rapports régénérés pour {project_code}.")
            else:
                self.report_engine.rebuild_all()
                self.append_log("Rapports régénérés pour tous les projets.")
        except Exception as exc:  # noqa: BLE001
            self.append_log(f"Erreur régénération rapports : {exc}")
            raise

    def _project_code_from_target(self, target: Path) -> str | None:
        try:
            relative = target.relative_to(Path(self.settings.base_folder))
        except ValueError:
            return None
        parts = relative.parts
        if len(parts) >= 3 and parts[0] == 'projects':
            return parts[1]
        return None

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    GrantProofDesktopSyncApp().run()
