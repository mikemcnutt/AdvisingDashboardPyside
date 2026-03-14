import json
import platform
import sys
import html
import uuid
import threading
import datetime as dt
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Tuple, Dict, DefaultDict
from collections import defaultdict
from urllib.parse import urlparse, parse_qs, urlencode

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import win32com.client
except Exception:
    win32com = None


APP_TITLE = "Advising Dashboard"
HEADER_TEXT = "One Dashboard To Rule Them All"

ROYAL_BG = "#0b1f5e"
CARD_BG = "#e0e7ff"
BORDER_BLUE = "#93c5fd"
ROYAL_BLUE_DARK = "#1e40af"
ROYAL_BLUE_LIGHT = "#3b82f6"

TEXT_DARK = "#0f172a"
TEXT_MUTED = "#334155"

TRACK_LABELS = {
    "BS": "Business Software and Support",
    "CL": "Cloud Computing Technologies",
    "GT": "General Track",
    "IS": "Information Security",
    "IT": "Internet Technologies",
    "NA": "Network Administration",
    "NT": "Network Technologies",
    "PR": "Programming",
}


@dataclass
class StudentInfo:
    first_name: str
    last_name: str
    student_id: str
    kctcs_email: str
    personal_email: str
    notes: str
    track_code: str
    subtrack_code: str
    json_path: str

    @property
    def display_name(self) -> str:
        name = f"{self.first_name} {self.last_name}".strip()
        return name if name else (self.kctcs_email or "(Unnamed Student)")

    @property
    def track_label(self) -> str:
        code = (self.track_code or "").strip()
        return TRACK_LABELS.get(code, code or "Unknown Track")


def safe_str(v) -> str:
    return "" if v is None else str(v)


def app_base_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent


def settings_path() -> Path:
    return app_base_dir() / "settings.json"


def load_settings() -> dict:
    p = settings_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(d: dict):
    p = settings_path()
    try:
        p.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception:
        pass


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def iter_json_files(root: Path):
    for p in root.rglob("*.json"):
        yield p


def extract_student_info(obj: dict, json_path: str) -> StudentInfo:
    student = obj.get("student") if isinstance(obj.get("student"), dict) else {}
    data = obj.get("data") if isinstance(obj.get("data"), dict) else {}
    sel = obj.get("selection") if isinstance(obj.get("selection"), dict) else {}

    return StudentInfo(
        first_name=safe_str(student.get("firstName")).strip(),
        last_name=safe_str(student.get("lastName")).strip(),
        student_id=safe_str(student.get("studentId")).strip(),
        kctcs_email=safe_str(student.get("kctcsEmail")).strip(),
        personal_email=safe_str(student.get("personalEmail")).strip(),
        notes=safe_str(data.get("notes")).strip(),
        track_code=safe_str(sel.get("scenario")).strip(),
        subtrack_code=safe_str(sel.get("subplan")).strip(),
        json_path=json_path
    )


def find_semester_plan(obj: dict, season: str, year: str) -> Optional[dict]:
    data = obj.get("data")
    if not isinstance(data, dict):
        return None
    plans = data.get("semesterPlans")
    if not isinstance(plans, list):
        return None

    for p in plans:
        if not isinstance(p, dict):
            continue
        if safe_str(p.get("season")).strip() == season and safe_str(p.get("year")).strip() == year:
            return p
    return None


def term_state(obj: dict, season: str, year: str) -> str:
    plan = find_semester_plan(obj, season, year)
    if plan is None:
        return "unadvised"

    courses = plan.get("courses", [])
    if not isinstance(courses, list) or len(courses) == 0:
        return "unadvised"

    if bool(plan.get("notComplete")):
        return "partial"

    return "done"


def classify_multi(obj: dict, terms: List[Tuple[str, str]]) -> str:
    any_unadvised = False
    any_partial = False

    for season, year in terms:
        st = term_state(obj, season, year)
        if st == "unadvised":
            any_unadvised = True
        elif st == "partial":
            any_partial = True

    if any_unadvised:
        return "needs"
    if any_partial:
        return "partial"
    return "done"


def term_badges(obj: dict, terms: List[Tuple[str, str]]) -> str:
    parts = []
    for season, year in terms:
        st = term_state(obj, season, year)
        if st == "unadvised":
            sym = "⛔"
        elif st == "partial":
            sym = "⚠️"
        else:
            sym = "✅"
        parts.append(f"{season}: {sym}")
    return "  ".join(parts)


class Tooltip:
    def __init__(self, parent: tk.Widget):
        self.parent = parent
        self.tip = None

    def show(self, x: int, y: int, text: str):
        self.hide()
        if not text:
            return
        self.tip = tk.Toplevel(self.parent)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")

        lbl = tk.Label(
            self.tip,
            text=text,
            justify="left",
            background="#0b1220",
            foreground="#e5e7eb",
            relief="solid",
            borderwidth=1,
            wraplength=520,
            padx=10,
            pady=8,
            font=("Segoe UI", 9),
        )
        lbl.pack()

    def hide(self):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
        self.tip = None


class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, background=CARD_BG)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel, add=True)

    def _on_inner_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.inner_id, width=event.width)

    def _on_mousewheel(self, event):
        if self.winfo_containing(event.x_root, event.y_root) in (self.canvas, self.inner):
            delta = -1 * int(event.delta / 120)
            self.canvas.yview_scroll(delta, "units")

    def clear(self):
        for child in self.inner.winfo_children():
            child.destroy()


def ensure_outlook_ready():
    if platform.system().lower() != "windows":
        raise RuntimeError("Outlook desktop automation is only supported on Windows.")
    if win32com is None:
        raise RuntimeError("pywin32 is not installed. Install on Windows with: pip install pywin32")


def _nl2br(text: str) -> str:
    return html.escape(text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")


def build_email_subject(base_subject: str, term_label: str) -> str:
    s = (base_subject or "").strip()
    if not s:
        s = "Advising Appointment Needed"
    if term_label and term_label.lower() in s.lower():
        return s
    return f"{s} — {term_label}" if term_label else s


def build_email_html(first_name: str, message_text: str, scheduling_link: str) -> str:
    first = (first_name or "").strip() or "there"
    msg_html = _nl2br(message_text)

    link = (scheduling_link or "").strip()
    button_block = ""
    if link:
        safe_link = html.escape(link, quote=True)
        button_block = f"""
          <div style="margin-top:18px;">
            <a href="{safe_link}"
               style="display:inline-block;background:#3b82f6;color:#ffffff;text-decoration:none;
                      padding:10px 14px;border-radius:999px;font-weight:700;font-size:14px;">
              Schedule Appointment
            </a>
          </div>
        """

    return f"""
<!doctype html>
<html>
  <body style="margin:0;padding:0;background:#f1f5f9;font-family:Segoe UI, Arial, sans-serif;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
      <tr>
        <td align="center" style="padding:18px;">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="640"
                 style="max-width:640px;background:#ffffff;border:1px solid #dbeafe;border-radius:14px;overflow:hidden;">
            <tr>
              <td style="background:#1e3a8a;padding:18px 20px;">
                <div style="color:#ffffff;font-size:16px;font-weight:700;letter-spacing:.2px;">
                  Advising Appointment
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:20px;">
                <div style="color:#0f172a;font-size:15px;font-weight:700;margin-bottom:12px;">
                  Hello {html.escape(first)},
                </div>

                <div style="color:#334155;font-size:14px;line-height:1.55;">
                  {msg_html}
                </div>

                {button_block}

                <div style="margin-top:18px;color:#0f172a;font-size:14px;">
                  Thanks,<br>
                  <span style="color:#334155;">(Your Advisor)</span>
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:14px 20px;background:#f8fafc;border-top:1px solid #e2e8f0;">
                <div style="color:#64748b;font-size:12px;line-height:1.4;">
                  This email was generated from the advising dashboard.
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
""".strip()


def outlook_create_email_html(kctcs_email: str, personal_email: str, subject: str, html_body: str, draft: bool = True):
    ensure_outlook_ready()

    to_list = [e.strip() for e in [kctcs_email, personal_email] if e and str(e).strip()]
    if not to_list:
        raise RuntimeError("Student has no email addresses in JSON (KCTCS or personal).")

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(to_list)
    mail.Subject = subject
    mail.HTMLBody = html_body

    if draft:
        mail.Save()
    else:
        mail.Send()


def open_outlook_web_email(kctcs_email: str, personal_email: str, subject: str, first_name: str, message_text: str, scheduling_link: str):
    to_list = [e.strip() for e in [kctcs_email, personal_email] if e and str(e).strip()]
    if not to_list:
        raise RuntimeError("Student has no email addresses in JSON (KCTCS or personal).")

    greeting_name = (first_name or "").strip() or "there"

    body_lines = [
        f"Hello {greeting_name},",
        "",
        message_text.strip()
    ]

    if scheduling_link and scheduling_link.strip():
        body_lines.extend([
            "",
            "Schedule Appointment:",
            scheduling_link.strip()
        ])

    body_lines.extend([
        "",
        "Thanks,",
        "(Your Advisor)"
    ])

    body_text = "\n".join(body_lines)
    body_text = body_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")

    params = {
        "to": ";".join(to_list),
        "subject": subject,
        "body": body_text
    }

    url = "https://outlook.office.com/mail/deeplink/compose?" + urllib.parse.urlencode(
        params,
        quote_via=urllib.parse.quote
    )

    webbrowser.open(url)


class LocalEditorServer:
    def __init__(self, base_dir: Path, html_filename: str = "advising.html"):
        self.base_dir = base_dir
        self.html_path = base_dir / html_filename
        self._httpd: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._port: Optional[int] = None
        self._token_to_json: Dict[str, Path] = {}

    @property
    def port(self) -> int:
        if self._port is None:
            raise RuntimeError("Server not started")
        return self._port

    def set_mapping(self, token: str, json_path: Path):
        self._token_to_json[token] = json_path

    def start(self):
        if self._httpd is not None:
            return

        if not self.html_path.exists():
            raise RuntimeError(f"Missing {self.html_path.name}. Put it in the same folder as the EXE.")

        server = self

        class Handler(BaseHTTPRequestHandler):
            def _send(self, code: int, body: bytes, content_type: str = "text/plain; charset=utf-8"):
                self.send_response(code)
                self.send_header("Content-Type", content_type)
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self):
                parsed = urlparse(self.path)

                if parsed.path == "/" or parsed.path.endswith("/"):
                    self.send_response(302)
                    self.send_header("Location", f"/{server.html_path.name}")
                    self.end_headers()
                    return

                if parsed.path == f"/{server.html_path.name}":
                    try:
                        data = server.html_path.read_bytes()
                        self._send(200, data, "text/html; charset=utf-8")
                    except Exception as e:
                        self._send(500, f"Error reading HTML: {e}".encode("utf-8"))
                    return

                if parsed.path == "/api/student":
                    qs = parse_qs(parsed.query)
                    token = (qs.get("token") or [""])[0]
                    if not token or token not in server._token_to_json:
                        self._send(404, b"Unknown token")
                        return

                    p = server._token_to_json[token]
                    try:
                        body = p.read_bytes()
                        self._send(200, body, "application/json; charset=utf-8")
                    except Exception as e:
                        self._send(500, f"Error reading JSON: {e}".encode("utf-8"))
                    return

                self._send(404, b"Not found")

            def do_POST(self):
                parsed = urlparse(self.path)
                if parsed.path != "/api/save":
                    self._send(404, b"Not found")
                    return

                qs = parse_qs(parsed.query)
                token = (qs.get("token") or [""])[0]
                if not token or token not in server._token_to_json:
                    self._send(404, b"Unknown token")
                    return

                length = int(self.headers.get("Content-Length", "0") or "0")
                raw = self.rfile.read(length) if length > 0 else b""

                try:
                    obj = json.loads(raw.decode("utf-8"))
                except Exception:
                    self._send(400, b"Invalid JSON")
                    return

                target = server._token_to_json[token]
                try:
                    target.parent.mkdir(parents=True, exist_ok=True)

                    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
                    backup = target.with_name(f"{target.stem}_backup_{stamp}{target.suffix}")
                    if target.exists():
                        backup.write_bytes(target.read_bytes())

                    pretty = json.dumps(obj, indent=2, ensure_ascii=False)
                    target.write_text(pretty, encoding="utf-8")

                    self._send(200, b'{"ok":true}', "application/json; charset=utf-8")
                except Exception as e:
                    self._send(500, f"Save failed: {e}".encode("utf-8"))

            def log_message(self, _format, *_args):
                return

        httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self._httpd = httpd
        self._port = httpd.server_address[1]

        t = threading.Thread(target=httpd.serve_forever, daemon=True)
        self._thread = t
        t.start()

    def stop(self):
        if self._httpd is not None:
            try:
                self._httpd.shutdown()
            except Exception:
                pass
            self._httpd = None
            self._thread = None
            self._port = None


class AdvisingDashboardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(1180, 740)
        self.configure(bg=ROYAL_BG)

        settings = load_settings()
        window_state = settings.get("window_state", "zoomed")

        if platform.system().lower() == "windows":
            if window_state == "zoomed":
                self.state("zoomed")
            else:
                geom = settings.get("window_geometry")
                if geom:
                    self.geometry(geom)
                else:
                    self.state("zoomed")
        else:
            if window_state == "zoomed":
                try:
                    self.attributes("-zoomed", True)
                except Exception:
                    geom = settings.get("window_geometry")
                    if geom:
                        self.geometry(geom)
                    else:
                        self.geometry("1400x900")
            else:
                geom = settings.get("window_geometry")
                if geom:
                    self.geometry(geom)
                else:
                    self.geometry("1400x900")

        self.tooltip = Tooltip(self)

        self.year_var = tk.StringVar(value="2026")
        self.folder_var = tk.StringVar(value=str(self.default_advising_folder()))

        self.spring_var = tk.BooleanVar(value=False)
        self.summer_var = tk.BooleanVar(value=False)
        self.fall_var = tk.BooleanVar(value=True)

        self.search_var = tk.StringVar(value="")
        self.track_filter_var = tk.StringVar(value="All Tracks")

        self.count_needs = tk.StringVar(value="Needs Advised: 0")
        self.count_partial = tk.StringVar(value="Advised Not Complete: 0")
        self.count_done = tk.StringVar(value="Advised: 0")

        self.all_needs_students: list[StudentInfo] = []
        self.all_partial_students: list[StudentInfo] = []
        self.all_done_students: list[StudentInfo] = []

        self.needs_students: list[StudentInfo] = []
        self.partial_students: list[StudentInfo] = []
        self.done_students: list[StudentInfo] = []

        self.needs_checks: dict[str, tk.BooleanVar] = {}

        s = load_settings()
        self.subject_var = tk.StringVar(value=s.get("subject", "Advising Appointment Needed"))
        self.scheduling_link_var = tk.StringVar(value=s.get("schedulingLink", ""))

        self.year_var.set(s.get("last_year", "2026"))
        self.folder_var.set(s.get("last_folder", str(self.default_advising_folder())))
        self.spring_var.set(bool(s.get("last_spring", False)))
        self.summer_var.set(bool(s.get("last_summer", False)))
        self.fall_var.set(bool(s.get("last_fall", True)))
        self.track_filter_var.set(s.get("last_track_filter", "All Tracks"))

        self._last_obj_by_path: dict = {}
        self._last_terms: List[Tuple[str, str]] = []

        self.server = LocalEditorServer(app_base_dir(), "advising.html")

        self._apply_theme()
        self._build_ui()

        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        self.track_filter_var.trace_add("write", lambda *_: (self._save_settings(), self.apply_filter()))
        self.year_var.trace_add("write", lambda *_: self._save_settings())
        self.folder_var.trace_add("write", lambda *_: self._save_settings())
        self.spring_var.trace_add("write", lambda *_: self._save_settings())
        self.summer_var.trace_add("write", lambda *_: self._save_settings())
        self.fall_var.trace_add("write", lambda *_: self._save_settings())

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def default_advising_folder(self) -> Path:
        return app_base_dir() / "Advising"

    def selected_terms(self) -> List[Tuple[str, str]]:
        year = self.year_var.get().strip()
        out: List[Tuple[str, str]] = []
        if self.spring_var.get():
            out.append(("Spring", year))
        if self.summer_var.get():
            out.append(("Summer", year))
        if self.fall_var.get():
            out.append(("Fall", year))
        return out

    def term_label(self) -> str:
        year = self.year_var.get().strip()
        terms = [t for (t, _y) in self.selected_terms()]
        if not terms:
            return year
        return f"{'/'.join(terms)} {year}"

    def _apply_theme(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Top.TLabelframe", background=CARD_BG, foreground=TEXT_DARK, bordercolor=BORDER_BLUE)
        style.configure("Top.TLabelframe.Label", background=CARD_BG, foreground=ROYAL_BLUE_DARK, font=("Segoe UI", 11, "bold"))

        style.configure("Card.TLabelframe", background=CARD_BG, foreground=TEXT_DARK, bordercolor=BORDER_BLUE)
        style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=ROYAL_BLUE_DARK, font=("Segoe UI", 11, "bold"))

        style.configure("Blue.TButton", background=ROYAL_BLUE_LIGHT, foreground="white",
                        font=("Segoe UI", 10, "bold"), padding=(12, 7), borderwidth=0, focusthickness=0)
        style.map("Blue.TButton", background=[("active", ROYAL_BLUE_DARK)])

        style.configure("Pill.TButton", background=ROYAL_BLUE_LIGHT, foreground="white",
                        font=("Segoe UI", 9, "bold"), padding=(12, 6), borderwidth=0, focusthickness=0)
        style.map("Pill.TButton", background=[("active", ROYAL_BLUE_DARK)])

        style.configure("Summary.TLabel", background=CARD_BG, foreground=TEXT_DARK, font=("Segoe UI", 10, "bold"))

    def _save_settings(self):
        existing = load_settings()
        existing.update({
            "subject": self.subject_var.get(),
            "schedulingLink": self.scheduling_link_var.get(),
            "last_year": self.year_var.get(),
            "last_folder": self.folder_var.get(),
            "last_spring": self.spring_var.get(),
            "last_summer": self.summer_var.get(),
            "last_fall": self.fall_var.get(),
            "last_track_filter": self.track_filter_var.get(),
        })
        save_settings(existing)

    def on_close(self):
        settings = load_settings()

        try:
            if platform.system().lower() == "windows":
                if self.state() == "zoomed":
                    settings["window_state"] = "zoomed"
                else:
                    settings["window_state"] = "normal"
                    settings["window_geometry"] = self.geometry()
            else:
                try:
                    if bool(self.attributes("-zoomed")):
                        settings["window_state"] = "zoomed"
                    else:
                        settings["window_state"] = "normal"
                        settings["window_geometry"] = self.geometry()
                except Exception:
                    settings["window_state"] = "normal"
                    settings["window_geometry"] = self.geometry()
        except Exception:
            pass

        settings["subject"] = self.subject_var.get()
        settings["schedulingLink"] = self.scheduling_link_var.get()
        settings["last_year"] = self.year_var.get()
        settings["last_folder"] = self.folder_var.get()
        settings["last_spring"] = self.spring_var.get()
        settings["last_summer"] = self.summer_var.get()
        settings["last_fall"] = self.fall_var.get()
        settings["last_track_filter"] = self.track_filter_var.get()

        save_settings(settings)
        self.destroy()

    def _quick_pair_summer_fall(self):
        self.spring_var.set(False)
        self.summer_var.set(True)
        self.fall_var.set(True)

    def _build_glow_title(self, parent):
        title_wrap = tk.Frame(parent, bg=ROYAL_BG, height=52)
        title_wrap.pack(fill="x", pady=(2, 8))

        dramatic_font = ("Georgia", 22, "bold italic")

        for dx, dy in [(-2, 0), (2, 0), (0, -2), (0, 2), (-1, -1), (1, 1), (-1, 1), (1, -1)]:
            glow = tk.Label(
                title_wrap,
                text=HEADER_TEXT,
                bg=ROYAL_BG,
                fg="#ffffff",
                font=dramatic_font
            )
            glow.place(relx=0.5, rely=0.5, anchor="center", x=dx, y=dy)

        main = tk.Label(
            title_wrap,
            text=HEADER_TEXT,
            bg=ROYAL_BG,
            fg="#f8fafc",
            font=dramatic_font
        )
        main.place(relx=0.5, rely=0.5, anchor="center")

    def _build_ui(self):
        self.container = tk.Frame(self, bg=ROYAL_BG, highlightthickness=0, bd=0)
        self.container.pack(fill="both", expand=True)

        header = tk.Frame(self.container, bg=ROYAL_BG, highlightthickness=0)
        header.pack(fill="x", pady=(10, 2))
        self._build_glow_title(header)

        top = ttk.Labelframe(self.container, text="Controls", padding=12, style="Top.TLabelframe")
        top.pack(fill="x", padx=12, pady=(8, 10))

        ttk.Label(top, text="Year:", foreground=TEXT_DARK, background=CARD_BG, font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Combobox(top, textvariable=self.year_var, state="readonly",
                     values=[str(y) for y in range(2026, 2041)], width=8).pack(side="left", padx=(8, 16))

        ttk.Label(top, text="Advise for:", foreground=TEXT_DARK, background=CARD_BG, font=("Segoe UI", 10, "bold")).pack(side="left")

        def blue_check(master, text, var):
            cb = tk.Checkbutton(
                master, text=text, variable=var,
                bg=CARD_BG, fg=TEXT_DARK,
                activebackground=CARD_BG, activeforeground=TEXT_DARK,
                selectcolor="#dbeafe",
                font=("Segoe UI", 10, "bold"),
                highlightthickness=0
            )
            cb.pack(side="left", padx=(10, 0))
            return cb

        blue_check(top, "Spring", self.spring_var)
        blue_check(top, "Summer", self.summer_var)
        blue_check(top, "Fall", self.fall_var)

        ttk.Button(top, text="Quick pair: Summer + Fall", style="Blue.TButton",
                   command=self._quick_pair_summer_fall).pack(side="left", padx=(16, 16))

        ttk.Label(top, text="Search:", foreground=TEXT_DARK, background=CARD_BG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(top, textvariable=self.search_var, width=22).pack(side="left", padx=(8, 16))

        ttk.Label(top, text="Track:", foreground=TEXT_DARK, background=CARD_BG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")

        self.track_filter_combo = ttk.Combobox(
            top,
            textvariable=self.track_filter_var,
            state="readonly",
            width=28,
            values=["All Tracks"]
        )
        self.track_filter_combo.pack(side="left", padx=(8, 16))

        ttk.Label(top, text="Advising folder:", foreground=TEXT_DARK, background=CARD_BG,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Entry(top, textvariable=self.folder_var, width=40).pack(side="left", padx=(8, 8))

        ttk.Button(top, text="Browse…", style="Blue.TButton", command=self.browse_folder).pack(side="left", padx=(0, 10))
        ttk.Button(top, text="Scan", style="Blue.TButton", command=self.scan).pack(side="left")

        self.status_label = ttk.Label(top, text="Ready", style="Summary.TLabel")
        self.status_label.pack(side="right")

        summary = ttk.Frame(self.container, padding=(12, 0), style="Top.TLabelframe")
        summary.pack(fill="x")
        ttk.Label(summary, textvariable=self.count_needs, style="Summary.TLabel").pack(side="left", padx=(0, 14))
        ttk.Label(summary, textvariable=self.count_partial, style="Summary.TLabel").pack(side="left", padx=(0, 14))
        ttk.Label(summary, textvariable=self.count_done, style="Summary.TLabel").pack(side="left")

        email_box = ttk.Labelframe(self.container, text="Email settings", padding=10, style="Card.TLabelframe")
        email_box.pack(fill="x", padx=12, pady=(10, 10))

        row1 = ttk.Frame(email_box)
        row1.pack(fill="x", pady=(0, 8))

        ttk.Label(row1, text="Subject:", background=CARD_BG, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        subj_entry = ttk.Entry(row1, textvariable=self.subject_var)
        subj_entry.pack(side="left", fill="x", expand=True, padx=(8, 12))

        ttk.Label(row1, text="Scheduling link:", background=CARD_BG, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        link_entry = ttk.Entry(row1, textvariable=self.scheduling_link_var, width=44)
        link_entry.pack(side="left", padx=(8, 0))

        subj_entry.bind("<FocusOut>", lambda _e: self._save_settings())
        link_entry.bind("<FocusOut>", lambda _e: self._save_settings())

        ttk.Label(email_box, text="Message:", background=CARD_BG, foreground=TEXT_DARK,
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.email_body = tk.Text(email_box, height=4, wrap="word", bd=1, relief="solid", highlightthickness=0)
        self.email_body.pack(fill="x", expand=True)
        self.email_body.insert("1.0", "Please reply to schedule an advising appointment for the selected semester(s).")

        main = ttk.Frame(self.container, padding=(12, 0, 12, 12))
        main.pack(fill="both", expand=True)

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(0, weight=1)

        self.frame_needs = ttk.Labelframe(main, text="Needs advised (0)", padding=10, style="Card.TLabelframe")
        self.frame_partial = ttk.Labelframe(main, text="Advised (not complete) (0)", padding=10, style="Card.TLabelframe")
        self.frame_done = ttk.Labelframe(main, text="Advised (0)", padding=10, style="Card.TLabelframe")

        self.frame_needs.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.frame_partial.grid(row=0, column=1, sticky="nsew", padx=8)
        self.frame_done.grid(row=0, column=2, sticky="nsew", padx=(8, 0))

        needs_controls = ttk.Frame(self.frame_needs)
        needs_controls.pack(fill="x", pady=(0, 8))

        ttk.Button(needs_controls, text="Select all", style="Blue.TButton", command=self.needs_select_all).pack(side="left")
        ttk.Button(needs_controls, text="Select none", style="Blue.TButton", command=self.needs_select_none).pack(side="left", padx=(8, 0))

        ttk.Button(needs_controls, text="Send Email", style="Blue.TButton",
                   command=lambda: self.email_selected_needs(draft=False)).pack(side="right")
        ttk.Button(needs_controls, text="Create Draft", style="Blue.TButton",
                   command=lambda: self.email_selected_needs(draft=True)).pack(side="right", padx=(0, 8))

        self.needs_list = ScrollableFrame(self.frame_needs)
        self.needs_list.pack(fill="both", expand=True)

        self.partial_list = ScrollableFrame(self.frame_partial)
        self.partial_list.pack(fill="both", expand=True)

        self.done_list = ScrollableFrame(self.frame_done)
        self.done_list.pack(fill="both", expand=True)

    def browse_folder(self):
        chosen = filedialog.askdirectory(title="Select Advising folder")
        if chosen:
            self.folder_var.set(chosen)

    def set_status(self, text: str):
        self.status_label.config(text=text)
        self.update_idletasks()

    def needs_select_all(self):
        for var in self.needs_checks.values():
            var.set(True)

    def needs_select_none(self):
        for var in self.needs_checks.values():
            var.set(False)

    def _current_subject(self) -> str:
        return build_email_subject(self.subject_var.get(), self.term_label())

    def _current_message_text(self) -> str:
        return self.email_body.get("1.0", "end").strip()

    def _current_link(self) -> str:
        return self.scheduling_link_var.get().strip()

    def _matches_search(self, s: StudentInfo, q: str) -> bool:
        selected_track = self.track_filter_var.get().strip()

        if selected_track and selected_track != "All Tracks":
            if s.track_label != selected_track:
                return False

        if not q:
            return True

        blob = " ".join([
            s.display_name,
            s.student_id,
            s.kctcs_email,
            s.personal_email,
            s.track_label
        ]).lower()

        return q in blob

    def _refresh_track_filter_options(self):
        track_names = sorted({
            s.track_label
            for s in (self.all_needs_students + self.all_partial_students + self.all_done_students)
            if s.track_label
        })

        values = ["All Tracks"] + track_names
        self.track_filter_combo["values"] = values

        if self.track_filter_var.get() not in values:
            self.track_filter_var.set("All Tracks")

    def apply_filter(self):
        q = self.search_var.get().strip().lower()

        self.needs_students = [s for s in self.all_needs_students if self._matches_search(s, q)]
        self.partial_students = [s for s in self.all_partial_students if self._matches_search(s, q)]
        self.done_students = [s for s in self.all_done_students if self._matches_search(s, q)]

        self._render_all()

    def _render_all(self):
        terms = self._last_terms
        obj_by_path = self._last_obj_by_path

        self._render_needs(obj_by_path, terms)
        self._render_partial(obj_by_path, terms)
        self._render_done(obj_by_path, terms)

        n_needs = len(self.needs_students)
        n_partial = len(self.partial_students)
        n_done = len(self.done_students)

        self.count_needs.set(f"Needs Advised: {n_needs}")
        self.count_partial.set(f"Advised Not Complete: {n_partial}")
        self.count_done.set(f"Advised: {n_done}")

        self.frame_needs.config(text=f"Needs advised ({n_needs})")
        self.frame_partial.config(text=f"Advised (not complete) ({n_partial})")
        self.frame_done.config(text=f"Advised ({n_done})")

    def email_selected_needs(self, draft: bool):
        selected = []
        for s in self.needs_students:
            var = self.needs_checks.get(s.json_path)
            if var and var.get():
                selected.append(s)

        if not selected:
            messagebox.showinfo("No selection", "Select at least one student to email.")
            return

        subject = self._current_subject()
        message_text = self._current_message_text()
        link = self._current_link()

        if platform.system().lower() != "windows":
            opened = 0
            errors = 0
            for s in selected:
                try:
                    open_outlook_web_email(
                        s.kctcs_email,
                        s.personal_email,
                        subject,
                        s.first_name,
                        message_text,
                        link
                    )
                    opened += 1
                except Exception:
                    errors += 1

            messagebox.showinfo("Email complete", f"Opened Outlook Web drafts: {opened}\nErrors: {errors}")
            return

        try:
            ensure_outlook_ready()
        except Exception as e:
            messagebox.showerror("Email unavailable", str(e))
            return

        if not draft:
            confirm = messagebox.askyesno("Confirm send", f"Send {len(selected)} email(s) now?")
            if not confirm:
                return

        ok = 0
        err = 0

        for s in selected:
            try:
                html_body = build_email_html(s.first_name, message_text, link)
                outlook_create_email_html(s.kctcs_email, s.personal_email, subject, html_body, draft=draft)
                ok += 1
            except Exception:
                err += 1

        mode = "Drafted" if draft else "Sent"
        messagebox.showinfo("Email complete", f"{mode}: {ok}\nErrors: {err}")

    def email_one_partial(self, s: StudentInfo):
        subject = self._current_subject()
        message_text = self._current_message_text()
        link = self._current_link()

        if platform.system().lower() != "windows":
            try:
                open_outlook_web_email(
                    s.kctcs_email,
                    s.personal_email,
                    subject,
                    s.first_name,
                    message_text,
                    link
                )
                messagebox.showinfo("Email opened", f"Outlook Web draft opened for {s.display_name}.")
            except Exception as e:
                messagebox.showerror("Email failed", str(e))
            return

        try:
            ensure_outlook_ready()
        except Exception as e:
            messagebox.showerror("Email unavailable", str(e))
            return

        try:
            html_body = build_email_html(s.first_name, message_text, link)
            outlook_create_email_html(s.kctcs_email, s.personal_email, subject, html_body, draft=True)
            messagebox.showinfo("Draft created", f"Draft email created for {s.display_name}.")
        except Exception as e:
            messagebox.showerror("Email failed", str(e))

    def open_in_editor(self, json_path: str):
        try:
            self.server.start()
        except Exception as e:
            messagebox.showerror("Editor unavailable", str(e))
            return

        token = uuid.uuid4().hex
        self.server.set_mapping(token, Path(json_path))

        params = urlencode({
            "token": token,
            "json": f"/api/student?token={token}",
            "save": f"/api/save?token={token}",
        })
        url = f"http://127.0.0.1:{self.server.port}/{self.server.html_path.name}?{params}"

        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Could not open browser", str(e))

    def _render_name_link(self, parent, text: str, json_path: str):
        lbl = tk.Label(
            parent,
            text=text,
            bg=CARD_BG,
            fg=ROYAL_BLUE_DARK,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2"
        )
        lbl.pack(anchor="w")
        lbl.bind("<Button-1>", lambda _e: self.open_in_editor(json_path))
        lbl.bind("<Enter>", lambda _e: lbl.config(fg=ROYAL_BLUE_LIGHT))
        lbl.bind("<Leave>", lambda _e: lbl.config(fg=ROYAL_BLUE_DARK))
        return lbl

    def _grouped_by_track(self, students: List[StudentInfo]) -> List[Tuple[str, List[StudentInfo]]]:
        buckets: DefaultDict[str, List[StudentInfo]] = defaultdict(list)
        for s in students:
            buckets[s.track_label].append(s)

        items = list(buckets.items())
        items.sort(key=lambda kv: kv[0].lower())
        for _track, lst in items:
            lst.sort(key=lambda s: s.display_name.lower())
        return items

    def _render_track_header(self, parent, track_label: str, count: int):
        hdr = tk.Frame(parent, bg="#c7d2fe", highlightthickness=0, bd=0)
        hdr.pack(fill="x", pady=(8, 4))
        txt = f"{track_label} ({count})"
        tk.Label(hdr, text=txt, bg="#c7d2fe", fg=ROYAL_BLUE_DARK,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=8, pady=6)

    def _render_needs(self, obj_by_path: dict, terms: List[Tuple[str, str]]):
        self.needs_list.clear()
        self.needs_checks.clear()

        for track_label, group_students in self._grouped_by_track(self.needs_students):
            self._render_track_header(self.needs_list.inner, track_label, len(group_students))

            for s in group_students:
                holder = tk.Frame(self.needs_list.inner, bg=CARD_BG, highlightbackground=BORDER_BLUE, highlightthickness=1)
                holder.pack(fill="x", pady=4)

                row = ttk.Frame(holder)
                row.pack(fill="x", padx=8, pady=6)

                var = tk.BooleanVar(value=True)
                self.needs_checks[s.json_path] = var

                tk.Checkbutton(
                    row,
                    variable=var,
                    bg=CARD_BG,
                    activebackground=CARD_BG,
                    highlightthickness=0
                ).pack(side="left", padx=(0, 10))

                left = ttk.Frame(row)
                left.pack(side="left", fill="x", expand=True)

                self._render_name_link(left, s.display_name, s.json_path)

                badges = ""
                obj = obj_by_path.get(s.json_path)
                if obj is not None:
                    badges = term_badges(obj, terms)

                ttk.Label(
                    left,
                    text=badges,
                    background=CARD_BG,
                    foreground=TEXT_MUTED,
                    font=("Segoe UI", 9)
                ).pack(anchor="w")

                ttk.Label(
                    row,
                    text=s.student_id,
                    background=CARD_BG,
                    foreground=TEXT_MUTED,
                    font=("Segoe UI", 9)
                ).pack(side="right")

    def _render_partial(self, obj_by_path: dict, terms: List[Tuple[str, str]]):
        self.partial_list.clear()

        for track_label, group_students in self._grouped_by_track(self.partial_students):
            self._render_track_header(self.partial_list.inner, track_label, len(group_students))

            for s in group_students:
                holder = tk.Frame(self.partial_list.inner, bg=CARD_BG, highlightbackground=BORDER_BLUE, highlightthickness=1)
                holder.pack(fill="x", pady=4)

                row = ttk.Frame(holder)
                row.pack(fill="x", padx=8, pady=8)

                left = ttk.Frame(row)
                left.pack(side="left", fill="x", expand=True)

                self._render_name_link(left, s.display_name, s.json_path)

                badges = ""
                obj = obj_by_path.get(s.json_path)
                if obj is not None:
                    badges = term_badges(obj, terms)

                ttk.Label(left, text=badges, background=CARD_BG, foreground=TEXT_MUTED,
                          font=("Segoe UI", 9)).pack(anchor="w")

                ttk.Label(left, text=s.student_id, background=CARD_BG, foreground=TEXT_MUTED,
                          font=("Segoe UI", 9)).pack(anchor="w")

                right = ttk.Frame(row)
                right.pack(side="right")

                if s.notes.strip():
                    notes_lbl = tk.Label(
                        right,
                        text="Notes",
                        bg="#dbeafe",
                        fg=ROYAL_BLUE_DARK,
                        padx=10,
                        pady=4,
                        font=("Segoe UI", 9, "bold"),
                        cursor="question_arrow"
                    )
                    notes_lbl.pack(side="left", padx=(0, 8))

                    notes_lbl.bind("<Enter>", lambda _e, n=s.notes: self.tooltip.show(self.winfo_pointerx()+12, self.winfo_pointery()+12, n))
                    notes_lbl.bind("<Leave>", lambda _e: self.tooltip.hide())

                ttk.Button(right, text="Email", style="Pill.TButton",
                           command=lambda stu=s: self.email_one_partial(stu)).pack(side="left")

    def _render_done(self, obj_by_path: dict, terms: List[Tuple[str, str]]):
        self.done_list.clear()

        for track_label, group_students in self._grouped_by_track(self.done_students):
            self._render_track_header(self.done_list.inner, track_label, len(group_students))

            for s in group_students:
                holder = tk.Frame(self.done_list.inner, bg=CARD_BG, highlightbackground=BORDER_BLUE, highlightthickness=1)
                holder.pack(fill="x", pady=4)

                row = ttk.Frame(holder)
                row.pack(fill="x", padx=8, pady=8)

                left = ttk.Frame(row)
                left.pack(side="left", fill="x", expand=True)

                self._render_name_link(left, s.display_name, s.json_path)

                badges = ""
                obj = obj_by_path.get(s.json_path)
                if obj is not None:
                    badges = term_badges(obj, terms)

                ttk.Label(left, text=badges, background=CARD_BG, foreground=TEXT_MUTED,
                          font=("Segoe UI", 9)).pack(anchor="w")

                ttk.Label(row, text=s.student_id, background=CARD_BG, foreground=TEXT_MUTED,
                          font=("Segoe UI", 9)).pack(side="right")

    def scan(self):
        terms = self.selected_terms()
        if not terms:
            messagebox.showerror("Select term(s)", "Pick at least one term (Spring/Summer/Fall) to scan.")
            return

        folder = Path(self.folder_var.get()).expanduser()
        if not folder.exists() or not folder.is_dir():
            messagebox.showerror("Folder not found", f"Advising folder does not exist:\n{folder}")
            return

        label = self.term_label()
        self.set_status(f"Scanning for {label}…")

        needs: list[StudentInfo] = []
        partial: list[StudentInfo] = []
        done: list[StudentInfo] = []

        files = list(iter_json_files(folder))
        bad_files = 0

        obj_by_path: dict = {}

        for p in files:
            try:
                obj = load_json(p)
                obj_by_path[str(p)] = obj

                bucket = classify_multi(obj, terms)
                info = extract_student_info(obj, str(p))

                if bucket == "needs":
                    needs.append(info)
                elif bucket == "partial":
                    partial.append(info)
                else:
                    done.append(info)
            except Exception:
                bad_files += 1
                continue

        self.all_needs_students = needs
        self.all_partial_students = partial
        self.all_done_students = done

        self._refresh_track_filter_options()

        self._last_obj_by_path = obj_by_path
        self._last_terms = terms

        self.apply_filter()

        msg = f"{len(files)} file(s) scanned"
        if bad_files:
            msg += f" • {bad_files} unreadable"
        self.set_status(msg)


def main():
    app = AdvisingDashboardApp()
    app.mainloop()


if __name__ == "__main__":
    main()
