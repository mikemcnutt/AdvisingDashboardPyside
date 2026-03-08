"""
Advising Dashboard - Glassmorphism Design
Beautiful frosted glass UI with glowing borders and modern aesthetics
"""

import json
import platform
import sys
import html
import uuid
import threading
import webbrowser
from pathlib import Path
from typing import Optional, List, Tuple, Dict
from collections import OrderedDict
from dataclasses import dataclass
from urllib.parse import urlparse, parse_qs
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox, QScrollArea,
    QFrame, QSizePolicy, QMessageBox, QFileDialog, QGridLayout, QGraphicsDropShadowEffect
)
from PySide6.QtCore import Qt, QTimer, Signal, QSize, QRect, QPropertyAnimation, QEasingCurve, Property, QByteArray
from PySide6.QtGui import (
    QFont, QPalette, QColor, QPainter, QPainterPath, QLinearGradient, 
    QBrush, QPen, QPixmap, QIcon
)

try:
    import win32com.client
except Exception:
    win32com = None


APP_TITLE = "Advising Dashboard"
HEADER_TEXT = "One Dashboard To Rule Them All"

# Glassmorphism color palette - cohesive blue frosted glass theme
COLORS = {
    # Backgrounds - unified blue gradient
    'bg_main': '#1a3d5c',           # Deep blue
    'bg_gradient_start': '#2b5278', # Medium blue
    'bg_gradient_end': '#4a7ba7',   # Lighter blue
    
    # Glass effects - softer, more transparent
    'glass_bg': 'rgba(255, 255, 255, 0.12)',  # More visible translucent
    'glass_border': 'rgba(255, 255, 255, 0.3)', # Soft white border
    'glass_glow': 'rgba(255, 255, 255, 0.5)',   # Bright white glow
    'glass_shadow': 'rgba(0, 0, 0, 0.2)',       # Subtle shadow
    
    # Accents - lighter, softer tones
    'accent_light': '#a8d8ea',      # Light cyan
    'accent_bright': '#e3f2fd',     # Very light blue
    'accent_white': '#ffffff',      # Pure white
    
    # Text - white and light blue
    'text_primary': '#ffffff',
    'text_secondary': '#e3f2fd',
    'text_muted': '#b3d9f2',
    
    # Status colors - softer tones
    'status_needed': 'rgba(255, 150, 150, 0.8)',   # Soft red
    'status_partial': 'rgba(255, 200, 100, 0.8)',  # Soft gold
    'status_complete': 'rgba(150, 255, 180, 0.8)', # Soft green
}

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


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.resolve()


def load_settings() -> dict:
    settings_file = app_base_dir() / "advising_dashboard_settings.json"
    if settings_file.exists():
        try:
            return json.loads(settings_file.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_settings(settings: dict):
    settings_file = app_base_dir() / "advising_dashboard_settings.json"
    try:
        settings_file.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    except Exception:
        pass


@dataclass
class SnapshotInfo:
    file_path: Path
    student_name: str
    student_id: str
    kctcs_email: str
    personal_email: str
    track: str
    track_name: str
    badges: List[str]
    spring_done: bool
    summer_done: bool
    fall_done: bool
    spring_partial: bool
    summer_partial: bool
    fall_partial: bool
    notes: str


class LocalEditorServer:
    def __init__(self, html_file: str):
        self.html_path = app_base_dir() / html_file
        self.token_map: Dict[str, Path] = {}
        self.server = None
        self.thread = None

    def start(self):
        if self.server is not None:
            return

        class Handler(BaseHTTPRequestHandler):
            def log_message(self2, format, *args):
                pass

            def do_GET(self2):
                parsed = urlparse(self2.path)
                path = parsed.path
                qs = parse_qs(parsed.query)

                if path == "/advising.html" or path == "/Advising.html":
                    if not self.html_path.exists():
                        self2.send_response(404)
                        self2.end_headers()
                        self2.wfile.write(b"HTML file not found")
                        return
                    self2.send_response(200)
                    self2.send_header("Content-Type", "text/html; charset=utf-8")
                    self2.end_headers()
                    self2.wfile.write(self.html_path.read_bytes())

                elif path == "/api/student":
                    token = qs.get("token", [""])[0]
                    json_file = self.token_map.get(token)
                    if json_file is None or not json_file.exists():
                        self2.send_response(404)
                        self2.end_headers()
                        self2.wfile.write(b"Not found")
                        return
                    self2.send_response(200)
                    self2.send_header("Content-Type", "application/json")
                    self2.end_headers()
                    self2.wfile.write(json_file.read_bytes())

                else:
                    self2.send_response(404)
                    self2.end_headers()

            def do_POST(self2):
                if self2.path.startswith("/api/save"):
                    parsed = urlparse(self2.path)
                    qs = parse_qs(parsed.query)
                    token = qs.get("token", [""])[0]
                    json_file = self.token_map.get(token)
                    if json_file is None:
                        self2.send_response(404)
                        self2.end_headers()
                        return

                    length = int(self2.headers.get("Content-Length", 0))
                    body = self2.rfile.read(length)

                    backup = json_file.with_suffix(".bak")
                    try:
                        if json_file.exists():
                            backup.write_bytes(json_file.read_bytes())
                    except Exception:
                        pass

                    json_file.write_bytes(body)

                    self2.send_response(200)
                    self2.send_header("Content-Type", "application/json")
                    self2.end_headers()
                    self2.wfile.write(b'{"status":"ok"}')
                else:
                    self2.send_response(404)
                    self2.end_headers()

        try:
            self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
            self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
            self.thread.start()
        except Exception as e:
            QMessageBox.critical(None, "Server Error", f"Could not start local server:\n{e}")

    def get_url(self, json_file: Path) -> str:
        if self.server is None:
            return ""
        tok = str(uuid.uuid4())
        self.token_map[tok] = json_file
        port = self.server.server_address[1]
        base = f"http://127.0.0.1:{port}"
        import urllib.parse
        json_url = f"{base}/api/student?token={tok}"
        save_url = f"{base}/api/save?token={tok}"
        return f"{base}/advising.html?token={tok}&json={urllib.parse.quote(json_url)}&save={urllib.parse.quote(save_url)}"


def build_email_subject(template: str, term_label: str) -> str:
    if "{term}" in template or "{TERM}" in template:
        return template.replace("{term}", term_label).replace("{TERM}", term_label)
    return f"{template} - {term_label}"


def send_outlook_emails(subject: str, body_html: str, recipients: List[str], draft: bool = False):
    if win32com is None:
        QMessageBox.critical(None, "Error", "Outlook automation not available (win32com not installed)")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        for addr in recipients:
            mail = outlook.CreateItem(0)
            mail.To = addr
            mail.Subject = subject
            mail.HTMLBody = body_html
            if draft:
                mail.Save()
            else:
                mail.Send()
    except Exception as e:
        QMessageBox.critical(None, "Outlook Error", f"Could not send emails:\n{e}")


class GlassCard(QFrame):
    """Glassmorphism card with translucent background and soft border"""
    def __init__(self, parent=None, glow_color=None, border_width=2):
        super().__init__(parent)
        
        self.glow_color = glow_color or COLORS['glass_border']
        self.border_width = border_width
        
        # Softer frosted glass effect
        self.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(255, 255, 255, 0.15);
                border: {border_width}px solid rgba(255, 255, 255, 0.3);
                border-radius: 24px;
            }}
        """)
        
        # Subtle shadow instead of bright glow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 60))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)


class GlassButton(QPushButton):
    """Glassmorphism button with soft glow"""
    def __init__(self, text, glow_color=None, parent=None):
        super().__init__(text, parent)
        
        self.glow_color = glow_color or COLORS['accent_light']
        
        self.setMinimumHeight(44)
        self.setMinimumWidth(130)
        self.setCursor(Qt.PointingHandCursor)
        self.setFont(QFont("Segoe UI", 11, QFont.Bold))
        
        self._update_style(False)
        
        # Soft shadow
        self.shadow = QGraphicsDropShadowEffect()
        self.shadow.setBlurRadius(15)
        self.shadow.setColor(QColor(0, 0, 0, 40))
        self.shadow.setOffset(0, 2)
        self.setGraphicsEffect(self.shadow)
    
    def _update_style(self, hovering):
        if hovering:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: rgba(255, 255, 255, 0.25);
                    color: {COLORS['text_primary']};
                    border: 2px solid rgba(255, 255, 255, 0.5);
                    border-radius: 22px;
                    padding: 10px 28px;
                    font-weight: bold;
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: rgba(255, 255, 255, 0.18);
                    color: {COLORS['text_primary']};
                    border: 2px solid rgba(255, 255, 255, 0.35);
                    border-radius: 22px;
                    padding: 10px 28px;
                    font-weight: bold;
                }}
            """)
    
    def enterEvent(self, event):
        self._update_style(True)
        super().enterEvent(event)
    
    def leaveEvent(self, event):
        self._update_style(False)
        super().leaveEvent(event)


class StudentCard(QFrame):
    """Student card with soft glass effect"""
    clicked = Signal(object)
    
    def __init__(self, student: SnapshotInfo, accent_color, parent=None,
                 show_checkbox=False, show_email_btn=False):
        super().__init__(parent)
        
        self.student = student
        self.accent_color = accent_color
        self.show_checkbox = show_checkbox
        self.show_email_btn = show_email_btn
        self.checkbox = None
        
        self.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(255, 255, 255, 0.12);
                border: 2px solid rgba(255, 255, 255, 0.2);
                border-radius: 14px;
                padding: 12px;
            }}
            QFrame:hover {{
                background-color: rgba(255, 255, 255, 0.18);
                border: 2px solid rgba(255, 255, 255, 0.35);
            }}
        """)
        
        self.setMinimumHeight(80)
        self.setCursor(Qt.PointingHandCursor)
        
        # Subtle shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(12)
        shadow.setColor(QColor(0, 0, 0, 40))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)
        
        # Top row
        top_row = QHBoxLayout()
        
        if self.show_checkbox:
            self.checkbox = QCheckBox()
            self.checkbox.setStyleSheet(f"""
                QCheckBox::indicator {{
                    width: 22px;
                    height: 22px;
                    border-radius: 7px;
                    border: 2px solid rgba(255, 255, 255, 0.4);
                    background-color: rgba(255, 255, 255, 0.1);
                }}
                QCheckBox::indicator:checked {{
                    background-color: rgba(255, 255, 255, 0.35);
                    border-color: rgba(255, 255, 255, 0.6);
                }}
                QCheckBox::indicator:hover {{
                    border-color: rgba(255, 255, 255, 0.6);
                }}
            """)
            top_row.addWidget(self.checkbox)
        
        # Name
        name_label = QLabel(self.student.student_name)
        name_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        name_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                background: transparent;
            }}
        """)
        top_row.addWidget(name_label)
        top_row.addStretch()
        
        # Notes badge
        if self.student.notes:
            notes_label = QLabel("Notes")
            notes_label.setToolTip(self.student.notes)
            notes_label.setFont(QFont("Segoe UI", 9, QFont.Bold))
            notes_label.setStyleSheet(f"""
                QLabel {{
                    background-color: rgba(255, 200, 100, 0.3);
                    color: {COLORS['text_primary']};
                    padding: 6px 10px;
                    border-radius: 8px;
                    border: 1px solid rgba(255, 200, 100, 0.5);
                }}
            """)
            top_row.addWidget(notes_label)
        
        # Email button
        if self.show_email_btn:
            email_btn = GlassButton("Email", glow_color=COLORS['status_needed'])
            email_btn.setMaximumWidth(90)
            email_btn.setMaximumHeight(34)
            top_row.addWidget(email_btn)
        
        # Student ID
        if self.student.student_id:
            id_label = QLabel(self.student.student_id)
            id_label.setFont(QFont("Consolas", 10))
            id_label.setStyleSheet(f"""
                QLabel {{
                    color: {COLORS['text_secondary']};
                    background-color: rgba(0, 0, 0, 0.2);
                    padding: 6px 12px;
                    border-radius: 8px;
                    border: 1px solid rgba(255, 255, 255, 0.15);
                }}
            """)
            top_row.addWidget(id_label)
        
        layout.addLayout(top_row)
        
        # Badges
        badges_label = QLabel("  ".join(self.student.badges))
        badges_label.setFont(QFont("Segoe UI", 9))
        badges_label.setStyleSheet(f"color: {COLORS['text_muted']}; background: transparent;")
        layout.addWidget(badges_label)
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.student)
        super().mousePressEvent(event)


class ColumnCard(QFrame):
    """Column container with soft glass header"""
    def __init__(self, title, glow_color, parent=None):
        super().__init__(parent)
        
        self.glow_color = glow_color
        self.title_label = None
        
        # Softer glassmorphism styling
        self.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(255, 255, 255, 0.12);
                border: 2px solid rgba(255, 255, 255, 0.25);
                border-radius: 28px;
            }}
        """)
        
        # Subtle shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(25)
        shadow.setColor(QColor(0, 0, 0, 50))
        shadow.setOffset(0, 5)
        self.setGraphicsEffect(shadow)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Header
        header = QFrame()
        header.setMinimumHeight(70)
        header.setMaximumHeight(70)
        header.setStyleSheet(f"""
            QFrame {{
                background: rgba(255, 255, 255, 0.08);
                border-top-left-radius: 26px;
                border-top-right-radius: 26px;
                border: none;
                border-bottom: 2px solid rgba(255, 255, 255, 0.2);
            }}
        """)
        
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(24, 0, 24, 0)
        
        # Title
        title_container = QWidget()
        title_container.setStyleSheet("background: transparent;")
        title_layout = QVBoxLayout(title_container)
        title_layout.setContentsMargins(0, 8, 0, 0)
        title_layout.setSpacing(4)
        
        self.title_label = QLabel(title)
        self.title_label.setFont(QFont("Segoe UI", 15, QFont.Bold))
        self.title_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                background: transparent;
            }}
        """)
        self.title_label.setAlignment(Qt.AlignCenter)
        
        # Count badge
        self.count_label = QLabel("0")
        self.count_label.setFont(QFont("Segoe UI", 28, QFont.Bold))
        self.count_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                background: transparent;
            }}
        """)
        self.count_label.setAlignment(Qt.AlignCenter)
        
        title_layout.addWidget(self.title_label)
        
        header_layout.addWidget(title_container)
        
        layout.addWidget(header)
        
        # Content area
        self.content = QWidget()
        self.content.setStyleSheet("background: transparent;")
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(16, 16, 16, 16)
        self.content_layout.setSpacing(12)
        
        layout.addWidget(self.content)
    
    def set_title(self, title):
        if self.title_label:
            self.title_label.setText(title)


class AdvisingDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle(APP_TITLE)
        
        # Set initial size to Full HD (1920x1080)
        self.resize(1920, 1080)
        
        # Set minimum size for proper scaling
        self.setMinimumSize(1280, 720)
        
        # Center window on screen
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - 1920) // 2
        y = (screen.height() - 1080) // 2
        self.move(max(0, x), max(0, y))
        
        # Variables
        self.snapshots: List[SnapshotInfo] = []
        self.needs_checks: Dict[str, QCheckBox] = {}
        
        # Settings
        settings = load_settings()
        self.current_year = settings.get("last_year", "2026")
        self.spring_enabled = settings.get("last_spring", False)
        self.summer_enabled = settings.get("last_summer", True)
        self.fall_enabled = settings.get("last_fall", True)
        self.folder_path = settings.get("last_folder", "")
        self.track_filter = settings.get("last_track_filter", "All Tracks")
        self.email_subject = settings.get("subject", "Advising Appointment Needed")
        self.scheduling_link = settings.get("schedulingLink", "")
        
        # Restore window geometry if available
        if "window_geometry" in settings:
            try:
                geom = settings["window_geometry"]
                if isinstance(geom, str):
                    from base64 import b64decode
                    self.restoreGeometry(QByteArray(b64decode(geom.encode())))
            except:
                pass
        
        # Restore window state (maximized or normal)
        if settings.get("window_state") == "maximized":
            self.setWindowState(Qt.WindowMaximized)
        
        # Server
        html_file = "advising.html"
        if not (app_base_dir() / html_file).exists():
            html_file = "Advising.html"
        self.server = LocalEditorServer(html_file)
        
        self._setup_ui()
        self._apply_styles()
    
    def _setup_ui(self):
        # Central widget with gradient background
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(24)
        
        # Header
        self._build_header(main_layout)
        
        # Control panel
        self._build_control_panel(main_layout)
        
        # Student columns
        self._build_columns(main_layout)
    
    def _build_header(self, parent_layout):
        header_card = GlassCard(glow_color=COLORS['accent_light'])
        header_card.setMinimumHeight(120)
        header_card.setMaximumHeight(120)
        
        layout = QVBoxLayout(header_card)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(8)
        
        # Main title with glow
        title = QLabel(HEADER_TEXT)
        title.setFont(QFont("Segoe UI", 36, QFont.Bold))
        title.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                background: transparent;
            }}
        """)
        title.setAlignment(Qt.AlignCenter)
        
        # Add text glow effect
        title_shadow = QGraphicsDropShadowEffect()
        title_shadow.setBlurRadius(30)
        title_shadow.setColor(QColor(COLORS['accent_light']))
        title_shadow.setOffset(0, 0)
        title.setGraphicsEffect(title_shadow)
        
        layout.addWidget(title)
        
        # Subtitle
        subtitle = QLabel("Advanced Student Management System")
        subtitle.setFont(QFont("Segoe UI", 12, QFont.Bold))
        subtitle.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['accent_light']};
                background: transparent;
            }}
        """)
        subtitle.setAlignment(Qt.AlignCenter)
        layout.addWidget(subtitle)
        
        parent_layout.addWidget(header_card)
    
    def _build_control_panel(self, parent_layout):
        panel = GlassCard(glow_color=COLORS['glass_border'])
        
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(20)
        
        # Top row
        top_row = QHBoxLayout()
        top_row.setSpacing(20)
        
        # Year
        year_widget = self._create_control_group("Academic Year")
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(y) for y in range(2026, 2041)])
        self.year_combo.setCurrentText(self.current_year)
        year_widget.layout().addWidget(self.year_combo)
        top_row.addWidget(year_widget)
        
        # Semesters
        sem_widget = self._create_control_group("Advising For")
        sem_checks = QHBoxLayout()
        
        self.spring_check = QCheckBox("Spring")
        self.spring_check.setChecked(self.spring_enabled)
        sem_checks.addWidget(self.spring_check)
        
        self.summer_check = QCheckBox("Summer")
        self.summer_check.setChecked(self.summer_enabled)
        sem_checks.addWidget(self.summer_check)
        
        self.fall_check = QCheckBox("Fall")
        self.fall_check.setChecked(self.fall_enabled)
        sem_checks.addWidget(self.fall_check)
        
        sem_widget.layout().addLayout(sem_checks)
        top_row.addWidget(sem_widget)
        
        # Quick button
        quick_btn = GlassButton("Quick: Summer + Fall", glow_color=COLORS['accent_light'])
        quick_btn.clicked.connect(self._quick_pair)
        quick_btn.setMaximumWidth(200)
        top_row.addWidget(quick_btn)
        
        # Search
        search_widget = self._create_control_group("Search")
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Name, ID, Email...")
        self.search_entry.textChanged.connect(self._on_search_changed)
        search_widget.layout().addWidget(self.search_entry)
        top_row.addWidget(search_widget, 1)
        
        # Track filter
        track_widget = self._create_control_group("Track Filter")
        self.track_combo = QComboBox()
        tracks = ["All Tracks"] + [f"{k}: {v}" for k, v in sorted(TRACK_LABELS.items())]
        self.track_combo.addItems(tracks)
        self.track_combo.setCurrentText(self.track_filter)
        self.track_combo.currentTextChanged.connect(self._on_filter_changed)
        track_widget.layout().addWidget(self.track_combo)
        top_row.addWidget(track_widget, 1)
        
        layout.addLayout(top_row)
        
        # Folder row
        folder_row = QHBoxLayout()
        folder_row.setSpacing(12)
        
        folder_widget = self._create_control_group("Advising Folder")
        self.folder_entry = QLineEdit(self.folder_path)
        folder_widget.layout().addWidget(self.folder_entry)
        folder_row.addWidget(folder_widget, 1)
        
        browse_btn = GlassButton("Browse...", glow_color=COLORS['accent_light'])
        browse_btn.clicked.connect(self._browse_folder)
        browse_btn.setMaximumWidth(120)
        folder_row.addWidget(browse_btn)
        
        scan_btn = GlassButton("Scan Folder", glow_color=COLORS['status_complete'])
        scan_btn.clicked.connect(self._scan_folder)
        scan_btn.setMaximumWidth(160)
        folder_row.addWidget(scan_btn)
        
        layout.addLayout(folder_row)
        
        # Status
        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Segoe UI", 10))
        self.status_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_muted']};
                background: transparent;
                font-style: italic;
            }}
        """)
        self.status_label.setAlignment(Qt.AlignRight)
        layout.addWidget(self.status_label)
        
        # Email templates
        email_row = QHBoxLayout()
        email_row.setSpacing(20)
        
        subj_widget = self._create_control_group("Email Subject")
        self.subject_entry = QLineEdit(self.email_subject)
        subj_widget.layout().addWidget(self.subject_entry)
        email_row.addWidget(subj_widget, 1)
        
        link_widget = self._create_control_group("Scheduling Link")
        self.link_entry = QLineEdit(self.scheduling_link)
        link_widget.layout().addWidget(self.link_entry)
        email_row.addWidget(link_widget, 1)
        
        layout.addLayout(email_row)
        
        parent_layout.addWidget(panel)
    
    def _create_control_group(self, label_text):
        """Create a labeled control group"""
        widget = QWidget()
        widget.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        label = QLabel(label_text)
        label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['accent_light']};
                background: transparent;
            }}
        """)
        layout.addWidget(label)
        
        return widget
    
    def _build_columns(self, parent_layout):
        columns_layout = QGridLayout()
        columns_layout.setSpacing(20)
        
        # Equal column widths with proper scaling
        columns_layout.setColumnStretch(0, 1)
        columns_layout.setColumnStretch(1, 1)
        columns_layout.setColumnStretch(2, 1)
        
        # Three glowing columns
        self.needs_column = ColumnCard("Needs Advised (0)", COLORS['status_needed'])
        self.needs_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.partial_column = ColumnCard("Advised (Not Complete) (0)", COLORS['status_partial'])
        self.partial_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.done_column = ColumnCard("Advised (0)", COLORS['status_complete'])
        self.done_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Add controls to needs column
        controls = QHBoxLayout()
        controls.setSpacing(12)
        
        select_all_btn = GlassButton("Select All", glow_color=COLORS['accent_light'])
        select_all_btn.clicked.connect(self._select_all_needs)
        select_all_btn.setMaximumWidth(120)
        controls.addWidget(select_all_btn)
        
        select_none_btn = GlassButton("Select None", glow_color=COLORS['accent_light'])
        select_none_btn.clicked.connect(self._select_none_needs)
        select_none_btn.setMaximumWidth(120)
        controls.addWidget(select_none_btn)
        
        controls.addStretch()
        
        draft_btn = GlassButton("Create Draft", glow_color=COLORS['accent_light'])
        draft_btn.clicked.connect(lambda: self._email_selected_needs(True))
        draft_btn.setMaximumWidth(130)
        controls.addWidget(draft_btn)
        
        send_btn = GlassButton("Send Email", glow_color=COLORS['status_needed'])
        send_btn.clicked.connect(lambda: self._email_selected_needs(False))
        send_btn.setMaximumWidth(130)
        controls.addWidget(send_btn)
        
        self.needs_column.content_layout.addLayout(controls)
        
        # Scrollable areas with custom styling
        self.needs_scroll = self._create_scroll_area()
        self.needs_list = QWidget()
        self.needs_list.setStyleSheet("background: transparent;")
        self.needs_list_layout = QVBoxLayout(self.needs_list)
        self.needs_list_layout.setAlignment(Qt.AlignTop)
        self.needs_list_layout.setSpacing(10)
        self.needs_scroll.setWidget(self.needs_list)
        self.needs_column.content_layout.addWidget(self.needs_scroll)
        
        self.partial_scroll = self._create_scroll_area()
        self.partial_list = QWidget()
        self.partial_list.setStyleSheet("background: transparent;")
        self.partial_list_layout = QVBoxLayout(self.partial_list)
        self.partial_list_layout.setAlignment(Qt.AlignTop)
        self.partial_list_layout.setSpacing(10)
        self.partial_scroll.setWidget(self.partial_list)
        self.partial_column.content_layout.addWidget(self.partial_scroll)
        
        self.done_scroll = self._create_scroll_area()
        self.done_list = QWidget()
        self.done_list.setStyleSheet("background: transparent;")
        self.done_list_layout = QVBoxLayout(self.done_list)
        self.done_list_layout.setAlignment(Qt.AlignTop)
        self.done_list_layout.setSpacing(10)
        self.done_scroll.setWidget(self.done_list)
        self.done_column.content_layout.addWidget(self.done_scroll)
        
        columns_layout.addWidget(self.needs_column, 0, 0)
        columns_layout.addWidget(self.partial_column, 0, 1)
        columns_layout.addWidget(self.done_column, 0, 2)
        
        parent_layout.addLayout(columns_layout, 1)
    
    def _create_scroll_area(self):
        """Create a styled scroll area"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background: transparent;")
        return scroll
    
    def _apply_styles(self):
        # Global stylesheet with cohesive glassmorphism
        self.setStyleSheet(f"""
            QMainWindow {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2b5278,
                    stop:0.5 #3d6a94,
                    stop:1 #4a7ba7);
            }}
            
            #centralWidget {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2b5278,
                    stop:0.5 #3d6a94,
                    stop:1 #4a7ba7);
            }}
            
            QWidget {{
                color: {COLORS['text_primary']};
            }}
            
            QLineEdit {{
                background-color: rgba(255, 255, 255, 0.15);
                color: {COLORS['text_primary']};
                border: 2px solid rgba(255, 255, 255, 0.3);
                border-radius: 12px;
                padding: 10px 16px;
                font-size: 11pt;
                selection-background-color: rgba(255, 255, 255, 0.3);
            }}
            QLineEdit:focus {{
                border-color: rgba(255, 255, 255, 0.5);
                background-color: rgba(255, 255, 255, 0.2);
            }}
            
            QComboBox {{
                background-color: rgba(255, 255, 255, 0.15);
                color: {COLORS['text_primary']};
                border: 2px solid rgba(255, 255, 255, 0.3);
                border-radius: 12px;
                padding: 8px 16px;
                font-size: 11pt;
                min-height: 28px;
            }}
            QComboBox:hover {{
                border-color: rgba(255, 255, 255, 0.5);
                background-color: rgba(255, 255, 255, 0.2);
            }}
            QComboBox::drop-down {{
                border: none;
                padding-right: 10px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border: none;
            }}
            QComboBox QAbstractItemView {{
                background-color: rgba(60, 100, 140, 0.95);
                color: {COLORS['text_primary']};
                selection-background-color: rgba(255, 255, 255, 0.25);
                border: 2px solid rgba(255, 255, 255, 0.3);
                border-radius: 10px;
                padding: 5px;
            }}
            
            QCheckBox {{
                color: {COLORS['text_primary']};
                spacing: 10px;
                background: transparent;
                font-size: 11pt;
            }}
            QCheckBox::indicator {{
                width: 24px;
                height: 24px;
                border-radius: 8px;
                border: 2px solid rgba(255, 255, 255, 0.4);
                background-color: rgba(255, 255, 255, 0.1);
            }}
            QCheckBox::indicator:checked {{
                background-color: rgba(255, 255, 255, 0.35);
                border-color: rgba(255, 255, 255, 0.6);
            }}
            QCheckBox::indicator:hover {{
                border-color: rgba(255, 255, 255, 0.6);
                background-color: rgba(255, 255, 255, 0.15);
            }}
            
            QLabel {{
                background: transparent;
            }}
            
            QScrollArea {{
                background: transparent;
                border: none;
            }}
            
            QScrollBar:vertical {{
                background-color: rgba(255, 255, 255, 0.08);
                width: 12px;
                border-radius: 6px;
                margin: 0px;
            }}
            QScrollBar::handle:vertical {{
                background-color: rgba(255, 255, 255, 0.25);
                border-radius: 6px;
                min-height: 30px;
            }}
            QScrollBar::handle:vertical:hover {{
                background-color: rgba(255, 255, 255, 0.35);
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
            }}
        """)
    
    def _quick_pair(self):
        self.spring_check.setChecked(False)
        self.summer_check.setChecked(True)
        self.fall_check.setChecked(True)
    
    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Advising Folder")
        if folder:
            self.folder_entry.setText(folder)
    
    def _scan_folder(self):
        folder = Path(self.folder_entry.text().strip())
        if not folder.is_dir():
            QMessageBox.critical(self, "Error", "Please select a valid advising folder.")
            return
        
        self.status_label.setText("Scanning...")
        QApplication.processEvents()
        
        self.snapshots.clear()
        self.needs_checks.clear()
        
        terms = []
        if self.spring_check.isChecked():
            terms.append(("Spring", self.year_combo.currentText()))
        if self.summer_check.isChecked():
            terms.append(("Summer", self.year_combo.currentText()))
        if self.fall_check.isChecked():
            terms.append(("Fall", self.year_combo.currentText()))
        
        if not terms:
            QMessageBox.warning(self, "No Semesters", "Please select at least one semester.")
            self.status_label.setText("Ready")
            return
        
        json_files = list(folder.rglob("*.json"))
        
        for jf in json_files:
            try:
                data = json.loads(jf.read_text(encoding="utf-8"))
            except:
                continue
            
            student_name = data.get("studentName", "Unknown")
            student_id = data.get("studentID", "")
            kctcs_email = data.get("kctcsEmail", "")
            personal_email = data.get("personalEmail", "")
            track = data.get("track", "GT")
            notes = data.get("notes", "")
            
            track_name = TRACK_LABELS.get(track, track)
            
            plan = data.get("semesterPlan", {})
            
            badges = []
            spring_done = False
            summer_done = False
            fall_done = False
            spring_partial = False
            summer_partial = False
            fall_partial = False
            
            for term_name, _year in terms:
                term_key = term_name.lower()
                term_data = plan.get(term_key, {})
                courses = term_data.get("courses", [])
                declined = term_data.get("declined", False)
                not_complete = term_data.get("notComplete", False)
                
                if declined:
                    badge = "[Not Started]"
                elif not_complete:
                    badge = "[In Progress]"
                elif courses:
                    badge = "[Complete]"
                else:
                    badge = "[Not Started]"
                
                badges.append(f"{term_name}: {badge}")
                
                if term_name == "Spring":
                    spring_done = (badge == "[Complete]")
                    spring_partial = (badge == "[In Progress]")
                elif term_name == "Summer":
                    summer_done = (badge == "[Complete]")
                    summer_partial = (badge == "[In Progress]")
                elif term_name == "Fall":
                    fall_done = (badge == "[Complete]")
                    fall_partial = (badge == "[In Progress]")
            
            snap = SnapshotInfo(
                file_path=jf,
                student_name=student_name,
                student_id=student_id,
                kctcs_email=kctcs_email,
                personal_email=personal_email,
                track=track,
                track_name=track_name,
                badges=badges,
                spring_done=spring_done,
                summer_done=summer_done,
                fall_done=fall_done,
                spring_partial=spring_partial,
                summer_partial=summer_partial,
                fall_partial=fall_partial,
                notes=notes
            )
            self.snapshots.append(snap)
        
        self.status_label.setText(f"Scanned {len(json_files)} files")
        self._populate_lists()
    
    def _on_search_changed(self):
        self._populate_lists()
    
    def _on_filter_changed(self):
        self._populate_lists()
    
    def _populate_lists(self):
        # Clear existing
        while self.needs_list_layout.count():
            child = self.needs_list_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        while self.partial_list_layout.count():
            child = self.partial_list_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        while self.done_list_layout.count():
            child = self.done_list_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        # Filter
        search_text = self.search_entry.text().strip().lower()
        track_filter = self.track_combo.currentText()
        
        filtered = []
        for s in self.snapshots:
            if search_text:
                if not (
                    search_text in s.student_name.lower()
                    or search_text in s.student_id.lower()
                    or search_text in s.kctcs_email.lower()
                    or search_text in s.personal_email.lower()
                    or search_text in s.track_name.lower()
                ):
                    continue
            
            if track_filter != "All Tracks":
                if not track_filter.startswith(s.track + ":"):
                    continue
            
            filtered.append(s)
        
        filtered.sort(key=lambda x: (x.track_name, x.student_name.lower()))
        
        # Categorize
        needs_list = []
        partial_list = []
        done_list = []
        
        for s in filtered:
            terms_selected = []
            if self.spring_check.isChecked():
                terms_selected.append("spring")
            if self.summer_check.isChecked():
                terms_selected.append("summer")
            if self.fall_check.isChecked():
                terms_selected.append("fall")
            
            all_selected_done = True
            any_selected_partial = False
            
            for t in terms_selected:
                if t == "spring":
                    if not s.spring_done:
                        all_selected_done = False
                    if s.spring_partial:
                        any_selected_partial = True
                elif t == "summer":
                    if not s.summer_done:
                        all_selected_done = False
                    if s.summer_partial:
                        any_selected_partial = True
                elif t == "fall":
                    if not s.fall_done:
                        all_selected_done = False
                    if s.fall_partial:
                        any_selected_partial = True
            
            if all_selected_done:
                done_list.append(s)
            elif any_selected_partial:
                partial_list.append(s)
            else:
                needs_list.append(s)
        
        # Build lists with track headers
        self._build_list(self.needs_list_layout, needs_list, COLORS['status_needed'], show_checkbox=True)
        self._build_list(self.partial_list_layout, partial_list, COLORS['status_partial'], show_email_btn=True)
        self._build_list(self.done_list_layout, done_list, COLORS['status_complete'])
        
        # Update counts
        self.needs_column.set_title(f"Needs Advised ({len(needs_list)})")
        self.partial_column.set_title(f"Advised (Not Complete) ({len(partial_list)})")
        self.done_column.set_title(f"Advised ({len(done_list)})")
    
    def _build_list(self, layout, students, accent_color, show_checkbox=False, show_email_btn=False):
        current_track = None
        
        for s in students:
            if s.track_name != current_track:
                current_track = s.track_name
                count = sum(1 for x in students if x.track_name == current_track)
                
                header = QLabel(f"  {current_track} ({count})")
                header.setFont(QFont("Segoe UI", 12, QFont.Bold))
                header.setStyleSheet(f"""
                    QLabel {{
                        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 rgba(77, 208, 225, 0.2),
                            stop:1 rgba(100, 181, 246, 0.1));
                        color: {COLORS['text_primary']};
                        padding: 10px;
                        border-radius: 10px;
                        border-left: 4px solid {accent_color};
                    }}
                """)
                layout.addWidget(header)
            
            card = StudentCard(s, accent_color, show_checkbox=show_checkbox, show_email_btn=show_email_btn)
            card.clicked.connect(self._open_student)
            
            if show_checkbox and card.checkbox:
                self.needs_checks[s.student_id] = card.checkbox
            
            layout.addWidget(card)
        
        layout.addStretch()
    
    def _open_student(self, student: SnapshotInfo):
        if not self.server.server:
            self.server.start()
        
        if not self.server.server:
            QMessageBox.critical(self, "Error", "Could not start local server")
            return
        
        url = self.server.get_url(student.file_path)
        webbrowser.open(url)
    
    def _select_all_needs(self):
        for cb in self.needs_checks.values():
            cb.setChecked(True)
    
    def _select_none_needs(self):
        for cb in self.needs_checks.values():
            cb.setChecked(False)
    
    def _email_selected_needs(self, draft: bool):
        selected = [s for s in self.snapshots 
                   if s.student_id in self.needs_checks 
                   and self.needs_checks[s.student_id].isChecked()]
        
        if not selected:
            QMessageBox.information(self, "No Selection", "No students selected.")
            return
        
        if not draft:
            reply = QMessageBox.question(self, "Confirm", 
                                        f"Send emails to {len(selected)} student(s)?",
                                        QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        
        recipients = []
        for s in selected:
            if s.kctcs_email:
                recipients.append(s.kctcs_email)
            elif s.personal_email:
                recipients.append(s.personal_email)
        
        subject = build_email_subject(self.subject_entry.text(), self._term_label())
        body = self._build_email_body()
        
        send_outlook_emails(subject, body, recipients, draft)
        QMessageBox.information(self, "Done", 
                               f"{'Drafts created' if draft else 'Emails sent'} for {len(recipients)} student(s).")
    
    def _term_label(self) -> str:
        year = self.year_combo.currentText()
        terms = []
        if self.spring_check.isChecked():
            terms.append("Spring")
        if self.summer_check.isChecked():
            terms.append("Summer")
        if self.fall_check.isChecked():
            terms.append("Fall")
        
        if not terms:
            return year
        return f"{'/'.join(terms)} {year}"
    
    def _build_email_body(self) -> str:
        term = self._term_label()
        msg = (
            f"<p>Hello,</p>"
            f"<p>This is a reminder that you need to complete your advising appointment for "
            f"{html.escape(term)}.</p>"
        )
        link = self.link_entry.text().strip()
        if link:
            msg += f'<p>Please schedule your appointment here: <a href="{html.escape(link)}">{html.escape(link)}</a></p>'
        else:
            msg += "<p>Please schedule your appointment at your earliest convenience.</p>"
        msg += "<p>Thank you!</p>"
        return msg
    
    def closeEvent(self, event):
        settings = {
            "last_year": self.year_combo.currentText(),
            "last_spring": self.spring_check.isChecked(),
            "last_summer": self.summer_check.isChecked(),
            "last_fall": self.fall_check.isChecked(),
            "last_folder": self.folder_entry.text(),
            "last_track_filter": self.track_combo.currentText(),
            "subject": self.subject_entry.text(),
            "schedulingLink": self.link_entry.text(),
            "window_geometry": self.saveGeometry().toBase64().data().decode(),
            "window_state": "maximized" if self.isMaximized() else "normal"
        }
        save_settings(settings)
        event.accept()


def main():
    # Enable high DPI scaling BEFORE creating QApplication
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application-wide font for consistent scaling
    app.setFont(QFont("Segoe UI", 10))
    
    window = AdvisingDashboard()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
