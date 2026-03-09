"""
Advising Dashboard - Dark Purple/Blue Glassmorphism Theme
Beautiful frosted glass UI with glowing purple effects
Version: 2.1 - FIXED name parsing for firstName/lastName
"""

import json
import platform
import sys
import html
import re
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
    QFrame, QSizePolicy, QMessageBox, QFileDialog, QGridLayout, QGraphicsDropShadowEffect,
    QTextEdit
)
from PySide6.QtCore import Qt, QTimer, Signal, QSize, QRect, QRectF, QPropertyAnimation, QEasingCurve, Property, QByteArray
from PySide6.QtGui import (
    QFont, QPalette, QColor, QPainter, QPainterPath, QLinearGradient, QRadialGradient,
    QBrush, QPen, QPixmap, QIcon
)

try:
    import win32com.client
except Exception:
    win32com = None


APP_TITLE = "Advising Dashboard"
HEADER_TEXT = "One Dashboard To Rule Them All"

# Glassmorphism color palette - balanced blue + purple with black undertones
COLORS = {
    # Backgrounds - black edges with blended midnight blue and deep violet core
    'bg_gradient_1': '#000000',
    'bg_gradient_2': '#02040b',
    'bg_gradient_3': '#050814',
    'bg_gradient_4': '#000000',
    
    # Glass effects - equal blue/purple glow mix
    'glass_bg': 'rgba(82, 103, 255, 0.11)',
    'glass_border': 'rgba(136, 168, 255, 0.34)',
    'glass_glow': 'rgba(115, 104, 255, 0.62)',
    'glass_shadow': 'rgba(6, 8, 25, 0.78)',
    
    # Card backgrounds - dark frosted indigo glass
    'card_bg': 'rgba(4, 6, 16, 0.78)',
    'card_hover': 'rgba(35, 32, 90, 0.50)',
    
    # Text
    'text_primary': '#f8fbff',
    'text_secondary': '#d8ddff',
    'text_muted': '#a5b4fc',
    
    # Accent colors - balanced blue and purple
    'accent_purple': '#8b5cf6',
    'accent_blue': '#4f8cff',
    'accent_pink': '#b794f6',
    
    # Status colors
    'status_needed': '#60a5fa',
    'status_partial': '#fbbf24',
    'status_complete': '#34d399',
    
    # Button colors
    'button_bg': 'rgba(79, 70, 229, 0.34)',
    'button_hover': 'rgba(96, 101, 255, 0.56)',
    'button_border': 'rgba(153, 170, 255, 0.52)',
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

INPUT_WIDGET_STYLE = """
    background-color: #ffffff;
    color: #355cff;
    border: 1px solid rgba(138, 154, 255, 0.55);
    border-radius: 16px;
    padding: 12px 14px;
    selection-background-color: rgba(79, 140, 255, 0.22);
    selection-color: #17306b;
    font-size: 14px;
    font-weight: 600;
"""


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


def html_to_plain_text(value: str) -> str:
    """Convert any saved HTML-ish template content into readable plain text."""
    if not value:
        return ""
    text = str(value)
    if "<" in text and ">" in text:
        text = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", text)
        text = re.sub(r"(?i)</\s*p\s*>", "\n\n", text)
        text = re.sub(r"(?i)<\s*p[^>]*>", "", text)
        text = re.sub(r"(?i)</\s*div\s*>", "\n", text)
        text = re.sub(r"(?i)<\s*div[^>]*>", "", text)
        text = re.sub(r"(?i)</\s*li\s*>", "\n", text)
        text = re.sub(r"(?i)<\s*li[^>]*>", "- ", text)
        text = re.sub(r"(?i)<[^>]+>", "", text)
        text = html.unescape(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def track_filter_value(snapshot: "SnapshotInfo") -> str:
    if not snapshot:
        return ""
    return (snapshot.track_name or snapshot.track or "").strip()


@dataclass
class SnapshotInfo:
    file_path: Path
    first_name: str
    last_name: str
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


def send_outlook_emails(messages: List[Tuple[str, str, str]], draft: bool = False):
    if win32com is None:
        QMessageBox.critical(None, "Error", "Outlook automation not available (win32com not installed)")
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        for addr, subject, body_text in messages:
            mail = outlook.CreateItem(0)
            mail.To = addr
            mail.Subject = subject
            mail.Body = body_text
            if draft:
                mail.Save()
            else:
                mail.Send()
    except Exception as e:
        QMessageBox.critical(None, "Outlook Error", f"Could not send emails:\n{e}")


class GlassCard(QFrame):
    """Glassmorphism card with translucent background and purple glow"""
    def __init__(self, parent=None, glow_color=None, border_width=2):
        super().__init__(parent)
        
        self.glow_color = glow_color or COLORS['glass_border']
        self.border_width = border_width
        
        # Purple-tinted frosted glass
        self.setStyleSheet(f"""
            QFrame {{
                background-color: rgba(8, 11, 28, 0.52);
                border: {border_width}px solid rgba(133, 150, 255, 0.34);
                border-radius: 28px;
            }}
        """)
        
        # Purple glow shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(COLORS['glass_shadow']))
        shadow.setOffset(0, 0)
        self.setGraphicsEffect(shadow)


class GlassButton(QPushButton):
    """Glassmorphism button with purple glow"""
    def __init__(self, text, glow_color=None, parent=None):
        super().__init__(text, parent)

        self.glow_color = glow_color or COLORS['button_bg']
        self.is_hovering = False

        self.setMinimumHeight(44)
        self.setMinimumWidth(130)
        self.setCursor(Qt.PointingHandCursor)
        self.setFont(QFont("Segoe UI", 11, QFont.Bold))

        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setOffset(0, 0)
        self.setGraphicsEffect(self.shadow)

        self._update_style(False)

    def _update_style(self, hovering):
        self.is_hovering = hovering

        if hovering:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {COLORS['button_hover']};
                    color: {COLORS['text_primary']};
                    border: 2px solid {COLORS['accent_purple']};
                    border-radius: 22px;
                    padding: 10px 28px;
                    font-weight: bold;
                }}
            """)
            self.shadow.setBlurRadius(35)
            self.shadow.setColor(QColor(COLORS['accent_purple']))
        else:
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {COLORS['button_bg']};
                    color: {COLORS['text_primary']};
                    border: 2px solid {COLORS['button_border']};
                    border-radius: 22px;
                    padding: 10px 28px;
                    font-weight: bold;
                }}
            """)
            self.shadow.setBlurRadius(25)
            self.shadow.setColor(QColor(COLORS['glass_shadow']))

    def enterEvent(self, event):
        self._update_style(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._update_style(False)
        super().leaveEvent(event)



class XCheckBox(QCheckBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedSize(22, 22)
        self.setStyleSheet("QCheckBox { spacing: 0; background: transparent; } QCheckBox::indicator { width: 0px; height: 0px; }")

    def sizeHint(self):
        return QSize(22, 22)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        r = self.rect().adjusted(1, 1, -1, -1)

        fill = QColor(255, 255, 255, 28)
        border = QColor(255, 255, 255, 102)
        if self.isChecked():
            fill = QColor(102, 123, 255, 180)
            border = QColor(210, 220, 255, 220)
        elif self.underMouse():
            border = QColor(210, 220, 255, 160)

        painter.setPen(QPen(border, 2))
        painter.setBrush(fill)
        painter.drawRoundedRect(r, 7, 7)

        if self.isChecked():
            painter.setPen(QPen(QColor("#ffffff"), 2.4, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            painter.setFont(QFont("Segoe UI", 11, QFont.Bold))
            painter.drawText(r, Qt.AlignCenter, "×")


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
                background-color: rgba(10, 13, 31, 0.72);
                border: 1px solid rgba(138, 154, 255, 0.32);
                border-radius: 18px;
                padding: 12px;
            }}
            QFrame:hover {{
                background-color: rgba(20, 24, 58, 0.84);
                border: 1px solid rgba(138, 170, 255, 0.75);
            }}
        """)
        
        self.setMinimumHeight(110)
        self.setCursor(Qt.PointingHandCursor)
        
        # Stronger shadow for better separation
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(18)
        shadow.setColor(QColor(0, 0, 0, 70))
        shadow.setOffset(0, 3)
        self.setGraphicsEffect(shadow)
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(6)
        
        # Top row
        top_row = QHBoxLayout()
        top_row.setSpacing(10)
        
        if self.show_checkbox:
            self.checkbox = XCheckBox()
            top_row.addWidget(self.checkbox)
        
        # Name
        name_label = QLabel(self.student.student_name)
        name_label.setFont(QFont("Segoe UI", 11, QFont.Bold))
        name_label.setWordWrap(False)
        name_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        name_label.setMinimumWidth(150)
        name_label.setStyleSheet(f"""
            QLabel {{
                color: #ffffff;
                background: transparent;
                padding: 0;
            }}
        """)
        
        # Add text shadow for readability
        name_shadow = QGraphicsDropShadowEffect()
        name_shadow.setBlurRadius(12)
        name_shadow.setColor(QColor(0, 0, 0, 120))
        name_shadow.setOffset(0, 1)
        name_label.setGraphicsEffect(name_shadow)
        
        top_row.addWidget(name_label)
        top_row.addStretch()
        
        # Notes button
        if self.student.notes:
            notes_btn = QPushButton("Notes")
            notes_btn.setToolTip(self.student.notes)
            notes_btn.setCursor(Qt.ArrowCursor)
            notes_btn.setFocusPolicy(Qt.NoFocus)
            notes_btn.setFixedSize(92, 34)
            notes_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {COLORS['button_bg']};
                    color: {COLORS['text_primary']};
                    border: 2px solid {COLORS['button_border']};
                    border-radius: 17px;
                    padding: 6px 16px;
                    font: 700 10pt 'Segoe UI';
                }}
            """)
            top_row.addWidget(notes_btn)
        
        # Email button
        if self.show_email_btn:
            email_btn = QPushButton("Email")
            email_btn.setCursor(Qt.PointingHandCursor)
            email_btn.setFocusPolicy(Qt.NoFocus)
            email_btn.setFixedSize(92, 34)
            email_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {COLORS['button_bg']};
                    color: {COLORS['text_primary']};
                    border: 2px solid {COLORS['button_border']};
                    border-radius: 17px;
                    padding: 6px 16px;
                    font: 700 10pt 'Segoe UI';
                }}
                QPushButton:hover {{
                    background-color: {COLORS['button_hover']};
                    border: 2px solid {COLORS['accent_blue']};
                }}
            """)
            top_row.addWidget(email_btn)
        
        # Student ID
        if self.student.student_id:
            id_label = QLabel(self.student.student_id)
            id_label.setFont(QFont("Consolas", 10))
            id_label.setAlignment(Qt.AlignCenter)
            id_label.setFixedWidth(82)
            id_label.setStyleSheet(f"""
                QLabel {{
                    color: #f8fbff;
                    background-color: rgba(0, 0, 0, 0.42);
                    padding: 6px 10px;
                    border-radius: 12px;
                    border: 1px solid rgba(153, 170, 255, 0.28);
                }}
            """)
            top_row.addWidget(id_label)
        
        layout.addLayout(top_row)
        
        # Badges with better contrast
        badges_label = QLabel("  ".join(self.student.badges))
        badges_label.setFont(QFont("Segoe UI", 10, QFont.DemiBold))
        badges_label.setStyleSheet(f"""
            QLabel {{
                color: #ffffff;
                background-color: rgba(2, 5, 18, 0.92);
                border: 1px solid rgba(108, 129, 255, 0.42);
                border-radius: 10px;
                padding: 8px 12px;
            }}
        """)
        
        # Add text shadow
        badges_shadow = QGraphicsDropShadowEffect()
        badges_shadow.setBlurRadius(10)
        badges_shadow.setColor(QColor(0, 0, 0, 100))
        badges_shadow.setOffset(0, 1)
        badges_label.setGraphicsEffect(badges_shadow)
        
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
        
        # Better glass styling for readability
        self.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['card_bg']};
                border: 2px solid {COLORS['glass_border']};
                border-radius: 28px;
            }}
        """)
        
        # Stronger shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 6)
        self.setGraphicsEffect(shadow)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Header
        header = QFrame()
        header.setMinimumHeight(86)
        header.setMaximumHeight(86)
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
        
        # Title with shadow for readability
        self.title_label = QLabel(title)
        self.title_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        self.title_label.setStyleSheet(f"""
            QLabel {{
                color: {COLORS['text_primary']};
                background: transparent;
            }}
        """)
        self.title_label.setAlignment(Qt.AlignCenter)
        
        # Strong text shadow for readability
        title_shadow = QGraphicsDropShadowEffect()
        title_shadow.setBlurRadius(20)
        title_shadow.setColor(QColor(0, 0, 0, 150))
        title_shadow.setOffset(0, 2)
        self.title_label.setGraphicsEffect(title_shadow)
        
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


class BloomBackground(QWidget):
    """Paints a dark blue-purple gradient with soft bloom lighting."""
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        r = self.rect()

        base = QLinearGradient(0, 0, 0, r.height())
        base.setColorAt(0.0, QColor(COLORS["bg_gradient_1"]))
        base.setColorAt(0.35, QColor(COLORS["bg_gradient_2"]))
        base.setColorAt(0.7, QColor(COLORS["bg_gradient_3"]))
        base.setColorAt(1.0, QColor(COLORS["bg_gradient_4"]))
        painter.fillRect(r, base)

        blooms = [
            ((r.width() * 0.22, r.height() * 0.12), r.width() * 0.24, QColor(79, 140, 255, 26)),
            ((r.width() * 0.78, r.height() * 0.18), r.width() * 0.20, QColor(139, 92, 246, 24)),
            ((r.width() * 0.20, r.height() * 0.64), r.width() * 0.16, QColor(79, 140, 255, 28)),
            ((r.width() * 0.82, r.height() * 0.68), r.width() * 0.15, QColor(139, 92, 246, 28)),
            ((r.width() * 0.52, r.height() * 0.90), r.width() * 0.24, QColor(100, 120, 255, 20)),
        ]
        painter.setPen(Qt.NoPen)
        for (cx, cy), radius, color in blooms:
            grad = QRadialGradient(cx, cy, radius)
            grad.setColorAt(0.0, color)
            grad.setColorAt(0.45, QColor(color.red(), color.green(), color.blue(), max(10, color.alpha() // 2)))
            grad.setColorAt(1.0, QColor(color.red(), color.green(), color.blue(), 0))
            painter.setBrush(QBrush(grad))
            painter.drawEllipse(QRectF(cx - radius, cy - radius, radius * 2, radius * 2))

        painter.setBrush(Qt.NoBrush)
        painter.setPen(QPen(QColor(255, 255, 255, 16), 2))
        painter.drawRoundedRect(QRectF(24, 24, r.width() - 48, r.height() - 48), 30, 30)


class AdvisingDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle(APP_TITLE)
        
        # Default to slightly under the available screen height so the Windows taskbar
        # does not make the bottom edge feel cut off.
        screen = QApplication.primaryScreen().availableGeometry()
        default_width = min(1920, max(1280, screen.width() - 40))
        default_height = min(1028, max(720, screen.height() - 40))
        self.resize(default_width, default_height)

        # Set minimum size for proper scaling
        self.setMinimumSize(1280, 720)

        # Center window on screen
        x = screen.x() + max(0, (screen.width() - default_width) // 2)
        y = screen.y() + max(0, (screen.height() - default_height) // 2)
        self.move(x, y)
        
        # Variables
        self.snapshots: List[SnapshotInfo] = []
        self.needs_checks: Dict[str, QCheckBox] = {}
        
        # Settings
        self.settings = load_settings()
        settings = self.settings
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
        central = BloomBackground()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(28, 26, 28, 26)
        main_layout.setSpacing(18)

        self._build_header(main_layout)
        self._build_control_panel(main_layout)
        self._build_columns(main_layout)

        main_layout.setStretch(0, 0)
        main_layout.setStretch(1, 1)
        main_layout.setStretch(2, 3)
    
    def _build_header(self, parent_layout):
        header_card = GlassCard(glow_color=COLORS['glass_glow'])
        header_card.setMinimumHeight(88)
        header_card.setMaximumHeight(96)
        header_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(header_card)
        layout.setContentsMargins(34, 10, 34, 10)
        layout.setSpacing(0)

        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.addStretch(1)

        title = QLabel("One Dashboard to Rule Them All")
        title.setWordWrap(False)
        title.setAlignment(Qt.AlignCenter)
        title.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Preferred)
        title.setFont(QFont("Segoe UI", 24, QFont.Bold))
        title.setStyleSheet(f"""
            color: {COLORS['text_primary']};
            background: transparent;
            border: none;
            padding: 0;
            margin: 0;
        """)
        title_shadow = QGraphicsDropShadowEffect()
        title_shadow.setBlurRadius(42)
        title_shadow.setColor(QColor(COLORS['accent_blue']))
        title_shadow.setOffset(0, 0)
        title.setGraphicsEffect(title_shadow)
        title_row.addWidget(title, 0, Qt.AlignCenter)
        title_row.addStretch(1)

        layout.addStretch(1)
        layout.addLayout(title_row)
        layout.addStretch(1)

        parent_layout.addWidget(header_card)

    def _build_control_panel(self, parent_layout):
        panel = GlassCard(glow_color=COLORS['glass_glow'])
        panel.setMaximumHeight(360)
        panel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(panel)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(16)

        top_row = QHBoxLayout()
        top_row.setSpacing(16)

        year_widget = self._create_control_group("Academic Year")
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(y) for y in range(2026, 2041)])
        self.year_combo.setCurrentText(self.current_year)
        self.year_combo.setStyleSheet(f"""
            QComboBox {{{INPUT_WIDGET_STYLE}}}
            QComboBox QAbstractItemView {{
                background-color: #ffffff;
                color: #355cff;
                border: 1px solid rgba(138, 154, 255, 0.45);
                selection-background-color: rgba(79, 140, 255, 0.14);
                selection-color: #17306b;
                outline: 0;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 26px;
                background: transparent;
            }}
        """)
        year_widget.layout().addWidget(self.year_combo)
        top_row.addWidget(year_widget, 0)

        sem_widget = self._create_control_group("Advising Terms")
        sem_checks = QHBoxLayout()
        sem_checks.setSpacing(14)
        self.spring_check = QCheckBox("Spring")
        self.spring_check.setChecked(self.spring_enabled)
        sem_checks.addWidget(self.spring_check)
        self.summer_check = QCheckBox("Summer")
        self.summer_check.setChecked(self.summer_enabled)
        sem_checks.addWidget(self.summer_check)
        self.fall_check = QCheckBox("Fall")
        self.fall_check.setChecked(self.fall_enabled)
        sem_checks.addWidget(self.fall_check)
        sem_checks.addStretch()
        sem_widget.layout().addLayout(sem_checks)
        top_row.addWidget(sem_widget, 1)

        quick_btn = GlassButton("Quick: Summer + Fall")
        quick_btn.clicked.connect(self._quick_pair)
        quick_btn.setMaximumWidth(220)
        top_row.addWidget(quick_btn, 0)

        search_widget = self._create_control_group("Search")
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("Name, ID, Email...")
        self.search_entry.setStyleSheet(f"QLineEdit {{{INPUT_WIDGET_STYLE}}}")
        self.search_entry.textChanged.connect(self._on_search_changed)
        search_widget.layout().addWidget(self.search_entry)
        top_row.addWidget(search_widget, 1)

        track_widget = self._create_control_group("Track Filter")
        self.track_combo = QComboBox()
        self.track_combo.addItem("All Tracks")
        self.track_combo.setStyleSheet(f"""
            QComboBox {{{INPUT_WIDGET_STYLE}}}
            QComboBox QAbstractItemView {{
                background-color: #ffffff;
                color: #355cff;
                border: 1px solid rgba(138, 154, 255, 0.45);
                selection-background-color: rgba(79, 140, 255, 0.14);
                selection-color: #17306b;
                outline: 0;
            }}
            QComboBox::drop-down {{
                border: none;
                width: 26px;
                background: transparent;
            }}
        """)
        self.track_combo.setCurrentText(self.track_filter)
        self.track_combo.currentTextChanged.connect(self._on_filter_changed)
        self.year_combo.currentTextChanged.connect(self._populate_lists)
        self.spring_check.toggled.connect(self._populate_lists)
        self.summer_check.toggled.connect(self._populate_lists)
        self.fall_check.toggled.connect(self._populate_lists)
        track_widget.layout().addWidget(self.track_combo)
        top_row.addWidget(track_widget, 1)

        layout.addLayout(top_row)

        middle_row = QHBoxLayout()
        middle_row.setSpacing(12)

        folder_widget = self._create_control_group("Advising Folder")
        self.folder_entry = QLineEdit(self.folder_path)
        self.folder_entry.setStyleSheet(f"QLineEdit {{{INPUT_WIDGET_STYLE}}}")
        folder_widget.layout().addWidget(self.folder_entry)
        middle_row.addWidget(folder_widget, 1)

        browse_btn = GlassButton("Browse")
        browse_btn.clicked.connect(self._browse_folder)
        browse_btn.setMaximumWidth(120)
        middle_row.addWidget(browse_btn)

        scan_btn = GlassButton("Scan Folder", glow_color=COLORS['status_complete'])
        scan_btn.clicked.connect(self._scan_folder)
        scan_btn.setMaximumWidth(150)
        middle_row.addWidget(scan_btn)

        layout.addLayout(middle_row)

        lower_row = QHBoxLayout()
        lower_row.setSpacing(16)

        subj_widget = self._create_control_group("Email Subject")
        self.subject_entry = QLineEdit(self.email_subject)
        self.subject_entry.setStyleSheet(f"QLineEdit {{{INPUT_WIDGET_STYLE}}}")
        subj_widget.layout().addWidget(self.subject_entry)
        lower_row.addWidget(subj_widget, 1)

        link_widget = self._create_control_group("Scheduling Link")
        self.link_entry = QLineEdit(self.scheduling_link)
        self.link_entry.setStyleSheet(f"QLineEdit {{{INPUT_WIDGET_STYLE}}}")
        link_widget.layout().addWidget(self.link_entry)
        lower_row.addWidget(link_widget, 1)

        self.status_label = QLabel("Ready")
        self.status_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.status_label.setStyleSheet(f"color: {COLORS['text_secondary']}; background: transparent; border: none; padding: 0; margin: 0;")
        lower_row.addWidget(self.status_label, 0, Qt.AlignBottom)

        layout.addLayout(lower_row)

        body_widget = self._create_control_group("Email Body (plain text)")
        self.email_body = QTextEdit()
        self.email_body.setPlaceholderText("Plain text email. Use {term}, {first_name}, or {student_name}.")
        self.email_body.setStyleSheet(f"QTextEdit {{{INPUT_WIDGET_STYLE}}}")
        self.email_body.setAcceptRichText(False)
        self.email_body.setLineWrapMode(QTextEdit.WidgetWidth)
        self.email_body.setMinimumHeight(88)
        self.email_body.setMaximumHeight(100)

        default_body = (
            "Hello {first_name},\n\n"
            "This is a reminder that you need to complete your advising appointment for {term}.\n\n"
            "Please schedule your appointment at your earliest convenience.\n\n"
            "Thank you!"
        )
        saved_body = html_to_plain_text(self.settings.get("email_body", default_body))
        if not saved_body:
            saved_body = default_body
        self.email_body.setPlainText(saved_body)
        body_widget.layout().addWidget(self.email_body)
        layout.addWidget(body_widget)

        parent_layout.addWidget(panel)


    def _quick_pair(self):
        """Quick-select Summer + Fall terms and refresh the filtered view."""
        self.spring_check.setChecked(False)
        self.summer_check.setChecked(True)
        self.fall_check.setChecked(True)
        if hasattr(self, "status_label"):
            self.status_label.setText("Quick pair applied: Summer and Fall")
        self._refresh_columns()

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
                color: {COLORS['text_primary']};
                background: transparent;
                border: none;
                padding: 0;
                margin: 0;
            }}
        """)
        
        layout.addWidget(label)
        
        return widget
    
    def _build_columns(self, parent_layout):
        columns_layout = QGridLayout()
        columns_layout.setHorizontalSpacing(18)
        columns_layout.setVerticalSpacing(18)
        columns_layout.setColumnStretch(0, 1)
        columns_layout.setColumnStretch(1, 1)
        columns_layout.setColumnStretch(2, 1)

        self.needs_column = ColumnCard("Needs Advised (0)", COLORS['status_needed'])
        self.needs_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.needs_column.setMinimumHeight(520)

        self.partial_column = ColumnCard("Advised (Not Complete) (0)", COLORS['status_partial'])
        self.partial_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.partial_column.setMinimumHeight(520)

        self.done_column = ColumnCard("Advised (0)", COLORS['status_complete'])
        self.done_column.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.done_column.setMinimumHeight(520)

        controls = QHBoxLayout()
        controls.setSpacing(10)
        select_all_btn = GlassButton("Select All")
        select_all_btn.clicked.connect(self._select_all_needs)
        select_all_btn.setMaximumWidth(112)
        controls.addWidget(select_all_btn)

        select_none_btn = GlassButton("Select None")
        select_none_btn.clicked.connect(self._select_none_needs)
        select_none_btn.setMaximumWidth(118)
        controls.addWidget(select_none_btn)
        controls.addStretch()

        draft_btn = GlassButton("Create Draft")
        draft_btn.clicked.connect(lambda: self._email_selected_needs(True))
        draft_btn.setMaximumWidth(130)
        controls.addWidget(draft_btn)

        send_btn = GlassButton("Send Email", glow_color=COLORS['status_needed'])
        send_btn.clicked.connect(lambda: self._email_selected_needs(False))
        send_btn.setMaximumWidth(130)
        controls.addWidget(send_btn)
        self.needs_column.content_layout.addLayout(controls)

        self.needs_scroll = self._create_scroll_area()
        self.needs_list = QWidget()
        self.needs_list.setStyleSheet("background: transparent;")
        self.needs_list_layout = QVBoxLayout(self.needs_list)
        self.needs_list_layout.setAlignment(Qt.AlignTop)
        self.needs_list_layout.setSpacing(12)
        self.needs_scroll.setWidget(self.needs_list)
        self.needs_column.content_layout.addWidget(self.needs_scroll, 1)

        self.partial_scroll = self._create_scroll_area()
        self.partial_list = QWidget()
        self.partial_list.setStyleSheet("background: transparent;")
        self.partial_list_layout = QVBoxLayout(self.partial_list)
        self.partial_list_layout.setAlignment(Qt.AlignTop)
        self.partial_list_layout.setSpacing(12)
        self.partial_scroll.setWidget(self.partial_list)
        self.partial_column.content_layout.addWidget(self.partial_scroll, 1)

        self.done_scroll = self._create_scroll_area()
        self.done_list = QWidget()
        self.done_list.setStyleSheet("background: transparent;")
        self.done_list_layout = QVBoxLayout(self.done_list)
        self.done_list_layout.setAlignment(Qt.AlignTop)
        self.done_list_layout.setSpacing(12)
        self.done_scroll.setWidget(self.done_list)
        self.done_column.content_layout.addWidget(self.done_scroll, 1)

        columns_layout.addWidget(self.needs_column, 0, 0)
        columns_layout.addWidget(self.partial_column, 0, 1)
        columns_layout.addWidget(self.done_column, 0, 2)

        parent_layout.addLayout(columns_layout, 1)

    def _create_scroll_area(self):
        """Create a styled scroll area"""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background: transparent;")
        return scroll
    
    def _apply_styles(self):
        self.setStyleSheet(f"""
            QMainWindow, #centralWidget {{
                background: transparent;
            }}

            QLabel {{
                background: transparent;
                border: none;
            }}

            QLineEdit, QComboBox, QTextEdit {{
                background-color: #ffffff;
                color: #355cff;
                border: 1px solid rgba(138, 154, 255, 0.55);
                border-radius: 16px;
                padding: 12px 14px;
                selection-background-color: rgba(79, 140, 255, 0.22);
                selection-color: #17306b;
                font-size: 14px;
                font-weight: 600;
            }}

            QLineEdit::placeholder, QTextEdit[placeholderText="true"] {{
                color: #6a79b8;
            }}

            QLineEdit:focus, QComboBox:focus, QTextEdit:focus {{
                border: 1px solid rgba(79, 140, 255, 0.95);
                background-color: #ffffff;
                color: #355cff;
            }}

            QComboBox QAbstractItemView {{
                background-color: #ffffff;
                color: #355cff;
                border: 1px solid rgba(138, 154, 255, 0.45);
                selection-background-color: rgba(79, 140, 255, 0.14);
                selection-color: #17306b;
                outline: 0;
            }}

            QToolTip {{
                background-color: rgba(10, 12, 24, 0.96);
                color: #f8fbff;
                border: 1px solid rgba(138, 154, 255, 0.65);
                padding: 8px 10px;
                border-radius: 8px;
                font-size: 12px;
            }}

            QComboBox::drop-down {{
                border: none;
                width: 26px;
                background: transparent;
            }}

            QComboBox QAbstractItemView {{
                background-color: #ffffff;
                color: #355cff;
                border: 1px solid rgba(138, 154, 255, 0.45);
                padding: 8px;
                selection-background-color: rgba(79, 140, 255, 0.14);
                selection-color: #17306b;
            }}

            QCheckBox {{
                color: {COLORS['text_primary']};
                spacing: 8px;
                font-size: 14px;
            }}

            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border-radius: 9px;
                border: 1px solid rgba(170, 186, 255, 0.40);
                background-color: rgba(255, 255, 255, 0.06);
            }}

            QCheckBox::indicator:checked {{
                background-color: rgba(102, 123, 255, 0.65);
                border: 1px solid rgba(192, 202, 255, 0.75);
            }}

            QScrollArea {{
                background: transparent;
                border: none;
            }}

            QScrollBar:vertical {{
                background: rgba(0, 0, 0, 0.18);
                width: 12px;
                margin: 2px 0 2px 0;
                border-radius: 6px;
            }}

            QScrollBar::handle:vertical {{
                background: rgba(121, 139, 255, 0.55);
                min-height: 28px;
                border-radius: 6px;
            }}

            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
                border: none;
                height: 0px;
            }}
        """)

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Advising Folder")
        if folder:
            self.folder_entry.setText(folder)
    
    def _scan_folder(self):
        folder_text = self.folder_entry.text().strip()
        if not folder_text:
            chosen = QFileDialog.getExistingDirectory(self, "Select Advising Folder")
            if not chosen:
                return
            self.folder_entry.setText(chosen)
            folder_text = chosen

        folder = Path(folder_text)
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
        
        json_files = list(folder.rglob("*.json")) + list(folder.rglob("*.JSON"))
        if not json_files:
            self.status_label.setText("Scanned 0 files")
            self._refresh_track_filter_options()
            self._populate_lists()
            QMessageBox.information(self, "No JSON Files", "No JSON files were found in that folder.")
            return

        for jf in json_files:
            try:
                data = json.loads(jf.read_text(encoding="utf-8"))
            except Exception:
                continue

            student_block = data.get("student", {}) if isinstance(data.get("student"), dict) else {}
            selection_block = data.get("selection", {}) if isinstance(data.get("selection"), dict) else {}
            data_block = data.get("data", {}) if isinstance(data.get("data"), dict) else {}

            # FIXED: Comprehensive name parsing that handles all formats
            first_name = ""
            last_name = ""
            student_name = ""
            
            # Try firstName/lastName fields first (preferred format)
            first_name = str(student_block.get("firstName") or data.get("firstName") or "").strip()
            last_name = str(student_block.get("lastName") or data.get("lastName") or "").strip()
            
            # Build full name from first/last
            if first_name or last_name:
                student_name = " ".join(part for part in [first_name, last_name] if part).strip()
            
            # Fall back to single 'name' field if firstName/lastName are empty
            if not student_name:
                student_name = str(student_block.get("name") or data.get("studentName") or "").strip()
                
                # If we got a name from the single field, split it for email personalization
                if student_name:
                    name_parts = student_name.split(None, 1)  # Split on first space
                    first_name = name_parts[0] if name_parts else student_name
                    last_name = name_parts[1] if len(name_parts) > 1 else ""
            
            # Final fallback: use filename
            if not student_name:
                student_name = jf.stem
                first_name = jf.stem
                last_name = ""

            student_id = str(student_block.get("studentId") or data.get("studentId") or data.get("studentID") or "").strip()
            kctcs_email = str(student_block.get("kctcsEmail") or data.get("kctcsEmail") or "").strip()
            personal_email = str(student_block.get("personalEmail") or data.get("personalEmail") or "").strip()
            track = str(selection_block.get("scenario") or data.get("track") or "GT").strip() or "GT"
            notes = str(data_block.get("notes") or data.get("notes") or student_block.get("notes") or "").strip()

            track_name = TRACK_LABELS.get(track, track)

            semester_plans = data_block.get("semesterPlans")
            if isinstance(semester_plans, list):
                plan_lookup = {}
                for item in semester_plans:
                    if not isinstance(item, dict):
                        continue
                    season = str(item.get("season", "")).strip().lower()
                    year = str(item.get("year", "")).strip()
                    if season:
                        plan_lookup[(season, year)] = item
                        if season not in plan_lookup:
                            plan_lookup[season] = item
            else:
                raw_plan = data.get("semesterPlan", {}) if isinstance(data.get("semesterPlan"), dict) else {}
                plan_lookup = {}
                for season in ("spring", "summer", "fall"):
                    item = raw_plan.get(season, {})
                    if isinstance(item, dict):
                        plan_lookup[season] = item

            badges = []
            spring_done = False
            summer_done = False
            fall_done = False
            spring_partial = False
            summer_partial = False
            fall_partial = False

            for term_name, year in terms:
                term_key = term_name.lower()
                term_data = plan_lookup.get((term_key, year), plan_lookup.get(term_key, {}))
                if not isinstance(term_data, dict):
                    term_data = {}
                courses = term_data.get("courses", [])
                declined = bool(term_data.get("declined", False))
                not_complete = bool(term_data.get("notComplete", False))

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
                    spring_done = badge == "[Complete]"
                    spring_partial = badge == "[In Progress]"
                elif term_name == "Summer":
                    summer_done = badge == "[Complete]"
                    summer_partial = badge == "[In Progress]"
                elif term_name == "Fall":
                    fall_done = badge == "[Complete]"
                    fall_partial = badge == "[In Progress]"

            snap = SnapshotInfo(
                file_path=jf,
                first_name=first_name,
                last_name=last_name,
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

        self.status_label.setText(f"Scanned {len(self.snapshots)} student(s)")
        if self.search_entry.text().strip():
            self.search_entry.clear()
        self.track_filter = "All Tracks"
        self._refresh_track_filter_options()
        self.track_combo.setCurrentIndex(0)
        self._populate_lists()
    
    def _refresh_track_filter_options(self):
        current = self.track_combo.currentText() if hasattr(self, "track_combo") else "All Tracks"
        tracks = sorted({track_filter_value(s) for s in self.snapshots if track_filter_value(s)})

        self.track_combo.blockSignals(True)
        self.track_combo.clear()
        self.track_combo.addItem("All Tracks")
        self.track_combo.addItems(tracks)

        preferred = self.track_filter if self.track_filter in tracks or self.track_filter == "All Tracks" else current
        if preferred in tracks or preferred == "All Tracks":
            self.track_combo.setCurrentText(preferred)
        else:
            self.track_combo.setCurrentIndex(0)

        self.track_combo.blockSignals(False)

    def _on_search_changed(self):
        self._populate_lists()
    
    def _on_filter_changed(self):
        self._populate_lists()
    
    def _refresh_columns(self):
        self._populate_lists()

    def _populate_lists(self):
        self.needs_checks = {}
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
                if track_filter_value(s) != track_filter:
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
                        color: {COLORS['text_primary']};
                        padding: 10px 14px;
                        border-radius: 14px;
                        border: 1px solid rgba(148, 166, 255, 0.42);
                        background: qlineargradient(
                            x1:0, y1:0, x2:1, y2:0,
                            stop:0 rgba(40, 62, 168, 0.92),
                            stop:0.55 rgba(66, 58, 168, 0.92),
                            stop:1 rgba(89, 50, 156, 0.92)
                        );
                    }}
                """)
                
                # Add text shadow
                header_shadow = QGraphicsDropShadowEffect()
                header_shadow.setBlurRadius(20)
                header_shadow.setColor(QColor(35, 20, 100, 110))
                header_shadow.setOffset(0, 1)
                header.setGraphicsEffect(header_shadow)
                
                layout.addWidget(header)
            
            card = StudentCard(s, accent_color, show_checkbox=show_checkbox, show_email_btn=show_email_btn)
            card.clicked.connect(self._open_student)
            
            if show_checkbox and card.checkbox:
                self.needs_checks[str(s.file_path)] = card.checkbox
            
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
                   if str(s.file_path) in self.needs_checks
                   and self.needs_checks[str(s.file_path)].isChecked()]
        
        if not selected:
            QMessageBox.information(self, "No Selection", "No students selected.")
            return
        
        if not draft:
            reply = QMessageBox.question(self, "Confirm", 
                                        f"Send emails to {len(selected)} student(s)?",
                                        QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        
        messages = []
        term_label = self._term_label()
        for s in selected:
            recipient = s.kctcs_email or s.personal_email
            if not recipient:
                continue
            subject = build_email_subject(self.subject_entry.text(), term_label)
            body = self._build_email_body(s)
            messages.append((recipient, subject, body))

        if not messages:
            QMessageBox.information(self, "No Email Address", "No selected students had an email address.")
            return

        send_outlook_emails(messages, draft)
        QMessageBox.information(self, "Done", 
                               f"{'Drafts created' if draft else 'Emails sent'} for {len(messages)} student(s).")
    
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
        if len(terms) == 1:
            term_text = terms[0]
        elif len(terms) == 2:
            term_text = f"{terms[0]} and {terms[1]}"
        else:
            term_text = f"{', '.join(terms[:-1])}, and {terms[-1]}"
        return f"{term_text} {year}"
    
    def _build_email_body(self, student: Optional[SnapshotInfo] = None) -> str:
        term = self._term_label()
        body_template = html_to_plain_text(self.email_body.toPlainText()).strip()
        first_name = "Student"
        student_name = "Student"
        if student is not None:
            first_name = (student.first_name or student.student_name or "Student").strip()
            student_name = (student.student_name or first_name or "Student").strip()

        body = body_template.replace("{term}", term).replace("{TERM}", term)
        body = body.replace("{first_name}", first_name).replace("{FIRST_NAME}", first_name)
        body = body.replace("{student_name}", student_name).replace("{STUDENT_NAME}", student_name)

        if not body.lower().startswith("hello"):
            body = f"Hello {first_name},\n\n" + body

        link = self.link_entry.text().strip()
        if link:
            body = body.rstrip() + f"\n\nSchedule here: {link}"

        return body
    
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
            "email_body": html_to_plain_text(self.email_body.toPlainText()),
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
