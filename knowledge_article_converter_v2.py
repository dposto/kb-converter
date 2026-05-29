#!/usr/bin/env python3
"""
Knowledge Article Converter
Converts HTML knowledge base articles to formatted DOCX files
"""

__version__ = "2.0"
__author__ = "David Posto"

import subprocess
import os
import json
import re
import time
import fnmatch
import hashlib
from collections import Counter
from pathlib import Path
from threading import Thread
import base64

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QListWidget, QListWidgetItem, QProgressBar, QMenuBar,
    QMenu, QFileDialog, QMessageBox, QDialog, QDialogButtonBox,
    QSplitter, QStatusBar, QTextEdit
)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QAction, QActionGroup, QColor, QPainter, QPen

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Embedded reference document (base64)
REFERENCE_DOCX_B64 = ""  # Will be filled during build

# Embedded Lua filter
LUA_FILTER = """-- FILTER CONTENT WILL BE EMBEDDED HERE --"""


# ─── Helper Functions ─────────────────────────────────────────────────────────

def analyze_html(content):
    """Analyze HTML content and return tag counts and structure info."""
    tags = re.findall(r'<(\w+)[\s>]', content)
    counts = Counter(t.lower() for t in tags)
    info = {
        'total_tags': len(tags),
        'unique_tags': len(counts),
        'headings': sum(counts.get(f'h{i}', 0) for i in range(1, 7)),
        'paragraphs': counts.get('p', 0),
        'tables': counts.get('table', 0),
        'rows': counts.get('tr', 0),
        'images': counts.get('img', 0),
        'links': counts.get('a', 0),
        'lists': counts.get('ul', 0) + counts.get('ol', 0),
        'list_items': counts.get('li', 0),
        'divs': counts.get('div', 0),
        'spans': counts.get('span', 0),
        'bold': counts.get('b', 0) + counts.get('strong', 0),
        'italic': counts.get('i', 0) + counts.get('em', 0),
    }
    # Detect encoding from meta tag
    enc_match = re.search(r'charset=["\']?([^"\'\s;>]+)', content, re.IGNORECASE)
    info['declared_encoding'] = enc_match.group(1) if enc_match else None
    return info


def file_hash(filepath, algo='sha256'):
    """Compute file hash."""
    h = hashlib.new(algo)
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def detect_system_dark_mode():
    """Detect if the desktop environment is using a dark theme."""
    try:
        result = subprocess.run([
            'dbus-send', '--session', '--print-reply=literal',
            '--dest=org.freedesktop.portal.Desktop',
            '/org/freedesktop/portal/desktop',
            'org.freedesktop.portal.Settings.Read',
            'string:org.freedesktop.appearance',
            'string:color-scheme'
        ], capture_output=True, text=True, timeout=2)
        if result.returncode == 0 and 'uint32 1' in result.stdout:
            return True
        if result.returncode == 0 and ('uint32 0' in result.stdout or
                                        'uint32 2' in result.stdout):
            return False
    except Exception:
        pass
    try:
        result = subprocess.run(
            ['gsettings', 'get', 'org.gnome.desktop.interface', 'color-scheme'],
            capture_output=True, text=True, timeout=2)
        if 'prefer-dark' in result.stdout:
            return True
    except Exception:
        pass
    try:
        result = subprocess.run(
            ['gsettings', 'get', 'org.gnome.desktop.interface', 'gtk-theme'],
            capture_output=True, text=True, timeout=2)
        theme = result.stdout.strip().strip("'\"").lower()
        if any(dark in theme for dark in ['dark', 'noir', 'night', 'gruvbox', 'dracula']):
            return True
    except Exception:
        pass
    return False


def extract_article_info(html_path):
    """Extract article number and title from the HTML <h1> tag."""
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            content = f.read()
        m = re.search(r'<h1[^>]*>\s*(?:\[.*?\]\s*)?(\d+)\s*-\s*(.+?)\s*</h1>',
                      content, re.IGNORECASE)
        if m:
            article_num = m.group(1)
            clean_title = m.group(2).strip()
            full_title = f"{article_num} - {clean_title}"
            return article_num, full_title, clean_title
    except Exception:
        pass
    return None, None, None


def strip_self_service_flow(html_content):
    """Remove the Self Service Flow section from HTML content."""
    pattern = r'<h2>\s*Self Service Flow\s*</h2>\s*<p>.*?</p>'
    return re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)


def strip_ai_usage(html_content):
    """Remove the AI Usage section from HTML content."""
    pattern = r'<h2>\s*AI Usage\s*</h2>\s*<p>.*?</p>'
    return re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)


def sanitize_filename(name, max_length=100):
    """Make a filename safe for Windows and Linux."""
    name = name.replace('/', '-')
    name = re.sub(r'[\\:*?"<>|]', '', name)
    name = re.sub(r'  +', ' ', name)
    name = name.strip()
    if len(name) > max_length:
        name = name[:max_length].rsplit(' ', 1)[0]
    name = name.rstrip('. ')
    return name


# ─── Themes ───────────────────────────────────────────────────────────────────

LIGHT_THEME = """
QMainWindow, QWidget {
    background-color: #f2f2f7;
    color: #1c1c1e;
}
QMenuBar {
    background-color: #e5e5ea;
    color: #1c1c1e;
    border-bottom: 1px solid #d1d1d6;
}
QMenuBar::item:selected {
    background-color: #0061AC;
    color: #ffffff;
}
QMenu {
    background-color: #ffffff;
    color: #1c1c1e;
    border: 1px solid #d1d1d6;
}
QMenu::item:selected {
    background-color: #0061AC;
    color: #ffffff;
}
QMenu::separator {
    height: 1px;
    background-color: #d1d1d6;
    margin: 4px 8px;
}
QListWidget {
    background-color: #ffffff;
    color: #1c1c1e;
    border: 1px solid #d1d1d6;
    font-family: monospace;
    font-size: 10pt;
}
QListWidget::item:selected {
    background-color: #0061AC;
    color: #ffffff;
}
QProgressBar {
    background-color: #e5e5ea;
    border: none;
    border-radius: 4px;
    height: 8px;
}
QProgressBar::chunk {
    background-color: #0061AC;
    border-radius: 4px;
}
QStatusBar {
    background-color: #e5e5ea;
    color: #6e6e73;
    border-top: 1px solid #d1d1d6;
}
QLabel#sectionHeader {
    font-weight: bold;
    font-size: 11pt;
    padding: 4px 0px;
}
QSplitter::handle {
    background-color: #d1d1d6;
}
QTextEdit#terminalLog {
    background-color: #0a0a0a;
    color: #33ff33;
    border: 2px solid #1a3a1a;
    border-radius: 6px;
    font-family: monospace;
    font-size: 9pt;
    padding: 6px 8px;
    selection-background-color: #33ff33;
    selection-color: #0a0a0a;
}
"""

DARK_THEME = """
QMainWindow, QWidget {
    background-color: #1c1c1e;
    color: #f2f2f7;
}
QMenuBar {
    background-color: #2c2c2e;
    color: #f2f2f7;
    border-bottom: 1px solid #3a3a3c;
}
QMenuBar::item:selected {
    background-color: #3a9bdc;
    color: #ffffff;
}
QMenu {
    background-color: #2c2c2e;
    color: #f2f2f7;
    border: 1px solid #3a3a3c;
}
QMenu::item:selected {
    background-color: #3a9bdc;
    color: #ffffff;
}
QMenu::separator {
    height: 1px;
    background-color: #48484a;
    margin: 4px 8px;
}
QListWidget {
    background-color: #2c2c2e;
    color: #f2f2f7;
    border: 1px solid #3a3a3c;
    font-family: monospace;
    font-size: 10pt;
}
QListWidget::item:selected {
    background-color: #3a9bdc;
    color: #ffffff;
}
QProgressBar {
    background-color: #3a3a3c;
    border: none;
    border-radius: 4px;
    height: 8px;
}
QProgressBar::chunk {
    background-color: #3a9bdc;
    border-radius: 4px;
}
QStatusBar {
    background-color: #2c2c2e;
    color: #98989d;
    border-top: 1px solid #3a3a3c;
}
QLabel#sectionHeader {
    font-weight: bold;
    font-size: 11pt;
    padding: 4px 0px;
}
QSplitter::handle {
    background-color: #48484a;
}
QTextEdit#terminalLog {
    background-color: #0a0a0a;
    color: #33ff33;
    border: 2px solid #1a3a1a;
    border-radius: 6px;
    font-family: monospace;
    font-size: 9pt;
    padding: 6px 8px;
    selection-background-color: #33ff33;
    selection-color: #0a0a0a;
}
"""


# ─── CRT Terminal ─────────────────────────────────────────────────────────────

class CRTTerminal(QTextEdit):
    """QTextEdit styled as an old green phosphor CRT monitor with scanlines."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setObjectName("terminalLog")
        self.viewport().setAttribute(
            Qt.WidgetAttribute.WA_TranslucentBackground, False)

    def paintEvent(self, event):
        super().paintEvent(event)
        # Draw scanlines over the viewport
        painter = QPainter(self.viewport())
        pen = QPen(QColor(0, 0, 0, 30))
        pen.setWidth(1)
        painter.setPen(pen)
        for y in range(0, self.viewport().height(), 3):
            painter.drawLine(0, y, self.viewport().width(), y)
        painter.end()


# ─── Signals ──────────────────────────────────────────────────────────────────

class ConvertSignals(QObject):
    """Signals for thread-safe UI updates from conversion threads."""
    add_source = pyqtSignal(str)
    add_output = pyqtSignal(str, str)  # text, color
    set_status = pyqtSignal(str)
    progress = pyqtSignal(int)
    complete = pyqtSignal()
    register_click = pyqtSignal(str, str)  # action_type, path
    spinner_stop = pyqtSignal()
    auto_convert_file = pyqtSignal(str)  # html file path from watchdog
    log = pyqtSignal(str)  # terminal log message


# ─── Main Window ──────────────────────────────────────────────────────────────

class KnowledgeArticleConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Knowledge Article Converter v{__version__}")
        self.setMinimumSize(800, 500)

        # Load settings
        self.config_file = Path.home() / ".config" / "ka-converter" / "settings_v2.json"
        self.load_settings()

        # State
        self.selected_files = []
        self.is_converting = False
        self.conversion_results = {'success': 0, 'failed': 0, 'skipped': 0}
        self.last_output_folder = None
        self.output_files = []
        self.click_actions = []

        # Auto-watch state
        self.watching = False
        self.observer = None

        # Spinner for busy indicator
        self._spinner_frames = ['🕐', '🕑', '🕒', '🕓', '🕔', '🕕', '🕖', '🕗', '🕘', '🕙', '🕚', '🕛']
        self._spinner_index = 0
        self._spinner_timer = QTimer(self)
        self._spinner_timer.timeout.connect(self._tick_spinner)

        # Signals for thread-safe updates
        self.signals = ConvertSignals()
        self.signals.add_source.connect(self._on_add_source)
        self.signals.add_output.connect(self._on_add_output)
        self.signals.set_status.connect(self._on_set_status)
        self.signals.progress.connect(self._on_progress)
        self.signals.complete.connect(self._on_complete)
        self.signals.register_click.connect(self._on_register_click)
        self.signals.spinner_stop.connect(self._stop_spinner)
        self.signals.auto_convert_file.connect(self._auto_convert)
        self.signals.log.connect(self._on_log)

        # Build UI
        self._build_menu()
        self._build_central()
        self._build_statusbar()

        # Apply theme
        self._apply_theme()

        # Restore geometry
        geo = self.settings.get("window_geometry")
        if isinstance(geo, dict):
            self.resize(geo.get("width", 900), geo.get("height", 600))
        else:
            self.resize(900, 600)

        pos = self.settings.get("window_position")
        if isinstance(pos, list) and len(pos) == 2:
            self.move(pos[0], pos[1])

        # Start auto-watch if enabled
        if self.settings.get('auto_watch', False):
            self.start_watching()

    # ── Settings ──────────────────────────────────────────────────────────

    def load_settings(self):
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        defaults = {
            "last_file_directory": str(Path.home() / "Downloads"),
            "window_geometry": None,
            "window_position": None,
            "overwrite_existing": True,
            "strip_self_service_flow": True,
            "strip_ai_usage": True,
            "naming_mode": "original",
            "auto_watch": False,
            "watch_folder": str(Path.home() / "Downloads"),
            "theme": "system",
        }
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r') as f:
                    self.settings = {**defaults, **json.load(f)}
            else:
                self.settings = defaults
        except Exception:
            self.settings = defaults

    def save_settings(self):
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except Exception as e:
            print(f"Could not save settings: {e}")

    # ── Menu Bar ──────────────────────────────────────────────────────────

    def _build_menu(self):
        menubar = self.menuBar()

        # ── File ──
        file_menu = menubar.addMenu("File")

        open_action = QAction("Open...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.select_files)
        file_menu.addAction(open_action)

        convert_action = QAction("Convert", self)
        convert_action.setShortcut("Ctrl+R")
        convert_action.triggered.connect(self.start_conversion)
        file_menu.addAction(convert_action)

        clear_action = QAction("Clear", self)
        clear_action.triggered.connect(self._clear_lists)
        file_menu.addAction(clear_action)

        file_menu.addSeparator()

        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # ── Settings ──
        settings_menu = menubar.addMenu("Settings")

        # File Naming options (flat — no submenu)
        naming_label = QAction("File Naming:", self)
        naming_label.setEnabled(False)
        settings_menu.addAction(naming_label)

        self.naming_group = QActionGroup(self)

        naming_options = [
            ("Keep original filename", "original"),
            ("Article # - Title", "full_title"),
            ("Revised Article # - Title", "revised_title"),
        ]
        for label, mode in naming_options:
            action = QAction(label, self, checkable=True)
            action.setChecked(self.settings.get('naming_mode', 'original') == mode)
            action.triggered.connect(lambda checked, m=mode: self._set_naming_mode(m))
            self.naming_group.addAction(action)
            settings_menu.addAction(action)

        settings_menu.addSeparator()

        # Processing toggles
        self.overwrite_action = QAction("Overwrite existing files", self, checkable=True)
        self.overwrite_action.setChecked(self.settings.get('overwrite_existing', True))
        self.overwrite_action.triggered.connect(
            lambda c: self._toggle_setting('overwrite_existing', c))
        settings_menu.addAction(self.overwrite_action)

        self.strip_ssf_action = QAction("Remove Self Service Flow", self, checkable=True)
        self.strip_ssf_action.setChecked(self.settings.get('strip_self_service_flow', True))
        self.strip_ssf_action.triggered.connect(
            lambda c: self._toggle_setting('strip_self_service_flow', c))
        settings_menu.addAction(self.strip_ssf_action)

        self.strip_ai_action = QAction("Remove AI Usage", self, checkable=True)
        self.strip_ai_action.setChecked(self.settings.get('strip_ai_usage', True))
        self.strip_ai_action.triggered.connect(
            lambda c: self._toggle_setting('strip_ai_usage', c))
        settings_menu.addAction(self.strip_ai_action)

        # ── Auto-Convert ──
        auto_menu = menubar.addMenu("Auto-Convert")

        self.auto_watch_action = QAction("Enable Auto-Convert", self, checkable=True)
        self.auto_watch_action.setChecked(self.settings.get('auto_watch', False))
        self.auto_watch_action.triggered.connect(self.toggle_auto_watch)
        auto_menu.addAction(self.auto_watch_action)

        watch_folder_action = QAction("Watch Folder...", self)
        watch_folder_action.triggered.connect(self.select_watch_folder)
        auto_menu.addAction(watch_folder_action)

        # ── View ──
        view_menu = menubar.addMenu("View")

        # Theme submenu
        theme_menu = view_menu.addMenu("Theme")
        self.theme_group = QActionGroup(self)

        current_theme = self.settings.get('theme', 'system')
        for label, value in [("System", "system"), ("Light", "light"), ("Dark", "dark")]:
            action = QAction(label, self, checkable=True)
            action.setChecked(current_theme == value)
            action.triggered.connect(lambda checked, v=value: self._set_theme(v))
            self.theme_group.addAction(action)
            theme_menu.addAction(action)

        # ── Help ──
        help_menu = menubar.addMenu("Help")

        getting_started = QAction("Getting Started", self)
        getting_started.triggered.connect(self._show_getting_started)
        help_menu.addAction(getting_started)

        user_guide = QAction("User Guide", self)
        user_guide.triggered.connect(self._show_user_guide)
        help_menu.addAction(user_guide)

        license_action = QAction("License && Third-Party Notices", self)
        license_action.triggered.connect(self._show_license)
        help_menu.addAction(license_action)

        help_menu.addSeparator()

        about_action = QAction("About Knowledge Article Converter", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    # ── Central Widget ────────────────────────────────────────────────────

    def _build_central(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 8, 12, 4)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        layout.addWidget(self.progress_bar)

        # Splitter for source/output columns
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # Source Files column
        source_widget = QWidget()
        source_layout = QVBoxLayout(source_widget)
        source_layout.setContentsMargins(0, 0, 0, 0)
        source_header = QLabel("Source Files")
        source_header.setObjectName("sectionHeader")
        source_layout.addWidget(source_header)
        self.source_list = QListWidget()
        self.source_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        source_layout.addWidget(self.source_list)

        # Output column
        output_widget = QWidget()
        output_layout = QVBoxLayout(output_widget)
        output_layout.setContentsMargins(0, 0, 0, 0)

        output_header_row = QHBoxLayout()
        output_header_row.setContentsMargins(0, 0, 0, 0)
        output_header = QLabel("Output")
        output_header.setObjectName("sectionHeader")
        output_header_row.addWidget(output_header)
        output_header_row.addStretch()
        self._spinner_label = QLabel("")
        self._spinner_label.setObjectName("sectionHeader")
        output_header_row.addWidget(self._spinner_label)
        output_layout.addLayout(output_header_row)
        self.output_list = QListWidget()
        self.output_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        self.output_list.setCursor(Qt.CursorShape.PointingHandCursor)
        self.output_list.itemClicked.connect(self._on_output_click)
        output_layout.addWidget(self.output_list)

        self.splitter.addWidget(source_widget)
        self.splitter.addWidget(output_widget)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 3)

        layout.addWidget(self.splitter)

        # Terminal log
        self.terminal = CRTTerminal()
        self.terminal.setFixedHeight(160)
        layout.addWidget(self.terminal)

    # ── Status Bar ────────────────────────────────────────────────────────

    def _build_statusbar(self):
        self.statusBar().showMessage("Ready")

    # ── Theme ─────────────────────────────────────────────────────────────

    def _apply_theme(self):
        theme = self.settings.get('theme', 'system')
        if theme == 'system':
            use_dark = detect_system_dark_mode()
        elif theme == 'dark':
            use_dark = True
        else:
            use_dark = False

        self.setStyleSheet(DARK_THEME if use_dark else LIGHT_THEME)
        self._current_dark = use_dark

    def _set_theme(self, theme):
        self.settings['theme'] = theme
        self.save_settings()
        self._apply_theme()

    def _get_color(self, role):
        """Get a color for the current theme."""
        if self._current_dark:
            colors = {
                'success': '#30d158', 'error': '#ff453a', 'warning': '#ff9f0a',
                'secondary': '#98989d',
            }
        else:
            colors = {
                'success': '#34c759', 'error': '#ff3b30', 'warning': '#ff9500',
                'secondary': '#6e6e73',
            }
        return colors.get(role, '#888888')

    # ── Menu Actions ──────────────────────────────────────────────────────

    def _set_naming_mode(self, mode):
        self.settings['naming_mode'] = mode
        self.save_settings()

    def _toggle_setting(self, key, checked):
        self.settings[key] = checked
        self.save_settings()

    def _clear_lists(self):
        """Clear file selection and output lists."""
        self.selected_files = []
        self.source_list.clear()
        self.output_list.clear()
        self.click_actions = []
        self.output_files = []
        self.progress_bar.setValue(0)
        self.terminal.clear()
        if self.watching:
            watch_path = self.settings.get('watch_folder', '~/Downloads')
            display = watch_path.replace(str(Path.home()), '~')
            self.statusBar().showMessage(f"Watching {display} for new exports...")
        else:
            self.statusBar().showMessage("Ready")

    # ── Help ──────────────────────────────────────────────────────────────

    def _show_getting_started(self):
        QMessageBox.information(self, "Getting Started",
            f"Knowledge Article Converter v{__version__}\n\n"
            "Quick Start:\n"
            "1. Enable Auto-Convert from the Auto-Convert menu\n"
            "2. Set your Watch Folder (defaults to ~/Downloads)\n"
            "3. Export articles from the knowledge base — they\n"
            "   will be converted to DOCX automatically\n\n"
            "Manual Mode:\n"
            "Use File → Open to select HTML files, then\n"
            "File → Convert to process them.\n\n"
            "Settings are saved automatically between sessions.")

    def _show_user_guide(self):
        QMessageBox.information(self, "User Guide",
            "File Naming Modes:\n"
            "• Keep original — uses the exported filename\n"
            "• Article # - Title — names using the article title\n"
            "• Revised Article # - Title — prepends 'Revised'\n\n"
            "Processing Options (Settings menu):\n"
            "• Remove Self Service Flow — strips the SSF section\n"
            "• Remove AI Usage — strips the AI Usage section\n"
            "• Overwrite existing — replaces files without prompting\n\n"
            "Output Column:\n"
            "• Click a filename to open the converted document\n"
            "• Click a folder path to open it in your file manager\n\n"
            "Theme:\n"
            "View → Theme lets you switch between System,\n"
            "Light, and Dark themes.")

    def _show_about(self):
        QMessageBox.about(self, "About Knowledge Article Converter",
            f"Knowledge Article Converter\n"
            f"Version {__version__}\n\n"
            f"Author: {__author__}\n\n"
            "Converts HTML knowledge base article exports\n"
            "to professionally formatted DOCX files.\n\n"
            "This software is distributed under the GNU General\n"
            "Public License v3. See Help \u2192 License & Third-Party\n"
            "Notices for full licensing information.")

    def _show_license(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("License & Third-Party Notices")
        dialog.setMinimumSize(620, 500)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(10)

        header = QLabel("License & Third-Party Notices")
        header.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(header)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setStyleSheet("font-family: monospace; font-size: 11px;")
        text.setPlainText(
            "Knowledge Article Converter\n"
            "======================================================\n\n"
            f"Copyright (C) 2025 {__author__}\n\n"
            "This program is free software: you can redistribute it\n"
            "and/or modify it under the terms of the GNU General\n"
            "Public License as published by the Free Software\n"
            "Foundation, either version 3 of the License, or (at\n"
            "your option) any later version.\n\n"
            "This program is distributed in the hope that it will\n"
            "be useful, but WITHOUT ANY WARRANTY; without even the\n"
            "implied warranty of MERCHANTABILITY or FITNESS FOR A\n"
            "PARTICULAR PURPOSE. See the GNU General Public License\n"
            "for more details.\n\n"
            "Full license text:\n"
            "https://www.gnu.org/licenses/gpl-3.0.txt\n\n\n"
            "Third-Party Components\n"
            "======================================================\n\n"
            "PyQt6\n"
            "------------------------------------------------------\n"
            "Copyright (C) Riverbank Computing Limited\n"
            "License: GNU General Public License v3\n"
            "https://riverbankcomputing.com/software/pyqt/\n\n"
            "pandoc\n"
            "------------------------------------------------------\n"
            "Copyright (C) 2006-2024 John MacFarlane\n"
            "License: GNU General Public License v2 or later\n"
            "https://pandoc.org/\n\n"
            "Note: this application uses a modified build of pandoc.\n"
            "The source modifications are included in the\n"
            "repository under pandoc-patch/.\n\n"
            "watchdog\n"
            "------------------------------------------------------\n"
            "Copyright (C) 2011 Yesudeep Mangalapilly and contributors\n"
            "License: Apache License 2.0\n"
            "https://github.com/gorakhargosh/watchdog\n\n"
            "Python\n"
            "------------------------------------------------------\n"
            "Copyright (C) 2001-2024 Python Software Foundation\n"
            "License: Python Software Foundation License (PSF)\n"
            "https://www.python.org/\n"
        )
        layout.addWidget(text)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        dialog.exec()

    # ── File Selection ────────────────────────────────────────────────────

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select HTML Files",
            self.settings.get('last_file_directory', str(Path.home() / "Downloads")),
            "HTML Files (*.html);;All Files (*)")
        if files:
            self.selected_files = files
            self.settings['last_file_directory'] = str(Path(files[0]).parent)
            self.save_settings()

            self.source_list.clear()
            self.output_list.clear()
            self.click_actions = []
            self.output_files = []

            for f in files:
                self.source_list.addItem(os.path.basename(f))

            count = len(files)
            self.statusBar().showMessage(
                f"{count} file{'s' if count != 1 else ''} selected — File → Convert to process")

    # ── Conversion ────────────────────────────────────────────────────────

    def determine_output_filename(self, html_file):
        mode = self.settings.get('naming_mode', 'original')
        original_base = os.path.splitext(os.path.basename(html_file))[0]

        if mode == 'original':
            return original_base + '.docx'

        article_num, full_title, clean_title = extract_article_info(html_file)
        if article_num and clean_title:
            if mode == 'full_title':
                name = f"Article {full_title}"
            elif mode == 'revised_title':
                name = f"Revised Article {full_title}"
            else:
                name = original_base
            return sanitize_filename(name) + '.docx'

        return original_base + '.docx'

    def start_conversion(self):
        if not self.selected_files:
            QMessageBox.warning(self, "No Files", "Use Open to select HTML files first.")
            return
        if self.is_converting:
            return

        self.is_converting = True
        self.output_list.clear()
        self.click_actions = []
        self.output_files = []
        self.progress_bar.setValue(0)
        self.conversion_results = {'success': 0, 'failed': 0, 'skipped': 0}
        self.terminal.clear()
        self._start_spinner()

        Thread(target=self._convert_files, daemon=True).start()

    def _convert_files(self):
        total = len(self.selected_files)

        for idx, html_file in enumerate(self.selected_files, 1):
            self.signals.set_status.emit(
                f"Converting {idx}/{total}: {os.path.basename(html_file)}")

            out_name = self.determine_output_filename(html_file)

            output_folder = str(Path(html_file).parent)

            output_file = os.path.join(output_folder, out_name)
            self.last_output_folder = output_folder

            success = self.convert_single_file(html_file, output_file)

            if success:
                self.conversion_results['success'] += 1
                self.signals.add_output.emit(f"✓ {out_name}", self._get_color('success'))
                self.signals.register_click.emit('file', output_file)
            else:
                self.conversion_results['failed'] += 1
                self.signals.add_output.emit(f"✗ {out_name}", self._get_color('error'))
                self.signals.register_click.emit('none', '')

            progress = int((idx / total) * 100)
            self.signals.progress.emit(progress)

        self.signals.complete.emit()

    def convert_single_file(self, input_file, output_file):
        """Convert single HTML file to DOCX."""
        basename = os.path.basename(input_file)
        out_basename = os.path.basename(output_file)
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            actual_input = input_file
            tmp_input = None

            # ── Phase 1: Read and analyze input ──
            self._log_cmd(f"stat {basename}")
            input_size = os.path.getsize(input_file)
            import stat as stat_mod
            fstat = os.stat(input_file)
            perms = stat_mod.filemode(fstat.st_mode)
            mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(fstat.st_mtime))
            self._log_info(f"{perms} {input_size:,} bytes modified {mtime}")

            self._log_cmd(f"sha256sum {basename}")
            input_hash = file_hash(input_file)
            self._log_info(f"{input_hash}  {basename}")

            self._log_cmd(f"cat {basename} | wc -l")
            with open(input_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            line_count = len(html_content.splitlines())
            self._log_info(f"{len(html_content):,} chars | {line_count:,} lines")

            # HTML structure analysis
            self._log_cmd(f"htmlparse --analyze {basename}")
            analysis = analyze_html(html_content)
            enc = analysis['declared_encoding'] or 'not declared'
            self._log_info(f"encoding: {enc}")
            self._log_info(f"dom: {analysis['total_tags']} nodes, {analysis['unique_tags']} unique tags")

            parts = []
            for key, label in [('headings', 'h1-h6'), ('paragraphs', 'p'),
                               ('divs', 'div'), ('spans', 'span')]:
                if analysis[key]:
                    parts.append(f"{analysis[key]} {label}")
            if parts:
                self._log_info(f"  layout: {' | '.join(parts)}")

            parts = []
            if analysis['tables']:
                parts.append(f"{analysis['tables']} table ({analysis['rows']} tr)")
            if analysis['lists']:
                parts.append(f"{analysis['lists']} list ({analysis['list_items']} li)")
            if parts:
                self._log_info(f"  data:   {' | '.join(parts)}")

            parts = []
            if analysis['images']:
                parts.append(f"{analysis['images']} img")
            if analysis['links']:
                parts.append(f"{analysis['links']} a")
            if analysis['bold']:
                parts.append(f"{analysis['bold']} bold")
            if analysis['italic']:
                parts.append(f"{analysis['italic']} italic")
            if parts:
                self._log_info(f"  inline: {' | '.join(parts)}")

            # ── Phase 2: Strip sections ──
            strip_ssf = self.settings.get('strip_self_service_flow', True)
            strip_ai = self.settings.get('strip_ai_usage', True)

            cleaned = html_content
            if strip_ssf or strip_ai:
                if strip_ssf:
                    self._log_cmd("sed -n '/Self Service Flow/p'")
                    before = len(cleaned)
                    cleaned = strip_self_service_flow(cleaned)
                    delta = before - len(cleaned)
                    if delta > 0:
                        self._log_info(f"matched — removed {delta:,} bytes ({delta * 100 // before}%)")
                    else:
                        self._log_info("no match (0 bytes removed)")
                if strip_ai:
                    self._log_cmd("sed -n '/AI Usage/p'")
                    before = len(cleaned)
                    cleaned = strip_ai_usage(cleaned)
                    delta = before - len(cleaned)
                    if delta > 0:
                        self._log_info(f"matched — removed {delta:,} bytes ({delta * 100 // before}%)")
                    else:
                        self._log_info("no match (0 bytes removed)")

                if cleaned != html_content:
                    import tempfile
                    tmp = tempfile.NamedTemporaryFile(
                        mode='w', suffix='.html', delete=False,
                        encoding='utf-8')
                    tmp.write(cleaned)
                    tmp.close()
                    actual_input = tmp.name
                    tmp_input = tmp.name
                    reduction = len(html_content) - len(cleaned)
                    self._log_cmd(f"tee /tmp/{os.path.basename(tmp.name)}")
                    self._log_info(
                        f"{len(cleaned):,} bytes written (stripped {reduction:,} from original)")

            # ── Phase 3: Resolve resources ──
            if REFERENCE_DOCX_B64 and len(REFERENCE_DOCX_B64) > 100:
                import tempfile
                self._log_cmd("base64 --decode < reference.docx.b64")
                ref_decoded = base64.b64decode(REFERENCE_DOCX_B64)
                self._log_info(f"decoded {len(ref_decoded):,} bytes")
                with tempfile.NamedTemporaryFile(
                        mode='w', suffix='.lua', delete=False) as lua_file:
                    lua_file.write(LUA_FILTER)
                    lua_path = lua_file.name
                self._log_info(f"lua filter: /tmp/{os.path.basename(lua_path)} ({len(LUA_FILTER):,} bytes)")
                with tempfile.NamedTemporaryFile(
                        mode='wb', suffix='.docx', delete=False) as ref_file:
                    ref_file.write(ref_decoded)
                    ref_path = ref_file.name
                self._log_info(f"reference: /tmp/{os.path.basename(ref_path)} ({len(ref_decoded):,} bytes)")
            else:
                self._log_cmd("ls -la filter.lua reference.docx")
                lua_path = os.path.join(script_dir, 'filter.lua')
                ref_path = os.path.join(script_dir, 'reference.docx')
                if not os.path.exists(lua_path):
                    self._log_err(f"ENOENT: {lua_path}")
                    return False
                if not os.path.exists(ref_path):
                    self._log_err(f"ENOENT: {ref_path}")
                    return False
                lua_size = os.path.getsize(lua_path)
                ref_size = os.path.getsize(ref_path)
                self._log_info(f"-rw-r--r-- {lua_size:>8,}  filter.lua")
                self._log_info(f"-rw-r--r-- {ref_size:>8,}  reference.docx")

            # ── Phase 4: pandoc version ──
            self._log_cmd("pandoc --version | head -1")
            try:
                ver_result = subprocess.run(
                    ['pandoc', '--version'], capture_output=True, text=True, timeout=5)
                ver_line = ver_result.stdout.splitlines()[0] if ver_result.stdout else 'unknown'
                self._log_info(ver_line)
            except Exception:
                self._log_warn("could not determine pandoc version")

            # ── Phase 5: Run pandoc with --verbose ──
            pandoc_cmd = [
                'pandoc', actual_input,
                '--verbose',
                '--lua-filter', lua_path,
                '--reference-doc', ref_path,
                '-o', output_file
            ]

            if os.environ.get('PANDOC_DATADIR'):
                pandoc_cmd.insert(2, '--data-dir')
                pandoc_cmd.insert(3, os.environ['PANDOC_DATADIR'])

            display_cmd = ' '.join([
                'pandoc', os.path.basename(actual_input),
                '--verbose',
                '--lua-filter', os.path.basename(lua_path),
                '--reference-doc', os.path.basename(ref_path),
                '-o', out_basename
            ])
            self._log_cmd(display_cmd)

            t_start = time.monotonic()
            result = subprocess.run(pandoc_cmd, capture_output=True, text=True)
            elapsed = time.monotonic() - t_start

            # Stream pandoc's verbose/warning output
            if result.stderr.strip():
                for line in result.stderr.strip().splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    if result.returncode != 0:
                        self._log_err(line)
                    elif line.startswith('[WARNING]'):
                        self._log_warn(line)
                    else:
                        self._log_info(line)

            # Cleanup
            if REFERENCE_DOCX_B64 and len(REFERENCE_DOCX_B64) > 100:
                os.unlink(lua_path)
                os.unlink(ref_path)
            if tmp_input:
                os.unlink(tmp_input)

            if result.returncode != 0:
                self._log_err(f"process exited with code {result.returncode} after {elapsed:.2f}s")
                return False

            # ── Phase 6: Output analysis ──
            self._log_info(f"[exit 0] completed in {elapsed:.2f}s")

            if os.path.exists(output_file):
                out_size = os.path.getsize(output_file)
                ratio = out_size / input_size if input_size > 0 else 0
                self._log_cmd(f"sha256sum {out_basename}")
                out_hash = file_hash(output_file)
                self._log_info(f"{out_hash}  {out_basename}")
                self._log_info(f"{out_size:,} bytes ({ratio:.1f}x input)")
                self._log_ok(f"✓ {out_basename}")
            else:
                self._log_ok(f"✓ {out_basename}")

            return True

        except Exception as e:
            self._log_err(f"fatal: {e}")
            import traceback
            traceback.print_exc()
            return False

    # ── Auto-Watch ────────────────────────────────────────────────────────

    def select_watch_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Watch Folder",
            self.settings.get('watch_folder', str(Path.home() / "Downloads")))
        if folder:
            self.settings['watch_folder'] = folder
            self.save_settings()
            if self.watching:
                self.stop_watching()
                self.start_watching()

    def toggle_auto_watch(self, checked):
        self.settings['auto_watch'] = checked
        self.save_settings()
        if checked:
            self.start_watching()
        else:
            self.stop_watching()

    def start_watching(self):
        watch_path = self.settings.get('watch_folder', str(Path.home() / "Downloads"))
        if not os.path.isdir(watch_path):
            QMessageBox.warning(self, "Invalid Folder",
                                f"Watch folder does not exist:\n{watch_path}")
            self.auto_watch_action.setChecked(False)
            self.settings['auto_watch'] = False
            self.save_settings()
            return

        self.watching = True
        handler = _ArticleFileHandler(self)
        self.observer = Observer()
        self.observer.schedule(handler, watch_path, recursive=True)
        self.observer.start()

        display = watch_path.replace(str(Path.home()), '~')
        self.statusBar().showMessage(f"Watching {display} for new exports...")

    def stop_watching(self):
        self.watching = False
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
            self.observer = None
        self.statusBar().showMessage("Ready")

    def _auto_convert(self, html_file):
        if self.is_converting:
            return

        basename = os.path.basename(html_file)
        self.statusBar().showMessage(f"Auto-converting: {basename}")
        self._start_spinner()

        # Add to source list
        self.source_list.addItem(basename)
        self.source_list.addItem("")  # Blank line to align

        out_name = self.determine_output_filename(html_file)
        output_dir = os.path.dirname(html_file)
        output_file = os.path.join(output_dir, out_name)
        self.last_output_folder = output_dir

        # Shorten path for display
        rel_dir = output_dir.replace(str(Path.home()), '~')
        parts = rel_dir.split(os.sep)
        if len(parts) > 4:
            short_dir = os.sep.join(['...'] + parts[-2:])
        else:
            short_dir = rel_dir

        if os.path.exists(output_file) and not self.settings.get('overwrite_existing', True):
            item = QListWidgetItem(f"⊘ {out_name} (exists)")
            item.setForeground(QColor(self._get_color('warning')))
            self.output_list.addItem(item)
            self.click_actions.append(None)

            item2 = QListWidgetItem(f"    → {short_dir}")
            item2.setForeground(QColor(self._get_color('secondary')))
            self.output_list.addItem(item2)
            self.click_actions.append(None)

            watch_path = self.settings.get('watch_folder', '~/Downloads')
            display = watch_path.replace(str(Path.home()), '~')
            self.statusBar().showMessage(f"Watching {display} for new exports...")
            return

        def do_convert():
            success = self.convert_single_file(html_file, output_file)

            if success:
                self.signals.add_output.emit(
                    f"✓ {out_name}", self._get_color('success'))
                self.signals.register_click.emit('file', output_file)
                self.signals.add_output.emit(
                    f"    → {short_dir}", self._get_color('secondary'))
                self.signals.register_click.emit('folder', output_dir)
            else:
                self.signals.add_output.emit(
                    f"✗ {out_name}", self._get_color('error'))
                self.signals.register_click.emit('none', '')
                self.signals.add_output.emit("", self._get_color('secondary'))
                self.signals.register_click.emit('none', '')

            self.output_files.append(output_file if success else None)

            watch_path = self.settings.get('watch_folder', '~/Downloads')
            display = watch_path.replace(str(Path.home()), '~')
            self.signals.set_status.emit(f"Watching {display} for new exports...")
            self.signals.spinner_stop.emit()

        Thread(target=do_convert, daemon=True).start()

    # ── Signal Handlers ───────────────────────────────────────────────────

    def _on_add_source(self, text):
        self.source_list.addItem(text)

    def _on_add_output(self, text, color):
        item = QListWidgetItem(text)
        item.setForeground(QColor(color))
        self.output_list.addItem(item)
        self.output_list.scrollToBottom()

    def _on_set_status(self, text):
        self.statusBar().showMessage(text)

    def _on_log(self, text):
        self.terminal.append(text)
        scrollbar = self.terminal.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _log_cmd(self, text):
        """Log a command line (bright green, $ prefix)."""
        self.signals.log.emit(
            f'<span style="color:#33ff33;font-weight:bold;">$ {text}</span>')

    def _log_info(self, text):
        """Log an info/detail line (dim green)."""
        self.signals.log.emit(
            f'<span style="color:#1a9a1a;">&nbsp;&nbsp;{text}</span>')

    def _log_ok(self, text):
        """Log a success line (bright green)."""
        self.signals.log.emit(
            f'<span style="color:#33ff33;">&nbsp;&nbsp;{text}</span>')

    def _log_err(self, text):
        """Log an error line (amber)."""
        self.signals.log.emit(
            f'<span style="color:#ff6600;">&nbsp;&nbsp;{text}</span>')

    def _log_warn(self, text):
        """Log a warning line (yellow)."""
        self.signals.log.emit(
            f'<span style="color:#cccc00;">&nbsp;&nbsp;{text}</span>')

    # ── Spinner ──────────────────────────────────────────────────────────

    def _start_spinner(self):
        self._spinner_index = 0
        self._spinner_label.setText(self._spinner_frames[0])
        self._spinner_timer.start(150)

    def _stop_spinner(self):
        self._spinner_timer.stop()
        self._spinner_label.setText("")

    def _tick_spinner(self):
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_frames)
        self._spinner_label.setText(self._spinner_frames[self._spinner_index])

    def _on_progress(self, value):
        self.progress_bar.setValue(value)

    def _on_register_click(self, action_type, path):
        if action_type == 'none':
            self.click_actions.append(None)
        else:
            self.click_actions.append((action_type, path))

    def _on_complete(self):
        self.is_converting = False
        self._stop_spinner()
        s = self.conversion_results['success']
        f = self.conversion_results['failed']
        total = s + f

        if f == 0:
            self.statusBar().showMessage(f"Successfully converted {s} file(s)")
        else:
            self.statusBar().showMessage(f"Converted {s} of {total} — {f} failed")

    def _on_output_click(self, item):
        index = self.output_list.row(item)
        if index < len(self.click_actions) and self.click_actions[index]:
            action_type, path = self.click_actions[index]
            if os.path.exists(path):
                try:
                    subprocess.run(['xdg-open', path])
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Could not open:\n{e}")
            else:
                QMessageBox.warning(self, "Not Found", f"Path does not exist:\n{path}")

    # ── Window Close ──────────────────────────────────────────────────────

    def closeEvent(self, event):
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)

        self.settings['window_geometry'] = {
            'width': self.width(), 'height': self.height()
        }
        self.settings['window_position'] = [self.x(), self.y()]
        self.save_settings()
        event.accept()


# ─── Watchdog Handler ─────────────────────────────────────────────────────────

class _ArticleFileHandler(FileSystemEventHandler):
    """Watchdog handler that triggers auto-convert for new Article *.html files."""

    def __init__(self, app):
        super().__init__()
        self.app = app

    def on_created(self, event):
        if event.is_directory:
            return
        filename = os.path.basename(event.src_path)
        if fnmatch.fnmatch(filename, 'Article *.html'):
            time.sleep(0.5)
            self.app.signals.auto_convert_file.emit(event.src_path)


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    import sys
    app = QApplication(sys.argv)
    window = KnowledgeArticleConverter()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
