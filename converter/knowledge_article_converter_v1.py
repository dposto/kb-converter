#!/usr/bin/env python3
"""
Knowledge Article Converter
Converts HTML knowledge base articles to formatted DOCX files
"""

__version__ = "1.3"
__author__ = "David Posto"

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import json
import re
import time
import fnmatch
from pathlib import Path
from threading import Thread
import base64
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Embedded reference document (base64)
REFERENCE_DOCX_B64 = ""  # Will be filled during build

# Embedded Lua filter
LUA_FILTER = """-- FILTER CONTENT WILL BE EMBEDDED HERE --"""


def detect_system_dark_mode():
    """Detect if the desktop environment is using a dark theme.

    Checks in order:
    1. XDG Desktop Portal (works on COSMIC, GNOME 42+, KDE, any freedesktop DE)
    2. gsettings color-scheme (GNOME)
    3. gsettings GTK theme name (GNOME/GTK-based DEs)
    4. xfconf (XFCE)
    """
    # Method 1: XDG Desktop Portal - most universal, covers COSMIC DE
    # Portal returns: 0 = no preference, 1 = prefer dark, 2 = prefer light
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
        # If we got a valid response of 0 or 2, that's explicitly not dark
        if result.returncode == 0 and ('uint32 0' in result.stdout or
                                        'uint32 2' in result.stdout):
            return False
    except Exception:
        pass

    # Method 2: gsettings color-scheme (GNOME 42+)
    try:
        result = subprocess.run(
            ['gsettings', 'get', 'org.gnome.desktop.interface', 'color-scheme'],
            capture_output=True, text=True, timeout=2
        )
        if 'prefer-dark' in result.stdout:
            return True
    except Exception:
        pass

    # Method 3: GTK theme name for common dark indicators
    try:
        result = subprocess.run(
            ['gsettings', 'get', 'org.gnome.desktop.interface', 'gtk-theme'],
            capture_output=True, text=True, timeout=2
        )
        theme = result.stdout.strip().strip("'\"").lower()
        if any(dark in theme for dark in ['dark', 'noir', 'night', 'gruvbox', 'dracula']):
            return True
    except Exception:
        pass

    # Method 4: xfconf for XFCE
    try:
        result = subprocess.run(
            ['xfconf-query', '-c', 'xsettings', '-p', '/Net/ThemeName'],
            capture_output=True, text=True, timeout=2
        )
        theme = result.stdout.strip().lower()
        if any(dark in theme for dark in ['dark', 'noir', 'night']):
            return True
    except Exception:
        pass

    return False


def extract_article_info(html_path):
    """Extract article number and title from the HTML <h1> tag.

    Handles optional prefixes like [Draft] before the article number.
    Returns (article_number, full_title, clean_title) or (None, None, None).
    Example: '[Draft] 12345 - Sample Article Title'
        -> ('12345', '12345 - Sample Article Title', 'Sample Article Title')
    """
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            content = f.read()
        # Match optional prefix in brackets (e.g. [Draft]) before the article number
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
    """Remove the Self Service Flow section from HTML content.

    The section starts with <h2>Self Service Flow</h2> followed by a <p> block
    containing the flow definition ending with nested closing braces.
    """
    pattern = r'<h2>\s*Self Service Flow\s*</h2>\s*<p>.*?</p>'
    return re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)


def strip_ai_usage(html_content):
    """Remove the AI Usage section from HTML content.

    Typically contains just '(No value)' and is not needed in output.
    """
    pattern = r'<h2>\s*AI Usage\s*</h2>\s*<p>.*?</p>'
    return re.sub(pattern, '', html_content, flags=re.DOTALL | re.IGNORECASE)


# ─── Color Palettes ───────────────────────────────────────────────────────────

LIGHT_COLORS = {
    'bg':           '#f2f2f7',
    'fg':           '#1c1c1e',
    'fg_secondary': '#6e6e73',
    'accent':       '#0061AC',
    'accent_fg':    '#ffffff',
    'accent_hover': '#004d8a',
    'success':      '#34c759',
    'error':        '#ff3b30',
    'warning':      '#ff9500',
    'card_bg':      '#ffffff',
    'card_border':  '#d1d1d6',
    'input_bg':     '#ffffff',
    'input_border': '#c7c7cc',
    'header_bg':    '#0061AC',
    'header_fg':    '#ffffff',
    'list_bg':      '#ffffff',
    'separator':    '#d1d1d6',
    'radio_sel':    '#ffffff',
    'check_sel':    '#ffffff',
    'disabled_bg':  '#e5e5ea',
    'disabled_fg':  '#aeaeb2',
    'btn_bg':       '#e5e5ea',
    'btn_fg':       '#1c1c1e',
    'btn_hover':    '#d1d1d6',
}

DARK_COLORS = {
    'bg':           '#1c1c1e',
    'fg':           '#f2f2f7',
    'fg_secondary': '#98989d',
    'accent':       '#3a9bdc',
    'accent_fg':    '#ffffff',
    'accent_hover': '#2d7ab8',
    'success':      '#30d158',
    'error':        '#ff453a',
    'warning':      '#ff9f0a',
    'card_bg':      '#2c2c2e',
    'card_border':  '#3a3a3c',
    'input_bg':     '#3a3a3c',
    'input_border': '#48484a',
    'header_bg':    '#2c2c2e',
    'header_fg':    '#f2f2f7',
    'list_bg':      '#2c2c2e',
    'separator':    '#48484a',
    'radio_sel':    '#3a3a3c',
    'check_sel':    '#3a3a3c',
    'disabled_bg':  '#2c2c2e',
    'disabled_fg':  '#636366',
    'btn_bg':       '#3a3a3c',
    'btn_fg':       '#f2f2f7',
    'btn_hover':    '#48484a',
}


class KnowledgeArticleConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Knowledge Article Converter")

        # Load settings
        self.config_file = Path.home() / ".config" / "ka-converter" / "settings.json"
        self.load_settings()

        # Always detect system theme
        self.settings['dark_mode'] = detect_system_dark_mode()

        # Restore window
        if self.settings.get("window_geometry"):
            self.root.geometry(self.settings["window_geometry"])
        else:
            self.root.geometry("900x850")

        self.root.minsize(800, 600)

        if self.settings.get("window_position"):
            self.root.geometry(
                f"+{self.settings['window_position'][0]}+{self.settings['window_position'][1]}")

        self.root.update_idletasks()
        if self.root.winfo_width() < 800 or self.root.winfo_height() < 600:
            self.root.geometry("900x850")

        # State
        self.selected_files = []
        self.is_converting = False
        self.conversion_results = {'success': 0, 'failed': 0, 'skipped': 0}
        self.last_output_folder = None
        self.output_files = []
        self.click_actions = []  # Maps each status listbox line to ('file', path) or ('folder', path) or None

        # Auto-watch state
        self.watching = False
        self.observer = None

        # Build UI
        self.colors = DARK_COLORS if self.settings.get('dark_mode', False) else LIGHT_COLORS
        self.root.configure(bg=self.colors['bg'])
        self.setup_styles()
        self.create_widgets()

        self.root.update_idletasks()
        self.root.minsize(800, 600)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ── Settings ──────────────────────────────────────────────────────────

    def load_settings(self):
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        defaults = {
            "output_folder": str(Path.home() / "Documents"),
            "same_folder_as_source": True,
            "last_file_directory": str(Path.home() / "Downloads"),
            "window_geometry": "900x850",
            "window_position": None,
            "overwrite_existing": True,
            "dark_mode": False,
            "strip_self_service_flow": True,
            "strip_ai_usage": True,
            "naming_mode": "original",
            "auto_watch": False,
            "watch_folder": str(Path.home() / "Downloads"),
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

    # ── Styles ────────────────────────────────────────────────────────────

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        c = self.colors

        style.configure('TButton',
                        background=c['btn_bg'], foreground=c['btn_fg'],
                        borderwidth=1, relief='flat',
                        padding=(14, 7), font=('Sans', 10))
        style.map('TButton',
                  background=[('active', c['btn_hover']),
                              ('disabled', c['disabled_bg'])],
                  foreground=[('disabled', c['disabled_fg'])])

        style.configure('Accent.TButton',
                        background=c['accent'], foreground=c['accent_fg'],
                        borderwidth=0, relief='flat',
                        padding=(14, 7), font=('Sans', 10, 'bold'))
        style.map('Accent.TButton',
                  background=[('active', c['accent_hover']),
                              ('disabled', c['disabled_bg'])],
                  foreground=[('disabled', c['disabled_fg'])])

        style.configure('TProgressbar',
                        background=c['accent'], troughcolor=c['card_bg'],
                        borderwidth=0, thickness=8)

        style.configure('TRadiobutton',
                        background=c['card_bg'], foreground=c['fg'],
                        font=('Sans', 10), indicatorcolor=c['input_bg'],
                        indicatorrelief='flat')
        style.map('TRadiobutton',
                  background=[('active', c['card_bg'])],
                  indicatorcolor=[('selected', c['accent'])])

        style.configure('TCheckbutton',
                        background=c['card_bg'], foreground=c['fg'],
                        font=('Sans', 10), indicatorcolor=c['input_bg'],
                        indicatorrelief='flat')
        style.map('TCheckbutton',
                  background=[('active', c['card_bg'])],
                  indicatorcolor=[('selected', c['accent'])])

        style.configure('TCombobox',
                        background=c['input_bg'], foreground=c['fg'],
                        fieldbackground=c['input_bg'],
                        arrowcolor=c['fg'],
                        font=('Sans', 10))
        style.map('TCombobox',
                  fieldbackground=[('readonly', c['input_bg'])],
                  foreground=[('readonly', c['fg'])])

    # ── UI Construction ───────────────────────────────────────────────────

    def create_widgets(self):
        c = self.colors

        main = tk.Frame(self.root, bg=c['bg'])
        main.grid(row=0, column=0, sticky='nsew', padx=16, pady=16)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # ── Card: Files ───────────────────────────────────────────────
        file_card = self._make_card(main, row=0)

        self._card_heading(file_card, "Files")

        btn_row = tk.Frame(file_card, bg=c['card_bg'])
        btn_row.pack(fill='x', padx=12, pady=10)

        self.select_btn = tk.Button(
            btn_row, text="Select HTML Files",
            bg=c['accent'], fg=c['accent_fg'],
            activebackground=c['accent_hover'], activeforeground=c['accent_fg'],
            font=('Sans', 9), relief='flat', cursor='hand2',
            padx=10, pady=3, command=self.select_files)
        self.select_btn.pack(side='left', padx=(0, 8))

        tk.Button(
            btn_row, text="Clear",
            bg=c['btn_bg'], fg=c['btn_fg'], activebackground=c['btn_hover'],
            font=('Sans', 9), relief='flat', cursor='hand2',
            padx=10, pady=3, command=self.clear_selection).pack(side='left')

        self.file_count_label = tk.Label(
            btn_row, text="No files selected", font=('Sans', 10),
            bg=c['card_bg'], fg=c['fg_secondary'])
        self.file_count_label.pack(side='left', padx=(12, 0))

        # Convert and Open Folder buttons on the right side of the same row
        tk.Button(
            btn_row, text="Open Output Folder",
            bg=c['btn_bg'], fg=c['btn_fg'], activebackground=c['btn_hover'],
            font=('Sans', 9), relief='flat', cursor='hand2',
            padx=10, pady=3, command=self.open_output_folder).pack(
            side='right', padx=(8, 0))

        self.convert_btn = tk.Button(
            btn_row, text="Convert Files",
            bg=c['accent'], fg=c['accent_fg'],
            activebackground=c['accent_hover'], activeforeground=c['accent_fg'],
            font=('Sans', 9), relief='flat', cursor='hand2',
            padx=10, pady=3, command=self.start_conversion)
        self.convert_btn.pack(side='right')

        # ── Card: Options ─────────────────────────────────────────────
        opt_card = self._make_card(main, row=1)

        self._card_heading(opt_card, "Options")

        options_inner = tk.Frame(opt_card, bg=c['card_bg'])
        options_inner.pack(fill='x', padx=12, pady=(8, 10))

        # -- Output Location dropdown --
        loc_row = tk.Frame(options_inner, bg=c['card_bg'])
        loc_row.pack(fill='x', pady=(0, 6))

        tk.Label(loc_row, text="Output Location:", font=('Sans', 10, 'bold'),
                 bg=c['card_bg'], fg=c['fg'], width=16, anchor='w').pack(side='left', padx=(0, 8))

        self.output_mode = tk.StringVar(
            value='Same folder as source files' if self.settings['same_folder_as_source']
            else 'Custom folder')

        self.output_combo = ttk.Combobox(
            loc_row, textvariable=self.output_mode,
            values=['Same folder as source files', 'Custom folder'],
            state='readonly', font=('Sans', 10), width=28)
        self.output_combo.pack(side='left')
        self.output_combo.bind('<<ComboboxSelected>>', lambda e: self.update_output_mode())

        # Custom folder row (shown/hidden based on selection)
        self.folder_row = tk.Frame(options_inner, bg=c['card_bg'])

        self.folder_entry = tk.Entry(
            self.folder_row, font=('Sans', 10),
            bg=c['input_bg'], fg=c['fg'], insertbackground=c['fg'],
            relief='solid', borderwidth=1,
            highlightbackground=c['input_border'], highlightcolor=c['accent'],
            disabledbackground=c['disabled_bg'], disabledforeground=c['disabled_fg'])
        self.folder_entry.pack(side='left', fill='x', expand=True, ipady=3)
        self.folder_entry.insert(0, self.settings['output_folder'])

        for txt, cmd in [("Browse…", self.select_output_folder),
                         ("Set Default", self.set_default_folder)]:
            tk.Button(self.folder_row, text=txt, bg=c['btn_bg'], fg=c['btn_fg'],
                      activebackground=c['btn_hover'], font=('Sans', 9),
                      relief='flat', cursor='hand2', padx=10, pady=3,
                      command=cmd).pack(side='left', padx=(6, 0))

        self.update_output_mode()

        # -- File Naming dropdown --
        naming_row = tk.Frame(options_inner, bg=c['card_bg'])
        naming_row.pack(fill='x', pady=(6, 6))

        tk.Label(naming_row, text="File Naming:", font=('Sans', 10, 'bold'),
                 bg=c['card_bg'], fg=c['fg'], width=16, anchor='w').pack(side='left', padx=(0, 8))

        self.naming_display_map = {
            'Keep original filename': 'original',
            'Article # - Title': 'full_title',
            'Revised Article # - Title': 'revised_title',
        }
        self.naming_reverse_map = {v: k for k, v in self.naming_display_map.items()}

        saved_mode = self.settings.get('naming_mode', 'original')
        self.naming_display = tk.StringVar(
            value=self.naming_reverse_map.get(saved_mode, 'Keep original filename'))

        naming_combo = ttk.Combobox(
            naming_row, textvariable=self.naming_display,
            values=list(self.naming_display_map.keys()),
            state='readonly', font=('Sans', 10), width=28)
        naming_combo.pack(side='left')
        naming_combo.bind('<<ComboboxSelected>>', lambda e: self.update_naming_mode())

        # -- Separator --
        self._card_separator(opt_card)

        # -- Processing checkboxes --
        self._card_heading(opt_card, "Processing", top_pad=8)

        checks_frame = tk.Frame(opt_card, bg=c['card_bg'])
        checks_frame.pack(fill='x', padx=12, pady=(0, 10))

        self.overwrite_var = tk.BooleanVar(
            value=self.settings.get('overwrite_existing', True))
        ttk.Checkbutton(checks_frame,
                        text="Overwrite existing files (otherwise prompts for each)",
                        variable=self.overwrite_var,
                        command=self.update_overwrite_setting).pack(anchor='w')

        self.strip_ssf_var = tk.BooleanVar(
            value=self.settings.get('strip_self_service_flow', True))
        ttk.Checkbutton(checks_frame,
                        text="Remove Self Service Flow section",
                        variable=self.strip_ssf_var,
                        command=self.update_ssf_setting).pack(anchor='w', pady=(4, 0))

        self.strip_ai_var = tk.BooleanVar(
            value=self.settings.get('strip_ai_usage', True))
        ttk.Checkbutton(checks_frame,
                        text="Remove AI Usage section",
                        variable=self.strip_ai_var,
                        command=self.update_ai_setting).pack(anchor='w', pady=(4, 0))

        self.auto_watch_var = tk.BooleanVar(
            value=self.settings.get('auto_watch', False))
        ttk.Checkbutton(checks_frame,
                        text="Auto-convert new exports",
                        variable=self.auto_watch_var,
                        command=self.toggle_auto_watch).pack(anchor='w', pady=(4, 0))

        # Watch folder row (shown/hidden based on auto-watch)
        self.watch_folder_row = tk.Frame(checks_frame, bg=c['card_bg'])

        tk.Label(self.watch_folder_row, text="Watch Folder:",
                 font=('Sans', 9), bg=c['card_bg'], fg=c['fg']).pack(
            side='left', padx=(20, 6))

        self.watch_folder_entry = tk.Entry(
            self.watch_folder_row, font=('Sans', 9),
            bg=c['input_bg'], fg=c['fg'], insertbackground=c['fg'],
            relief='solid', borderwidth=1,
            highlightbackground=c['input_border'], highlightcolor=c['accent'])
        self.watch_folder_entry.pack(side='left', fill='x', expand=True, ipady=2)
        self.watch_folder_entry.insert(0, self.settings.get(
            'watch_folder', str(Path.home() / "Downloads")))

        tk.Button(self.watch_folder_row, text="Browse…",
                  bg=c['btn_bg'], fg=c['btn_fg'],
                  activebackground=c['btn_hover'], font=('Sans', 9),
                  relief='flat', cursor='hand2', padx=10, pady=2,
                  command=self.select_watch_folder).pack(side='left', padx=(6, 0))

        # Show/hide watch folder row based on checkbox state
        if self.auto_watch_var.get():
            self.watch_folder_row.pack(anchor='w', fill='x', pady=(4, 0))

        # ── Card: Progress ────────────────────────────────────────────
        prog_card = self._make_card(main, row=2, expand=True)

        self._card_heading(prog_card, "Progress")

        self.progress_bar = ttk.Progressbar(prog_card, mode='determinate')
        self.progress_bar.pack(fill='x', padx=12, pady=(10, 4))

        self.status_label = tk.Label(
            prog_card, text="Ready", font=('Sans', 10),
            bg=c['card_bg'], fg=c['fg_secondary'])
        self.status_label.pack(anchor='w', padx=12, pady=(0, 8))

        # Two-column list with visible column borders
        list_outer = tk.Frame(prog_card, bg=c['card_bg'])
        list_outer.pack(fill='both', expand=True, padx=12, pady=(0, 8))
        list_outer.grid_columnconfigure(0, weight=1)
        list_outer.grid_columnconfigure(1, weight=3)
        list_outer.grid_rowconfigure(1, weight=1)

        # Column headers
        left_header = tk.Frame(list_outer, bg=c['input_border'])
        left_header.grid(row=0, column=0, sticky='ew', padx=(0, 4))
        tk.Label(left_header, text="  Source Files", font=('Sans', 10, 'bold'),
                 bg=c['input_border'], fg=c['fg'],
                 anchor='w', pady=4).pack(fill='x')

        right_header = tk.Frame(list_outer, bg=c['input_border'])
        right_header.grid(row=0, column=1, sticky='ew', padx=(4, 0))
        tk.Label(right_header, text="  Output", font=('Sans', 10, 'bold'),
                 bg=c['input_border'], fg=c['fg'],
                 anchor='w', pady=4).pack(fill='x')

        # Left column with border
        left_border = tk.Frame(list_outer, bg=c['input_border'],
                               highlightthickness=0)
        left_border.grid(row=1, column=0, sticky='nsew', padx=(0, 4))

        left_inner = tk.Frame(left_border, bg=c['list_bg'])
        left_inner.pack(fill='both', expand=True, padx=1, pady=(0, 1))

        ls = tk.Scrollbar(left_inner, troughcolor=c['list_bg'])
        ls.pack(side='right', fill='y')
        self.file_listbox = tk.Listbox(
            left_inner, font=('Monospace', 10), bg=c['list_bg'], fg=c['fg'],
            selectbackground=c['accent'], relief='flat', borderwidth=0,
            yscrollcommand=ls.set, selectmode='none', highlightthickness=0)
        self.file_listbox.pack(side='left', fill='both', expand=True)
        ls.config(command=self.file_listbox.yview)

        # Right column with border
        right_border = tk.Frame(list_outer, bg=c['input_border'],
                                highlightthickness=0)
        right_border.grid(row=1, column=1, sticky='nsew', padx=(4, 0))

        right_inner = tk.Frame(right_border, bg=c['list_bg'])
        right_inner.pack(fill='both', expand=True, padx=1, pady=(0, 1))

        rs = tk.Scrollbar(right_inner, troughcolor=c['list_bg'])
        rs.pack(side='right', fill='y')
        self.status_listbox = tk.Listbox(
            right_inner, font=('Monospace', 10), bg=c['list_bg'], fg=c['fg'],
            selectbackground=c['accent'], relief='flat', borderwidth=0,
            yscrollcommand=rs.set, selectmode='none', highlightthickness=0,
            cursor='hand2')
        self.status_listbox.pack(side='left', fill='both', expand=True)
        rs.config(command=self.status_listbox.yview)
        self.status_listbox.bind('<Button-1>', self.on_status_click)

        main.grid_columnconfigure(0, weight=1)

        # Start watching on launch if previously enabled
        if self.auto_watch_var.get():
            self.start_watching()

    # ── Card Helpers ──────────────────────────────────────────────────────

    def _make_card(self, parent, row, expand=False):
        c = self.colors
        card = tk.Frame(parent, bg=c['card_bg'],
                        highlightbackground=c['card_border'],
                        highlightthickness=1)
        card.grid(row=row, column=0, sticky='nsew' if expand else 'ew',
                  pady=(0, 8))
        if expand:
            parent.grid_rowconfigure(row, weight=1)
        return card

    def _card_heading(self, card, text, top_pad=10):
        c = self.colors
        tk.Label(card, text=text, font=('Sans', 11, 'bold'),
                 bg=c['card_bg'], fg=c['fg']).pack(
            anchor='w', padx=12, pady=(top_pad, 6))
        tk.Frame(card, bg=c['separator'], height=1).pack(fill='x', padx=12)

    def _card_separator(self, card):
        c = self.colors
        tk.Frame(card, bg=c['separator'], height=1).pack(
            fill='x', padx=12, pady=(10, 0))

    # ── File Selection ────────────────────────────────────────────────────

    def select_files(self):
        try:
            result = subprocess.run([
                'zenity', '--file-selection', '--multiple', '--separator=|',
                '--title=Select HTML Files',
                '--file-filter=HTML Files | *.html',
                '--file-filter=All Files | *',
                f'--filename={self.settings["last_file_directory"]}/'
            ], capture_output=True, text=True, timeout=300)

            if result.returncode == 0 and result.stdout.strip():
                files = result.stdout.strip().split('|')
                self.selected_files = files
                if files:
                    self.settings['last_file_directory'] = str(Path(files[0]).parent)
                    self.save_settings()
                self.update_file_count()
                return
            elif result.returncode == 1:
                return
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        files = filedialog.askopenfilenames(
            title="Select HTML Files",
            filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")],
            initialdir=self.settings.get('last_file_directory', str(Path.home())))
        if files:
            self.selected_files = list(files)
            self.settings['last_file_directory'] = str(Path(files[0]).parent)
            self.save_settings()
            self.update_file_count()

    def clear_selection(self):
        self.selected_files = []
        self.file_listbox.delete(0, tk.END)
        self.status_listbox.delete(0, tk.END)
        self.output_files = []
        self.click_actions = []
        self.update_file_count()
        self.progress_bar['value'] = 0
        self.status_label.config(text="Ready")

    def update_file_count(self):
        count = len(self.selected_files)
        if count == 0:
            self.file_count_label.config(text="No files selected")
        elif count == 1:
            self.file_count_label.config(text="1 file selected")
        else:
            self.file_count_label.config(text=f"{count} files selected")

        self.file_listbox.delete(0, tk.END)
        for fp in self.selected_files:
            self.file_listbox.insert(tk.END, os.path.basename(fp))

    # ── Output Folder ─────────────────────────────────────────────────────

    def select_output_folder(self):
        try:
            result = subprocess.run([
                'zenity', '--file-selection', '--directory',
                '--title=Select Output Folder',
                f'--filename={self.folder_entry.get()}/'
            ], capture_output=True, text=True, timeout=300)
            if result.returncode == 0 and result.stdout.strip():
                self.folder_entry.delete(0, tk.END)
                self.folder_entry.insert(0, result.stdout.strip())
                return
            elif result.returncode == 1:
                return
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        folder = filedialog.askdirectory(
            title="Select Output Folder", initialdir=self.folder_entry.get())
        if folder:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)

    def set_default_folder(self):
        folder = self.folder_entry.get()
        if folder and os.path.isdir(folder):
            self.settings['output_folder'] = folder
            self.save_settings()
            messagebox.showinfo("Success", "Default output folder updated!")
        else:
            messagebox.showerror("Error", "Please select a valid folder")

    def update_output_mode(self):
        if self.output_mode.get() == 'Custom folder':
            self.folder_row.pack(fill='x', pady=(4, 6))
            self.settings['same_folder_as_source'] = False
        else:
            self.folder_row.pack_forget()
            self.settings['same_folder_as_source'] = True
        self.save_settings()

    def update_overwrite_setting(self):
        self.settings['overwrite_existing'] = self.overwrite_var.get()
        self.save_settings()

    def update_ssf_setting(self):
        self.settings['strip_self_service_flow'] = self.strip_ssf_var.get()
        self.save_settings()

    def update_ai_setting(self):
        self.settings['strip_ai_usage'] = self.strip_ai_var.get()
        self.save_settings()

    def update_naming_mode(self):
        display = self.naming_display.get()
        self.settings['naming_mode'] = self.naming_display_map.get(display, 'original')
        self.save_settings()

    # ── Conversion ────────────────────────────────────────────────────────

    def start_conversion(self):
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select HTML files to convert")
            return
        if self.is_converting:
            return

        if self.output_mode.get() == 'Custom folder':
            output_folder = self.folder_entry.get()
            if not output_folder or not os.path.isdir(output_folder):
                messagebox.showerror("Invalid Folder",
                                     "Please select a valid output folder")
                return

        self.is_converting = True
        self.convert_btn.config(state='disabled')
        self.select_btn.config(state='disabled')
        self.status_listbox.delete(0, tk.END)
        self.progress_bar['value'] = 0
        self.conversion_results = {'success': 0, 'failed': 0, 'skipped': 0}
        self.output_files = []

        Thread(target=self.convert_files, daemon=True).start()

    def determine_output_filename(self, html_file):
        """Determine output filename based on naming mode."""
        mode = self.naming_display_map.get(self.naming_display.get(), 'original')
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
            return self._sanitize_filename(name) + '.docx'

        # Fallback to original if parsing fails
        return original_base + '.docx'

    @staticmethod
    def _sanitize_filename(name, max_length=100):
        """Make a filename safe for Windows and Linux.

        Replaces / with -, removes illegal Windows characters,
        collapses double spaces, truncates to max_length on a word
        boundary, and strips trailing periods and spaces.
        """
        # Replace forward slash with hyphen
        name = name.replace('/', '-')
        # Remove remaining Windows-illegal characters: \ : * ? " < > |
        name = re.sub(r'[\\:*?"<>|]', '', name)
        # Collapse multiple spaces into one
        name = re.sub(r'  +', ' ', name)
        # Strip leading/trailing whitespace
        name = name.strip()
        # Truncate to max length on a word boundary
        if len(name) > max_length:
            name = name[:max_length].rsplit(' ', 1)[0]
        # Strip trailing periods and spaces (invalid on Windows)
        name = name.rstrip('. ')
        return name

    def convert_files(self):
        total = len(self.selected_files)

        for idx, html_file in enumerate(self.selected_files, 1):
            filename = os.path.basename(html_file)
            self.root.after(0, self.status_label.config,
                            {'text': f"Converting {idx}/{total}: {filename}"})

            out_name = self.determine_output_filename(html_file)

            if self.output_mode.get() == 'Same folder as source files':
                output_folder = str(Path(html_file).parent)
            else:
                output_folder = self.folder_entry.get()

            output_file = os.path.join(output_folder, out_name)
            self.last_output_folder = output_folder

            # Check existing file
            if os.path.exists(output_file) and not self.overwrite_var.get():
                import queue
                response_queue = queue.Queue()

                def ask_user(of=output_file):
                    base_name = os.path.splitext(os.path.basename(of))[0]
                    counter = 1
                    base_path = of.rsplit('.', 1)[0]
                    while os.path.exists(f"{base_path} ({counter}).docx"):
                        counter += 1
                    suggested = f"{base_name} ({counter})"

                    dlg = tk.Toplevel(self.root)
                    dlg.title("File Exists")
                    dlg.geometry("550x200")
                    dlg.resizable(False, False)
                    dlg.transient(self.root)
                    dlg.grab_set()
                    cc = self.colors
                    dlg.configure(bg=cc['card_bg'])

                    dlg.update_idletasks()
                    x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 275
                    y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 100
                    dlg.geometry(f"+{x}+{y}")

                    content = tk.Frame(dlg, bg=cc['card_bg'])
                    content.pack(fill='both', expand=True, padx=20, pady=16)

                    tk.Label(content,
                             text=f"{os.path.basename(of)} already exists.",
                             font=('Sans', 11, 'bold'),
                             bg=cc['card_bg'], fg=cc['fg']).pack(pady=(0, 12))

                    rename_row = tk.Frame(content, bg=cc['card_bg'])
                    rename_row.pack(fill='x', pady=(0, 12))

                    tk.Label(rename_row, text="Rename to:", font=('Sans', 10),
                             bg=cc['card_bg'], fg=cc['fg']).pack(
                        anchor='w', pady=(0, 4))

                    rename_entry = tk.Entry(
                        rename_row, font=('Sans', 10),
                        bg=cc['input_bg'], fg=cc['fg'],
                        insertbackground=cc['fg'],
                        relief='solid', borderwidth=1)
                    rename_entry.insert(0, suggested)
                    rename_entry.pack(fill='x', ipady=3)

                    btn_row = tk.Frame(content, bg=cc['card_bg'])
                    btn_row.pack()

                    def on_replace():
                        response_queue.put(('replace', None))
                        dlg.destroy()

                    def on_skip():
                        response_queue.put(('skip', None))
                        dlg.destroy()

                    def on_rename():
                        name = rename_entry.get().strip()
                        if name:
                            response_queue.put(('rename', name))
                            dlg.destroy()

                    for txt, cmd in [("Replace", on_replace),
                                     ("Save as Rename", on_rename),
                                     ("Skip", on_skip)]:
                        tk.Button(
                            btn_row, text=txt, command=cmd,
                            font=('Sans', 10), width=14,
                            bg=cc['btn_bg'], fg=cc['btn_fg'],
                            activebackground=cc['btn_hover'],
                            relief='flat', cursor='hand2').pack(
                            side='left', padx=3)

                    dlg.protocol("WM_DELETE_WINDOW", on_skip)
                    rename_entry.focus()
                    rename_entry.select_range(0, tk.END)

                self.root.after(0, ask_user)
                choice, custom_name = response_queue.get()

                if choice == 'skip':
                    self.conversion_results['skipped'] += 1
                    self.root.after(0, self.add_status_to_list,
                                    f"⊘ {os.path.basename(output_file)}",
                                    self.colors['warning'])
                    self.root.after(0, self.add_output_file, None, False)
                    progress = (idx / total) * 100
                    self.root.after(0, self.progress_bar.config,
                                    {'value': progress})
                    continue
                elif choice == 'rename' and custom_name:
                    output_dir = os.path.dirname(output_file)
                    if not custom_name.endswith('.docx'):
                        custom_name += '.docx'
                    output_file = os.path.join(output_dir, custom_name)

            # Convert (single call — was duplicated in original)
            success = self.convert_single_file(html_file, output_file)

            if success:
                self.conversion_results['success'] += 1
            else:
                self.conversion_results['failed'] += 1

            out_basename = os.path.basename(output_file)
            status = f"✓ {out_basename}" if success else f"✗ {out_basename}"
            color = self.colors['success'] if success else self.colors['error']
            self.root.after(0, self.add_status_to_list, status, color)
            self.root.after(0, self.add_output_file, output_file, success)

            progress = (idx / total) * 100
            self.root.after(0, self.progress_bar.config, {'value': progress})

        self.root.after(0, self.conversion_complete)

    def convert_single_file(self, input_file, output_file):
        """Convert single HTML file to DOCX, optionally stripping SSF."""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            actual_input = input_file
            tmp_input = None

            # Pre-process: strip sections as configured
            strip_ssf = self.strip_ssf_var.get()
            strip_ai = self.strip_ai_var.get()

            if strip_ssf or strip_ai:
                try:
                    with open(input_file, 'r', encoding='utf-8') as f:
                        html_content = f.read()

                    cleaned = html_content
                    if strip_ssf:
                        cleaned = strip_self_service_flow(cleaned)
                    if strip_ai:
                        cleaned = strip_ai_usage(cleaned)

                    if cleaned != html_content:
                        import tempfile
                        tmp = tempfile.NamedTemporaryFile(
                            mode='w', suffix='.html', delete=False,
                            encoding='utf-8')
                        tmp.write(cleaned)
                        tmp.close()
                        actual_input = tmp.name
                        tmp_input = tmp.name
                except Exception as e:
                    print(f"Warning: Could not pre-process {input_file}: {e}")

            # Lua filter and reference doc
            if REFERENCE_DOCX_B64 and len(REFERENCE_DOCX_B64) > 100:
                import tempfile
                with tempfile.NamedTemporaryFile(
                        mode='w', suffix='.lua', delete=False) as lua_file:
                    lua_file.write(LUA_FILTER)
                    lua_path = lua_file.name
                with tempfile.NamedTemporaryFile(
                        mode='wb', suffix='.docx', delete=False) as ref_file:
                    ref_file.write(base64.b64decode(REFERENCE_DOCX_B64))
                    ref_path = ref_file.name
            else:
                lua_path = os.path.join(script_dir,
                                        'filter.lua')
                ref_path = os.path.join(script_dir, 'reference.docx')
                if not os.path.exists(lua_path):
                    print(f"Error: Lua filter not found at {lua_path}")
                    return False
                if not os.path.exists(ref_path):
                    print(f"Error: Reference doc not found at {ref_path}")
                    return False

            pandoc_cmd = [
                'pandoc', actual_input,
                '--lua-filter', lua_path,
                '--reference-doc', ref_path,
                '-o', output_file
            ]

            if os.environ.get('PANDOC_DATADIR'):
                pandoc_cmd.insert(2, '--data-dir')
                pandoc_cmd.insert(3, os.environ['PANDOC_DATADIR'])

            result = subprocess.run(pandoc_cmd, capture_output=True, text=True)

            # Clean up temp files
            if REFERENCE_DOCX_B64 and len(REFERENCE_DOCX_B64) > 100:
                os.unlink(lua_path)
                os.unlink(ref_path)
            if tmp_input:
                os.unlink(tmp_input)

            if result.returncode != 0:
                print(f"Pandoc error: {result.stderr}")

            return result.returncode == 0

        except Exception as e:
            print(f"Error converting {input_file}: {e}")
            import traceback
            traceback.print_exc()
            return False

    # ── Auto-Watch ────────────────────────────────────────────────────────

    def select_watch_folder(self):
        """Browse for watch folder."""
        try:
            result = subprocess.run([
                'zenity', '--file-selection', '--directory',
                '--title=Select Watch Folder',
                f'--filename={self.watch_folder_entry.get()}/'
            ], capture_output=True, text=True, timeout=300)
            if result.returncode == 0 and result.stdout.strip():
                self.watch_folder_entry.delete(0, tk.END)
                self.watch_folder_entry.insert(0, result.stdout.strip())
                self.settings['watch_folder'] = result.stdout.strip()
                self.save_settings()
                # Restart watcher if active
                if self.watching:
                    self.stop_watching()
                    self.start_watching()
                return
            elif result.returncode == 1:
                return
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        folder = filedialog.askdirectory(
            title="Select Watch Folder",
            initialdir=self.watch_folder_entry.get())
        if folder:
            self.watch_folder_entry.delete(0, tk.END)
            self.watch_folder_entry.insert(0, folder)
            self.settings['watch_folder'] = folder
            self.save_settings()
            # Restart watcher if active
            if self.watching:
                self.stop_watching()
                self.start_watching()

    def toggle_auto_watch(self):
        """Toggle auto-convert watching on/off."""
        self.settings['auto_watch'] = self.auto_watch_var.get()
        self.save_settings()
        if self.auto_watch_var.get():
            self.watch_folder_row.pack(anchor='w', fill='x', pady=(4, 0))
            self.start_watching()
        else:
            self.watch_folder_row.pack_forget()
            self.stop_watching()

    def start_watching(self):
        """Start watching the configured folder for new Article *.html files."""
        watch_path = self.watch_folder_entry.get()
        if not os.path.isdir(watch_path):
            messagebox.showerror("Invalid Folder",
                                 f"Watch folder does not exist:\n{watch_path}")
            self.auto_watch_var.set(False)
            self.settings['auto_watch'] = False
            self.save_settings()
            return

        self.watching = True
        handler = _ArticleFileHandler(self)
        self.observer = Observer()
        self.observer.schedule(handler, watch_path, recursive=True)
        self.observer.start()

        # Shorten path for display
        display_path = watch_path.replace(str(Path.home()), '~')
        self.status_label.config(text=f"Watching {display_path} for new exports...")

    def stop_watching(self):
        """Stop watching."""
        self.watching = False
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
            self.observer = None
        self.status_label.config(text="Ready")

    def _auto_convert(self, html_file):
        """Auto-convert a single file, saving to the same folder as the source."""
        if self.is_converting:
            return

        basename = os.path.basename(html_file)
        self.status_label.config(text=f"Auto-converting: {basename}")

        # Add to source list for visibility (blank second line to align with output)
        self.file_listbox.insert(tk.END, basename)
        self.file_listbox.insert(tk.END, "")

        # Determine output file — always same folder as source for auto-convert
        out_name = self.determine_output_filename(html_file)
        output_dir = os.path.dirname(html_file)
        output_file = os.path.join(output_dir, out_name)
        self.last_output_folder = output_dir

        # Handle overwrite — silently skip if file exists and overwrite is off
        if os.path.exists(output_file) and not self.overwrite_var.get():
            self.add_status_to_list(f"⊘ {out_name} (exists)", self.colors['warning'])
            display_path = self.watch_folder_entry.get().replace(str(Path.home()), '~')
            self.status_label.config(text=f"Watching {display_path} for new exports...")
            return

        # Shorten destination for display in output column
        rel_dir = output_dir.replace(str(Path.home()), '~')
        # Show just the last 2-3 folder segments if path is long
        parts = rel_dir.split(os.sep)
        if len(parts) > 4:
            short_dir = os.sep.join(['...'] + parts[-2:])
        else:
            short_dir = rel_dir

        # Run conversion in a thread to keep UI responsive
        def do_convert():
            success = self.convert_single_file(html_file, output_file)

            if success:
                self.root.after(0, self._add_auto_status,
                                out_name, output_dir, short_dir, output_file, True)
            else:
                self.root.after(0, self._add_auto_status,
                                out_name, output_dir, short_dir, output_file, False)

            display_path = self.watch_folder_entry.get().replace(str(Path.home()), '~')
            self.root.after(0, self.status_label.config,
                            {'text': f"Watching {display_path} for new exports..."})

        Thread(target=do_convert, daemon=True).start()

    # ── Status Helpers ────────────────────────────────────────────────────

    def _add_auto_status(self, out_name, output_dir, short_dir, output_file, success):
        """Add two-line status entry for auto-converted file with click actions."""
        if success:
            self.add_status_to_list(f"✓ {out_name}", self.colors['success'])
            self.click_actions.append(('file', output_file))
            self.add_status_to_list(f"    → {short_dir}", self.colors['fg_secondary'])
            self.click_actions.append(('folder', output_dir))
        else:
            self.add_status_to_list(f"✗ {out_name}", self.colors['error'])
            self.click_actions.append(None)
        self.output_files.append(output_file if success else None)

    def add_status_to_list(self, text, color):
        self.status_listbox.insert(tk.END, text)
        idx = self.status_listbox.size() - 1
        self.status_listbox.itemconfig(idx, fg=color)
        self.status_listbox.see(idx)
        self.file_listbox.see(idx)

    def add_output_file(self, output_file, success):
        self.output_files.append(output_file if success else None)

    def on_status_click(self, event):
        index = self.status_listbox.nearest(event.y)

        # Try click_actions first (auto-convert entries)
        if index < len(self.click_actions) and self.click_actions[index]:
            action_type, path = self.click_actions[index]
            if os.path.exists(path):
                try:
                    subprocess.run(['xdg-open', path])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open:\n{e}")
            else:
                messagebox.showerror("Not Found", f"Path does not exist:\n{path}")
            return

        # Fallback for manual conversion entries (no click_actions entry)
        if index < len(self.output_files) and self.output_files[index]:
            path = self.output_files[index]
            if os.path.exists(path):
                try:
                    subprocess.run(['xdg-open', path])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open file:\n{e}")
            else:
                messagebox.showerror("Not Found",
                                     f"File does not exist:\n{path}")

    def conversion_complete(self):
        self.is_converting = False
        self.convert_btn.config(state='normal')
        self.select_btn.config(state='normal')

        total = sum(self.conversion_results.values())
        s = self.conversion_results['success']
        f = self.conversion_results['failed']
        sk = self.conversion_results['skipped']

        if f == 0 and sk == 0:
            self.status_label.config(
                text=f"Successfully converted {s} file(s)!")
            messagebox.showinfo("Success",
                                f"Successfully converted {s} file(s)!")
        elif s == 0 and sk > 0 and f == 0:
            self.status_label.config(text=f"All {sk} file(s) were skipped.")
            messagebox.showinfo("Skipped", f"All {sk} file(s) were skipped.")
        elif s == 0:
            msg = f"Conversion failed. 0 of {total} converted."
            if sk > 0:
                msg += f" {sk} skipped."
            self.status_label.config(text=msg)
            messagebox.showerror("Failed", msg)
        else:
            parts = [f"Converted {s} of {total} files."]
            if f > 0:
                parts.append(f"{f} failed.")
            if sk > 0:
                parts.append(f"{sk} skipped.")
            msg = " ".join(parts)
            self.status_label.config(text=msg)
            messagebox.showwarning("Partial Success", msg)

    def open_output_folder(self):
        if self.last_output_folder and os.path.isdir(self.last_output_folder):
            folder = self.last_output_folder
        elif self.output_mode.get() == 'Custom folder':
            folder = self.folder_entry.get()
        elif self.selected_files:
            folder = str(Path(self.selected_files[0]).parent)
        else:
            messagebox.showinfo("No Output",
                                "No files have been converted yet.")
            return

        if not os.path.isdir(folder):
            messagebox.showerror("Not Found",
                                 f"Folder does not exist:\n{folder}")
            return

        try:
            subprocess.run(['xdg-open', folder])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def on_closing(self):
        # Stop the watchdog observer on exit
        if self.observer:
            self.observer.stop()
            self.observer.join(timeout=2)
        self.settings['watch_folder'] = self.watch_folder_entry.get()
        self.settings['window_geometry'] = self.root.geometry().split('+')[0]
        self.settings['window_position'] = [
            self.root.winfo_x(), self.root.winfo_y()]
        self.save_settings()
        self.root.destroy()


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
            # Small delay to let the file finish writing
            time.sleep(0.5)
            self.app.root.after(0, self.app._auto_convert, event.src_path)


def main():
    root = tk.Tk()
    app = KnowledgeArticleConverter(root)
    root.mainloop()


if __name__ == '__main__':
    main()
