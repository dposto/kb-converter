# Changelog

## v2.0 — PyQt6 rewrite

A full rewrite of the original tkinter application, replacing the minimal
proof-of-concept UI with a proper desktop application.

**UI**
- Migrated from tkinter to PyQt6
- Menu bar with File, Settings, View, Auto-Convert, and Help menus
- Splitter layout — file list and conversion log in resizable panes
- CRT-style terminal log with phosphor green text and live status output
- Live theme switching — System, Light, and Dark (View → Theme)
- Clock emoji spinner during active conversions

**Conversion**
- Deep pandoc verbose output piped directly to the log panel
- Cleaner error reporting with inline failure indicators

**Help menu**
- Getting Started and User Guide dialogs
- License & Third-Party Notices dialog covering all dependencies
- About dialog

**Settings**
- Watch folder configurable from the UI
- File naming modes: keep original, Article # - Title, Revised Article # - Title
- Processing options: Remove Self Service Flow, Remove AI Usage, Overwrite existing
- Settings persist between sessions

---

## v1.x — Initial release (tkinter)

Original proof-of-concept built to solve an immediate workflow problem:
automatically converting knowledge base HTML exports to formatted DOCX
without manual intervention.

**Features**
- tkinter UI — minimal, functional
- watchdog-based folder monitoring
- pandoc conversion with reference.docx template and Lua highlight filter
- Two-line status display
- Clickable output entries to open converted files or folders
- Configurable watch folder
- Basic theme switching

**Foundation**
- Established the core pipeline: HTML export → watched folder → pandoc →
  formatted DOCX
- Proved out the Lua filter approach for HTML highlight span handling
- Identified the pandoc list numbering and indent issues that led to
  the Docx.hs patch
