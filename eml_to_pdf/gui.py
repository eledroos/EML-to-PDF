"""Modern graphical user interface for EML to PDF Converter."""

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional, Callable, List, Dict

from .config import ConversionConfig, AVAILABLE_THEMES, PAGE_SIZES, AVAILABLE_FONTS
from .converter import convert_batch, create_skipped_files_report, BatchConversionResult

# Try to import ttkbootstrap, fall back to standard tkinter
try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    from ttkbootstrap.dialogs import Messagebox
    from ttkbootstrap.scrolled import ScrolledFrame
    TTKBOOTSTRAP_AVAILABLE = True
except ImportError:
    import tkinter as tk
    from tkinter import ttk
    TTKBOOTSTRAP_AVAILABLE = False

# Try to import tkinterdnd2 for drag-and-drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False

# Recent folders storage
RECENT_FOLDERS_FILE = Path.home() / ".eml_to_pdf_recent.json"
MAX_RECENT_FOLDERS = 5


def load_recent_folders() -> List[Dict]:
    """Load recent folders from storage."""
    if RECENT_FOLDERS_FILE.exists():
        try:
            with open(RECENT_FOLDERS_FILE, 'r') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return []


def save_recent_folders(folders: List[Dict]) -> None:
    """Save recent folders to storage."""
    try:
        with open(RECENT_FOLDERS_FILE, 'w') as f:
            json.dump(folders[:MAX_RECENT_FOLDERS], f)
    except IOError:
        pass


def add_recent_folder(path: str, file_count: int) -> None:
    """Add a folder to recent folders list."""
    folders = load_recent_folders()
    # Remove if already exists
    folders = [f for f in folders if f.get('path') != path]
    # Add to front
    folders.insert(0, {
        'path': path,
        'file_count': file_count,
        'last_used': datetime.now().isoformat()
    })
    save_recent_folders(folders)


class ProgressWindow:
    """Modern progress window with detailed information."""

    def __init__(self, parent, total_files: int):
        if TTKBOOTSTRAP_AVAILABLE:
            self.popup = ttk.Toplevel(parent)
        else:
            self.popup = tk.Toplevel(parent)

        self.popup.title("Converting...")
        self.popup.geometry("520x220")
        self.popup.resizable(False, False)
        self.popup.protocol("WM_DELETE_WINDOW", self.on_close)
        self.popup.transient(parent)
        self.popup.grab_set()

        # Center on parent
        self.popup.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 520) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 220) // 2
        self.popup.geometry(f"+{x}+{y}")

        self.cancelled = False
        self.start_time = time.time()
        self.total_files = total_files

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.popup, padding=25)
        main_frame.pack(fill="both", expand=True)

        # Title
        title = ttk.Label(main_frame, text="Converting EML to PDF", font=("TkDefaultFont", 14, "bold"))
        title.pack(anchor="w")

        # Current file
        self.file_var = self._sv("Preparing...")
        file_label = ttk.Label(main_frame, textvariable=self.file_var, foreground="gray")
        file_label.pack(anchor="w", pady=(5, 15))

        # Progress bar
        if TTKBOOTSTRAP_AVAILABLE:
            self.progress = ttk.Progressbar(main_frame, length=460, mode="determinate",
                                             bootstyle="success-striped", maximum=self.total_files)
        else:
            self.progress = ttk.Progressbar(main_frame, length=460, mode="determinate",
                                             maximum=self.total_files)
        self.progress.pack(fill="x")

        # Stats row
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(fill="x", pady=(10, 0))

        self.count_var = self._sv("0 / 0 files")
        ttk.Label(stats_frame, textvariable=self.count_var).pack(side="left")

        self.eta_var = self._sv("")
        ttk.Label(stats_frame, textvariable=self.eta_var, foreground="gray").pack(side="right")

        # Cancel button
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(20, 0))

        if TTKBOOTSTRAP_AVAILABLE:
            self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.cancel,
                                          bootstyle="danger-outline", width=15)
        else:
            self.cancel_btn = ttk.Button(btn_frame, text="Cancel", command=self.cancel, width=15)
        self.cancel_btn.pack(side="right")

    def _sv(self, value: str):
        if TTKBOOTSTRAP_AVAILABLE:
            return ttk.StringVar(value=value)
        return tk.StringVar(value=value)

    def update(self, current_file: str, files_processed: int) -> bool:
        if self.cancelled:
            return False

        # Truncate filename
        display_name = current_file if len(current_file) < 50 else current_file[:47] + "..."
        self.file_var.set(display_name)
        self.progress["value"] = files_processed
        self.count_var.set(f"{files_processed} / {self.total_files} files")

        # ETA calculation
        elapsed = time.time() - self.start_time
        if files_processed > 0:
            remaining = self.total_files - files_processed
            eta = (elapsed / files_processed) * remaining
            if eta < 60:
                self.eta_var.set(f"~{int(eta)}s remaining")
            else:
                self.eta_var.set(f"~{int(eta/60)}m {int(eta%60)}s remaining")

        self.popup.update()
        return True

    def cancel(self):
        self.cancelled = True
        self.cancel_btn.config(state="disabled", text="Cancelling...")

    def on_close(self):
        self.cancel()

    def destroy(self):
        self.popup.destroy()


class SettingsWindow:
    """Settings dialog window."""

    def __init__(self, parent, config: ConversionConfig, on_save: Callable):
        self.config = config
        self.on_save = on_save

        if TTKBOOTSTRAP_AVAILABLE:
            self.window = ttk.Toplevel(parent)
        else:
            self.window = tk.Toplevel(parent)

        self.window.title("Settings")
        self.window.geometry("420x500")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        # Center on parent
        self.window.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 420) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
        self.window.geometry(f"+{x}+{y}")

        self._create_widgets()

    def _create_widgets(self):
        main = ttk.Frame(self.window, padding=25)
        main.pack(fill="both", expand=True)

        # PDF Settings
        self._section_label(main, "PDF Output")

        row = ttk.Frame(main)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text="Page Size:", width=15, anchor="w").pack(side="left")
        self.page_var = self._sv(self.config.page_size)
        ttk.Combobox(row, textvariable=self.page_var, values=PAGE_SIZES,
                     state="readonly", width=18).pack(side="right")

        row = ttk.Frame(main)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text="Font:", width=15, anchor="w").pack(side="left")
        self.font_var = self._sv(self.config.font_family)
        ttk.Combobox(row, textvariable=self.font_var, values=AVAILABLE_FONTS,
                     state="readonly", width=18).pack(side="right")

        # Processing
        self._section_label(main, "Processing", pady=(20, 5))

        self.organize_var = self._bv(self.config.organize_by_date)
        ttk.Checkbutton(main, text="Organize output by year/month folders",
                        variable=self.organize_var).pack(anchor="w", pady=2)

        self.attach_var = self._bv(self.config.extract_attachments)
        ttk.Checkbutton(main, text="Extract email attachments",
                        variable=self.attach_var).pack(anchor="w", pady=2)

        self.weasy_var = self._bv(self.config.use_weasyprint)
        ttk.Checkbutton(main, text="Use WeasyPrint for HTML emails",
                        variable=self.weasy_var).pack(anchor="w", pady=2)

        self.addressbook_var = self._bv(self.config.generate_address_book)
        ttk.Checkbutton(main, text="Generate address book (CSV)",
                        variable=self.addressbook_var).pack(anchor="w", pady=2)

        # Metadata
        self._section_label(main, "Include in PDF", pady=(20, 5))

        meta_frame = ttk.Frame(main)
        meta_frame.pack(fill="x")

        left = ttk.Frame(meta_frame)
        left.pack(side="left", anchor="nw")
        right = ttk.Frame(meta_frame)
        right.pack(side="left", anchor="nw", padx=(40, 0))

        self.meta_vars = {
            'subject': self._bv(self.config.include_subject),
            'from': self._bv(self.config.include_from),
            'to': self._bv(self.config.include_to),
            'cc': self._bv(self.config.include_cc),
            'bcc': self._bv(self.config.include_bcc),
            'date': self._bv(self.config.include_date),
        }

        ttk.Checkbutton(left, text="Subject", variable=self.meta_vars['subject']).pack(anchor="w")
        ttk.Checkbutton(left, text="From", variable=self.meta_vars['from']).pack(anchor="w")
        ttk.Checkbutton(left, text="To", variable=self.meta_vars['to']).pack(anchor="w")
        ttk.Checkbutton(right, text="CC", variable=self.meta_vars['cc']).pack(anchor="w")
        ttk.Checkbutton(right, text="BCC", variable=self.meta_vars['bcc']).pack(anchor="w")
        ttk.Checkbutton(right, text="Date", variable=self.meta_vars['date']).pack(anchor="w")

        # Theme (if available)
        if TTKBOOTSTRAP_AVAILABLE:
            self._section_label(main, "Appearance", pady=(20, 5))
            row = ttk.Frame(main)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text="Theme:", width=15, anchor="w").pack(side="left")
            self.theme_var = self._sv(self.config.theme)
            ttk.Combobox(row, textvariable=self.theme_var, values=AVAILABLE_THEMES,
                         state="readonly", width=18).pack(side="right")

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(30, 0))

        if TTKBOOTSTRAP_AVAILABLE:
            ttk.Button(btn_frame, text="Cancel", command=self.window.destroy,
                       bootstyle="secondary", width=12).pack(side="left")
            ttk.Button(btn_frame, text="Save", command=self.save,
                       bootstyle="success", width=12).pack(side="right")
        else:
            ttk.Button(btn_frame, text="Cancel", command=self.window.destroy, width=12).pack(side="left")
            ttk.Button(btn_frame, text="Save", command=self.save, width=12).pack(side="right")

    def _section_label(self, parent, text, pady=(0, 5)):
        lbl = ttk.Label(parent, text=text, font=("TkDefaultFont", 11, "bold"))
        lbl.pack(anchor="w", pady=pady)

    def _sv(self, value):
        return ttk.StringVar(value=value) if TTKBOOTSTRAP_AVAILABLE else tk.StringVar(value=value)

    def _bv(self, value):
        return ttk.BooleanVar(value=value) if TTKBOOTSTRAP_AVAILABLE else tk.BooleanVar(value=value)

    def save(self):
        self.config.page_size = self.page_var.get()
        self.config.font_family = self.font_var.get()
        self.config.organize_by_date = self.organize_var.get()
        self.config.extract_attachments = self.attach_var.get()
        self.config.use_weasyprint = self.weasy_var.get()
        self.config.generate_address_book = self.addressbook_var.get()
        self.config.include_subject = self.meta_vars['subject'].get()
        self.config.include_from = self.meta_vars['from'].get()
        self.config.include_to = self.meta_vars['to'].get()
        self.config.include_cc = self.meta_vars['cc'].get()
        self.config.include_bcc = self.meta_vars['bcc'].get()
        self.config.include_date = self.meta_vars['date'].get()

        if TTKBOOTSTRAP_AVAILABLE and hasattr(self, 'theme_var'):
            self.config.theme = self.theme_var.get()

        self.config.save()
        self.on_save(self.config)
        self.window.destroy()


class EMLConverterApp:
    """Modern EML to PDF Converter application."""

    def __init__(self):
        self.config = ConversionConfig.load()
        self.selected_folder = None
        self.eml_files = []
        self.root = self._create_root()
        self._setup_ui()

    def _create_root(self):
        if TTKBOOTSTRAP_AVAILABLE:
            if TKDND_AVAILABLE:
                root = TkinterDnD.Tk()
                root.style = ttk.Style(self.config.theme)
            else:
                root = ttk.Window(themename=self.config.theme)
        else:
            if TKDND_AVAILABLE:
                root = TkinterDnD.Tk()
            else:
                root = tk.Tk()

        root.title("EML to PDF Converter")
        root.geometry("580x620")
        root.resizable(False, False)
        return root

    def _sv(self, value=""):
        return ttk.StringVar(value=value) if TTKBOOTSTRAP_AVAILABLE else tk.StringVar(value=value)

    def _bv(self, value=False):
        return ttk.BooleanVar(value=value) if TTKBOOTSTRAP_AVAILABLE else tk.BooleanVar(value=value)

    def _setup_ui(self):
        # Main container
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill="both", expand=True)

        # Header
        self._create_header()

        # Content area (switches between main view and preview view)
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill="both", expand=True, pady=(15, 0))

        # Show main view
        self._show_main_view()

    def _create_header(self):
        header = ttk.Frame(self.main_frame)
        header.pack(fill="x")

        # Title
        title = ttk.Label(header, text="EML to PDF", font=("TkDefaultFont", 22, "bold"))
        title.pack(side="left")

        subtitle = ttk.Label(header, text="Converter", font=("TkDefaultFont", 22), foreground="gray")
        subtitle.pack(side="left", padx=(8, 0))

        # Settings button
        if TTKBOOTSTRAP_AVAILABLE:
            settings_btn = ttk.Button(header, text="Settings", command=self.open_settings,
                                       bootstyle="secondary-outline", width=10)
        else:
            settings_btn = ttk.Button(header, text="Settings", command=self.open_settings, width=10)
        settings_btn.pack(side="right")

    def _show_main_view(self):
        """Show the main drop zone view."""
        # Clear content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        self.selected_folder = None
        self.eml_files = []

        # Drop zone card
        self._create_drop_zone()

        # Quick settings
        self._create_quick_settings()

        # Recent folders
        self._create_recent_folders()

    def _create_drop_zone(self):
        """Create the large drop zone."""
        # Container with border
        if TTKBOOTSTRAP_AVAILABLE:
            self.drop_card = ttk.Frame(self.content_frame, bootstyle="light", padding=3)
        else:
            self.drop_card = ttk.Frame(self.content_frame, relief="solid", borderwidth=1)
        self.drop_card.pack(fill="x", pady=(0, 15))

        # Inner frame
        if TTKBOOTSTRAP_AVAILABLE:
            self.drop_inner = ttk.Frame(self.drop_card, bootstyle="light", padding=40)
        else:
            self.drop_inner = ttk.Frame(self.drop_card, padding=40)
        self.drop_inner.pack(fill="both", expand=True)

        # Icon (using text as placeholder)
        icon_label = ttk.Label(self.drop_inner, text="📧", font=("TkDefaultFont", 48))
        icon_label.pack()

        # Main text
        main_text = ttk.Label(self.drop_inner, text="Drop folder here",
                               font=("TkDefaultFont", 16, "bold"))
        main_text.pack(pady=(10, 5))

        # Sub text
        sub_text = ttk.Label(self.drop_inner, text="or click to browse for EML files",
                              foreground="gray")
        sub_text.pack()

        # Browse button
        if TTKBOOTSTRAP_AVAILABLE:
            browse_btn = ttk.Button(self.drop_inner, text="Browse Folder",
                                     command=self.browse_folder, bootstyle="primary", width=15)
        else:
            browse_btn = ttk.Button(self.drop_inner, text="Browse Folder",
                                     command=self.browse_folder, width=15)
        browse_btn.pack(pady=(20, 0))

        # Make entire drop zone clickable
        for widget in [self.drop_card, self.drop_inner, icon_label, main_text, sub_text]:
            widget.bind("<Button-1>", lambda e: self.browse_folder())

        # Setup drag-and-drop
        if TKDND_AVAILABLE:
            self.drop_card.drop_target_register(DND_FILES)
            self.drop_card.dnd_bind("<<Drop>>", self.on_drop)
            self.drop_card.dnd_bind("<<DragEnter>>", self.on_drag_enter)
            self.drop_card.dnd_bind("<<DragLeave>>", self.on_drag_leave)

    def _create_quick_settings(self):
        """Create quick settings toggles."""
        settings_frame = ttk.Frame(self.content_frame)
        settings_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(settings_frame, text="Quick Options:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")

        options_row = ttk.Frame(settings_frame)
        options_row.pack(fill="x", pady=(5, 0))

        self.attach_var = self._bv(self.config.extract_attachments)
        self.organize_var = self._bv(self.config.organize_by_date)
        self.addressbook_var = self._bv(self.config.generate_address_book)

        ttk.Checkbutton(options_row, text="Extract attachments",
                        variable=self.attach_var, command=self._update_config).pack(side="left")
        ttk.Checkbutton(options_row, text="Organize by date",
                        variable=self.organize_var, command=self._update_config).pack(side="left", padx=(20, 0))
        ttk.Checkbutton(options_row, text="Generate address book",
                        variable=self.addressbook_var, command=self._update_config).pack(side="left", padx=(20, 0))

    def _update_config(self):
        """Update config from quick settings."""
        self.config.extract_attachments = self.attach_var.get()
        self.config.organize_by_date = self.organize_var.get()
        self.config.generate_address_book = self.addressbook_var.get()
        self.config.save()

    def _create_recent_folders(self):
        """Create recent folders section."""
        folders = load_recent_folders()
        if not folders:
            return

        # Header
        header = ttk.Frame(self.content_frame)
        header.pack(fill="x")

        ttk.Label(header, text="Recent Folders", font=("TkDefaultFont", 10, "bold")).pack(side="left")

        # Folder list
        for folder in folders[:3]:
            self._create_folder_item(folder)

    def _create_folder_item(self, folder: Dict):
        """Create a recent folder item."""
        path = folder.get('path', '')
        count = folder.get('file_count', 0)

        # Check if folder still exists and count files
        if os.path.isdir(path):
            current_count = len([f for f in os.listdir(path) if f.lower().endswith('.eml')])
        else:
            return  # Skip non-existent folders

        item = ttk.Frame(self.content_frame)
        item.pack(fill="x", pady=2)

        # Folder path (truncated)
        display_path = path if len(path) < 45 else "..." + path[-42:]

        if TTKBOOTSTRAP_AVAILABLE:
            btn = ttk.Button(item, text=f"📁 {display_path} ({current_count} files)",
                             command=lambda p=path: self.select_folder(p),
                             bootstyle="link", padding=0)
        else:
            btn = ttk.Button(item, text=f"📁 {display_path} ({current_count} files)",
                             command=lambda p=path: self.select_folder(p))
        btn.pack(anchor="w")

    def _show_preview_view(self):
        """Show the file preview view."""
        # Clear content
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # Selected folder info
        info_frame = ttk.Frame(self.content_frame)
        info_frame.pack(fill="x")

        ttk.Label(info_frame, text="Selected Folder:", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")

        # Truncate path
        display_path = self.selected_folder
        if len(display_path) > 60:
            display_path = "..." + display_path[-57:]
        ttk.Label(info_frame, text=display_path, foreground="gray").pack(anchor="w")

        ttk.Label(info_frame, text=f"Found {len(self.eml_files)} EML files",
                   font=("TkDefaultFont", 12, "bold")).pack(anchor="w", pady=(10, 0))

        # File list
        list_frame = ttk.Frame(self.content_frame)
        list_frame.pack(fill="both", expand=True, pady=15)

        # Create scrollable list
        if TTKBOOTSTRAP_AVAILABLE:
            canvas_frame = ttk.Frame(list_frame, bootstyle="light", padding=1)
        else:
            canvas_frame = ttk.Frame(list_frame, relief="solid", borderwidth=1)
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, height=200, highlightthickness=0) if not TTKBOOTSTRAP_AVAILABLE else ttk.Canvas(canvas_frame, height=200)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)

        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Add files to list
        for i, (filename, mtime) in enumerate(self.eml_files[:50]):  # Show max 50
            row = ttk.Frame(scrollable)
            row.pack(fill="x", padx=10, pady=2)

            ttk.Label(row, text=filename[:50] + "..." if len(filename) > 50 else filename).pack(side="left")
            ttk.Label(row, text=mtime.strftime("%Y-%m-%d"), foreground="gray").pack(side="right")

        if len(self.eml_files) > 50:
            ttk.Label(scrollable, text=f"... and {len(self.eml_files) - 50} more files",
                       foreground="gray").pack(pady=5)

        # Output path
        output_path = os.path.join(self.selected_folder, "PDF")
        out_frame = ttk.Frame(self.content_frame)
        out_frame.pack(fill="x")

        ttk.Label(out_frame, text="Output:", font=("TkDefaultFont", 10, "bold")).pack(side="left")
        ttk.Label(out_frame, text=output_path, foreground="gray").pack(side="left", padx=(5, 0))

        # Quick settings (again for this view)
        settings_frame = ttk.Frame(self.content_frame)
        settings_frame.pack(fill="x", pady=(15, 0))

        self.attach_var = self._bv(self.config.extract_attachments)
        self.organize_var = self._bv(self.config.organize_by_date)
        self.addressbook_var = self._bv(self.config.generate_address_book)

        ttk.Checkbutton(settings_frame, text="Extract attachments",
                        variable=self.attach_var, command=self._update_config).pack(side="left")
        ttk.Checkbutton(settings_frame, text="Organize by date",
                        variable=self.organize_var, command=self._update_config).pack(side="left", padx=(20, 0))
        ttk.Checkbutton(settings_frame, text="Generate address book",
                        variable=self.addressbook_var, command=self._update_config).pack(side="left", padx=(20, 0))

        # Action buttons
        btn_frame = ttk.Frame(self.content_frame)
        btn_frame.pack(fill="x", pady=(20, 0))

        if TTKBOOTSTRAP_AVAILABLE:
            ttk.Button(btn_frame, text="Cancel", command=self._show_main_view,
                       bootstyle="secondary", width=12).pack(side="left")
            ttk.Button(btn_frame, text=f"Convert {len(self.eml_files)} Files",
                       command=self.start_conversion, bootstyle="success", width=20).pack(side="right")
        else:
            ttk.Button(btn_frame, text="Cancel", command=self._show_main_view, width=12).pack(side="left")
            ttk.Button(btn_frame, text=f"Convert {len(self.eml_files)} Files",
                       command=self.start_conversion, width=20).pack(side="right")

    def browse_folder(self):
        """Open folder browser."""
        folder = filedialog.askdirectory(title="Select Folder Containing EML Files")
        if folder:
            self.select_folder(folder)

    def select_folder(self, path: str):
        """Select a folder and show preview."""
        if not os.path.isdir(path):
            self.show_error("Folder not found.")
            return

        # Find EML files
        eml_files = []
        for f in os.listdir(path):
            if f.lower().endswith('.eml'):
                fpath = os.path.join(path, f)
                mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                eml_files.append((f, mtime))

        if not eml_files:
            self.show_error("No EML files found in this folder.")
            return

        # Sort by date (newest first)
        eml_files.sort(key=lambda x: x[1], reverse=True)

        self.selected_folder = path
        self.eml_files = eml_files

        # Add to recent
        add_recent_folder(path, len(eml_files))

        # Show preview
        self._show_preview_view()

    def on_drop(self, event):
        """Handle dropped files/folders."""
        data = event.data
        if data.startswith("{") and data.endswith("}"):
            paths = [data[1:-1]]
        else:
            paths = data.split()

        if paths:
            self.select_folder(paths[0])

        self.on_drag_leave(event)

    def on_drag_enter(self, event):
        """Visual feedback on drag enter."""
        if TTKBOOTSTRAP_AVAILABLE:
            self.drop_card.configure(bootstyle="info")
            self.drop_inner.configure(bootstyle="info")

    def on_drag_leave(self, event):
        """Visual feedback on drag leave."""
        if TTKBOOTSTRAP_AVAILABLE:
            self.drop_card.configure(bootstyle="light")
            self.drop_inner.configure(bootstyle="light")

    def start_conversion(self):
        """Start the conversion process."""
        if not self.selected_folder or not self.eml_files:
            return

        # Update config from current settings
        self.config.extract_attachments = self.attach_var.get()
        self.config.organize_by_date = self.organize_var.get()

        # Create progress window
        progress = ProgressWindow(self.root, len(self.eml_files))

        def progress_callback(current: int, total: int, filename: str) -> bool:
            return progress.update(filename, current)

        # Run conversion
        result = convert_batch(
            input_folder=self.selected_folder,
            config=self.config,
            progress_callback=progress_callback
        )

        progress.destroy()

        # Create skipped files report
        if result.failed > 0:
            create_skipped_files_report(result.results, result.output_folder)

        # Show result and return to main view
        self._show_result(result)
        self._show_main_view()

    def _show_result(self, result: BatchConversionResult):
        """Show conversion result message."""
        if result.cancelled:
            message = (f"Conversion cancelled.\n\n"
                      f"Processed: {result.successful} of {result.total_files} files\n"
                      f"Location: {result.output_folder}")
            title = "Cancelled"
        else:
            message = f"Successfully converted {result.successful} emails to PDF.\n\n"
            message += f"Output: {result.output_folder}\n"

            if result.address_book_path:
                message += f"\nAddress book: address_book.csv"

            if result.failed > 0:
                message += f"\n{result.failed} files skipped (see Skipped_Files_Report.pdf)"

            title = "Conversion Complete"

        self.show_info(title, message)

    def open_settings(self):
        """Open settings window."""
        def on_save(config):
            self.config = config
            if TTKBOOTSTRAP_AVAILABLE:
                self.root.style.theme_use(config.theme)
            # Update quick settings vars if they exist
            if hasattr(self, 'attach_var'):
                self.attach_var.set(config.extract_attachments)
            if hasattr(self, 'organize_var'):
                self.organize_var.set(config.organize_by_date)

        SettingsWindow(self.root, self.config, on_save)

    def show_info(self, title: str, message: str):
        if TTKBOOTSTRAP_AVAILABLE:
            Messagebox.show_info(message, title=title, parent=self.root)
        else:
            messagebox.showinfo(title, message)

    def show_error(self, message: str):
        if TTKBOOTSTRAP_AVAILABLE:
            Messagebox.show_error(message, title="Error", parent=self.root)
        else:
            messagebox.showerror("Error", message)

    def run(self):
        """Run the application."""
        self.root.mainloop()


def launch_gui():
    """Launch the GUI application."""
    app = EMLConverterApp()
    app.run()


if __name__ == '__main__':
    launch_gui()
