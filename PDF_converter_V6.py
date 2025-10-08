"""
PDF_converter.py
Improved GUI with:
- Dual mode: Enhance PDF (original features) and Compact Layout Maker (n-up)
- Live split preview (original vs processed)
- Presets, Apply-to-All, Auto-Optimize, Revert
- Compact PDF generator (2x2, 3x1, 3x2) with borders, margins, reading direction
- Quality slider + file size estimator (sampling)
- Pre-export verification modal with thumbnails
- Progress bar with thread-safe updates, keyboard shortcuts, and small UX polish
"""

import fitz  # PyMuPDF
from PIL import Image, ImageOps, ImageEnhance, ImageDraw, ImageFont, ImageTk, ImageChops
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import threading
import io
import os
import math
import traceback

# -----------------------------
# Constants & Utilities
# -----------------------------
DEFAULT_SETTINGS_FILE = "settings.json"
DEFAULT_DPI = 200  # used when rendering pages from PDF
ESTIMATE_SAMPLE_COUNT = 3  # pages to sample for size estimator

# paper sizes in mm
PAPER_SIZES_MM = {
    "A4": (210.0, 297.0),
    "Letter": (216.0, 279.0)
}

def mm_to_px(mm, dpi):
    return int(mm * dpi / 25.4)

def safe_action(func):
    """Decorator to run UI-affecting functions via tk.after when called from threads."""
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception:
            traceback.print_exc()
    return wrapper

# -----------------------------
# PDF Manager
# -----------------------------
class PDFManager:
    def __init__(self):
        self.doc = None
        self.original_pages = []  # list of PIL.Image (original)
        self.pages = []           # list of PIL.Image (processed)
        self.selected = []        # booleans

    def load_pdf(self, filepath, settings, progress_callback=None):
        """Load PDF and process each page using given settings.
           progress_callback: callable(float) where float in [0,1]"""
        self.doc = fitz.open(filepath)
        self.original_pages = []
        self.pages = []
        total = len(self.doc)
        for idx, p in enumerate(self.doc):
            pix = p.get_pixmap(dpi=DEFAULT_DPI)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.original_pages.append(img)
            self.pages.append(self.process_image(img, settings))
            if progress_callback:
                progress_callback((idx + 1) / total)
        self.selected = [True] * len(self.pages)

    def process_image(self, img, settings):
        """Apply invert -> grayscale optional -> contrast/brightness/sharpness."""
        # Keep identical logic to original for compatibility
        inverted = ImageOps.invert(img)
        processed = inverted.convert("L").convert("RGB") if settings.get("grayscale", True) else inverted
        processed = ImageEnhance.Contrast(processed).enhance(settings.get("contrast", 1.2))
        processed = ImageEnhance.Brightness(processed).enhance(settings.get("brightness", 1.0))
        processed = ImageEnhance.Sharpness(processed).enhance(settings.get("sharpness", 1.0))
        return processed

    def reprocess_current_page(self, index, settings):
        if 0 <= index < len(self.original_pages):
            self.pages[index] = self.process_image(self.original_pages[index], settings)

    def reprocess_all_pages(self, settings, progress_callback=None):
        total = len(self.original_pages)
        for idx, img in enumerate(self.original_pages):
            self.pages[idx] = self.process_image(img, settings)
            if progress_callback:
                progress_callback((idx + 1) / total)

    def insert_blank_page(self, index):
        if not self.pages:
            return
        blank = Image.new("RGB", self.pages[0].size, "white")
        self.original_pages.insert(index, blank.copy())
        self.pages.insert(index, blank)
        self.selected.insert(index, True)

    def insert_text_page(self, index, text):
        if not self.pages:
            return
        img = Image.new("RGB", self.pages[0].size, "white")
        draw = ImageDraw.Draw(img)
        font = ImageFont.load_default()
        draw.text((50, 50), text, fill="black", font=font)
        self.original_pages.insert(index, img.copy())
        self.pages.insert(index, img)
        self.selected.insert(index, True)

    def move_page(self, index, direction):
        new_index = index + direction
        if 0 <= index < len(self.pages) and 0 <= new_index < len(self.pages):
            for arr in [self.pages, self.original_pages, self.selected]:
                arr[index], arr[new_index] = arr[new_index], arr[index]

    def export_pdf(self, out_path, page_indices, progress_callback=None, quality=85):
        """Export processed pages specified by page_indices (list of indices) into a PDF.
           We compress each image to JPEG in-memory and then save combined PDF pages to keep size small."""
        export_pages = []
        total = len(page_indices)
        for i, idx in enumerate(page_indices):
            if self.selected[idx]:
                # reduce image to reasonable size for PDF export if necessary
                img = self.pages[idx]
                export_pages.append(img.convert("RGB"))
            if progress_callback:
                progress_callback((i + 1) / total)

        if not export_pages:
            raise Exception("No pages selected for export!")

        # Save images using PIL.save(..., format='PDF') - combine
        # To keep file size lower, we will convert images to JPEG bytes then read back into PIL for PDF save
        pil_pages_for_pdf = []
        for img in export_pages:
            # downscale moderately based on DPI used originally
            # For simplicity, save as is then use save_all to create PDF
            pil_pages_for_pdf.append(img)

        # save
        pil_pages_for_pdf[0].save(out_path, save_all=True, append_images=pil_pages_for_pdf[1:])

    def parse_ranges(self, s):
        """Parse a string like '1-4,6,8-9' into zero-based indices list."""
        result = []
        if not s or not s.strip():
            return list(range(len(self.pages)))
        parts = s.split(",")
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start, end = part.split("-")
                result.extend(range(int(start) - 1, int(end)))
            else:
                result.append(int(part) - 1)
        # clamp valid indices
        return [i for i in result if 0 <= i < len(self.pages)]

# -----------------------------
# Settings Save/Load & Presets
# -----------------------------
def save_settings(settings, filename=DEFAULT_SETTINGS_FILE):
    with open(filename, "w") as f:
        json.dump(settings, f, indent=2)

def load_settings(filename=DEFAULT_SETTINGS_FILE):
    try:
        with open(filename, "r") as f:
            return json.load(f)
    except Exception:
        # default
        return {
            "contrast": 1.2, "brightness": 1.0, "sharpness": 1.0, "grayscale": True,
            "last_folder": os.getcwd(),
            "presets": {
                "Print Clear": {"contrast": 1.3, "brightness": 1.05, "sharpness": 1.1, "grayscale": True},
                "Dark Notes Fix": {"contrast": 1.6, "brightness": 1.2, "sharpness": 1.0, "grayscale": True},
                "Read on Screen": {"contrast": 1.1, "brightness": 1.1, "sharpness": 1.0, "grayscale": False}
            },
            "last_preset": "Print Clear"
        }

# -----------------------------
# Main GUI
# -----------------------------
class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.manager = PDFManager()
        self.settings = load_settings()
        self.current_index = 0

        # State for layout maker
        self.layout_options = {
            "layout": tk.StringVar(value="2x2"),
            "paper": tk.StringVar(value="A4"),
            "orientation": tk.StringVar(value="Portrait"),
            "with_border": tk.BooleanVar(value=False),
            "outer_margin_mm": tk.DoubleVar(value=5.0),
            "inner_margin_mm": tk.DoubleVar(value=2.0),
            "reading_direction": tk.StringVar(value="Left to Right"),
            "quality": tk.IntVar(value=85)
        }

        self.root.title("PW Notes Converter")
        self.root.geometry("1200x760")
        self.root.minsize(1000, 650)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # top navigation for modes
        nav = tk.Frame(root, pady=6)
        nav.pack(fill="x")
        self.btn_enhance = ttk.Button(nav, text="Enhance PDF", command=self.show_enhance_mode)
        self.btn_layout = ttk.Button(nav, text="Compact Layout Maker", command=self.show_layout_mode)
        self.btn_enhance.pack(side="left", padx=6)
        self.btn_layout.pack(side="left")

        # status / progress
        status_frame = tk.Frame(root)
        status_frame.pack(fill="x", padx=10, pady=2)
        self.status_label = tk.Label(status_frame, text="Ready", anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=1.0)
        self.progress_bar.pack(side="right", fill="x", ipadx=150)

        # container for two modes
        self.container = tk.Frame(root)
        self.container.pack(fill="both", expand=True)

        # initialize frames
        self.enhance_frame = tk.Frame(self.container)
        self.layout_frame = tk.Frame(self.container)

        self.setup_enhance_frame()
        self.setup_layout_frame()

        # start with enhance mode
        self.show_enhance_mode()

        # keyboard shortcuts
        self.root.bind("<Left>", lambda e: self.change_page(-1))
        self.root.bind("<Right>", lambda e: self.change_page(1))
        self.root.bind("<Control-s>", lambda e: self.export_pdf())
        self.root.bind("<Control-o>", lambda e: self.open_pdf())

    # -----------------------------
    # Enhance Mode UI
    # -----------------------------
    def setup_enhance_frame(self):
        f = self.enhance_frame
        f.pack(fill="both", expand=True)

        # top file buttons
        file_frame = tk.Frame(f, pady=6)
        file_frame.pack(fill="x", padx=8)
        self.btn_open = ttk.Button(file_frame, text="Open PDF", command=self.open_pdf)
        self.btn_open.pack(side="left", padx=3)
        self.btn_batch = ttk.Button(file_frame, text="Batch Mode", command=self.open_batch)
        self.btn_batch.pack(side="left", padx=3)

        # Page indicator and quick actions
        info_frame = tk.Frame(f)
        info_frame.pack(fill="x", padx=8, pady=4)
        self.page_label = tk.Label(info_frame, text="Page 0 of 0")
        self.page_label.pack(side="left")
        ttk.Button(info_frame, text="Revert Page", command=self.revert_current_page).pack(side="left", padx=6)
        ttk.Button(info_frame, text="Apply to All Pages", command=self.apply_to_all_pages).pack(side="left", padx=6)
        ttk.Button(info_frame, text="Auto Optimize for Print", command=self.auto_optimize).pack(side="left", padx=6)

        # preview area (split view)
                # preview area (split view)
        preview_outer = tk.Frame(f)
        preview_outer.pack(fill="both", expand=True, padx=8, pady=6)
        preview_outer.columnconfigure(0, weight=1)
        preview_outer.columnconfigure(1, weight=1)
        preview_outer.rowconfigure(0, weight=1)

        left_box = tk.LabelFrame(preview_outer, text="Original")
        left_box.grid(row=0, column=0, sticky="nsew", padx=5, pady=2)
        left_box.rowconfigure(0, weight=1)
        left_box.columnconfigure(0, weight=1)
        self.canvas_left = tk.Canvas(left_box, bg="white")
        self.canvas_left.grid(row=0, column=0, sticky="nsew")

        right_box = tk.LabelFrame(preview_outer, text="Processed (Live)")
        right_box.grid(row=0, column=1, sticky="nsew", padx=5, pady=2)
        right_box.rowconfigure(0, weight=1)
        right_box.columnconfigure(0, weight=1)
        self.canvas_right = tk.Canvas(right_box, bg="white")
        self.canvas_right.grid(row=0, column=0, sticky="nsew")

        # page navigation
        nav = tk.Frame(f, pady=4)
        nav.pack()
        ttk.Button(nav, text="⏮ Prev", command=lambda: self.change_page(-1)).grid(row=0, column=0, padx=6)
        ttk.Button(nav, text="Next ⏭", command=lambda: self.change_page(1)).grid(row=0, column=1, padx=6)
        ttk.Button(nav, text="Insert Blank", command=self.insert_blank).grid(row=0, column=2, padx=6)
        ttk.Button(nav, text="Insert Text", command=self.insert_text).grid(row=0, column=3, padx=6)

        # settings / sliders / presets
        settings_frame = tk.LabelFrame(f, text="Image Settings (Live)", padx=8, pady=6)
        settings_frame.pack(fill="x", padx=8, pady=6)

        self.contrast = tk.DoubleVar(value=self.settings.get("contrast", 1.2))
        self.brightness = tk.DoubleVar(value=self.settings.get("brightness", 1.0))
        self.sharpness = tk.DoubleVar(value=self.settings.get("sharpness", 1.0))
        self.grayscale = tk.BooleanVar(value=self.settings.get("grayscale", True))

        left_col = tk.Frame(settings_frame)
        left_col.pack(side="left", fill="x", expand=True, padx=6)

        tk.Label(left_col, text="Contrast").pack(anchor="w")
        tk.Scale(left_col, from_=0.5, to=3.0, resolution=0.1, orient="horizontal",
                 variable=self.contrast, command=self.update_live_preview).pack(fill="x")
        tk.Label(left_col, text="Brightness").pack(anchor="w")
        tk.Scale(left_col, from_=0.5, to=3.0, resolution=0.1, orient="horizontal",
                 variable=self.brightness, command=self.update_live_preview).pack(fill="x")
        tk.Label(left_col, text="Sharpness").pack(anchor="w")
        tk.Scale(left_col, from_=0.5, to=3.0, resolution=0.1, orient="horizontal",
                 variable=self.sharpness, command=self.update_live_preview).pack(fill="x")
        tk.Checkbutton(left_col, text="Grayscale", variable=self.grayscale,
                       command=self.update_live_preview).pack(anchor="w", pady=4)

        # presets
        right_col = tk.Frame(settings_frame)
        right_col.pack(side="right", padx=6)
        tk.Label(right_col, text="Presets").pack(anchor="w")
        self.preset_names = list(self.settings.get("presets", {}).keys())
        self.preset_var = tk.StringVar(value=self.settings.get("last_preset", (self.preset_names[0] if self.preset_names else "")))
        self.preset_menu = ttk.Combobox(right_col, values=self.preset_names, state="readonly",
                                        textvariable=self.preset_var)
        self.preset_menu.pack()
        ttk.Button(right_col, text="Load Preset", command=self.load_preset).pack(pady=4)
        ttk.Button(right_col, text="Save Current as Preset", command=self.save_current_as_preset).pack(pady=4)
        ttk.Button(right_col, text="Revert All (reload originals)", command=self.revert_all).pack(pady=4)

        # export area
        export_frame = tk.LabelFrame(f, text="Export", padx=8, pady=6)
        export_frame.pack(fill="x", padx=8, pady=6)
        tk.Label(export_frame, text="Page ranges (e.g. 1-5,7,10-12):").grid(row=0, column=0, sticky="w")
        self.range_entry = tk.Entry(export_frame)
        self.range_entry.grid(row=0, column=1, sticky="ew", padx=6)
        export_frame.columnconfigure(1, weight=1)
        ttk.Button(export_frame, text="Preview All Pages before Export", command=self.preview_all_before_export).grid(row=0, column=2, padx=6)
        ttk.Button(export_frame, text="Export PDF", command=self.export_pdf).grid(row=1, column=1, pady=6)

        # tooltips: lightweight
        self._add_tooltip(self.btn_open, "Ctrl+O — Open PDF")
        self._add_tooltip(self.btn_batch, "Batch convert PDFs")
        self._add_tooltip(export_frame, "Export uses processed pages")

        # hide until shown by mode switch
        f.pack_forget()

    # -----------------------------
    # Layout Maker UI
    # -----------------------------
    def setup_layout_frame(self):
        f = self.layout_frame
        f.pack(fill="both", expand=True)

        # top: options
        opts = tk.Frame(f, pady=6)
        opts.pack(fill="x", padx=8)

        # layout type
        tk.Label(opts, text="Pages per sheet:").grid(row=0, column=0, sticky="w")
        layout_menu = ttk.Combobox(opts, values=["2x2", "3x1", "3x2"], state="readonly",
                                   textvariable=self.layout_options["layout"], width=8)
        layout_menu.grid(row=0, column=1, padx=6)

        tk.Checkbutton(opts, text="Add border", variable=self.layout_options["with_border"]).grid(row=0, column=2, padx=6)

        # page size and orientation
        tk.Label(opts, text="Paper:").grid(row=1, column=0, sticky="w")
        paper_menu = ttk.Combobox(opts, values=list(PAPER_SIZES_MM.keys()), state="readonly",
                                  textvariable=self.layout_options["paper"], width=8)
        paper_menu.grid(row=1, column=1, padx=6)
        orient_menu = ttk.Combobox(opts, values=["Portrait", "Landscape"], state="readonly",
                                   textvariable=self.layout_options["orientation"], width=10)
        orient_menu.grid(row=1, column=2, padx=6)

        # reading direction
        tk.Label(opts, text="Reading direction:").grid(row=2, column=0, sticky="w")
        rd = ttk.Combobox(opts, values=["Left to Right", "Top to Bottom"], state="readonly",
                          textvariable=self.layout_options["reading_direction"], width=15)
        rd.grid(row=2, column=1, padx=6)

        # margins
        margin_frame = tk.Frame(f)
        margin_frame.pack(fill="x", padx=8, pady=6)
        tk.Label(margin_frame, text="Outer margin (mm)").grid(row=0, column=0, sticky="w")
        tk.Scale(margin_frame, from_=0, to=30, orient="horizontal", variable=self.layout_options["outer_margin_mm"]).grid(row=0, column=1, sticky="ew")
        tk.Label(margin_frame, text="Inner gap (mm)").grid(row=1, column=0, sticky="w")
        tk.Scale(margin_frame, from_=0, to=20, orient="horizontal", variable=self.layout_options["inner_margin_mm"]).grid(row=1, column=1, sticky="ew")

        # quality and estimator
        qframe = tk.Frame(f, pady=6)
        qframe.pack(fill="x", padx=8)
        tk.Label(qframe, text="JPEG Quality:").grid(row=0, column=0, sticky="w")
        ttk.Scale(qframe, from_=50, to=100, variable=self.layout_options["quality"], orient="horizontal").grid(row=0, column=1, sticky="ew")
        ttk.Button(qframe, text="Estimate File Size", command=self.estimate_compact_size).grid(row=0, column=2, padx=6)
        self.estimate_label = tk.Label(qframe, text="Estimated: N/A")
        self.estimate_label.grid(row=0, column=3, padx=6)

        # actions
        actions = tk.Frame(f, pady=6)
        actions.pack(fill="x", padx=8)
        ttk.Button(actions, text="Generate Compact PDF", command=self.generate_compact_pdf).pack(side="left", padx=6)
        ttk.Button(actions, text="Preview Compact Pages", command=self.preview_compact_before_export).pack(side="left", padx=6)

        # hide until shown
        f.pack_forget()

    # -----------------------------
    # Mode switching
    # -----------------------------
    def show_enhance_mode(self):
        self.layout_frame.pack_forget()
        self.enhance_frame.pack(fill="both", expand=True)
        self.btn_enhance.state(["disabled"])
        self.btn_layout.state(["!disabled"])
        self.status_label.config(text="Enhance mode")

    def show_layout_mode(self):
        self.enhance_frame.pack_forget()
        self.layout_frame.pack(fill="both", expand=True)
        self.btn_layout.state(["disabled"])
        self.btn_enhance.state(["!disabled"])
        self.status_label.config(text="Layout mode")

    # -----------------------------
    # File operations & threading helpers
    # -----------------------------
    def open_pdf(self):
        initial = self.settings.get("last_folder", os.getcwd())
        path = filedialog.askopenfilename(initialdir=initial, filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        # remember folder
        self.settings["last_folder"] = os.path.dirname(path)
        # load in background
        threading.Thread(target=self._load_pdf_thread, args=(path,), daemon=True).start()

    def _run_in_ui_thread(self, func, *args, **kwargs):
        self.root.after(0, lambda: func(*args, **kwargs))

    def _load_pdf_thread(self, path):
        try:
            self._set_progress(0, "Loading PDF...")
            settings = self.get_settings()
            self.manager.load_pdf(path, settings, progress_callback=self._thread_progress_callback)
            self.current_index = 0
            self._run_in_ui_thread(self.update_previews)
            self._set_progress(1, "Loaded")
            # small pause then reset progress
            self._run_in_ui_thread(lambda: self.root.after(300, lambda: self._set_progress(0, "Ready")))
        except Exception as e:
            self._run_in_ui_thread(lambda: messagebox.showerror("Error", f"Failed to load PDF:\n{e}"))
            self._set_progress(0, "Error")
            traceback.print_exc()

    def open_batch(self):
        initial = self.settings.get("last_folder", os.getcwd())
        paths = filedialog.askopenfilenames(initialdir=initial, filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        self.settings["last_folder"] = os.path.dirname(paths[0])
        threading.Thread(target=self._batch_thread, args=(paths,), daemon=True).start()

    def _batch_thread(self, paths):
        try:
            for path in paths:
                self._set_progress(0, f"Processing {os.path.basename(path)}")
                self.manager.load_pdf(path, self.get_settings(), progress_callback=self._thread_progress_callback)
                out = path.replace(".pdf", "_converted.pdf")
                # direct export using manager - uses processed images
                self.manager.export_pdf(out, range(len(self.manager.pages)), progress_callback=self._thread_progress_callback)
            self._set_progress(0, "Batch done")
            self._run_in_ui_thread(lambda: messagebox.showinfo("Done", "Batch conversion completed!"))
        except Exception as e:
            traceback.print_exc()
            self._run_in_ui_thread(lambda: messagebox.showerror("Error", str(e)))
            self._set_progress(0, "Error")

    def _export_thread(self, out_path, page_indices, quality):
        try:
            self._set_progress(0, "Exporting...")
            # Export uses manager.export_pdf (already uses processed images)
            self.manager.export_pdf(out_path, page_indices, progress_callback=self._thread_progress_callback, quality=quality)
            self._set_progress(1, "Saved")
            self._run_in_ui_thread(lambda: messagebox.showinfo("Export", f"Saved to {out_path}"))
            self._run_in_ui_thread(lambda: self.root.after(400, lambda: self._set_progress(0, "Ready")))
        except Exception as e:
            traceback.print_exc()
            self._run_in_ui_thread(lambda: messagebox.showerror("Export Error", str(e)))
            self._set_progress(0, "Error")

    def _thread_progress_callback(self, value):
        # called from background thread; update progress via UI thread
        self._run_in_ui_thread(lambda v=value: self._set_progress(v, f"Processing: {int(v*100)}%"))

    def _set_progress(self, value, status_text=None):
        self.progress_var.set(value)
        if status_text:
            self.status_label.config(text=status_text)

    # -----------------------------
    # Preview & Page controls
    # -----------------------------
    def update_previews(self):
        """Update both canvases with current page images (original & processed)."""
        if not self.manager.pages:
            self.page_label.config(text="Page 0 of 0")
            self.canvas_left.delete("all")
            self.canvas_right.delete("all")
            return
        total = len(self.manager.pages)
        self.page_label.config(text=f"Page {self.current_index + 1} of {total}")

        orig = self.manager.original_pages[self.current_index].copy()
        proc = self.manager.pages[self.current_index].copy()

        # compute available size
        lw = max(200, self.canvas_left.winfo_width())
        lh = max(150, self.canvas_left.winfo_height())
        rw = max(200, self.canvas_right.winfo_width())
        rh = max(150, self.canvas_right.winfo_height())

        # thumbnail both to fit their canvases maintaining ratio
        orig.thumbnail((lw - 20, lh - 20), Image.LANCZOS)
        proc.thumbnail((rw - 20, rh - 20), Image.LANCZOS)

        self.tk_orig = ImageTk.PhotoImage(orig)
        self.tk_proc = ImageTk.PhotoImage(proc)

        self.canvas_left.delete("all")
        self.canvas_right.delete("all")

        self.canvas_left.create_image(lw // 2, lh // 2, image=self.tk_orig, anchor="center")
        self.canvas_right.create_image(rw // 2, rh // 2, image=self.tk_proc, anchor="center")

    def update_live_preview(self, *_):
        """Reprocess only current page and update preview (live when sliders change)."""
        if not self.manager.pages:
            return
        self.manager.reprocess_current_page(self.current_index, self.get_settings())
        self.update_previews()

    def change_page(self, delta):
        if not self.manager.pages:
            return
        self.current_index = (self.current_index + delta) % len(self.manager.pages)
        self.update_previews()

    def toggle_page_keep(self):
        if not self.manager.pages:
            return
        self.manager.selected[self.current_index] = not self.manager.selected[self.current_index]
        self.update_previews()

    def insert_blank(self):
        if not self.manager.pages:
            return
        self.manager.insert_blank_page(self.current_index + 1)
        self.update_previews()

    def insert_text(self):
        if not self.manager.pages:
            return
        text = tk.simpledialog.askstring("Text Page", "Enter text:")
        if text:
            self.manager.insert_text_page(self.current_index + 1, text)
            self.update_previews()

    def move_page(self, direction):
        if not self.manager.pages:
            return
        self.manager.move_page(self.current_index, direction)
        self.update_previews()

    def revert_current_page(self):
        """Restore processed page to match original (without altering original image)."""
        if not self.manager.pages:
            return
        # simply copy original to processed
        self.manager.pages[self.current_index] = self.manager.original_pages[self.current_index].copy()
        self.update_previews()
        messagebox.showinfo("Reverted", f"Page {self.current_index + 1} reverted to original.")

    def revert_all(self):
        """Reset all processed pages to original images (keeps originals unchanged)."""
        if not self.manager.pages:
            return
        for i in range(len(self.manager.pages)):
            self.manager.pages[i] = self.manager.original_pages[i].copy()
        self.update_previews()
        messagebox.showinfo("Reverted", "All pages restored to original images.")

    # -----------------------------
    # Settings & Presets 
    # -----------------------------
    def get_settings(self):
        return {
            "contrast": self.contrast.get(),
            "brightness": self.brightness.get(),
            "sharpness": self.sharpness.get(),
            "grayscale": self.grayscale.get()
        }

    def save_settings(self):
        # persist current sliders and preset list + last folder
        self.settings.update(self.get_settings())
        self.settings["last_preset"] = self.preset_var.get()
        save_settings(self.settings)
        messagebox.showinfo("Settings", "Saved!")

    def auto_optimize(self):
        preset = {"contrast": 1.3, "brightness": 1.05, "sharpness": 1.1, "grayscale": True}
        self.contrast.set(preset["contrast"])
        self.brightness.set(preset["brightness"])
        self.sharpness.set(preset["sharpness"])
        self.grayscale.set(preset["grayscale"])
        self.update_live_preview()

    def apply_to_all_pages(self):
        """Reprocess all pages with current settings and update progress bar."""
        if not self.manager.pages:
            messagebox.showinfo("No pages", "No pages loaded.")
            return

        def worker():
            try:
                self._set_progress(0, "Reprocessing all pages...")
                self.manager.reprocess_all_pages(self.get_settings(), progress_callback=self._thread_progress_callback)
                self._run_in_ui_thread(self.update_previews)
                self._set_progress(1, "Done")
                self._run_in_ui_thread(lambda: self.root.after(400, lambda: self._set_progress(0, "Ready")))
                self._run_in_ui_thread(lambda: messagebox.showinfo("Reprocess", "All pages updated successfully."))
            except Exception as e:
                self._run_in_ui_thread(lambda: messagebox.showerror("Error", str(e)))
                self._set_progress(0, "Error")

        threading.Thread(target=worker, daemon=True).start()

    def load_preset(self):
        key = self.preset_var.get()
        if not key:
            return
        preset = self.settings.get("presets", {}).get(key)
        if not preset:
            messagebox.showerror("Preset", "Preset not found.")
            return
        self.contrast.set(preset.get("contrast", 1.2))
        self.brightness.set(preset.get("brightness", 1.0))
        self.sharpness.set(preset.get("sharpness", 1.0))
        self.grayscale.set(preset.get("grayscale", True))
        self.update_live_preview()
        self.settings["last_preset"] = key

    def save_current_as_preset(self):
        name = tk.simpledialog.askstring("Preset name", "Enter a name for this preset:")
        if not name:
            return
        ps = self.settings.get("presets", {})
        ps[name] = self.get_settings()
        self.settings["presets"] = ps
        self.preset_names = list(ps.keys())
        self.preset_menu['values'] = self.preset_names
        messagebox.showinfo("Preset", f"Saved preset '{name}'")


    # -----------------------------
    # Export / Preview All / Pre-Export Verification
    # -----------------------------
    def preview_all_before_export(self):
        """Show a modal with thumbnails of all processed pages for user confirmation."""
        if not self.manager.pages:
            messagebox.showinfo("No pages", "No pages loaded.")
            return
        # open modal
        top = tk.Toplevel(self.root)
        top.title("Preview All Pages")
        top.geometry("800x600")
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(frame)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor="nw")

        thumbs = []
        for i, img in enumerate(self.manager.pages):
            thumb = img.copy()
            thumb.thumbnail((240, 170), Image.LANCZOS)
            tkimg = ImageTk.PhotoImage(thumb)
            lbl = tk.Label(inner, image=tkimg)
            lbl.image = tkimg
            lbl.grid(row=i//3, column=i%3, padx=6, pady=6)
            tk.Label(inner, text=f"{i+1} {'(Kept)' if self.manager.selected[i] else '(Removed)'}").grid(row=(i//3)+1000, column=i%3)  # small hack: not needed but keeps spacing
            thumbs.append(lbl)

        inner.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

        btn_frame = tk.Frame(top)
        btn_frame.pack(fill="x", pady=6)
        ttk.Button(btn_frame, text="Reprocess all then preview", command=lambda: [top.destroy(), self.apply_to_all_pages(), self.preview_all_before_export()]).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Confirm & Export", command=lambda: [top.destroy(), self.export_pdf()]).pack(side="right", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=top.destroy).pack(side="right", padx=6)

    def export_pdf(self):
        if not self.manager.pages:
            messagebox.showinfo("No pages", "No pages loaded.")
            return
        # ask default filename
        initial = self.settings.get("last_folder", os.getcwd())
        out = filedialog.asksaveasfilename(initialdir=initial, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not out:
            return
        # confirm overwrite
        if os.path.exists(out):
            if not messagebox.askyesno("Overwrite", f"{out} exists. Overwrite?"):
                return
        self.settings["last_folder"] = os.path.dirname(out)
        page_range_text = self.range_entry.get()
        page_indices = self.manager.parse_ranges(page_range_text) if page_range_text else list(range(len(self.manager.pages)))
        # Launch background export thread
        quality = 85  # default for regular export (we use quality only for compact generator)
        threading.Thread(target=self._export_thread, args=(out, page_indices, quality), daemon=True).start()

    # -----------------------------
    # Compact Layout Generator: estimator + generator + preview
    # -----------------------------
    def estimate_compact_size(self):
        """Estimate final compact PDF size by sampling a few pages compressed to chosen quality."""
        if not self.manager.pages:
            self.estimate_label.config(text="Estimated: N/A (no pages)")
            return

        # sample min(ESTIMATE_SAMPLE_COUNT, total) pages evenly across document
        total_pages = len(self.manager.pages)
        count = min(ESTIMATE_SAMPLE_COUNT, total_pages)
        indices = [math.floor(i * total_pages / count) for i in range(count)]
        q = self.layout_options["quality"].get()
        # compress each sample to JPEG bytes with size similar to what generator will create
        sizes = []
        for idx in indices:
            img = self.manager.pages[idx]
            # create a cell image (resized) similar to what will be placed on compact page
            cell = self._make_cell_image_for_estimate(img)
            buf = io.BytesIO()
            try:
                cell.save(buf, format="JPEG", quality=q)
                sizes.append(buf.tell())
            finally:
                buf.close()
        if not sizes:
            self.estimate_label.config(text="Estimated: N/A")
            return
        avg = sum(sizes) / len(sizes)  # bytes per cell
        layout = self.layout_options["layout"].get()
        cells_per_sheet = 4 if layout == "2x2" else (3 if layout == "3x1" else 6 if layout == "3x2" else 1)
        # estimate pages after packing
        sheets = math.ceil(total_pages / cells_per_sheet)
        # estimated total bytes = avg * cells_per_sheet * sheets (rough) + PDF overhead
        estimated_bytes = avg * total_pages * 1.15  # multiplier for overhead
        estimated_mb = estimated_bytes / (1024 * 1024)
        self.estimate_label.config(text=f"Estimated: ~{estimated_mb:.2f} MB")

    def _make_cell_image_for_estimate(self, img):
        """Return resized cell image for estimator based on paper and layout."""
        # compute cell size in px based on paper and layout using chosen DPI
        paper = self.layout_options["paper"].get()
        orientation = self.layout_options["orientation"].get()
        dpi = DEFAULT_DPI
        w_mm, h_mm = PAPER_SIZES_MM.get(paper, PAPER_SIZES_MM["A4"])
        if orientation == "Landscape":
            w_mm, h_mm = h_mm, w_mm
        # layout grid size
        layout = self.layout_options["layout"].get()
        if layout == "2x2":
            cols, rows = 2, 2
        elif layout == "3x1":
            cols, rows = 1, 3
        elif layout == "3x2":
            cols, rows = 2, 3
        else:
            cols, rows = 1, 1

        outer_mm = self.layout_options["outer_margin_mm"].get()
        inner_mm = self.layout_options["inner_margin_mm"].get()
        # compute sheet px
        sheet_w = mm_to_px(w_mm, dpi)
        sheet_h = mm_to_px(h_mm, dpi)
        # compute cell area
        usable_w = sheet_w - mm_to_px(outer_mm*2, dpi) - mm_to_px(inner_mm*(cols - 1), dpi)
        usable_h = sheet_h - mm_to_px(outer_mm*2, dpi) - mm_to_px(inner_mm*(rows - 1), dpi)
        cell_w = max(10, usable_w // cols)
        cell_h = max(10, usable_h // rows)
        cell = img.copy()
        cell.thumbnail((cell_w, cell_h), Image.LANCZOS)
        return cell

    def generate_compact_pdf(self):
        """Generate the compact PDF according to layout options. Uses progress bar and runs in separate thread."""
        if not self.manager.pages:
            messagebox.showinfo("No pages", "No pages loaded.")
            return

        initial = self.settings.get("last_folder", os.getcwd())
        out = filedialog.asksaveasfilename(initialdir=initial, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not out:
            return
        if os.path.exists(out):
            if not messagebox.askyesno("Overwrite", f"{out} exists. Overwrite?"):
                return
        self.settings["last_folder"] = os.path.dirname(out)

        # launch background generator
        threading.Thread(target=self._generate_compact_thread, args=(out,), daemon=True).start()

    def _generate_compact_thread(self, out_path):
        try:
            self._set_progress(0, "Generating compact PDF...")
            # build list of sheet images
            layout = self.layout_options["layout"].get()
            if layout == "2x2":
                cols, rows = 2, 2
            elif layout == "3x1":
                cols, rows = 1, 3
            elif layout == "3x2":
                cols, rows = 2, 3
            else:
                cols, rows = 2, 2

            dpi = DEFAULT_DPI
            paper = self.layout_options["paper"].get()
            orientation = self.layout_options["orientation"].get()
            w_mm, h_mm = PAPER_SIZES_MM.get(paper, PAPER_SIZES_MM["A4"])
            if orientation == "Landscape":
                w_mm, h_mm = h_mm, w_mm
            sheet_w = mm_to_px(w_mm, dpi)
            sheet_h = mm_to_px(h_mm, dpi)

            outer_mm = self.layout_options["outer_margin_mm"].get()
            inner_mm = self.layout_options["inner_margin_mm"].get()
            outer_px = mm_to_px(outer_mm, dpi)
            inner_px = mm_to_px(inner_mm, dpi)

            # compute cell sizes
            usable_w = sheet_w - outer_px*2 - inner_px*(cols - 1)
            usable_h = sheet_h - outer_px*2 - inner_px*(rows - 1)
            cell_w = max(50, usable_w // cols)
            cell_h = max(50, usable_h // rows)

            # reading order
            left_to_right = (self.layout_options["reading_direction"].get() == "Left to Right")
            with_border = self.layout_options["with_border"].get()
            quality = int(self.layout_options["quality"].get())

            total_pages = len(self.manager.pages)
            cells_per_sheet = cols * rows
            sheets = math.ceil(total_pages / cells_per_sheet)
            sheet_images = []

            for s in range(sheets):
                # create a blank A4 sheet (portrait or landscape)
                sheet_img = Image.new("RGB", (sheet_w, sheet_h), "white")
                draw = ImageDraw.Draw(sheet_img)

                # place images (2×2, 3×1, etc.)
                for cell_idx in range(cells_per_sheet):
                    page_idx = s * cells_per_sheet + cell_idx
                    if page_idx >= total_pages:
                        break

                    # reading order
                    if left_to_right:
                        r = cell_idx // cols
                        c = cell_idx % cols
                    else:
                        c = cell_idx // rows
                        r = cell_idx % rows

                    x = outer_px + c * (cell_w + inner_px)
                    y = outer_px + r * (cell_h + inner_px)
                    img = self.manager.pages[page_idx].copy()
                    img.thumbnail((cell_w, cell_h), Image.LANCZOS)
                    paste_x = x + (cell_w - img.width) // 2
                    paste_y = y + (cell_h - img.height) // 2
                    sheet_img.paste(img, (paste_x, paste_y))
                    if with_border:
                        draw.rectangle([x, y, x + cell_w - 1, y + cell_h - 1], outline="black", width=1)

                    # progress update
                    overall_idx = page_idx
                    self._run_in_ui_thread(lambda v=(overall_idx + 1) / total_pages:
                                           self._set_progress(v, f"Placing page {overall_idx + 1}/{total_pages}"))

                # ✅ Rotate entire A4 sheet for landscape output (true physical rotation)
                if self.layout_options["orientation"].get() == "Landscape":
                    sheet_img = sheet_img.transpose(Image.Transpose.ROTATE_270)

                sheet_images.append(sheet_img)

            # save sheet images as PDF
            # To control file size, we save each sheet as JPEG into BytesIO and then create PIL images from bytes for PDF save
            pil_pages_for_pdf = []
            temp_bytes = []
            for idx, sheet in enumerate(sheet_images):
                buf = io.BytesIO()
                sheet.save(buf, format="JPEG", quality=quality)
                temp_bytes.append(buf.getvalue())
                buf.close()
                # reopen into PIL to ensure compactness
                pil_pages_for_pdf.append(Image.open(io.BytesIO(temp_bytes[-1])).convert("RGB"))
                # small progress update
                self._run_in_ui_thread(lambda v=(idx + 1) / len(sheet_images): self._set_progress(v * 0.9, f"Preparing sheet {idx+1}/{len(sheet_images)}"))

            if not pil_pages_for_pdf:
                raise Exception("Nothing to save for compact PDF")

            # final save
            pil_pages_for_pdf[0].save(out_path, save_all=True, append_images=pil_pages_for_pdf[1:])
            self._set_progress(1, "Saved")
            self._run_in_ui_thread(lambda: messagebox.showinfo("Done", f"Saved compact PDF to {out_path}"))
            self._run_in_ui_thread(lambda: self.root.after(400, lambda: self._set_progress(0, "Ready")))
        except Exception as e:
            traceback.print_exc()
            self._run_in_ui_thread(lambda: messagebox.showerror("Error", str(e)))
            self._set_progress(0, "Error")

    def preview_compact_before_export(self):
        """Preview what the first few sheets will look like in a modal."""
        if not self.manager.pages:
            messagebox.showinfo("No pages", "No pages loaded.")
            return
        # Create a small set of sheet images in-memory (first 2 sheets) and show thumbnails
        try:
            sheet_images = self._compose_sheets_preview(count=2)
        except Exception as e:
            messagebox.showerror("Preview Error", str(e))
            return
        top = tk.Toplevel(self.root)
        top.title("Compact Preview")
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True)
        for i, sh in enumerate(sheet_images):
            img = sh.copy()
            img.thumbnail((600, 800), Image.LANCZOS)
            tkimg = ImageTk.PhotoImage(img)
            lbl = tk.Label(frame, image=tkimg)
            lbl.image = tkimg
            lbl.pack(padx=6, pady=6)

    def _compose_sheets_preview(self, count=2):
        """Return in-memory sheet PIL images for preview (small number of sheets)."""
        layout = self.layout_options["layout"].get()
        if layout == "2x2":
            cols, rows = 2, 2
        elif layout == "3x1":
            cols, rows = 1, 3
        elif layout == "3x2":
            cols, rows = 2, 3
        else:
            cols, rows = 2, 2

        dpi = DEFAULT_DPI
        paper = self.layout_options["paper"].get()
        orientation = self.layout_options["orientation"].get()
        w_mm, h_mm = PAPER_SIZES_MM.get(paper, PAPER_SIZES_MM["A4"])
        if orientation == "Landscape":
            w_mm, h_mm = h_mm, w_mm
        sheet_w = mm_to_px(w_mm, dpi)
        sheet_h = mm_to_px(h_mm, dpi)

        outer_mm = self.layout_options["outer_margin_mm"].get()
        inner_mm = self.layout_options["inner_margin_mm"].get()
        outer_px = mm_to_px(outer_mm, dpi)
        inner_px = mm_to_px(inner_mm, dpi)

        usable_w = sheet_w - outer_px*2 - inner_px*(cols - 1)
        usable_h = sheet_h - outer_px*2 - inner_px*(rows - 1)
        cell_w = max(50, usable_w // cols)
        cell_h = max(50, usable_h // rows)

        left_to_right = (self.layout_options["reading_direction"].get() == "Left to Right")
        with_border = self.layout_options["with_border"].get()

        total_pages = len(self.manager.pages)
        cells_per_sheet = cols * rows
        sheets = math.ceil(total_pages / cells_per_sheet)
        sheet_images = []
        
        for s in range(min(sheets, count)):
            # create a blank A4 sheet
            sheet_img = Image.new("RGB", (sheet_w, sheet_h), "white")
            draw = ImageDraw.Draw(sheet_img)

            for cell_idx in range(cells_per_sheet):
                page_idx = s * cells_per_sheet + cell_idx
                if page_idx >= total_pages:
                    break

                if left_to_right:
                    r = cell_idx // cols
                    c = cell_idx % cols
                else:
                    c = cell_idx // rows
                    r = cell_idx % rows

                x = outer_px + c * (cell_w + inner_px)
                y = outer_px + r * (cell_h + inner_px)
                img = self.manager.pages[page_idx].copy()
                img.thumbnail((cell_w, cell_h), Image.LANCZOS)
                paste_x = x + (cell_w - img.width) // 2
                paste_y = y + (cell_h - img.height) // 2
                sheet_img.paste(img, (paste_x, paste_y))
                if with_border:
                    draw.rectangle([x, y, x + cell_w - 1, y + cell_h - 1], outline="black", width=1)

                    # ✅ Rotate for landscape preview (apply rotation after full layout)
                if self.layout_options["orientation"].get() == "Landscape":
                    # Rotate after all images have been placed, not before
                    sheet_img = sheet_img.transpose(Image.Transpose.ROTATE_270)

                    # Swap width/height since image was rotated
                   # sheet_w, sheet_h = sheet_img.size

                sheet_images.append(sheet_img)

            # ✅ Return the list of composed sheet previews
            return sheet_images


    # -----------------------------
    # Window close & autosave
    # -----------------------------
    def on_close(self):
        # autosave current settings
        self.settings.update(self.get_settings())
        self.settings["last_preset"] = self.preset_var.get()
        save_settings(self.settings)
        self.root.destroy()

    # -----------------------------
    # Small helpers
    # -----------------------------
    def _add_tooltip(self, widget, text):
        # minimal tooltip implementation
        def enter(e):
            self._tooltip = tk.Toplevel(self.root)
            self._tooltip.wm_overrideredirect(True)
            x = e.x_root + 10
            y = e.y_root + 10
            self._tooltip.wm_geometry(f"+{x}+{y}")
            tk.Label(self._tooltip, text=text, background="yellow").pack()
        def leave(e):
            if hasattr(self, "_tooltip") and self._tooltip:
                self._tooltip.destroy()
                self._tooltip = None
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop()
