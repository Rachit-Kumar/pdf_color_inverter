import fitz  # PyMuPDF
from PIL import Image, ImageOps, ImageEnhance, ImageDraw, ImageFont, ImageTk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import threading
import time


# -----------------------------
# PDF Manager
# -----------------------------
class PDFManager:
    def __init__(self):
        self.doc = None
        self.original_pages = []  # unprocessed pages
        self.pages = []           # processed pages
        self.selected = []

    def load_pdf(self, filepath, settings, progress_callback=None):
        self.doc = fitz.open(filepath)
        self.original_pages = []
        self.pages = []

        total = len(self.doc)
        for idx, p in enumerate(self.doc):
            pix = p.get_pixmap(dpi=200)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.original_pages.append(img)
            self.pages.append(self.process_image(img, settings))
            if progress_callback:
                progress_callback((idx + 1) / total)
        self.selected = [True] * len(self.pages)

    def process_image(self, img, settings):
        inverted = ImageOps.invert(img)
        processed = inverted.convert("L").convert("RGB") if settings["grayscale"] else inverted
        processed = ImageEnhance.Contrast(processed).enhance(settings["contrast"])
        processed = ImageEnhance.Brightness(processed).enhance(settings["brightness"])
        processed = ImageEnhance.Sharpness(processed).enhance(settings["sharpness"])
        return processed

    def reprocess_current_page(self, index, settings):
        if 0 <= index < len(self.original_pages):
            img = self.original_pages[index]
            self.pages[index] = self.process_image(img, settings)

    def reprocess_all_pages(self, settings, progress_callback=None):
        total = len(self.original_pages)
        for idx, img in enumerate(self.original_pages):
            self.pages[idx] = self.process_image(img, settings)
            if progress_callback:
                progress_callback((idx + 1) / total)

    def insert_blank_page(self, index):
        blank = Image.new("RGB", self.pages[0].size, "white")
        self.original_pages.insert(index, blank.copy())
        self.pages.insert(index, blank)
        self.selected.insert(index, True)

    def insert_text_page(self, index, text):
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

    def export_pdf(self, out_path, settings, page_range=None, progress_callback=None):
        export_pages = []
        indices = self.parse_ranges(page_range) if page_range else range(len(self.pages))
        total = len(indices)
        for i, idx in enumerate(indices):
            if self.selected[idx]:
                export_pages.append(self.pages[idx])
            if progress_callback:
                progress_callback((i + 1) / total)

        if not export_pages:
            raise Exception("No pages selected for export!")

        export_pages[0].save(out_path, save_all=True, append_images=export_pages[1:])

    def parse_ranges(self, s):
        result = []
        parts = s.split(",")
        for part in parts:
            if "-" in part:
                start, end = part.split("-")
                result.extend(range(int(start) - 1, int(end)))
            else:
                result.append(int(part) - 1)
        return result


# -----------------------------
# Settings Save/Load
# -----------------------------
def save_settings(settings, filename="settings.json"):
    with open(filename, "w") as f:
        json.dump(settings, f)


def load_settings(filename="settings.json"):
    try:
        with open(filename, "r") as f:
            return json.load(f)
    except:
        return {"contrast": 2.4, "brightness": 1.9, "sharpness": 1.6, "grayscale": True}


# -----------------------------
# GUI
# -----------------------------
class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.manager = PDFManager()
        self.settings = load_settings()
        self.current_index = 0

        self.root.title("PW Notes Converter - Split View + Live Tools")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 600)

        # Layout config
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # --- Progress Bar ---
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(root, variable=self.progress_var, maximum=1)
        self.progress_bar.grid(row=5, column=0, sticky="ew", padx=10, pady=2)

        # --- File Controls ---
        top_frame = tk.Frame(root, padx=10, pady=5)
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure((0, 1), weight=1)

        tk.Button(top_frame, text="📂 Open PDF", command=self.open_pdf, width=15).grid(row=0, column=0, padx=5)
        tk.Button(top_frame, text="🗂 Batch Mode", command=self.open_batch, width=15).grid(row=0, column=1, padx=5)

        # --- Split Preview (Left = Original, Right = Processed) ---
        preview_frame = tk.Frame(root)
        preview_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        preview_frame.columnconfigure((0, 1), weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Left Canvas (Original)
        left_frame = tk.LabelFrame(preview_frame, text="Original Page")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        self.canvas_left = tk.Canvas(left_frame, bg="lightgray")
        self.canvas_left.pack(fill="both", expand=True)

        # Right Canvas (Processed)
        right_frame = tk.LabelFrame(preview_frame, text="Processed Preview")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        self.canvas_right = tk.Canvas(right_frame, bg="lightgray")
        self.canvas_right.pack(fill="both", expand=True)

        # --- Navigation ---
        nav_frame = tk.Frame(root, pady=5)
        nav_frame.grid(row=2, column=0)
        tk.Button(nav_frame, text="⏮ Prev", command=lambda: self.change_page(-1)).grid(row=0, column=0, padx=10)
        tk.Button(nav_frame, text="Next ⏭", command=lambda: self.change_page(1)).grid(row=0, column=1, padx=10)

        # --- Settings ---
        settings_frame = tk.LabelFrame(root, text="Image Settings (Live)", padx=10, pady=5)
        settings_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        settings_frame.columnconfigure(0, weight=1)

        self.contrast = tk.DoubleVar(value=self.settings["contrast"])
        self.brightness = tk.DoubleVar(value=self.settings["brightness"])
        self.sharpness = tk.DoubleVar(value=self.settings["sharpness"])
        self.grayscale = tk.BooleanVar(value=self.settings["grayscale"])

        for label, var in [("Contrast", self.contrast),
                           ("Brightness", self.brightness),
                           ("Sharpness", self.sharpness)]:
            tk.Label(settings_frame, text=label).pack()
            tk.Scale(settings_frame, from_=0.5, to=3, resolution=0.1,
                     orient="horizontal", variable=var, command=self.update_live_preview).pack(fill="x")

        tk.Checkbutton(settings_frame, text="Grayscale", variable=self.grayscale,
                       command=self.update_live_preview).pack(pady=5)

        tk.Button(settings_frame, text="Apply to All Pages", command=self.apply_to_all_pages).pack(pady=3)
        tk.Button(settings_frame, text="Auto Optimize for Print", command=self.auto_optimize).pack(pady=3)
        ttk.Button(settings_frame, text="💾 Save Settings", command=self.save_settings).pack(pady=5)

        # --- Export Section ---
        export_frame = tk.LabelFrame(root, text="Export", padx=10, pady=5)
        export_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
        export_frame.columnconfigure(1, weight=1)

        tk.Label(export_frame, text="Page ranges (e.g. 1-5,7,10-12):").grid(row=0, column=0, sticky="w")
        self.range_entry = tk.Entry(export_frame)
        self.range_entry.grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(export_frame, text="Export PDF", command=self.export_pdf).grid(row=0, column=2, padx=5)

        # --- Keyboard Shortcuts ---
        self.root.bind("<Left>", lambda e: self.change_page(-1))
        self.root.bind("<Right>", lambda e: self.change_page(1))
        self.root.bind("<Control-s>", lambda e: self.export_pdf())
        self.root.bind("<Control-o>", lambda e: self.open_pdf())

    # -----------------------------
    # File Operations
    # -----------------------------
    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            threading.Thread(target=self._load_pdf_thread, args=(path,)).start()

    def _load_pdf_thread(self, path):
        try:
            self.progress_var.set(0)
            self.manager.load_pdf(path, self.get_settings(), progress_callback=self.update_progress)
            self.progress_var.set(1)
            self.current_index = 0
            self.update_previews()
            time.sleep(0.3)
            self.progress_var.set(0)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_batch(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return

        def batch_convert():
            for path in paths:
                self.progress_var.set(0)
                self.manager.load_pdf(path, self.get_settings(), progress_callback=self.update_progress)
                out = path.replace(".pdf", "_converted.pdf")
                self.manager.export_pdf(out, self.get_settings(), progress_callback=self.update_progress)
                time.sleep(0.3)
            self.progress_var.set(0)
            messagebox.showinfo("Done", "Batch conversion completed!")

        threading.Thread(target=batch_convert).start()

    def export_pdf(self):
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if out:
            try:
                threading.Thread(target=self._export_thread, args=(out,)).start()
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def _export_thread(self, out):
        self.progress_var.set(0)
        self.manager.export_pdf(out, self.get_settings(), self.range_entry.get(), progress_callback=self.update_progress)
        self.progress_var.set(1)
        time.sleep(0.4)
        self.progress_var.set(0)
        messagebox.showinfo("Export", f"Saved to {out}")

    # -----------------------------
    # Page Controls
    # -----------------------------
    def change_page(self, delta):
        if self.manager.pages:
            self.current_index = (self.current_index + delta) % len(self.manager.pages)
            self.update_previews()

    # -----------------------------
    # Preview & Live Update
    # -----------------------------
    def update_previews(self):
        if not self.manager.pages:
            return
        orig = self.manager.original_pages[self.current_index].copy()
        proc = self.manager.pages[self.current_index].copy()

        w = self.canvas_left.winfo_width() - 20
        h = self.canvas_left.winfo_height() - 20
        orig.thumbnail((w, h), Image.LANCZOS)
        proc.thumbnail((w, h), Image.LANCZOS)

        self.tk_orig = ImageTk.PhotoImage(orig)
        self.tk_proc = ImageTk.PhotoImage(proc)

        self.canvas_left.delete("all")
        self.canvas_left.create_image(w // 2, h // 2, image=self.tk_orig, anchor="center")

        self.canvas_right.delete("all")
        self.canvas_right.create_image(w // 2, h // 2, image=self.tk_proc, anchor="center")

    def update_live_preview(self, *_):
        if not self.manager.pages:
            return
        self.manager.reprocess_current_page(self.current_index, self.get_settings())
        self.update_previews()

    def apply_to_all_pages(self):
        if not self.manager.pages:
            return
        threading.Thread(target=self._apply_all_thread).start()

    def _apply_all_thread(self):
        self.progress_var.set(0)
        self.manager.reprocess_all_pages(self.get_settings(), progress_callback=self.update_progress)
        self.progress_var.set(1)
        time.sleep(0.4)
        self.progress_var.set(0)
        self.update_previews()
        messagebox.showinfo("Done", "Settings applied to all pages!")

    def auto_optimize(self):
        # One-click print-ready settings
        preset = {"contrast": 1.3, "brightness": 1.05, "sharpness": 1.1, "grayscale": True}
        for k, v in preset.items():
            getattr(self, k).set(v)
        self.update_live_preview()

    def update_progress(self, value):
        self.progress_var.set(value)

    # -----------------------------
    # Settings
    # -----------------------------
    def get_settings(self):
        return {
            "contrast": self.contrast.get(),
            "brightness": self.brightness.get(),
            "sharpness": self.sharpness.get(),
            "grayscale": self.grayscale.get()
        }

    def save_settings(self):
        save_settings(self.get_settings())
        messagebox.showinfo("Settings", "Saved!")


# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterGUI(root)
    root.mainloop()

