# PW Notes PDF Converter (Invert + Enhance)

A simple desktop GUI to convert PDFs into high-contrast, inverted-color pages with live preview, brightness/contrast/sharpness controls, grayscale mode, selective export by page ranges, and batch processing. Built with PyMuPDF, Pillow, and Tkinter.

- Split preview: original vs processed side-by-side with live updates.
- One-click invert with optional grayscale; adjustable contrast, brightness, sharpness.
- Apply settings to one or all pages with a progress bar.
- Export full PDF or page ranges like 1-5,7,10-12.
- Batch convert multiple PDFs.
- Settings persist to settings.json.

***
## - You can easily set multiple pages per sheet (e.g. for printing) in your PDF with this online tool.
- https://online2pdf.com/en/multiple-pages-per-sheet
## Quick start (zero prior experience)

Follow these steps exactly. If something doesn’t work, tell me which step you’re on and what error you see.

### 1) Install Python
- Windows: install Python 3.10+ from the official website. During setup, check “Add Python to PATH”.
- macOS: Python 3 usually available via Homebrew (`brew install python`) or from the official site.
- Linux: use your package manager (e.g., `sudo apt-get install python3 python3-venv python3-tk`).

### 2) Create a project folder
- Make a new folder anywhere, e.g., `PW-PDF-Converter`.
- Put the provided Python file inside it as `PDF_converter_V4.py`.

### 3) Open a terminal in that folder
- Windows: Shift + Right-click in the folder → “Open PowerShell window here” (or use Terminal, `cd path/to/folder`).
- macOS/Linux: Right-click → “Open in Terminal” (or `cd path/to/folder`).

### 4) Create and activate a virtual environment
- Windows:
  - `python -m venv .venv`
  - `.\.venv\Scripts\activate`
- macOS/Linux:
  - `python3 -m venv .venv`
  - `source .venv/bin/activate`

You should see `(.venv)` at the start of your terminal line. Keep this activated whenever you work on the project.

### 5) Install required libraries
Run these commands in your activated virtual environment:
- `pip install --upgrade pip`
- `pip install pymupdf pillow`

If Tkinter is missing:
- Ubuntu/Debian: `sudo apt-get install python3-tk`
- Fedora: `sudo dnf install python3-tkinter`
- Arch: `sudo pacman -S tk`

### 6) Run the app
- Windows: `python PDF_converter_V4.py`
- macOS/Linux: `python3 PDF_converter_V4.py`

A window titled “PW Notes Converter - Split View + Live Tools” should open.

***

## How to use
- **Open PDF**: Click “📂 Open PDF” and choose a file. Preview shows Original (left) and Processed (right).
- **Adjust settings**: Move sliders for Contrast, Brightness, Sharpness; toggle Grayscale. Preview updates live.
- **Apply to all pages**: Click “Apply to All Pages” to reprocess entire document with current settings.
- **Auto optimize**: “Auto Optimize for Print” sets contrast 1.3, brightness 1.05, sharpness 1.1, grayscale on.
- **For Best Results use these Setting for PW Notes** : ["contrast": 2.4, "brightness": 1.9, "sharpness": 1.6, "grayscale": On]
- **Navigate pages**: Use “⏮ Prev” and “Next ⏭”.
- **Export**: In Export box, optionally enter ranges like `1-5,7,10-12`. Click “Export PDF” and choose save path.
- **Batch mode**: Click “🗂 Batch Mode” to select multiple PDFs; each is processed and saved as `_converted`.

### Page ranges format
- Comma-separated tokens.
- Individual pages: `3,7`
- Ranges: `1-5`
- Mixed: `1-3,6,9-12`
- Input uses 1-based page numbers (the tool converts internally).

***

## Troubleshooting
- **Module not found**: Ensure your virtual environment is activated and you ran `pip install` commands in it.
- **Tkinter errors on Linux**: install your distro’s `python3-tk` package (see step 5).
- **Slow or large memory use**: Very large PDFs take longer because pages are rasterized at 200 DPI.
- **Exported PDF is image-based**: Text won’t be selectable; this is expected.

***

## Code
Copy this entire block into `PDF_converter_V4.py` in your project folder.


