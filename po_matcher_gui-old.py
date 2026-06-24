#!/usr/bin/env python3
"""
PO Invoice Matcher  —  GUI application
=======================================
Matches invoice PDFs with purchase order PDFs by Purchase Order number.
Copies matched pairs into individual subfolders named by PO number.

Build into .exe (Windows):
    pip install pyinstaller pytesseract pdfplumber Pillow
    pyinstaller --onefile --windowed --name "PO Matcher" po_matcher_gui.py

Requirements at runtime:
    • Tesseract OCR  –  https://github.com/UB-Mannheim/tesseract/wiki
    • Poppler        –  https://github.com/oschwartz10612/poppler-windows/releases
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── optional OCR deps ─────────────────────────────────────────────────────────
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

# ── Windows: suppress console popups + auto-detect install paths ──────────────
if sys.platform == "win32":
    import ctypes
    _CREATE_NO_WINDOW = 0x08000000  # prevents subprocess console flashing

    _TESSERACT_PATHS = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    _POPPLER_PATHS = [
        r"C:\poppler\Library\bin",
        r"C:\poppler\bin",
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files\poppler\bin",
    ]
    if OCR_AVAILABLE:
        for _p in _TESSERACT_PATHS:
            if os.path.isfile(_p):
                pytesseract.pytesseract.tesseract_cmd = _p
                break
    for _p in _POPPLER_PATHS:
        if os.path.isdir(_p):
            os.environ["PATH"] = _p + os.pathsep + os.environ.get("PATH", "")
            break
else:
    _CREATE_NO_WINDOW = 0  # no-op on non-Windows


def _run_silent(cmd, **kwargs):
    """Run a subprocess with no visible console window on Windows."""
    return subprocess.run(
        cmd,
        creationflags=_CREATE_NO_WINDOW,
        **kwargs,
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  PO NUMBER LOGIC
# ═══════════════════════════════════════════════════════════════════════════════

_CANDIDATE_RE = re.compile(
    r"""
    (?:
        (?:p/?o|purchase\s*(?:order|no\.?|number)|p\.o\.|order\s*no\.?|po\s*no\.?|p/o\s*(?:number|no\.?)?)
        [\s:.\-#]*
        ([2][0-9]{4,7})
    )
    |
    \b(2[0-9]{7})\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

_SKIP_RE = re.compile(r"^(202[0-9]|203[0-9]|6[0-9]{7}|8[0-9]{7}|9[0-9]{7})$")


def _normalise_po(raw: str) -> str | None:
    raw = raw.strip()
    if not raw.startswith("2") or _SKIP_RE.match(raw):
        return None
    if len(raw) == 8:
        return raw
    if len(raw) < 5 or len(raw) > 9:
        return None
    prefix, suffix = raw[:2], raw[2:]
    canonical = prefix + "0" * (8 - len(raw)) + suffix
    return canonical if len(canonical) == 8 else None


def extract_po_numbers(text: str) -> set[str]:
    found: set[str] = set()
    for m in _CANDIDATE_RE.finditer(text):
        raw = m.group(1) or m.group(2)
        if raw:
            c = _normalise_po(raw)
            if c:
                found.add(c)
    return found


# ═══════════════════════════════════════════════════════════════════════════════
#  PDF TEXT EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_text_pdfplumber(path: Path) -> str:
    if not PDFPLUMBER_AVAILABLE:
        return ""
    try:
        with pdfplumber.open(str(path)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return ""


def _extract_text_ocr(path: Path, stop_event: threading.Event) -> str:
    """
    Rasterise each page with pdftoppm (no console flash) then OCR with
    Tesseract.  Checks stop_event between pages so Stop works mid-file.
    """
    if not OCR_AVAILABLE:
        return ""
    parts: list[str] = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        try:
            result = _run_silent(
                ["pdftoppm", "-jpeg", "-r", "200", str(path), str(tmp / "page")],
                capture_output=True,
                timeout=120,
            )
            if result.returncode != 0:
                return ""
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return ""

        images = sorted(tmp.glob("page-*.jpg")) + sorted(tmp.glob("page-*.ppm"))
        for img_path in images:
            if stop_event.is_set():
                break
            try:
                img = Image.open(img_path)
                parts.append(pytesseract.image_to_string(img, lang="eng"))
            except Exception:
                pass
    return "\n".join(parts)


def get_po_numbers(path: Path, stop_event: threading.Event) -> set[str]:
    text = _extract_text_pdfplumber(path)
    pos = extract_po_numbers(text)
    if not pos and not stop_event.is_set():
        text = _extract_text_ocr(path, stop_event)
        pos = extract_po_numbers(text)
    return pos


# ═══════════════════════════════════════════════════════════════════════════════
#  MATCHING LOGIC  (runs in background thread)
# ═══════════════════════════════════════════════════════════════════════════════

def run_matching(
    inv_folder: Path,
    po_folder: Path,
    out_folder: Path,
    msg_queue: queue.Queue,
    stop_event: threading.Event,
) -> None:
    def log(msg: str):
        msg_queue.put(("log", msg))

    def stopped() -> bool:
        return stop_event.is_set()

    try:
        inv_pdfs = sorted(inv_folder.glob("*.[pP][dD][fF]"))
        po_pdfs  = sorted(po_folder.glob("*.[pP][dD][fF]"))
        total    = len(inv_pdfs) + len(po_pdfs)

        if total == 0:
            msg_queue.put(("error", "No PDF files found in the selected folders."))
            return

        inv_map: dict[str, list[Path]] = {}
        po_map:  dict[str, list[Path]] = {}
        done = 0

        log("── Scanning invoice folder ──")
        for pdf in inv_pdfs:
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return
            log(f"  {pdf.name}")
            msg_queue.put(("progress", (done, total, f"Scanning {pdf.name}")))
            pos = get_po_numbers(pdf, stop_event)
            log(f"    → PO numbers: {pos or '(none found)'}")
            for po in pos:
                inv_map.setdefault(po, []).append(pdf)
            done += 1

        log("\n── Scanning PO folder ──")
        for pdf in po_pdfs:
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return
            log(f"  {pdf.name}")
            msg_queue.put(("progress", (done, total, f"Scanning {pdf.name}")))
            pos = get_po_numbers(pdf, stop_event)
            log(f"    → PO numbers: {pos or '(none found)'}")
            for po in pos:
                po_map.setdefault(po, []).append(pdf)
            done += 1

        if stopped():
            msg_queue.put(("stopped", "Stopped by user."))
            return

        msg_queue.put(("progress", (total, total, "Copying matched files…")))

        common = set(inv_map) & set(po_map)
        if not common:
            log("\n⚠  No matching PO numbers found between the two folders.")
            log(f"   Invoice POs : {sorted(inv_map) or '(none)'}")
            log(f"   PO file POs : {sorted(po_map) or '(none)'}")
            msg_queue.put(("done", "No matches found. Check the log for details."))
            return

        log("\n── Copying matched files ──")
        out_folder.mkdir(parents=True, exist_ok=True)
        copied = 0

        for po in sorted(common):
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return
            dest = out_folder / po
            dest.mkdir(parents=True, exist_ok=True)
            seen: set[Path] = set()
            for src in inv_map[po] + po_map[po]:
                if src in seen:
                    continue
                seen.add(src)
                dst_path = dest / src.name
                if dst_path.exists():
                    counter = 1
                    while dst_path.exists():
                        dst_path = dest / f"{src.stem}_{counter}{src.suffix}"
                        counter += 1
                shutil.copy2(src, dst_path)
                log(f"  [{po}] {src.name}")
                copied += 1

        summary = (
            f"✓ Matched {len(common)} PO number(s), copied {copied} file(s).\n"
            f"Output folder: {out_folder.resolve()}"
        )
        log(f"\n{summary}")
        msg_queue.put(("done", summary))

    except Exception as exc:
        msg_queue.put(("error", str(exc)))


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════════════════════════

class POMatcherApp(tk.Tk):
    BG         = "#1e2330"
    PANEL      = "#252b3b"
    ACCENT     = "#4f8ef7"
    ACCENT_HOV = "#699eff"
    SUCCESS    = "#3ecf8e"
    WARNING    = "#f6a623"
    TEXT       = "#e8ecf4"
    MUTED      = "#7a859e"
    LOG_BG     = "#141820"
    LOG_FG     = "#c5cede"

    def __init__(self):
        super().__init__()
        self.title("PO Invoice Matcher")
        self.resizable(True, True)
        self.minsize(680, 560)
        self.configure(bg=self.BG)

        self._q: queue.Queue        = queue.Queue()
        self._thread: threading.Thread | None = None
        self._stop_event            = threading.Event()

        self._inv_var = tk.StringVar()
        self._po_var  = tk.StringVar()
        self._out_var = tk.StringVar()

        self._build_ui()
        self._center_window(720, 620)
        self._poll_queue()

    def _center_window(self, w: int, h: int):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        hdr = tk.Frame(self, bg=self.PANEL, pady=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="PO Invoice Matcher",
                 font=("Segoe UI", 18, "bold"),
                 bg=self.PANEL, fg=self.TEXT).pack()
        tk.Label(hdr, text="Match invoice PDFs to purchase orders by PO number",
                 font=("Segoe UI", 9),
                 bg=self.PANEL, fg=self.MUTED).pack()

        pickers = tk.Frame(self, bg=self.BG, padx=24, pady=16)
        pickers.pack(fill="x")
        self._folder_row(pickers, "Invoice Folder",        self._inv_var, 0)
        self._folder_row(pickers, "Purchase Order Folder", self._po_var,  1)
        self._folder_row(pickers, "Output Folder",         self._out_var, 2)

        prog_frame = tk.Frame(self, bg=self.BG, padx=24)
        prog_frame.pack(fill="x")
        self._prog_label = tk.Label(prog_frame, text="Ready",
                                    font=("Segoe UI", 9),
                                    bg=self.BG, fg=self.MUTED, anchor="w")
        self._prog_label.pack(fill="x", pady=(0, 4))

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("PO.Horizontal.TProgressbar",
                         troughcolor=self.LOG_BG, background=self.ACCENT,
                         borderwidth=0, thickness=10)
        self._progress = ttk.Progressbar(prog_frame,
                                          style="PO.Horizontal.TProgressbar",
                                          orient="horizontal",
                                          mode="determinate", maximum=100)
        self._progress.pack(fill="x", pady=(0, 12))

        btn_frame = tk.Frame(self, bg=self.BG, padx=24, pady=4)
        btn_frame.pack(fill="x")
        self._start_btn = self._make_button(
            btn_frame, "▶  Start", self.ACCENT, self.ACCENT_HOV, self._on_start)
        self._start_btn.pack(side="left", ipadx=18, ipady=6)
        self._stop_btn = self._make_button(
            btn_frame, "■  Stop", "#c0392b", "#e74c3c", self._on_stop,
            state="disabled")
        self._stop_btn.pack(side="left", padx=(10, 0), ipadx=18, ipady=6)
        self._open_btn = self._make_button(
            btn_frame, "📂  Open Output", self.PANEL, "#2e3650",
            self._open_output, fg=self.TEXT)
        self._open_btn.pack(side="right", ipadx=12, ipady=6)

        log_frame = tk.Frame(self, bg=self.BG, padx=24, pady=8)
        log_frame.pack(fill="both", expand=True)
        tk.Label(log_frame, text="LOG", font=("Segoe UI", 8, "bold"),
                 bg=self.BG, fg=self.MUTED).pack(anchor="w")
        text_outer = tk.Frame(log_frame, bg=self.ACCENT, bd=1)
        text_outer.pack(fill="both", expand=True, pady=(2, 0))
        scroll = tk.Scrollbar(text_outer)
        scroll.pack(side="right", fill="y")
        self._log = tk.Text(text_outer, bg=self.LOG_BG, fg=self.LOG_FG,
                             font=("Consolas", 9), bd=0, relief="flat",
                             wrap="word", yscrollcommand=scroll.set,
                             state="disabled", padx=10, pady=8)
        self._log.pack(fill="both", expand=True)
        scroll.config(command=self._log.yview)
        self._log.tag_config("success", foreground=self.SUCCESS)
        self._log.tag_config("warn",    foreground=self.WARNING)
        self._log.tag_config("muted",   foreground=self.MUTED)

        self._status = tk.Label(self, text="  Select folders and press Start.",
                                 font=("Segoe UI", 8),
                                 bg=self.PANEL, fg=self.MUTED,
                                 anchor="w", pady=5)
        self._status.pack(fill="x", side="bottom")

    def _folder_row(self, parent, label, var, row):
        tk.Label(parent, text=label, font=("Segoe UI", 9, "bold"),
                 bg=self.BG, fg=self.MUTED, width=22, anchor="w"
                 ).grid(row=row, column=0, padx=(0, 8), pady=5, sticky="w")
        tk.Entry(parent, textvariable=var, font=("Segoe UI", 9),
                 bg=self.LOG_BG, fg=self.TEXT, insertbackground=self.TEXT,
                 relief="flat", bd=6
                 ).grid(row=row, column=1, sticky="ew", pady=5)
        tk.Button(parent, text="Browse", font=("Segoe UI", 9),
                  bg=self.PANEL, fg=self.TEXT,
                  activebackground=self.ACCENT, activeforeground="white",
                  relief="flat", cursor="hand2", bd=0, padx=10,
                  command=lambda v=var: self._browse(v)
                  ).grid(row=row, column=2, padx=(8, 0), pady=5)
        parent.columnconfigure(1, weight=1)

    def _make_button(self, parent, text, bg, hover_bg, cmd,
                     state="normal", fg="white"):
        btn = tk.Button(parent, text=text, font=("Segoe UI", 10, "bold"),
                        bg=bg, fg=fg, activebackground=hover_bg,
                        activeforeground="white", relief="flat",
                        cursor="hand2", bd=0, command=cmd, state=state)
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    # ── actions ───────────────────────────────────────────────────────────────
    def _browse(self, var):
        path = filedialog.askdirectory(title="Select folder")
        if path:
            var.set(path)

    def _on_start(self):
        inv = self._inv_var.get().strip()
        po  = self._po_var.get().strip()
        out = self._out_var.get().strip()
        if not inv or not po or not out:
            messagebox.showwarning("Missing folders", "Please select all three folders.")
            return
        if not Path(inv).is_dir():
            messagebox.showerror("Not found", f"Invoice folder not found:\n{inv}")
            return
        if not Path(po).is_dir():
            messagebox.showerror("Not found", f"Purchase order folder not found:\n{po}")
            return

        self._log_clear()
        self._log_write("Starting…\n", "muted")
        self._set_state("running")
        self._progress["value"] = 0
        self._prog_label.configure(text="Starting…")

        # Fresh stop event for this run
        self._stop_event = threading.Event()

        self._thread = threading.Thread(
            target=run_matching,
            args=(Path(inv), Path(po), Path(out), self._q, self._stop_event),
            daemon=True,
        )
        self._thread.start()

    def _on_stop(self):
        """Signal the worker thread to stop after its current file."""
        self._stop_event.set()
        self._stop_btn.configure(state="disabled")
        self._log_write("\n⚠  Stop requested — finishing current file…\n", "warn")
        self._status.configure(text="  Stopping after current file…")

    def _open_output(self):
        out = self._out_var.get().strip()
        if not out:
            messagebox.showinfo("Output folder", "No output folder set yet.")
            return
        path = Path(out)
        if not path.exists():
            messagebox.showinfo("Output folder", "Output folder does not exist yet.")
            return
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            _run_silent(["open", str(path)])
        else:
            _run_silent(["xdg-open", str(path)])

    # ── queue polling ─────────────────────────────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                msg_type, payload = self._q.get_nowait()

                if msg_type == "log":
                    tag = ""
                    if payload.startswith("✓") or "[" in payload:
                        tag = "success"
                    elif "⚠" in payload or "warn" in payload.lower():
                        tag = "warn"
                    elif payload.startswith("  ") or payload.startswith("──"):
                        tag = "muted"
                    self._log_write(payload + "\n", tag)

                elif msg_type == "progress":
                    current, total, label = payload
                    pct = int(current / total * 100) if total else 0
                    self._progress["value"] = pct
                    self._prog_label.configure(text=f"{label}  ({current}/{total})")
                    self._status.configure(text=f"  Processing… {pct}%")

                elif msg_type == "done":
                    self._progress["value"] = 100
                    self._prog_label.configure(text="Complete")
                    self._set_state("idle")
                    self._status.configure(text=f"  {payload}")
                    messagebox.showinfo("Done", payload)

                elif msg_type == "stopped":
                    self._prog_label.configure(text="Stopped")
                    self._set_state("idle")
                    self._status.configure(text="  Stopped by user.")
                    self._log_write("\n■  Stopped by user.\n", "warn")

                elif msg_type == "error":
                    self._set_state("idle")
                    self._log_write(f"\n✗ Error: {payload}\n", "warn")
                    self._status.configure(text=f"  Error: {payload}")
                    messagebox.showerror("Error", payload)

        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    # ── helpers ───────────────────────────────────────────────────────────────
    def _set_state(self, state: str):
        if state == "running":
            self._start_btn.configure(state="disabled")
            self._stop_btn.configure(state="normal")
        else:
            self._start_btn.configure(state="normal")
            self._stop_btn.configure(state="disabled")

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _log_write(self, msg: str, tag: str = ""):
        self._log.configure(state="normal")
        if tag:
            self._log.insert("end", msg, tag)
        else:
            self._log.insert("end", msg)
        self._log.see("end")
        self._log.configure(state="disabled")


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

def _check_dependencies() -> list[str]:
    issues = []
    if not OCR_AVAILABLE:
        issues.append("pytesseract / Pillow not installed  →  pip install pytesseract Pillow")
    if not PDFPLUMBER_AVAILABLE:
        issues.append("pdfplumber not installed  →  pip install pdfplumber")
    try:
        _run_silent(["pdftoppm", "-v"], capture_output=True, timeout=5)
    except FileNotFoundError:
        issues.append(
            "Poppler (pdftoppm) not found.\n"
            "  Windows: https://github.com/oschwartz10612/poppler-windows/releases\n"
            "           Add the bin\\ folder to PATH."
        )
    except Exception:
        pass
    if OCR_AVAILABLE:
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            issues.append(
                "Tesseract OCR not found.\n"
                "  Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
                "           Add C:\\Program Files\\Tesseract-OCR\\ to PATH."
            )
    return issues


if __name__ == "__main__":
    issues = _check_dependencies()
    app = POMatcherApp()
    if issues:
        msg = "Some dependencies are missing:\n\n" + "\n\n".join(issues)
        messagebox.showwarning("Setup required", msg)
    app.mainloop()