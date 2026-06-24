#!/usr/bin/env python3
"""
PO Invoice Matcher  —  GUI application
=======================================
Matches invoice PDFs with purchase order PDFs by Purchase Order number.
When a PO folder contains more than 2 files, further groups them into
subfolders by D/N (Delivery Note) number.

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
    _CREATE_NO_WINDOW = 0x08000000
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
    _CREATE_NO_WINDOW = 0


def _run_silent(cmd, **kwargs):
    """Run subprocess with no visible console window on Windows."""
    return subprocess.run(cmd, creationflags=_CREATE_NO_WINDOW, **kwargs)


# ═══════════════════════════════════════════════════════════════════════════════
#  PO NUMBER EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

_PO_RE = re.compile(
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
    for m in _PO_RE.finditer(text):
        raw = m.group(1) or m.group(2)
        if raw:
            c = _normalise_po(raw)
            if c:
                found.add(c)
    return found


# ═══════════════════════════════════════════════════════════════════════════════
#  D/N NUMBER EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════
#
#  Patterns seen across document types:
#    Tamaco invoices/DOs : "D/N : 02"  or  "DIN: 02"  (OCR reads / as I)
#    UE internal DNs     : "D/N  02"   or  "Vendor: ... D/N  02"
#    Seen Joo DOs        : "D/O NUMBER : DO-159552"  (different — these are
#                          internal DO refs, not delivery batch numbers)
#
#  We capture the small zero-padded batch number (01–99) that groups
#  multiple invoices/DOs belonging to the same delivery batch under one PO.

_DN_RE = re.compile(
    r"""
    (?:
        # "D/N", "DN", "DIN" (OCR artefact), "D.N" followed by number
        d[/\.]?[in][\s:]*(\d{1,2})\b
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

# Extra pattern: UE-style "D/N" label on its own line then number next token
_DN_LABEL_RE = re.compile(
    r"d\s*/\s*n[\s:]*(\d{1,2})\b",
    re.IGNORECASE,
)

# Pattern: Tamaco "DIN: 02" / "D/N: 02" anywhere in text
_DN_COLON_RE = re.compile(
    r"\b(?:d[/.]?n|din)\s*[:\-]\s*(\d{1,2})\b",
    re.IGNORECASE,
)


def extract_dn_numbers(text: str) -> set[str]:
    """
    Return the set of D/N batch numbers found in *text*, zero-padded to 2 digits.
    e.g. "02", "03"
    """
    found: set[str] = set()

    for pattern in (_DN_COLON_RE, _DN_LABEL_RE):
        for m in pattern.finditer(text):
            n = int(m.group(1))
            if 1 <= n <= 99:
                found.add(f"{n:02d}")

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
    if not OCR_AVAILABLE:
        return ""
    parts: list[str] = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        try:
            result = _run_silent(
                ["pdftoppm", "-jpeg", "-r", "200", str(path), str(tmp / "page")],
                capture_output=True, timeout=120,
            )
            if result.returncode != 0:
                return ""
        except (FileNotFoundError, subprocess.TimeoutExpired):
            return ""

        for img_path in sorted(tmp.glob("page-*.jpg")) + sorted(tmp.glob("page-*.ppm")):
            if stop_event.is_set():
                break
            try:
                img = Image.open(img_path)
                parts.append(pytesseract.image_to_string(img, lang="eng"))
            except Exception:
                pass
    return "\n".join(parts)


def get_document_info(path: Path, stop_event: threading.Event) -> tuple[set[str], set[str]]:
    """
    Return (po_numbers, dn_numbers) extracted from the PDF.
    Tries text layer first; falls back to OCR for scanned files.
    """
    text = _extract_text_pdfplumber(path)
    pos  = extract_po_numbers(text)
    dns  = extract_dn_numbers(text)

    if (not pos or not dns) and not stop_event.is_set():
        ocr_text = _extract_text_ocr(path, stop_event)
        if not pos:
            pos = extract_po_numbers(ocr_text)
        if not dns:
            dns = extract_dn_numbers(ocr_text)
        # Merge: combine both sources in case one had PO and other had DN
        if not pos:
            pos = extract_po_numbers(text)
        dns |= extract_dn_numbers(text)

    return pos, dns


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

        # po_map  : {po_number: [(path, dn_numbers), ...]}
        inv_map: dict[str, list[tuple[Path, set[str]]]] = {}
        po_map:  dict[str, list[tuple[Path, set[str]]]] = {}
        done = 0

        log("── Scanning invoice folder ──")
        for pdf in inv_pdfs:
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return
            log(f"  {pdf.name}")
            msg_queue.put(("progress", (done, total, f"Scanning {pdf.name}")))
            pos, dns = get_document_info(pdf, stop_event)
            log(f"    → PO: {pos or '(none)'}  D/N: {dns or '(none)'}")
            for po in pos:
                inv_map.setdefault(po, []).append((pdf, dns))
            done += 1

        log("\n── Scanning PO folder ──")
        for pdf in po_pdfs:
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return
            log(f"  {pdf.name}")
            msg_queue.put(("progress", (done, total, f"Scanning {pdf.name}")))
            pos, dns = get_document_info(pdf, stop_event)
            log(f"    → PO: {pos or '(none)'}  D/N: {dns or '(none)'}")
            for po in pos:
                po_map.setdefault(po, []).append((pdf, dns))
            done += 1

        if stopped():
            msg_queue.put(("stopped", "Stopped by user."))
            return

        msg_queue.put(("progress", (total, total, "Organising matched files…")))

        common = set(inv_map) & set(po_map)
        if not common:
            log("\n⚠  No matching PO numbers found between the two folders.")
            log(f"   Invoice POs  : {sorted(inv_map) or '(none)'}")
            log(f"   PO file POs  : {sorted(po_map) or '(none)'}")
            msg_queue.put(("done", "No matches found. Check the log for details."))
            return

        log(f"\n── Organising matched files ──")
        out_folder.mkdir(parents=True, exist_ok=True)
        copied = 0

        for po in sorted(common):
            if stopped():
                msg_queue.put(("stopped", "Stopped by user."))
                return

            all_entries: list[tuple[Path, set[str]]] = inv_map[po] + po_map[po]

            # De-duplicate paths
            seen_paths: set[Path] = set()
            unique_entries: list[tuple[Path, set[str]]] = []
            for path, dns in all_entries:
                if path not in seen_paths:
                    seen_paths.add(path)
                    unique_entries.append((path, dns))

            po_dest = out_folder / po
            po_dest.mkdir(parents=True, exist_ok=True)

            # ── Decide whether to use D/N subfolders ──────────────────────────
            # Collect all D/N numbers across all files in this PO group
            all_dns: set[str] = set()
            for _, dns in unique_entries:
                all_dns |= dns

            use_dn_folders = len(unique_entries) > 2 and len(all_dns) > 1

            if use_dn_folders:
                log(f"\n  [PO {po}]  {len(unique_entries)} files, "
                    f"D/N numbers found: {sorted(all_dns)} → using D/N subfolders")
            else:
                log(f"\n  [PO {po}]  {len(unique_entries)} file(s) → flat folder")

            for src, dns in unique_entries:
                if use_dn_folders and dns:
                    # File belongs to specific D/N(s) — put in each matching subfolder
                    for dn in sorted(dns):
                        dn_dest = po_dest / f"DN{dn}"
                        dn_dest.mkdir(parents=True, exist_ok=True)
                        dst_path = _unique_path(dn_dest / src.name)
                        shutil.copy2(src, dst_path)
                        log(f"    [{po}/DN{dn}] {src.name}")
                        copied += 1
                elif use_dn_folders and not dns:
                    # File has no D/N — goes into a special "unmatched" subfolder
                    unk_dest = po_dest / "DN_unknown"
                    unk_dest.mkdir(parents=True, exist_ok=True)
                    dst_path = _unique_path(unk_dest / src.name)
                    shutil.copy2(src, dst_path)
                    log(f"    [{po}/DN_unknown] {src.name}  ⚠ no D/N detected")
                    copied += 1
                else:
                    # ≤2 files or only one D/N — flat into PO folder
                    dst_path = _unique_path(po_dest / src.name)
                    shutil.copy2(src, dst_path)
                    log(f"    [{po}] {src.name}")
                    copied += 1

        summary = (
            f"✓ Matched {len(common)} PO number(s), copied {copied} file(s).\n"
            f"Output folder: {out_folder.resolve()}"
        )
        log(f"\n{summary}")
        msg_queue.put(("done", summary))

    except Exception as exc:
        msg_queue.put(("error", str(exc)))


def _unique_path(path: Path) -> Path:
    """Return path, incrementing a counter suffix if it already exists."""
    if not path.exists():
        return path
    counter = 1
    while True:
        candidate = path.parent / f"{path.stem}_{counter}{path.suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


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

        self._q: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()

        self._inv_var = tk.StringVar()
        self._po_var  = tk.StringVar()
        self._out_var = tk.StringVar()

        self._build_ui()
        self._center_window(720, 640)
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
        tk.Label(hdr, text="Match invoice PDFs to purchase orders · groups by D/N when needed",
                 font=("Segoe UI", 9), bg=self.PANEL, fg=self.MUTED).pack()

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
        self._stop_event = threading.Event()

        self._thread = threading.Thread(
            target=run_matching,
            args=(Path(inv), Path(po), Path(out), self._q, self._stop_event),
            daemon=True,
        )
        self._thread.start()

    def _on_stop(self):
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

    def _poll_queue(self):
        try:
            while True:
                msg_type, payload = self._q.get_nowait()
                if msg_type == "log":
                    tag = ""
                    if payload.startswith("✓") or ("[" in payload and "/DN" in payload):
                        tag = "success"
                    elif "[" in payload and "/DN" not in payload:
                        tag = "success"
                    elif "⚠" in payload or "warn" in payload.lower() or "unknown" in payload:
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