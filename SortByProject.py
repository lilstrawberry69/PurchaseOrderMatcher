#!/usr/bin/env python3
"""
Project Sorter  —  GUI application
====================================
Reads combined PDFs from a PO Matcher output folder, extracts the delivery
location / project name from each PDF, and copies (or moves) them into:

    <projects_folder>/<Project Name>/<po_number>/[DN##/]<filename>.pdf

Build into .exe (Windows):
    pip install pyinstaller pytesseract pdfplumber Pillow
    pyinstaller --onefile --windowed --name "Project Sorter" project_sorter.py

Requirements at runtime:
    • Tesseract OCR  –  https://github.com/UB-Mannheim/tesseract/wiki
    • Poppler        –  https://github.com/oschwartz10612/poppler-windows/releases
"""

import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── optional OCR / PDF deps ───────────────────────────────────────────────────
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
    _CNW = 0x08000000
    for _p in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]:
        if OCR_AVAILABLE and os.path.isfile(_p):
            pytesseract.pytesseract.tesseract_cmd = _p
            break
    for _p in [
        r"C:\poppler\Library\bin",
        r"C:\poppler\bin",
        r"C:\Program Files\poppler\Library\bin",
        r"C:\Program Files\poppler\bin",
    ]:
        if os.path.isdir(_p):
            os.environ["PATH"] = _p + os.pathsep + os.environ.get("PATH", "")
            break
else:
    _CNW = 0


def _run(cmd, **kw):
    """Run a subprocess with no visible console window on Windows."""
    return subprocess.run(cmd, creationflags=_CNW, **kw)


# ═══════════════════════════════════════════════════════════════════════════════
#  LOCATION EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════

# Singapore estate / town names — extend this list as your project portfolio grows
KNOWN_AREAS = [
    'BUKIT BATOK', 'BUKIT TIMAH', 'BUKIT PANJANG', 'BUKIT MERAH',
    'QUEENSWAY', 'TENGAH', 'JURONG EAST', 'JURONG WEST',
    'CLEMENTI', 'WOODLANDS', 'TAMPINES', 'BEDOK', 'PUNGGOL', 'SENGKANG',
    'YISHUN', 'ANG MO KIO', 'BISHAN', 'TOA PAYOH', 'HOUGANG', 'SERANGOON',
    'PASIR RIS', 'CHANGI', 'KALLANG', 'GEYLANG', 'MARINE PARADE',
    'NOVENA', 'NEWTON', 'ORCHARD', 'RIVER VALLEY', 'BUONA VISTA',
    'WEST COAST', 'PIONEER', 'BOON LAY', 'CHOA CHU KANG', 'DOVER',
    'HOLLAND', 'KENT RIDGE', 'ONE NORTH', 'SEMBAWANG', 'CANBERRA',
    'MARSILING', 'KRANJI', 'TUAS', 'JURONG ISLAND',
]

# Matches area name + optional cluster codes e.g. "BUKIT BATOK N4", "TENGAH C1A"
_AREA_RE = re.compile(
    r'\b(' + '|'.join(re.escape(a) for a in KNOWN_AREAS) + r')'
    r'(?:\s+(?:[A-Z]\d+[A-Z]?(?:\s*[&,]\s*[A-Z]\d+[A-Z]?)*|\d+[A-Z]?))*',
    re.IGNORECASE,
)


def _extract_location(text: str) -> str | None:
    """
    Extract the best delivery location string from document text.

    Priority order:
      1. Known area name  — most reliable; present in almost every document type
      2. UE Purchase Order "Deliver To:" column value
      3. Tamaco / checklist "Project: <name>" line
      4. "FSA at <location>" pattern
    """
    lines = [ln.strip() for ln in text.splitlines()]

    # ── Priority 1: Known area name ───────────────────────────────────────────
    candidates: list[str] = []
    for line in lines:
        for m in _AREA_RE.finditer(line):
            candidates.append(m.group(0).strip())
    if candidates:
        return max(candidates, key=len)   # longest = most specific

    # ── Priority 2: UE Purchase Order "Deliver To:" ───────────────────────────
    # pdfplumber renders the two-column layout as merged lines, e.g.:
    #   "SEEN JOO COMPANY PTE LTD  FSA at Queensway C1"
    for i, line in enumerate(lines):
        if re.search(r'\bdeliver\s*to\s*:', line, re.I):
            for j in range(i + 1, min(i + 4, len(lines))):
                c = lines[j]
                c2 = re.sub(r'^SEEN JOO[^A-Z]*', '', c, flags=re.I).strip()
                c2 = re.sub(r'\s+PURCHASE\s+NO.*', '', c2, flags=re.I).strip()
                if len(c2) > 4:
                    return c2

    # ── Priority 3: "Project: <name>" ────────────────────────────────────────
    for line in lines:
        m = re.search(r'\bproject\s*[:\-]\s*(.+)', line, re.I)
        if m:
            val = m.group(1).strip()
            if re.match(r'[A-Z]{2,5}/[A-Z]/', val):   # skip TMC/Q/25-W217 codes
                continue
            if re.match(r'UNIT\)', val, re.I):          # skip column-merge artefact
                continue
            if len(val) > 3:
                return val

    # ── Priority 4: "FSA at <location>" ──────────────────────────────────────
    for line in lines:
        m = re.search(r'\bFSA\s+(?:at|AT)\s+(.+)', line, re.I)
        if m:
            return 'FSA ' + m.group(1).strip()

    return None


def _normalise(raw: str) -> str:
    """Clean a raw location string into a safe, readable folder name."""
    raw = re.sub(r'\s*\(\s*[CN]\d+[-\s&,CN\d]+\)', '', raw)          # (C24-C25)
    raw = re.sub(r'\s*\(\s*\d[\d\s]*(?:unit|units?)?\s*\)', '', raw, flags=re.I)
    raw = re.sub(r'\s*\(S?\s*\d{6}\s*\)', '', raw)                    # postal codes
    raw = re.sub(r'\s+\d{6}\s*$', '', raw)
    raw = re.sub(r'\s*[-\u2013]\s*CON\b.*', '', raw, flags=re.I)
    raw = re.sub(r'\bCHINA\b.*', '', raw, flags=re.I)
    raw = re.sub(r'Lorry crane.*', '', raw, flags=re.I)
    raw = re.sub(r'\bSCS Cert.*', '', raw, flags=re.I)
    raw = re.sub(r'\(GATE\s*\d+\)', '', raw, flags=re.I)
    raw = re.sub(r'LAMP POST\s*\d+', '', raw, flags=re.I)
    raw = re.sub(r',\s*$', '', raw)
    raw = re.sub(r'[<>:"/\\|?*]', '', raw)   # characters invalid in folder names
    raw = re.sub(r'&', 'and', raw)
    raw = re.sub(r'\s+', ' ', raw).strip('. ')

    def cap(w: str) -> str:
        if re.match(r'^[NC]\d+[A-Za-z]?$', w, re.I):   # N4, C24, C1A → uppercase
            return w.upper()
        return w.capitalize()

    return ' '.join(cap(w) for w in raw.split())


# ── PDF text extraction ───────────────────────────────────────────────────────

def _pdfplumber_text(path: Path) -> str:
    if not PDFPLUMBER_AVAILABLE:
        return ""
    try:
        with pdfplumber.open(str(path)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return ""


def _ocr_text(path: Path) -> str:
    if not OCR_AVAILABLE:
        return ""
    parts: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        r = _run(
            ["pdftoppm", "-jpeg", "-r", "200", str(path), f"{tmp}/page"],
            capture_output=True, timeout=120,
        )
        if r.returncode != 0:
            return ""
        for img in sorted(Path(tmp).glob("page-*.jpg")):
            try:
                parts.append(pytesseract.image_to_string(Image.open(img), lang="eng"))
            except Exception:
                pass
    return "\n".join(parts)


def get_project(pdf: Path) -> str:
    """Return a normalised project / delivery-area name extracted from the PDF."""
    text = _pdfplumber_text(pdf)
    raw  = _extract_location(text)
    if not raw:
        ocr  = _ocr_text(pdf)
        raw  = _extract_location(ocr) or _extract_location(text + "\n" + ocr)
    if not raw:
        return "Unknown Project"
    return _normalise(raw) or "Unknown Project"


# ── helpers ───────────────────────────────────────────────────────────────────

def _unique(path: Path) -> Path:
    """Return path unchanged, or with an incremented suffix if it already exists."""
    if not path.exists():
        return path
    i = 1
    while True:
        c = path.parent / f"{path.stem}_{i}{path.suffix}"
        if not c.exists():
            return c
        i += 1


# ═══════════════════════════════════════════════════════════════════════════════
#  WORKER  (runs in background thread, feeds messages to GUI queue)
# ═══════════════════════════════════════════════════════════════════════════════

def _worker(
    input_folder: Path,
    output_folder: Path,
    move: bool,
    msg_queue: queue.Queue,
    stop_event: threading.Event,
) -> None:
    """
    Sorts combined PDFs into project folders.
    Sends to msg_queue:
        ("log",      "text")
        ("progress", (current: int, total: int, label: str))
        ("done",     "summary text")
        ("stopped",  "reason")
        ("error",    "error text")
    """
    def log(msg: str):
        msg_queue.put(("log", msg))

    try:
        combined_pdfs = sorted(input_folder.rglob("combined/*.pdf"))
        total = len(combined_pdfs)

        if total == 0:
            msg_queue.put((
                "error",
                "No combined PDFs found in the selected folder.\n\n"
                "Make sure you select the PO Matcher output folder\n"
                "(the one containing PO number subfolders with 'combined' inside).",
            ))
            return

        log(f"Found {total} combined PDF(s).\n")
        output_folder.mkdir(parents=True, exist_ok=True)

        stats:  dict[str, list[str]] = {}
        errors: list[str] = []
        done = 0

        for pdf in combined_pdfs:
            if stop_event.is_set():
                msg_queue.put(("stopped", "Stopped by user."))
                return

            rel = pdf.relative_to(input_folder)
            log(f"  {rel}")
            msg_queue.put(("progress", (done, total, str(rel))))

            try:
                project = get_project(pdf)
            except Exception as e:
                log(f"    ⚠ Could not read: {e}")
                errors.append(str(rel))
                project = "Unknown Project"

            log(f"    → {project}")

            # Build destination — strip "combined" folder level, keep PO/DN structure
            rel_parts = list(rel.parts)
            rel_parts.remove("combined")
            dest_dir = output_folder / project / Path(*rel_parts[:-1])
            dest_dir.mkdir(parents=True, exist_ok=True)

            dest_file = _unique(dest_dir / pdf.name)
            if move:
                shutil.move(str(pdf), dest_file)
            else:
                shutil.copy2(pdf, dest_file)

            short = str(dest_file.relative_to(output_folder))
            log(f"    ✓ {short}\n")
            stats.setdefault(project, []).append(short)

            done += 1
            msg_queue.put(("progress", (done, total, str(rel))))

        # Build summary
        total_files = sum(len(v) for v in stats.values())
        lines = [
            f"{'Moved' if move else 'Copied'} {total_files} file(s) into "
            f"{len(stats)} project folder(s):\n"
        ]
        for proj in sorted(stats):
            lines.append(f"  📁  {proj}  ({len(stats[proj])} file(s))")
        if errors:
            lines.append(
                f"\n⚠  {len(errors)} file(s) unreadable → placed in 'Unknown Project'"
            )
        summary = "\n".join(lines)

        log("\n" + "─" * 55)
        log(summary)
        msg_queue.put(("done", summary))

    except Exception as exc:
        import traceback
        msg_queue.put(("error", f"{exc}\n\n{traceback.format_exc()}"))


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════════════════════════

class ProjectSorterApp(tk.Tk):

    # ── palette ───────────────────────────────────────────────────────────────
    BG         = "#1b001f"
    PANEL      = "#3A0053"
    ACCENT     = "#00276a"
    ACCENT_HOV = "#699eff"
    SUCCESS    = "#3ecf8e"
    WARNING    = "#f6a623"
    PURPLE     = "#b57bee"
    PURPLE_HOV = "#c99df5"
    TEXT       = "#e8ecf4"
    MUTED      = "#7a859e"
    LOG_BG     = "#30022B"
    LOG_FG     = "#c5cede"

    def __init__(self):
        super().__init__()
        self.title("Project Sorter")
        self.resizable(True, True)
        self.minsize(680, 580)
        self.configure(bg=self.BG)

        self._q          = queue.Queue()
        self._thread     = None
        self._stop_event = threading.Event()
        self._move_var   = tk.BooleanVar(value=False)
        self._input_var  = tk.StringVar()
        self._output_var = tk.StringVar()

        self._build_ui()
        self._center(740, 660)
        self._poll()

    def _center(self, w: int, h: int):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=self.PANEL, pady=16)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="Project Sorter",
            font=("Segoe UI", 18, "bold"),
            bg=self.PANEL, fg=self.TEXT,
        ).pack()
        tk.Label(
            hdr,
            text="Reads combined PDFs and sorts them into project folders by delivery location",
            font=("Segoe UI", 9), bg=self.PANEL, fg=self.MUTED,
        ).pack()

        # Folder pickers
        pickers = tk.Frame(self, bg=self.BG, padx=24, pady=16)
        pickers.pack(fill="x")

        self._folder_row(
            pickers, "Matched Folder (input)", self._input_var, row=0,
            tip="The PO Matcher output folder — contains PO subfolders with 'combined' PDFs inside",
        )
        self._folder_row(
            pickers, "Projects Folder (output)", self._output_var, row=1,
            tip="Destination — project subfolders will be created here",
        )

        # Options
        # opts = tk.Frame(self, bg=self.BG, padx=28)
        # opts.pack(fill="x")
        # tk.Checkbutton(
        #     opts, text="Move files instead of copying  (saves disk space — original files will be removed)",
        #     variable=self._move_var,
        #     font=("Segoe UI", 9),
        #     bg=self.BG, fg=self.MUTED,
        #     selectcolor=self.LOG_BG,
        #     activebackground=self.BG, activeforeground=self.TEXT,
        # ).pack(side="left", pady=4)

        # Progress
        prog_frame = tk.Frame(self, bg=self.BG, padx=24, pady=10)
        prog_frame.pack(fill="x")
        self._prog_label = tk.Label(
            prog_frame, text="Ready",
            font=("Segoe UI", 9), bg=self.BG, fg=self.MUTED, anchor="w",
        )
        self._prog_label.pack(fill="x", pady=(0, 4))

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure(
            "PS.Horizontal.TProgressbar",
            troughcolor=self.LOG_BG, background=self.PURPLE,
            borderwidth=0, thickness=10,
        )
        self._progress = ttk.Progressbar(
            prog_frame, style="PS.Horizontal.TProgressbar",
            orient="horizontal", mode="determinate", maximum=100,
        )
        self._progress.pack(fill="x", pady=(0, 4))

        # Buttons
        btn_frame = tk.Frame(self, bg=self.BG, padx=24, pady=6)
        btn_frame.pack(fill="x")

        self._start_btn = self._make_btn(
            btn_frame, "▶  Sort",
            self.PURPLE, self.PURPLE_HOV, self._on_start,
        )
        self._start_btn.pack(side="left", ipadx=18, ipady=6)

        self._stop_btn = self._make_btn(
            btn_frame, "■  Stop",
            "#c0392b", "#e74c3c", self._on_stop, state="disabled",
        )
        self._stop_btn.pack(side="left", padx=(10, 0), ipadx=18, ipady=6)

        self._open_btn = self._make_btn(
            btn_frame, "📂  Open Output",
            self.PANEL, "#2e3650", self._open_output, fg=self.TEXT,
        )
        self._open_btn.pack(side="right", ipadx=12, ipady=6)

        # Log
        log_frame = tk.Frame(self, bg=self.BG, padx=24, pady=8)
        log_frame.pack(fill="both", expand=True)
        tk.Label(
            log_frame, text="LOG", font=("Segoe UI", 8, "bold"),
            bg=self.BG, fg=self.MUTED,
        ).pack(anchor="w")
        outer = tk.Frame(log_frame, bg=self.PURPLE, bd=1)
        outer.pack(fill="both", expand=True, pady=(2, 0))
        scroll = tk.Scrollbar(outer)
        scroll.pack(side="right", fill="y")
        self._log = tk.Text(
            outer, bg=self.LOG_BG, fg=self.LOG_FG,
            font=("Consolas", 9), bd=0, relief="flat",
            wrap="word", yscrollcommand=scroll.set,
            state="disabled", padx=10, pady=8,
        )
        self._log.pack(fill="both", expand=True)
        scroll.config(command=self._log.yview)

        self._log.tag_config("success", foreground=self.SUCCESS)
        self._log.tag_config("warn",    foreground=self.WARNING)
        self._log.tag_config("project", foreground=self.PURPLE_HOV)
        self._log.tag_config("muted",   foreground=self.MUTED)

        # Status bar
        self._status = tk.Label(
            self,
            text="  Select the Matched folder and an output folder, then press Sort.",
            font=("Segoe UI", 8),
            bg=self.PANEL, fg=self.MUTED,
            anchor="w", pady=5,
        )
        self._status.pack(fill="x", side="bottom")

    def _folder_row(self, parent, label: str, var: tk.StringVar,
                    row: int, tip: str = ""):
        label_row = row * 2
        tip_row   = row * 2 + 1

        tk.Label(
            parent, text=label, font=("Segoe UI", 9, "bold"),
            bg=self.BG, fg=self.MUTED, width=24, anchor="w",
        ).grid(row=label_row, column=0, padx=(0, 8), pady=(8, 0), sticky="w")

        tk.Entry(
            parent, textvariable=var, font=("Segoe UI", 9),
            bg=self.LOG_BG, fg=self.TEXT, insertbackground=self.TEXT,
            relief="flat", bd=6,
        ).grid(row=label_row, column=1, sticky="ew", pady=(8, 0))

        tk.Button(
            parent, text="Browse", font=("Segoe UI", 9),
            bg=self.PANEL, fg=self.TEXT,
            activebackground=self.ACCENT, activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=10,
            command=lambda v=var: self._browse(v),
        ).grid(row=label_row, column=2, padx=(8, 0), pady=(8, 0))

        if tip:
            tk.Label(
                parent, text=tip, font=("Segoe UI", 7, "italic"),
                bg=self.BG, fg=self.MUTED, anchor="w",
            ).grid(row=tip_row, column=0, columnspan=3,
                   padx=(0, 8), pady=(1, 4), sticky="w")

        parent.columnconfigure(1, weight=1)

    def _make_btn(self, parent, text: str, bg: str, hover: str,
                  cmd, state: str = "normal", fg: str = "white") -> tk.Button:
        btn = tk.Button(
            parent, text=text, font=("Segoe UI", 10, "bold"),
            bg=bg, fg=fg, activebackground=hover, activeforeground="white",
            relief="flat", cursor="hand2", bd=0,
            command=cmd, state=state,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
        btn.bind("<Leave>", lambda e: btn.configure(bg=bg))
        return btn

    # ── actions ───────────────────────────────────────────────────────────────
    def _browse(self, var: tk.StringVar):
        p = filedialog.askdirectory(title="Select folder")
        if p:
            var.set(p)

    def _on_start(self):
        inp = self._input_var.get().strip()
        out = self._output_var.get().strip()

        if not inp or not out:
            messagebox.showwarning(
                "Missing folders",
                "Please select both:\n"
                "  • the Matched folder (PO Matcher output)\n"
                "  • an output folder for sorted projects",
            )
            return
        if not Path(inp).is_dir():
            messagebox.showerror("Not found", f"Matched folder not found:\n{inp}")
            return

        self._log_clear()
        self._log_write("Starting…\n", "muted")
        self._set_running(True)
        self._progress["value"] = 0
        self._prog_label.configure(text="Starting…")
        self._stop_event = threading.Event()

        self._thread = threading.Thread(
            target=_worker,
            args=(
                Path(inp), Path(out),
                self._move_var.get(),
                self._q, self._stop_event,
            ),
            daemon=True,
        )
        self._thread.start()

    def _on_stop(self):
        self._stop_event.set()
        self._stop_btn.configure(state="disabled")
        self._log_write("\n⚠  Stop requested — finishing current file…\n", "warn")
        self._status.configure(text="  Stopping…")

    def _open_output(self):
        out = self._output_var.get().strip()
        if not out:
            messagebox.showinfo("Output folder", "No output folder selected yet.")
            return
        p = Path(out)
        if not p.exists():
            messagebox.showinfo("Output folder", "Output folder does not exist yet.")
            return
        if sys.platform == "win32":
            os.startfile(p)
        elif sys.platform == "darwin":
            _run(["open", str(p)])
        else:
            _run(["xdg-open", str(p)])

    # ── queue polling ─────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                kind, payload = self._q.get_nowait()

                if kind == "log":
                    if payload.startswith("    →"):
                        self._log_write(payload + "\n", "project")
                    elif payload.startswith("    ✓") or "📁" in payload:
                        self._log_write(payload + "\n", "success")
                    elif "⚠" in payload or "Error" in payload:
                        self._log_write(payload + "\n", "warn")
                    elif (payload.startswith("  ")
                          or payload.startswith("──")
                          or payload.startswith("Found")):
                        self._log_write(payload + "\n", "muted")
                    else:
                        self._log_write(payload + "\n")

                elif kind == "progress":
                    current, total, label = payload
                    pct = int(current / total * 100) if total else 0
                    self._progress["value"] = pct
                    name = Path(label).name
                    self._prog_label.configure(
                        text=f"{name}  ({current}/{total})"
                    )
                    self._status.configure(text=f"  Sorting… {pct}%")

                elif kind == "done":
                    self._progress["value"] = 100
                    self._prog_label.configure(text="Complete ✓")
                    self._set_running(False)
                    self._status.configure(text="  Done.")
                    messagebox.showinfo("Done", payload)

                elif kind == "stopped":
                    self._prog_label.configure(text="Stopped")
                    self._set_running(False)
                    self._status.configure(text="  Stopped by user.")
                    self._log_write("\n■  Stopped by user.\n", "warn")

                elif kind == "error":
                    self._set_running(False)
                    self._log_write(f"\n✗  {payload}\n", "warn")
                    self._status.configure(text="  Error — see log.")
                    messagebox.showerror("Error", payload)

        except queue.Empty:
            pass
        self.after(100, self._poll)

    # ── helpers ───────────────────────────────────────────────────────────────
    def _set_running(self, running: bool):
        self._start_btn.configure(state="disabled" if running else "normal")
        self._stop_btn.configure(state="normal"    if running else "disabled")

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
#  DEPENDENCY CHECK
# ═══════════════════════════════════════════════════════════════════════════════

def _check_deps() -> list[str]:
    issues = []
    if not OCR_AVAILABLE:
        issues.append(
            "pytesseract / Pillow not installed\n"
            "  → pip install pytesseract Pillow"
        )
    if not PDFPLUMBER_AVAILABLE:
        issues.append(
            "pdfplumber not installed\n"
            "  → pip install pdfplumber"
        )
    try:
        _run(["pdftoppm", "-v"], capture_output=True, timeout=5)
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


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    issues = _check_deps()
    app = ProjectSorterApp()
    if issues:
        messagebox.showwarning(
            "Setup required",
            "Some dependencies are missing:\n\n" + "\n\n".join(issues),
        )
    app.mainloop()