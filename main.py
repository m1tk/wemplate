#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Word Print Server
=================
A tiny local print service for the "Word Template Filler" web app.

What it does
------------
* Runs a localhost-only HTTP server with a tkinter GUI
  (host / port fields, printer dropdown, temp-folder override,
  Start/Stop button, live log).
* Uploaded .docx files are buffered into a per-job temp folder. The GUI
  "Temp folder" field chooses where: the folder you set, else the Windows
  default (%TEMP%), else a 'print_jobs' folder next to this program —
  whatever is actually writable. Each candidate is probed at Start by
  writing a REAL .docx inside a sub-folder (exactly what a job does).
  Note: "you can write there but this program can't" means per-application
  blocking — Windows Defender 'Controlled folder access' or antivirus
  ransomware protection. Allow this program (wemplate.exe) in it.
* POST /print : body = raw .docx bytes
                headers: X-Filename (URI-encoded, Arabic OK), X-Copies (int)
* GET  /ping  : status JSON the web app uses for its "Connect" check
* Printing is done by Microsoft Word itself via COM automation (pywin32),
  so Arabic shaping, fonts, page breaks and copies are exactly what Word
  would print — no browser rendering involved.
* The printer is chosen HERE, in the GUI dropdown ("Windows default" or any
  installed printer). This is fully transparent to the web app: it never
  sends a printer name, it only displays what this server reports in /ping.

Security
--------
Only clients from the local machine (127.0.0.0/8, ::1) are accepted;
everyone else gets 403. CORS is open because the web app usually runs from
a file:// page, but that is safe since the client IP is still checked.

Lifecycle
---------
* The Start/Stop button stops the HTTP listener and releases the port.
* Closing the window (and atexit as a safety net) performs a graceful
  shutdown: listener closed, worker signalled, Word told to Quit, port
  released — no "port already in use" on the next start.

Usage
-----
    python word_print_server.py              # GUI
    python word_print_server.py --selftest   # headless self-test (fake printer)

Dependencies: pywin32 on Windows  ->  pip install pywin32
              tkinter (ships with the standard python.org Windows installer)
"""

from __future__ import annotations

import atexit
import getpass
import json
import os
import queue
import re
import shutil
import socket
import sys
import tempfile
import threading
import time
import urllib.parse
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_NAME              = "Word Print Server"
SERVICE_ID            = "word-print-server"
VERSION               = "1.2.0"
DOCX_MIME             = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
DEFAULT_HOST          = "127.0.0.1"
DEFAULT_PORT          = 9100
DEFAULT_PRINTER_LABEL = "Windows default"
MAX_LOG_LINES         = 1000          # log buffer cap (GUI + self-test)
MAX_BODY_BYTES        = 100 * 1024 * 1024   # 100 MB upload limit
JOB_TIMEOUT_S         = 240           # handler gives up waiting after this
CONFIG_FILE_NAME      = "word_print_server.ini"

# --- optional dependencies (script stays importable without them) ---------
try:
    import pythoncom                       # noqa: F401  (pywin32)
    import win32com.client as _w32c        # noqa: F401
    import win32print as _w32p
    PYWIN32 = True
    PYWIN32_ERROR = ""
except Exception as _exc:                  # pragma: no cover
    PYWIN32 = False
    PYWIN32_ERROR = str(_exc)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TK_OK = True
except Exception:                          # pragma: no cover
    TK_OK = False

# ---------------------------------------------------------------------------
# Shared state + logging (thread-safe: log() only enqueues)
# ---------------------------------------------------------------------------

STATE = {
    "server": None,          # PrintHTTPServer instance or None
    "server_thread": None,   # serve_forever thread
    "worker": None,          # WordWorker instance
    "printer": DEFAULT_PRINTER_LABEL,
    "temp_dir": "",          # "" = Windows default (%TEMP%); set in the GUI
    "jobs_total": 0,
    "boot_t": time.time(),
    "started_at": None,
    "app": None,             # App instance (GUI mode only)
}

LOGQ: "queue.Queue[tuple[str, str]]" = queue.Queue()

LOG_TAGS = {                # GUI text-widget tag -> foreground color
    "info":  "#1a1a1a",
    "ok":    "#0a7a2f",
    "warn":  "#a86800",
    "error": "#b00020",
    "debug": "#6b6b6b",
}


def log(msg: str, level: str = "info") -> None:
    """Queue a log line (safe to call from ANY thread)."""
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    LOGQ.put((level if level in LOG_TAGS else "info", line))


def drain_log_to_console() -> None:
    """Headless helper: print everything currently queued."""
    while True:
        try:
            level, line = LOGQ.get_nowait()
        except queue.Empty:
            return
        _console(line, error=level in ("error", "warn"))


def _first_line(text: str, limit: int = 300) -> str:
    text = (text or "").strip().replace("\r", "")
    line = text.splitlines()[0] if text else ""
    return line[:limit] if line else "unknown error"


def _console(msg: str, error: bool = False) -> None:
    """
    Console output that never crashes in headless / frozen contexts:
    PyInstaller --windowed EXEs have sys.stdout / sys.stderr set to None.
    """
    order = ("stderr", "stdout") if error else ("stdout", "stderr")
    for name in order:
        stream = getattr(sys, name, None)
        if stream is not None:
            try:
                print(msg, file=stream)
            except Exception:
                pass
            return


# ---------------------------------------------------------------------------
# Printer helpers
# ---------------------------------------------------------------------------

def list_printers() -> list:
    """Installed printer names (pywin32). Empty list when unavailable."""
    if not PYWIN32:
        return []
    try:
        flags = _w32p.PRINTER_ENUM_LOCAL | _w32p.PRINTER_ENUM_CONNECTIONS
        return sorted({p[2] for p in _w32p.EnumPrinters(flags, None, 2)})
    except Exception as exc:
        log(f"Cannot enumerate printers: {_first_line(str(exc))}", "warn")
        return []


def get_default_printer():
    """Name of the Windows default printer, or None."""
    if not PYWIN32:
        return None
    try:
        return _w32p.GetDefaultPrinter()
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Temp-folder handling (uploaded .docx files are buffered here before printing)
# ---------------------------------------------------------------------------

def get_temp_base() -> str:
    """
    Base folder that receives one 'wps_<job>_' sub-folder per print job.
    A custom folder configured in the GUI (STATE["temp_dir"]) takes
    precedence; an empty value means the Windows default (%TEMP%).
    """
    custom = (STATE.get("temp_dir") or "").strip()
    return custom if custom else tempfile.gettempdir()


def _normalize_dir(path: str) -> str:
    """
    Clean a user-typed folder path into an absolute path: trims spaces and
    the surrounding quotes Windows Explorer's "Copy as path" adds, expands
    %VARS% and ~.
    """
    return os.path.abspath(
        os.path.expandvars(
            os.path.expanduser((path or "").strip().strip('"'))))


def probe_temp_dir(path: str):
    """
    Check that *path* can host job temp folders by doing EXACTLY what a
    print job does: create a 'wps_probe_…' sub-folder, write a .docx file
    into it, delete both. (A plain .tmp probe would miss antivirus rules
    that block office documents specifically.) The folder is created when
    missing.
    Returns (ok, message): on success message is the normalized absolute
    path; on failure it is "<path>: <reason>".
    """
    try:
        path = _normalize_dir(path)
        os.makedirs(path, exist_ok=True)
        sub = tempfile.mkdtemp(prefix="wps_probe_", dir=path)
        try:
            docx = os.path.join(sub, f"probe_{uuid.uuid4().hex[:8]}.docx")
            with open(docx, "wb") as fh:
                fh.write(b"PK\x03\x04" + b"\x00" * 64)
        finally:
            shutil.rmtree(sub, ignore_errors=True)
        return True, path
    except OSError as exc:
        detail = _first_line(str(exc))
        winerr = getattr(exc, "winerror", None)
        if winerr is not None:
            detail = f"{detail} (winerror {winerr})"
        return False, f"{path}: {detail}"


def _current_user() -> str:
    """Best-effort process user name — helps diagnose permission issues."""
    try:
        return getpass.getuser()
    except Exception:                          # pragma: no cover
        return ""


def pick_temp_base(custom: str = ""):
    """
    Decide where print-job temp folders go — first candidate that is
    actually WRITABLE wins:
      1. the folder typed in the GUI ("Temp folder"), if any
      2. the Windows default temp dir (%TEMP%)
      3. a 'print_jobs' folder next to the program (script or .exe)
    Returns (chosen, tried):
      chosen = (label, verified_path) or None when nothing is writable
      tried  = [(label, path, ok, message), ...] for the log
    """
    candidates = []
    if (custom or "").strip():
        candidates.append(("Temp folder set in the GUI",
                           _normalize_dir(custom)))
    candidates.append(("Windows default temp (%TEMP%)",
                       tempfile.gettempdir()))
    try:
        if getattr(sys, "frozen", False):
            prog_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            prog_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:                          # pragma: no cover
        prog_dir = os.getcwd()
    candidates.append(("'print_jobs' next to the program",
                       os.path.join(prog_dir, "print_jobs")))
    tried = []
    for label, path in candidates:
        ok, msg = probe_temp_dir(path)
        tried.append((label, path, ok, msg))
        if ok:
            return (label, msg), tried
    return None, tried


# ---------------------------------------------------------------------------
# Word COM print worker (single dedicated thread = its own COM apartment)
# ---------------------------------------------------------------------------

class WordWorker(threading.Thread):
    """
    Serial print queue. One background thread owns the (hidden) Word
    instance for its whole life — this keeps COM apartment rules happy and
    avoids re-launching Word for every job (Word startup is the slow part).
    """

    _SENTINEL = object()

    def __init__(self, fake: bool = False):
        super().__init__(name="word-print-worker", daemon=True)
        self.fake = fake          # --selftest mode: simulate Word
        self.jobs: "queue.Queue" = queue.Queue()
        self.word = None          # COM Application object (worker thread only)
        self.active_printer = None
        self._lock = threading.Lock()

    # ---------------- public API (called from other threads) ----------------

    def submit(self, path: str, copies: int, printer: str,
               result: dict, done: threading.Event) -> None:
        """Queue a print job. Result dict + event are filled by the worker."""
        self.jobs.put((path, copies, printer, result, done))

    def stop(self) -> None:
        """Ask the worker to finish and exit (current job completes first)."""
        self.jobs.put(self._SENTINEL)

    # --------------------------- thread body --------------------------------

    def run(self) -> None:
        if PYWIN32 and not self.fake:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
        log("Print worker started (" + ("FAKE backend" if self.fake
                                        else "Microsoft Word COM backend") + ").")
        while True:
            item = self.jobs.get()
            if item is self._SENTINEL:
                break
            path, copies, printer_sel, result, done = item
            t0 = time.perf_counter()
            try:
                pages, resolved = self._print(path, copies, printer_sel)
                result["ok"] = True
                result["pages"] = pages
                result["printer"] = resolved
                result["print_ms"] = int((time.perf_counter() - t0) * 1000)
            except BaseException as exc:      # noqa: BLE001  (report, never die)
                result["ok"] = False
                result["error"] = _first_line(str(exc) or repr(exc))
                result["print_ms"] = int((time.perf_counter() - t0) * 1000)
                self._crash_word()            # fresh Word for the next job
            finally:
                try:
                    shutil.rmtree(os.path.dirname(path), ignore_errors=True)
                except Exception:
                    pass
                done.set()
        self._quit_word()
        if PYWIN32 and not self.fake:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        log("Print worker stopped.")

    # ----------------------------- internals --------------------------------

    def _ensure_word(self):
        if self.fake:
            time.sleep(0.15)                 # simulate Word startup cost
            return
        if self.word is None:
            log("Starting Microsoft Word (hidden)…")
            word = _w32c.DispatchEx("Word.Application")   # dedicated instance,
            try:                                          # never the user's copy
                word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
            except Exception:
                pass
            word.Visible = False
            word.DisplayAlerts = 0            # wdAlertsNone — no modal dialogs
            self.word = word
            self.active_printer = None

    def _print(self, path: str, copies: int, printer_sel: str):
        """Runs INSIDE the worker thread. Returns (pages, resolved_printer)."""
        self._ensure_word()

        if self.fake:
            return 3, (printer_sel or DEFAULT_PRINTER_LABEL)

        # Printer selection (transparent to the web app: chosen in the GUI)
        target = None
        if printer_sel and printer_sel != DEFAULT_PRINTER_LABEL:
            target = printer_sel
        if target and target != self.active_printer:
            log(f"Selecting printer: {target}")
            try:
                self.word.ActivePrinter = target
                self.active_printer = target
            except Exception as exc:
                self.active_printer = None
                raise RuntimeError(
                    f"Cannot select printer \"{target}\": {_first_line(str(exc))}"
                ) from exc
        if self.active_printer is None:
            try:
                self.active_printer = self.word.ActivePrinter
            except Exception:
                self.active_printer = None

        doc = self.word.Documents.Open(
            FileName=os.path.abspath(path),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            ConfirmConversions=False,
        )
        try:
            try:
                pages = int(doc.ComputeStatistics(2))     # 2 = wdStatisticPages
            except Exception:
                pages = None
            doc.PrintOut(Background=False, Range=0, Copies=copies)  # 0 = all
        finally:
            doc.Close(SaveChanges=0)                       # wdDoNotSaveChanges

        resolved = self.active_printer or target or DEFAULT_PRINTER_LABEL
        return pages, resolved

    def _crash_word(self) -> None:
        """Word failed mid-job — drop the instance so the next job starts clean."""
        self._quit_word()

    def _quit_word(self) -> None:
        with self._lock:
            word, self.word, self.active_printer = self.word, None, None
        if word is None:
            return
        try:
            word.DisplayAlerts = 0
            word.Quit(SaveChanges=0)
            log("Microsoft Word closed.")
        except Exception as exc:
            log(f"Word quit issue: {_first_line(str(exc))}", "warn")


# ---------------------------------------------------------------------------
# HTTP layer
# ---------------------------------------------------------------------------

def _sanitize_docx_name(raw: str) -> str:
    """
    Return a safe file name for logging/temp use.
    '' -> caller should use the default name
    a string ending in .docx -> ok
    anything else -> caller must reject (415)
    """
    name = (raw or "").replace("\\", "/").split("/")[-1].strip()
    if not name:
        return "document.docx"
    name = re.sub(r'[<>:"|?*\x00-\x1f]', "_", name).strip(" .")
    if len(name) > 120:
        name = name[:114] + ".docx"
    if not name.lower().endswith(".docx"):
        return ""
    return name


def _client_allowed(ip: str) -> bool:
    """Localhost-only security: 127.0.0.0/8, ::1 and v4-mapped loopback."""
    return (ip.startswith("127.") or ip == "::1"
            or ip.startswith("::ffff:127."))


class PrintHandler(BaseHTTPRequestHandler):
    server_version = "WordPrintServer/" + VERSION
    protocol_version = "HTTP/1.1"     # we always send Content-Length

    # -- silence the default stderr request logging (we log explicitly) ------
    def log_message(self, fmt, *args):     # noqa: A003
        pass

    # -- helpers --------------------------------------------------------------
    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers",
                         "Content-Type, X-Filename, X-Copies")

    def _send_json(self, payload: dict, code: int = 200,
                   close: bool = False) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self._cors()
        if close:
            self.send_header("Connection", "close")
            self.close_connection = True
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _read_exact(self, n: int):
        buf = b""
        while len(buf) < n:
            chunk = self.rfile.read(min(65536, n - len(buf)))
            if not chunk:
                return None
            buf += chunk
        return buf

    def _guard_local(self) -> bool:
        ip = self.client_address[0]
        if _client_allowed(ip):
            return True
        log(f"Rejected request from non-local client {ip} "
            f"(localhost-only policy).", "warn")
        self._send_json({"ok": False,
                         "error": "forbidden: this print server accepts "
                                  "clients on this machine only"},
                        403, close=True)
        return False

    # -- routes ----------------------------------------------------------------
    def do_OPTIONS(self):                  # noqa: N802 (stdlib naming)
        if not self._guard_local():
            return
        self.send_response(204)
        self._cors()
        self.send_header("Access-Control-Max-Age", "86400")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):                      # noqa: N802
        if not self._guard_local():
            return
        route = self.path.split("?", 1)[0]
        if route in ("/", "/ping"):
            self._send_json(self._status_payload())
            return
        self._send_json({"ok": False,
                         "error": f"GET {route} not found — use GET /ping "
                                  "or POST /print"}, 404)

    def do_POST(self):                     # noqa: N802
        t_req = time.perf_counter()
        if not self._guard_local():
            return
        route = self.path.split("?", 1)[0]
        if route != "/print":
            self._send_json({"ok": False,
                             "error": f"POST {route} not found — use POST /print"},
                            404, close=True)
            return

        # -- size checks -------------------------------------------------------
        try:
            length = int(self.headers.get("Content-Length") or "0")
        except ValueError:
            length = -1
        if length <= 0:
            self._send_json({"ok": False,
                             "error": "empty body — POST the raw .docx bytes"},
                            400, close=True)
            return
        if length > MAX_BODY_BYTES:
            self._send_json({"ok": False,
                             "error": f"file too large ({length} bytes; "
                                      f"limit {MAX_BODY_BYTES})"},
                            413, close=True)   # body intentionally not read
            return
        body = self._read_exact(length)
        if body is None:
            self._send_json({"ok": False, "error": "incomplete request body"},
                            400, close=True)
            return

        # -- filename ----------------------------------------------------------
        raw_name = self.headers.get("X-Filename") or ""
        try:
            raw_name = urllib.parse.unquote(raw_name)
        except Exception:
            pass
        fname = _sanitize_docx_name(raw_name)
        if not fname:
            self._send_json({"ok": False,
                             "error": "only .docx files are accepted "
                                      "(X-Filename must end with .docx)"},
                            415)
            return

        # -- copies --------------------------------------------------------------
        c_raw = (self.headers.get("X-Copies") or "1").strip()
        try:
            copies = int(c_raw)
        except ValueError:
            self._send_json({"ok": False,
                             "error": f"X-Copies must be an integer, got "
                                      f"\"{c_raw}\""}, 400)
            return
        copies = max(1, min(999, copies))

        # -- basic DOCX sanity (ZIP magic) --------------------------------------
        if not body.startswith(b"PK"):
            self._send_json({"ok": False,
                             "error": "body is not a .docx file "
                                      "(ZIP signature missing)"}, 400)
            return

        # -- queue the real work -------------------------------------------------
        worker = STATE.get("worker")
        if worker is None:
            self._send_json({"ok": False, "error": "print worker is not running"},
                            503)
            return
        if not PYWIN32 and not worker.fake:
            self._send_json({"ok": False,
                             "error": "pywin32 is not installed on the server "
                                      "machine — run:  pip install pywin32"},
                            500)
            return

        job_id = uuid.uuid4().hex[:8]
        temp_base = get_temp_base()
        try:
            tmpdir = tempfile.mkdtemp(prefix=f"wps_{job_id}_", dir=temp_base)
            path = os.path.join(tmpdir, f"job_{job_id}.docx")
            with open(path, "wb") as fh:
                fh.write(body)
        except OSError as exc:
            detail = _first_line(str(exc))
            winerr = getattr(exc, "winerror", None)
            if winerr is not None:
                detail = f"{detail} (winerror {winerr})"
            log(f"JOB {job_id} ✗ cannot buffer upload into temp folder \""
                f"{temp_base}\": {detail}", "error")
            log("Folder passed the Start probe, so something is blocking "
                "THIS program specifically — Windows Defender 'Controlled "
                "folder access' / antivirus ransomware protection. Allow "
                "wemplate.exe in it, or choose another folder.", "warn")
            self._send_json({
                "ok": False,
                "error": (f"cannot buffer upload into temp folder "
                          f"\"{temp_base}\" — {detail}. Fix: allow this "
                          f"program in Windows Defender 'Controlled folder "
                          f"access' / antivirus ransomware protection, or "
                          f"set another 'Temp folder' in the {APP_NAME} "
                          f"window, then Start again."),
            }, 500, close=True)
            return

        printer_sel = STATE.get("printer") or DEFAULT_PRINTER_LABEL
        result, done = {}, threading.Event()
        worker.submit(path, copies, printer_sel, result, done)
        log(f"JOB {job_id} ▶ \"{fname}\" ×{copies} → printer: {printer_sel}")

        if not done.wait(JOB_TIMEOUT_S):
            log(f"JOB {job_id} ✗ timed out after {JOB_TIMEOUT_S}s", "error")
            self._send_json({"ok": False, "job_id": job_id, "filename": fname,
                             "error": f"print timed out after {JOB_TIMEOUT_S}s — "
                                      "Word may still be spooling; check the "
                                      "printer"},
                            504)
            return

        elapsed_ms = int((time.perf_counter() - t_req) * 1000)
        print_ms = result.get("print_ms")
        wait_ms = max(0, elapsed_ms - print_ms) if print_ms is not None else None
        STATE["jobs_total"] += 1

        if result.get("ok"):
            pages = result.get("pages")
            log(f"JOB {job_id} ✔ printed "
                f"{pages if pages is not None else '?'} page(s) ×{copies} on "
                f"\"{result.get('printer')}\" in {(print_ms or 0) / 1000:.1f}s "
                f"(waited {(wait_ms or 0) / 1000:.1f}s in queue)")
            self._send_json({
                "ok": True,
                "job_id": job_id,
                "filename": fname,
                "printer": result.get("printer"),
                "copies": copies,
                "pages": pages,
                "elapsed_ms": elapsed_ms,
                "print_ms": print_ms,
                "wait_ms": wait_ms,
                "service": SERVICE_ID,
                "version": VERSION,
            })
        else:
            log(f"JOB {job_id} ✗ FAILED: {result.get('error')}", "error")
            self._send_json({
                "ok": False,
                "job_id": job_id,
                "filename": fname,
                "error": result.get("error", "unknown printing error"),
                "elapsed_ms": elapsed_ms,
                "print_ms": print_ms,
                "service": SERVICE_ID,
            }, 500)

    # -- status -----------------------------------------------------------------
    def _status_payload(self) -> dict:
        sel = STATE.get("printer") or DEFAULT_PRINTER_LABEL
        default = get_default_printer()
        worker = STATE.get("worker")
        backend_ok = PYWIN32 or bool(worker is not None and worker.fake)
        return {
            "ok": True,
            "service": SERVICE_ID,
            "version": VERSION,
            "app": APP_NAME,
            "backend": ("Microsoft Word (COM automation)" if backend_ok
                        else "UNAVAILABLE — pip install pywin32"),
            "printer": sel,
            "resolved_printer": (sel if sel != DEFAULT_PRINTER_LABEL
                                 else (default or DEFAULT_PRINTER_LABEL)),
            "default_printer": default,
            "printers": list_printers(),
            "temp_dir": get_temp_base(),
            "temp_dir_custom": bool((STATE.get("temp_dir") or "").strip()),
            "jobs_total": STATE.get("jobs_total", 0),
            "uptime_s": int(time.time() - STATE["boot_t"]),
            "localhost_only": True,
            "endpoints": {
                "ping": "GET /ping",
                "print": "POST /print — raw .docx bytes; "
                         "headers: X-Filename (URI-encoded), X-Copies (int)",
            },
        }


class PrintHTTPServer(ThreadingHTTPServer):
    daemon_threads = True          # in-flight requests never block exit
    allow_reuse_address = True     # instant re-bind after a graceful stop

    def handle_error(self, request, client_address):   # noqa: D102
        exc = sys.exc_info()[1]
        if isinstance(exc, (ConnectionResetError, BrokenPipeError,
                            socket.timeout, TimeoutError)):
            return
        log(f"Request error from {client_address[0]}: {_first_line(str(exc))}",
            "error")


# ---------------------------------------------------------------------------
# Server / worker lifecycle (shared by GUI, atexit and self-test)
# ---------------------------------------------------------------------------

def start_server(host: str, port: int):
    """Bind + serve. Returns the server object, or None on bind failure."""
    stop_server()                                   # just in case
    try:
        srv = PrintHTTPServer((host, port), PrintHandler)
    except OSError as exc:
        log(f"Cannot bind {host}:{port} — {_first_line(str(exc))}", "error")
        return None
    thread = threading.Thread(target=srv.serve_forever,
                              kwargs={"poll_interval": 0.2},
                              name="http-server", daemon=True)
    STATE["server"], STATE["server_thread"] = srv, thread
    STATE["started_at"] = time.time()
    thread.start()
    log(f"HTTP server listening on http://{host}:{port}"
        + ("" if host.startswith("127.") else "   (Windows may ask about "
                                             "firewall access — safe to deny; "
                                             "clients are localhost-only)"))
    log("Endpoints:  GET /ping   ·   POST /print (raw .docx bytes)")
    return srv


def stop_server() -> None:
    """Graceful stop: shutdown loop, close socket, join thread → port released."""
    srv, STATE["server"] = STATE.get("server"), None
    thread, STATE["server_thread"] = STATE.get("server_thread"), None
    if srv is None:
        return
    try:
        srv.shutdown()            # stops serve_forever (called from other thread)
        srv.server_close()        # closes the listening socket
        if thread is not None:
            thread.join(timeout=2)
        log("HTTP server stopped — port released.")
    except Exception as exc:
        log(f"Server stop issue: {_first_line(str(exc))}", "warn")


def start_worker(fake: bool = False) -> None:
    if STATE.get("worker") is not None:
        return
    worker = WordWorker(fake=fake)
    STATE["worker"] = worker
    worker.start()
    if not PYWIN32 and not fake:
        log("pywin32 is missing — printing is DISABLED. Fix with:  "
            "pip install pywin32", "error")


def stop_worker() -> None:
    worker, STATE["worker"] = STATE.get("worker"), None
    if worker is None:
        return
    worker.stop()
    worker.join(timeout=5)
    if worker.is_alive():
        log("Word is still finishing a job — if WINWORD.EXE lingers in Task "
            "Manager, close it manually.", "warn")


_CLEANUP_DONE = False
_CLEANUP_LOCK = threading.Lock()


def cleanup() -> None:
    """
    Idempotent graceful shutdown — registered with atexit AND called after
    the GUI mainloop returns. Guarantees: listener closed & port released,
    print worker signalled, Word asked to quit.
    """
    global _CLEANUP_DONE
    with _CLEANUP_LOCK:
        if _CLEANUP_DONE:
            return
        _CLEANUP_DONE = True
    try:
        stop_server()
    except Exception as exc:                   # pragma: no cover
        _console(f"[cleanup] stop_server: {exc}", error=True)
    try:
        stop_worker()
    except Exception as exc:                   # pragma: no cover
        _console(f"[cleanup] stop_worker: {exc}", error=True)


atexit.register(cleanup)

# ---------------------------------------------------------------------------
# Config persistence (remembers host / port / printer / temp folder)
# ---------------------------------------------------------------------------

def _config_path() -> str:
    if getattr(sys, "frozen", False):
        # Frozen (PyInstaller) exe: __file__ lives in a throwaway extraction
        # folder — keep the settings file next to the real .exe instead.
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except NameError:                      # pragma: no cover
            base = os.getcwd()
    return os.path.join(base, CONFIG_FILE_NAME)


def load_config() -> dict:
    cfg = {}
    try:
        import configparser
        parser = configparser.ConfigParser()
        if parser.read(_config_path(), encoding="utf-8"):
            cfg["host"] = parser.get("server", "host", fallback="").strip()
            cfg["port"] = parser.get("server", "port", fallback="").strip()
            cfg["printer"] = parser.get("printer", "name", fallback="").strip()
            cfg["temp_dir"] = parser.get("server", "temp_dir",
                                         fallback="").strip()
    except Exception:
        pass
    return cfg


def save_config(host: str, port: str, printer: str,
                temp_dir: str = "") -> None:
    try:
        import configparser
        parser = configparser.ConfigParser()
        parser["server"] = {"host": host, "port": port,
                            "temp_dir": temp_dir}
        parser["printer"] = {"name": printer}
        with open(_config_path(), "w", encoding="utf-8") as fh:
            parser.write(fh)
    except Exception as exc:
        log(f"Cannot save settings: {_first_line(str(exc))}", "warn")


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class App:
    """tkinter front-end: host/port, printer dropdown, temp folder,
    start/stop, log."""

    def __init__(self, root: "tk.Tk"):
        self.root = root
        self.running = False
        self._nlines = 0

        cfg = load_config()
        # "last started" values — also the fallback persisted on close, so
        # settings survive even if the user closes without pressing Start.
        self._last_host = cfg.get("host") or DEFAULT_HOST
        self._last_port = cfg.get("port") or str(DEFAULT_PORT)
        self._last_temp = cfg.get("temp_dir") or ""

        self.host_var = tk.StringVar(value=self._last_host)
        self.port_var = tk.StringVar(value=self._last_port)
        self.temp_var = tk.StringVar(value=self._last_temp)
        self.printer_var = tk.StringVar(value=DEFAULT_PRINTER_LABEL)

        root.title(f"{APP_NAME}  v{VERSION}")
        root.minsize(780, 460)
        self._build_ui()

        saved_printer = cfg.get("printer")
        if saved_printer:
            self.printer_var.set(saved_printer)
        self.refresh_printers(quiet=not PYWIN32)
        if saved_printer and saved_printer in self.printer_cb["values"]:
            self.printer_var.set(saved_printer)
        STATE["printer"] = self.printer_var.get()

        start_worker(fake=False)          # one worker for the whole session
        self._pump_log()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        log(f"{APP_NAME} {VERSION} ready."
            + ("" if PYWIN32 else "  Printing unavailable until pywin32 is installed."))

    # ----------------------------- UI building ------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 4}

        top = ttk.LabelFrame(self.root, text="Server")
        top.pack(fill="x", **pad)

        row1 = ttk.Frame(top); row1.pack(fill="x", padx=6, pady=4)
        ttk.Label(row1, text="Host:").pack(side="left")
        self.host_entry = ttk.Entry(row1, textvariable=self.host_var, width=14)
        self.host_entry.pack(side="left", padx=(2, 12))
        ttk.Label(row1, text="Port:").pack(side="left")
        self.port_entry = ttk.Entry(row1, textvariable=self.port_var, width=7)
        self.port_entry.pack(side="left", padx=(2, 12))
        ttk.Label(row1, text="Printer:").pack(side="left")
        self.printer_cb = ttk.Combobox(row1, textvariable=self.printer_var,
                                       state="readonly", width=34)
        self.printer_cb.pack(side="left", padx=(2, 6))
        self.printer_cb.bind("<<ComboboxSelected>>", self.on_printer_change)
        ttk.Button(row1, text="↻ Refresh printers",
                   command=self.refresh_printers).pack(side="left")

        # -- temp folder row: where uploads are buffered before printing. -----
        # The Windows default (%TEMP%) is not writable on some machines —
        # point this at any writable folder (e.g. C:\Temp) to fix
        # "cannot buffer upload: permission denied".
        row_t = ttk.Frame(top); row_t.pack(fill="x", padx=6, pady=(0, 4))
        ttk.Label(row_t, text="Temp folder:").pack(side="left")
        self.temp_entry = ttk.Entry(row_t, textvariable=self.temp_var)
        self.temp_entry.pack(side="left", fill="x", expand=True, padx=(2, 6))
        self.temp_browse_btn = ttk.Button(row_t, text="Browse…", width=10,
                                          command=self.browse_temp)
        self.temp_browse_btn.pack(side="left")
        self.temp_reset_btn = ttk.Button(row_t, text="Use Windows default",
                                         width=18, command=self.reset_temp)
        self.temp_reset_btn.pack(side="left", padx=(6, 0))
        ttk.Label(row_t, foreground="#888",
                  text="empty = Windows default (%TEMP%)").pack(
            side="left", padx=(6, 0))

        row2 = ttk.Frame(top); row2.pack(fill="x", padx=6, pady=(0, 6))
        self.start_btn = ttk.Button(row2, text="▶ Start server",
                                    command=self.on_toggle)
        self.start_btn.pack(side="left")
        self.status_lbl = ttk.Label(row2, text="○ Stopped", foreground="#777")
        self.status_lbl.pack(side="left", padx=12)
        ttk.Label(row2, foreground="#888",
                  text="clients: this machine only · POST /print (.docx) · "
                       "GET /ping").pack(side="right")

        logframe = ttk.LabelFrame(self.root, text="Log")
        logframe.pack(fill="both", expand=True, **pad)

        bar = ttk.Frame(logframe); bar.pack(fill="x", padx=6, pady=(4, 0))
        ttk.Button(bar, text="Clear log", width=10,
                   command=self.clear_log).pack(side="left")
        ttk.Button(bar, text="Save log as…", width=14,
                   command=self.save_log).pack(side="left", padx=6)
        self.line_count_lbl = ttk.Label(bar, foreground="#888",
                                        text=f"0 / {MAX_LOG_LINES} lines")
        self.line_count_lbl.pack(side="right")

        wrap = ttk.Frame(logframe); wrap.pack(fill="both", expand=True,
                                              padx=6, pady=(2, 6))
        self.log_text = tk.Text(wrap, height=16, wrap="word", state="disabled",
                                relief="flat", background="#fafafa",
                                font=("Consolas", 9))
        ysb = ttk.Scrollbar(wrap, orient="vertical",
                            command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=ysb.set)
        ysb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        for tag, color in LOG_TAGS.items():
            self.log_text.tag_configure(tag, foreground=color)

    # ------------------------------ log pump --------------------------------

    def _append_log(self, level: str, line: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n", level)
        self._nlines += 1
        excess = self._nlines - MAX_LOG_LINES
        if excess > 0:                       # hard cap → bounded memory
            self.log_text.delete("1.0", f"{excess}.0")
            self._nlines -= excess
        self.log_text.configure(state="disabled")
        self.log_text.see("end")
        self.line_count_lbl.configure(
            text=f"{self._nlines} / {MAX_LOG_LINES} lines")

    def _pump_log(self) -> None:
        try:
            while True:
                level, line = LOGQ.get_nowait()
                self._append_log(level, line)
        except queue.Empty:
            pass
        self.root.after(100, self._pump_log)   # tkinter updates stay on main thread

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self._nlines = 0
        self.line_count_lbl.configure(text=f"0 / {MAX_LOG_LINES} lines")
        log("Log cleared.")

    def save_log(self) -> None:
        name = time.strftime("wordprintserver_%Y%m%d_%H%M%S.log")
        path = filedialog.asksaveasfilename(
            defaultextension=".log", initialfile=name,
            filetypes=[("Log files", "*.log"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(self.log_text.get("1.0", "end"))
            log(f"Log saved to {path}", "ok")
        except OSError as exc:
            log(f"Cannot save log: {_first_line(str(exc))}", "error")

    # ---------------------------- temp folder --------------------------------

    def browse_temp(self) -> None:
        """Pick a folder for print-job temp files (upload buffering)."""
        chosen = filedialog.askdirectory(
            title="Choose a writable folder for print-job temporary files",
            initialdir=self.temp_var.get().strip() or tempfile.gettempdir())
        if chosen:
            self.temp_var.set(os.path.normpath(chosen))

    def reset_temp(self) -> None:
        """Back to the Windows default temp folder (%TEMP%)."""
        self.temp_var.set("")

    # ------------------------------ actions ---------------------------------

    def refresh_printers(self, quiet: bool = False) -> None:
        plist = list_printers()
        values = [DEFAULT_PRINTER_LABEL] + plist
        self.printer_cb["values"] = values
        if self.printer_var.get() not in values:
            self.printer_var.set(DEFAULT_PRINTER_LABEL)
        STATE["printer"] = self.printer_var.get()
        if quiet:
            return
        if not PYWIN32:
            log("Printer list unavailable (pywin32 not installed).", "warn")
            return
        default = get_default_printer()
        log(f"Printer list refreshed — {len(plist)} installed printer(s)"
            + (f", system default: \"{default}\"" if default else ""))

    def on_printer_change(self, _event=None) -> None:
        STATE["printer"] = self.printer_var.get() or DEFAULT_PRINTER_LABEL
        log(f"Printer for new jobs: {STATE['printer']}")
        if self.running:
            save_config(self._last_host, self._last_port, STATE["printer"],
                        self._last_temp)

    def on_toggle(self) -> None:
        if self.running:
            self.stop()
        else:
            self.start()

    def start(self) -> None:
        host = (self.host_var.get() or "").strip() or DEFAULT_HOST
        try:
            port = int(self.port_var.get())
            if not 1 <= port <= 65535:
                raise ValueError
        except ValueError:
            log("Port must be an integer between 1 and 65535.", "error")
            return

        # -- temp folder: pick a WRITABLE location BEFORE binding, so a bad
        #    folder never produces "cannot buffer upload" mid-printing.
        #    Order: GUI setting -> Windows default -> beside the program.
        temp_custom = self.temp_var.get().strip()
        chosen, tried = pick_temp_base(temp_custom)
        if chosen is None:
            for label, path, ok, msg in tried:
                log(f"Temp folder NOT usable — {label} ({path}): {msg}",
                    "error")
            log("No writable temp folder found. This machine lets YOU write "
                "but blocks THIS program from writing documents — that is "
                "per-application protection: Windows Defender 'Controlled "
                "folder access' or your antivirus' ransomware shield. Add "
                "this program (wemplate.exe) to its allowed apps, then "
                "Start again.", "error")
            return
        if temp_custom and chosen[1] != _normalize_dir(temp_custom):
            fail_msg = next((msg for _l, _p, ok, msg in tried if not ok), "")
            log(f"Your Temp folder is NOT writable for this program — "
                f"{fail_msg}", "error")
            log("YOU can write there, but THIS program can't — per-app "
                "blocking (Windows Defender 'Controlled folder access' / "
                "antivirus ransomware protection). Allow wemplate.exe in "
                "it, or pick a folder outside Documents/Desktop.", "warn")
            log(f"Falling back to: {chosen[0]} -> {chosen[1]}", "warn")
        else:
            for label, path, ok, msg in tried:
                if ok:
                    log(f"Temp folder candidate: {label} -> {path}", "debug")
        user = _current_user()
        STATE["temp_dir"] = chosen[1]        # effective folder for jobs
        self._last_temp = temp_custom        # keep what the user typed
        log(f"Temp folder: {chosen[1]}  (verified writable"
            + (f"; program runs as '{user}'" if user else "") + ")", "ok")

        STATE["printer"] = self.printer_var.get() or DEFAULT_PRINTER_LABEL
        if start_server(host, port) is None:
            self._set_running(False)
            return
        self._last_host, self._last_port = host, str(port)
        save_config(host, str(port), STATE["printer"], temp_custom)
        self._set_running(True)
        extra = ""
        if PYWIN32:
            default = get_default_printer()
            if STATE["printer"] == DEFAULT_PRINTER_LABEL and default:
                extra = f"  (system default: \"{default}\")"
        else:
            extra = "   [printing DISABLED — pywin32 missing]"
        log(f"Printer for new jobs: {STATE['printer']}{extra}", "ok")

    def stop(self) -> None:
        stop_server()
        self._set_running(False)

    def _set_running(self, running: bool) -> None:
        self.running = running
        if running:
            host = (self.host_var.get() or "").strip() or DEFAULT_HOST
            port = (self.port_var.get() or "").strip()
            self.status_lbl.configure(
                text=f"● Running — http://{host}:{port}   (localhost clients only)",
                foreground="#0a7a2f")
            self.start_btn.configure(text="■ Stop server")
        else:
            self.status_lbl.configure(text="○ Stopped", foreground="#777")
            self.start_btn.configure(text="▶ Start server")
        state = "disabled" if running else "normal"
        self.host_entry.configure(state=state)
        self.port_entry.configure(state=state)
        self.temp_entry.configure(state=state)
        self.temp_browse_btn.configure(state=state)
        self.temp_reset_btn.configure(state=state)

    def on_close(self) -> None:
        try:
            if self.running and not messagebox.askokcancel(
                    APP_NAME, "The server is running.\n\nStop it and exit?"):
                return
        except Exception:
            pass
        save_config(self._last_host, self._last_port, STATE["printer"],
                    self._last_temp)
        self.root.destroy()      # cleanup() runs via atexit / after mainloop


# ---------------------------------------------------------------------------
# Self-test (headless; exercises HTTP + security with a FAKE printer)
# ---------------------------------------------------------------------------

def _free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _nonloopback_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
        s.close()
        return None if ip.startswith("127.") else ip
    except Exception:
        return None


def run_selftest() -> int:
    import http.client

    print(f"== {APP_NAME} {VERSION} self-test ==")
    STATE["printer"] = DEFAULT_PRINTER_LABEL
    start_worker(fake=True)

    results = []

    def check(name, cond, extra=""):
        results.append(bool(cond))
        status = "PASS" if cond else "FAIL"
        tail = f"   [{extra}]" if (extra and not cond) else ""
        print(f"  {status}  {name}{tail}")

    def request(method, path, headers=None, body=None, port=None, host="127.0.0.1"):
        conn = http.client.HTTPConnection(host, port, timeout=15)
        conn.request(method, path, body=body, headers=headers or {})
        resp = conn.getresponse()
        data = resp.read()
        conn.close()
        return resp, data

    port = _free_port()
    srv = start_server("127.0.0.1", port)
    check("server starts and binds", srv is not None)
    check("server restarts on the same port (graceful release)",
          (stop_server(), start_server("127.0.0.1", port))[1] is not None)

    fake_docx = b"PK\x03\x04" + b"\x00" * 64
    arabic_name = urllib.parse.quote("تقرير_مبيعات filled.docx")

    # 1. ping
    r, data = request("GET", "/ping", port=port)
    payload = json.loads(data)
    check("GET /ping → 200 with service info",
          r.status == 200 and payload.get("ok") and
          payload.get("service") == SERVICE_ID and
          isinstance(payload.get("printers"), list))

    # 2. CORS preflight
    r, data = request("OPTIONS", "/print", port=port)
    check("OPTIONS /print → 204 + CORS headers",
          r.status == 204 and
          r.getheader("Access-Control-Allow-Origin") == "*" and
          "POST" in (r.getheader("Access-Control-Allow-Methods") or "") and
          "X-Filename" in (r.getheader("Access-Control-Allow-Headers") or ""))

    # 3. successful print (fake Word)
    r, data = request("POST", "/print", port=port,
                      headers={"Content-Type": DOCX_MIME,
                               "X-Filename": arabic_name,
                               "X-Copies": "2"},
                      body=fake_docx)
    payload = json.loads(data)
    check("POST /print (Arabic filename, 2 copies) → printed",
          r.status == 200 and payload.get("ok") and
          payload.get("copies") == 2 and payload.get("pages") == 3 and
          payload.get("printer") == DEFAULT_PRINTER_LABEL and
          payload.get("filename") == "تقرير_مبيعات filled.docx" and
          isinstance(payload.get("job_id"), str) and
          isinstance(payload.get("print_ms"), int) and
          payload.get("wait_ms") is not None,
          extra=json.dumps(payload, ensure_ascii=False)[:200])
    job_dir_re = re.compile(r"^wps_[0-9a-f]{8}_.")   # per-job dirs only
    check("temp job folders cleaned up",
          not [n for n in os.listdir(get_temp_base())
               if job_dir_re.match(n)])

    # 4. copies clamp
    r, data = request("POST", "/print", port=port,
                      headers={"X-Filename": "a.docx", "X-Copies": "5000"},
                      body=fake_docx)
    check("X-Copies clamped to 999",
          json.loads(data).get("copies") == 999)

    # 5. bad magic
    r, data = request("POST", "/print", port=port,
                      headers={"X-Filename": "a.docx"}, body=b"not a zip")
    check("non-ZIP body rejected (400)", r.status == 400)

    # 6. wrong extension
    r, data = request("POST", "/print", port=port,
                      headers={"X-Filename": "a.pdf"}, body=fake_docx)
    check("non-.docx extension rejected (415)", r.status == 415)

    # 7. bad copies
    r, data = request("POST", "/print", port=port,
                      headers={"X-Filename": "a.docx", "X-Copies": "two"},
                      body=fake_docx)
    check("non-integer X-Copies rejected (400)", r.status == 400)

    # 8. empty body
    r, data = request("POST", "/print", port=port,
                      headers={"X-Filename": "a.docx"})
    check("empty body rejected (400)", r.status == 400)

    # 9. 404
    r, data = request("GET", "/nope", port=port)
    check("unknown route → 404 JSON", r.status == 404)

    # 10. oversize Content-Length (raw socket; body deliberately short)
    s = socket.create_connection(("127.0.0.1", port), timeout=10)
    s.sendall(b"POST /print HTTP/1.1\r\nHost: t\r\n"
              b"Content-Length: 209715200\r\n\r\nXXXX")
    first_line = s.recv(4096).split(b"\r\n", 1)[0]
    s.close()
    check("oversize body rejected (413)", b"413" in first_line,
          extra=first_line.decode(errors="replace"))

    # 11. non-local client (needs a real non-loopback interface)
    foreign_ip = _nonloopback_ip()
    if foreign_ip:
        port2 = _free_port()
        srv2 = start_server("0.0.0.0", port2)
        if srv2:
            try:
                r, data = request("GET", "/ping", port=port2, host=foreign_ip)
                check("non-local client rejected (403)", r.status == 403)
            finally:
                stop_server()
    else:
        print("  SKIP  non-local client rejected (no non-loopback interface)")

    # 12. custom temp folder (the GUI "Temp folder" setting, server-side)
    custom_dir = tempfile.mkdtemp(prefix="wps_custom_probe_")
    STATE["temp_dir"] = custom_dir
    ok, msg = probe_temp_dir(custom_dir)
    check("probe_temp_dir accepts a writable folder",
          ok and msg == os.path.abspath(custom_dir), extra=msg)
    port3 = _free_port()
    if start_server("127.0.0.1", port3):
        r, data = request("POST", "/print", port=port3,
                          headers={"X-Filename": "custom.docx"},
                          body=fake_docx)
        check("print job buffered in the custom temp folder",
              r.status == 200 and json.loads(data).get("ok"))
        r, data = request("GET", "/ping", port=port3)
        check("ping reports the effective temp folder",
              json.loads(data).get("temp_dir") == os.path.abspath(custom_dir))
        check("custom temp folder cleaned after the job",
              not os.listdir(custom_dir))
        stop_server()
    STATE["temp_dir"] = ""
    shutil.rmtree(custom_dir, ignore_errors=True)

    # 13. probe rejects an unusable folder (target path is inside a file)
    blocker_dir = tempfile.mkdtemp(prefix="wps_probe_reject_")
    blocker = os.path.join(blocker_dir, "file.txt")
    with open(blocker, "w", encoding="utf-8") as fh:
        fh.write("not a folder")
    ok, msg = probe_temp_dir(os.path.join(blocker, "sub"))
    check("probe_temp_dir rejects an unusable folder", not ok, extra=msg)
    chosen, _tried = pick_temp_base(os.path.join(blocker, "sub"))
    check("temp base auto-falls-back when the custom folder is unusable",
          chosen is not None and
          chosen[0] != "Temp folder set in the GUI",
          extra=str(chosen))
    shutil.rmtree(blocker_dir, ignore_errors=True)

    stop_server()
    stop_worker()
    cleanup()
    drain_log_to_console()

    passed = sum(results)
    total = len(results)
    print(f"\n  {passed}/{total} checks passed.")
    print("Self-test OK." if passed == total else "Self-test FAILED!")
    return 0 if passed == total else 1


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--selftest" in argv:
        return run_selftest()
    if not TK_OK:
        _console("tkinter is not available in this Python installation; "
                 "cannot open the GUI.", error=True)
        _console("Try:  python word_print_server.py --selftest", error=True)
        return 2
    root = tk.Tk()
    app = App(root)
    STATE["app"] = app
    try:
        root.mainloop()
    finally:
        cleanup()                # graceful stop on every exit path (plus atexit)
    return 0


if __name__ == "__main__":
    sys.exit(main())
