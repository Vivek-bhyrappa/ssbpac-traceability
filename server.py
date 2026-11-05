#!/usr/bin/env python3
"""
SSBPAC Traceability Generator Web Application Server
===================================================
A self-contained web application powered by Python's standard library.
Direct disk-level processing with ZERO file uploads:
  - Reads folders and documents directly from disk at full SSD speed.
  - Native Windows Folder and File Choosers (opens native OS dialog when clicked).
  - Dynamically builds SWRS Traceability Reports based on active stages:
      * SWRS Requirements Template (.xlsx) [MANDATORY]
      * SHI Word Reports Folder (.docx) [OPTIONAL]
      * Overall System Test Word Reports Folder (.docx) [OPTIONAL]
      * Software Component Testing TCF Folder (.tcf) [OPTIONAL]
      * Software Integration Testing TCF Folder (.tcf) [OPTIONAL]

Supports:
  * SWRS vs SHI only (6 columns)
  * SWRS vs Full Suite (16 columns across all 4 stages)
  * Any custom combination of test stages
"""

import os
import sys
import json
import time
import shutil
import urllib.parse
import threading
import subprocess
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import generator_core

DEFAULT_CONFIG = {
    "input_excel": os.path.join(PARENT_DIR, "overall_traceability.xlsx"),
    "unit_dir": os.path.join(PARENT_DIR, "Unit testing"),
    "integ_dir": os.path.join(PARENT_DIR, "Integration testing"),
    "shi_dir": os.path.join(PARENT_DIR, "SHI"),
    "overall_dir": os.path.join(PARENT_DIR, "Overall Software Test report"),
    "output_excel": os.path.join(CURRENT_DIR, "overall_traceability_report.xlsx")
}

task_state = {
    "is_running": False,
    "progress": 0,
    "message": "Idle",
    "logs": [],
    "result": None,
    "error": None,
    "active_mode": "SWRS vs Full Suite"
}
state_lock = threading.Lock()


def count_files(dir_path, extension):
    """Counts files ending with given extension recursively."""
    if not dir_path or not os.path.exists(dir_path):
        return 0, False
    if os.path.isfile(dir_path):
        return 1 if dir_path.lower().endswith(extension.lower()) else 0, True
    cnt = 0
    for root, _, files in os.walk(dir_path):
        for f in files:
            if f.lower().endswith(extension.lower()) and not f.startswith("~$"):
                cnt += 1
    return cnt, True


def ask_folder_dialog(title="Select Folder", initial_dir=None):
    """Opens native Windows folder selection dialog in a dedicated clean subprocess."""
    py_code = (
        "import sys, os\n"
        "try:\n"
        "    import tkinter as tk\n"
        "    from tkinter import filedialog\n"
        "    root = tk.Tk()\n"
        "    root.withdraw()\n"
        "    root.attributes('-topmost', True)\n"
        f"    p = filedialog.askdirectory(title={repr(title)}, initialdir={repr(initial_dir or PARENT_DIR)})\n"
        "    root.destroy()\n"
        "    if p:\n"
        "        sys.stdout.write(p)\n"
        "except Exception as e:\n"
        "    sys.stderr.write(str(e))\n"
    )
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        res = subprocess.run(
            [sys.executable, "-c", py_code],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=flags
        )
        folder = res.stdout.strip()
        if folder:
            return folder
    except Exception as e:
        sys.stderr.write(f"Subprocess folder picker error: {e}\n")

    if os.name == 'nt':
        try:
            ps_code = (
                "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; "
                "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
                f"$f.Description = '{title}'; "
                "$f.ShowNewFolderButton = $true; "
                "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.SelectedPath }"
            )
            res = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_code],
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            folder = res.stdout.strip()
            if folder:
                return folder
        except Exception:
            pass

    return ""


def ask_file_dialog(title="Select Master SWRS Excel Template", initial_dir=None):
    """Opens native Windows file selection dialog in a dedicated clean subprocess."""
    py_code = (
        "import sys, os\n"
        "try:\n"
        "    import tkinter as tk\n"
        "    from tkinter import filedialog\n"
        "    root = tk.Tk()\n"
        "    root.withdraw()\n"
        "    root.attributes('-topmost', True)\n"
        f"    p = filedialog.askopenfilename(title={repr(title)}, initialdir={repr(initial_dir or PARENT_DIR)}, filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')])\n"
        "    root.destroy()\n"
        "    if p:\n"
        "        sys.stdout.write(p)\n"
        "except Exception as e:\n"
        "    sys.stderr.write(str(e))\n"
    )
    try:
        flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        res = subprocess.run(
            [sys.executable, "-c", py_code],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=flags
        )
        file_path = res.stdout.strip()
        if file_path:
            return file_path
    except Exception as e:
        sys.stderr.write(f"Subprocess file picker error: {e}\n")

    if os.name == 'nt':
        try:
            ps_code = (
                "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; "
                "$f = New-Object System.Windows.Forms.OpenFileDialog; "
                f"$f.Title = '{title}'; "
                "$f.Filter = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'; "
                "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.FileName }"
            )
            res = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_code],
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            file_path = res.stdout.strip()
            if file_path:
                return file_path
        except Exception:
            pass

    return ""


class TraceabilityHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        if len(args) >= 3:
            sys.stdout.write(f"[{time.strftime('%H:%M:%S')}] {args[0]} {args[1]} -> {args[2]}\n")
        elif len(args) > 0:
            sys.stdout.write(f"[{time.strftime('%H:%M:%S')}] {format % args}\n")
        else:
            sys.stdout.write(f"[{time.strftime('%H:%M:%S')}] {format}\n")

    def send_json(self, data, status_code=200):
        body = json.dumps(data, indent=2).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_HEAD(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
            return

        if path in ["/", "/index.html"]:
            self.serve_ui()
        elif path == "/api/defaults":
            is_local = self.client_address[0] in ("127.0.0.1", "::1", "localhost")
            data = dict(DEFAULT_CONFIG)
            data["is_local"] = is_local
            self.send_json(data)
        elif path == "/api/progress":
            with state_lock:
                self.send_json(dict(task_state))
        elif path == "/api/download":
            self.serve_download()
        elif path == "/api/browse-folder":
            stage = query.get("stage", [""])[0]
            self.handle_browse_folder(stage)
        elif path == "/api/browse-file":
            stage = query.get("stage", [""])[0]
            self.handle_browse_file(stage)
        else:
            self.send_error(404, "Not Found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        content_len = int(self.headers.get("Content-Length", 0))

        if path == "/api/upload":
            self.handle_upload(query, content_len)
            return

        post_body = self.rfile.read(content_len).decode("utf-8") if content_len > 0 else "{}"
        try:
            req_data = json.loads(post_body)
        except Exception:
            req_data = {}

        if path == "/api/generate":
            self.handle_generate(req_data)
        elif path == "/api/verify":
            self.handle_verify(req_data)
        elif path == "/api/open-folder":
            self.handle_open_folder(req_data)
        else:
            self.send_error(404, "Not Found")

    def handle_upload(self, query, content_len):
        try:
            stage = query.get("stage", ["swrs"])[0]
            raw_name = query.get("filename", ["uploaded_requirements.xlsx"])[0]
            subfolder = query.get("folder", [""])[0]
            rel_path = query.get("relpath", [""])[0]
            filename = os.path.basename(raw_name)

            if subfolder:
                base_upload = os.path.join(CURRENT_DIR, "uploads", os.path.basename(subfolder))
                if rel_path:
                    # Clean rel_path to prevent path traversal while preserving subdirectories
                    clean_rel = os.path.normpath(rel_path).lstrip("\\/").replace("..", "")
                    dest_path = os.path.join(CURRENT_DIR, "uploads", clean_rel)
                    uploads_dir = os.path.dirname(dest_path)
                else:
                    uploads_dir = base_upload
                    dest_path = os.path.join(uploads_dir, filename)
            else:
                base_upload = os.path.join(CURRENT_DIR, "uploads")
                uploads_dir = base_upload
                dest_path = os.path.join(uploads_dir, filename)

            os.makedirs(uploads_dir, exist_ok=True)

            remaining = content_len
            with open(dest_path, "wb") as f:
                while remaining > 0:
                    chunk_size = min(remaining, 65536)
                    chunk = self.rfile.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    remaining -= len(chunk)

            ext = ".xlsx" if stage == "swrs" else (".docx" if stage in ["shi", "sys"] else ".tcf")
            target_check = base_upload if subfolder else dest_path
            cnt, ok = count_files(target_check, ext)

            self.send_json({
                "status": "success",
                "path": target_check,
                "folder_path": base_upload if subfolder else os.path.dirname(dest_path),
                "name": filename if not subfolder else os.path.basename(subfolder) + "/",
                "count": cnt,
                "exists": ok
            })
        except Exception as e:
            self.send_json({
                "status": "error",
                "message": f"Upload processing error: {str(e)}"
            }, status_code=500)

    def handle_browse_folder(self, stage):
        titles = {
            "swrs": "Select Master SWRS Requirements Folder",
            "shi": "Select Complete SHI Verification Folder",
            "sys": "Select Overall System Software Test Folder",
            "comp": "Select Software Component (Unit Testing) Folder",
            "integ": "Select Software Integration Testing Folder"
        }
        exts = {
            "swrs": ".xlsx",
            "shi": ".docx",
            "sys": ".docx",
            "comp": ".tcf",
            "integ": ".tcf"
        }
        title = titles.get(stage, "Select Folder")
        folder = ask_folder_dialog(title=title)

        if not folder:
            self.send_json({"status": "cancelled"})
            return

        ext = exts.get(stage, "")
        cnt, ok = count_files(folder, ext) if ext else (0, True)

        self.send_json({
            "status": "success",
            "path": folder,
            "name": os.path.basename(folder) + "/",
            "count": cnt,
            "exists": ok
        })

    def handle_browse_file(self, stage):
        file_path = ask_file_dialog(title="Select Master SWRS Excel Template")
        if not file_path:
            self.send_json({"status": "cancelled"})
            return

        cnt, ok = count_files(file_path, ".xlsx")
        self.send_json({
            "status": "success",
            "path": file_path,
            "name": os.path.basename(file_path),
            "count": cnt,
            "exists": ok
        })

    def serve_download(self):
        with state_lock:
            out_file = task_state.get("result", {}).get("output_file") if task_state.get("result") else None
            if not out_file or not os.path.exists(out_file):
                out_file = DEFAULT_CONFIG["output_excel"]

        if not os.path.exists(out_file):
            self.send_error(404, "Output file not found. Please generate the report first.")
            return

        try:
            with open(out_file, "rb") as f:
                content = f.read()
            filename = os.path.basename(out_file)
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, f"Error reading download file: {e}")

    def handle_open_folder(self, data):
        folder = data.get("path")
        if not folder:
            with state_lock:
                folder = task_state.get("result", {}).get("output_file") if task_state.get("result") else None
        if not folder:
            folder = DEFAULT_CONFIG["output_excel"]
        target = os.path.dirname(folder) if os.path.isfile(folder) else folder
        if os.name == "nt" and os.path.exists(target):
            try:
                subprocess.Popen(f'explorer "{target}"')
                self.send_json({"status": "success", "opened": target})
                return
            except Exception as e:
                self.send_json({"status": "error", "message": str(e)}, 500)
                return
        self.send_json({"status": "error", "message": "Directory does not exist"}, 400)

    def handle_verify(self, data):
        excel_path = data.get("input_excel", "").strip() or DEFAULT_CONFIG["input_excel"]
        unit_dir = data.get("unit_dir", "").strip() or DEFAULT_CONFIG["unit_dir"]
        integ_dir = data.get("integ_dir", "").strip() or DEFAULT_CONFIG["integ_dir"]
        shi_dir = data.get("shi_dir", "").strip() or DEFAULT_CONFIG["shi_dir"]
        overall_dir = data.get("overall_dir", "").strip() or DEFAULT_CONFIG["overall_dir"]

        excel_exists = os.path.exists(excel_path)
        if excel_exists and os.path.isdir(excel_path):
            excel_cnt, _ = count_files(excel_path, ".xlsx")
            excel_exists = excel_cnt > 0
        else:
            excel_cnt = 1 if excel_exists else 0

        unit_cnt, unit_ok = count_files(unit_dir, ".tcf")
        integ_cnt, integ_ok = count_files(integ_dir, ".tcf")
        shi_cnt, shi_ok = count_files(shi_dir, ".docx")
        overall_cnt, overall_ok = count_files(overall_dir, ".docx")

        response = {
            "input_excel": {"exists": excel_exists, "count": excel_cnt, "name": os.path.basename(excel_path)},
            "unit_dir": {"exists": unit_ok and unit_cnt > 0, "count": unit_cnt, "name": os.path.basename(unit_dir)},
            "integ_dir": {"exists": integ_ok and integ_cnt > 0, "count": integ_cnt, "name": os.path.basename(integ_dir)},
            "shi_dir": {"exists": shi_ok and shi_cnt > 0, "count": shi_cnt, "name": os.path.basename(shi_dir)},
            "overall_dir": {"exists": overall_ok and overall_cnt > 0, "count": overall_cnt, "name": os.path.basename(overall_dir)},
            "all_valid": bool(excel_exists)
        }
        self.send_json(response)

    def handle_generate(self, data):
        """
        Executes report generation directly from local paths on disk.
        NO network upload, zero data transfer delay.
        """
        global task_state
        with state_lock:
            if task_state["is_running"]:
                self.send_json({"status": "busy", "message": "A generation task is already in progress."}, 409)
                return

        input_excel = data.get("input_excel", "").strip() or DEFAULT_CONFIG["input_excel"]
        if not input_excel or not os.path.exists(input_excel):
            self.send_json({
                "status": "error",
                "message": "SWRS Requirements Excel file (.xlsx) or folder is MANDATORY."
            }, 400)
            return

        include_shi = bool(data.get("include_shi", True))
        include_sys = bool(data.get("include_sys", True))
        include_comp = bool(data.get("include_comp", True))
        include_integ = bool(data.get("include_integ", True))

        shi_dir = data.get("shi_dir", "").strip() or DEFAULT_CONFIG["shi_dir"]
        overall_dir = data.get("overall_dir", "").strip() or DEFAULT_CONFIG["overall_dir"]
        unit_dir = data.get("unit_dir", "").strip() or DEFAULT_CONFIG["unit_dir"]
        integ_dir = data.get("integ_dir", "").strip() or DEFAULT_CONFIG["integ_dir"]

        if include_shi and not (shi_dir and os.path.exists(shi_dir)):
            include_shi = False
        if include_sys and not (overall_dir and os.path.exists(overall_dir)):
            include_sys = False
        if include_comp and not (unit_dir and os.path.exists(unit_dir)):
            include_comp = False
        if include_integ and not (integ_dir and os.path.exists(integ_dir)):
            include_integ = False

        out_path = DEFAULT_CONFIG["output_excel"]

        pipeline_params = {
            "input_excel": input_excel,
            "shi_dir": shi_dir if include_shi else None,
            "overall_dir": overall_dir if include_sys else None,
            "unit_dir": unit_dir if include_comp else None,
            "integ_dir": integ_dir if include_integ else None,
            "output_excel": out_path,
            "include_shi": include_shi,
            "include_sys": include_sys,
            "include_comp": include_comp,
            "include_integ": include_integ,
            "tag_tcfs": bool(data.get("tag_tcfs", False))
        }

        with state_lock:
            task_state["is_running"] = True
            task_state["progress"] = 0
            task_state["message"] = "Initializing generation pipeline..."
            task_state["logs"] = []
            task_state["result"] = None
            task_state["error"] = None

        thread = threading.Thread(target=self.run_background_pipeline, args=(pipeline_params,))
        thread.daemon = True
        thread.start()

        self.send_json({
            "status": "started",
            "message": "Generation started directly from disk paths.",
            "active_stages": {
                "shi": include_shi,
                "system_test": include_sys,
                "component_test": include_comp,
                "integration_test": include_integ
            }
        })

    def run_background_pipeline(self, data):
        global task_state

        def progress_cb(msg, pct):
            with state_lock:
                task_state["progress"] = pct
                task_state["message"] = msg
                timestamp = time.strftime("%H:%M:%S")
                task_state["logs"].append(f"[{timestamp}] [{pct}%] {msg}")

        try:
            res = generator_core.run_traceability_pipeline(
                unit_dir=data.get("unit_dir"),
                integ_dir=data.get("integ_dir"),
                shi_dir=data.get("shi_dir"),
                overall_dir=data.get("overall_dir"),
                input_excel=data.get("input_excel"),
                output_excel=data.get("output_excel"),
                tag_tcfs=data.get("tag_tcfs", False),
                include_shi=data.get("include_shi"),
                include_sys=data.get("include_sys"),
                include_comp=data.get("include_comp"),
                include_integ=data.get("include_integ"),
                progress_callback=progress_cb
            )
            with state_lock:
                task_state["result"] = res
                task_state["progress"] = 100
                task_state["message"] = "Completed successfully!"
                task_state["is_running"] = False
        except Exception as e:
            with state_lock:
                task_state["error"] = str(e)
                task_state["progress"] = 100
                task_state["message"] = f"Error: {e}"
                task_state["logs"].append(f"[{time.strftime('%H:%M:%S')}] [ERROR] {e}")
                task_state["is_running"] = False

    def serve_ui(self):
        html = self.get_ui_html()
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(body)

    def get_ui_html(self):
        return """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SSBPAC Traceability Suite</title>
    <style>
        :root {
            --primary: #0f172a;
            --primary-light: #1e293b;
            --accent: #059669;
            --accent-hover: #047857;
            --accent-glow: rgba(5, 150, 105, 0.25);
            --bg: #f8fafc;
            --card-bg: #ffffff;
            --border: #e2e8f0;
            --border-hover: #cbd5e1;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --radius-lg: 14px;
            --radius-md: 10px;
            --radius-sm: 6px;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
        body { background: var(--bg); color: var(--text-main); padding: 28px 20px; min-height: 100vh; -webkit-font-smoothing: antialiased; }
        .container { max-width: 1080px; margin: 0 auto; }
        
        /* Executive Header */
        .header {
            background: linear-gradient(135deg, #090d16 0%, #111827 50%, #1e293b 100%);
            border: 1px solid rgba(255, 255, 255, 0.08);
            color: white;
            padding: 22px 28px;
            border-radius: var(--radius-lg);
            margin-bottom: 20px;
            box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15), 0 4px 6px -2px rgba(15, 23, 42, 0.05);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .header-brand {
            display: flex;
            align-items: center;
            gap: 14px;
        }
        .brand-icon {
            width: 42px;
            height: 42px;
            border-radius: 10px;
            background: linear-gradient(135deg, #059669 0%, #0284c7 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px rgba(5, 150, 105, 0.35);
        }
        .header h1 { font-size: 21px; font-weight: 800; letter-spacing: -0.4px; line-height: 1.2; }
        .header-sub { font-size: 12.5px; color: #94a3b8; font-weight: 500; margin-top: 2px; }

        /* Sleek Author Badge */
        .author-badge {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(255, 255, 255, 0.12);
            border: 1px solid rgba(255, 255, 255, 0.22);
            padding: 6px 16px;
            border-radius: 9999px;
            font-size: 13.5px;
            font-weight: 600;
            letter-spacing: 0.3px;
            color: #ffffff;
            backdrop-filter: blur(10px);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
            transition: all 0.2s ease;
        }
        .author-badge:hover {
            background: rgba(255, 255, 255, 0.18);
            border-color: rgba(255, 255, 255, 0.35);
            transform: translateY(-1px);
        }

        /* Scope & Filter Bar */
        .matrix-banner {
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 14px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        }
        .matrix-title { font-size: 11px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 2px; }
        .matrix-badge {
            background: #ecfdf5;
            color: #065f46;
            border: 1px solid #a7f3d0;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 13.5px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .stages-pills { display: flex; gap: 8px; flex-wrap: wrap; }
        .stage-pill {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            background: #f1f5f9;
            color: #64748b;
            border: 1px solid #e2e8f0;
            transition: all 0.2s;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .stage-pill.active {
            background: #dcfce7;
            color: #15803d;
            border-color: #86efac;
            font-weight: 700;
        }
        .stage-pill.mandatory {
            background: #fee2e2;
            color: #b91c1c;
            border-color: #fca5a5;
            font-weight: 700;
            cursor: default;
        }

        /* Card Container */
        .card {
            background: var(--card-bg);
            border-radius: var(--radius-lg);
            padding: 24px 26px;
            margin-bottom: 20px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04), 0 8px 16px -4px rgba(0,0,0,0.02);
            border: 1px solid var(--border);
        }
        .card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 18px;
            padding-bottom: 14px;
            border-bottom: 1px solid #f1f5f9;
        }
        .card-title { font-size: 16.5px; font-weight: 700; color: var(--primary); display: flex; align-items: center; gap: 10px; }
        .card-badge { font-size: 11.5px; font-weight: 600; color: #64748b; background: #f1f5f9; padding: 4px 10px; border-radius: 12px; border: 1px solid #e2e8f0; }

        /* Stage Card */
        .stage-card {
            border: 1.5px solid #e2e8f0;
            border-radius: var(--radius-md);
            padding: 16px 20px;
            margin-bottom: 14px;
            background: #ffffff;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .stage-card:hover { border-color: #cbd5e1; box-shadow: 0 2px 6px rgba(0,0,0,0.02); }
        .stage-card.disabled { opacity: 0.45; background: #f8fafc; border-style: dashed; }
        .stage-card-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
        }
        .stage-title {
            font-size: 14.5px;
            font-weight: 700;
            color: #1e293b;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        /* Modern Switch */
        .switch {
            position: relative;
            display: inline-block;
            width: 38px;
            height: 22px;
        }
        .switch input { opacity: 0; width: 0; height: 0; }
        .slider {
            position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0;
            background-color: #cbd5e1; transition: .2s; border-radius: 22px;
        }
        .slider:before {
            position: absolute; content: ""; height: 16px; width: 16px; left: 3px; bottom: 3px;
            background-color: white; transition: .2s; border-radius: 50%; box-shadow: 0 1px 3px rgba(0,0,0,0.25);
        }
        input:checked + .slider { background-color: var(--accent); }
        input:checked + .slider:before { transform: translateX(16px); }

        .stage-badge {
            font-size: 11px;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: 8px;
            letter-spacing: 0.3px;
        }
        .badge-req { background: #fee2e2; color: #dc2626; border: 1px solid #fca5a5; }
        .badge-opt { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }

        /* Source Box */
        .source-box {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: var(--radius-sm);
            padding: 12px 16px;
            transition: all 0.2s;
        }
        .source-box:hover { border-color: #cbd5e1; background: #ffffff; }
        
        .source-info { display: flex; align-items: center; gap: 12px; }
        .source-icon { font-size: 20px; line-height: 1; }
        .source-name { font-size: 13.5px; font-weight: 700; color: #0f172a; margin-bottom: 2px; }
        .source-desc { font-size: 12px; color: #64748b; }
        
        .source-actions { display: flex; gap: 8px; align-items: center; }
        .btn-folder {
            padding: 7px 13px;
            font-size: 12.5px;
            font-weight: 600;
            border-radius: var(--radius-sm);
            border: 1px solid #cbd5e1;
            background: #ffffff;
            color: #334155;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.15s;
            white-space: nowrap;
            box-shadow: 0 1px 2px rgba(0,0,0,0.03);
        }
        .btn-folder:hover { background: #f8fafc; border-color: #94a3b8; }
        .btn-folder-accent { background: #f0fdf4; color: #166534; border-color: #bbf7d0; font-weight: 700; }
        .btn-folder-accent:hover { background: #dcfce7; border-color: #86efac; }

        /* Actions Bar */
        .actions-bar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 14px;
            padding-top: 18px;
            margin-top: 6px;
            border-top: 1px solid #e2e8f0;
        }
        .status-pulse {
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #10b981;
            box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.2);
            margin-right: 6px;
        }
        .btn {
            padding: 12px 26px;
            font-size: 14px;
            font-weight: 700;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
            display: inline-flex;
            align-items: center;
            gap: 8px;
            text-decoration: none;
        }
        .btn-primary {
            background: linear-gradient(135deg, #059669 0%, #047857 100%);
            color: white;
            box-shadow: 0 4px 14px var(--accent-glow);
        }
        .btn-primary:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 20px rgba(5, 150, 105, 0.35);
        }
        .btn-secondary { background: #ffffff; color: #334155; border: 1.5px solid #cbd5e1; }
        .btn-secondary:hover { background: #f8fafc; border-color: #94a3b8; }
        .btn:disabled { opacity: 0.55; cursor: not-allowed; transform: none; box-shadow: none; }

        /* Terminal Logs & Progress */
        .progress-box { margin-top: 20px; display: none; }
        .progress-bar-bg { background: #e2e8f0; height: 10px; border-radius: 5px; overflow: hidden; margin-bottom: 8px; }
        .progress-bar-fill {
            background: linear-gradient(90deg, #059669, #10b981, #0284c7);
            background-size: 200% 100%;
            animation: shimmer 2s infinite linear;
            height: 100%;
            width: 0%;
            transition: width 0.35s ease;
        }
        @keyframes shimmer { 0% { background-position: 100% 0; } 100% { background-position: -100% 0; } }
        .progress-text { font-size: 12.5px; color: var(--text-muted); display: flex; justify-content: space-between; font-weight: 600; }
        .terminal-header {
            background: #090d16;
            padding: 8px 14px;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            display: flex;
            align-items: center;
            gap: 6px;
            margin-top: 14px;
            border-bottom: 1px solid rgba(255,255,255,0.06);
        }
        .terminal-dot { width: 10px; height: 10px; border-radius: 50%; }
        .log-box {
            background: #090d16;
            color: #94a3b8;
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            font-size: 12px;
            padding: 14px;
            border-bottom-left-radius: 8px;
            border-bottom-right-radius: 8px;
            height: 160px;
            overflow-y: auto;
            line-height: 1.6;
        }

        /* Results Panel */
        .results-panel { display: none; margin-top: 24px; animation: fadeIn 0.4s ease; border-left: 4px solid var(--accent); }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 14px;
            margin: 18px 0;
        }
        .metric-card {
            background: #f8fafc;
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 16px;
            text-align: center;
            transition: transform 0.2s;
        }
        .metric-card:hover { transform: translateY(-1px); }
        .metric-value { font-size: 28px; font-weight: 800; color: var(--primary); margin-top: 4px; }
        .metric-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; font-weight: 700; letter-spacing: 0.6px; }

        .sheet-pill {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 6px;
            background: #e0f2fe;
            color: #0369a1;
            font-size: 12px;
            font-weight: 700;
            margin-right: 6px;
            margin-bottom: 4px;
            border: 1px solid #bae6fd;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Executive Navbar Header -->
        <div class="header">
            <div class="header-brand">
                <div class="brand-icon">
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                        <polygon points="12 2 2 7 12 12 22 7 12 2"></polygon>
                        <polyline points="2 17 12 22 22 17"></polyline>
                        <polyline points="2 12 12 17 22 12"></polyline>
                    </svg>
                </div>
                <div>
                    <h1>SSBPAC Traceability Suite</h1>
                </div>
            </div>
            <div class="author-badge">
                <span>👤</span> Vivek
            </div>
        </div>

        <!-- Dynamic Scope Indicator -->
        <div class="matrix-banner">
            <div>
                <div class="matrix-title">Target Compliance Matrix Scope</div>
                <div id="matrix-mode-label" class="matrix-badge">
                    <span style="font-size: 15px;">📊</span> SWRS vs Full Suite (16 Columns)
                </div>
            </div>
            <div class="stages-pills">
                <span id="pill-swrs" class="stage-pill mandatory">● SWRS (Mandatory)</span>
                <span id="pill-shi" class="stage-pill active" onclick="toggleStage('shi')">● SHI Testing</span>
                <span id="pill-sys" class="stage-pill active" onclick="toggleStage('sys')">● System Testing</span>
                <span id="pill-comp" class="stage-pill active" onclick="toggleStage('comp')">● Component Testing</span>
                <span id="pill-integ" class="stage-pill active" onclick="toggleStage('integ')">● Integration Testing</span>
            </div>
        </div>

        <!-- Main Form Card -->
        <div class="card">
            <div class="card-header">
                <div class="card-title">
                    <svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
                    Input Documents & Test Artifacts
                </div>
                <span class="card-badge">5 Configured Stages</span>
            </div>

            <!-- 1. Master SWRS Requirements (Mandatory) -->
            <div id="stage-card-swrs" class="stage-card" style="border-left: 3.5px solid #dc2626;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <span>1. Master SWRS Requirements</span>
                        <span class="stage-badge badge-req">Mandatory</span>
                    </div>
                </div>
                <div class="source-box">
                    <div class="source-info">
                        <div class="source-icon">📊</div>
                        <div>
                            <div id="name-swrs" class="source-name">SWRS_ID.xlsx</div>
                            <div id="desc-swrs" class="source-desc">Master requirements template (Section, SL No., Requirement & ID's)</div>
                        </div>
                    </div>
                    <div class="source-actions">
                        <input type="file" id="file-input-swrs" accept=".xlsx" style="display:none;" onchange="handleFileInput(event, 'swrs')">
                        <button type="button" class="btn-folder btn-folder-accent" onclick="browseFile('swrs')">
                            📄 Choose File
                        </button>
                    </div>
                </div>
            </div>

            <!-- 2. SHI Word Reports Folder (Optional) -->
            <div id="stage-card-shi" class="stage-card" style="border-left: 3.5px solid #0284c7;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <label class="switch">
                            <input type="checkbox" id="check-shi" checked onchange="updateUIState()">
                            <span class="slider"></span>
                        </label>
                        <span>2. SHI Verification Folder</span>
                        <span class="stage-badge badge-opt">Optional</span>
                    </div>
                </div>
                <div class="source-box">
                    <div class="source-info">
                        <div class="source-icon">📁</div>
                        <div>
                            <div id="name-shi" class="source-name">SHI Folder</div>
                            <div id="desc-shi" class="source-desc">9 DOCX reports in Single line & Double line subfolders</div>
                        </div>
                    </div>
                    <div class="source-actions">
                        <input type="file" id="folder-input-shi" webkitdirectory directory multiple style="display:none;" onchange="handleFolderUpload(event, 'shi')">
                        <button type="button" class="btn-folder btn-folder-accent" onclick="browseFolder('shi')">
                            📁 Choose Folder
                        </button>
                    </div>
                </div>
            </div>

            <!-- 3. Overall System Test Reports Folder (Optional) -->
            <div id="stage-card-sys" class="stage-card" style="border-left: 3.5px solid #6366f1;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <label class="switch">
                            <input type="checkbox" id="check-sys" checked onchange="updateUIState()">
                            <span class="slider"></span>
                        </label>
                        <span>3. Overall System Software Test Folder</span>
                        <span class="stage-badge badge-opt">Optional</span>
                    </div>
                </div>
                <div class="source-box">
                    <div class="source-info">
                        <div class="source-icon">📁</div>
                        <div>
                            <div id="name-sys" class="source-name">Overall Software Test report Folder</div>
                            <div id="desc-sys" class="source-desc">System verification DOCX reports</div>
                        </div>
                    </div>
                    <div class="source-actions">
                        <input type="file" id="folder-input-sys" webkitdirectory directory multiple style="display:none;" onchange="handleFolderUpload(event, 'sys')">
                        <button type="button" class="btn-folder btn-folder-accent" onclick="browseFolder('sys')">
                            📁 Choose Folder
                        </button>
                    </div>
                </div>
            </div>

            <!-- 4. Software Component Testing Folder (Optional) -->
            <div id="stage-card-comp" class="stage-card" style="border-left: 3.5px solid #059669;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <label class="switch">
                            <input type="checkbox" id="check-comp" checked onchange="updateUIState()">
                            <span class="slider"></span>
                        </label>
                        <span>4. Software Component (Unit Testing) Folder</span>
                        <span class="stage-badge badge-opt">Optional</span>
                    </div>
                </div>
                <div class="source-box">
                    <div class="source-info">
                        <div class="source-icon">📁</div>
                        <div>
                            <div id="name-comp" class="source-name">Unit testing Folder</div>
                            <div id="desc-comp" class="source-desc">156 TCF test cases (BVA, Statement Coverage, etc.)</div>
                        </div>
                    </div>
                    <div class="source-actions">
                        <input type="file" id="folder-input-comp" webkitdirectory directory multiple style="display:none;" onchange="handleFolderUpload(event, 'comp')">
                        <button type="button" class="btn-folder btn-folder-accent" onclick="browseFolder('comp')">
                            📁 Choose Folder
                        </button>
                    </div>
                </div>
            </div>

            <!-- 5. Software Integration Testing Folder (Optional) -->
            <div id="stage-card-integ" class="stage-card" style="border-left: 3.5px solid #8b5cf6;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <label class="switch">
                            <input type="checkbox" id="check-integ" checked onchange="updateUIState()">
                            <span class="slider"></span>
                        </label>
                        <span>5. Software Integration Testing Folder</span>
                        <span class="stage-badge badge-opt">Optional</span>
                    </div>
                </div>
                <div class="source-box">
                    <div class="source-info">
                        <div class="source-icon">📁</div>
                        <div>
                            <div id="name-integ" class="source-name">Integration testing Folder</div>
                            <div id="desc-integ" class="source-desc">25 TCF integration test cases</div>
                        </div>
                    </div>
                    <div class="source-actions">
                        <input type="file" id="folder-input-integ" webkitdirectory directory multiple style="display:none;" onchange="handleFolderUpload(event, 'integ')">
                        <button type="button" class="btn-folder btn-folder-accent" onclick="browseFolder('integ')">
                            📁 Choose Folder
                        </button>
                    </div>
                </div>
            </div>

            <!-- Actions Bar -->
            <div class="actions-bar">
                <div>
                    <div id="status-summary-text" style="font-size: 14px; font-weight: 700; color: #1e293b; display: flex; align-items: center;">
                        <span class="status-pulse"></span> Configured Output: SWRS vs Full Suite (16 Columns)
                    </div>
                    <div style="font-size: 12px; color: #64748b; margin-top: 3px;">
                        All artifacts loaded. Click below to generate the compliance report.
                    </div>
                </div>
                <button id="btn-generate" class="btn btn-primary" onclick="startGeneration()">
                    <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
                    Generate Traceability Report
                </button>
            </div>

            <!-- Progress Box -->
            <div id="progress-box" class="progress-box">
                <div class="progress-bar-bg">
                    <div id="progress-bar-fill" class="progress-bar-fill"></div>
                </div>
                <div class="progress-text">
                    <span id="progress-msg">Initializing generation pipeline...</span>
                    <span id="progress-pct">0%</span>
                </div>
                <div class="terminal-header">
                    <div class="terminal-dot" style="background: #ef4444;"></div>
                    <div class="terminal-dot" style="background: #eab308;"></div>
                    <div class="terminal-dot" style="background: #22c55e;"></div>
                    <span style="color: #64748b; font-size: 11px; margin-left: 8px; font-family: monospace;">execution_console</span>
                </div>
                <div id="log-box" class="log-box"></div>
            </div>
        </div>

        <!-- Results Card -->
        <div id="results-panel" class="card results-panel">
            <div class="card-header">
                <div>
                    <div class="card-title">
                        <span>Report Generated Successfully!</span>
                        <span id="result-badge" class="stage-pill active" style="margin-left: 8px;">Completed</span>
                    </div>
                    <div id="result-mode-text" style="font-size: 13px; color: var(--text-muted); margin-top: 4px;">
                        Mode: SWRS vs SHI (6 Columns)
                    </div>
                </div>
                <div style="display: flex; gap: 10px;">
                    <button id="btn-download" class="btn btn-primary" onclick="downloadReport()">
                        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        Download Report (.xlsx)
                    </button>
                    <button class="btn btn-secondary" onclick="openFolder()">
                        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
                        Open Folder
                    </button>
                </div>
            </div>

            <div id="sheet-pills-container" style="margin-bottom: 14px;"></div>

            <div class="metrics-grid">
                <div class="metric-card">
                    <div class="metric-label">Total SWRS Requirements</div>
                    <div id="m-total" class="metric-value">283</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Covered Requirements</div>
                    <div id="m-cov" class="metric-value" style="color: #059669;">92</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Overall Coverage</div>
                    <div id="m-pct" class="metric-value" style="color: #059669;">32.5%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Untraced Requirements</div>
                    <div id="m-untraced" class="metric-value" style="color: #dc2626;">191</div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Transparent interceptor for ngrok tunnels to prevent warning intercepts on API requests
        const _origFetch = window.fetch;
        window.fetch = function(url, options = {}) {
            options = options || {};
            options.headers = options.headers || {};
            if (options.headers instanceof Headers) {
                options.headers.set('ngrok-skip-browser-warning', 'true');
            } else {
                options.headers['ngrok-skip-browser-warning'] = 'true';
            }
            return _origFetch(url, options);
        };

        let pollInterval = null;
        let isLocalClient = true;
        
        // Disk paths stored in memory only — never shown on screen
        const pathsState = {
            swrs: '',
            shi: '',
            sys: '',
            comp: '',
            integ: ''
        };

        window.addEventListener('DOMContentLoaded', () => {
            fetch('/api/defaults')
                .then(r => r.json())
                .then(data => {
                    pathsState.swrs = data.input_excel || '';
                    pathsState.shi = data.shi_dir || '';
                    pathsState.sys = data.overall_dir || '';
                    pathsState.comp = data.unit_dir || '';
                    pathsState.integ = data.integ_dir || '';
                    isLocalClient = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

                    // Verify default folder counts
                    fetch('/api/verify', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(data)
                    })
                    .then(r => r.json())
                    .then(res => {
                        if (res.input_excel && res.input_excel.exists) {
                            document.getElementById('desc-swrs').innerText = `Master template verified (${res.input_excel.count} .xlsx)`;
                        }
                        if (res.shi_dir && res.shi_dir.exists) {
                            document.getElementById('desc-shi').innerText = `${res.shi_dir.count} DOCX reports in Single line & Double line subfolders`;
                        }
                        if (res.overall_dir && res.overall_dir.exists) {
                            document.getElementById('desc-sys').innerText = `${res.overall_dir.count} system verification DOCX reports`;
                        }
                        if (res.unit_dir && res.unit_dir.exists) {
                            document.getElementById('desc-comp').innerText = `${res.unit_dir.count} TCF test cases (BVA, Statement Coverage, etc.)`;
                        }
                        if (res.integ_dir && res.integ_dir.exists) {
                            document.getElementById('desc-integ').innerText = `${res.integ_dir.count} TCF integration test cases`;
                        }
                        updateUIState();
                    });
                });
        });

        function browseFolder(category) {
            if (!isLocalClient) {
                const el = document.getElementById('folder-input-' + category);
                if (el) el.click();
                return;
            }
            fetch('/api/browse-folder?stage=' + category)
                .then(r => r.json())
                .then(res => {
                    if (res.status === 'success' && res.path) {
                        pathsState[category] = res.path;
                        document.getElementById('name-' + category).innerText = res.name;
                        document.getElementById('desc-' + category).innerText = `${res.count} matching file(s) detected in folder`;
                        if (category !== 'swrs') {
                            document.getElementById('check-' + category).checked = true;
                        }
                        updateUIState();
                    }
                })
                .catch(() => {
                    const el = document.getElementById('folder-input-' + category);
                    if (el) el.click();
                });
        }

        function formatBytes(bytes) {
            if (!bytes || bytes <= 0) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
        }

        function browseFile(category) {
            if (isLocalClient) {
                fetch('/api/browse-file?stage=' + category)
                    .then(r => r.json())
                    .then(res => {
                        if (res.status === 'success' && res.path) {
                            pathsState[category] = res.path;
                            document.getElementById('name-' + category).innerText = res.name;
                            document.getElementById('desc-' + category).innerText = 'Master template verified (' + res.count + ' .xlsx)';
                            updateUIState();
                        }
                    })
                    .catch(() => {
                        const fileInput = document.getElementById('file-input-' + category);
                        if (fileInput) fileInput.click();
                    });
                return;
            }
            const fileInput = document.getElementById('file-input-' + category);
            if (fileInput) {
                fileInput.click();
            }
        }

        async function handleFileInput(e, category) {
            const file = e.target.files[0];
            if (!file) return;
            const descEl = document.getElementById('desc-' + category);
            const origDesc = descEl.innerHTML;
            const fileSizeStr = formatBytes(file.size);

            descEl.innerHTML = `
                <div style="margin-top: 6px;">
                    <div style="display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11.5px; font-weight: 600; color: #0f172a; margin-bottom: 4px;">
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px;" title="${file.name}">
                            Uploading: <span style="font-weight: 500; color: #475569;">${file.name}</span>
                        </span>
                        <span style="background: #059669; color: #ffffff; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700;">${fileSizeStr}</span>
                    </div>
                    <div style="width: 100%; height: 6px; background: #e2e8f0; border-radius: 999px; overflow: hidden;">
                        <div style="width: 100%; height: 100%; background: linear-gradient(90deg, #059669, #0284c7); border-radius: 999px;"></div>
                    </div>
                </div>
            `;

            try {
                const r = await fetch('/api/upload?stage=' + category + '&filename=' + encodeURIComponent(file.name), {
                    method: 'POST',
                    headers: { 
                        'Content-Type': 'application/octet-stream',
                        'ngrok-skip-browser-warning': 'true'
                    },
                    body: file
                });

                const resText = await r.text();
                let res;
                try {
                    res = JSON.parse(resText);
                } catch (parseErr) {
                    throw new Error(`Upload failed (Status ${r.status}): ${resText.slice(0, 120)}`);
                }

                if (res.status === 'success' && res.path) {
                    pathsState[category] = res.path;
                    document.getElementById('name-' + category).innerText = res.name;
                    descEl.innerHTML = `<span style="color: #065f46; font-weight: 600;">✓ Master template verified (${fileSizeStr})</span>`;
                    if (category !== 'swrs') {
                        document.getElementById('check-' + category).checked = true;
                    }
                    updateUIState();
                } else {
                    descEl.innerHTML = origDesc;
                    alert('Upload failed: ' + (res.message || 'Unknown error'));
                }
            } catch (err) {
                descEl.innerHTML = origDesc;
                alert('Upload error: ' + err.message);
            } finally {
                e.target.value = '';
            }
        }

        async function handleFolderUpload(e, category) {
            const rawFiles = Array.from(e.target.files);
            if (!rawFiles || rawFiles.length === 0) return;

            // Filter out system and temp word/excel lock files (like ~$WordDoc.docx)
            const files = rawFiles.filter(f => !f.name.startsWith('~$') && !f.name.startsWith('.'));
            if (files.length === 0) {
                alert('No valid test/document files found in the selected folder.');
                e.target.value = '';
                return;
            }

            const descEl = document.getElementById('desc-' + category);
            const origDesc = descEl.innerHTML;
            const folderName = files[0].webkitRelativePath ? files[0].webkitRelativePath.split('/')[0] : (category + '_folder');

            const totalBytes = files.reduce((acc, f) => acc + (f.size || 0), 0);
            const totalBytesStr = formatBytes(totalBytes);
            let uploadedBytes = 0;
            let uploadedCount = 0;
            let lastRes = null;

            try {
                // Upload files sequentially with detailed file size, dynamic progress percentage, and animated bar
                for (let i = 0; i < files.length; i++) {
                    const f = files[i];
                    const fSize = f.size || 0;
                    const fSizeStr = formatBytes(fSize);
                    
                    // Progress percentage calculation
                    const pct = totalBytes > 0 
                        ? Math.min(99, Math.max(1, Math.round((uploadedBytes / totalBytes) * 100)))
                        : Math.round(((i) / files.length) * 100);

                    descEl.innerHTML = `
                        <div style="margin-top: 6px;">
                            <div style="display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11.5px; font-weight: 600; color: #0f172a; margin-bottom: 4px;">
                                <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 320px;" title="${f.name}">
                                    Uploading [${i + 1}/${files.length}]: <span style="font-weight: 500; color: #475569;">${f.name}</span>
                                </span>
                                <div style="display: inline-flex; align-items: center; gap: 6px; flex-shrink: 0;">
                                    <span style="background: #f1f5f9; color: #334155; padding: 2px 7px; border-radius: 999px; font-size: 11px; border: 1px solid #cbd5e1;">${fSizeStr}</span>
                                    <span style="background: #059669; color: #ffffff; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700;">${pct}%</span>
                                </div>
                            </div>
                            <div style="width: 100%; height: 6px; background: #e2e8f0; border-radius: 999px; overflow: hidden;">
                                <div style="width: ${pct}%; height: 100%; background: linear-gradient(90deg, #059669, #0284c7); border-radius: 999px; transition: width 0.25s ease;"></div>
                            </div>
                            <div style="display: flex; justify-content: space-between; font-size: 10.5px; color: #64748b; margin-top: 3px;">
                                <span>${formatBytes(uploadedBytes)} of ${totalBytesStr} uploaded</span>
                                <span>${files.length - i} file(s) remaining</span>
                            </div>
                        </div>
                    `;

                    const relPath = f.webkitRelativePath || (folderName + '/' + f.name);
                    const url = '/api/upload?stage=' + category + 
                                '&folder=' + encodeURIComponent(folderName) + 
                                '&relpath=' + encodeURIComponent(relPath) + 
                                '&filename=' + encodeURIComponent(f.name);

                    const r = await fetch(url, {
                        method: 'POST',
                        headers: { 
                            'Content-Type': 'application/octet-stream',
                            'ngrok-skip-browser-warning': 'true'
                        },
                        body: f
                    });

                    const resText = await r.text();
                    let res;
                    try {
                        res = JSON.parse(resText);
                    } catch (parseErr) {
                        throw new Error(`Upload failed for "${f.name}" (Status ${r.status}): ${resText.slice(0, 120)}`);
                    }

                    if (res.status !== 'success') {
                        throw new Error(res.message || `Failed to upload ${f.name}`);
                    }

                    uploadedBytes += fSize;
                    uploadedCount++;
                    lastRes = res;
                }

                if (lastRes && lastRes.status === 'success') {
                    pathsState[category] = lastRes.path;
                    document.getElementById('name-' + category).innerText = folderName + '/';
                    descEl.innerHTML = `<span style="color: #065f46; font-weight: 600;">✓ Uploaded & verified ${uploadedCount} file(s) (${totalBytesStr})</span>`;
                    if (category !== 'swrs') {
                        document.getElementById('check-' + category).checked = true;
                    }
                    updateUIState();
                }
            } catch (err) {
                descEl.innerHTML = origDesc;
                alert('Folder upload error: ' + err.message);
            } finally {
                e.target.value = '';
            }
        }

        function updateUIState() {
            const shiActive = document.getElementById('check-shi').checked;
            const sysActive = document.getElementById('check-sys').checked;
            const compActive = document.getElementById('check-comp').checked;
            const integActive = document.getElementById('check-integ').checked;

            setCardActive('stage-card-shi', shiActive);
            setCardActive('stage-card-sys', sysActive);
            setCardActive('stage-card-comp', compActive);
            setCardActive('stage-card-integ', integActive);

            document.getElementById('pill-swrs').className = 'stage-pill mandatory';
            document.getElementById('pill-shi').className = 'stage-pill ' + (shiActive ? 'active' : '');
            document.getElementById('pill-sys').className = 'stage-pill ' + (sysActive ? 'active' : '');
            document.getElementById('pill-comp').className = 'stage-pill ' + (compActive ? 'active' : '');
            document.getElementById('pill-integ').className = 'stage-pill ' + (integActive ? 'active' : '');

            const stages = [];
            let cols = 4;
            if (shiActive) { stages.push('SHI'); cols += 2; }
            if (sysActive) { stages.push('System Test'); cols += 2; }
            if (compActive) { stages.push('Component Testing'); cols += 4; }
            if (integActive) { stages.push('Integration Testing'); cols += 4; }

            const labelEl = document.getElementById('matrix-mode-label');
            const summaryEl = document.getElementById('status-summary-text');
            const btnGen = document.getElementById('btn-generate');

            btnGen.disabled = false;
            const modeStr = stages.length > 0 ? 'SWRS vs ' + stages.join(' + ') : 'SWRS Base Template Only';
            labelEl.innerHTML = `📊 ${modeStr} (${cols} Columns)`;
            summaryEl.innerHTML = `Configured Output: <strong>${modeStr}</strong> (${cols} columns).`;
        }

        function setCardActive(cardId, active) {
            const el = document.getElementById(cardId);
            if (!el) return;
            if (active) el.classList.remove('disabled');
            else el.classList.add('disabled');
        }

        function toggleStage(stage) {
            const chk = document.getElementById('check-' + stage);
            if (chk) {
                chk.checked = !chk.checked;
                updateUIState();
            }
        }

        // --- Start Generation Directly on Disk (0 MB Upload) ---
        function startGeneration() {
            const data = {
                input_excel: pathsState.swrs,
                shi_dir: pathsState.shi,
                overall_dir: pathsState.sys,
                unit_dir: pathsState.comp,
                integ_dir: pathsState.integ,
                include_shi: document.getElementById('check-shi').checked,
                include_sys: document.getElementById('check-sys').checked,
                include_comp: document.getElementById('check-comp').checked,
                include_integ: document.getElementById('check-integ').checked
            };

            document.getElementById('btn-generate').disabled = true;
            document.getElementById('progress-box').style.display = 'block';
            document.getElementById('results-panel').style.display = 'none';
            document.getElementById('log-box').innerHTML = '';
            document.getElementById('progress-bar-fill').style.width = '10%';
            document.getElementById('progress-pct').innerText = '10%';
            document.getElementById('progress-msg').innerText = 'Starting pipeline directly from disk...';

            fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
            })
            .then(r => r.json())
            .then(res => {
                if (res.status === 'started') {
                    pollInterval = setInterval(pollProgress, 1000);
                } else {
                    alert('Error: ' + (res.message || 'Generation could not start'));
                    document.getElementById('btn-generate').disabled = false;
                }
            })
            .catch(err => {
                alert('Connection error: ' + err);
                document.getElementById('btn-generate').disabled = false;
            });
        }

        function pollProgress() {
            fetch('/api/progress')
                .then(r => r.json())
                .then(state => {
                    document.getElementById('progress-bar-fill').style.width = state.progress + '%';
                    document.getElementById('progress-pct').innerText = state.progress + '%';
                    document.getElementById('progress-msg').innerText = state.message;

                    const logBox = document.getElementById('log-box');
                    logBox.innerHTML = state.logs.join('\\n');
                    logBox.scrollTop = logBox.scrollHeight;

                    if (!state.is_running && state.progress >= 100) {
                        clearInterval(pollInterval);
                        document.getElementById('btn-generate').disabled = false;

                        if (state.result && state.result.metrics) {
                            showResults(state.result);
                        } else if (state.error) {
                            alert('Generation error: ' + state.error);
                        }
                    }
                });
        }

        function showResults(result) {
            document.getElementById('results-panel').style.display = 'block';
            const m = result.metrics || {};
            document.getElementById('m-total').innerText = m.total_requirements || 0;
            document.getElementById('m-cov').innerText = m.covered_requirements || 0;
            document.getElementById('m-pct').innerText = m.coverage_pct || '0%';
            document.getElementById('m-untraced').innerText = m.untraced_requirements || 0;

            const modeText = result.mode || 'Complete Traceability';
            document.getElementById('result-mode-text').innerText = `Generated Mode: ${modeText}`;

            let pills = '';
            if (result.sheets && Array.isArray(result.sheets) && result.sheets.length > 0) {
                pills = result.sheets.map((s, idx) => `<span class="sheet-pill">${idx + 1}. ${s}</span>`).join('');
            } else {
                let sheetIdx = 1;
                pills += `<span class="sheet-pill">${sheetIdx++}. Traceability Matrix</span>`;
                pills += `<span class="sheet-pill">${sheetIdx++}. Summary Statistics</span>`;
                const hasTcf = (m.component_swrs > 0 || m.integration_swrs > 0 || m.unmapped_tcs > 0 ||
                               (result.active_stages && (result.active_stages.component_test || result.active_stages.integration_test)));
                if (hasTcf) {
                    pills += `<span class="sheet-pill">${sheetIdx++}. Unmapped Test Cases</span>`;
                }
                pills += `<span class="sheet-pill">${sheetIdx++}. Unmapped SWRS Requirements</span>`;
            }
            document.getElementById('sheet-pills-container').innerHTML = pills;

            document.getElementById('results-panel').scrollIntoView({ behavior: 'smooth' });
        }

        function downloadReport() {
            window.location.href = '/api/download';
        }

        function openFolder() {
            fetch('/api/open-folder', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({})
            });
        }
    </script>
</body>
</html>"""


def run_server(host="0.0.0.0", port=8080):
    import socket
    server = ThreadingHTTPServer((host, port), TraceabilityHandler)
    
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        local_ip = s.getsockname()[0]
        s.close()
    except Exception:
        local_ip = "127.0.0.1"

    print("=" * 70)
    print(" SSBPAC TRACEABILITY WEB APPLICATION SERVER")
    print("=" * 70)
    print(f" Local Access (Your PC):        http://localhost:{port}")
    print(f" Local Network (Same Wi-Fi):    http://{local_ip}:{port}")
    print("=" * 70)
    print(" Direct disk processing (0 MB network upload).")
    print(" Press Ctrl+C to stop the server.")
    print("=" * 70)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        server.server_close()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", 8080))
    run_server(port=port)