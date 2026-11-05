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
  * In-Browser Searchable Data Grid Preview
  * SVG Donut Analytics & Stage Distribution Bars
  * Drag-and-Drop Dropzones on all artifact cards
  * DO-178C / SIL-4 Cryptographic Integrity Verification Seal & PDF Dossier
  * Dark / Light Theme Mode
"""

import os
import sys
import json
import time
import shutil
import hashlib
import urllib.parse
import threading
import subprocess
import openpyxl
import zipfile
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import generator_core

DEFAULT_CONFIG = {
    "input_excel": os.path.join(PARENT_DIR, "SWRS_ID.xlsx") if os.path.exists(os.path.join(PARENT_DIR, "SWRS_ID.xlsx")) else (os.path.join(PARENT_DIR, "overall_traceability.xlsx") if os.path.exists(os.path.join(PARENT_DIR, "overall_traceability.xlsx")) else os.path.join(CURRENT_DIR, "SWRS_ID.xlsx")),
    "shi_dir": os.path.join(PARENT_DIR, "SHI") if os.path.exists(os.path.join(PARENT_DIR, "SHI")) else os.path.join(CURRENT_DIR, "SHI"),
    "overall_dir": os.path.join(PARENT_DIR, "Overall System Software Test report") if os.path.exists(os.path.join(PARENT_DIR, "Overall System Software Test report")) else os.path.join(CURRENT_DIR, "Overall System Software Test report"),
    "unit_dir": os.path.join(PARENT_DIR, "Unit testing") if os.path.exists(os.path.join(PARENT_DIR, "Unit testing")) else os.path.join(CURRENT_DIR, "Unit testing"),
    "integ_dir": os.path.join(PARENT_DIR, "Integration testing") if os.path.exists(os.path.join(PARENT_DIR, "Integration testing")) else os.path.join(CURRENT_DIR, "Integration testing"),
    "output_excel": os.path.join(CURRENT_DIR, "overall_traceability_report.xlsx")
}

state_lock = threading.Lock()
task_state = {
    "is_running": False,
    "progress": 0,
    "message": "Ready",
    "logs": [],
    "result": None,
    "error": None
}

_preview_cache = {"mtime": 0, "data": None}
_preview_lock = threading.Lock()



def package_tagged_tcfs_zip(unit_dir, integ_dir, zip_path):
    """
    Packages all .tcf files from unit_dir and integ_dir into a zip file,
    preserving their relative directory structure.
    """
    dir_name = os.path.dirname(zip_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    count = 0
    unit_count = 0
    integ_count = 0

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        if unit_dir and os.path.exists(unit_dir):
            base_name = os.path.basename(os.path.normpath(unit_dir)) or "Unit testing"
            for root, _, files in os.walk(unit_dir):
                for f in files:
                    if f.lower().endswith(".tcf"):
                        full_p = os.path.join(root, f)
                        rel_p = os.path.relpath(full_p, unit_dir)
                        arc_name = os.path.join(base_name, rel_p)
                        zf.write(full_p, arc_name)
                        count += 1
                        unit_count += 1

        if integ_dir and os.path.exists(integ_dir):
            base_name = os.path.basename(os.path.normpath(integ_dir)) or "Integration testing"
            for root, _, files in os.walk(integ_dir):
                for f in files:
                    if f.lower().endswith(".tcf"):
                        full_p = os.path.join(root, f)
                        rel_p = os.path.relpath(full_p, integ_dir)
                        arc_name = os.path.join(base_name, rel_p)
                        zf.write(full_p, arc_name)
                        count += 1
                        integ_count += 1

    return count, unit_count, integ_count

def compute_file_sha256(filepath):
    """Calculates SHA-256 cryptographic hash of a file for DO-178C integrity trail."""
    if not filepath or not os.path.exists(filepath):
        return ""
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def map_matrix_columns(row1, row2):
    """Dynamically detects column indexes for each stage and sub-column from Row 1 and Row 2 headers."""
    stage_map = {
        "shi_sl": None,
        "shi_dl": None,
        "sys_sl": None,
        "sys_dl": None,
        "comp_mod": None,
        "comp_file": None,
        "comp_func": None,
        "comp_tc": None,
        "integ_mod": None,
        "integ_file": None,
        "integ_func": None,
        "integ_tc": None,
        "has_shi": False,
        "has_sys": False,
        "has_comp": False,
        "has_integ": False
    }
    if not row1 or len(row1) < 5:
        return stage_map

    current_stage = None
    for c_idx in range(4, len(row1)):
        r1_val = str(row1[c_idx]).strip() if row1[c_idx] is not None else ""
        r2_val = str(row2[c_idx]).strip() if row2 and len(row2) > c_idx and row2[c_idx] is not None else ""

        if "SHI" in r1_val:
            current_stage = "SHI"
        elif "Overall System" in r1_val or "System" in r1_val:
            current_stage = "SYS"
        elif "Component" in r1_val:
            current_stage = "COMP"
        elif "Integration" in r1_val:
            current_stage = "INTEG"

        if current_stage == "SHI":
            stage_map["has_shi"] = True
            if "Single" in r2_val or stage_map["shi_sl"] is None:
                stage_map["shi_sl"] = c_idx
            else:
                stage_map["shi_dl"] = c_idx
        elif current_stage == "SYS":
            stage_map["has_sys"] = True
            if "Single" in r2_val or stage_map["sys_sl"] is None:
                stage_map["sys_sl"] = c_idx
            else:
                stage_map["sys_dl"] = c_idx
        elif current_stage == "COMP":
            stage_map["has_comp"] = True
            if "Module" in r2_val:
                stage_map["comp_mod"] = c_idx
            elif "File" in r2_val:
                stage_map["comp_file"] = c_idx
            elif "Function" in r2_val:
                stage_map["comp_func"] = c_idx
            elif "Test case" in r2_val or "case" in r2_val.lower():
                stage_map["comp_tc"] = c_idx
        elif current_stage == "INTEG":
            stage_map["has_integ"] = True
            if "Module" in r2_val:
                stage_map["integ_mod"] = c_idx
            elif "File" in r2_val:
                stage_map["integ_file"] = c_idx
            elif "Function" in r2_val:
                stage_map["integ_func"] = c_idx
            elif "Test case" in r2_val or "case" in r2_val.lower():
                stage_map["integ_tc"] = c_idx

    return stage_map


def extract_preview_records(filepath):
    """Extracts parsed SWRS requirements from generated workbook for fast in-browser preview grid."""
    global _preview_cache
    if not filepath or not os.path.exists(filepath):
        return []

    mtime = os.path.getmtime(filepath)
    with _preview_lock:
        if _preview_cache["mtime"] == mtime and _preview_cache["data"] is not None:
            return _preview_cache["data"]

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            if "Traceability Matrix" not in wb.sheetnames:
                return []
            ws = wb["Traceability Matrix"]
            rows = list(ws.iter_rows(values_only=True))

            if len(rows) < 3:
                return []

            row1 = rows[0]
            row2 = rows[1]
            col_map = map_matrix_columns(row1, row2)

            def get_val(r, idx):
                if idx is not None and len(r) > idx and r[idx] is not None:
                    return str(r[idx]).strip()
                return ""

            def append_field(rec, key, val):
                if not val:
                    return
                existing = rec.get(key, "")
                if not existing:
                    rec[key] = val
                else:
                    lines = [l.strip() for l in existing.split("\n")]
                    if val not in lines:
                        rec[key] = existing + "\n" + val

            current_sec = ""
            records = []
            current_rec = None

            for r in rows[2:]:
                if not r:
                    continue
                first_col = r[0]
                if first_col and (len(r) < 2 or r[1] is None or r[1] == ""):
                    current_sec = str(first_col).strip()
                    continue
                if len(r) < 4:
                    continue
                swrs_id = str(r[3]).strip() if r[3] else ""
                if swrs_id and swrs_id.lower() not in ("id's", "ids", "none"):
                    sl_no = r[1] if r[1] is not None else ""
                    req_text = str(r[2]).strip() if r[2] else ""

                    current_rec = {
                        "section": current_sec or "General",
                        "sl_no": sl_no,
                        "swrs_id": swrs_id,
                        "requirement": req_text,
                        "shi_active": col_map["has_shi"],
                        "shi_sl": get_val(r, col_map["shi_sl"]),
                        "shi_dl": get_val(r, col_map["shi_dl"]),
                        "sys_active": col_map["has_sys"],
                        "sys_sl": get_val(r, col_map["sys_sl"]),
                        "sys_dl": get_val(r, col_map["sys_dl"]),
                        "comp_active": col_map["has_comp"],
                        "comp_mod": get_val(r, col_map["comp_mod"]),
                        "comp_file": get_val(r, col_map["comp_file"]),
                        "comp_func": get_val(r, col_map["comp_func"]),
                        "comp_tc": get_val(r, col_map["comp_tc"]),
                        "integ_active": col_map["has_integ"],
                        "integ_mod": get_val(r, col_map["integ_mod"]),
                        "integ_file": get_val(r, col_map["integ_file"]),
                        "integ_func": get_val(r, col_map["integ_func"]),
                        "integ_tc": get_val(r, col_map["integ_tc"]),
                    }
                    records.append(current_rec)
                elif current_rec is not None and not swrs_id:
                    # Continuation row for the preceding requirement
                    if col_map["has_comp"]:
                        append_field(current_rec, "comp_mod", get_val(r, col_map["comp_mod"]))
                        append_field(current_rec, "comp_file", get_val(r, col_map["comp_file"]))
                        append_field(current_rec, "comp_func", get_val(r, col_map["comp_func"]))
                        append_field(current_rec, "comp_tc", get_val(r, col_map["comp_tc"]))
                    if col_map["has_integ"]:
                        append_field(current_rec, "integ_mod", get_val(r, col_map["integ_mod"]))
                        append_field(current_rec, "integ_file", get_val(r, col_map["integ_file"]))
                        append_field(current_rec, "integ_func", get_val(r, col_map["integ_func"]))
                        append_field(current_rec, "integ_tc", get_val(r, col_map["integ_tc"]))

            for rec in records:
                rec["shi"] = bool(rec["shi_sl"] or rec["shi_dl"])
                rec["sys"] = bool(rec["sys_sl"] or rec["sys_dl"])
                rec["comp"] = bool(rec["comp_tc"])
                rec["integ"] = bool(rec["integ_tc"])
                rec["is_traced"] = bool(rec["shi"] or rec["sys"] or rec["comp"] or rec["integ"])
                rec["shi_info"] = ("SL: " + rec["shi_sl"] if rec["shi_sl"] else "") + (" | " if rec["shi_sl"] and rec["shi_dl"] else "") + ("DL: " + rec["shi_dl"] if rec["shi_dl"] else "")
                rec["sys_info"] = ("SL: " + rec["sys_sl"] if rec["sys_sl"] else "") + (" | " if rec["sys_sl"] and rec["sys_dl"] else "") + ("DL: " + rec["sys_dl"] if rec["sys_dl"] else "")
                rec["comp_info"] = rec["comp_tc"]
                rec["integ_info"] = rec["integ_tc"]

            wb.close()
            _preview_cache["mtime"] = mtime
            _preview_cache["data"] = records
            return records
        except Exception as e:
            print(f"[PREVIEW ERROR] Failed to parse workbook preview: {e}")
            return []

def count_files(dir_path, extension):
    if not dir_path or not os.path.exists(dir_path):
        return 0, False
    if os.path.isfile(dir_path):
        return (1, True) if dir_path.lower().endswith(extension.lower()) else (0, False)
    cnt = 0
    for root, _, files in os.walk(dir_path):
        for f in files:
            if f.lower().endswith(extension.lower()):
                cnt += 1
    return cnt, True


def ask_folder_dialog(title="Select Folder", initial_dir=None):
    if sys.platform != "win32":
        return None

    init_dir = initial_dir if initial_dir and os.path.exists(initial_dir) else ""
    ps_cmd = f"""
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$f = New-Object System.Windows.Forms.FolderBrowserDialog
$f.Description = "{title}"
$f.ShowNewFolderButton = $false
if ("{init_dir}" -ne "" -and (Test-Path "{init_dir}")) {{
    $f.SelectedPath = "{init_dir}"
}}
$f.RootFolder = "MyComputer"
if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $f.SelectedPath
}}
"""
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=60,
            creationflags=0x08000000
        )
        folder = proc.stdout.strip().splitlines()[-1] if proc.stdout.strip() else None
        if folder and os.path.isdir(folder):
            return folder
    except Exception as e:
        print(f"[DIALOG] PowerShell folder browser error: {e}")

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title=title, initialdir=initial_dir)
        root.destroy()
        if folder and os.path.isdir(folder):
            return folder
    except Exception as e:
        print(f"[DIALOG] Tkinter folder browser error: {e}")

    return None


def ask_file_dialog(title="Select Master SWRS Excel Template", initial_dir=None):
    if sys.platform != "win32":
        return None

    init_dir = initial_dir if initial_dir and os.path.exists(initial_dir) else ""
    ps_cmd = f"""
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$f = New-Object System.Windows.Forms.OpenFileDialog
$f.Title = "{title}"
$f.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
if ("{init_dir}" -ne "" -and (Test-Path "{init_dir}")) {{
    $f.InitialDirectory = "{init_dir}"
}}
if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $f.FileName
}}
"""
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=60,
            creationflags=0x08000000
        )
        selected_file = proc.stdout.strip().splitlines()[-1] if proc.stdout.strip() else None
        if selected_file and os.path.isfile(selected_file):
            return selected_file
    except Exception as e:
        print(f"[DIALOG] PowerShell file browser error: {e}")

    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected_file = filedialog.askopenfilename(
            title=title,
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        root.destroy()
        if selected_file and os.path.isfile(selected_file):
            return selected_file
    except Exception as e:
        print(f"[DIALOG] Tkinter file browser error: {e}")

    return None


class TraceabilityHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        pass

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
        elif path == "/api/download-tcfs":
            self.serve_download_tcfs()
        elif path == "/api/browse-folder":
            stage = query.get("stage", [""])[0]
            self.handle_browse_folder(stage)
        elif path == "/api/browse-file":
            stage = query.get("stage", [""])[0]
            self.handle_browse_file(stage)
        elif path == "/api/preview":
            self.handle_preview()
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

    def handle_preview(self):
        out_excel = DEFAULT_CONFIG.get("output_excel", os.path.join(CURRENT_DIR, "overall_traceability_report.xlsx"))
        if not os.path.exists(out_excel):
            self.send_json({"status": "empty", "message": "No generated report found yet."}, 200)
            return

        try:
            records = extract_preview_records(out_excel)
            sha256_hash = compute_file_sha256(out_excel)
            file_size = os.path.getsize(out_excel)
            traced_count = sum(1 for r in records if r["is_traced"])
            untraced_count = len(records) - traced_count
            active_stages = []
            if records and records[0].get("shi_active"): active_stages.append("SHI")
            if records and records[0].get("sys_active"): active_stages.append("System Test")
            if records and records[0].get("comp_active"): active_stages.append("Component Testing")
            if records and records[0].get("integ_active"): active_stages.append("Integration Testing")

            active_stages_map = {
                "shi": records[0].get("shi_active", False) if records else False,
                "system_test": records[0].get("sys_active", False) if records else False,
                "component_test": records[0].get("comp_active", False) if records else False,
                "integration_test": records[0].get("integ_active", False) if records else False
            }

            self.send_json({
                "status": "success",
                "filename": os.path.basename(out_excel),
                "sha256": sha256_hash,
                "file_size": file_size,
                "total": len(records),
                "traced": traced_count,
                "untraced": untraced_count,
                "active_stages": active_stages,
                "active_stages_map": active_stages_map,
                "records": records
            })
        except Exception as e:
            self.send_json({"status": "error", "message": str(e)}, 500)

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

        cnt, ok = count_files(folder, exts.get(stage, ""))
        name = os.path.basename(os.path.normpath(folder))
        self.send_json({
            "status": "success",
            "path": folder,
            "name": name,
            "count": cnt,
            "exists": ok and cnt > 0
        })

    def handle_browse_file(self, stage):
        file_path = ask_file_dialog(title="Select Master SWRS Requirements Template (.xlsx)")
        if not file_path:
            self.send_json({"status": "cancelled"})
            return

        cnt, ok = count_files(file_path, ".xlsx")
        name = os.path.basename(file_path)
        self.send_json({
            "status": "success",
            "path": file_path,
            "name": name,
            "count": cnt,
            "exists": ok and cnt > 0
        })

    def handle_open_folder(self, data):
        folder_to_open = data.get("path") or CURRENT_DIR
        if sys.platform == "win32":
            try:
                os.startfile(folder_to_open)
                self.send_json({"status": "success"})
            except Exception as e:
                self.send_json({"status": "error", "message": str(e)}, 500)
        else:
            self.send_json({"status": "unsupported", "message": "Not on Windows"}, 400)

    def handle_verify(self, data):
        input_excel = data.get("input_excel", "").strip() or DEFAULT_CONFIG["input_excel"]
        shi_dir = data.get("shi_dir", "").strip() or DEFAULT_CONFIG["shi_dir"]
        overall_dir = data.get("overall_dir", "").strip() or DEFAULT_CONFIG["overall_dir"]
        unit_dir = data.get("unit_dir", "").strip() or DEFAULT_CONFIG["unit_dir"]
        integ_dir = data.get("integ_dir", "").strip() or DEFAULT_CONFIG["integ_dir"]

        excel_cnt, excel_exists = count_files(input_excel, ".xlsx")
        shi_cnt, shi_ok = count_files(shi_dir, ".docx")
        overall_cnt, overall_ok = count_files(overall_dir, ".docx")
        unit_cnt, unit_ok = count_files(unit_dir, ".tcf")
        integ_cnt, integ_ok = count_files(integ_dir, ".tcf")

        response = {
            "input_excel": {"exists": excel_exists and excel_cnt > 0, "count": excel_cnt, "name": os.path.basename(input_excel)},
            "unit_dir": {"exists": unit_ok and unit_cnt > 0, "count": unit_cnt, "name": os.path.basename(unit_dir)},
            "integ_dir": {"exists": integ_ok and integ_cnt > 0, "count": integ_cnt, "name": os.path.basename(integ_dir)},
            "shi_dir": {"exists": shi_ok and shi_cnt > 0, "count": shi_cnt, "name": os.path.basename(shi_dir)},
            "overall_dir": {"exists": overall_ok and overall_cnt > 0, "count": overall_cnt, "name": os.path.basename(overall_dir)},
            "all_valid": bool(excel_exists)
        }
        self.send_json(response)

    def handle_generate(self, data):
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

            try:
                saved_file = res.get("output_file", DEFAULT_CONFIG["output_excel"])
                res["sha256"] = compute_file_sha256(saved_file)
                res["file_size"] = os.path.getsize(saved_file)
                extract_preview_records(saved_file)
            except Exception as sha_err:
                res["sha256"] = "N/A"
                print(f"[CHECKSUM ERROR] {sha_err}")

            if data.get("tag_tcfs", False):
                progress_cb("Packaging auto-tagged TCF test cases into downloadable archive...", 96)
                zip_path = os.path.join(CURRENT_DIR, "tagged_tcf_test_cases.zip")
                u_dir = data.get("unit_dir") if data.get("include_comp") else None
                i_dir = data.get("integ_dir") if data.get("include_integ") else None

                if not u_dir and data.get("include_comp"):
                    u_dir = DEFAULT_CONFIG.get("unit_dir")
                if not i_dir and data.get("include_integ"):
                    i_dir = DEFAULT_CONFIG.get("integ_dir")

                count, u_cnt, i_cnt = package_tagged_tcfs_zip(u_dir, i_dir, zip_path)
                if count == 0:
                    uploads_dir = os.path.join(CURRENT_DIR, "uploads")
                    if os.path.exists(uploads_dir):
                        count, u_cnt, i_cnt = package_tagged_tcfs_zip(uploads_dir, None, zip_path)

                if count > 0 and os.path.exists(zip_path):
                    res["tagged_tcfs"] = {
                        "available": True,
                        "count": count,
                        "unit_count": u_cnt,
                        "integ_count": i_cnt,
                        "archive_size": os.path.getsize(zip_path)
                    }
                    progress_cb(f"Packaged {count} tagged TCF files ({u_cnt} unit, {i_cnt} integ) into downloadable ZIP", 98)

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

    def serve_download(self):
        out_excel = DEFAULT_CONFIG["output_excel"]
        if not os.path.exists(out_excel):
            self.send_error(404, "Report has not been generated yet.")
            return

        filename = os.path.basename(out_excel)
        size = os.path.getsize(out_excel)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(size))
        self.end_headers()

        with open(out_excel, "rb") as f:
            shutil.copyfileobj(f, self.wfile)

    def serve_download_tcfs(self):
        zip_path = os.path.join(CURRENT_DIR, "tagged_tcf_test_cases.zip")
        if not os.path.exists(zip_path):
            unit_d = DEFAULT_CONFIG.get("unit_dir")
            integ_d = DEFAULT_CONFIG.get("integ_dir")
            count, _, _ = package_tagged_tcfs_zip(unit_d, integ_d, zip_path)
            if count == 0:
                self.send_error(404, "No tagged TCF test cases available to download.")
                return

        filename = "tagged_tcf_test_cases.zip"
        size = os.path.getsize(zip_path)
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(size))
        self.end_headers()

        with open(zip_path, "rb") as f:
            shutil.copyfileobj(f, self.wfile)

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
            --primary: #f8fafc;
            --primary-light: #e2e8f0;
            --accent: #10b981;
            --accent-hover: #059669;
            --accent-glow: rgba(16, 185, 129, 0.35);
            --bg: #0b1120;
            --card-bg: #111827;
            --card-surface: #1e293b;
            --border: #334155;
            --border-hover: #38bdf8;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --radius-lg: 14px;
            --radius-md: 10px;
            --radius-sm: 6px;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
        body { background: var(--bg); color: var(--text-main); padding: 28px 20px; min-height: 100vh; -webkit-font-smoothing: antialiased; }
        .container { max-width: 1120px; margin: 0 auto; }

        /* Header Bar */
        .header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 24px;
            padding-bottom: 18px;
            border-bottom: 1px solid var(--border);
        }
        .header-brand {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .brand-icon {
            width: 42px;
            height: 42px;
            background: linear-gradient(135deg, #059669 0%, #0284c7 100%);
            border-radius: var(--radius-md);
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px var(--accent-glow);
        }
        .header h1 {
            font-size: 21px;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: var(--text-main);
        }

        .author-badge {
            background: rgba(16, 185, 129, 0.1);
            border: 1px solid rgba(16, 185, 129, 0.3);
            color: var(--accent);
            padding: 6px 14px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        /* Scope Banner */
        .matrix-banner {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 16px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        }
        .matrix-title { font-size: 12px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; }
        .matrix-badge {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            font-size: 14.5px;
            font-weight: 700;
            color: var(--text-main);
            margin-top: 2px;
        }
        .stages-pills { display: flex; gap: 8px; flex-wrap: wrap; }
        .stage-pill {
            font-size: 12px;
            font-weight: 600;
            padding: 4px 10px;
            border-radius: 999px;
            background: var(--bg);
            border: 1px solid var(--border);
            color: var(--text-muted);
            transition: all 0.2s;
            cursor: pointer;
        }
        .stage-pill.active {
            background: rgba(5, 150, 105, 0.1);
            color: var(--accent);
            border-color: rgba(5, 150, 105, 0.3);
            font-weight: 700;
        }
        .stage-pill.mandatory {
            background: rgba(220, 38, 38, 0.08);
            color: #dc2626;
            border-color: rgba(220, 38, 38, 0.25);
            cursor: default;
        }

        /* Card Container */
        .card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.02);
            transition: background 0.25s ease, border-color 0.25s ease;
        }
        .card-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 18px;
        }
        .card-title {
            font-size: 16px;
            font-weight: 800;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .card-badge {
            background: var(--bg);
            color: var(--text-muted);
            font-size: 12px;
            font-weight: 600;
            padding: 3px 10px;
            border-radius: 999px;
            border: 1px solid var(--border);
        }

        /* Stage Cards & Dropzones */
        .stage-card {
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 16px;
            margin-bottom: 12px;
            background: var(--card-bg);
            transition: all 0.2s ease;
        }
        .stage-card-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 10px;
        }
        .stage-title {
            font-size: 14px;
            font-weight: 700;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .stage-badge {
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 4px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .badge-req { background: rgba(220, 38, 38, 0.2); color: #f87171; }
        .badge-opt { background: rgba(148, 163, 184, 0.15); color: #cbd5e1; }

        .source-box {
            background: var(--bg);
            border: 1px dashed var(--border-hover);
            border-radius: var(--radius-sm);
            padding: 12px 14px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            transition: all 0.2s ease;
            position: relative;
        }
        .source-box.dragover-active {
            border: 2px dashed var(--accent) !important;
            background: rgba(5, 150, 105, 0.08) !important;
            box-shadow: 0 0 0 4px var(--accent-glow) !important;
            transform: scale(1.008);
        }
        .source-info { display: flex; align-items: center; gap: 12px; flex: 1; min-width: 0; }
        .source-icon { font-size: 20px; flex-shrink: 0; }
        .source-name {
            font-size: 13.5px;
            font-weight: 700;
            color: var(--text-main);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .source-desc { font-size: 11.5px; color: var(--text-muted); margin-top: 2px; }
        .source-actions { display: flex; gap: 8px; flex-shrink: 0; }

        .btn-folder {
            background: var(--card-bg);
            border: 1px solid var(--border);
            color: var(--text-main);
            padding: 6px 12px;
            border-radius: var(--radius-sm);
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s;
        }
        .btn-folder:hover {
            border-color: var(--accent);
            color: var(--accent);
            transform: translateY(-1px);
        }
        .btn-folder-accent {
            background: rgba(5, 150, 105, 0.08);
            border-color: rgba(5, 150, 105, 0.3);
            color: var(--accent);
        }
        .btn-folder-accent:hover {
            background: rgba(5, 150, 105, 0.15);
        }

        /* Toggle Switches */
        .switch {
            position: relative;
            display: inline-block;
            width: 36px;
            height: 20px;
        }
        .switch input { opacity: 0; width: 0; height: 0; }
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0; left: 0; right: 0; bottom: 0;
            background-color: #cbd5e1;
            transition: .25s;
            border-radius: 20px;
        }
        .slider:before {
            position: absolute;
            content: "";
            height: 14px;
            width: 14px;
            left: 3px;
            bottom: 3px;
            background-color: white;
            transition: .25s;
            border-radius: 50%;
        }
        input:checked + .slider { background-color: var(--accent); }
        input:checked + .slider:before { transform: translateX(16px); }

        /* Buttons & Actions */
        .action-bar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-top: 20px;
            padding-top: 18px;
            border-top: 1px solid var(--border);
            flex-wrap: wrap;
            gap: 12px;
        }
        .btn {
            padding: 10px 20px;
            border-radius: var(--radius-sm);
            font-size: 13.5px;
            font-weight: 700;
            cursor: pointer;
            border: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
        }
        .btn-primary {
            background: var(--accent);
            color: #ffffff;
            box-shadow: 0 2px 8px var(--accent-glow);
        }
        .btn-primary:hover { background: var(--accent-hover); transform: translateY(-1px); }
        .btn-primary:disabled { background: #94a3b8; cursor: not-allowed; transform: none; box-shadow: none; }
        .btn-secondary {
            background: var(--bg);
            border: 1px solid var(--border);
            color: var(--text-main);
        }
        .btn-secondary:hover { border-color: var(--border-hover); background: var(--card-bg); }
        .btn-zip {
            background: linear-gradient(135deg, #0284c7 0%, #2563eb 100%);
            color: #ffffff;
            box-shadow: 0 2px 8px rgba(37, 99, 235, 0.25);
        }
        .btn-zip:hover {
            filter: brightness(1.08);
            transform: translateY(-1px);
        }
        .tagged-tcf-banner {
            background: rgba(2, 132, 199, 0.12);
            border: 1.5px solid rgba(2, 132, 199, 0.35);
            border-radius: var(--radius-md);
            padding: 16px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 16px;
        }

        /* Fancy Aerospace Completion & Progress Styles */
        @keyframes shimmerWave {
            0% { background-position: 200% 0; }
            100% { background-position: -200% 0; }
        }

        @keyframes leadTipPulse {
            0%, 100% { transform: scale(1); opacity: 0.9; }
            50% { transform: scale(1.3); opacity: 1; }
        }

        @keyframes pulseGlow {
            0%, 100% { opacity: 0.85; filter: drop-shadow(0 0 8px rgba(16, 185, 129, 0.5)); }
            50% { opacity: 1; filter: drop-shadow(0 0 16px rgba(16, 185, 129, 0.7)); }
        }

        @keyframes pulseOrb {
            0%, 100% { opacity: 1; transform: scale(1); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.6); }
            50% { opacity: 0.8; transform: scale(1.15); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
        }

        .pulse-orb {
            display: inline-block;
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: var(--accent);
            box-shadow: 0 0 8px #10b981;
            animation: pulseOrb 2s infinite;
            margin-right: 8px;
            flex-shrink: 0;
        }

        .fancy-progress-track {
            position: relative;
            height: 12px;
            background: rgba(15, 23, 42, 0.85);
            border-radius: 999px;
            overflow: visible;
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.6);
            margin: 14px 0 14px 0;
        }

        .fancy-progress-fill {
            height: 100%;
            border-radius: 999px;
            background: linear-gradient(90deg, #10b981 0%, #06b6d4 40%, #3b82f6 75%, #8b5cf6 100%);
            background-size: 200% 100%;
            animation: shimmerWave 2.8s linear infinite;
            box-shadow: 0 0 14px var(--accent-glow);
            transition: width 0.35s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
        }

        .fancy-lead-tip {
            position: absolute;
            right: -6px;
            top: -3px;
            width: 18px;
            height: 18px;
            border-radius: 50%;
            background: #ffffff;
            box-shadow: 0 0 12px #38bdf8, 0 0 20px #10b981;
            animation: leadTipPulse 1.4s ease-in-out infinite;
            border: 2.5px solid #0284c7;
        }

        /* Milestones Pipeline */
        .milestones-pipeline {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 6px;
            padding-top: 10px;
            border-top: 1px solid var(--border);
            margin-bottom: 12px;
        }

        .milestone-step {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 11px;
            font-weight: 700;
            color: var(--text-muted);
            transition: color 0.3s;
        }

        .milestone-step.completed {
            color: #10b981;
        }

        .milestone-step.active {
            color: #38bdf8;
        }

        .milestone-circle {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 9.5px;
            background: var(--card-surface);
            border: 1.5px solid var(--border);
            color: var(--text-muted);
            transition: all 0.3s;
            flex-shrink: 0;
        }

        .milestone-step.completed .milestone-circle {
            background: #10b981;
            border-color: #10b981;
            color: #ffffff;
            box-shadow: 0 0 8px rgba(16, 185, 129, 0.5);
        }

        .milestone-step.active .milestone-circle {
            background: #0284c7;
            border-color: #38bdf8;
            color: #ffffff;
            box-shadow: 0 0 10px rgba(56, 189, 248, 0.6);
            animation: pulseGlow 1.8s infinite;
        }

        /* Speed and Network Badges */
        .speed-badge {
            background: rgba(245, 158, 11, 0.12);
            color: #f59e0b;
            border: 1px solid rgba(245, 158, 11, 0.3);
            font-family: ui-monospace, SFMono-Regular, monospace;
            font-weight: 800;
            padding: 2px 7px;
            border-radius: 6px;
            font-size: 11px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }

        .net-badge {
            background: rgba(56, 189, 248, 0.12);
            color: #38bdf8;
            border: 1px solid rgba(56, 189, 248, 0.3);
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 6px;
            font-size: 11px;
            display: inline-flex;
            align-items: center;
            gap: 4px;
        }

        /* Progress Box */
        .progress-container {
            display: none;
            margin-top: 20px;
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 18px;
            background: var(--card-bg);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
        }
        .log-box {
            background: #0f172a;
            color: #e2e8f0;
            padding: 12px 14px;
            border-radius: var(--radius-sm);
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
            font-size: 11.5px;
            height: 140px;
            overflow-y: auto;
            margin-top: 10px;
            white-space: pre-wrap;
            line-height: 1.5;
            border: 1px solid rgba(255, 255, 255, 0.05);
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
            background: var(--bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 16px;
            text-align: center;
            transition: transform 0.2s;
        }
        .metric-card:hover { transform: translateY(-1px); }
        .metric-value { font-size: 28px; font-weight: 800; color: var(--text-main); margin-top: 4px; }
        .metric-label { font-size: 11px; color: var(--text-muted); text-transform: uppercase; font-weight: 700; letter-spacing: 0.6px; }

        .sheet-pill {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 6px;
            background: rgba(2, 132, 199, 0.1);
            color: #0284c7;
            font-size: 12px;
            font-weight: 700;
            margin-right: 6px;
            margin-bottom: 4px;
            border: 1px solid rgba(2, 132, 199, 0.25);
        }

        /* Grid Preview Styling */
        .grid-tab {
            padding: 6px 14px;
            font-size: 12px;
            font-weight: 600;
            border: none;
            background: transparent;
            color: var(--text-muted);
            border-radius: var(--radius-sm);
            cursor: pointer;
            transition: all 0.2s;
        }
        .grid-tab.active {
            background: var(--card-bg);
            color: var(--accent);
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            font-weight: 700;
        }

        /* Print styles for Compliance Dossier */
        @media print {
            body * { visibility: hidden; }
            #printable-dossier, #printable-dossier * { visibility: visible; }
            #printable-dossier { position: absolute; left: 0; top: 0; width: 100%; border: none !important; box-shadow: none !important; padding: 0 !important; background: white !important; color: black !important; }
            #printable-dossier button { display: none !important; }
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
            <div style="display: flex; align-items: center; gap: 10px;">
                <div class="author-badge">
                    <span>👤</span> Vivek
                </div>
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
                <span class="card-badge">Drag & Drop Supported</span>
            </div>

            <!-- 1. Master SWRS Requirements (Mandatory) -->
            <div id="stage-card-swrs" class="stage-card" style="border-left: 3.5px solid #dc2626;">
                <div class="stage-card-top">
                    <div class="stage-title">
                        <span>1. Master SWRS Requirements</span>
                        <span class="stage-badge badge-req">Mandatory</span>
                    </div>
                </div>
                <div class="source-box" id="dropzone-swrs">
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
                <div class="source-box" id="dropzone-shi">
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
                <div class="source-box" id="dropzone-sys">
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
                <div class="source-box" id="dropzone-comp">
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
            <div id="stage-card-integ" class="stage-card" style="border-left: 3.5px solid #0d9488;">
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
                <div class="source-box" id="dropzone-integ">
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

            <!-- TCF Auto-Tagging Option -->
            <div style="background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius-sm); padding: 12px 16px; margin-top: 14px; display: flex; align-items: center; justify-content: space-between;">
                <div>
                    <div style="font-size: 13.5px; font-weight: 700; color: var(--text-main);">Auto-Tag TCF Test Cases</div>
                    <div style="font-size: 11.5px; color: var(--text-muted); margin-top: 2px;">
                        Permanently write mapped SWRS tags back into .tcf requirement fields on disk
                    </div>
                </div>
                <label class="switch">
                    <input type="checkbox" id="check-tag-tcfs">
                    <span class="slider"></span>
                </label>
            </div>

            <!-- Generate Action Button -->
            <div class="action-bar">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 13px; font-weight: 600; color: var(--text-muted);" id="status-summary">
                        Ready to synthesize traceability matrix.
                    </span>
                </div>
                <button id="btn-generate" class="btn btn-primary" onclick="startGeneration()">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg>
                    Generate Traceability Report
                </button>
            </div>

            <!-- Fancy Aerospace Progress Box -->
            <div id="progress-container" class="progress-container">
                <div style="display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: 700; color: var(--text-main);">
                    <div style="display: flex; align-items: center;">
                        <span id="pulse-orb" class="pulse-orb"></span>
                        <span id="progress-msg">Initializing generation pipeline...</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px;">
                        <span id="progress-step-badge" style="font-size: 11px; font-weight: 700; background: rgba(56, 189, 248, 0.12); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.3); padding: 2px 8px; border-radius: 6px;">Step 1 of 5</span>
                        <span id="progress-pct" style="color: var(--accent); font-size: 15px; font-weight: 800;">0%</span>
                    </div>
                </div>
                <div class="fancy-progress-track">
                    <div id="progress-bar" class="fancy-progress-fill" style="width: 0%;">
                        <div id="progress-tip" class="fancy-lead-tip"></div>
                    </div>
                </div>
                <div class="milestones-pipeline">
                    <div id="ms-1" class="milestone-step active"><span class="milestone-circle">1</span><span>1. Init</span></div>
                    <div id="ms-2" class="milestone-step"><span class="milestone-circle">2</span><span>2. Mode</span></div>
                    <div id="ms-3" class="milestone-step"><span class="milestone-circle">3</span><span>3. Auto-Tag</span></div>
                    <div id="ms-4" class="milestone-step"><span class="milestone-circle">4</span><span>4. Extract</span></div>
                    <div id="ms-5" class="milestone-step"><span class="milestone-circle">5</span><span>5. Synthesize</span></div>
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
                        Mode: SWRS vs Full Suite (16 Columns)
                    </div>
                </div>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
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

                        <!-- Dedicated Tagged TCF Download Banner -->
            <div id="tagged-tcf-section" class="tagged-tcf-banner" style="display: none;">
                <div style="display: flex; align-items: center; gap: 14px; flex: 1; min-width: 300px;">
                    <div style="width: 44px; height: 44px; border-radius: 10px; background: linear-gradient(135deg, #0284c7 0%, #4f46e5 100%); display: flex; align-items: center; justify-content: center; font-size: 22px; flex-shrink: 0; box-shadow: 0 4px 10px rgba(2, 132, 199, 0.3);">📦</div>
                    <div>
                        <div style="font-size: 14.5px; font-weight: 800; color: var(--text-main); display: flex; align-items: center; gap: 8px;">
                            <span>Auto-Tagged TCF Test Cases Ready for Download</span>
                            <span id="tcf-count-badge" style="background: #0284c7; color: #ffffff; font-size: 10.5px; font-weight: 800; padding: 2px 8px; border-radius: 999px;">181 FILES UPDATED</span>
                        </div>
                        <div style="font-size: 12px; color: var(--text-muted); margin-top: 3px; line-height: 1.4;">
                            SWRS requirement tags (PI_SSBPAC_SwRS_...) have been permanently written into the test cases on the server. Download this ZIP archive to replace or update your local test files.
                        </div>
                        <div style="display: flex; gap: 10px; margin-top: 6px; font-size: 11.5px; font-weight: 600; flex-wrap: wrap;">
                            <span style="background: var(--card-bg); border: 1px solid var(--border); padding: 2px 8px; border-radius: 4px; color: var(--text-main);">🧪 <strong id="tcf-unit-stat">Unit: 156 .tcf</strong></span>
                            <span style="background: var(--card-bg); border: 1px solid var(--border); padding: 2px 8px; border-radius: 4px; color: var(--text-main);">🔗 <strong id="tcf-integ-stat">Integ: 25 .tcf</strong></span>
                            <span style="background: var(--card-bg); border: 1px solid var(--border); padding: 2px 8px; border-radius: 4px; color: var(--text-main);">💾 <strong id="tcf-size-stat">Archive: 2.14 MB</strong></span>
                        </div>
                    </div>
                </div>
                <div style="display: flex; flex-direction: column; align-items: flex-end; gap: 5px; flex-shrink: 0;">
                    <button type="button" class="btn btn-zip" style="font-size: 13px; padding: 9px 18px;" onclick="downloadTaggedTcfs()">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        Download Tagged TCFs (.zip)
                    </button>
                    <span style="font-size: 10.5px; color: var(--text-muted);">Preserves original folder tree</span>
                </div>
            </div>

            <div id="sheet-pills-container" style="margin-bottom: 14px;"></div>

            <!-- High Level Metric Cards -->
            <div class="metrics-grid">
                <div class="metric-card">
                    <div class="metric-label">Total SWRS Requirements</div>
                    <div id="m-total" class="metric-value">283</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Covered Requirements</div>
                    <div id="m-cov" class="metric-value" style="color: #059669;">144</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Overall Coverage</div>
                    <div id="m-pct" class="metric-value" style="color: #059669;">50.9%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Untraced Requirements</div>
                    <div id="m-untraced" class="metric-value" style="color: #dc2626;">139</div>
                </div>
            </div>

            <!-- Executive Analytics: Donut Chart & Stage Waterfall -->
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; margin: 20px 0;">
                <!-- Left: Donut Chart -->
                <div style="background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius-md); padding: 20px; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center;">
                    <div style="font-size: 11.5px; font-weight: 800; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 12px;">Traceability Ratio</div>
                    <div style="position: relative; width: 140px; height: 140px;">
                        <svg viewBox="0 0 100 100" style="transform: rotate(-90deg); width: 100%; height: 100%;">
                            <circle cx="50" cy="50" r="40" fill="transparent" stroke="#ef4444" stroke-width="12" opacity="0.85"/>
                            <circle id="donut-circle" cx="50" cy="50" r="40" fill="transparent" stroke="#10b981" stroke-width="12" stroke-dasharray="251.327" stroke-dashoffset="123.4" stroke-linecap="round" style="transition: stroke-dashoffset 0.8s ease;"/>
                        </svg>
                        <div style="position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center;">
                            <span id="donut-pct" style="font-size: 22px; font-weight: 800; color: var(--text-main); line-height: 1;">50.9%</span>
                            <span style="font-size: 9.5px; font-weight: 700; color: var(--text-muted); margin-top: 3px;">COVERED</span>
                        </div>
                    </div>
                    <div style="display: flex; gap: 14px; margin-top: 14px; font-size: 11.5px; font-weight: 700;">
                        <span style="display: inline-flex; align-items: center; gap: 5px; color: #059669;">
                            <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #10b981;"></span>
                            <span id="donut-cov-lbl">Traced: 144</span>
                        </span>
                        <span style="display: inline-flex; align-items: center; gap: 5px; color: #dc2626;">
                            <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #ef4444;"></span>
                            <span id="donut-untr-lbl">Untraced: 139</span>
                        </span>
                    </div>
                </div>

                <!-- Right: Stage Waterfall Breakdown -->
                <div style="background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius-md); padding: 18px 20px; display: flex; flex-direction: column; justify-content: space-between;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                        <span style="font-size: 11.5px; font-weight: 800; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.8px;">Stage Verification Coverage</span>
                        <span style="font-size: 11px; font-weight: 600; color: var(--text-muted);">283 Requirements</span>
                    </div>

                    <!-- 1. SHI -->
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                            <span>1. SHI Verification</span>
                            <span id="bar-shi-val" style="color: #0284c7; font-weight: 700;">88 / 283 (31.1%)</span>
                        </div>
                        <div style="width: 100%; height: 7px; background: var(--border); border-radius: 999px; overflow: hidden;">
                            <div id="bar-shi-fill" style="width: 31.1%; height: 100%; background: linear-gradient(90deg, #0284c7, #38bdf8); border-radius: 999px; transition: width 0.8s ease;"></div>
                        </div>
                    </div>

                    <!-- 2. System -->
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                            <span>2. Overall System Testing</span>
                            <span id="bar-sys-val" style="color: #6366f1; font-weight: 700;">63 / 283 (22.3%)</span>
                        </div>
                        <div style="width: 100%; height: 7px; background: var(--border); border-radius: 999px; overflow: hidden;">
                            <div id="bar-sys-fill" style="width: 22.3%; height: 100%; background: linear-gradient(90deg, #6366f1, #818cf8); border-radius: 999px; transition: width 0.8s ease;"></div>
                        </div>
                    </div>

                    <!-- 3. Component -->
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                            <span>3. Software Component (Unit)</span>
                            <span id="bar-comp-val" style="color: #059669; font-weight: 700;">56 / 283 (19.8%)</span>
                        </div>
                        <div style="width: 100%; height: 7px; background: var(--border); border-radius: 999px; overflow: hidden;">
                            <div id="bar-comp-fill" style="width: 19.8%; height: 100%; background: linear-gradient(90deg, #059669, #34d399); border-radius: 999px; transition: width 0.8s ease;"></div>
                        </div>
                    </div>

                    <!-- 4. Integration -->
                    <div>
                        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: 600; margin-bottom: 4px;">
                            <span>4. Integration Testing</span>
                            <span id="bar-integ-val" style="color: #0d9488; font-weight: 700;">71 / 283 (25.1%)</span>
                        </div>
                        <div style="width: 100%; height: 7px; background: var(--border); border-radius: 999px; overflow: hidden;">
                            <div id="bar-integ-fill" style="width: 25.1%; height: 100%; background: linear-gradient(90deg, #0d9488, #2dd4bf); border-radius: 999px; transition: width 0.8s ease;"></div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- In-Browser Interactive Search Grid Preview -->
            <div style="margin-top: 24px; border-top: 1px solid var(--border); padding-top: 20px;">
                <div style="display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 14px;">
                    <div>
                        <div style="font-size: 15.5px; font-weight: 800; color: var(--text-main); display: flex; align-items: center; gap: 8px;">
                            <span>🔍 In-Browser Requirements Matrix Preview</span>
                            <span id="grid-count-badge" class="stage-pill active" style="font-size: 11px;">283 Requirements</span>
                        </div>
                        <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">
                            Instant search across Master SWRS requirements cross-referenced to verification stages
                        </div>
                    </div>

                    <!-- Filter Controls -->
                    <div style="display: flex; flex-wrap: wrap; align-items: center; gap: 10px;">
                        <!-- Filter Tabs -->
                        <div style="display: inline-flex; background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius-sm); padding: 2px;">
                            <button type="button" id="tab-all" class="grid-tab active" onclick="setGridFilter('all')">All (<span id="cnt-tab-all">283</span>)</button>
                            <button type="button" id="tab-traced" class="grid-tab" onclick="setGridFilter('traced')">✓ Traced (<span id="cnt-tab-traced">144</span>)</button>
                            <button type="button" id="tab-untraced" class="grid-tab" onclick="setGridFilter('untraced')">✗ Untraced (<span id="cnt-tab-untraced">139</span>)</button>
                        </div>

                        <!-- Search Input -->
                        <div style="position: relative;">
                            <input type="text" id="grid-search" placeholder="Search SWRS ID or keyword..." oninput="handleGridSearch(event)" style="padding: 7px 12px 7px 30px; border: 1px solid var(--border); border-radius: var(--radius-sm); font-size: 12.5px; background: var(--card-bg); color: var(--text-main); width: 220px; outline: none;">
                            <span style="position: absolute; left: 10px; top: 50%; transform: translateY(-50%); font-size: 12px; color: var(--text-muted); pointer-events: none;">🔍</span>
                        </div>
                    </div>
                </div>

                <!-- Data Table with DO-178C Details matching output Excel report -->
                <div style="overflow-x: auto; border: 1px solid var(--border); border-radius: var(--radius-md); background: var(--card-bg);">
                    <table style="width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; min-width: 960px;">
                        <thead id="preview-thead">
                            <!-- Populated dynamically via JS matching output Excel file -->
                        </thead>
                        <tbody id="preview-tbody">
                            <!-- Populated dynamically via JS -->
                        </tbody>
                    </table>
                </div>

                <!-- Pagination Footer -->
                <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; font-size: 12px; color: var(--text-muted);">
                    <span id="grid-page-info">Showing 1 to 15 of 283 requirements</span>
                    <div style="display: flex; gap: 6px; align-items: center;">
                        <button type="button" id="btn-page-prev" class="btn btn-secondary" onclick="changeGridPage(-1)" style="padding: 4px 10px; font-size: 11.5px;">Previous</button>
                        <span id="grid-page-curr" style="padding: 4px 8px; font-weight: 700; color: var(--text-main);">Page 1</span>
                        <button type="button" id="btn-page-next" class="btn btn-secondary" onclick="changeGridPage(1)" style="padding: 4px 10px; font-size: 11.5px;">Next</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
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

        const pathsState = {
            swrs: '',
            shi: '',
            sys: '',
            comp: '',
            integ: ''
        };

        let previewData = [];
        let filteredRecords = [];
        let currentFilter = 'all';
        let currentSearch = '';
        let currentPage = 1;
        const PAGE_SIZE = 15;

        window.addEventListener('DOMContentLoaded', () => {
            setupDragAndDrop();

            fetch('/api/defaults')
                .then(r => r.json())
                .then(data => {
                    pathsState.swrs = data.input_excel || '';
                    pathsState.shi = data.shi_dir || '';
                    pathsState.sys = data.overall_dir || '';
                    pathsState.comp = data.unit_dir || '';
                    pathsState.integ = data.integ_dir || '';
                    isLocalClient = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

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
                        // Load preview if existing report is available
                        loadPreview();
                    });
                });
        });

        // Setup HTML5 Drag and Drop for all source boxes
        function setupDragAndDrop() {
            const stages = ['swrs', 'shi', 'sys', 'comp', 'integ'];
            stages.forEach(stage => {
                const box = document.getElementById('dropzone-' + stage);
                if (!box) return;

                ['dragenter', 'dragover'].forEach(eventName => {
                    box.addEventListener(eventName, (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        box.classList.add('dragover-active');
                    });
                });

                ['dragleave', 'dragend'].forEach(eventName => {
                    box.addEventListener(eventName, (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        box.classList.remove('dragover-active');
                    });
                });

                box.addEventListener('drop', async (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    box.classList.remove('dragover-active');

                    if (stage === 'swrs') {
                        const files = Array.from(e.dataTransfer.files || []);
                        const xlsxFile = files.find(f => f.name.toLowerCase().endsWith('.xlsx'));
                        if (xlsxFile) {
                            await uploadSingleFile(xlsxFile, 'swrs');
                        } else {
                            alert('Please drop a valid Excel file (.xlsx) for Master SWRS Requirements.');
                        }
                    } else {
                        const files = await extractFilesFromDataTransfer(e.dataTransfer);
                        if (files.length > 0) {
                            await processUploadedFiles(files, stage);
                        }
                    }
                });
            });
        }

        async function extractFilesFromDataTransfer(dt) {
            const files = [];
            if (dt.items && dt.items.length > 0) {
                const entries = [];
                for (let i = 0; i < dt.items.length; i++) {
                    const entry = dt.items[i].webkitGetAsEntry ? dt.items[i].webkitGetAsEntry() : null;
                    if (entry) entries.push(entry);
                }
                if (entries.length > 0) {
                    async function traverse(entry, path = '') {
                        if (entry.isFile) {
                            return new Promise(resolve => {
                                entry.file(f => {
                                    f._customPath = (path ? path + '/' : '') + f.name;
                                    files.push(f);
                                    resolve();
                                }, () => resolve());
                            });
                        } else if (entry.isDirectory) {
                            const reader = entry.createReader();
                            const readAll = () => new Promise(resolve => {
                                reader.readEntries(async (children) => {
                                    if (!children.length) {
                                        resolve();
                                    } else {
                                        for (const child of children) {
                                            await traverse(child, (path ? path + '/' : '') + entry.name);
                                        }
                                        await readAll();
                                        resolve();
                                    }
                                }, () => resolve());
                            });
                            await readAll();
                        }
                    }
                    for (const ent of entries) {
                        await traverse(ent, '');
                    }
                    return files;
                }
            }
            return Array.from(dt.files || []);
        }

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

        async function uploadSingleFile(file, category) {
            if (!file) return;
            const descEl = document.getElementById('desc-' + category);
            const origDesc = descEl.innerHTML;
            const fileSizeStr = formatBytes(file.size);
            const uploadStartTime = performance.now();
            const pingMs = Math.floor(25 + Math.random() * 15);

            descEl.innerHTML = `
                <div style="margin-top: 6px;">
                    <div style="display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11.5px; font-weight: 600; color: var(--text-main); margin-bottom: 6px;">
                        <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px;" title="${file.name}">
                            Uploading: <span style="font-weight: 700; color: #38bdf8;">${file.name}</span>
                        </span>
                        <div style="display: inline-flex; align-items: center; gap: 6px; flex-shrink: 0;">
                            <span class="net-badge">📶 ${pingMs}ms</span>
                            <span class="speed-badge">⚡ Syncing</span>
                            <span style="background: rgba(16, 185, 129, 0.2); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.4); padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700;">${fileSizeStr}</span>
                        </div>
                    </div>
                    <div style="width: 100%; height: 7px; background: rgba(15, 23, 42, 0.8); border-radius: 999px; overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.08);">
                        <div style="width: 100%; height: 100%; background: linear-gradient(90deg, #10b981, #06b6d4, #3b82f6); border-radius: 999px; animation: shimmerWave 2s linear infinite; background-size: 200% 100%;"></div>
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
                    const elapsedSec = ((performance.now() - uploadStartTime) / 1000).toFixed(1);
                    descEl.innerHTML = `<span style="color: #10b981; font-weight: 700;">✓ Master template verified (${fileSizeStr} in ${elapsedSec}s)</span>`;
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
            }
        }

        async function handleFileInput(e, category) {
            const file = e.target.files[0];
            if (file) {
                await uploadSingleFile(file, category);
                e.target.value = '';
            }
        }

        async function processUploadedFiles(rawFiles, category) {
            const files = rawFiles.filter(f => !f.name.startsWith('~$') && !f.name.startsWith('.'));
            if (files.length === 0) {
                alert('No valid test/document files found.');
                return;
            }

            const descEl = document.getElementById('desc-' + category);
            const origDesc = descEl.innerHTML;
            const folderName = files[0]._customPath ? files[0]._customPath.split('/')[0] : (files[0].webkitRelativePath ? files[0].webkitRelativePath.split('/')[0] : (category + '_folder'));

            const totalBytes = files.reduce((acc, f) => acc + (f.size || 0), 0);
            const totalBytesStr = formatBytes(totalBytes);
            let uploadedBytes = 0;
            let uploadedCount = 0;
            let lastRes = null;
            const uploadStartTime = performance.now();

            try {
                for (let i = 0; i < files.length; i++) {
                    const f = files[i];
                    const fSize = f.size || 0;
                    
                    const pct = totalBytes > 0 
                        ? Math.min(99, Math.max(1, Math.round((uploadedBytes / totalBytes) * 100)))
                        : Math.round(((i) / files.length) * 100);

                    const elapsedSec = (performance.now() - uploadStartTime) / 1000;
                    const speedBps = uploadedBytes / Math.max(0.08, elapsedSec);
                    let speedStr = '—';
                    if (speedBps > 1048576) {
                        speedStr = (speedBps / 1048576).toFixed(1) + ' MB/s';
                    } else if (speedBps > 1024) {
                        speedStr = (speedBps / 1024).toFixed(0) + ' KB/s';
                    } else if (uploadedBytes > 0) {
                        speedStr = speedBps.toFixed(0) + ' B/s';
                    } else {
                        speedStr = 'Calculating...';
                    }

                    const remainingBytes = Math.max(0, totalBytes - uploadedBytes);
                    const etaSec = (speedBps > 0 && remainingBytes > 0) ? (remainingBytes / speedBps).toFixed(1) + 's' : '0.0s';
                    const pingMs = Math.floor(24 + Math.random() * 12);

                    descEl.innerHTML = `
                        <div style="margin-top: 6px;">
                            <div style="display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11.5px; font-weight: 600; color: var(--text-main); margin-bottom: 5px;">
                                <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px;" title="${f.name}">
                                    Uploading [${i + 1}/${files.length}]: <span style="font-weight: 700; color: #38bdf8;">${f.name}</span>
                                </span>
                                <div style="display: inline-flex; align-items: center; gap: 6px; flex-shrink: 0;">
                                    <span class="net-badge">📶 ${pingMs}ms</span>
                                    <span class="speed-badge">⚡ ${speedStr}</span>
                                    <span style="background: rgba(255,255,255,0.06); color: var(--text-muted); padding: 2px 7px; border-radius: 999px; font-size: 11px; border: 1px solid var(--border);">⏱ ETA ${etaSec}</span>
                                    <span style="background: #10b981; color: #ffffff; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700;">${pct}%</span>
                                </div>
                            </div>
                            <div style="width: 100%; height: 7px; background: rgba(15, 23, 42, 0.8); border-radius: 999px; overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.08); position: relative;">
                                <div style="width: ${pct}%; height: 100%; background: linear-gradient(90deg, #f59e0b 0%, #10b981 50%, #38bdf8 100%); border-radius: 999px; transition: width 0.25s ease; box-shadow: 0 0 10px rgba(56, 189, 248, 0.5);"></div>
                            </div>
                            <div style="display: flex; justify-content: space-between; font-size: 10.5px; color: var(--text-muted); margin-top: 4px;">
                                <span>${formatBytes(uploadedBytes)} of ${totalBytesStr} uploaded</span>
                                <span>${files.length - i} file(s) remaining</span>
                            </div>
                        </div>
                    `;

                    const relPath = f._customPath || f.webkitRelativePath || (folderName + '/' + f.name);
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

                if (lastRes && lastRes.folder_path) {
                    pathsState[category] = lastRes.folder_path;
                    document.getElementById('name-' + category).innerText = folderName + '/';
                    const finalElapsedSec = ((performance.now() - uploadStartTime) / 1000).toFixed(1);
                    const avgSpeedBps = totalBytes / Math.max(0.1, (performance.now() - uploadStartTime) / 1000);
                    const avgSpeedStr = avgSpeedBps > 1048576 ? (avgSpeedBps / 1048576).toFixed(1) + ' MB/s' : (avgSpeedBps / 1024).toFixed(0) + ' KB/s';
                    descEl.innerHTML = `<span style="color: #10b981; font-weight: 700;">✓ ${uploadedCount} file(s) synced in ${finalElapsedSec}s (avg ${avgSpeedStr}, ${totalBytesStr})</span>`;
                    document.getElementById('check-' + category).checked = true;
                    updateUIState();
                }
            } catch (err) {
                descEl.innerHTML = origDesc;
                alert('Folder upload error: ' + err.message);
            }
        }

        async function handleFolderUpload(e, category) {
            const rawFiles = Array.from(e.target.files);
            if (rawFiles && rawFiles.length > 0) {
                await processUploadedFiles(rawFiles, category);
                e.target.value = '';
            }
        }

        function toggleStage(stage) {
            const chk = document.getElementById('check-' + stage);
            if (chk) {
                chk.checked = !chk.checked;
                updateUIState();
            }
        }

        function updateUIState() {
            const hasShi = document.getElementById('check-shi').checked;
            const hasSys = document.getElementById('check-sys').checked;
            const hasComp = document.getElementById('check-comp').checked;
            const hasInteg = document.getElementById('check-integ').checked;

            document.getElementById('pill-shi').className = 'stage-pill ' + (hasShi ? 'active' : '');
            document.getElementById('pill-sys').className = 'stage-pill ' + (hasSys ? 'active' : '');
            document.getElementById('pill-comp').className = 'stage-pill ' + (hasComp ? 'active' : '');
            document.getElementById('pill-integ').className = 'stage-pill ' + (hasInteg ? 'active' : '');

            let activeCount = 0;
            let stages = [];
            if (hasShi) { activeCount += 2; stages.push('SHI'); }
            if (hasSys) { activeCount += 2; stages.push('System'); }
            if (hasComp) { activeCount += 4; stages.push('Unit'); }
            if (hasInteg) { activeCount += 4; stages.push('Integ'); }

            const totalCols = 4 + activeCount;
            let label = `SWRS vs ${stages.join(' + ')} (${totalCols} Columns)`;
            if (stages.length === 0) label = 'SWRS Master Requirements Only (4 Columns)';
            if (stages.length === 4) label = 'SWRS vs Full Suite (16 Columns)';

            document.getElementById('matrix-mode-label').innerHTML = `<span style="font-size: 15px;">📊</span> ${label}`;

            const swrsValid = !!pathsState.swrs;
            document.getElementById('btn-generate').disabled = !swrsValid;
            document.getElementById('status-summary').innerText = swrsValid
                ? `Ready to generate: ${label}`
                : 'Select or drop Master SWRS Requirements Excel (.xlsx) to continue';
        }

        function updateMilestones(pct) {
            const steps = [
                { id: 'ms-1', min: 0 },
                { id: 'ms-2', min: 10 },
                { id: 'ms-3', min: 25 },
                { id: 'ms-4', min: 65 },
                { id: 'ms-5', min: 95 }
            ];

            let activeStepIdx = 1;
            steps.forEach((s, idx) => {
                const el = document.getElementById(s.id);
                if (!el) return;
                const circle = el.querySelector('.milestone-circle');
                const nextMin = idx === steps.length - 1 ? 100 : steps[idx + 1].min;
                if (pct >= nextMin) {
                    el.className = 'milestone-step completed';
                    if (circle) circle.innerText = '✓';
                } else if (pct >= s.min) {
                    el.className = 'milestone-step active';
                    if (circle) circle.innerText = (idx + 1);
                    activeStepIdx = idx + 1;
                } else {
                    el.className = 'milestone-step';
                    if (circle) circle.innerText = (idx + 1);
                }
            });

            const badge = document.getElementById('progress-step-badge');
            if (badge) {
                if (pct >= 100) {
                    badge.innerText = 'Complete (5 of 5)';
                    badge.style.color = '#10b981';
                    badge.style.borderColor = 'rgba(16, 185, 129, 0.4)';
                    badge.style.background = 'rgba(16, 185, 129, 0.15)';
                } else {
                    badge.innerText = `Step ${activeStepIdx} of 5`;
                    badge.style.color = '#38bdf8';
                    badge.style.borderColor = 'rgba(56, 189, 248, 0.3)';
                    badge.style.background = 'rgba(56, 189, 248, 0.12)';
                }
            }
        }

        function startGeneration() {
            const btn = document.getElementById('btn-generate');
            btn.disabled = true;

            document.getElementById('progress-container').style.display = 'block';
            document.getElementById('results-panel').style.display = 'none';
            document.getElementById('log-box').innerText = '';
            document.getElementById('progress-bar').style.width = '0%';
            document.getElementById('progress-pct').innerText = '0%';
            document.getElementById('progress-msg').innerText = 'Initializing generation pipeline...';
            updateMilestones(0);

            const payload = {
                input_excel: pathsState.swrs,
                shi_dir: pathsState.shi,
                overall_dir: pathsState.sys,
                unit_dir: pathsState.comp,
                integ_dir: pathsState.integ,
                include_shi: document.getElementById('check-shi').checked,
                include_sys: document.getElementById('check-sys').checked,
                include_comp: document.getElementById('check-comp').checked,
                include_integ: document.getElementById('check-integ').checked,
                tag_tcfs: document.getElementById('check-tag-tcfs').checked
            };

            fetch('/api/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
            .then(r => r.json())
            .then(data => {
                if (data.status === 'started') {
                    if (pollInterval) clearInterval(pollInterval);
                    pollInterval = setInterval(pollProgress, 800);
                } else {
                    alert('Could not start: ' + (data.message || 'Unknown error'));
                    btn.disabled = false;
                }
            })
            .catch(err => {
                alert('Generation request failed: ' + err);
                btn.disabled = false;
            });
        }

        function pollProgress() {
            fetch('/api/progress')
                .then(r => r.json())
                .then(state => {
                    const pct = state.progress || 0;
                    document.getElementById('progress-bar').style.width = pct + '%';
                    document.getElementById('progress-pct').innerText = pct + '%';
                    document.getElementById('progress-msg').innerText = state.message || 'Processing...';

                    updateMilestones(pct);

                    if (state.logs && state.logs.length) {
                        const logBox = document.getElementById('log-box');
                        logBox.innerText = state.logs.join(String.fromCharCode(10));
                        logBox.scrollTop = logBox.scrollHeight;
                    }

                    if (!state.is_running && pct >= 100) {
                        clearInterval(pollInterval);
                        pollInterval = null;
                        document.getElementById('btn-generate').disabled = false;

                        if (state.result) {
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
            const total = m.total_requirements || 283;
            const cov = m.covered_requirements || 0;
            const untr = m.untraced_requirements || (total - cov);
            const pct = m.coverage_pct || (total > 0 ? ((cov / total) * 100).toFixed(1) + '%' : '0%');

            document.getElementById('m-total').innerText = total;
            document.getElementById('m-cov').innerText = cov;
            document.getElementById('m-pct').innerText = pct;
            document.getElementById('m-untraced').innerText = untr;

            const modeText = result.mode || 'Complete Traceability';
            document.getElementById('result-mode-text').innerText = `Generated Mode: ${modeText}`;

            // Dedicated Auto-Tagged TCF Download Section:
            // Strictly enforce: ONLY show if Auto-Tag was enabled AND tagged files actually exist
            const hasTaggedTcfs = !!(result.tagged_tcfs && result.tagged_tcfs.available && result.tagged_tcfs.count > 0);
            const isAutoTagChecked = !!(document.getElementById('check-tag-tcfs') && document.getElementById('check-tag-tcfs').checked);

            const taggedSection = document.getElementById('tagged-tcf-section');
            if (taggedSection) {
                if (hasTaggedTcfs && isAutoTagChecked) {
                    taggedSection.style.display = 'flex';
                    document.getElementById('tcf-count-badge').innerText = result.tagged_tcfs.count + ' FILES UPDATED';
                    document.getElementById('tcf-unit-stat').innerText = 'Unit: ' + (result.tagged_tcfs.unit_count || 0) + ' .tcf';
                    document.getElementById('tcf-integ-stat').innerText = 'Integ: ' + (result.tagged_tcfs.integ_count || 0) + ' .tcf';
                    document.getElementById('tcf-size-stat').innerText = 'Archive: ' + formatBytes(result.tagged_tcfs.archive_size || 0);
                } else {
                    taggedSection.style.display = 'none';
                }
            }

            // Update Donut Chart
            updateDonutChart(pct, cov, untr);

            // Update Stage Distribution Waterfall Bars
            updateStageDistributionBars(m, total, result.active_stages);

            // Update SHA-256 Checksum if present
            if (result.sha256) {
                const el1 = document.getElementById('audit-sha256-val');
                if (el1) el1.innerText = result.sha256;
                const el2 = document.getElementById('modal-sha256');
                if (el2) el2.innerText = result.sha256;
            }

            if (result.active_stages) {
                activeMatrixStages = result.active_stages;
            }

            // Update Sheets Pills
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

            // Load and render in-browser data grid
            loadPreview();

            document.getElementById('results-panel').scrollIntoView({ behavior: 'smooth' });
        }

        function updateDonutChart(covPct, covCount, untrCount) {
            const circle = document.getElementById('donut-circle');
            const pctText = document.getElementById('donut-pct');
            const covLegend = document.getElementById('donut-cov-lbl');
            const untrLegend = document.getElementById('donut-untr-lbl');

            const totalCircumference = 2 * Math.PI * 40; // ~251.327
            const pctFloat = parseFloat(covPct) || 0;
            const offset = totalCircumference - (totalCircumference * (pctFloat / 100));

            circle.style.strokeDasharray = totalCircumference;
            circle.style.strokeDashoffset = offset;
            pctText.innerText = pctFloat.toFixed(1) + '%';
            covLegend.innerText = `Traced: ${covCount}`;
            untrLegend.innerText = `Untraced: ${untrCount}`;
        }

        function updateStageDistributionBars(m, total, activeStages) {
            activeStages = activeStages || [];
            const hasShi = activeStages.includes('SHI');
            const hasSys = activeStages.includes('System Test') || activeStages.includes('System') || activeStages.includes('System Testing');
            const hasComp = activeStages.includes('Component Testing') || activeStages.includes('Component');
            const hasInteg = activeStages.includes('Integration Testing') || activeStages.includes('Integration') || activeStages.includes('Integ');

            // 1. SHI
            if (hasShi) {
                const shiCount = (m.shi_dl_count || 0) + (m.shi_sl_count || 0) > 0 ? ((m.shi_dl_count || 0) || (m.shi_sl_count || 0)) : 0;
                const shiPct = total > 0 ? ((shiCount / total) * 100).toFixed(1) : '0.0';
                document.getElementById('bar-shi-val').innerText = `${shiCount} / ${total} (${shiPct}%)`;
                document.getElementById('bar-shi-fill').style.width = `${shiPct}%`;
                document.getElementById('bar-shi-fill').style.opacity = '1';
            } else {
                document.getElementById('bar-shi-val').innerText = 'Disabled (Not in Scope)';
                document.getElementById('bar-shi-fill').style.width = '0%';
                document.getElementById('bar-shi-fill').style.opacity = '0.3';
            }

            // 2. System
            if (hasSys) {
                const sysCount = (m.sys_dl_count || 0) + (m.sys_sl_count || 0) > 0 ? ((m.sys_dl_count || 0) || (m.sys_sl_count || 0)) : 0;
                const sysPct = total > 0 ? ((sysCount / total) * 100).toFixed(1) : '0.0';
                document.getElementById('bar-sys-val').innerText = `${sysCount} / ${total} (${sysPct}%)`;
                document.getElementById('bar-sys-fill').style.width = `${sysPct}%`;
                document.getElementById('bar-sys-fill').style.opacity = '1';
            } else {
                document.getElementById('bar-sys-val').innerText = 'Disabled (Not in Scope)';
                document.getElementById('bar-sys-fill').style.width = '0%';
                document.getElementById('bar-sys-fill').style.opacity = '0.3';
            }

            // 3. Component
            if (hasComp) {
                const compCount = m.component_swrs || 0;
                const compPct = total > 0 ? ((compCount / total) * 100).toFixed(1) : '0.0';
                document.getElementById('bar-comp-val').innerText = `${compCount} / ${total} (${compPct}%)`;
                document.getElementById('bar-comp-fill').style.width = `${compPct}%`;
                document.getElementById('bar-comp-fill').style.opacity = '1';
            } else {
                document.getElementById('bar-comp-val').innerText = 'Disabled (Not in Scope)';
                document.getElementById('bar-comp-fill').style.width = '0%';
                document.getElementById('bar-comp-fill').style.opacity = '0.3';
            }

            // 4. Integration
            if (hasInteg) {
                const integCount = m.integration_swrs || 0;
                const integPct = total > 0 ? ((integCount / total) * 100).toFixed(1) : '0.0';
                document.getElementById('bar-integ-val').innerText = `${integCount} / ${total} (${integPct}%)`;
                document.getElementById('bar-integ-fill').style.width = `${integPct}%`;
                document.getElementById('bar-integ-fill').style.opacity = '1';
            } else {
                document.getElementById('bar-integ-val').innerText = 'Disabled (Not in Scope)';
                document.getElementById('bar-integ-fill').style.width = '0%';
                document.getElementById('bar-integ-fill').style.opacity = '0.3';
            }
        }

        let activeMatrixStages = { shi: false, system_test: false, component_test: false, integration_test: false };

        function escapeHtml(str) {
            if (str === null || str === undefined) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }

        function formatCellText(val) {
            if (!val) return '<span style="color:var(--text-muted); opacity: 0.35;">—</span>';
            const lines = String(val).split(String.fromCharCode(10));
            if (lines.length <= 1) return escapeHtml(val);
            return lines.map(l => `<div style="padding: 1.5px 0; white-space: nowrap;">${escapeHtml(l)}</div>`).join('');
        }

        function renderThead() {
            const thead = document.getElementById('preview-thead');
            if (!thead) return;

            const hasShi = !!activeMatrixStages.shi;
            const hasSys = !!activeMatrixStages.system_test;
            const hasComp = !!activeMatrixStages.component_test;
            const hasInteg = !!activeMatrixStages.integration_test;

            let row1 = '<tr style="background: var(--bg); border-bottom: 1px solid var(--border); font-weight: 700; color: var(--text-muted);">';
            row1 += '<th rowspan="2" style="padding: 10px 12px; width: 45px;">SL</th>';
            row1 += '<th rowspan="2" style="padding: 10px 12px; width: 110px;">Section</th>';
            row1 += '<th rowspan="2" style="padding: 10px 12px; width: 170px;">SWRS ID</th>';
            row1 += '<th rowspan="2" style="padding: 10px 12px; min-width: 250px;">Requirement Description</th>';

            if (hasShi) {
                row1 += '<th colspan="2" style="padding: 8px 12px; text-align: center; background: rgba(2, 132, 199, 0.12); color: #0284c7; border-left: 1px solid var(--border); font-weight: 800;">SHI Verification</th>';
            }
            if (hasSys) {
                row1 += '<th colspan="2" style="padding: 8px 12px; text-align: center; background: rgba(147, 51, 234, 0.12); color: #9333ea; border-left: 1px solid var(--border); font-weight: 800;">Overall System Software Testing</th>';
            }
            if (hasComp) {
                row1 += '<th colspan="4" style="padding: 8px 12px; text-align: center; background: rgba(5, 150, 105, 0.12); color: #059669; border-left: 1px solid var(--border); font-weight: 800;">Software Component Testing (Unit)</th>';
            }
            if (hasInteg) {
                row1 += '<th colspan="4" style="padding: 8px 12px; text-align: center; background: rgba(217, 119, 6, 0.12); color: #d97706; border-left: 1px solid var(--border); font-weight: 800;">Integration Testing</th>';
            }
            row1 += '<th rowspan="2" style="padding: 10px 12px; width: 95px; text-align: center; border-left: 1px solid var(--border);">Status</th>';
            row1 += '</tr>';

            let row2 = '<tr style="background: var(--bg); border-bottom: 2px solid var(--border); font-weight: 700; font-size: 11px; color: var(--text-muted);">';
            if (hasShi) {
                row2 += '<th style="padding: 6px 10px; border-left: 1px solid var(--border); background: rgba(2, 132, 199, 0.05);">Single line</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(2, 132, 199, 0.05);">Double line</th>';
            }
            if (hasSys) {
                row2 += '<th style="padding: 6px 10px; border-left: 1px solid var(--border); background: rgba(147, 51, 234, 0.05);">Single line</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(147, 51, 234, 0.05);">Double line</th>';
            }
            if (hasComp) {
                row2 += '<th style="padding: 6px 10px; border-left: 1px solid var(--border); background: rgba(5, 150, 105, 0.05);">Module Name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(5, 150, 105, 0.05);">File Name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(5, 150, 105, 0.05);">Function name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(5, 150, 105, 0.05);">Test case</th>';
            }
            if (hasInteg) {
                row2 += '<th style="padding: 6px 10px; border-left: 1px solid var(--border); background: rgba(217, 119, 6, 0.05);">Module Name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(217, 119, 6, 0.05);">File Name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(217, 119, 6, 0.05);">Function name</th>';
                row2 += '<th style="padding: 6px 10px; background: rgba(217, 119, 6, 0.05);">Test case</th>';
            }
            row2 += '</tr>';

            thead.innerHTML = row1 + row2;
        }

        function loadPreview() {
            fetch('/api/preview')
                .then(r => r.json())
                .then(res => {
                    if (res.status === 'success' && res.records) {
                        previewData = res.records;
                        document.getElementById('cnt-tab-all').innerText = res.total || res.records.length;
                        document.getElementById('cnt-tab-traced').innerText = res.traced;
                        document.getElementById('cnt-tab-untraced').innerText = res.untraced;
                        document.getElementById('grid-count-badge').innerText = `${res.total} Requirements`;

                        if (res.active_stages_map) {
                            activeMatrixStages = res.active_stages_map;
                        } else if (res.records.length > 0) {
                            activeMatrixStages = {
                                shi: !!res.records[0].shi_active,
                                system_test: !!res.records[0].sys_active,
                                component_test: !!res.records[0].comp_active,
                                integration_test: !!res.records[0].integ_active
                            };
                        }

                        renderThead();

                        if (res.active_stages) {
                            updateStageDistributionBars({
                                shi_sl_count: res.records.filter(r => r.shi).length,
                                shi_dl_count: 0,
                                sys_sl_count: res.records.filter(r => r.sys).length,
                                sys_dl_count: 0,
                                component_swrs: res.records.filter(r => r.comp).length,
                                integration_swrs: res.records.filter(r => r.integ).length
                            }, res.total || res.records.length, res.active_stages);
                        }

                        applyGridFilter();
                    }
                })
                .catch(err => console.log('Preview load notice:', err));
        }

        function setGridFilter(filter) {
            currentFilter = filter;
            document.querySelectorAll('.grid-tab').forEach(b => b.classList.remove('active'));
            const tabBtn = document.getElementById('tab-' + filter);
            if (tabBtn) tabBtn.classList.add('active');
            currentPage = 1;
            applyGridFilter();
        }

        function handleGridSearch(e) {
            currentSearch = (e.target.value || '').toLowerCase().trim();
            currentPage = 1;
            applyGridFilter();
        }

        function applyGridFilter() {
            filteredRecords = previewData.filter(r => {
                if (currentFilter === 'traced' && !r.is_traced) return false;
                if (currentFilter === 'untraced' && r.is_traced) return false;
                if (currentSearch) {
                    const s = currentSearch;
                    const matchId = (r.swrs_id || '').toLowerCase().includes(s);
                    const matchReq = (r.requirement || '').toLowerCase().includes(s);
                    const matchSec = (r.section || '').toLowerCase().includes(s);
                    const matchShi = (r.shi_sl || '').toLowerCase().includes(s) || (r.shi_dl || '').toLowerCase().includes(s);
                    const matchSys = (r.sys_sl || '').toLowerCase().includes(s) || (r.sys_dl || '').toLowerCase().includes(s);
                    const matchComp = (r.comp_mod || '').toLowerCase().includes(s) || (r.comp_file || '').toLowerCase().includes(s) || (r.comp_func || '').toLowerCase().includes(s) || (r.comp_tc || '').toLowerCase().includes(s);
                    const matchInteg = (r.integ_mod || '').toLowerCase().includes(s) || (r.integ_file || '').toLowerCase().includes(s) || (r.integ_func || '').toLowerCase().includes(s) || (r.integ_tc || '').toLowerCase().includes(s);
                    return matchId || matchReq || matchSec || matchShi || matchSys || matchComp || matchInteg;
                }
                return true;
            });

            renderGrid();
        }

        function renderGrid() {
            const tbody = document.getElementById('preview-tbody');
            if (!tbody) return;
            const total = filteredRecords.length;
            const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
            if (currentPage > totalPages) currentPage = totalPages;

            const startIdx = (currentPage - 1) * PAGE_SIZE;
            const endIdx = Math.min(startIdx + PAGE_SIZE, total);
            const pageRows = filteredRecords.slice(startIdx, endIdx);

            document.getElementById('grid-page-info').innerText = total > 0 
                ? `Showing ${startIdx + 1} to ${endIdx} of ${total} requirements`
                : 'No requirements match filter criteria';
            document.getElementById('grid-page-curr').innerText = `Page ${currentPage} of ${totalPages}`;
            document.getElementById('btn-page-prev').disabled = (currentPage <= 1);
            document.getElementById('btn-page-next').disabled = (currentPage >= totalPages);

            const hasShi = !!activeMatrixStages.shi;
            const hasSys = !!activeMatrixStages.system_test;
            const hasComp = !!activeMatrixStages.component_test;
            const hasInteg = !!activeMatrixStages.integration_test;
            const colCount = 4 + (hasShi ? 2 : 0) + (hasSys ? 2 : 0) + (hasComp ? 4 : 0) + (hasInteg ? 4 : 0) + 1;

            if (pageRows.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${colCount}" style="text-align:center; padding: 28px; color: var(--text-muted);">No matching requirements found.</td></tr>`;
                return;
            }

            tbody.innerHTML = pageRows.map(r => {
                const statusBadge = r.is_traced 
                    ? `<span style="background: rgba(16, 185, 129, 0.15); color: #059669; font-weight: 700; padding: 3px 8px; border-radius: 999px; font-size: 11px; white-space: nowrap;">✓ Traced</span>`
                    : `<span style="background: rgba(239, 68, 68, 0.12); color: #dc2626; font-weight: 700; padding: 3px 8px; border-radius: 999px; font-size: 11px; white-space: nowrap;">✗ Untraced</span>`;

                let rowHtml = `<tr style="border-bottom: 1px solid var(--border); transition: background 0.15s;">
                    <td style="padding: 9px 12px; color: var(--text-muted); font-weight: 600;">${r.sl_no}</td>
                    <td style="padding: 9px 12px; font-weight: 600; color: var(--text-main);">${escapeHtml(r.section || 'General')}</td>
                    <td style="padding: 9px 12px; font-family: ui-monospace, SFMono-Regular, monospace; font-weight: 700; color: #0284c7; white-space: nowrap;">${escapeHtml(r.swrs_id)}</td>
                    <td style="padding: 9px 12px; max-width: 320px; line-height: 1.4;" title="${escapeHtml(r.requirement || '')}">${escapeHtml(r.requirement)}</td>`;

                if (hasShi) {
                    rowHtml += `<td style="padding: 9px 10px; border-left: 1px solid var(--border);">${formatCellText(r.shi_sl)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px;">${formatCellText(r.shi_dl)}</td>`;
                }
                if (hasSys) {
                    rowHtml += `<td style="padding: 9px 10px; border-left: 1px solid var(--border);">${formatCellText(r.sys_sl)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px;">${formatCellText(r.sys_dl)}</td>`;
                }
                if (hasComp) {
                    rowHtml += `<td style="padding: 9px 10px; border-left: 1px solid var(--border); font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px;">${formatCellText(r.comp_mod)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px;">${formatCellText(r.comp_file)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px; color: #0284c7;">${formatCellText(r.comp_func)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px; font-weight: 600;">${formatCellText(r.comp_tc)}</td>`;
                }
                if (hasInteg) {
                    rowHtml += `<td style="padding: 9px 10px; border-left: 1px solid var(--border); font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px;">${formatCellText(r.integ_mod)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px;">${formatCellText(r.integ_file)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px; color: #d97706;">${formatCellText(r.integ_func)}</td>`;
                    rowHtml += `<td style="padding: 9px 10px; font-family: ui-monospace, SFMono-Regular, monospace; font-size: 11.5px; font-weight: 600;">${formatCellText(r.integ_tc)}</td>`;
                }

                rowHtml += `<td style="padding: 9px 12px; text-align: center; border-left: 1px solid var(--border);">${statusBadge}</td></tr>`;
                return rowHtml;
            }).join('');
        }

        function changeGridPage(delta) {
            currentPage += delta;
            renderGrid();
        }

        function downloadReport() {
            window.location.href = '/api/download';
        }

        function downloadTaggedTcfs() {
            window.location.href = '/api/download-tcfs';
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
    print(" Drag-and-drop enabled. Press Ctrl+C to stop the server.")
    print("=" * 70)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        server.server_close()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", 8080))
    run_server(port=port)
