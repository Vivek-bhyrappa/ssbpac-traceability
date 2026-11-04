#!/usr/bin/env python3
"""
================================================================================
Comprehensive SWRS End-to-End Traceability Report Generator
================================================================================
Consolidates and automates:
  1. TCF Test Case Tagging (add_test_case_numbers.py)
  2. DOCX Test Report Extraction (generate_single_and_double_line_report.py)
     - SHI (Software Hardware Integration): Single Line & Double Line
     - Overall System Software Testing: Single Line & Double Line
  3. TCF Test Suite Extraction (generate_traceability_report.py)
     - Software Component Testing (Unit Testing)
     - Integration Testing
     - Distinguishes based on SET_NAME (e.g. CPU0_unit_testing vs CPU123_Software_Integration)

Output:
  Excel workbook formatted with a two-tier green header matching the design specification:
  Row 1: [None, 'SL No.', 'Requirement', "ID's", 'SHI' (merged), 'Overall System Software Testing' (merged), 'Software Component Testing' (merged 4 cols), 'Integration Testing' (merged 4 cols)]
  Row 2: [None, '', '', '', 'Single line', 'Double line', 'Single line', 'Double line', 'Module Name', 'File Name', 'Function name', 'Test case', 'Module Name', 'File Name', 'Function name', 'Test case']
================================================================================
"""

import os
import sys
import re
import time
import shutil
import zipfile
import argparse
import logging
import ctypes
import hashlib
from collections import OrderedDict
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==============================================================================
# CONFIGURATION - CHANGE PATHS HERE
# ==============================================================================
# Set your absolute paths below. You can change any path to point to your selected inputs.
# Folders can also be specified as a list of absolute paths: [r"C:\path1", r"C:\path2"]

# 1. Unit Testing folder(s) containing .tcf files (Software Component Testing)

UNIT_TESTING_DIR = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\Unit testing"

# 2. Integration Testing folder(s) containing .tcf files
INTEGRATION_TESTING_DIR = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\Integration testing"

# 3. SHI Reports folder(s) containing Word (.docx) test reports
SHI_REPORTS_DIR = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\SHI"

# 4. Overall Software Test report folder(s) containing Word (.docx) reports
OVERALL_SOFTWARE_TEST_DIR = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\Overall Software Test report"

# 5. Base Excel file containing the SWRS Requirements template
REQUIREMENTS_EXCEL_PATH = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\overall_traceability.xlsx"

# 6. Output Excel report destination path
OUTPUT_EXCEL_PATH = r"C:\Users\hbt_l\OneDrive\Desktop\ssbpac\all-tracability\overall_traceability_report.xlsx"


# ------------------------------------------------------------------------------
# TCF TEST CASE TAGGING OPTIONS
# ------------------------------------------------------------------------------
# Set to True to automatically insert 'Test Case <N>' tags into .tcf files on disk
TAG_TCF_FILES = True

# Set to True to preview TCF tagging changes in the console without writing files
DRY_RUN_TAGGING = False

# Set to True to create a '.bak' backup before modifying any .tcf file
CREATE_BACKUP_ON_TAG = True

# ------------------------------------------------------------------------------
# HIGHLIGHTING & STYLING SETTINGS
# ------------------------------------------------------------------------------
# Highlight requirements that have no test coverage across any of the stages
HIGHLIGHT_UNTRACED_REQUIREMENTS = True

# Colors matching reference standard
UNTRACED_FILL_COLOR = "FFF2CC"   # Soft Light Yellow highlight for untested reqs
UNTRACED_FONT_COLOR = "000000"   # Clean, readable black text
HEADER_GREEN_FILL   = "92D050"   # Green accent header fill as shown in reference design

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("TraceabilityGenerator")


# ==============================================================================
# 1. FILE UTILITIES & SWRS NORMALIZATION
# ==============================================================================
def is_valid_hex_color(val):
    """Checks whether a value is a valid 6-char or 8-char aRGB hex string."""
    if not val:
        return False
    s = str(val).strip()
    return bool(re.match(r"^[0-9A-Fa-f]{6}$|^[0-9A-Fa-f]{8}$", s))


def safe_make_fill(color_val):
    """Safely constructs an openpyxl PatternFill only for valid hex colors."""
    if not color_val:
        return None
    s = str(color_val).strip()
    if s in ["00000000", "0", "None"]:
        return None
    if is_valid_hex_color(s):
        return PatternFill(start_color=s, end_color=s, fill_type="solid")
    return None


def to_path_list(item):
    """Normalizes a single path or list/tuple of paths into a flat list of strings."""
    if item is None:
        return []
    if isinstance(item, (list, tuple, set)):
        result = []
        for x in item:
            result.extend(to_path_list(x))
        return result
    s = str(item).strip()
    return [s] if s else []


def safe_copy_file(src, dst):
    """Safely copies files on Windows even if open in Word or Excel."""
    try:
        if os.name == "nt":
            res = ctypes.windll.kernel32.CopyFileW(src, dst, False)
            if res:
                return True
        shutil.copyfile(src, dst)
        return True
    except Exception as e:
        logger.warning(f"Error copying {src} -> {dst}: {e}")
        return False


def normalize_swrs_id(sid: str) -> str:
    """
    Standardizes requirement IDs across documents and TCF files:
      - Strips whitespace inside ID (e.g. 'GENR _0 40' -> 'GENR_040')
      - Distinguishes and preserves SWRS vs HWRS domain prefixes (e.g. SWRS_GENR_001 vs HWRS_GENR_001)
      - Formats numeric suffixes with 3-digit zero-padding (e.g. SWRS_POST_2 -> SWRS_POST_002)
      - Preserves architectural references (e.g. 'SAD')
    """
    if not sid:
        return ""
    s = re.sub(r"\s+", "", str(sid)).upper().strip()
    if s == "SAD":
        return "SAD"
    m = re.search(r"(HWRS|SWRS)[_\s]*([A-Z]{3,5})[_\s]*(\d+)", s)
    if m:
        rtype = m.group(1)
        cat = m.group(2)
        num = int(m.group(3))
        return f"{rtype}_{cat}_{num:03d}"
    m2 = re.search(r"([A-Z]{3,5})_(\d+)", s)
    if m2:
        return f"SWRS_{m2.group(1)}_{int(m2.group(2)):03d}"
    return s


def sort_tc_key(tc_str: str):
    """Sorts test case strings naturally by numeric value (e.g. Test Case 2 before Test Case 10)."""
    m = re.search(r"\d+", str(tc_str))
    if m:
        return (0, int(m.group()))
    return (1, str(tc_str))


def chunk_items(items, per_line=2, sep=", ", line_sep="\n") -> str:
    """Formats a list of strings into chunked multi-line text (e.g. 2 per line)."""
    if not items:
        return ""
    lines = []
    for i in range(0, len(items), per_line):
        lines.append(sep.join(items[i:i+per_line]))
    return line_sep.join(lines)


# ==============================================================================
# 2. TCF TEST CASE TAGGING (add_test_case_numbers.py)
# ==============================================================================
def tag_single_tcf_block(block_content: str, tc_num: int) -> tuple[str, bool]:
    """Tags a single test case block with 'Test Case <tc_num>' if missing."""
    end_tag = "# End Test Case"
    has_end_tag = end_tag in block_content
    if has_end_tag:
        clean_part, after_part = block_content.split(end_tag, 1)
        suffix = end_tag + after_part
    else:
        clean_part = block_content
        suffix = ""

    if re.search(r"\bTest Case\s+\d+\b", clean_part, re.IGNORECASE):
        return block_content, False

    nl = "\r\n" if "\r\n" in clean_part else "\n"

    # Case 1: # Begin Text already exists
    begin_text_m = re.search(r"([ \t]*# Begin Text[^\r\n]*(\r?\n))", clean_part)
    if begin_text_m:
        prefix = begin_text_m.group(1)
        insertion = prefix + f"Test Case {tc_num}" + nl
        new_clean = clean_part[:begin_text_m.start()] + insertion + clean_part[begin_text_m.end():]
        return new_clean + suffix, True

    # Case 2: Insert before # Begin Variable
    var_m = re.search(r"([ \t]*# Begin Variable\b)", clean_part)
    if var_m:
        text_block = f"      # Begin Text{nl}Test Case {tc_num}{nl}      # End Text{nl}{nl}"
        new_clean = clean_part[:var_m.start()] + text_block + clean_part[var_m.start():]
        return new_clean + suffix, True

    # Case 3: Fallback -> append at end
    text_block = f"{nl}      # Begin Text{nl}Test Case {tc_num}{nl}      # End Text{nl}"
    new_clean = clean_part.rstrip() + nl + text_block
    return new_clean + suffix, True


def tag_tcf_files_in_folder(folder_path, dry_run: bool = False, create_backup: bool = True):
    """Recursively scans and tags .tcf files in folder_path with explicit Test Case numbers.
    Skips duplicate files only if both the file name and content are identical."""
    folders = to_path_list(folder_path)
    seen_files = {}  # (filename_lower, content_hash) -> original_filepath

    for fld in folders:
        if not os.path.exists(fld):
            logger.warning(f"Folder not found for TCF tagging: {fld}")
            continue

        logger.info(f"Checking / Tagging TCF files in: {fld} (dry_run={dry_run})")
        total_files = 0
        modified_files = 0
        total_cases = 0
        modified_cases = 0

        for root, _, files in os.walk(fld):
            for f in sorted(files):
                if f.lower().endswith(".tcf"):
                    fp = os.path.join(root, f)
                    content = None
                    chosen_enc = "utf-8"
                    for enc in ["utf-8", "latin-1", "cp1252"]:
                        try:
                            with open(fp, "r", encoding=enc, errors="ignore") as file_obj:
                                content = file_obj.read()
                            chosen_enc = enc
                            break
                        except Exception:
                            continue
                    if content is None:
                        continue

                    # Skip duplicate files with identical filename and identical content
                    content_hash = hashlib.sha256(content.replace("\r\n", "\n").encode("utf-8")).hexdigest()
                    file_key = (f.lower(), content_hash)
                    if file_key in seen_files:
                        logger.info(f"Skipping duplicate file: {fp} (identical name & content to {seen_files[file_key]})")
                        continue
                    seen_files[file_key] = fp

                    total_files += 1

                    parts = content.split("# Begin Test Case")
                    if len(parts) <= 1:
                        continue

                    header = parts[0]
                    new_parts = [header]
                    file_modified = False

                    for tc_num, raw_block in enumerate(parts[1:], 1):
                        total_cases += 1
                        updated_block, mod = tag_single_tcf_block(raw_block, tc_num)
                        new_parts.append(updated_block)
                        if mod:
                            modified_cases += 1
                            file_modified = True

                    if file_modified:
                        modified_files += 1
                        new_content = "# Begin Test Case".join(new_parts)
                        if not dry_run:
                            if create_backup:
                                bak_path = fp + ".bak"
                                if not os.path.exists(bak_path):
                                    shutil.copy2(fp, bak_path)
                            with open(fp, "w", encoding=chosen_enc, newline="") as file_obj:
                                file_obj.write(new_content)

        mode_str = "[DRY-RUN]" if dry_run else "[TAGGED]"
        logger.info(f"{mode_str} {fld}: {modified_cases}/{total_cases} test cases updated in {modified_files}/{total_files} files.")


# ==============================================================================
# 3. DOCX TEST REPORT PARSER (SHI & OVERALL SOFTWARE TEST)
# ==============================================================================
def classify_docx(doc_path: str) -> tuple[str, str]:
    """
    Classifies a Word test report:
      1. Test Stage: 'SHI' vs 'SYSTEM_TEST'
      2. Configuration: 'SINGLE_LINE' vs 'DOUBLE_LINE'
    """
    W_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    TAG_T = f"{W_NAMESPACE}t"
    first_page_text = ""

    try:
        with zipfile.ZipFile(doc_path, "r") as z:
            with z.open("word/document.xml") as zf:
                tree = ET.parse(zf)
                root = tree.getroot()
                body = root.find(f"{W_NAMESPACE}body")
                paras = []
                if body is not None:
                    for elem in list(body)[:30]:
                        t = " ".join([n.text for n in elem.iter(TAG_T) if n.text]).strip()
                        if t:
                            paras.append(t)
                first_page_text = " ".join(paras).upper()
    except Exception as e:
        logger.warning(f"Error reading first page of {os.path.basename(doc_path)}: {e}")

    norm_text = re.sub(r"[\s_]+", " ", first_page_text)

    # 1. Test Stage
    if any(k in norm_text for k in ["HARDWARE INTEGRATION", "SHITR", "SHI TR", "SHI_TR"]):
        test_type = "SHI"
    elif any(k in norm_text for k in ["OVERALL SOFTWARE TEST", "OSTR", "SYSTEM TEST"]):
        test_type = "SYSTEM_TEST"
    else:
        p_lower = doc_path.lower()
        if "shi" in p_lower:
            test_type = "SHI"
        elif "overall" in p_lower or "system" in p_lower:
            test_type = "SYSTEM_TEST"
        else:
            test_type = "UNKNOWN"

    # 2. Configuration
    p_lower = doc_path.lower()
    if "single" in p_lower:
        config = "SINGLE_LINE"
    elif "double" in p_lower:
        config = "DOUBLE_LINE"
    elif "SINGLE LINE" in norm_text or "SINGLE_LINE" in norm_text or re.search(r"\bSL\b|_SL\b", norm_text):
        config = "SINGLE_LINE"
    elif "DOUBLE LINE" in norm_text or "DOUBLE_LINE" in norm_text or re.search(r"\bDL\b|_DL\b", norm_text):
        config = "DOUBLE_LINE"
    else:
        config = "UNKNOWN"

    return test_type, config


def extract_swrs_ids_from_text(text: str) -> set[str]:
    """Extracts and normalizes requirement IDs from arbitrary text blocks, distinguishing SWRS and HWRS."""
    ids = set()
    if not text:
        return ids
    # Token matching HWRS/SWRS + category code + digits with optional prefix
    pattern = re.compile(r"(?:PI_SSBPAC_|HASSDAC_)?(HWRS|SWRS)[_\s]*([A-Za-z]{3,5})[_\s]*(\d+)", re.IGNORECASE)
    for m in pattern.finditer(text):
        rtype = m.group(1).upper()
        cat = m.group(2).upper()
        num = int(m.group(3))
        ids.add(f"{rtype}_{cat}_{num:03d}")
    return ids


def parse_docx_traceability(doc_path: str) -> dict[str, set[str]]:
    """
    Parses a DOCX document via fast XML parsing.
    Returns: dict mapping norm_swrs_id -> set of formatted Test Case strings.
    """
    if not os.path.exists(doc_path):
        return {}

    t0 = time.time()
    swrs_to_tcs = {}

    try:
        with zipfile.ZipFile(doc_path, "r") as z:
            with z.open("word/document.xml") as zf:
                tree = ET.parse(zf)
                root = tree.getroot()

        W_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        TAG_P = f"{W_NAMESPACE}p"
        TAG_TBL = f"{W_NAMESPACE}tbl"
        TAG_TR = f"{W_NAMESPACE}tr"
        TAG_TC = f"{W_NAMESPACE}tc"
        TAG_T = f"{W_NAMESPACE}t"

        # 1. Parse Traceability Tables
        for tbl in root.iter(TAG_TBL):
            for tr in tbl.iter(TAG_TR):
                row_sids = set()
                row_tcs = set()
                for tc in tr.findall(TAG_TC):
                    t_raw = "".join([n.text for n in tc.iter(TAG_T) if n.text])
                    t_spaced = " ".join([n.text for n in tc.iter(TAG_T) if n.text])
                    for t in (t_raw, t_spaced):
                        row_sids.update(extract_swrs_ids_from_text(t))
                        tc_matches = re.findall(r"Test\s*Case\s*\d+", t, re.IGNORECASE)
                        for tc_str in tc_matches:
                            m = re.search(r"\d+", tc_str)
                            if m:
                                row_tcs.add(f"Test Case {int(m.group()):02d}")
                if row_sids and row_tcs:
                    for sid in row_sids:
                        swrs_to_tcs.setdefault(sid, set()).update(row_tcs)

        # 2. Parse Test Case Headings and execution paragraphs
        body = root.find(f"{W_NAMESPACE}body")
        if body is not None:
            curr_tc = ""
            for elem in body:
                if elem.tag == TAG_P:
                    txt = "".join([node.text for node in elem.iter(TAG_T) if node.text]).strip()
                    m = re.search(r"Test\s*Case\s*(\d+)", txt, re.IGNORECASE)
                    if m:
                        curr_tc = f"Test Case {int(m.group(1)):02d}"
                elif elem.tag == TAG_TBL:
                    txt_raw = "".join([node.text for node in elem.iter(TAG_T) if node.text])
                    txt_spaced = " ".join([node.text for node in elem.iter(TAG_T) if node.text])
                    for txt in (txt_raw, txt_spaced):
                        if curr_tc:
                            for sid in extract_swrs_ids_from_text(txt):
                                swrs_to_tcs.setdefault(sid, set()).add(curr_tc)

    except Exception as e:
        logger.error(f"Error parsing DOCX {doc_path}: {e}")

    elapsed = time.time() - t0
    logger.info(f"Parsed {os.path.basename(doc_path)} in {elapsed:.2f}s -> {len(swrs_to_tcs)} unique IDs mapped.")
    return swrs_to_tcs


def scan_all_docx_reports(shi_folders, overall_folders):
    """
    Scans SHI and Overall Software Test report folders or individual files.
    Accepts single path strings or lists of paths (files or directories).
    Returns:
      shi_sl_map: norm_id -> set of TCs
      shi_dl_map: norm_id -> set of TCs
      sys_sl_map: norm_id -> set of TCs
      sys_dl_map: norm_id -> set of TCs
    """
    shi_sl = {}
    shi_dl = {}
    sys_sl = {}
    sys_dl = {}

    targets = []
    for fld in to_path_list(shi_folders):
        targets.append((fld, "SHI"))
    for fld in to_path_list(overall_folders):
        targets.append((fld, "SYSTEM_TEST"))

    for target_path, fallback_type in targets:
        if not os.path.exists(target_path):
            logger.warning(f"DOCX path not found: {target_path}")
            continue

        if os.path.isfile(target_path):
            file_list = [target_path]
        else:
            file_list = []
            for root, _, files in os.walk(target_path):
                for f in sorted(files):
                    if f.lower().endswith(".docx") and not f.startswith("~$"):
                        file_list.append(os.path.join(root, f))

        for p in file_list:
            if p.lower().endswith(".docx") and not os.path.basename(p).startswith("~$"):
                tt, cfg = classify_docx(p)
                if tt == "UNKNOWN":
                    tt = fallback_type
                
                doc_map = parse_docx_traceability(p)
                for nid, tcs in doc_map.items():
                    if tt == "SHI":
                        if cfg == "SINGLE_LINE":
                            shi_sl.setdefault(nid, set()).update(tcs)
                        else:
                            shi_dl.setdefault(nid, set()).update(tcs)
                    else:  # SYSTEM_TEST
                        if cfg == "SINGLE_LINE":
                            sys_sl.setdefault(nid, set()).update(tcs)
                        else:
                            sys_dl.setdefault(nid, set()).update(tcs)

    return shi_sl, shi_dl, sys_sl, sys_dl


# ==============================================================================
# 4. TCF TEST CASE PARSER (UNIT & INTEGRATION TESTING)
# ==============================================================================
def extract_module_and_file(file_path_val: str, full_tcf_path: str, tcf_root: str) -> tuple[str, str]:
    """
    Extracts module_name (e.g. 'cpu_123' or 'cpu_0') and file_basename (e.g. 'app.c')
    from the TCF 'File =' path, matching the reference traceability design.
    """
    norm_path = (file_path_val + " " + full_tcf_path + " " + tcf_root).replace("\\", "/").lower()
    module_name = ""
    if "cpu_123" in norm_path or "cpu123" in norm_path:
        module_name = "cpu_123"
    elif "cpu_0" in norm_path or "cpu0" in norm_path:
        module_name = "cpu_0"
    else:
        parts = [pt for pt in file_path_val.replace("\\", "/").split("/") if pt]
        source_indices = [i for i, part in enumerate(parts) if part.lower() == "source"]
        if source_indices and source_indices[-1] > 0:
            candidate = parts[source_indices[-1] - 1]
            if candidate.lower() in ["code", "codes"] and source_indices[-1] > 1:
                module_name = parts[source_indices[-1] - 2]
            else:
                module_name = candidate

    file_basename = os.path.basename(file_path_val.replace("\\", "/")) if file_path_val else ""
    return module_name, file_basename


def parse_all_tcf_files(unit_folders, integ_folders, master_norm_ids: set[str] = None):
    """
    Recursively scans and parses .tcf files.
    Accepts single path strings or lists of folder paths.
    Distinguishes Software Component Testing vs Integration Testing:
      - If SET_NAME contains 'integration' (e.g. CPU123_Software_Integration) -> Integration Testing
      - If SET_NAME contains 'unit_testing', 'unit', or 'component' (e.g. CPU0_unit_testing) -> Component Testing
      - Fallback: folder path containing 'integration' or 'unit'.

    Returns:
      comp_map: norm_id -> list of dicts: [{"module": ..., "file": ..., "function": ..., "tc_text": ..., "tcs": [...]}, ...]
      integ_map: norm_id -> list of dicts: [{"module": ..., "file": ..., "function": ..., "tc_text": ..., "tcs": [...]}, ...]
      unmapped_tcs: list of unmapped test case details
    """
    raw_comp = {}   # norm_id -> dict of (mod, file, proc) -> set(tc_names)
    raw_integ = {}  # norm_id -> dict of (mod, file, proc) -> set(tc_names)
    unmapped_tcs = []
    seen_files = {}  # (filename_lower, content_hash) -> original_filepath

    folders = []
    for fld in to_path_list(unit_folders):
        folders.append((fld, "COMPONENT"))
    for fld in to_path_list(integ_folders):
        folders.append((fld, "INTEGRATION"))

    for folder_path, fallback_type in folders:
        if not os.path.exists(folder_path):
            logger.warning(f"TCF path not found: {folder_path}")
            continue

        if os.path.isfile(folder_path):
            tcf_files = [(os.path.dirname(folder_path), os.path.basename(folder_path))]
            base_folder = os.path.dirname(folder_path)
        else:
            tcf_files = []
            base_folder = folder_path
            logger.info(f"Scanning TCF directory: {folder_path}")
            for root, _, files in os.walk(folder_path):
                for f in sorted(files):
                    if f.lower().endswith(".tcf"):
                        tcf_files.append((root, f))

        for root, f in tcf_files:
            fp = os.path.join(root, f)
            try:
                rel_path = os.path.relpath(fp, base_folder)
            except ValueError:
                rel_path = fp

            content = None
            for enc in ["utf-8", "latin-1", "cp1252"]:
                try:
                    with open(fp, "r", encoding=enc, errors="ignore") as f_obj:
                        content = f_obj.read()
                    break
                except Exception:
                    continue
            if content is None:
                continue

            # Skip duplicate files only if both the file name and content are identical
            content_hash = hashlib.sha256(content.replace("\r\n", "\n").encode("utf-8")).hexdigest()
            file_key = (f.lower(), content_hash)
            if file_key in seen_files:
                logger.info(f"Skipping duplicate file: {fp} (identical name & content to {seen_files[file_key]})")
                continue
            seen_files[file_key] = fp

            # Extract SET_NAME
            set_m = re.search(r"^\s*SET_NAME\s*=\s*(.+)$", content, re.MULTILINE)
            set_name = set_m.group(1).strip() if set_m else ""

            # Determine test classification
            set_lower = set_name.lower()
            path_lower = fp.lower()
            if "integration" in set_lower or "integration" in path_lower:
                test_category = "INTEGRATION"
            elif "unit" in set_lower or "unit" in path_lower or "component" in set_lower or "component" in path_lower:
                test_category = "COMPONENT"
            else:
                test_category = fallback_type

            target_dict = raw_integ if test_category == "INTEGRATION" else raw_comp
            category_label = "Software Component Testing" if test_category == "COMPONENT" else "Integration Testing"

            blocks = content.split("# Begin Test Case")
            for b_idx, raw_block in enumerate(blocks[1:], 1):
                clean_block = raw_block.split("# End Test Case")[0] if "# End Test Case" in raw_block else raw_block

                proc_m = re.search(r"^\s*Procedure\s*=\s*(.+)$", clean_block, re.MULTILINE)
                proc_val = proc_m.group(1).strip() if proc_m else "UNKNOWN_PROC"

                desc_m = re.search(r"^\s*Description\s*=\s*(.+)$", clean_block, re.MULTILINE)
                desc_val = desc_m.group(1).strip() if desc_m else ""

                tc_m = re.search(r"\bTest Case\s+(\d+)\b", clean_block, re.IGNORECASE)
                tc_name = f"Test Case {int(tc_m.group(1))}" if tc_m else f"Test Case {b_idx}"

                file_m = re.search(r"^\s*File\s*=\s*(.+)$", clean_block, re.MULTILINE)
                file_path_val = file_m.group(1).strip() if file_m else ""
                mod_name, file_basename = extract_module_and_file(file_path_val, fp, folder_path)

                swrs_matches = re.findall(r"PI_SSBPAC_(?:SwRS|HwRS)_\w+", clean_block)
                swrs_uniq = list(dict.fromkeys(swrs_matches))

                if swrs_uniq:
                    for sid in swrs_uniq:
                        norm = normalize_swrs_id(sid)
                        key = (mod_name, file_basename, proc_val)
                        target_dict.setdefault(norm, {}).setdefault(key, set()).add(tc_name)

                        # If ID is tagged in TCF but not found in the SWRS requirements document:
                        if master_norm_ids is not None and norm not in master_norm_ids:
                            unmapped_tcs.append({
                                "file": rel_path,
                                "category": category_label,
                                "set_name": set_name,
                                "module": mod_name,
                                "file_name": file_basename,
                                "procedure": proc_val,
                                "tc_name": tc_name,
                                "swrs_id": sid,
                                "description": desc_val,
                                "status": f"({sid}) is not found in TCF or SWRS Excel document"
                            })
                else:
                    proc_id = f"{proc_val} - {tc_name}" if proc_val and proc_val != "UNKNOWN_PROC" else tc_name
                    unmapped_tcs.append({
                        "file": rel_path,
                        "category": category_label,
                        "set_name": set_name,
                        "module": mod_name,
                        "file_name": file_basename,
                        "procedure": proc_val,
                        "tc_name": tc_name,
                        "swrs_id": "Not Specified in TCF",
                        "description": desc_val,
                        "status": f"({proc_id}) is not found in TCF or SWRS Excel document"
                    })

    def finalize_tcf_map(raw_map: dict) -> dict[str, list[dict]]:
        final_map = {}
        for norm, func_dict in raw_map.items():
            entries = []
            for (mod, fil, fn) in sorted(func_dict.keys(), key=lambda x: (x[0], x[1], x[2])):
                tcs = sorted(list(func_dict[(mod, fil, fn)]), key=sort_tc_key)
                lines = []
                for i in range(0, len(tcs), 2):
                    lines.append(", ".join(tcs[i:i+2]))
                tc_text = ",\n".join(lines)
                entries.append({
                    "module": mod,
                    "file": fil,
                    "function": fn,
                    "tc_text": tc_text,
                    "tcs": tcs
                })
            final_map[norm] = entries
        return final_map

    comp_map = finalize_tcf_map(raw_comp)
    integ_map = finalize_tcf_map(raw_integ)

    # Collect all procedures that are successfully mapped to requirements across the test suite
    all_mapped_procedures = set()
    for norm, func_dict in list(raw_comp.items()) + list(raw_integ.items()):
        for (mod, fil, fn) in func_dict.keys():
            if fn:
                all_mapped_procedures.add(fn.strip().lower())

    filtered_unmapped_tcs = []
    for u in unmapped_tcs:
        proc_lower = u.get("procedure", "").strip().lower()
        rel_f_lower = u.get("file", "").lower()
        is_backup = "bckup" in rel_f_lower or "backup" in rel_f_lower
        
        # If the test case is from an older archive/backup file, but that procedure is already
        # fully mapped to requirements in an active test file, do not flag it as unmapped.
        if is_backup and proc_lower in all_mapped_procedures:
            continue
        filtered_unmapped_tcs.append(u)

    logger.info(
        f"TCF extraction complete. Component mapped: {len(comp_map)} SWRS, "
        f"Integration mapped: {len(integ_map)} SWRS, Unmapped TCs: {len(filtered_unmapped_tcs)} "
        f"(suppressed {len(unmapped_tcs) - len(filtered_unmapped_tcs)} legacy backup copies of already-mapped procedures)"
    )
    return comp_map, integ_map, filtered_unmapped_tcs


# ==============================================================================
# 5. REQUIREMENTS MASTER LIST EXTRACTOR
# ==============================================================================
def load_requirements_master_list(input_excel_path: str) -> list[dict]:
    """
    Loads all rows (sections, requirements, and sub-headers) from the base Excel template.
    Preserves exact row-by-row structure, SL No., requirement text, original ID's,
    and cell styles (fonts, fills, alignments, borders).
    Supports:
      1. Single-sheet format with section names in Col A (like overall_traceability.xlsx)
      2. Multi-sheet format where each sheet is a section (like SwRS Document Final.xlsx)
    """
    logger.info(f"Loading master requirements template from: {input_excel_path}")
    if not os.path.exists(input_excel_path):
        raise FileNotFoundError(f"Requirements master file not found: {input_excel_path}")

    if os.path.isdir(input_excel_path):
        candidates = []
        for root, _, files in os.walk(input_excel_path):
            for f in sorted(files):
                if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                    candidates.append(os.path.join(root, f))
        if not candidates:
            raise FileNotFoundError(f"No .xlsx file found in folder: {input_excel_path}")
        best = candidates[0]
        for c in candidates:
            c_name = os.path.basename(c).lower()
            if "overall" in c_name or "traceability" in c_name:
                best = c
                break
        logger.info(f"Resolved directory {input_excel_path} to requirements Excel: {best}")
        input_excel_path = best

    temp_excel = os.path.join(BASE_DIR, "temp_req_master.xlsx")
    safe_copy_file(input_excel_path, temp_excel)
    read_path = temp_excel if os.path.exists(temp_excel) else input_excel_path

    records = []
    try:
        wb = openpyxl.load_workbook(read_path, data_only=True)
        ws = wb.active

        # Check if first sheet has ID's or Requirement in row 1
        headers_row1 = [str(ws.cell(1, c).value or "") for c in range(1, 15)]
        has_overall_format = any("ID" in h.upper() or "REQ" in h.upper() for h in headers_row1)

        if has_overall_format:
            # Process every row from row 2 to max_row
            for r in range(2, ws.max_row + 1):
                c1_cell = ws.cell(r, 1)
                c2_cell = ws.cell(r, 2)
                c3_cell = ws.cell(r, 3)
                c4_cell = ws.cell(r, 4)

                c1_val = c1_cell.value
                c2_val = c2_cell.value
                c3_val = c3_cell.value
                c4_val = c4_cell.value

                def extract_style_dict(c):
                    fill_col = c.fill.fgColor.value if c.fill and c.fill.fgColor else None
                    if fill_col:
                        fill_col = str(fill_col)
                    return {
                        "fill_color": fill_col,
                        "font_bold": c.font.bold if c.font else False,
                        "font_size": c.font.size if c.font else 11,
                        "align_h": c.alignment.horizontal if c.alignment else None,
                        "align_v": c.alignment.vertical if c.alignment else None,
                    }

                styles = {
                    1: extract_style_dict(c1_cell),
                    2: extract_style_dict(c2_cell),
                    3: extract_style_dict(c3_cell),
                    4: extract_style_dict(c4_cell),
                }

                # Check if this row is a major section header
                is_section = bool(
                    c1_val and str(c1_val).strip() in [
                        "General", "POST", "INTERRUPTS", "ERRORS", "Operations",
                        "Reset_State", "SYS_Safety", "Communication"
                    ]
                )

                if is_section:
                    records.append({
                        "type": "section",
                        "row_idx": r,
                        "section_title": str(c1_val).strip(),
                        "c1": c1_val, "c2": c2_val, "c3": c3_val, "c4": c4_val,
                        "styles": styles
                    })
                elif c4_val and str(c4_val).strip():
                    sid_str = str(c4_val).strip()
                    records.append({
                        "type": "requirement",
                        "row_idx": r,
                        "sl_no": c2_val if c2_val is not None else "",
                        "requirement": str(c3_val).strip() if c3_val is not None else "",
                        "original_id": sid_str,
                        "norm_id": normalize_swrs_id(sid_str),
                        "c1": c1_val, "c2": c2_val, "c3": c3_val, "c4": c4_val,
                        "styles": styles
                    })
                else:
                    # Sub-header or row without ID
                    records.append({
                        "type": "sub_header",
                        "row_idx": r,
                        "sl_no": c2_val if c2_val is not None else "",
                        "requirement": str(c3_val).strip() if c3_val is not None else "",
                        "original_id": c4_val if c4_val is not None else "",
                        "norm_id": normalize_swrs_id(str(c4_val)) if c4_val else "",
                        "c1": c1_val, "c2": c2_val, "c3": c3_val, "c4": c4_val,
                        "styles": styles
                    })
        else:
            # Multi-sheet SWRS document format
            seen_ids = set()
            for sheet_name in wb.sheetnames:
                if sheet_name.strip().lower() in ["header", "metadata", "index"]:
                    continue
                sheet_ws = wb[sheet_name]
                sheet_has_id = False
                sheet_records = []
                for r in range(1, sheet_ws.max_row + 1):
                    row_vals = [str(sheet_ws.cell(r, c).value or "").strip() for c in range(1, sheet_ws.max_column + 1)]
                    row_text = " ".join(v for v in row_vals if v)
                    matches = re.findall(r"PI_SSBPAC_(?:SwRS|HwRS)_\w+", row_text)
                    for sid in matches:
                        norm = normalize_swrs_id(sid)
                        if norm not in seen_ids:
                            seen_ids.add(norm)
                            req_parts = [v for v in row_vals if v and sid not in v]
                            req_desc = " ".join(req_parts)
                            sheet_records.append({
                                "type": "requirement",
                                "row_idx": r,
                                "sl_no": len(sheet_records) + 1,
                                "requirement": req_desc,
                                "original_id": sid,
                                "norm_id": norm,
                                "c1": None, "c2": len(sheet_records) + 1, "c3": req_desc, "c4": sid,
                                "styles": {}
                            })
                            sheet_has_id = True
                if sheet_has_id:
                    records.append({
                        "type": "section",
                        "row_idx": 0,
                        "section_title": sheet_name,
                        "c1": sheet_name, "c2": None, "c3": None, "c4": None,
                        "styles": {}
                    })
                    records.extend(sheet_records)

    finally:
        if os.path.exists(temp_excel):
            try:
                os.remove(temp_excel)
            except Exception:
                pass

    req_count = sum(1 for x in records if x["type"] == "requirement")
    sec_count = sum(1 for x in records if x["type"] == "section")
    sub_count = sum(1 for x in records if x["type"] == "sub_header")
    logger.info(f"Loaded {req_count} requirements, {sec_count} section headers, and {sub_count} sub-headers ({len(records)} total rows).")
    return records


# ==============================================================================
# 6. EXCEL WORKBOOK GENERATOR (TWO-TIER GREEN HEADER MATCHING DESIGN)
# ==============================================================================
def build_overall_traceability_workbook(
    req_records: list[dict],
    shi_sl_map: dict[str, set[str]],
    shi_dl_map: dict[str, set[str]],
    sys_sl_map: dict[str, set[str]],
    sys_dl_map: dict[str, set[str]],
    comp_map: dict[str, list[dict]],
    integ_map: dict[str, list[dict]],
    unmapped_tcs: list[dict],
    output_path: str,
    include_shi: bool = True,
    include_sys: bool = True,
    include_comp: bool = True,
    include_integ: bool = True
):
    """
    Generates the master Excel report with:
      - Exact template styling, colors, and row layout from overall_traceability.xlsx
      - Dynamically generated headers and column layouts matching only the active test stages
      - Two-tier green header layout matching the attached reference image
      - Section header rows with gold accent (FFFFC000) matching template
      - All requirements populated with their corresponding test cases across active stages:
          * SHI (Single Line & Double Line)
          * Overall System Software Testing (Single Line & Double Line)
          * Software Component Testing (Module Name, File Name, Function name, Test case)
          * Integration Testing (Module Name, File Name, Function name, Test case)
      - Each function mapped to its own separate row and test cases
      - Untraced requirements highlighted in soft light yellow (#FFF2CC)
      - Auxiliary Summary and Unmapped Test Cases sheets
    """
    wb = openpyxl.Workbook()

    # --------------------------------------------------------------------------
    # Styles Definition matching overall_traceability.xlsx
    # --------------------------------------------------------------------------
    green_fill = PatternFill(start_color=HEADER_GREEN_FILL, end_color=HEADER_GREEN_FILL, fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    section_fill = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid") # Gold accent from template
    untraced_fill = PatternFill(start_color=UNTRACED_FILL_COLOR, end_color=UNTRACED_FILL_COLOR, fill_type="solid")

    header_font_14 = Font(name="Calibri", size=14, bold=True, color="000000")
    header_font_12 = Font(name="Calibri", size=12, bold=True, color="000000")
    header_font_11 = Font(name="Calibri", size=11, bold=True, color="000000")
    header_tier1_font = Font(name="Calibri", size=11, bold=True, color="000000")
    section_font = Font(name="Calibri", size=14, bold=True, color="000000")
    
    data_font = Font(name="Calibri", size=11)
    data_bold_font = Font(name="Calibri", size=11, bold=True)
    tc_font = Font(name="Calibri", size=10)
    untraced_font = Font(name="Calibri", size=10, color=UNTRACED_FONT_COLOR)

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # --------------------------------------------------------------------------
    # Define Dynamic Stage Blocks
    # --------------------------------------------------------------------------
    stage_blocks = []
    current_col = 5

    if include_shi:
        stage_blocks.append({
            "key": "SHI",
            "title": "SHI",
            "start_col": current_col,
            "end_col": current_col + 1,
            "subheaders": ["Single line", "Double line"],
            "widths": [20.0, 20.0],
            "types": ["tc", "tc"]
        })
        current_col += 2

    if include_sys:
        stage_blocks.append({
            "key": "SYS",
            "title": "Overall System Software Testing",
            "start_col": current_col,
            "end_col": current_col + 1,
            "subheaders": ["Single line", "Double line"],
            "widths": [20.0, 20.0],
            "types": ["tc", "tc"]
        })
        current_col += 2

    if include_comp:
        stage_blocks.append({
            "key": "COMP",
            "title": "Software Component Testing",
            "start_col": current_col,
            "end_col": current_col + 3,
            "subheaders": ["Module Name", "File Name", "Function name", "Test case"],
            "widths": [15.0, 18.0, 26.0, 30.0],
            "types": ["mod", "file", "func", "tc"]
        })
        current_col += 4

    if include_integ:
        stage_blocks.append({
            "key": "INTEG",
            "title": "Integration Testing",
            "start_col": current_col,
            "end_col": current_col + 3,
            "subheaders": ["Module Name", "File Name", "Function name", "Test case"],
            "widths": [15.0, 18.0, 26.0, 30.0],
            "types": ["mod", "file", "func", "tc"]
        })
        current_col += 4

    total_cols = current_col - 1

    # Set Workbook Properties & Author Watermark
    wb.properties.creator = "Vivek"
    wb.properties.lastModifiedBy = "Vivek"
    wb.properties.title = "SSBPAC SWRS Traceability Matrix"

    # --------------------------------------------------------------------------
    # SHEET 1: Master Traceability Matrix
    # --------------------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Traceability Matrix"
    ws1.views.sheetView[0].showGridLines = True

    # Row Heights matching template
    ws1.row_dimensions[1].height = 31.5
    ws1.row_dimensions[2].height = 24.0

    # Row 1 Base Values
    ws1.cell(1, 1, value=None)
    ws1.cell(1, 2, value="SL No.")
    ws1.cell(1, 3, value="Requirement")
    ws1.cell(1, 4, value="ID's")

    # Row 2 Base Values
    ws1.cell(2, 1, value=None)
    ws1.cell(2, 2, value=None)
    ws1.cell(2, 3, value=None)
    ws1.cell(2, 4, value=None)

    # Base Header Merging
    ws1.merge_cells("B1:B2")   # SL No.
    ws1.merge_cells("C1:C2")   # Requirement
    ws1.merge_cells("D1:D2")   # ID's

    # Stage Headers in Row 1 & Row 2
    for block in stage_blocks:
        sc = block["start_col"]
        ec = block["end_col"]
        ws1.cell(1, sc, value=block["title"])
        for c in range(sc + 1, ec + 1):
            ws1.cell(1, c, value=None)
        if sc < ec:
            ws1.merge_cells(start_row=1, start_column=sc, end_row=1, end_column=ec)

        for c_idx, subh in zip(range(sc, ec + 1), block["subheaders"]):
            ws1.cell(2, c_idx, value=subh)

    # Header Styling
    for col in range(1, total_cols + 1):
        c1 = ws1.cell(1, col)
        c2 = ws1.cell(2, col)
        c1.border = thin_border
        c2.border = thin_border
        c1.alignment = center_align
        c2.alignment = center_align
        c2.font = header_font_11

        if col == 1:
            pass
        elif col == 2:
            c1.font = header_font_11
            c1.fill = white_fill
            c2.fill = white_fill
        elif col in [3, 4]:
            c1.font = header_font_14
            c1.fill = green_fill
            c2.fill = green_fill
        else:
            c1.fill = green_fill
            c2.fill = green_fill
            # Set font based on stage size
            is_large = False
            for b in stage_blocks:
                if b["start_col"] <= col <= b["end_col"]:
                    if (b["end_col"] - b["start_col"] + 1) >= 4:
                        is_large = True
                    break
            c1.font = header_font_12 if is_large else header_font_14

    # Column Widths
    ws1.column_dimensions["A"].width = 14.71
    ws1.column_dimensions["B"].width = 9.14
    ws1.column_dimensions["C"].width = 120.0
    ws1.column_dimensions["D"].width = 26.85

    for block in stage_blocks:
        for c_idx, w in zip(range(block["start_col"], block["end_col"] + 1), block["widths"]):
            ws1.column_dimensions[get_column_letter(c_idx)].width = w

    # Populate Data Rows
    current_out_row = 3
    traced_stats = {
        "total": 0,
        "traced": 0,
        "untraced": 0,
        "shi_sl": 0,
        "shi_dl": 0,
        "sys_sl": 0,
        "sys_dl": 0,
        "component": 0,
        "integration": 0
    }

    for item in req_records:
        if item["type"] == "section":
            # Section header banner row matching template (FFFFC000 fill)
            sec_title = item["section_title"]
            ws1.cell(current_out_row, 1, value=sec_title)
            for c in range(1, total_cols + 1):
                cell = ws1.cell(current_out_row, c)
                cell.fill = section_fill
                cell.font = section_font
                cell.border = thin_border
            ws1.row_dimensions[current_out_row].height = 20.0
            current_out_row += 1

        elif item["type"] == "sub_header":
            # Sub-header or title row without ID (e.g. INTERLOCK, SM Panel, WatchDog)
            sl_no = item.get("sl_no", "")
            req_text = item.get("requirement", "")
            orig_id = item.get("original_id", "")
            styles = item.get("styles", {})

            row_vals = [
                item.get("c1"),
                sl_no,
                req_text,
                orig_id
            ] + [None] * (total_cols - 4)

            for c_idx, val in enumerate(row_vals, 1):
                cell = ws1.cell(current_out_row, c_idx, value=val)
                cell.border = thin_border
                cell.font = data_bold_font if not orig_id else data_font
                if c_idx in [2, 4]:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

                style_c = styles.get(c_idx, {}) if c_idx <= 4 else {}
                fill_color = style_c.get("fill_color")
                fill_obj = safe_make_fill(fill_color)
                if fill_obj:
                    cell.fill = fill_obj

            current_out_row += 1

        elif item["type"] == "requirement":
            sl_no = item["sl_no"]
            req_text = item["requirement"]
            orig_id = item["original_id"]
            nid = item["norm_id"]
            styles = item.get("styles", {})

            traced_stats["total"] += 1

            # 1. SHI Single & Double Line
            shi_sl_tcs = sorted(list(shi_sl_map.get(nid, set())), key=sort_tc_key) if include_shi else []
            shi_dl_tcs = sorted(list(shi_dl_map.get(nid, set())), key=sort_tc_key) if include_shi else []
            if shi_sl_tcs: traced_stats["shi_sl"] += 1
            if shi_dl_tcs: traced_stats["shi_dl"] += 1

            # 2. System Testing Single & Double Line
            sys_sl_tcs = sorted(list(sys_sl_map.get(nid, set())), key=sort_tc_key) if include_sys else []
            sys_dl_tcs = sorted(list(sys_dl_map.get(nid, set())), key=sort_tc_key) if include_sys else []
            if sys_sl_tcs: traced_stats["sys_sl"] += 1
            if sys_dl_tcs: traced_stats["sys_dl"] += 1

            # 3. Software Component Testing
            comp_entries = comp_map.get(nid, []) if include_comp else []
            if comp_entries: traced_stats["component"] += 1

            # 4. Integration Testing
            integ_entries = integ_map.get(nid, []) if include_integ else []
            if integ_entries: traced_stats["integration"] += 1

            # Check if requirement has coverage in active stages
            has_coverage = bool(
                shi_sl_tcs or shi_dl_tcs or sys_sl_tcs or sys_dl_tcs or comp_entries or integ_entries
            )
            if has_coverage:
                traced_stats["traced"] += 1
            else:
                traced_stats["untraced"] += 1

            n_rows = max(1, len(comp_entries) if include_comp else 0, len(integ_entries) if include_integ else 0)
            start_row = current_out_row
            end_row = current_out_row + n_rows - 1

            for k in range(n_rows):
                r = start_row + k

                # Base columns 1 to 4
                row_vals = [
                    item.get("c1") if k == 0 else None,
                    sl_no if k == 0 else None,
                    req_text if k == 0 else None,
                    orig_id if k == 0 else None
                ]

                # Append columns for each active stage block
                for block in stage_blocks:
                    bkey = block["key"]
                    if bkey == "SHI":
                        row_vals.append(chunk_items(shi_sl_tcs, per_line=2) if k == 0 else None)
                        row_vals.append(chunk_items(shi_dl_tcs, per_line=2) if k == 0 else None)
                    elif bkey == "SYS":
                        row_vals.append(chunk_items(sys_sl_tcs, per_line=2) if k == 0 else None)
                        row_vals.append(chunk_items(sys_dl_tcs, per_line=2) if k == 0 else None)
                    elif bkey == "COMP":
                        if k < len(comp_entries):
                            ce = comp_entries[k]
                            row_vals.extend([ce["module"], ce["file"], ce["function"], ce["tc_text"]])
                        else:
                            row_vals.extend(["", "", "", ""])
                    elif bkey == "INTEG":
                        if k < len(integ_entries):
                            ie = integ_entries[k]
                            row_vals.extend([ie["module"], ie["file"], ie["function"], ie["tc_text"]])
                        else:
                            row_vals.extend(["", "", "", ""])

                for c_idx, val in enumerate(row_vals, 1):
                    cell = ws1.cell(r, c_idx, value=val)
                    cell.border = thin_border

                    if c_idx <= 4:
                        cell.font = data_font
                        if c_idx in [2, 4]:
                            cell.alignment = center_align
                        else:
                            cell.alignment = left_align

                        # Preserve template fill on Col 4 if present
                        style_c = styles.get(c_idx, {}) if c_idx <= 4 else {}
                        fill_color = style_c.get("fill_color")
                        fill_obj = safe_make_fill(fill_color)
                        if fill_obj:
                            cell.fill = fill_obj

                    else:
                        # Find stage column type
                        col_type = "tc"
                        for b in stage_blocks:
                            if b["start_col"] <= c_idx <= b["end_col"]:
                                offset = c_idx - b["start_col"]
                                col_type = b["types"][offset]
                                break

                        if col_type in ["mod", "file", "func"]:
                            cell.font = data_font
                            cell.alignment = center_align
                        else:
                            cell.font = tc_font
                            cell.alignment = left_align

                        if HIGHLIGHT_UNTRACED_REQUIREMENTS and (not has_coverage) and (total_cols > 4):
                            cell.fill = untraced_fill
                            cell.font = untraced_font

            # Merge vertical spans across n_rows
            if n_rows > 1:
                if item.get("c1"):
                    ws1.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
                ws1.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
                ws1.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)
                ws1.merge_cells(start_row=start_row, start_column=4, end_row=end_row, end_column=4)

                for block in stage_blocks:
                    bkey = block["key"]
                    sc = block["start_col"]
                    ec = block["end_col"]
                    if bkey in ["SHI", "SYS"]:
                        for c_col in range(sc, ec + 1):
                            ws1.merge_cells(start_row=start_row, start_column=c_col, end_row=end_row, end_column=c_col)
                    elif bkey == "COMP":
                        if len(comp_entries) <= 1:
                            for c_col in range(sc, ec + 1):
                                ws1.merge_cells(start_row=start_row, start_column=c_col, end_row=end_row, end_column=c_col)
                    elif bkey == "INTEG":
                        if len(integ_entries) <= 1:
                            for c_col in range(sc, ec + 1):
                                ws1.merge_cells(start_row=start_row, start_column=c_col, end_row=end_row, end_column=c_col)

            current_out_row += n_rows

    # Freeze panes at E3 if additional columns exist, else C3
    ws1.freeze_panes = "E3" if total_cols > 4 else "C3"

    # --------------------------------------------------------------------------
    # SHEET 2: Traceability Statistics & Summary
    # --------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Summary Statistics")
    ws2.views.sheetView[0].showGridLines = True
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 18

    ws2.cell(1, 1, value="Metric / Category").font = header_tier1_font
    ws2.cell(1, 2, value="Requirement Count").font = header_tier1_font
    ws2.cell(1, 3, value="Coverage (%)").font = header_tier1_font
    for c in range(1, 4):
        ws2.cell(1, c).fill = green_fill
        ws2.cell(1, c).border = thin_border
        ws2.cell(1, c).alignment = center_align

    tot = traced_stats["total"] or 1
    metrics = [
        ("Total SWRS Requirements", traced_stats["total"], "100.0%"),
    ]
    if total_cols > 4:
        metrics.append(("Covered by At Least One Test Stage", traced_stats["traced"], f"{(traced_stats['traced']/tot)*100:.1f}%"))
        metrics.append(("Completely Untraced Requirements", traced_stats["untraced"], f"{(traced_stats['untraced']/tot)*100:.1f}%"))

    if include_shi:
        metrics.append(("Traced in SHI (Single Line)", traced_stats["shi_sl"], f"{(traced_stats['shi_sl']/tot)*100:.1f}%"))
        metrics.append(("Traced in SHI (Double Line)", traced_stats["shi_dl"], f"{(traced_stats['shi_dl']/tot)*100:.1f}%"))

    if include_sys:
        metrics.append(("Traced in Overall System Test (Single Line)", traced_stats["sys_sl"], f"{(traced_stats['sys_sl']/tot)*100:.1f}%"))
        metrics.append(("Traced in Overall System Test (Double Line)", traced_stats["sys_dl"], f"{(traced_stats['sys_dl']/tot)*100:.1f}%"))

    if include_comp:
        metrics.append(("Traced in Software Component Testing", traced_stats["component"], f"{(traced_stats['component']/tot)*100:.1f}%"))

    if include_integ:
        metrics.append(("Traced in Integration Testing", traced_stats["integration"], f"{(traced_stats['integration']/tot)*100:.1f}%"))

    # Identify unmapped SWRS requirements for the active stages
    unmapped_swrs = []
    if total_cols > 4:
        for item in req_records:
            if item.get("type") == "requirement":
                nid = item.get("norm_id")
                orig_id = item.get("original_id") or item.get("norm_id")
                cov = False
                if include_shi and (shi_sl_map.get(nid) or shi_dl_map.get(nid)): cov = True
                if include_sys and (sys_sl_map.get(nid) or sys_dl_map.get(nid)): cov = True
                if include_comp and comp_map.get(nid): cov = True
                if include_integ and integ_map.get(nid): cov = True

                if not cov:
                    if include_comp or include_integ:
                        status_str = f"({orig_id}) is not found in TCF or SWRS Excel document"
                    elif include_shi and not include_sys:
                        status_str = f"({orig_id}) is not found in SHI test reports"
                    elif include_sys and not include_shi:
                        status_str = f"({orig_id}) is not found in System Test reports"
                    else:
                        status_str = f"({orig_id}) has no coverage in provided test reports"

                    unmapped_swrs.append({
                        "sl_no": item.get("sl_no", ""),
                        "swrs_id": orig_id,
                        "norm_id": nid,
                        "requirement": item.get("requirement", ""),
                        "status": status_str
                    })

    if include_comp or include_integ:
        unmapped_tcf_swrs_count = sum(
            1 for item in req_records
            if item.get("type") == "requirement" and not (comp_map.get(item.get("norm_id")) or integ_map.get(item.get("norm_id")))
        )
        metrics.append(("SWRS Reqs Not Covered in Component/Integ TCF", unmapped_tcf_swrs_count, f"{(unmapped_tcf_swrs_count/tot)*100:.1f}%"))
        metrics.append(("Unmapped TCF Test Cases (No SWRS Tag in TCF)", len(unmapped_tcs), "-"))

    for r_idx, (m_label, m_val, m_pct) in enumerate(metrics, start=2):
        c1 = ws2.cell(r_idx, 1, value=m_label)
        c2 = ws2.cell(r_idx, 2, value=m_val)
        c3 = ws2.cell(r_idx, 3, value=m_pct)
        c1.border = thin_border; c2.border = thin_border; c3.border = thin_border
        c1.font = data_bold_font if "Total" in m_label or "At Least One" in m_label else data_font
        c2.font = data_bold_font if "Total" in m_label or "At Least One" in m_label else data_font
        c3.font = data_bold_font if "Total" in m_label or "At Least One" in m_label else data_font
        c2.alignment = center_align
        c3.alignment = center_align

    # --------------------------------------------------------------------------
    # SHEET 3: Unmapped Test Cases (only when TCFs are provided)
    # --------------------------------------------------------------------------
    if include_comp or include_integ:
        ws3 = wb.create_sheet(title="Unmapped Test Cases")
        ws3.views.sheetView[0].showGridLines = True
        headers_ws3 = [
            "TCF File", "Stage", "Module Name", "File Name",
            "Procedure / Function", "Test Case", "SWRS ID", "Description", "Status"
        ]
        for c_idx, h in enumerate(headers_ws3, 1):
            cell = ws3.cell(1, c_idx, value=h)
            cell.fill = green_fill
            cell.font = header_tier1_font
            cell.border = thin_border
            cell.alignment = center_align

        ws3_widths = [32, 26, 16, 18, 26, 16, 32, 50, 55]
        for c_idx, w in enumerate(ws3_widths, 1):
            ws3.column_dimensions[get_column_letter(c_idx)].width = w

        current_ws3_row = 2
        for tc in unmapped_tcs:
            proc = tc.get("procedure", "")
            tc_num = tc.get("tc_name", "")
            fallback_id = f"{proc} - {tc_num}" if proc and proc != "UNKNOWN_PROC" else (tc_num or "Test Case")
            default_status = f"({fallback_id}) is not found in TCF or SWRS Excel document"
            row_vals = [
                tc["file"],
                tc["category"],
                tc.get("module", ""),
                tc.get("file_name", ""),
                tc["procedure"],
                tc["tc_name"],
                tc.get("swrs_id", "Not Specified in TCF"),
                tc["description"],
                tc.get("status", default_status)
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws3.cell(current_ws3_row, c_idx, value=val)
                cell.border = thin_border
                cell.font = data_font
                if c_idx in [2, 3, 4, 6]:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align
            current_ws3_row += 1

    # --------------------------------------------------------------------------
    # SHEET 4: Unmapped SWRS Requirements (when test stages were provided)
    # --------------------------------------------------------------------------
    if total_cols > 4:
        ws4 = wb.create_sheet(title="Unmapped SWRS Requirements")
        ws4.views.sheetView[0].showGridLines = True
        headers_ws4 = ["SL No.", "SWRS ID", "Requirement Description", "Status"]
        for c_idx, h in enumerate(headers_ws4, 1):
            cell = ws4.cell(1, c_idx, value=h)
            cell.fill = green_fill
            cell.font = header_tier1_font
            cell.border = thin_border
            cell.alignment = center_align

        ws4_widths = [12, 32, 80, 55]
        for c_idx, w in enumerate(ws4_widths, 1):
            ws4.column_dimensions[get_column_letter(c_idx)].width = w

        for r_idx, sw in enumerate(unmapped_swrs, start=2):
            row_vals = [
                sw["sl_no"],
                sw["swrs_id"],
                sw["requirement"],
                sw["status"]
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws4.cell(r_idx, c_idx, value=val)
                cell.border = thin_border
                cell.font = data_font
                if c_idx in [1, 2]:
                    cell.alignment = center_align
                else:
                    cell.alignment = left_align

    # --------------------------------------------------------------------------
    # Save Output
    # --------------------------------------------------------------------------
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    target_paths = [output_path] + [output_path.replace(".xlsx", f"_updated_{i}.xlsx" if i > 1 else "_updated.xlsx") for i in range(1, 10)]
    for cand in target_paths:
        try:
            wb.save(cand)
            if cand != output_path:
                logger.warning(f"Target '{output_path}' was open in Excel. Saved to alternate path: {cand}")
            else:
                logger.info(f"Successfully saved complete traceability report to: {cand}")
            return cand
        except PermissionError:
            continue
    raise PermissionError(f"Could not save workbook: all candidate file paths were locked by Excel.")


# ==============================================================================
# 7. MAIN CLI & EXECUTION PIPELINE
# ==============================================================================
def run_traceability_pipeline(
    unit_dir=None,
    integ_dir=None,
    shi_dir=None,
    overall_dir=None,
    input_excel=None,
    output_excel=None,
    tag_tcfs=False,
    dry_run=False,
    create_backup=True,
    progress_callback=None,
    include_shi=None,
    include_sys=None,
    include_comp=None,
    include_integ=None
):
    """
    Executes the complete SWRS traceability generation pipeline.
    Invokes progress_callback(msg, pct) if provided.
    Returns a dict with summary metrics and saved output file path.
    """
    def log_progress(msg, pct):
        logger.info(f"[{pct}%] {msg}")
        if progress_callback:
            try:
                progress_callback(msg, pct)
            except Exception:
                pass

    unit_paths = to_path_list(unit_dir)
    integ_paths = to_path_list(integ_dir)
    shi_paths = to_path_list(shi_dir)
    overall_paths = to_path_list(overall_dir)

    # Auto-detect stages if not explicitly specified
    if include_shi is None:
        include_shi = len(shi_paths) > 0
    if include_sys is None:
        include_sys = len(overall_paths) > 0
    if include_comp is None:
        include_comp = len(unit_paths) > 0
    if include_integ is None:
        include_integ = len(integ_paths) > 0

    input_excel = input_excel if input_excel else REQUIREMENTS_EXCEL_PATH
    output_excel = output_excel if output_excel else OUTPUT_EXCEL_PATH

    log_progress("Initializing generation pipeline...", 5)

    active_stages = []
    if include_shi: active_stages.append("SHI")
    if include_sys: active_stages.append("System Test")
    if include_comp: active_stages.append("Component Testing")
    if include_integ: active_stages.append("Integration Testing")
    mode_str = " + ".join(active_stages) if active_stages else "SWRS Template Only"
    log_progress(f"Configured Mode: SWRS vs {mode_str}", 8)

    # Step 1: Optional TCF Tagging
    if tag_tcfs or dry_run:
        if include_comp:
            log_progress("Tagging Unit Testing TCF files with test case numbers...", 10)
            tag_tcf_files_in_folder(unit_paths, dry_run=dry_run, create_backup=create_backup)
        if include_integ:
            log_progress("Tagging Integration Testing TCF files with test case numbers...", 20)
            tag_tcf_files_in_folder(integ_paths, dry_run=dry_run, create_backup=create_backup)

    # Step 2: Load Requirements Master List
    log_progress("Reading master requirements template...", 30)
    req_records = load_requirements_master_list(input_excel)
    master_norm_ids = {r["norm_id"] for r in req_records if r.get("norm_id")}

    # Step 3: Scan DOCX Reports
    shi_sl, shi_dl, sys_sl, sys_dl = {}, {}, {}, {}
    if include_shi or include_sys:
        log_progress("Extracting SHI & Overall System Test Word reports...", 50)
        s_shi = shi_paths if include_shi else []
        s_sys = overall_paths if include_sys else []
        shi_sl, shi_dl, sys_sl, sys_dl = scan_all_docx_reports(s_shi, s_sys)

    # Step 4: Scan TCF Test Files
    comp_map, integ_map, unmapped_tcs = {}, {}, []
    if include_comp or include_integ:
        log_progress("Parsing Unit & Integration TCF test cases...", 70)
        u_paths = unit_paths if include_comp else []
        i_paths = integ_paths if include_integ else []
        comp_map, integ_map, unmapped_tcs = parse_all_tcf_files(u_paths, i_paths, master_norm_ids=master_norm_ids)

    # Step 5: Build Formatted Excel Output
    log_progress("Generating formatted Excel report...", 85)
    saved_path = build_overall_traceability_workbook(
        req_records=req_records,
        shi_sl_map=shi_sl,
        shi_dl_map=shi_dl,
        sys_sl_map=sys_sl,
        sys_dl_map=sys_dl,
        comp_map=comp_map,
        integ_map=integ_map,
        unmapped_tcs=unmapped_tcs,
        output_path=output_excel,
        include_shi=include_shi,
        include_sys=include_sys,
        include_comp=include_comp,
        include_integ=include_integ
    )

    log_progress(f"Traceability Report successfully saved to: {os.path.basename(saved_path)}", 100)

    # Compute high-level stats
    all_covered = set()
    for m in [shi_sl, shi_dl, sys_sl, sys_dl, comp_map, integ_map]:
        all_covered.update(m.keys())
    traced_master = all_covered.intersection(master_norm_ids)

    generated_sheets = ["Traceability Matrix", "Summary Statistics"]
    if include_comp or include_integ:
        generated_sheets.append("Unmapped Test Cases")
    generated_sheets.append("Unmapped SWRS Requirements")

    return {
        "status": "success",
        "output_file": saved_path,
        "filename": os.path.basename(saved_path),
        "mode": f"SWRS vs {mode_str}",
        "active_stages": active_stages,
        "sheets": generated_sheets,
        "metrics": {
            "total_requirements": len(master_norm_ids),
            "covered_requirements": len(traced_master),
            "coverage_pct": f"{(len(traced_master) / len(master_norm_ids) * 100):.1f}%" if master_norm_ids else "0%",
            "untraced_requirements": len(master_norm_ids) - len(traced_master),
            "shi_sl_count": len(shi_sl),
            "shi_dl_count": len(shi_dl),
            "sys_sl_count": len(sys_sl),
            "sys_dl_count": len(sys_dl),
            "component_swrs": len(comp_map),
            "integration_swrs": len(integ_map),
            "unmapped_tcs": len(unmapped_tcs)
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Unified All-In-One SWRS Traceability Report Generator"
    )
    parser.add_argument("--unit-dir", nargs="*", default=None, help="Unit testing TCF folder(s)")
    parser.add_argument("--integ-dir", nargs="*", default=None, help="Integration testing TCF folder(s)")
    parser.add_argument("--shi-dir", nargs="*", default=None, help="SHI DOCX folder(s)")
    parser.add_argument("--overall-dir", nargs="*", default=None, help="Overall Software Test report DOCX folder(s)")
    parser.add_argument("--input-excel", default=None, help="Input base Excel with SWRS requirements")
    parser.add_argument("--output", default=None, help="Output Excel report path")
    parser.add_argument("--tag-tcfs", action="store_true", help="Write explicit Test Case numbers into .tcf files")
    parser.add_argument("--dry-run-tag", action="store_true", help="Preview TCF tagging without writing to files")
    parser.add_argument("--no-backup", action="store_true", help="Skip creating .bak files when tagging TCFs")
    args = parser.parse_args()

    # Use CLI args if explicitly provided, otherwise fallback to top-level CONFIGURATION variables
    unit_dir = args.unit_dir if args.unit_dir is not None else UNIT_TESTING_DIR
    integ_dir = args.integ_dir if args.integ_dir is not None else INTEGRATION_TESTING_DIR
    shi_dir = args.shi_dir if args.shi_dir is not None else SHI_REPORTS_DIR
    overall_dir = args.overall_dir if args.overall_dir is not None else OVERALL_SOFTWARE_TEST_DIR
    input_excel = args.input_excel if args.input_excel is not None else REQUIREMENTS_EXCEL_PATH
    output_excel = args.output if args.output is not None else OUTPUT_EXCEL_PATH

    tag_tcfs = args.tag_tcfs or TAG_TCF_FILES
    dry_run = args.dry_run_tag or DRY_RUN_TAGGING
    create_backup = (not args.no_backup) if args.no_backup else CREATE_BACKUP_ON_TAG

    print("=" * 80)
    print(" SWRS TRACEABILITY: ALL-IN-ONE AUTOMATED REPORT GENERATOR")
    print("=" * 80)
    print(f"Base Directory       : {BASE_DIR}")
    print(f"Unit Testing (TCF)   : {unit_dir}")
    print(f"Integration (TCF)    : {integ_dir}")
    print(f"SHI Reports (DOCX)   : {shi_dir}")
    print(f"Overall Reports(DOCX): {overall_dir}")
    print(f"Requirements Master  : {input_excel}")
    print(f"Output Report File   : {output_excel}")
    print(f"Auto-Tag TCFs        : {tag_tcfs} (dry_run={dry_run})")
    print("-" * 80)

    # Step 1: Optional TCF Tagging
    if tag_tcfs or dry_run:
        tag_tcf_files_in_folder(unit_dir, dry_run=dry_run, create_backup=create_backup)
        tag_tcf_files_in_folder(integ_dir, dry_run=dry_run, create_backup=create_backup)

    # Step 2: Load Requirements Master List
    req_records = load_requirements_master_list(input_excel)
    master_norm_ids = {r["norm_id"] for r in req_records if r.get("norm_id")}

    # Step 3: Scan DOCX Reports (SHI & System Test)
    shi_sl, shi_dl, sys_sl, sys_dl = scan_all_docx_reports(shi_dir, overall_dir)

    # Step 4: Scan TCF Test Files (Component & Integration)
    comp_map, integ_map, unmapped_tcs = parse_all_tcf_files(unit_dir, integ_dir, master_norm_ids=master_norm_ids)

    # Step 5: Build Formatted Excel Output
    saved_path = build_overall_traceability_workbook(
        req_records=req_records,
        shi_sl_map=shi_sl,
        shi_dl_map=shi_dl,
        sys_sl_map=sys_sl,
        sys_dl_map=sys_dl,
        comp_map=comp_map,
        integ_map=integ_map,
        unmapped_tcs=unmapped_tcs,
        output_path=output_excel
    )

    print("\n" + "=" * 80)
    print(f"SUCCESS: Master Traceability Report generated at:\n  {saved_path}")
    print("=" * 80)


if __name__ == "__main__":
    main()

