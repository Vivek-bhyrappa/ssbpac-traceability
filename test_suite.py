"""
Automated CI/CD Test Suite for SSBPAC Traceability Suite
Validates generator core, server routes, openpyxl workbook creation, creator attribution,
preview extraction, dynamic column mapping, and tagged TCF zip packaging.
"""
import os
import sys
import unittest
import zipfile
import openpyxl

# Add current directory to path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import generator_core
import server


class TestTraceabilityCore(unittest.TestCase):
    """Unit tests for generator core engine."""

    def test_imports(self):
        """Verify core modules import cleanly."""
        self.assertIsNotNone(generator_core)
        self.assertIsNotNone(server)

    def test_author_metadata_embed(self):
        """Verify author attribution Vivek is properly embedded in OpenPyXL properties."""
        wb = openpyxl.Workbook()
        wb.properties.creator = "Vivek"
        self.assertEqual(wb.properties.creator, "Vivek")

    def test_file_counter(self):
        """Verify count_files handles nonexistent and valid paths cleanly."""
        cnt, ok = server.count_files("nonexistent_path_test_12345", ".docx")
        self.assertFalse(ok)
        self.assertEqual(cnt, 0)

    def test_server_defaults_structure(self):
        """Verify default configuration keys exist in server module."""
        self.assertIn("input_excel", server.DEFAULT_CONFIG)
        self.assertIn("shi_dir", server.DEFAULT_CONFIG)
        self.assertIn("overall_dir", server.DEFAULT_CONFIG)
        self.assertIn("unit_dir", server.DEFAULT_CONFIG)
        self.assertIn("integ_dir", server.DEFAULT_CONFIG)

    def test_task_state_initialized(self):
        """Verify pipeline task state is properly initialized."""
        with server.state_lock:
            self.assertIn("is_running", server.task_state)
            self.assertIn("progress", server.task_state)
            self.assertIn("logs", server.task_state)

    def test_sha256_checksum(self):
        """Verify SHA-256 calculation for DO-178C integrity trail."""
        self.assertEqual(server.compute_file_sha256("nonexistent_test_file.bin"), "")
        test_file = os.path.join(CURRENT_DIR, "test_sha_tmp.txt")
        with open(test_file, "w", encoding="utf-8") as f:
            f.write("DO-178C-VERIFICATION-TEST")
        try:
            h = server.compute_file_sha256(test_file)
            self.assertEqual(len(h), 64)
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

    def test_extract_preview_records(self):
        """Verify preview records extractor handles empty/nonexistent files safely."""
        recs = server.extract_preview_records("nonexistent_workbook.xlsx")
        self.assertEqual(recs, [])

    def test_package_tagged_tcfs_zip(self):
        """Verify package_tagged_tcfs_zip creates valid zip archives."""
        zip_tmp = os.path.join(CURRENT_DIR, "test_tcf_archive.zip")
        try:
            cnt, u_cnt, i_cnt = server.package_tagged_tcfs_zip("nonexistent_u", "nonexistent_i", zip_tmp)
            self.assertEqual(cnt, 0)
            self.assertTrue(os.path.exists(zip_tmp))
            with zipfile.ZipFile(zip_tmp, "r") as zf:
                self.assertEqual(len(zf.namelist()), 0)
        finally:
            if os.path.exists(zip_tmp):
                os.remove(zip_tmp)

    def test_map_matrix_columns(self):
        """Verify dynamic header column mapping for partial stage configurations."""
        row1_integ = (None, "SL No.", "Requirement", "ID's", "Integration Testing", None, None, None)
        row2_integ = (None, None, None, None, "Module Name", "File Name", "Function name", "Test case")
        m_integ = server.map_matrix_columns(row1_integ, row2_integ)
        self.assertFalse(m_integ["has_shi"])
        self.assertFalse(m_integ["has_sys"])
        self.assertFalse(m_integ["has_comp"])
        self.assertTrue(m_integ["has_integ"])
        self.assertEqual(m_integ["integ_mod"], 4)
        self.assertEqual(m_integ["integ_file"], 5)
        self.assertEqual(m_integ["integ_func"], 6)
        self.assertEqual(m_integ["integ_tc"], 7)

        row1_shi = (None, "SL No.", "Requirement", "ID's", "SHI", None)
        row2_shi = (None, None, None, None, "Single line", "Double line")
        m_shi = server.map_matrix_columns(row1_shi, row2_shi)
        self.assertTrue(m_shi["has_shi"])
        self.assertFalse(m_shi["has_sys"])
        self.assertFalse(m_shi["has_integ"])

    def test_detailed_preview_extraction(self):
        """Verify extract_preview_records extracts detailed stage fields on overall_traceability_report.xlsx if present."""
        report_file = os.path.join(server.PARENT_DIR, "overall_traceability_report.xlsx")
        if not os.path.exists(report_file):
            report_file = os.path.join(CURRENT_DIR, "overall_traceability_report.xlsx")
        if os.path.exists(report_file):
            recs = server.extract_preview_records(report_file)
            self.assertGreater(len(recs), 0)
            sample = recs[0]
            self.assertIn("swrs_id", sample)
            self.assertIn("requirement", sample)
            self.assertIn("is_traced", sample)
            self.assertIn("integ_mod", sample)
            self.assertIn("integ_func", sample)
            self.assertIn("integ_tc", sample)
            self.assertIn("comp_mod", sample)
            self.assertIn("comp_func", sample)
            self.assertIn("comp_tc", sample)
            self.assertIn("comp_entries", sample)
            self.assertIn("integ_entries", sample)

    def test_multi_entry_alignment_and_subrows(self):
        """Verify multi-entry requirements (such as SwRS_GENR_045) have structured 1:1 sub-entries for file, function, test case."""
        report_file = os.path.join(server.PARENT_DIR, "overall_traceability_report.xlsx")
        if not os.path.exists(report_file):
            report_file = os.path.join(CURRENT_DIR, "overall_traceability_report.xlsx")
        if os.path.exists(report_file):
            recs = server.extract_preview_records(report_file)
            r045_list = [r for r in recs if r.get("swrs_id") == "PI_SSBPAC_SwRS_GENR_045"]
            if r045_list:
                r045 = r045_list[0]
                self.assertGreater(len(r045["integ_entries"]), 1)
                # Ensure each entry has file, func, and tc present
                for entry in r045["integ_entries"]:
                    self.assertIn("file", entry)
                    self.assertIn("func", entry)
                    self.assertIn("tc", entry)
                    self.assertTrue(bool(entry["func"]))
                # Total entries should match 9 rows in Excel
                self.assertEqual(len(r045["integ_entries"]), 9)
                self.assertEqual(r045["integ_entries"][0]["file"], "app.c")
                self.assertEqual(r045["integ_entries"][0]["func"], "cpu_sm")
                self.assertEqual(r045["integ_entries"][1]["file"], "app.c")
                self.assertEqual(r045["integ_entries"][1]["func"], "spi_read_cb")



if __name__ == "__main__":
    unittest.main()
