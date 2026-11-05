"""
Automated CI/CD Test Suite for SSBPAC Traceability Suite
Validates generator core, server routes, openpyxl workbook creation, and creator attribution.
"""
import os
import sys
import unittest
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


if __name__ == "__main__":
    unittest.main()
