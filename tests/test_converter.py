import unittest
import os
import zipfile
import xml.etree.ElementTree as ET
from converter.engine import convert_file
from converter.inspect import inspect_file

class TestConverterEngine(unittest.TestCase):
    def test_csv_to_xlsx_and_back(self):
        csv_data = "Item,Qty,Price\nLaptop,5,1200\nMouse,25,25.5\nKeyboard,10,75.00\n"
        # CSV to XLSX
        xlsx_bytes, xlsx_name, xlsx_mime = convert_file(
            csv_data.encode("utf-8"), "csv", "xlsx", original_filename="inventory.csv"
        )
        self.assertEqual(xlsx_name, "inventory.xlsx")
        self.assertIn("spreadsheetml", xlsx_mime)
        self.assertTrue(len(xlsx_bytes) > 500)

        # Inspect XLSX
        insp = inspect_file(xlsx_bytes, "inventory.xlsx")
        self.assertEqual(insp["format"], "xlsx")
        self.assertEqual(insp["preview_type"], "table")
        self.assertIn("Item", insp["preview_data"]["headers"])

        # XLSX to CSV
        csv_bytes, csv_name, csv_mime = convert_file(
            xlsx_bytes, "xlsx", "csv", original_filename="inventory.xlsx"
        )
        self.assertEqual(csv_name, "inventory.csv")
        self.assertEqual(csv_mime, "text/csv")
        reconstructed = csv_bytes.decode("utf-8")
        self.assertIn("Laptop", reconstructed)
        self.assertIn("Keyboard", reconstructed)

    def test_docx_and_pdf_roundtrip(self):
        # Create a test DOCX
        from converter.document import create_docx_package
        paras = [
            {"text": "Project Report 2026", "bold": True, "heading_level": 1},
            {"text": "This document contains a summary of system architecture.", "bold": False},
            {"text": "The file conversion is executed locally and safely.", "bold": False}
        ]
        sample_docx = create_docx_package(paras)

        # Inspect DOCX
        insp_docx = inspect_file(sample_docx, "report.docx")
        self.assertEqual(insp_docx["format"], "docx")
        self.assertEqual(insp_docx["preview_type"], "document")
        self.assertIn("Project Report", insp_docx["preview_data"]["snippet"])

        # DOCX to PDF
        pdf_bytes, pdf_name, pdf_mime = convert_file(
            sample_docx, "docx", "pdf", original_filename="report.docx"
        )
        self.assertEqual(pdf_name, "report.pdf")
        self.assertEqual(pdf_mime, "application/pdf")
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))

        # Inspect PDF
        insp_pdf = inspect_file(pdf_bytes, "report.pdf")
        self.assertEqual(insp_pdf["format"], "pdf")
        self.assertEqual(insp_pdf["preview_type"], "document")
        self.assertIn("Project Report", insp_pdf["preview_data"]["snippet"])

        # PDF to DOCX
        re_docx, re_docx_name, re_docx_mime = convert_file(
            pdf_bytes, "pdf", "docx", original_filename="report.pdf"
        )
        self.assertEqual(re_docx_name, "report.docx")
        self.assertIn("wordprocessingml", re_docx_mime)
        self.assertTrue(len(re_docx) > 500)

        # Check reconstructed DOCX content
        insp_re = inspect_file(re_docx, "report.docx")
        self.assertIn("Project Report", insp_re["preview_data"]["snippet"])

if __name__ == "__main__":
    unittest.main()
