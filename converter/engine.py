import os
from typing import Tuple, Dict, Any, Optional
from converter.spreadsheet import csv_to_xlsx, xlsx_to_csv
from converter.document import pdf_to_docx, docx_to_pdf

SUPPORTED_CONVERSIONS = {
    ("xlsx", "csv"): ("text/csv", ".csv"),
    ("csv", "xlsx"): ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"),
    ("pdf", "docx"): ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
    ("docx", "pdf"): ("application/pdf", ".pdf"),
}

def get_base_filename(filename: str) -> str:
    """Extracts base filename without extension."""
    if not filename:
        return "converted"
    base = os.path.basename(filename)
    if "." in base:
        return base.rsplit(".", 1)[0]
    return base

def convert_file(file_bytes: bytes,
                 from_fmt: str,
                 to_fmt: str,
                 original_filename: str = "",
                 options: Optional[Dict[str, Any]] = None) -> Tuple[bytes, str, str]:
    """
    Executes conversion from one supported format to another.
    Returns: (output_bytes, output_filename, content_type)
    """
    src = from_fmt.lower().strip().lstrip(".")
    dst = to_fmt.lower().strip().lstrip(".")
    opts = options or {}

    pair = (src, dst)
    if pair not in SUPPORTED_CONVERSIONS:
        raise ValueError(f"Konversi dari '{src.upper()}' ke '{dst.upper()}' belum didukung. Pasangan yang didukung: XLSX ke CSV, CSV ke XLSX, PDF ke DOCX, dan DOCX ke PDF.")

    content_type, ext = SUPPORTED_CONVERSIONS[pair]
    base_name = get_base_filename(original_filename)
    out_filename = f"{base_name}{ext}"

    if pair == ("xlsx", "csv"):
        delimiter = opts.get("delimiter", ",")
        sheet_idx = int(opts.get("sheet_index", 0))
        csv_str = xlsx_to_csv(file_bytes, delimiter=delimiter, sheet_index=sheet_idx)
        out_bytes = csv_str.encode("utf-8")

    elif pair == ("csv", "xlsx"):
        delimiter = opts.get("delimiter")
        out_bytes = csv_to_xlsx(file_bytes, delimiter=delimiter)

    elif pair == ("pdf", "docx"):
        out_bytes = pdf_to_docx(file_bytes)

    elif pair == ("docx", "pdf"):
        out_bytes = docx_to_pdf(file_bytes)

    else:
        raise ValueError("Operasi tidak valid.")

    return out_bytes, out_filename, content_type
