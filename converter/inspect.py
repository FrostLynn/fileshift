import io
import csv
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from typing import Dict, Any, List
from converter.spreadsheet import xlsx_to_csv, detect_csv_delimiter

def inspect_file(file_bytes: bytes, filename: str, format_hint: str = "") -> Dict[str, Any]:
    """Inspects a file and returns format metadata and preview structure."""
    fmt = format_hint.lower()
    if not fmt:
        ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        fmt = ext

    result: Dict[str, Any] = {
        "filename": filename,
        "format": fmt,
        "size_bytes": len(file_bytes),
        "preview_type": "text",
        "metadata": {},
        "preview_data": None
    }

    try:
        if fmt == "csv":
            try:
                text = file_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = file_bytes.decode("latin-1")
            
            delim = detect_csv_delimiter(text)
            reader = csv.reader(io.StringIO(text), delimiter=delim)
            rows = []
            for idx, r in enumerate(reader):
                if idx < 15:
                    rows.append(r)
                if idx >= 100:
                    break
            
            line_count = len(text.splitlines())
            result["preview_type"] = "table"
            result["metadata"] = {
                "delimiter": delim,
                "total_rows": line_count,
                "columns_count": len(rows[0]) if rows else 0
            }
            result["preview_data"] = {
                "headers": rows[0] if rows else [],
                "rows": rows[1:11] if len(rows) > 1 else rows[:1]
            }

        elif fmt == "xlsx":
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                sheet_names = []
                if "xl/workbook.xml" in z.namelist():
                    wb = ET.fromstring(z.read("xl/workbook.xml"))
                    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                    for s in wb.findall(".//s:sheet", ns):
                        name = s.get("name", "")
                        if name:
                            sheet_names.append(name)
                
            csv_text = xlsx_to_csv(file_bytes, delimiter=",", sheet_index=0)
            reader = csv.reader(io.StringIO(csv_text))
            rows = []
            for idx, r in enumerate(reader):
                if idx < 15:
                    rows.append(r)
                if idx >= 100:
                    break

            result["preview_type"] = "table"
            result["metadata"] = {
                "sheets": sheet_names or ["Sheet1"],
                "total_rows": len(csv_text.splitlines()),
                "columns_count": len(rows[0]) if rows else 0
            }
            result["preview_data"] = {
                "headers": rows[0] if rows else [],
                "rows": rows[1:11] if len(rows) > 1 else rows[:1]
            }

        elif fmt == "pdf":
            # Extract plain text
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                tf.write(file_bytes)
                tf_path = tf.name
            try:
                # Get page count and text via pdftotext
                proc = subprocess.run(["pdftotext", tf_path, "-"], capture_output=True, text=True, check=False)
                raw_text = proc.stdout if proc.returncode == 0 else ""
                
                # Count pages via form feed \x0c
                page_count = max(1, raw_text.count("\x0c"))
                words = raw_text.split()
                snippet = " ".join(words[:120]) + ("..." if len(words) > 120 else "")
                
                result["preview_type"] = "document"
                result["metadata"] = {
                    "pages": page_count,
                    "words_count": len(words),
                    "character_count": len(raw_text)
                }
                result["preview_data"] = {
                    "snippet": snippet,
                    "paragraphs": [p.strip() for p in raw_text.split("\n\n")[:5] if p.strip()]
                }
            finally:
                if os.path.exists(tf_path):
                    os.remove(tf_path)

        elif fmt == "docx":
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
                xml_content = z.read("word/document.xml")
            tree = ET.fromstring(xml_content)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            
            paragraphs = []
            for p in tree.findall(".//w:p", ns):
                texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
                if texts:
                    paragraphs.append("".join(texts).strip())
            
            all_text = " ".join(paragraphs)
            words = all_text.split()
            snippet = " ".join(words[:120]) + ("..." if len(words) > 120 else "")

            result["preview_type"] = "document"
            result["metadata"] = {
                "paragraphs_count": len(paragraphs),
                "words_count": len(words),
                "character_count": len(all_text)
            }
            result["preview_data"] = {
                "snippet": snippet,
                "paragraphs": paragraphs[:6]
            }

    except Exception as err:
        result["error"] = str(err)

    return result
