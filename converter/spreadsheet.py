import io
import csv
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Tuple, Dict, Any, Optional

def col_name_to_index(col_name: str) -> int:
    """Converts Excel column name (e.g. 'A', 'Z', 'AA') to 0-based column index."""
    idx = 0
    for char in col_name.upper():
        if 'A' <= char <= 'Z':
            idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

def index_to_col_name(idx: int) -> str:
    """Converts 0-based column index to Excel column name (e.g. 0 -> 'A', 26 -> 'AA')."""
    name = ""
    idx += 1
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        name = chr(65 + remainder) + name
    return name

def split_cell_ref(ref: str) -> Tuple[str, int]:
    """Splits 'C14' into ('C', 14)."""
    col_str = ""
    row_str = ""
    for ch in ref:
        if ch.isalpha():
            col_str += ch
        elif ch.isdigit():
            row_str += ch
    row_num = int(row_str) if row_str else 1
    return col_str, row_num

def detect_csv_delimiter(sample_text: str) -> str:
    """Detects delimiter from CSV sample text, defaulting to comma."""
    if not sample_text:
        return ','
    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample_text[:4096], delimiters=[',', ';', '\t', '|'])
        return dialect.delimiter
    except Exception:
        first_line = sample_text.splitlines()[0] if sample_text.splitlines() else ""
        counts = {',': first_line.count(','), ';': first_line.count(';'), '\t': first_line.count('\t'), '|': first_line.count('|')}
        best = max(counts, key=counts.get)
        return best if counts[best] > 0 else ','

def csv_to_xlsx(csv_data: bytes | str, delimiter: Optional[str] = None) -> bytes:
    """Converts CSV bytes or string into a valid Microsoft Excel XLSX binary."""
    if isinstance(csv_data, bytes):
        # Attempt UTF-8, fall back to latin-1
        try:
            text = csv_data.decode('utf-8')
        except UnicodeDecodeError:
            text = csv_data.decode('latin-1')
    else:
        text = csv_data

    if not delimiter:
        delimiter = detect_csv_delimiter(text)

    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)

    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n',
        '  <sheetData>\n'
    ]

    for r_idx, row in enumerate(rows, start=1):
        if not row:
            continue
        sheet_lines.append(f'    <row r="{r_idx}">\n')
        for c_idx, cell_value in enumerate(row):
            col_letter = index_to_col_name(c_idx)
            cell_ref = f"{col_letter}{r_idx}"
            
            raw = cell_value.strip()
            # Determine if numeric
            is_num = False
            if raw:
                try:
                    # Do not treat leading zeros (like phone numbers '08123') as pure numbers
                    if raw.isdigit() and len(raw) > 1 and raw.startswith('0'):
                        is_num = False
                    else:
                        float(raw)
                        is_num = True
                except ValueError:
                    is_num = False

            escaped = (cell_value
                       .replace('&', '&amp;')
                       .replace('<', '&lt;')
                       .replace('>', '&gt;')
                       .replace('"', '&quot;'))

            if is_num:
                sheet_lines.append(f'      <c r="{cell_ref}"><v>{escaped}</v></c>\n')
            else:
                sheet_lines.append(f'      <c r="{cell_ref}" t="inlineStr"><is><t xml:space="preserve">{escaped}</t></is></c>\n')
        sheet_lines.append('    </row>\n')

    sheet_lines.append('  </sheetData>\n</worksheet>')
    sheet_xml = "".join(sheet_lines)

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
        '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        '</Types>'
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        '</Relationships>'
    )

    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n'
        '</Relationships>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        '  <sheets>\n'
        '    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>\n'
        '  </sheets>\n'
        '</workbook>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    return buf.getvalue()

def xlsx_to_csv(xlsx_data: bytes, delimiter: str = ',', sheet_index: int = 0) -> str:
    """Converts XLSX binary data to CSV string."""
    with zipfile.ZipFile(io.BytesIO(xlsx_data)) as z:
        namelist = z.namelist()
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        # 1. Parse shared strings if present
        shared_strings: List[str] = []
        if 'xl/sharedStrings.xml' in namelist:
            ss_tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in ss_tree.findall('.//s:si', ns):
                # Handle direct <t> or nested text runs <r><t>
                text_parts = []
                for t in si.findall('.//s:t', ns):
                    if t.text:
                        text_parts.append(t.text)
                shared_strings.append("".join(text_parts))

        # 2. Find target worksheet
        sheet_path = 'xl/worksheets/sheet1.xml'
        if 'xl/workbook.xml' in namelist:
            wb_tree = ET.fromstring(z.read('xl/workbook.xml'))
            sheets = wb_tree.findall('.//s:sheet', ns)
            if 0 <= sheet_index < len(sheets):
                # Get target sheet file
                sheet_elem = sheets[sheet_index]
                # Default path pattern
                sheet_path = f'xl/worksheets/sheet{sheet_index + 1}.xml'
                if sheet_path not in namelist:
                    # fallback to any available sheet in namelist
                    candidates = [n for n in namelist if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
                    if candidates:
                        sheet_path = sorted(candidates)[sheet_index if sheet_index < len(candidates) else 0]

        if sheet_path not in namelist:
            # Fallback to first xml in worksheets
            ws_files = [n for n in namelist if n.startswith('xl/worksheets/') and n.endswith('.xml')]
            if not ws_files:
                raise ValueError("No worksheets found in XLSX archive.")
            sheet_path = ws_files[0]

        sheet_tree = ET.fromstring(z.read(sheet_path))
        rows_map: Dict[int, Dict[int, str]] = {}
        max_col = 0

        for row_el in sheet_tree.findall('.//s:row', ns):
            r_val = row_el.get('r')
            row_idx = int(r_val) if r_val and r_val.isdigit() else (len(rows_map) + 1)
            rows_map[row_idx] = {}

            auto_col = 0
            for c_el in row_el.findall('./s:c', ns):
                c_ref = c_el.get('r')
                if c_ref:
                    col_str, _ = split_cell_ref(c_ref)
                    col_idx = col_name_to_index(col_str)
                else:
                    col_idx = auto_col
                auto_col = col_idx + 1

                if col_idx > max_col:
                    max_col = col_idx

                c_type = c_el.get('t', '')
                val_el = c_el.find('./s:v', ns)
                inline_el = c_el.find('.//s:is/s:t', ns)

                cell_value = ""
                if inline_el is not None and inline_el.text:
                    cell_value = inline_el.text
                elif val_el is not None and val_el.text is not None:
                    raw_val = val_el.text
                    if c_type == 's':
                        idx = int(raw_val)
                        cell_value = shared_strings[idx] if idx < len(shared_strings) else ""
                    elif c_type == 'b':
                        cell_value = "TRUE" if raw_val == "1" else "FALSE"
                    else:
                        cell_value = raw_val

                rows_map[row_idx][col_idx] = cell_value

        # Build tabular grid
        out = io.StringIO()
        writer = csv.writer(out, delimiter=delimiter)
        if rows_map:
            sorted_row_indices = sorted(rows_map.keys())
            max_row = max(sorted_row_indices)
            for r in range(1, max_row + 1):
                row_cells = rows_map.get(r, {})
                row_list = [row_cells.get(c, "") for c in range(0, max_col + 1)]
                writer.writerow(row_list)

        return out.getvalue()
