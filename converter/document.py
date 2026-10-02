import io
import os
import re
import tempfile
import zipfile
import subprocess
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional

def create_docx_package(paragraphs: List[Dict[str, Any]], tables: Optional[List[List[List[str]]]] = None) -> bytes:
    """Builds a standard OpenXML DOCX archive from structured paragraphs and optional tables."""
    doc_body = []
    
    for p in paragraphs:
        if p.get("is_page_break"):
            doc_body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n')
            continue

        text = p.get("text", "")
        # XML Escape text
        escaped_text = (text.replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;"))
        
        bold_tag = "<w:b/>" if p.get("bold") else ""
        italic_tag = "<w:i/>" if p.get("italic") else ""
        
        # Heading styles or font size
        heading_level = p.get("heading_level", 0)
        p_style = ""
        size = p.get("size", 24) # 24 half-pts = 12pt
        
        if heading_level == 1:
            p_style = '<w:pStyle w:val="Heading1"/>'
            size = max(size, 36) # 18pt
            bold_tag = "<w:b/>"
        elif heading_level == 2:
            p_style = '<w:pStyle w:val="Heading2"/>'
            size = max(size, 28) # 14pt
            bold_tag = "<w:b/>"
        elif heading_level == 3:
            p_style = '<w:pStyle w:val="Heading3"/>'
            size = max(size, 24)
            bold_tag = "<w:b/>"
            
        sz_tag = f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        
        p_pr = f'<w:pPr>{p_style}</w:pPr>' if p_style else ""
        r_pr = f'<w:rPr>{bold_tag}{italic_tag}{sz_tag}</w:rPr>' if (bold_tag or italic_tag or sz_tag) else ""
        
        doc_body.append(f"""    <w:p>
      {p_pr}
      <w:r>
        {r_pr}
        <w:t xml:space="preserve">{escaped_text}</w:t>
      </w:r>
    </w:p>\n""")

    # Render optional tables
    if tables:
        for tbl in tables:
            tbl_lines = [
                '    <w:tbl>\n',
                '      <w:tblPr>\n',
                '        <w:tblBorders>\n',
                '          <w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '          <w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '          <w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>\n',
                '        </w:tblBorders>\n',
                '      </w:tblPr>\n'
            ]
            for row in tbl:
                tbl_lines.append('      <w:tr>\n')
                for cell in row:
                    c_esc = (cell.replace("&", "&amp;")
                                 .replace("<", "&lt;")
                                 .replace(">", "&gt;"))
                    tbl_lines.append(f"""        <w:tc>
          <w:p><w:r><w:t xml:space="preserve">{c_esc}</w:t></w:r></w:p>
        </w:tc>\n""")
                tbl_lines.append('      </w:tr>\n')
            tbl_lines.append('    </w:tbl>\n')
            doc_body.append("".join(tbl_lines))

    # Section properties (A4 with 1 inch / 1440 twips margins)
    doc_body.append("""    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>\n""")

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
{"".join(doc_body)}  </w:body>
</w:document>"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
      <w:color w:val="2F5496"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
      <w:color w:val="2F5496"/>
    </w:rPr>
  </w:style>
</w:styles>"""

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles_xml)
        z.writestr("word/document.xml", document_xml)

    return buf.getvalue()

def pdf_to_docx(pdf_data: bytes) -> bytes:
    """Converts PDF bytes into a formatted DOCX binary."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        temp_pdf.write(pdf_data)
        temp_pdf_path = temp_pdf.name

    try:
        # Run pdftotext with bbox-layout
        res = subprocess.run(["pdftotext", "-bbox-layout", temp_pdf_path, "-"],
                             capture_output=True, text=True, check=False)
        paragraphs: List[Dict[str, Any]] = []

        if res.returncode == 0 and "<html" in res.stdout:
            try:
                root = ET.fromstring(res.stdout)
                ns = {"h": "http://www.w3.org/1999/xhtml"}
                
                # Check for pages
                pages = root.findall(".//h:page", ns)
                if not pages:
                    # Try without namespace
                    pages = root.findall(".//page")

                for page_idx, page in enumerate(pages):
                    if page_idx > 0:
                        paragraphs.append({"is_page_break": True})
                    
                    blocks = page.findall(".//h:block", ns) or page.findall(".//block")
                    for block in blocks:
                        lines = block.findall("./h:line", ns) or block.findall("./line")
                        block_text_lines = []
                        font_heights = []
                        
                        for line in lines:
                            words = line.findall("./h:word", ns) or line.findall("./word")
                            w_texts = []
                            for w in words:
                                if w.text:
                                    w_texts.append(w.text)
                                try:
                                    y_min = float(w.get("yMin", 0))
                                    y_max = float(w.get("yMax", 0))
                                    if y_max > y_min:
                                        font_heights.append(y_max - y_min)
                                except Exception:
                                    pass
                            if w_texts:
                                block_text_lines.append(" ".join(w_texts))
                                
                        if block_text_lines:
                            combined_text = " ".join(block_text_lines).strip()
                            if not combined_text:
                                continue
                            avg_height = sum(font_heights) / len(font_heights) if font_heights else 10.0
                            
                            is_h1 = avg_height >= 15.0 or (avg_height >= 13.0 and len(combined_text) < 60)
                            is_h2 = (not is_h1) and (avg_height >= 12.0 and len(combined_text) < 90)
                            
                            h_lvl = 1 if is_h1 else (2 if is_h2 else 0)
                            paragraphs.append({
                                "text": combined_text,
                                "bold": is_h1 or is_h2,
                                "heading_level": h_lvl,
                                "size": 36 if is_h1 else (28 if is_h2 else 24)
                            })
            except Exception:
                paragraphs = []

        # Fallback to plain pdftotext if bbox-layout was empty
        if not paragraphs:
            plain_res = subprocess.run(["pdftotext", "-layout", temp_pdf_path, "-"],
                                       capture_output=True, text=True, check=False)
            text_content = plain_res.stdout if plain_res.returncode == 0 else ""
            raw_paras = [p.strip() for p in text_content.split("\n\n") if p.strip()]
            for p in raw_paras:
                is_title = len(p) < 60 and not p.endswith(".")
                paragraphs.append({
                    "text": p.replace("\n", " "),
                    "bold": is_title,
                    "heading_level": 1 if is_title else 0,
                    "size": 32 if is_title else 24
                })

        return create_docx_package(paragraphs)
    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

def docx_to_pdf(docx_data: bytes) -> bytes:
    """Converts DOCX binary data into a standard PDF binary."""
    with zipfile.ZipFile(io.BytesIO(docx_data)) as z:
        if "word/document.xml" not in z.namelist():
            raise ValueError("Invalid DOCX file: word/document.xml not found.")
        tree = ET.fromstring(z.read("word/document.xml"))
        
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    elements: List[Dict[str, Any]] = []

    # Iterate over child elements of w:body to preserve document order
    body = tree.find(".//w:body", ns)
    if body is not None:
        for child in body:
            tag = child.tag
            if tag.endswith("p"):
                # Paragraph
                p_text_parts = []
                is_bold = False
                is_italic = False
                is_heading = False
                
                # Check style
                p_pr = child.find("./w:pPr", ns)
                if p_pr is not None:
                    p_style = p_pr.find("./w:pStyle", ns)
                    if p_style is not None:
                        val = p_style.get(f"{{{ns['w']}}}val", "").lower()
                        if "heading" in val or "title" in val:
                            is_heading = True

                for r in child.findall("./w:r", ns):
                    r_pr = r.find("./w:rPr", ns)
                    if r_pr is not None:
                        if r_pr.find("./w:b", ns) is not None:
                            is_bold = True
                        if r_pr.find("./w:i", ns) is not None:
                            is_italic = True
                    for t in r.findall("./w:t", ns):
                        if t.text:
                            p_text_parts.append(t.text)
                
                full_text = "".join(p_text_parts).strip()
                if full_text:
                    elements.append({
                        "type": "paragraph",
                        "text": full_text,
                        "bold": is_bold or is_heading,
                        "italic": is_italic,
                        "heading": is_heading
                    })
                elif child.find(".//w:br[@w:type='page']", ns) is not None:
                    elements.append({"type": "page_break"})
                    
            elif tag.endswith("tbl"):
                # Table
                table_rows = []
                for tr in child.findall("./w:tr", ns):
                    row_cells = []
                    for tc in tr.findall("./w:tc", ns):
                        cell_texts = []
                        for t in tc.findall(".//w:t", ns):
                            if t.text:
                                cell_texts.append(t.text)
                        row_cells.append(" ".join(cell_texts).strip())
                    if row_cells:
                        table_rows.append(row_cells)
                if table_rows:
                    elements.append({
                        "type": "table",
                        "rows": table_rows
                    })

    # Generate PostScript with pagination and word wrapping
    # A4 dimensions: 595 x 842 pt. Margins: left=54, right=541, top=788, bottom=54.
    PAGE_WIDTH = 595
    PAGE_HEIGHT = 842
    LEFT_MARGIN = 54
    RIGHT_MARGIN = 541
    MAX_LINE_WIDTH = RIGHT_MARGIN - LEFT_MARGIN # 487 pt
    START_Y = 780
    BOTTOM_LIMIT = 60

    ps_lines = [
        "%!PS",
        "/Helvetica findfont 10.5 scalefont setfont",
        "/Helvetica-Bold findfont 14 scalefont setfont",
        f"/y {START_Y} def",
        f"/newline {{ /y y 16 sub def {LEFT_MARGIN} y moveto }} def",
        f"/sectiongap {{ /y y 22 sub def {LEFT_MARGIN} y moveto }} def",
        f"/newpage {{ showpage /y {START_Y} def {LEFT_MARGIN} y moveto }} def",
        f"{LEFT_MARGIN} y moveto"
    ]

    current_y = START_Y

    def escape_ps(s: str) -> str:
        # PostScript string escaping
        res = ""
        for ch in s:
            if ch in "\\()":
                res += "\\" + ch
            elif ord(ch) < 128:
                res += ch
            else:
                # ASCII transliteration fallback for standard PostScript fonts
                res += "?"
        return res

    for el in elements:
        if el["type"] == "page_break":
            ps_lines.append("newpage")
            current_y = START_Y
            continue

        if el["type"] == "paragraph":
            is_head = el.get("heading") or el.get("bold")
            font_size = 14 if is_head else 10.5
            line_height = 20 if is_head else 15
            
            # Rough character width estimate (approx 0.52 * font_size)
            chars_per_line = int(MAX_LINE_WIDTH / (font_size * 0.52))

            # Check if need new page
            if current_y - line_height < BOTTOM_LIMIT:
                ps_lines.append("newpage")
                current_y = START_Y

            font_name = "/Helvetica-Bold" if is_head else ("/Helvetica-Oblique" if el.get("italic") else "/Helvetica")
            ps_lines.append(f"{font_name} findfont {font_size} scalefont setfont")
            
            words = el["text"].split()
            current_line_words = []
            
            for w in words:
                candidate = " ".join(current_line_words + [w])
                if len(candidate) > chars_per_line and current_line_words:
                    if current_y - line_height < BOTTOM_LIMIT:
                        ps_lines.append("newpage")
                        current_y = START_Y
                        ps_lines.append(f"{font_name} findfont {font_size} scalefont setfont")

                    ps_str = escape_ps(" ".join(current_line_words))
                    ps_lines.append(f"({ps_str}) show")
                    ps_lines.append("newline")
                    current_y -= line_height
                    current_line_words = [w]
                else:
                    current_line_words.append(w)
                    
            if current_line_words:
                if current_y - line_height < BOTTOM_LIMIT:
                    ps_lines.append("newpage")
                    current_y = START_Y
                    ps_lines.append(f"{font_name} findfont {font_size} scalefont setfont")
                ps_str = escape_ps(" ".join(current_line_words))
                ps_lines.append(f"({ps_str}) show")
                ps_lines.append("newline")
                current_y -= line_height

            # Paragraph gap
            ps_lines.append("/y y 6 sub def")
            current_y -= 6

        elif el["type"] == "table":
            table_rows = el.get("rows", [])
            if not table_rows:
                continue
            
            num_cols = max(len(r) for r in table_rows)
            if num_cols == 0:
                continue
            col_width = MAX_LINE_WIDTH / num_cols
            cell_font_size = 9.5
            row_height = 20

            ps_lines.append(f"/Helvetica findfont {cell_font_size} scalefont setfont")
            ps_lines.append("0.5 setlinewidth")

            for r_idx, row in enumerate(table_rows):
                if current_y - row_height < BOTTOM_LIMIT:
                    ps_lines.append("newpage")
                    current_y = START_Y
                    ps_lines.append(f"/Helvetica findfont {cell_font_size} scalefont setfont")
                    ps_lines.append("0.5 setlinewidth")

                row_top = current_y
                row_bottom = current_y - row_height
                
                # Draw cells
                for c_idx, cell_text in enumerate(row):
                    cell_x = LEFT_MARGIN + c_idx * col_width
                    # Truncate text if needed
                    max_chars = int(col_width / (cell_font_size * 0.52)) - 1
                    display_text = cell_text[:max_chars] if len(cell_text) > max_chars else cell_text
                    ps_text = escape_ps(display_text)
                    
                    # Position and print text
                    ps_lines.append(f"{cell_x + 4} {row_bottom + 5} moveto ({ps_text}) show")

                # Draw row underline
                ps_lines.append(f"{LEFT_MARGIN} {row_bottom} moveto {RIGHT_MARGIN} {row_bottom} lineto stroke")
                current_y -= row_height

            ps_lines.append("/y y 10 sub def")
            current_y -= 10

    ps_lines.append("showpage")
    ps_content = "\n".join(ps_lines)

    # Convert PS to PDF using ps2pdf
    with tempfile.NamedTemporaryFile(suffix=".ps", delete=False) as temp_ps:
        temp_ps.write(ps_content.encode("utf-8", errors="replace"))
        temp_ps_path = temp_ps.name

    temp_pdf_path = temp_ps_path.replace(".ps", ".pdf")
    try:
        subprocess.run(["ps2pdf", temp_ps_path, temp_pdf_path], check=True)
        with open(temp_pdf_path, "rb") as f:
            pdf_bytes = f.read()
        return pdf_bytes
    finally:
        if os.path.exists(temp_ps_path):
            os.remove(temp_ps_path)
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
