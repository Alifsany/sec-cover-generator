import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import simpleSplit

# Register TTF fonts for ReportLab if available
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def find_asset(filename):
    p1 = os.path.join(BASE_DIR, filename)
    if os.path.exists(p1):
        return p1
    p2 = os.path.join(BASE_DIR, 'static', filename)
    if os.path.exists(p2):
        return p2
    return p1

ANTONIO_FONT = find_asset('Antonio-Bold.ttf')
COMFORTAA_FONT = find_asset('Comfortaa-Regular.ttf')
WATERMARK_IMG = find_asset('watermark_transparent.png')

try:
    if os.path.exists(ANTONIO_FONT):
        pdfmetrics.registerFont(TTFont('Antonio-Bold', ANTONIO_FONT))
    if os.path.exists(COMFORTAA_FONT):
        pdfmetrics.registerFont(TTFont('Comfortaa-Regular', COMFORTAA_FONT))
except Exception as e:
    print('Font registration error:', e)

import re

def wrap_text(text, font_name, font_size, max_width):
    """Wraps text into lines that strictly fit within max_width, breaking long unbroken tokens if needed."""
    if text is None:
        return []
    text_str = str(text).strip()
    if not text_str:
        return []
    lines = simpleSplit(text_str, font_name, font_size, max_width)
    result = []
    for line in lines:
        if pdfmetrics.stringWidth(line, font_name, font_size) <= max_width:
            result.append(line)
        else:
            cur = ''
            parts = re.split(r'([-\s/_])', line)
            for p in parts:
                if not p:
                    continue
                test_str = (cur + p) if cur else p
                if pdfmetrics.stringWidth(test_str, font_name, font_size) <= max_width:
                    cur = test_str
                else:
                    if cur:
                        result.append(cur)
                        cur = ''
                    if pdfmetrics.stringWidth(p, font_name, font_size) > max_width:
                        char_cur = ''
                        for ch in p:
                            if pdfmetrics.stringWidth(char_cur + ch, font_name, font_size) <= max_width:
                                char_cur += ch
                            else:
                                if char_cur:
                                    result.append(char_cur)
                                char_cur = ch
                        cur = char_cur
                    else:
                        cur = p
            if cur:
                result.append(cur)
    return result

def generate_docx(title, course, sname, sid, semester, tname, tdesignation, output_path):
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.different_first_page_header_footer = False

    title_str = str(title).strip()
    course_str = str(course).strip()
    sname_str = str(sname).strip()
    sid_str = str(sid).strip()
    sem_str = str(semester).strip()
    tname_str = str(tname).strip()
    tdesig_str = str(tdesignation).strip()

    total_chars = len(title_str) + len(course_str) + len(sname_str) + len(tname_str) + len(tdesig_str) + len(sid_str) + len(sem_str)
    is_very_long = total_chars > 180 or len(title_str) > 75
    is_long = total_chars > 90 or len(title_str) > 40

    if is_very_long:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.5)
        space_after_sub = 28
        space_after_dept = 22
        space_after_title = 18
        space_after_course = 22
        p1_size = Pt(40)
        title_sz = Pt(15)
        course_sz = Pt(12)
        tbl_head_sz = 14
        tbl_body_sz = 11.5
        tbl_space_after = 2
        tbl_line_spacing = 1.05
    elif is_long:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.65)
        space_after_sub = 42
        space_after_dept = 32
        space_after_title = 28
        space_after_course = 36
        p1_size = Pt(41.5)
        title_sz = Pt(17.5)
        course_sz = Pt(13)
        tbl_head_sz = 15
        tbl_body_sz = 12.5
        tbl_space_after = 2
        tbl_line_spacing = 1.1
    else:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(0.8)
        space_after_sub = 60
        space_after_dept = 44
        space_after_title = 42
        space_after_course = 56
        p1_size = Pt(43)
        title_sz = Pt(20)
        course_sz = Pt(14)
        tbl_head_sz = 16
        tbl_body_sz = 14
        tbl_space_after = 3
        tbl_line_spacing = 1.15

    # Watermark in Header
    if os.path.exists(WATERMARK_IMG):
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hrun = hp.add_run()
        pic = hrun.add_picture(WATERMARK_IMG, width=Inches(5.55))
        
        inline = pic._inline
        cx = inline.extent.cx
        cy = inline.extent.cy
        doc_id = inline.docPr.id
        doc_name = inline.docPr.name
        top_emu = int(2.42 * 914400)

        anchor_xml = f"""
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240"
                   behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>
          <wp:positionV relativeFrom="page"><wp:posOffset>{top_emu}</wp:posOffset></wp:positionV>
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:wrapNone/>
          <wp:docPr id="{doc_id}" name="{doc_name}"/>
          <wp:cNvGraphicFramePr/>
        </wp:anchor>
        """
        anchor = parse_xml(anchor_xml)
        graphic = inline.graphic
        anchor.append(graphic)
        inline.getparent().replace(inline, anchor)

    # Paragraph 1: SHYAMOLI ENGINEERING COLLEGE
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(12)
    p1.paragraph_format.space_after = Pt(4)
    p1.paragraph_format.line_spacing = 1.0
    r1 = p1.add_run("SHYAMOLI ENGINEERING COLLEGE")
    r1.font.name = "Antonio"
    r1.font.size = p1_size
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0, 0, 0)
    r1._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Antonio" w:hAnsi="Antonio" w:cs="Antonio"/>'))

    # Paragraph 2: A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(space_after_sub)
    p2.paragraph_format.line_spacing = 1.0
    r2 = p2.add_run("A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(7, 122, 205)

    # Paragraph 3: DEPARTMENT OF
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(4)
    p3.paragraph_format.line_spacing = 1.15
    r3 = p3.add_run("DEPARTMENT OF")
    r3.font.name = "Comfortaa"
    r3.font.size = Pt(19 if is_very_long else 20)
    r3.font.bold = True
    r3._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Paragraph 4: COMPUTER SCIENCE AND ENGINEERING
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after = Pt(space_after_dept)
    p4.paragraph_format.line_spacing = 1.15
    r4 = p4.add_run("COMPUTER SCIENCE AND ENGINEERING")
    r4.font.name = "Comfortaa"
    r4.font.size = Pt(19 if is_very_long else 20)
    r4.font.bold = True
    r4._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Paragraph 5: ASSIGNMENT
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(0)
    p5.paragraph_format.space_after = Pt(4)
    p5.paragraph_format.line_spacing = 1.15
    r5 = p5.add_run("ASSIGNMENT")
    r5.font.name = "Comfortaa"
    r5.font.size = Pt(19 if is_very_long else 20)
    r5.font.bold = True
    r5._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Paragraph 6: Assignment Title
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after = Pt(space_after_title)
    p6.paragraph_format.line_spacing = 1.15
    r6 = p6.add_run(title_str)
    r6.font.name = "Comfortaa"
    r6.font.size = title_sz
    r6.font.bold = True
    r6._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Paragraph 7: Course Code
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.space_before = Pt(0)
    p7.paragraph_format.space_after = Pt(space_after_course)
    p7.paragraph_format.line_spacing = 1.15
    course_display = course_str if course_str.upper().startswith("COURSE CODE:") else f"COURSE CODE: {course_str}"
    r7 = p7.add_run(course_display)
    r7.font.name = "Comfortaa"
    r7.font.size = course_sz
    r7.font.bold = True
    r7._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Lower Two-Column Section: Table
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(r"""
        <w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
    """)
    tblPr.append(tblBorders)

    row = table.rows[0]
    cell_left = row.cells[0]
    cell_right = row.cells[1]
    cell_left.width = Inches(3.35)
    cell_right.width = Inches(3.15)

    tcPr_left = cell_left._tc.get_or_add_tcPr()
    tcMar_left = parse_xml(r"""
        <w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:w="0" w:type="dxa"/>
            <w:left w:w="260" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right w:w="140" w:type="dxa"/>
        </w:tcMar>
    """)
    tcPr_left.append(tcMar_left)

    tcPr_right = cell_right._tc.get_or_add_tcPr()
    tcMar_right = parse_xml(r"""
        <w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:w="0" w:type="dxa"/>
            <w:left w:w="140" w:type="dxa"/>
            <w:bottom w:w="0" w:type="dxa"/>
            <w:right w:w="0" w:type="dxa"/>
        </w:tcMar>
    """)
    tcPr_right.append(tcMar_right)

    def add_cell_line(cell, text, size_pt=14, bold=False, color_rgb=(0,0,0), space_after=3):
        if len(cell.paragraphs) == 1 and cell.paragraphs[0].text == "":
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = tbl_line_spacing
        run = p.add_run(text)
        run.font.name = "Comfortaa"
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)
        run._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))
        return p

    # Left Column: SUBMITTED BY
    add_cell_line(cell_left, "SUBMITTED BY", tbl_head_sz, bold=True, color_rgb=(67, 67, 67), space_after=tbl_space_after + 4)
    add_cell_line(cell_left, f"Name: {sname_str}", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=tbl_space_after)
    add_cell_line(cell_left, f"Student ID: {sid_str}", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=tbl_space_after)
    add_cell_line(cell_left, f"Semester: {sem_str}", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=0)

    # Right Column: SUBMITTED TO (Teacher Name directly below without "Teacher Name:" label)
    add_cell_line(cell_right, "SUBMITTED TO", tbl_head_sz, bold=True, color_rgb=(67, 67, 67), space_after=tbl_space_after + 4)
    add_cell_line(cell_right, f"{tname_str}", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=tbl_space_after)
    add_cell_line(cell_right, f"{tdesig_str}", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=tbl_space_after)
    add_cell_line(cell_right, "SHYAMOLI ENGINEERING COLLEGE", tbl_body_sz, bold=False, color_rgb=(0, 0, 0), space_after=0)

    doc.save(output_path)
    return output_path

def generate_pdf(title, course, sname, sid, semester, tname, tdesignation, output_path):
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter # 612 x 792 pt

    # Background Watermark
    if os.path.exists(WATERMARK_IMG):
        c.drawImage(WATERMARK_IMG, 106, 174, width=399.5, height=399.5, mask='auto', preserveAspectRatio=True)

    font_title = 'Antonio-Bold' if 'Antonio-Bold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold'
    font_sub = 'Times-Bold'
    font_body = 'Comfortaa-Regular' if 'Comfortaa-Regular' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'
    font_bold = 'Comfortaa-Regular' if 'Comfortaa-Regular' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold'

    # 1. Heading: SHYAMOLI ENGINEERING COLLEGE
    c.setFillColor(HexColor('#000000'))
    c.setFont(font_title, 43)
    c.drawCentredString(width / 2.0, 630.3, "SHYAMOLI ENGINEERING COLLEGE")

    # 2. Institutional subtitle: A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA
    c.setFillColor(HexColor('#077ACD'))
    c.setFont(font_sub, 14)
    c.drawCentredString(width / 2.0, 594.5, "A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA")

    # 3. DEPARTMENT OF & COMPUTER SCIENCE AND ENGINEERING
    c.setFillColor(HexColor('#000000'))
    c.setFont(font_bold, 20)
    c.drawCentredString(width / 2.0, 501.2, "DEPARTMENT OF")
    c.drawCentredString(width / 2.0, 473.9, "COMPUTER SCIENCE AND ENGINEERING")

    # 4. ASSIGNMENT
    c.drawCentredString(width / 2.0, 401.6, "ASSIGNMENT")

    # 5. Assignment Title: Dynamic Font Sizing & Complete Wrapping (No Truncation)
    title_str = str(title).strip()
    title_fs = 20
    for fs_cand in [20, 19, 18, 17, 16, 15, 14, 13]:
        t_lines = wrap_text(title_str, font_bold, fs_cand, 480)
        if len(t_lines) <= 2:
            title_fs = fs_cand
            break
        title_fs = fs_cand
    title_lines = wrap_text(title_str, font_bold, title_fs, 480)
    title_leading = title_fs * 1.22

    if len(title_lines) == 1:
        y_title_start = 374.3
    elif len(title_lines) == 2:
        y_title_start = 378.0
    else:
        y_title_start = min(382.0, 401.6 - title_fs - 6.0)

    c.setFont(font_bold, title_fs)
    y_curr = y_title_start
    for t_line in title_lines:
        c.drawCentredString(width / 2.0, y_curr, t_line)
        y_curr -= title_leading
    y_title_bottom = y_curr + title_leading - (title_fs * 0.3)

    # 6. Course Code: Dynamic Font Sizing & Safe Wrapping
    course_str = str(course).strip()
    course_display = course_str if course_str.upper().startswith("COURSE CODE:") else f"COURSE CODE: {course_str}"
    course_fs = 14
    for c_cand in [14, 13, 12, 11]:
        c_lines = wrap_text(course_display, font_bold, c_cand, 480)
        if len(c_lines) <= 1:
            course_fs = c_cand
            break
        course_fs = c_cand
    course_lines = wrap_text(course_display, font_bold, course_fs, 480)
    course_leading = course_fs * 1.22

    y_course_start = min(305.5, y_title_bottom - 26.0)
    c.setFont(font_bold, course_fs)
    y_c_curr = y_course_start
    for cline in course_lines:
        c.drawCentredString(width / 2.0, y_c_curr, cline)
        y_c_curr -= course_leading
    course_bottom = y_c_curr + course_leading - (course_fs * 0.3)

    # 7. Lower Section: SUBMITTED BY (x = 90.3 pt) & SUBMITTED TO (x = 345.3 pt)
    sname_str = str(sname).strip()
    sid_str = str(sid).strip()
    sem_str = str(semester).strip()
    tname_str = str(tname).strip()
    tdesig_str = str(tdesignation).strip()

    # Per-field adaptive base sizing for long entries
    sname_base_fs = 11.0 if len(sname_str) > 50 else (12.5 if len(sname_str) > 28 else 14.0)
    tname_base_fs = 11.0 if len(tname_str) > 50 else (12.5 if len(tname_str) > 28 else 14.0)
    tdesig_base_fs = 10.5 if len(tdesig_str) > 55 else (12.0 if len(tdesig_str) > 30 else 14.0)
    sid_base_fs = 12.0 if len(sid_str) > 24 else 14.0
    sem_base_fs = 12.0 if len(sem_str) > 28 else 14.0
    college_base_fs = 14.0

    # Iterative layout scaling to strictly satisfy bottom page boundary (lowest_y >= 45 pt)
    col_width = 230.0
    best_scale = 1.0
    best_heading_y = 216.3
    best_heading_fs = 16.0
    best_left_items = []
    best_right_items = []
    best_gap_under_heading = 16.0
    best_field_gap = 4.0

    for scale in [1.0, 0.95, 0.9, 0.85, 0.8, 0.75]:
        heading_fs = max(13.0, round(16.0 * scale, 1))
        gap_under_heading = round(16.0 * scale, 1)
        field_gap = round(4.0 * scale, 1)

        left_items = [
            (f"Name: {sname_str}", max(10.0, sname_base_fs * scale)),
            (f"Student ID: {sid_str}", max(10.0, sid_base_fs * scale)),
            (f"Semester: {sem_str}", max(10.0, sem_base_fs * scale))
        ]
        # Teacher Name displayed directly without "Teacher Name:" label
        right_items = [
            (tname_str, max(10.0, tname_base_fs * scale)),
            (tdesig_str, max(10.0, tdesig_base_fs * scale)),
            ("SHYAMOLI ENGINEERING COLLEGE", max(10.0, college_base_fs * scale))
        ]

        # Calculate left column height
        left_h = gap_under_heading
        for text, fs in left_items:
            lines = wrap_text(text, font_body, fs, col_width)
            lh = fs * 1.2
            left_h += len(lines) * lh + field_gap

        # Calculate right column height
        right_h = gap_under_heading
        for text, fs in right_items:
            lines = wrap_text(text, font_body, fs, col_width)
            lh = fs * 1.2
            right_h += len(lines) * lh + field_gap

        total_meta_h = max(left_h, right_h)
        heading_y = 216.3 if total_meta_h < 130 else min(226.0, course_bottom - 26.0)
        lowest_y = heading_y - total_meta_h

        best_scale = scale
        best_heading_y = heading_y
        best_heading_fs = heading_fs
        best_left_items = left_items
        best_right_items = right_items
        best_gap_under_heading = gap_under_heading
        best_field_gap = field_gap

        if lowest_y >= 45.0:
            break

    # Draw Metadata Headings
    c.setFillColor(HexColor('#434343'))
    c.setFont(font_bold, best_heading_fs)
    c.drawString(90.3, best_heading_y, "SUBMITTED BY")
    c.drawString(345.3, best_heading_y, "SUBMITTED TO")

    # Draw Left Column: SUBMITTED BY
    c.setFillColor(HexColor('#000000'))
    y_l_curr = best_heading_y - best_gap_under_heading
    for text, fs in best_left_items:
        lines = wrap_text(text, font_body, fs, col_width)
        c.setFont(font_body, fs)
        lh = fs * 1.2
        for l in lines:
            c.drawString(90.3, y_l_curr, l)
            y_l_curr -= lh
        y_l_curr -= best_field_gap

    # Draw Right Column: SUBMITTED TO (Teacher Name, Designation, College Name)
    y_r_curr = best_heading_y - best_gap_under_heading
    for text, fs in best_right_items:
        lines = wrap_text(text, font_body, fs, col_width)
        c.setFont(font_body, fs)
        lh = fs * 1.2
        for l in lines:
            c.drawString(345.3, y_r_curr, l)
            y_r_curr -= lh
        y_r_curr -= best_field_gap

    c.showPage()
    c.save()
    return output_path

