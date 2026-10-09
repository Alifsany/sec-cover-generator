import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls
import os

def create_cover_page():
    doc = docx.Document()
    section = doc.sections[0]
    
    # 1. Page Geometry (Standard Letter 8.5 x 11 inches)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.different_first_page_header_footer = False

    # 2. Add Background Watermark Logo in Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hrun = hp.add_run()
    
    watermark_path = os.path.abspath('c:/cover page/watermark_transparent.png')
    pic = hrun.add_picture(watermark_path, width=Inches(5.55))
    
    inline = pic._inline
    cx = inline.extent.cx
    cy = inline.extent.cy
    doc_id = inline.docPr.id
    doc_name = inline.docPr.name
    top_emu = int(2.42 * 914400) # Centered watermark vertical offset matching PDF

    anchor_xml = f"""
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
               xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
               distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240"
               behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page">
        <wp:align>center</wp:align>
      </wp:positionH>
      <wp:positionV relativeFrom="page">
        <wp:posOffset>{top_emu}</wp:posOffset>
      </wp:positionV>
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

    # 3. Document Body Content
    # Title: SHYAMOLI ENGINEERING COLLEGE
    # Font: Antonio (fallback Arial Narrow / Trebuchet MS / sans-serif), 43pt Bold
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(14)
    p1.paragraph_format.space_after = Pt(6)
    p1.paragraph_format.line_spacing = 1.0
    r1 = p1.add_run("SHYAMOLI ENGINEERING COLLEGE")
    r1.font.name = "Antonio"
    r1.font.size = Pt(43)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    # Ensure eastAsia / ascii font hints
    r1._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Antonio" w:hAnsi="Antonio" w:cs="Antonio"/>'))

    # Subtitle: A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA
    # Font: Times New Roman, 14pt Bold, Blue #077ACD
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(66)
    p2.paragraph_format.line_spacing = 1.0
    r2 = p2.add_run("A CONSTITUENT COLLEGE OF UNIVERSITY OF DHAKA")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x07, 0x7A, 0xCD) # rgb(7, 122, 205) matching (0.0275, 0.478, 0.804)

    # Department: DEPARTMENT OF \n COMPUTER SCIENCE AND ENGINEERING
    # Font: Comfortaa (SemiBold), 20pt
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(4)
    p3.paragraph_format.line_spacing = 1.15
    r3 = p3.add_run("DEPARTMENT OF")
    r3.font.name = "Comfortaa"
    r3.font.size = Pt(20)
    r3.font.bold = True
    r3.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r3._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after = Pt(48)
    p4.paragraph_format.line_spacing = 1.15
    r4 = p4.add_run("COMPUTER SCIENCE AND ENGINEERING")
    r4.font.name = "Comfortaa"
    r4.font.size = Pt(20)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r4._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Assignment Heading & Editable Title Placeholder
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(0)
    p5.paragraph_format.space_after = Pt(4)
    p5.paragraph_format.line_spacing = 1.15
    r5 = p5.add_run("ASSIGNMENT")
    r5.font.name = "Comfortaa"
    r5.font.size = Pt(20)
    r5.font.bold = True
    r5.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r5._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after = Pt(46)
    p6.paragraph_format.line_spacing = 1.15
    r6 = p6.add_run("[ENTER ASSIGNMENT TITLE]")
    r6.font.name = "Comfortaa"
    r6.font.size = Pt(20)
    r6.font.bold = True
    r6.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r6._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Course Code: COURSE CODE: CSE-1101
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.space_before = Pt(0)
    p7.paragraph_format.space_after = Pt(62)
    p7.paragraph_format.line_spacing = 1.15
    r7 = p7.add_run("COURSE CODE: CSE-1101")
    r7.font.name = "Comfortaa"
    r7.font.size = Pt(16)
    r7.font.bold = True
    r7.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r7._r.get_or_add_rPr().append(parse_xml(r'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Comfortaa" w:hAnsi="Comfortaa" w:cs="Comfortaa"/>'))

    # Two-Column Section: SUBMITTED BY & SUBMITTED TO
    # Exactly matching PDF layout: Left column at x=90pt, Right column at x=345pt
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Remove all table borders
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

    # Set column widths: Left = 3.3 in, Right = 3.2 in (Total 6.5 in printable)
    row = table.rows[0]
    cell_left = row.cells[0]
    cell_right = row.cells[1]
    cell_left.width = Inches(3.35)
    cell_right.width = Inches(3.15)

    # Zero margins on cells for exact alignment
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

    # Helper function to add lines to table cells
    def add_cell_line(cell, text, size_pt, bold=False, color_rgb=(0,0,0), space_after=2, font_name="Comfortaa"):
        # If cell has default empty paragraph and no text yet
        if len(cell.paragraphs) == 1 and cell.paragraphs[0].text == "":
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)
        run._r.get_or_add_rPr().append(parse_xml(f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'))
        return p

    # Left Column: SUBMITTED BY
    add_cell_line(cell_left, "SUBMITTED BY", 16, bold=True, color_rgb=(67, 67, 67), space_after=10)
    add_cell_line(cell_left, "Name: [ENTER YOUR FULL NAME]", 13, bold=False, color_rgb=(0, 0, 0), space_after=4)
    add_cell_line(cell_left, "Student ID: [ENTER YOUR STUDENT ID]", 11.5, bold=False, color_rgb=(0, 0, 0), space_after=4)
    add_cell_line(cell_left, "Semester: [ENTER YOUR YEAR AND SEMESTER]", 11.5, bold=False, color_rgb=(0, 0, 0), space_after=0)

    # Right Column: SUBMITTED TO
    add_cell_line(cell_right, "SUBMITTED TO", 16, bold=True, color_rgb=(67, 67, 67), space_after=10)
    add_cell_line(cell_right, "Teacher Name: [ENTER TEACHER'S FULL NAME]", 13, bold=False, color_rgb=(0, 0, 0), space_after=4)
    add_cell_line(cell_right, "Designation: [ENTER TEACHER'S DESIGNATION]", 11.5, bold=False, color_rgb=(0, 0, 0), space_after=4)
    add_cell_line(cell_right, "Institution: SHYAMOLI ENGINEERING COLLEGE", 11.5, bold=False, color_rgb=(0, 0, 0), space_after=0)

    output_path = os.path.abspath("c:/cover page/CSE-1101_Assignment_Cover_Template.docx")
    doc.save(output_path)
    print(f"Template successfully created at: {output_path}")

if __name__ == "__main__":
    create_cover_page()
