import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

TARGET_DIR = r"C:\\Users\\MIP-USER\\Documents\\Project Soal"

# ==========================================
# HEADER TABLE STYLE HELPER
# ==========================================
def set_cell_shading(cell, color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def create_header_table(doc, subject, paket_num):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, "EBF5FB")
    # Set borders
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = parse_xml(
        '<w:tcBorders ' + nsdecls('w') + '>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/>'
        '</w:tcBorders>'
    )
    tc_pr.append(tc_borders)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Line 1
    r1 = p.add_run("TRY OUT TES KEMAMPUAN AKADEMIK (TKA) SD")
    r1.bold = True
    r1.font.size = Pt(14)
    r1.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    r1.font.name = "Calibri"
    
    # Line 2
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"MATA PELAJARAN: {subject.upper()} — PAKET {paket_num}")
    r2.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor(0x2E, 0x5B, 0x88)
    r2.font.name = "Calibri"
    
    # Line 3
    p3 = cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p3.add_run("Standar Kurikulum Merdeka & K-13 (Kerangka Resmi Pusmendik Kemendikdasmen RI)")
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(0x64, 0x64, 0x64)
    r3.font.name = "Calibri"
    
    # Line 4
    p4 = cell.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run("Jumlah Soal: 20 Butir | Alokasi Waktu: 45 Menit | Fokus: Numerasi, Aljabar & Geometri Dasar")
    r4.bold = True
    r4.font.size = Pt(10)
    r4.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)
    r4.font.name = "Calibri"
    
    return table

# ==========================================
# PETUNJUK UMUM
# ==========================================
def add_petunjuk_umum(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("PETUNJUK UMUM:")
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = "Calibri"
    
    petunjuk_text = [
        "1. Berdoalah sebelum mengerjakan soal.",
        "2. Periksa dan bacalah setiap butir soal dengan teliti.",
        "3. Soal terdiri atas Pilihan Ganda Sederhana, Pilihan Ganda Kompleks (jawaban lebih dari satu), dan Tabel Kategori (Benar/Salah).",
        "4. Dahulukan soal yang Anda anggap mudah."
    ]
    for i, pt in enumerate(petunjuk_text):
        pp = doc.add_paragraph()
        pp.paragraph_format.space_before = Pt(2)
        pp.paragraph_format.space_after = Pt(2)
        pp.paragraph_format.left_indent = Inches(0.3)
        rp = pp.add_run(pt)
        rp.font.size = Pt(11)
        rp.font.name = "Calibri"
    
    # Garis pemisah
    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_div = p_div.add_run("═" * 55)
    r_div.font.color.rgb = RGBColor(0x93, 0xC5, 0xFD)
    r_div.font.size = Pt(10)

# ==========================================
# BAGIAN A HEADER
# ==========================================
def add_bagian_a_header(doc):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(8)
    r = h.add_run("BAGIAN A: NASKAH SOAL UJIAN")
    r.font.size = Pt(12)
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    r.font.name = "Calibri"

# ==========================================
# ADD QUESTION TO DOCX
# ==========================================
def add_question_docx(doc, q, q_idx):
    # Question text
    p_q = doc.add_paragraph()
    p_q.paragraph_format.space_before = Pt(8)
    p_q.paragraph_format.space_after = Pt(4)
    p_q.paragraph_format.left_indent = Inches(0.2)
    
    run_num = p_q.add_run(f"{q['number']}. ")
    run_num.bold = True
    run_num.font.size = Pt(11)
    run_num.font.name = "Calibri"
    
    # Handle multi-line questions
    q_lines = q["question"].split("\n")
    if len(q_lines) > 1:
        p_q.add_run(q_lines[0] + "\n")
        for line in q_lines[1:]:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(2)
            p_sub.paragraph_format.left_indent = Inches(0.5)
            if line.strip().startswith(("•", "-", "*")) or "Pernyataan" in q.get("tipe", ""):
                r = p_sub.add_run(line.strip())
                r.bold = True
                r.font.size = Pt(11)
                r.font.name = "Calibri"
            else:
                p_sub.add_run(line.strip()).font.size = Pt(11)
    else:
        p_q.add_run(q["question"]).font.size = Pt(11)
    
    # Options
    tipe = q.get("tipe", "Pilihan Ganda")
    
    if tipe == "MCMA" or tipe == "Pilihan Ganda Kompleks":
        # Checkbox style for MCMA
        for opt_idx, opt_text in enumerate(q["options"]):
            opt_letter = chr(65 + opt_idx)
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Inches(0.5)
            p_opt.paragraph_format.space_before = Pt(1)
            p_opt.paragraph_format.space_after = Pt(1)
            r_opt_l = p_opt.add_run(f"☐ {opt_letter}. ")
            r_opt_l.bold = True
            r_opt_l.font.size = Pt(11)
            r_opt_l.font.name = "Calibri"
            p_opt.add_run(opt_text).font.size = Pt(11)
        
        # Label MCMA
        p_label = doc.add_paragraph()
        p_label.paragraph_format.left_indent = Inches(0.5)
        p_label.paragraph_format.space_before = Pt(2)
        p_label.paragraph_format.space_after = Pt(4)
        r_label = p_label.add_run("[Pilihan Ganda Kompleks: Jawaban benar lebih dari satu]")
        r_label.italic = True
        r_label.font.size = Pt(9.5)
        r_label.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
        r_label.font.name = "Calibri"
        
    elif tipe == "Kategori" or tipe == "Tabel Benar/Salah":
        # Create table for Kategori
        table = doc.add_table(rows=len(q["options"]) + 1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        
        # Header row
        hdr = table.rows[0].cells
        hdr[0].text = "Pernyataan"
        hdr[1].text = "Benar"
        hdr[2].text = "Salah"
        for cell in hdr:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(10)
                    r.font.name = "Calibri"
        
        # Data rows
        for i, stmt in enumerate(q["options"]):
            row = table.rows[i + 1].cells
            row[0].text = f"{chr(65+i)}. {stmt}"
            row[1].text = "[   ]"
            row[2].text = "[   ]"
            for cell in row:
                for p in cell.paragraphs:
                    if cell is row[0]:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.font.size = Pt(10)
                        r.font.name = "Calibri"
        
        # Label Kategori
        p_label = doc.add_paragraph()
        p_label.paragraph_format.space_before = Pt(2)
        p_label.paragraph_format.space_after = Pt(4)
        r_label = p_label.add_run("[Pilihan Ganda Kompleks: Kategori Benar / Salah]")
        r_label.italic = True
        r_label.font.size = Pt(9.5)
        r_label.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
        r_label.font.name = "Calibri"
        
    else:
        # Pilihan Ganda Biasa (Single Select)
        for opt_idx, opt_text in enumerate(q["options"]):
            opt_letter = chr(65 + opt_idx)
            p_opt = doc.add_paragraph()
            p_opt.paragraph_format.left_indent = Inches(0.5)
            p_opt.paragraph_format.space_before = Pt(1)
            p_opt.paragraph_format.space_after = Pt(1)
            r_opt_l = p_opt.add_run(f"{opt_letter}. ")
            r_opt_l.bold = True
            r_opt_l.font.size = Pt(11)
            r_opt_l.font.name = "Calibri"
            p_opt.add_run(opt_text).font.size = Pt(11)

# ==========================================
# BAGIAN B HEADER
# ==========================================
def add_bagian_b_header(doc):
    doc.add_page_break()
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(8)
    r = h.add_run("BAGIAN B: KUNCI JAWABAN & PEMBAHASAN LENGKAP")
    r.font.size = Pt(13)
    r.bold = True
    r.font.color.rgb = RGBColor(0x15, 0x80, 0x3D)
    r.font.name = "Calibri"

# ==========================================
# KUNCI JAWABAN CEPAT TABLE
# ==========================================
def add_kunci_cepat_table(doc, paket_data):
    p = doc.add_paragraph()
    p.add_run("TABEL KUNCI JAWABAN CEPAT:\n").bold = True
    p.paragraph_format.space_after = Pt(6)
    
    table = doc.add_table(rows=5, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Light Shading Accent 1'
    
    # Headers
    hdr = table.rows[0].cells
    hdr[0].text = "No 1-5"
    hdr[1].text = "No 6-10"
    hdr[2].text = "No 11-15"
    hdr[3].text = "No 16-20"
    for cell in hdr:
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9)
    
    # Fill
    for r in range(1, 5):
        row_cells = table.rows[r].cells
        for c in range(4):
            idx = c * 5 + (r - 1)
            if idx < len(paket_data):
                q = paket_data[idx]
                correct = q["correct"]
                if isinstance(correct, list):
                    correct = ", ".join(correct)
                elif isinstance(correct, dict):
                    correct = ", ".join([f"{k}:{v}" for k,v in correct.items()])
                row_cells[c].text = f"No {q['number']}: {correct}"
                for p in row_cells[c].paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.font.size = Pt(9)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

# ==========================================
# PEMBAHASAN PER SOAL
# ==========================================
def add_pembahasan_docx(doc, q):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    
    correct = q["correct"]
    if isinstance(correct, list):
        correct = ", ".join(correct)
    elif isinstance(correct, dict):
        correct = ", ".join([f"{k}:{v}" for k,v in correct.items()])
    
    r_num = p.add_run(f"Soal Nomor {q['number']}  |  Kunci Jawaban: {correct}\n")
    r_num.bold = True
    r_num.font.size = Pt(11)
    r_num.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    r_num.font.name = "Calibri"
    
    r_mat = p.add_run(f"Materi / Lingkup Kompetensi: {q['materi']}\n")
    r_mat.italic = True
    r_mat.font.size = Pt(10)
    r_mat.font.color.rgb = RGBColor(0x4B, 0x55, 0x63)
    r_mat.font.name = "Calibri"
    
    p.add_run(f"Pembahasan & Langkah Penyelesaian:\n{q['explanation']}").font.size = Pt(11)
    
    # Separator line
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(2)
    p_line.paragraph_format.space_after = Pt(6)
    r_line = p_line.add_run("—" * 50)
    r_line.font.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)
    r_line.font.size = Pt(9)

# ==========================================
# BUILD COMPLETE DOCX
# ==========================================
def build_docx(paket_data, subject, paket_num, filename):
    doc = docx.Document()
    
    # Margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
    
    # Default style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
    
    # Build
    create_header_table(doc, subject, paket_num)
    doc.add_paragraph()  # spacer
    add_petunjuk_umum(doc)
    add_bagian_a_header(doc)
    
    for q in paket_data:
        add_question_docx(doc, q, q["number"])
    
    add_bagian_b_header(doc)
    add_kunci_cepat_table(doc, paket_data)
    
    for q in paket_data:
        add_pembahasan_docx(doc, q)
    
    out_path = os.path.join(TARGET_DIR, filename)
    doc.save(out_path)
    print(f"[OK] DOCX: {out_path}")
    return out_path

# ==========================================
# PDF GENERATOR (ReportLab)
# ==========================================
def build_pdf(paket_data, subject, paket_num, filename):
    out_path = os.path.join(TARGET_DIR, filename)
    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    
    styles = getSampleStyleSheet()
    style_normal = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontName='Helvetica', fontSize=9.5, leading=13, textColor=colors.HexColor('#1F2937'))
    style_title = ParagraphStyle('CustomTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=colors.HexColor('#1E3A8A'), alignment=1)
    style_subtitle = ParagraphStyle('CustomSub', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=12, textColor=colors.HexColor('#4B5563'), alignment=1)
    style_heading = ParagraphStyle('CustomHeading', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=colors.HexColor('#1E3A8A'), spaceBefore=10, spaceAfter=6)
    style_q_bold = ParagraphStyle('QBold', parent=style_normal, fontName='Helvetica-Bold', fontSize=10, leading=13, textColor=colors.HexColor('#1E3A8A'))
    style_kunci = ParagraphStyle('Kunci', parent=style_normal, fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=colors.HexColor('#15803D'))
    
    story = []
    
    # Header
    header_data = [[
        f"TRY OUT TES KEMAMPUAN AKADEMIK (TKA) SD\n"
        f"MATA PELAJARAN: {subject.upper()} — PAKET {paket_num}\n"
        f"Standar Kurikulum Merdeka & K-13 (Kerangka Resmi Pusmendik Kemendikdasmen RI)\n"
        f"Jumlah Soal: 20 Butir | Alokasi Waktu: 45 Menit | Fokus: Literasi & Numerasi Dasar"
    ]]
    header_table = Table(header_data, colWidths=[490])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#EBF5FB')),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#4472C4')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#4472C4')),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 12))
    
    # Petunjuk Umum
    petunjuk = [
        "PETUNJUK UMUM:",
        "1. Berdoalah sebelum mengerjakan soal.",
        "2. Periksa dan bacalah setiap butir soal dengan teliti.",
        "3. Soal terdiri atas Pilihan Ganda Sederhana, Pilihan Ganda Kompleks (jawaban lebih dari satu), dan Tabel Kategori (Benar/Salah).",
        "4. Dahulukan soal yang Anda anggap mudah."
    ]
    for pt in petunjuk:
        story.append(Paragraph(f"<b>{pt}</b>" if pt.startswith("PETUNJUK") else pt, style_normal))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#93C5FD')))
    story.append(Spacer(1, 8))
    
    # BAGIAN A
    story.append(Paragraph("BAGIAN A: NASKAH SOAL UJIAN", style_heading))
    
    for q in paket_data:
        tipe = q.get("tipe", "Pilihan Ganda")
        
        # Question
        q_text = f"<b>{q['number']}.</b> {q['question'].replace(chr(10), '<br/>')}"
        story.append(Paragraph(q_text, style_q_bold))
        story.append(Spacer(1, 4))
        
        # Options
        if tipe in ["MCMA", "Pilihan Ganda Kompleks"]:
            for opt_idx, opt_text in enumerate(q["options"]):
                opt_letter = chr(65 + opt_idx)
                story.append(Paragraph(f"☐ <b>{opt_letter}.</b> {opt_text}", style_normal))
            story.append(Paragraph("<i>[Pilihan Ganda Kompleks: Jawaban benar lebih dari satu]</i>", ParagraphStyle('LabelMCMA', parent=style_normal, fontSize=9, textColor=colors.HexColor('#4B5563'), fontName='Helvetica-Oblique')))
            
        elif tipe in ["Kategori", "Tabel Benar/Salah"]:
            # Table for Kategori
            table_data = [["Pernyataan", "Benar", "Salah"]]
            for i, stmt in enumerate(q["options"]):
                table_data.append([f"{chr(65+i)}. {stmt}", "[   ]", "[   ]"])
            
            t = Table(table_data, colWidths=[370, 60, 60])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EBF5FB')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('ALIGN', (1,1), (-1,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(t)
            story.append(Paragraph("<i>[Pilihan Ganda Kompleks: Kategori Benar / Salah]</i>", ParagraphStyle('LabelKat', parent=style_normal, fontSize=9, textColor=colors.HexColor('#4B5563'), fontName='Helvetica-Oblique')))
            
        else:
            for opt_idx, opt_text in enumerate(q["options"]):
                opt_letter = chr(65 + opt_idx)
                story.append(Paragraph(f"<b>{opt_letter}.</b> {opt_text}", style_normal))
        
        story.append(Spacer(1, 8))
    
    # BAGIAN B
    story.append(PageBreak())
    story.append(Paragraph("BAGIAN B: KUNCI JAWABAN & PEMBAHASAN LENGKAP", ParagraphStyle('SecB', parent=style_heading, fontSize=12, textColor=colors.HexColor('#15803D'))))
    story.append(Spacer(1, 6))
    
    # Kunci Cepat Table
    story.append(Paragraph("<b>TABEL KUNCI JAWABAN CEPAT:</b>", style_normal))
    story.append(Spacer(1, 4))
    
    table_data = [["No 1-5", "No 6-10", "No 11-15", "No 16-20"]]
    for r in range(5):
        row = []
        for c in range(4):
            idx = c * 5 + r
            if idx < len(paket_data):
                q = paket_data[idx]
                correct = q["correct"]
                if isinstance(correct, list):
                    correct = ", ".join(correct)
                elif isinstance(correct, dict):
                    correct = ", ".join([f"{k}:{v}" for k,v in correct.items()])
                row.append(f"No {q['number']}: {correct}")
            else:
                row.append("")
        table_data.append(row)
    
    t = Table(table_data, colWidths=[120]*4)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EBF5FB')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t)
    story.append(Spacer(1, 12))
    
    # Pembahasan per soal
    for q in paket_data:
        correct = q["correct"]
        if isinstance(correct, list):
            correct = ", ".join(correct)
        elif isinstance(correct, dict):
            correct = ", ".join([f"{k}:{v}" for k,v in correct.items()])
        
        story.append(Paragraph(f"<b><font color='#1E3A8A'>Soal Nomor {q['number']}  |  Kunci Jawaban: {correct}</font></b>", style_normal))
        story.append(Paragraph(f"<i>Materi / Lingkup Kompetensi: {q['materi']}</i>", style_normal))
        story.append(Paragraph(f"Pembahasan & Langkah Penyelesaian:<br/>{q['explanation'].replace(chr(10), '<br/>')}", style_normal))
        story.append(HRFlowable(width="80%", thickness=0.5, color=colors.HexColor('#E5E7EB')))
        story.append(Spacer(1, 8))
    
    doc.build(story)
    print(f"[OK] PDF: {out_path}")
    return out_path

print("Generator functions loaded.")
