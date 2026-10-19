import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.graphics.shapes import Drawing, Rect, Line, Circle, Polygon, String, Ellipse

pdf_path = r"C:/Users/MIP-USER/Documents/Project Soal/Bonus_Kumpulan_Rumus_Matematika_SD.pdf"
doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
story = []

styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, leading=24,
    textColor=colors.HexColor('#4338ca'), alignment=1, spaceAfter=8
)
subtitle_style = ParagraphStyle(
    'DocSubTitle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14,
    textColor=colors.HexColor('#475569'), alignment=1, spaceAfter=15
)
h1_style = ParagraphStyle(
    'SectionH1', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, leading=16,
    textColor=colors.HexColor('#ffffff'), spaceBefore=12, spaceAfter=6,
    backColor=colors.HexColor('#4f46e5'), borderPadding=5
)
h2_style = ParagraphStyle(
    'SectionH2', parent=styles['Heading3'], fontName='Helvetica-Bold', fontSize=10, leading=14,
    textColor=colors.HexColor('#312e81'), spaceBefore=8, spaceAfter=4
)
body_style = ParagraphStyle(
    'BodyTextCustom', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=12,
    textColor=colors.HexColor('#1e293b'), spaceAfter=4
)
code_style = ParagraphStyle(
    'CodeBox', parent=styles['Normal'], fontName='Courier', fontSize=8.5, leading=11,
    textColor=colors.HexColor('#0f172a'), backColor=colors.HexColor('#f1f5f9'),
    borderPadding=4, spaceAfter=4
)

# Title Block
story.append(Paragraph("🎁 BONUS EKSKLUSIF: KUMPULAN RUMUS CEPAT MATEMATIKA SD", ParagraphStyle('Banner', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#4f46e5'), alignment=1, spaceAfter=6)))
story.append(Paragraph("SmartDrill SD 2026 — Panduan Praktis Siap Ujian TKA & Sekolah", title_style))
story.append(Paragraph("Ringkasan rumus aritmatika, geometri lengkap dengan ilustrasi gambar bangun datar/ruang, satuan ukur, dan trik hitung cepat.", subtitle_style))
story.append(Spacer(1, 6))

# 1. Aritmatika
story.append(Paragraph("1. ARITMATIKA DASAR", h1_style))
story.append(Paragraph("1.1 Operasi Hitung Campuran (PMDAS / KAKATUA)", h2_style))
story.append(Paragraph("<b>Urutan Pengerjaan:</b> Pangkat/Akar → Kurung <code>( )</code> → Kali & Bagi (kiri ke kanan) → Tambah & Kurang (kiri ke kanan).", body_style))

story.append(Paragraph("1.2 Pecahan, Desimal & Persen", h2_style))
story.append(Paragraph("• Penjumlahan/Pengurangan: <code>a/b ± c/d = (ad ± bc) / bd</code><br/>• Perkalian & Pembagian: <code>a/b × c/d = ac / bd</code> | <code>a/b ÷ c/d = ad / bc</code><br/>• Konversi Cepat: <code>½ = 0,5 = 50%</code> | <code>¼ = 0,25 = 25%</code> | <code>¾ = 0,75 = 75%</code>", code_style))

story.append(Paragraph("1.3 FPB dan KPK", h2_style))
story.append(Paragraph("• <b>FPB:</b> Faktor sama pangkat terkecil | <b>KPK:</b> Semua faktor pangkat terbesar<br/>• Rumus Cepat: <code>a × b = FPB(a,b) × KPK(a,b)</code>", body_style))

story.append(Spacer(1, 8))

# 2. Geometri 2D with vector diagrams
story.append(Paragraph("2. BANGUN DATAR (2D) & ILUSTRASI", h1_style))

def draw_triangle():
    d = Drawing(120, 70)
    d.add(Polygon([10,5, 110,5, 60,65], fillColor=colors.HexColor('#eef2ff'), strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Line(60,5, 60,65, strokeColor=colors.HexColor('#dc2626'), strokeWidth=1, strokeDashArray=[2,2]))
    d.add(String(60, -2, "alas (a)", fontSize=7, fontName="Helvetica-Bold", textAnchor="middle", fillColor=colors.HexColor('#4338ca')))
    d.add(String(48, 35, "t", fontSize=8, fontName="Helvetica-Bold", fillColor=colors.HexColor('#dc2626')))
    return d

def draw_rectangle():
    d = Drawing(120, 70)
    d.add(Rect(10, 10, 100, 50, fillColor=colors.HexColor('#fef3c7'), strokeColor=colors.HexColor('#f59e0b'), strokeWidth=1.5))
    d.add(String(60, 0, "panjang (p)", fontSize=7, fontName="Helvetica-Bold", textAnchor="middle", fillColor=colors.HexColor('#92400e')))
    d.add(String(2, 32, "l", fontSize=7, fontName="Helvetica-Bold", fillColor=colors.HexColor('#92400e')))
    return d

def draw_circle():
    d = Drawing(120, 70)
    d.add(Circle(60, 35, 28, fillColor=colors.HexColor('#faf5ff'), strokeColor=colors.HexColor('#7c3aed'), strokeWidth=1.5))
    d.add(Line(60, 35, 88, 35, strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(String(70, 38, "r", fontSize=8, fontName="Helvetica-Bold", fillColor=colors.HexColor('#4f46e5')))
    return d

t_data1 = [
    [draw_triangle(), Paragraph("<b>Segitiga:</b><br/>• Luas = <code>½ × a × t</code><br/>• Keliling = <code>s₁ + s₂ + s₃</code><br/>• Siku-siku: <code>a² + b² = c²</code>", body_style)],
    [draw_rectangle(), Paragraph("<b>Persegi & Persegi Panjang:</b><br/>• Persegi: L = <code>s²</code>, K = <code>4s</code><br/>• P. Panjang: L = <code>p × l</code>, K = <code>2(p+l)</code>", body_style)],
    [draw_circle(), Paragraph("<b>Lingkaran:</b><br/>• Luas = <code>π × r²</code><br/>• Keliling = <code>2 × π × r</code><br/>• d = <code>2r</code> (Diameter)", body_style)]
]
t1 = Table(t_data1, colWidths=[130, 370])
t1.setStyle(TableStyle([
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('TOPPADDING', (0,0), (-1,-1), 4),
    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0'))
]))
story.append(t1)

story.append(Spacer(1, 8))

# 3. Bangun Ruang
story.append(Paragraph("3. BANGUN RUANG (3D) & ILUSTRASI", h1_style))

def draw_cube():
    d = Drawing(120, 70)
    d.add(Polygon([25,8, 75,8, 75,42, 25,42], fillColor=colors.HexColor('#eef2ff'), strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Polygon([38,18, 88,18, 88,52, 38,52], fillColor=colors.HexColor('#c7d2fe'), strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Line(25,8, 38,18, strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Line(75,8, 88,18, strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Line(75,42, 88,52, strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(Line(25,42, 38,52, strokeColor=colors.HexColor('#4f46e5'), strokeWidth=1.5))
    d.add(String(50, 0, "sisi (s)", fontSize=7, fontName="Helvetica-Bold", textAnchor="middle", fillColor=colors.HexColor('#4338ca')))
    return d

def draw_cylinder():
    d = Drawing(120, 70)
    # top ellipse
    d.add(Ellipse(60, 55, 25, 8, fillColor=colors.HexColor('#fef3c7'), strokeColor=colors.HexColor('#f59e0b'), strokeWidth=1.5))
    # body rectangle (side view)
    d.add(Rect(35, 15, 50, 35, fillColor=colors.HexColor('#fef3c7'), strokeColor=colors.HexColor('#f59e0b'), strokeWidth=1.5))
    # bottom ellipse (hidden, dashed)
    d.add(Ellipse(60, 15, 25, 8, fillColor=None, strokeColor=colors.HexColor('#f59e0b'), strokeWidth=1.5, strokeDashArray=[2,2]))
    d.add(String(60, 2, "Tabung (r, t)", fontSize=7, fontName="Helvetica-Bold", textAnchor="middle", fillColor=colors.HexColor('#92400e')))
    return d

def draw_ball():
    d = Drawing(120, 70)
    d.add(Circle(60, 35, 25, fillColor=colors.HexColor('#faf5ff'), strokeColor=colors.HexColor('#7c3aed'), strokeWidth=1.5))
    d.add(Line(60, 35, 85, 35, strokeColor=colors.HexColor('#dc2626'), strokeWidth=1.5))
    d.add(String(72, 38, "r", fontSize=8, fontName="Helvetica-Bold", fillColor=colors.HexColor('#7c3aed')))
    return d

t_data2 = [
    [Image(r"C:/Users/MIP-USER/Documents/Project Soal/gambar kubus.jpg", width=70, height=50), Paragraph("<b>Kubus:</b> Volume = <code>s³</code>, L.Permukaan = <code>6s²</code>", body_style)],
    [Image(r"C:/Users/MIP-USER/Documents/Project Soal/gambar balok.jpg", width=70, height=50), Paragraph("<b>Balok:</b> Volume = <code>p×l×t</code>, L.Permukaan = <code>2(pl+pt+lt)</code>", body_style)],
    [Image(r"C:/Users/MIP-USER/Documents/Project Soal/gambar tabung.jpg", width=70, height=50), Paragraph("<b>Tabung:</b> Volume = <code>πr²t</code>, Selimut = <code>2πrt</code>", body_style)],
    [Image(r"C:/Users/MIP-USER/Documents/Project Soal/gambar kerucut.jpg", width=70, height=50), Paragraph("<b>Kerucut:</b> Volume = <code>⅓πr²t</code>", body_style)],
    [Image(r"C:/Users/MIP-USER/Documents/Project Soal/gambar bola.jpg", width=70, height=50), Paragraph("<b>Bola:</b> Volume = <code>⁴/₃πr³</code>, L.Permukaan = <code>4πr²</code>", body_style)]
]
t2 = Table(t_data2, colWidths=[130, 370])
t2.setStyle(TableStyle([
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f8fafc')),
    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ('TOPPADDING', (0,0), (-1,-1), 4),
    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0'))
]))
story.append(t2)

story.append(Spacer(1, 8))

# 4. Satuan & Trik Cepat
story.append(Paragraph("4. SATUAN UKUR & TRIK HITUNG CEPAT", h1_style))
story.append(Paragraph("• <b>Panjang:</b> 1 km = 1.000 m = 100.000 cm<br/>• <b>Luas:</b> 1 ha = 10.000 m² | 1 km² = 100 ha<br/>• <b>Volume & Kapasitas:</b> 1 m³ = 1.000 Liter | 1 dm³ = 1 Liter<br/>• <b>Perkalian 11:</b> <code>ab × 11 = a (a+b) b</code> (Contoh: 45 × 11 = 4 9 5 = 495)<br/>• <b>Kuadrat Berakhiran 5:</b> <code>(a5)² = a(a+1) | 25</code> (Contoh: 35² = 3×4 | 25 = 1225)", code_style))

doc.build(story)
print("PDF generated successfully with vector diagrams at:", pdf_path)
