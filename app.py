import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
import os
from io import BytesIO
import requests

# =============== CONFIGURATION ===============
EXCEL_FILE = "phd_candidates.xlsx"  # your Excel file path
OUTPUT_FOLDER = "halltickets_pdf"   # output folder for PDFs
LOGO_URL = "https://mrdu.edu.in/wp-content/uploads/2025/08/Logo.png"
# ============================================

# Ensure output directory exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def download_logo(url):
    """Download logo and return as BytesIO object"""
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return BytesIO(response.content)
    except Exception as e:
        print(f"⚠️  Could not download logo: {e}")
    return None

def create_hall_ticket(name, phone, email, branch, reg_no, pdf_path, logo_data):
    """Generate a single hall ticket PDF"""
    
    # Create PDF
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=15*mm,
        leftMargin=15*mm,
        topMargin=15*mm,
        bottomMargin=15*mm
    )
    
    # Container for the 'Flowable' objects
    elements = []
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#721c24'),
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#721c24'),
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=9,
        alignment=TA_CENTER
    )
    
    # Header section with logo and university info
    header_data = []
    
    if logo_data:
        try:
            logo = Image(logo_data, width=60, height=60)
            header_info = [
                [logo, Paragraph("<b>MALLA REDDY (MR) DEEMED TO BE UNIVERSITY</b><br/>"
                                "Maisammaguda, Dhulapally (Post), Hyderabad – 500100, Telangana, India<br/>"
                                "Website: www.mrdu.edu.in | Email: mrdeemeduniversity@gmail.com | Phone: 9348161125", 
                                header_style)]
            ]
        except:
            header_info = [[Paragraph("<b>MALLA REDDY (MR) DEEMED TO BE UNIVERSITY</b><br/>"
                                     "Maisammaguda, Dhulapally (Post), Hyderabad – 500100, Telangana, India<br/>"
                                     "Website: www.mrdu.edu.in | Email: mrdeemeduniversity@gmail.com | Phone: 9348161125", 
                                     header_style)]]
    else:
        header_info = [[Paragraph("<b>MALLA REDDY (MR) DEEMED TO BE UNIVERSITY</b><br/>"
                                 "Maisammaguda, Dhulapally (Post), Hyderabad – 500100, Telangana, India<br/>"
                                 "Website: www.mrdu.edu.in | Email: mrdeemeduniversity@gmail.com | Phone: 9348161125", 
                                 header_style)]]
    
    header_table = Table(header_info, colWidths=[70, 450] if logo_data else [520])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#fce5e5')),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 12),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#e9ecef')),
    ]))
    
    elements.append(header_table)
    elements.append(Spacer(1, 10*mm))
    
    # Title
    elements.append(Paragraph("MRDUCET - PH.D 2025 HALLTICKET", title_style))
    elements.append(Spacer(1, 8*mm))
    
    # Candidate Details Table
    candidate_data = [
        ['Name of the Candidate', name],
        ['Registration No', reg_no],
        ['Phone No', phone],
        ['Branch', branch]
    ]
    
    candidate_table = Table(candidate_data, colWidths=[150, 250])
    candidate_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f5ff')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.black),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#dee2e6')),
        ('TEXTCOLOR', (1, 1), (1, 1), colors.HexColor('#d42229')),
        ('FONTNAME', (1, 1), (1, 1), 'Helvetica-Bold'),
    ]))
    
    elements.append(candidate_table)
    elements.append(Spacer(1, 8*mm))
    
    # Schedule Section
    schedule_title = Paragraph("<b>Schedule:</b>", styles['Heading3'])
    elements.append(schedule_title)
    elements.append(Spacer(1, 3*mm))
    
    schedule_data = [
        ['Sl.No', 'Details', 'Date', 'Timings'],
        ['1', 
         'Written Examination and Interview\n\n'
         '• Written Exam (100 Marks)\n'
         '  - Research Methodology: 50 Marks\n'
         '  - Subjective Questions: 50 Marks\n\n'
         '• Interview (30 Marks) for qualified candidates\n\n'
         '• Final Selection: Written Exam (scaled to 70) + Interview (30)',
         '18-10-2025',
         'Exam:\n10:00 AM – 11:30 AM\n\nInterview:\n1:00 PM – 5:30 PM']
    ]
    
    schedule_table = Table(schedule_data, colWidths=[40, 270, 80, 130])
    schedule_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#fce5e5')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#721c24')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (2, 1), 'CENTER'),
        ('ALIGN', (3, 1), (3, 1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('PADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#dee2e6')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f9f9f9')),
        ('TEXTCOLOR', (2, 1), (2, 1), colors.HexColor('#d42229')),
        ('FONTNAME', (2, 1), (2, 1), 'Helvetica-Bold'),
    ]))
    
    elements.append(schedule_table)
    elements.append(Spacer(1, 8*mm))
    
    # Important Instructions
    instructions_text = """<b>Important Instructions</b><br/><br/>
    <b>Reporting Time:</b> Please report by <b>9:00 AM</b> on 18-10-2025 for verification and registration.<br/><br/>
    <b>Documents to Carry:</b><br/>
    • Copy of this invitation letter (printout)<br/>
    • Government-issued photo ID (Aadhaar, Passport, etc.)<br/>
    • Originals and copies of academic certificates (10th, 12th, UG, PG)<br/>
    • Two (2) recent passport-size photographs
    """
    
    instructions_para = Paragraph(instructions_text, styles['Normal'])
    instructions_table = Table([[instructions_para]], colWidths=[520])
    instructions_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#fffaf0')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#fde2e2')),
        ('PADDING', (0, 0), (-1, -1), 12),
    ]))
    
    elements.append(instructions_table)
    elements.append(Spacer(1, 10*mm))
    
    # Footer
    footer_text = """Warm regards,<br/>
    <b>Registrar</b><br/>
    Malla Reddy (MR) Deemed to be University<br/><br/>
    <font size="8">© 2025 Malla Reddy (MR) Deemed to be University. All Rights Reserved.</font>
    """
    elements.append(Paragraph(footer_text, styles['Normal']))
    
    # Build PDF
    doc.build(elements)

# Main execution
print("🚀 Starting Hall Ticket Generation...\n")

# Download logo once
print("📥 Downloading logo...")
logo_data = download_logo(LOGO_URL)

# Load Excel file
print("📊 Loading Excel file...")
df = pd.read_excel(EXCEL_FILE)

# Generate PDFs
print(f"📝 Generating {len(df)} hall tickets...\n")

for i, row in df.iterrows():
    name = str(row.iloc[1]).strip()
    phone = str(row.iloc[2]).strip()
    email = str(row.iloc[3]).strip()
    branch = str(row.iloc[4]).strip()
    reg_no = str(row.iloc[5]).strip()

    # Create safe filename
    safe_name = "".join([c for c in name if c.isalnum() or c in (" ", "_")]).rstrip()
    pdf_filename = os.path.join(OUTPUT_FOLDER, f"{safe_name}_{reg_no}.pdf")

    # Generate PDF
    create_hall_ticket(name, phone, email, branch, reg_no, pdf_filename, logo_data)
    print(f"✅ Generated: {pdf_filename}")

print(f"\n🎉 All {len(df)} hall tickets generated successfully!")
print(f"📁 Output folder: {OUTPUT_FOLDER}")