import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image as RLImage

def generate_sample():
    # Setup output file path
    output_pdf_path = os.path.join(os.path.dirname(__file__), 'sample_offer.pdf')
    
    # Document setup
    doc = SimpleDocTemplate(output_pdf_path, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=24, bottomMargin=24)
    story = []
    
    # ReportLab Styles
    styles = getSampleStyleSheet()
    
    # Brand Colors
    blue_brand = colors.HexColor('#0c5ca8')
    green_brand = colors.HexColor('#3fb549')
    dark_gray = colors.HexColor('#2c2c2a')
    
    # Custom Styles
    brand_style1 = ParagraphStyle('Brand1', fontName='Helvetica-Bold', fontSize=26, leading=30, textColor=blue_brand, alignment=1)
    brand_style2 = ParagraphStyle('Brand2', fontName='Helvetica-Bold', fontSize=14, leading=16, textColor=green_brand, alignment=1)
    address_style = ParagraphStyle('Address', fontName='Helvetica-Bold', fontSize=9.5, leading=11.5, textColor=dark_gray, alignment=1)
    header_contact_style = ParagraphStyle('HeaderContact', fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=dark_gray, alignment=1)
    
    title_style = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=14, leading=16, textColor=colors.black, alignment=1, spaceAfter=4)
    intro_style = ParagraphStyle('Intro', fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=colors.HexColor('#1D9E75'), alignment=1, spaceBefore=2, spaceAfter=4)
    
    cell_style = ParagraphStyle('Cell', fontName='Helvetica', fontSize=9, leading=11, textColor=colors.black, alignment=1)
    cell_bold_style = ParagraphStyle('CellBold', fontName='Helvetica-Bold', fontSize=9.5, leading=11.5, textColor=colors.black, alignment=1)
    
    left_cell_style = ParagraphStyle('LeftCell', fontName='Helvetica', fontSize=9, leading=11, textColor=colors.black, alignment=0)
    left_cell_bold_style = ParagraphStyle('LeftCellBold', fontName='Helvetica-Bold', fontSize=9.5, leading=11.5, textColor=colors.black, alignment=0)
    
    terms_title_style = ParagraphStyle('TermsTitle', fontName='Helvetica-Bold', fontSize=10.5, leading=12.5, textColor=colors.black, spaceBefore=6, spaceAfter=3)
    terms_item_style = ParagraphStyle('TermsItem', fontName='Helvetica', fontSize=9, leading=12, textColor=colors.black, spaceAfter=2)
    
    footer_text_style = ParagraphStyle('FooterText', fontName='Helvetica', fontSize=9, leading=12, textColor=colors.black, alignment=0)
    
    # 1. Noble Air Gases Header Layout
    logo_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'static', 'img', 'noble_logo.png'))
    
    if os.path.exists(logo_path):
        logo_img = RLImage(logo_path, width=240, height=60, kind='proportional')
        logo_img.hAlign = 'CENTER'
        story.append(logo_img)
        story.append(Spacer(1, 4))
    else:
        # Fallback to text
        story.append(Paragraph("NOBLE", brand_style1))
        story.append(Paragraph("air gases", brand_style2))
        story.append(Spacer(1, 2))
        
    story.append(Paragraph("Plot No. A/12, MIDC Waluj, Chhatrapati Sambhajinagar", address_style))
    story.append(Spacer(1, 2))
    story.append(Paragraph("Email: sales@nobleairgases.com &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mobile: +91 9225309555", header_contact_style))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#d8d9d4'), spaceAfter=4))
    
    # 2. Document Title
    story.append(Paragraph("COMMERCIAL OFFER", title_style))
    
    # 3. Metadata block table
    customer_name = "Aditya Gases Pvt Ltd"
    attn = "Mr. R. K. Sharma (Purchase Manager)"
    tel = "+91 9876543210"
    q_date = "16th June 2026"
    q_no = "NAG/26-27/3482"
    ref = "Your enquiry email dated 15-06-2026"
    
    meta_data = [
        [
            Paragraph("<b>To</b>", left_cell_bold_style),
            Paragraph(f": M/s {customer_name}", left_cell_style),
            Paragraph("<b>K.Attn</b>", left_cell_bold_style),
            Paragraph(f": {attn}", left_cell_style)
        ],
        [
            Paragraph("<b>Tel</b>", left_cell_bold_style),
            Paragraph(f": {tel}", left_cell_style),
            Paragraph("<b>Date</b>", left_cell_bold_style),
            Paragraph(f": {q_date}", left_cell_style)
        ],
        [
            Paragraph("<b>Q. No.</b>", left_cell_bold_style),
            Paragraph(f": {q_no}", left_cell_style),
            Paragraph("<b>Ref</b>", left_cell_bold_style),
            Paragraph(f": {ref}", left_cell_style)
        ]
    ]
    meta_table = Table(meta_data, colWidths=[50, 220, 50, 220])
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor('#888888')),
    ]))
    
    story.append(meta_table)
    story.append(Spacer(1, 6))
    
    # 4. Intro text
    story.append(Paragraph("Thank you for your interest in our products & services. We are pleased to offer our most Competitive quote for your consideration with regards to your requirements", intro_style))
    
    # 5. Quote product table
    product_names = [
        "Argon (5.5 Grade)", "Argon (5.0 Grade)", "Helium (5.0 Grade)",
        "Nitrogen (5.0 Grade)", "Hydrogen (5.0 Grade)", "SF6 Gas"
    ]
    product_contents = ["07 Cum", "07 Cum", "07 Cum", "07 Cum", "07 Cum", "50 Kg"]
    product_prices = ["630", "390", "5700", "135", "250", "1050"]
    product_units = ["Per Cum", "Per Cum", "Per Cum", "Per Cum", "Per Cum", "Per Kg"]
    
    table_data = [
        [
            Paragraph("<b>No.</b>", cell_bold_style),
            Paragraph("<b>Product</b>", cell_bold_style),
            Paragraph("<b>Content per Cylinder</b>", cell_bold_style),
            Paragraph("<b>Price<br/>(Rs/unit)</b>", cell_bold_style),
            Paragraph("<b>Unit</b>", cell_bold_style)
        ]
    ]
    
    idx_no = 1
    for name, content, price, unit in zip(product_names, product_contents, product_prices, product_units):
        table_data.append([
            Paragraph(f"{idx_no:02d}", cell_style),
            Paragraph(f"<b>{name}</b>", left_cell_bold_style),
            Paragraph(content, cell_style),
            Paragraph(f"{price}/-", cell_style),
            Paragraph(unit, cell_style)
        ])
        idx_no += 1
        
    prod_table = Table(table_data, colWidths=[40, 200, 120, 90, 90])
    prod_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#d8d9d4')),
        ('LINEBELOW', (0,0), (-1,0), 1.5, colors.black),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('BACKGROUND', (0,0), (-1,0), colors.white),
    ]))
    story.append(prod_table)
    story.append(Spacer(1, 6))
    
    # 6. Terms and conditions
    story.append(Paragraph("<b>TERMS &amp; CONDITIONS:</b>", terms_title_style))
    terms_pairs = [
        ("1. Prices", "As Quoted"),
        ("2. GST", "@18% Extra."),
        ("3. Transportation", "Inclusive Delivery."),
        ("4. Payment", "30 Days"),
        ("5. Valve Damage", "Rs. 1000/- per Valve"),
        ("6. Cylinder Lost / Damage", "Rs.10,000/- Per Cylinder")
    ]
    for label, value in terms_pairs:
        story.append(Paragraph(f"<b>{label}</b> : {value}", terms_item_style))
        
    story.append(Spacer(1, 6))
    
    # 7. Footer text
    story.append(Paragraph("For any further queries, please feel free to contact us. We value your business association.", footer_text_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Thanking you,", footer_text_style))
    story.append(Spacer(1, 8))
    
    # Signature line
    sig_data = [
        [
            Paragraph("<b>For Noble Air Gases</b>", left_cell_bold_style),
            Paragraph("", cell_style)
        ],
        [
            Paragraph("<br/>Authorized Signatory", left_cell_style),
            Paragraph("", cell_style)
        ]
    ]
    sig_table = Table(sig_data, colWidths=[270, 270])
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(sig_table)
    
    doc.build(story)
    print("Sample PDF generated successfully at:", output_pdf_path)

if __name__ == '__main__':
    generate_sample()
