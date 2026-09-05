import os
import sys
import docx
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfgen import canvas

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super(NumberedCanvas, self).__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_header_footer(num_pages)
            super(NumberedCanvas, self).showPage()
        super(NumberedCanvas, self).save()

    def draw_header_footer(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 8)
        self.setFillColor(colors.HexColor("#64748B"))
        
        # Header (Top line & logo/text)
        self.setStrokeColor(colors.HexColor("#CBD5E1"))
        self.setLineWidth(0.5)
        self.line(36, 756, 576, 756)
        
        self.drawString(36, 762, "FACTURANEX SAS — NIT 902.049.529-8")
        self.drawRightString(576, 762, "Documento Legal Oficial")

        # Footer (Bottom line & Page Number)
        self.line(36, 45, 576, 45)
        self.drawString(36, 32, "© 2026 FacturaNex SAS — Todos los derechos reservados")
        page_str = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(576, 32, page_str)
        self.restoreState()

def docx_to_pdf(docx_path, pdf_path, document_title):
    print(f"Converting {docx_path} -> {pdf_path}...")
    doc = docx.Document(docx_path)
    
    pdf = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=36,
        rightMargin=36,
        topMargin=54,
        bottomMargin=54
    )
    
    story = []
    styles = getSampleStyleSheet()
    
    COLOR_PRIMARY = colors.HexColor('#D81B60')
    COLOR_NAVY = colors.HexColor('#0A1128')
    COLOR_TEXT = colors.HexColor('#1E293B')
    COLOR_MUTED = colors.HexColor('#64748B')
    
    title_style = ParagraphStyle(
        'DocTitle',
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=19,
        textColor=COLOR_PRIMARY,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    
    subtitle_style = ParagraphStyle(
        'DocSub',
        fontName='Helvetica-Bold',
        fontSize=10,
        leading=14,
        textColor=COLOR_NAVY,
        alignment=TA_CENTER,
        spaceAfter=12
    )

    h1_style = ParagraphStyle(
        'Heading1_Legal',
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=COLOR_NAVY,
        spaceBefore=10,
        spaceAfter=4,
        keepWithNext=True
    )

    body_style = ParagraphStyle(
        'Body_Legal',
        fontName='Helvetica',
        fontSize=9,
        leading=12.5,
        textColor=COLOR_TEXT,
        alignment=TA_JUSTIFY,
        spaceAfter=5
    )

    bullet_style = ParagraphStyle(
        'Bullet_Legal',
        fontName='Helvetica',
        fontSize=9,
        leading=12.5,
        textColor=COLOR_TEXT,
        alignment=TA_LEFT,
        leftIndent=15,
        spaceAfter=4
    )

    # Logo Header Table
    logo_file = os.path.abspath('img/logo_facturanex_v3.png')
    if os.path.exists(logo_file):
        img_logo = Image(logo_file, width=52, height=52)
    else:
        img_logo = Paragraph("<b>FacturaNex SAS</b>", title_style)

    head_right = Paragraph("<font color='#0A1128' size=8.5><b>FACTURANEX SAS</b></font><br/><font color='#64748B' size=8>NIT 902.049.529-8<br/>contacto@facturanex.com | facturanex.com</font>", ParagraphStyle('R1', alignment=TA_RIGHT))
    header_table = Table([[img_logo, head_right]], colWidths=[100, 440])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

    story.append(header_table)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1, color=COLOR_PRIMARY, spaceBefore=2, spaceAfter=12))

    # Title
    story.append(Paragraph(document_title.upper(), title_style))
    story.append(Paragraph("FACTURANEX SAS — NIT 902.049.529-8", subtitle_style))
    story.append(Spacer(1, 6))

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        
        # Determine styling based on content format
        if text.startswith(("1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9.", "10.", "11.", "12.", "13.", "14.", "15.")) or text.isupper():
            if len(text) < 120 and not text.endswith("."):
                story.append(Paragraph(text, h1_style))
                continue
        
        if text.startswith("•") or text.startswith("-"):
            story.append(Paragraph(text, bullet_style))
        else:
            # Escape HTML characters in text except if it's plain
            clean_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(clean_text, body_style))

    pdf.build(story, canvasmaker=NumberedCanvas)
    print(f"Finished building {pdf_path}")

if __name__ == "__main__":
    docx_to_pdf(
        "POLÍTICA DE PRIVACIDAD Y TRATAMIENTO DE DATOS PERSONALES.docx",
        "politica_de_privacidad_facturanex.pdf",
        "Política de Privacidad y Tratamiento de Datos Personales"
    )
    docx_to_pdf(
        "TÉRMINOS Y CONDICIONES DE USO Y PRESTACIÓN DEL SERVICIO.docx",
        "terminos_y_condiciones_facturanex.pdf",
        "Términos y Condiciones de Uso y Prestación del Servicio"
    )
    docx_to_pdf(
        "HABEAS DATA.docx",
        "habeas_data_facturanex.pdf",
        "Procedimiento para el Ejercicio del Derecho de Habeas Data"
    )
