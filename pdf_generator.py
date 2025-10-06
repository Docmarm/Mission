"""
Module pour la génération de PDF professionnels pour les procès-verbaux
Utilise reportlab pour créer des PDF structurés et formatés
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from datetime import datetime
import io

def create_pv_pdf(content, title="Procès-verbal de Mission", author="Responsable Mission"):
    """
    Génère un PDF professionnel pour un procès-verbal
    
    Args:
        content (str): Contenu du procès-verbal en texte
        title (str): Titre du document
        author (str): Nom de l'auteur/responsable
    
    Returns:
        bytes: Contenu du PDF généré
    """
    
    # Créer un buffer en mémoire
    buffer = io.BytesIO()
    
    # Configuration du document
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Style pour le titre principal
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    # Style pour les en-têtes de section
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=12,
        spaceBefore=20,
        fontName='Helvetica-Bold',
        textColor=colors.black
    )
    
    # Style pour le texte normal
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=6,
        alignment=TA_JUSTIFY,
        fontName='Helvetica'
    )
    
    # Style pour la signature
    signature_style = ParagraphStyle(
        'Signature',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        spaceAfter=6
    )
    
    # Contenu du document
    story = []
    
    # En-tête du document
    story.append(Paragraph(title.upper(), title_style))
    story.append(Spacer(1, 12))
    
    # Date de génération
    date_str = f"Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}"
    story.append(Paragraph(date_str, signature_style))
    story.append(Spacer(1, 20))
    
    # Traitement du contenu
    lines = content.split('\n')
    current_section = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            story.append(Spacer(1, 6))
            continue
            
        # Détection des titres (commencent par I., II., III., etc. ou A., B., C.)
        if (line.startswith(('I.', 'II.', 'III.', 'IV.', 'V.', 'VI.')) or 
            line.startswith(('A.', 'B.', 'C.', 'D.', 'E.', 'F.'))):
            story.append(Paragraph(line, heading_style))
        # Détection des sous-titres (numérotés 1., 2., 3.)
        elif line.startswith(('1.', '2.', '3.', '4.', '5.')):
            sub_heading_style = ParagraphStyle(
                'SubHeading',
                parent=normal_style,
                fontSize=12,
                fontName='Helvetica-Bold',
                spaceBefore=10,
                spaceAfter=6
            )
            story.append(Paragraph(line, sub_heading_style))
        # Texte normal
        else:
            story.append(Paragraph(line, normal_style))
    
    # Signature
    story.append(Spacer(1, 30))
    story.append(Paragraph(f"Fait à Dakar, le {datetime.now().strftime('%d/%m/%Y')}", signature_style))
    story.append(Spacer(1, 40))
    story.append(Paragraph("_" * 30, signature_style))
    story.append(Paragraph(f"<b>{author}</b>", signature_style))
    
    # Construction du PDF
    doc.build(story)
    
    # Récupération du contenu
    pdf_data = buffer.getvalue()
    buffer.close()
    
    return pdf_data

def create_word_document(content, title="Procès-verbal de Mission"):
    """
    Génère un document Word simple (format RTF) pour compatibilité
    
    Args:
        content (str): Contenu du document
        title (str): Titre du document
    
    Returns:
        str: Contenu RTF du document
    """
    
    rtf_content = f"""{{\\rtf1\\ansi\\deff0
{{\\fonttbl{{\\f0 Times New Roman;}}}}
{{\\colortbl;\\red0\\green0\\blue0;}}
\\f0\\fs24
\\qc\\b {title.upper()}\\b0\\par
\\par
\\qc Généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}\\par
\\par
\\ql
{content.replace(chr(10), '\\par ')}
\\par
\\par
\\qc Fait à Dakar, le {datetime.now().strftime('%d/%m/%Y')}\\par
\\par
\\par
\\qc ________________________________\\par
\\qc\\b Responsable Mission\\b0\\par
}}"""
    
    return rtf_content