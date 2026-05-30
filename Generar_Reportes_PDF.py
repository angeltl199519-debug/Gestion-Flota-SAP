from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.units import inch
from datetime import datetime

def exportar_a_pdf(nombre_archivo_pdf="Reporte_Mantenimiento.pdf"):
    """Genera un reporte PDF profesional"""
    
    doc = SimpleDocTemplate(nombre_archivo_pdf, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    titulo = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1072BA'),
        spaceAfter=30,
        alignment=1
    )
    
    story.append(Paragraph("🚛 REPORTE DE GESTIÓN DE MANTENIMIENTO DE FLOTA", titulo))
    story.append(Spacer(1, 0.3*inch))
    
    info_empresa = ParagraphStyle(
        'Info',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.grey
    )
    
    fecha_texto = f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    story.append(Paragraph(fecha_texto, info_empresa))
    story.append(Paragraph("Empresa: Transporte XYZ S.A.", info_empresa))
    story.append(Spacer(1, 0.3*inch))
    
    # Tabla de KPIs
    kpi_style = ParagraphStyle(
        'KPI',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=colors.HexColor('#1072BA'),
        spaceAfter=12
    )
    
    story.append(Paragraph("INDICADORES PRINCIPALES", kpi_style))
    
    kpi_data = [
        ["Métrica", "Valor", "Estado"],
        ["Total de Unidades", "40", "✓ Activo"],
        ["Órdenes Completadas", "38", "✓ OK"],
        ["Costo Total", "$125,450", "📊 Monitor"],
        ["Repuestos Bajo Stock", "2", "⚠️ Crítico"],
    ]
    
    kpi_table = Table(kpi_data, colWidths=[2.5*inch, 1.5*inch, 1.5*inch])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1072BA')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F0F0F0')]),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Resumen de Órdenes
    story.append(Paragraph("RESUMEN DE ÓRDENES RECIENTES", kpi_style))
    
    ordenes_data = [
        ["Nº Orden", "Unidad", "Tipo", "Estado", "Costo"],
        ["ORD-2024001", "UNIT-001", "Preventivo", "Completado", "$1,200"],
        ["ORD-2024002", "UNIT-005", "Correctivo", "Activo", "$2,500"],
        ["ORD-2024003", "UNIT-012", "Inspección", "Pendiente", "$800"],
    ]
    
    ordenes_table = Table(ordenes_data, colWidths=[1.3*inch, 1*inch, 1*inch, 1*inch, 1*inch])
    ordenes_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#70AD47')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
    ]))
    story.append(ordenes_table)
    story.append(Spacer(1, 0.5*inch))
    
    # Pie
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.grey,
        alignment=1
    )
    story.append(Paragraph("Documento Generado - Sistema de Gestión de Mantenimiento v2.0", footer_style))
    
    doc.build(story)
    print(f"✅ PDF generado: {nombre_archivo_pdf}")

if __name__ == "__main__":
    exportar_a_pdf("Reporte_Mantenimiento_Profesional.pdf")