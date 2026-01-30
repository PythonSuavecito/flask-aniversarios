from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.units import inch
from fpdf import FPDF
from datetime import datetime
import io
import os
import tempfile


def format_currency(value):
    """Formatea números como moneda (ej: 1000 → '1 000')."""
    try:
        num = float(value)
        return f"{num:,.0f}".replace(",", " ") + " "
    except:
        return "NULO "

class PDF(FPDF):
    def __init__(self):
        super().__init__(orientation='L')
        self.set_auto_page_break(auto=True, margin=15)
        self.set_font('Arial', '', 9)
        self.total_general = 0
    
    def header(self):
        self.set_font('Arial', 'B', 10)
        self.cell(0, 5, "CONGRESO 2025 - RESUMEN DE BONOS", 0, 1, 'C')
        self.cell(0, 5, datetime.now().strftime('%d/%m/%Y'), 0, 1, 'C')
        self.ln(5)
        
        headers = ["GRUPO", "GUIA", "BON1", "BON2", "BON3", "BON4", "BON5", 
                   "MONT1", "MONT2", "MONT3", "MONT4", "MONT5", "ASIST.", "TOTAL"]
        widths = [30, 20] + [12]*5 + [18]*5 + [12, 18]
        
        for header, width in zip(headers, widths):
            self.cell(width, 7, header, border=1, align='C')
        self.ln()
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')
    
    def clean_text(self, text, max_width):
        """Ajusta el texto al ancho de celda."""
        original_text = str(text)
        self.set_font('Arial', '', 8)
        
        if self.get_string_width(original_text) <= max_width:
            return original_text
        
        for size in [7, 6]:
            self.set_font('Arial', '', size)
            if self.get_string_width(original_text) <= max_width:
                return original_text
        
        self.set_font('Arial', '', 7)
        parts = original_text.split()
        shortened = ""
        
        for part in parts:
            if self.get_string_width(shortened + " " + part) <= max_width - 5:
                shortened += (" " if shortened else "") + part
            else:
                break
        
        if not shortened:
            shortened = original_text[:int(max_width/3)]
        
        return shortened + ".." if len(shortened) < len(original_text) else shortened
    
    def add_table_row(self, grupo, guia, bonos, montos, asistentes, total, is_continuation=False):
        """Añade una fila a la tabla, manejando continuaciones."""
        self.set_font('Arial', '', 8)
        
        if is_continuation:
            row_data = ["", ""] + bonos + montos + ["", ""]
        else:
            row_data = [
                self.clean_text(grupo, 30),
                self.clean_text(guia, 20)
            ] + bonos + montos + [
                "{:.2f}".format(asistentes).replace(".", ","),
                format_currency(total)
            ]
        
        col_widths = [30, 20] + [12]*5 + [18]*5 + [12, 18]
        
        for data, width in zip(row_data, col_widths):
            self.cell(width, 6, str(data), border=1)
        self.ln()
    
    def add_total_general(self):
        """Añade el total general al final del PDF."""
        self.ln(10)
        self.set_font('Arial', 'B', 10)
        total_formatted = f"{self.total_general:,.0f}".replace(",", " ") + " "
        self.cell(0, 8, f"TOTAL GENERAL: {total_formatted}", 0, 1, 'R')

def crear_pdf_aniversarios(df):
    """Crea PDF de aniversarios similar a genera_lista_junta.py"""
    buffer = io.BytesIO()
    
    # Configurar PDF
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # Estilos
    estilo_encabezado = ParagraphStyle(
        'encabezado',
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=16,
        alignment=1,
        spaceAfter=12
    )
    
    estilo_titulo = ParagraphStyle(
        'titulo',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=14,
        spaceAfter=6
    )
    
    estilo_nombre = ParagraphStyle(
        'nombre',
        fontName='Helvetica',
        fontSize=11,
        leading=12,
        leftIndent=15
    )

    # Procesar datos
    df['ANIVERSARIO'] = df['ANIVERSARIO'].str.extract(r'(\d+)').astype(int)
    total_festejados = len(df)
    
    # Encabezado
    encabezado_texto = f"ANIVERSARIO MAYO 2026 - TOTAL: {total_festejados} FESTEJADOS"
    p_encabezado = Paragraph(encabezado_texto, estilo_encabezado)
    p_encabezado.wrapOn(c, width - 2*inch, 50)
    p_encabezado.drawOn(c, inch, height - 0.5*inch)
    
    # Organizar en columnas
    x_positions = [30, width/4, width/2, 3*width/4]
    y_position = height - 80
    columna_actual = 0

    for años, nombres in sorted(df.groupby('ANIVERSARIO')['NOMBRE']):
        # Título del año
        titulo = f"<b>{años} {'AÑO' if años == 1 else 'AÑOS'}</b>"
        p = Paragraph(titulo, estilo_titulo)
        p.wrapOn(c, width/4, 30)
        p.drawOn(c, x_positions[columna_actual], y_position)
        y_position -= 20
        
        # Nombres
        for nombre in nombres:
            p = Paragraph(nombre, estilo_nombre)
            p.wrapOn(c, width/4, 20)
            p.drawOn(c, x_positions[columna_actual], y_position)
            y_position -= 15
        
        # Cambiar de columna si es necesario
        if y_position < 50:
            columna_actual += 1
            y_position = height - 80
            if columna_actual > 3:
                c.showPage()
                p_encabezado.drawOn(c, inch, height - 0.5*inch)
                columna_actual = 0
                y_position = height - 80

    c.save()
    buffer.seek(0)
    return buffer


