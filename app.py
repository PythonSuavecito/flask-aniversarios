from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.units import inch
from datetime import datetime
import io
import os
import re
from werkzeug.utils import secure_filename

# Configuración de la aplicación
app = Flask(__name__)
app.config['SECRET_KEY'] = 'clave-secreta-para-flash-messages'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB máximo

# Asegurar que la carpeta de uploads existe
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Extensiones permitidas
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    """Verifica si el archivo tiene una extensión permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extraer_numero_aniversario(valor):
    """Extrae número de aniversario de cualquier formato"""
    if pd.isna(valor):
        return 0
    
    # Convertir a string
    valor_str = str(valor).strip()
    
    # Buscar números en el string
    numeros = re.findall(r'\d+', valor_str)
    
    if numeros:
        try:
            return int(numeros[0])
        except:
            return 0
    
    # Si no se encontraron números, intentar convertir directamente
    try:
        # Intentar convertir a número
        num = float(valor_str)
        return int(num)
    except:
        return 0

def procesar_archivo(file):
    """Procesa archivo de manera robusta"""
    filename = secure_filename(file.filename)
    extension = filename.rsplit('.', 1)[1].lower()
    
    # Leer archivo
    if extension in ['xlsx', 'xls']:
        df = pd.read_excel(file)
    elif extension == 'csv':
        # Intentar diferentes encodings
        for encoding in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
            try:
                file.seek(0)  # Reiniciar posición del archivo
                df = pd.read_csv(file, encoding=encoding)
                break
            except:
                continue
        else:
            raise ValueError("No se pudo leer el archivo CSV con ningún encoding")
    
    # Buscar columnas (insensible a mayúsculas)
    df_columns_upper = [col.upper() for col in df.columns]
    
    nombre_col = None
    aniversario_col = None
    
    # Buscar columna NOMBRE
    for col, col_upper in zip(df.columns, df_columns_upper):
        if 'NOMBRE' in col_upper or 'NAME' in col_upper:
            nombre_col = col
            break
    
    # Buscar columna ANIVERSARIO
    for col, col_upper in zip(df.columns, df_columns_upper):
        if 'ANIVERSARIO' in col_upper or 'ANNIVERSARY' in col_upper or 'AÑOS' in col_upper or 'YEARS' in col_upper:
            aniversario_col = col
            break
    
    if not nombre_col:
        # Si no encuentra, usar la primera columna
        if len(df.columns) >= 1:
            nombre_col = df.columns[0]
        else:
            raise ValueError("El archivo debe tener al menos una columna")
    
    if not aniversario_col:
        # Si no encuentra, usar la segunda columna o crear una con valores por defecto
        if len(df.columns) >= 2:
            aniversario_col = df.columns[1]
        else:
            # Crear columna por defecto si no existe
            df['ANIVERSARIO'] = 1
            aniversario_col = 'ANIVERSARIO'
    
    # Crear nuevo DataFrame con columnas estandarizadas
    df_procesado = pd.DataFrame()
    df_procesado['NOMBRE'] = df[nombre_col].fillna('').astype(str)
    df_procesado['ANIVERSARIO_NUM'] = df[aniversario_col].apply(extraer_numero_aniversario)
    
    # Filtrar filas válidas (con nombre y aniversario > 0)
    df_procesado = df_procesado[(df_procesado['NOMBRE'].str.strip() != '') & (df_procesado['ANIVERSARIO_NUM'] > 0)]
    
    return df_procesado

def crear_pdf_madrino(df, mes="DICIEMBRE", año="2025"):
    """Crea PDF EXACTAMENTE como genera_lista_junta.py - 4 columnas, letras grandes"""
    buffer = io.BytesIO()
    
    if df.empty:
        raise ValueError("No hay datos válidos para generar el PDF")
    
    total_festejados = len(df)
    
    # Configurar PDF
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    
    # Estilos
    estilo_encabezado = ParagraphStyle(
        'encabezado',
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=16,
        alignment=1,  # Centrado
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
    
    # ENCABEZADO
    encabezado_texto = f"ANIVERSARIO {mes.upper()} {año} - TOTAL: {total_festejados} FESTEJADOS"
    p_encabezado = Paragraph(encabezado_texto, estilo_encabezado)
    p_encabezado.wrapOn(c, width - 2*inch, 50)
    p_encabezado.drawOn(c, inch, height - 0.5*inch)
    
    # Organizar en 4 COLUMNAS (EXACTAMENTE como genera_lista_junta.py)
    x_positions = [30, width/4, width/2, 3*width/4]
    y_position = height - 80  # Empezar más abajo por el encabezado
    columna_actual = 0
    
    # Ordenar por años
    df_sorted = df.sort_values('ANIVERSARIO_NUM')
    
    for años, grupo in df_sorted.groupby('ANIVERSARIO_NUM'):
        nombres = grupo['NOMBRE'].tolist()
        
        # Título del año (en negrita)
        titulo = f"<b>{años} {'AÑO' if años == 1 else 'AÑOS'}</b>"
        p = Paragraph(titulo, estilo_titulo)
        p.wrapOn(c, width/4, 30)
        p.drawOn(c, x_positions[columna_actual], y_position)
        y_position -= 20
        
        # Nombres (uno por línea)
        for nombre in nombres:
            p = Paragraph(str(nombre), estilo_nombre)
            p.wrapOn(c, width/4, 20)
            p.drawOn(c, x_positions[columna_actual], y_position)
            y_position -= 15
        
        # Cambiar de columna si es necesario
        if y_position < 50:  # Margen inferior más grande
            columna_actual += 1
            y_position = height - 80
            if columna_actual > 3:  # Si se llenan las 4 columnas
                c.showPage()  # Nueva página
                # Repetir encabezado en nueva página
                p_encabezado.drawOn(c, inch, height - 0.5*inch)
                columna_actual = 0
                y_position = height - 80
    
    c.save()
    buffer.seek(0)
    return buffer, total_festejados

# RUTAS DE LA APLICACIÓN

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html')

@app.route('/subir', methods=['POST'])
def subir_archivo():
    """Procesa el archivo subido y genera el PDF"""
    if 'archivo' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    file = request.files['archivo']
    
    if file.filename == '':
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        flash('Formato de archivo no permitido. Use Excel (.xlsx, .xls) o CSV (.csv)', 'error')
        return redirect(url_for('index'))
    
    try:
        # Procesar archivo
        df = procesar_archivo(file)
        
        # Obtener parámetros del formulario
        mes = request.form.get('mes', 'MAYO')
        año = request.form.get('año', '2026')
        
        # Crear PDF
        pdf_buffer, total_festejados = crear_pdf_madrino(df, mes, año)
        
        # Nombre del archivo de salida
        nombre_archivo = f"aniversarios_{mes.lower()}_{año}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        # Enviar el archivo directamente
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name=nombre_archivo,
            mimetype='application/pdf'
        )
        
    except ValueError as e:
        flash(str(e), 'error')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error procesando el archivo: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/descargar_ejemplo')
def descargar_ejemplo():
    """Descarga un archivo de ejemplo"""
    # Crear un DataFrame de ejemplo
    data = {
        'NOMBRE': [
            'Juan Pérez', 'María García', 'Carlos López', 'Ana Martínez',
            'Luis Rodríguez', 'Laura Fernández', 'Pedro Sánchez', 'Carmen Torres',
            'Miguel Ruiz', 'Isabel Gómez', 'Francisco Díaz', 'Teresa Romero'
        ],
        'ANIVERSARIO': ['5 años', '10 años', '15 años', '5 años', 
                       '20 años', '10 años', '25 años', '15 años', 
                       '30 años', '35 años', '40 años', '45 años']
    }
    
    df = pd.DataFrame(data)
    
    # Guardar como Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Aniversarios')
    
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name='ejemplo_aniversarios.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/generar_muestra')
def generar_muestra():
    """Genera un PDF de muestra con datos de ejemplo"""
    try:
        # Datos de ejemplo
        data = {
            'NOMBRE': [
                'Juan Pérez', 'María García', 'Carlos López', 'Ana Martínez',
                'Luis Rodríguez', 'Laura Fernández', 'Pedro Sánchez', 'Carmen Torres'
            ],
            'ANIVERSARIO': ['5 años', '10 años', '15 años', '5 años', 
                           '20 años', '10 años', '25 años', '15 años']
        }
        
        df = pd.DataFrame(data)
        df['NOMBRE'] = df['NOMBRE'].astype(str)
        df['ANIVERSARIO_NUM'] = df['ANIVERSARIO'].apply(extraer_numero_aniversario)
        df = df[df['ANIVERSARIO_NUM'] > 0]
        
        pdf_buffer, total_festejados = crear_pdf_madrino(df, "MAYO", "2026")
        
        return send_file(
            pdf_buffer,
            as_attachment=True,
            download_name='aniversarios_muestra.pdf',
            mimetype='application/pdf'
        )
        
    except Exception as e:
        flash(f'Error generando PDF: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/ayuda')
def ayuda():
    """Página de ayuda"""
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Ayuda - Generador de Aniversarios</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #2c3e50; }
            .container { max-width: 800px; margin: 0 auto; }
            .section { margin: 30px 0; padding: 20px; background: #f8f9fa; border-radius: 10px; }
            code { background: #e9ecef; padding: 2px 6px; border-radius: 4px; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📚 Ayuda - Generador de Listas de Aniversarios</h1>
            
            <div class="section">
                <h2>📋 Formato del Archivo</h2>
                <p>El archivo debe contener al menos una columna con nombres y otra con los años de aniversario.</p>
                <p><strong>Formatos aceptados:</strong> Excel (.xlsx, .xls) o CSV (.csv)</p>
                
                <h3>Columnas reconocidas automáticamente:</h3>
                <ul>
                    <li><strong>Nombres:</strong> "NOMBRE", "NAME" (no sensible a mayúsculas)</li>
                    <li><strong>Aniversarios:</strong> "ANIVERSARIO", "ANNIVERSARY", "AÑOS", "YEARS"</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>🔢 Formatos de Aniversario Aceptados</h2>
                <p>El sistema extraerá automáticamente los números de estos formatos:</p>
                <ul>
                    <li><code>5 años</code> (texto con número)</li>
                    <li><code>5</code> (solo número)</li>
                    <li><code>5.0</code> (número decimal)</li>
                    <li><code>5 años y 6 meses</code> (se extraerá el 5)</li>
                    <li><code>10th anniversary</code> (se extraerá el 10)</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>🖨️ Formato del PDF Generado</h2>
                <p>El PDF se genera con las siguientes características:</p>
                <ul>
                    <li><strong>4 columnas</strong> organizadas verticalmente</li>
                    <li><strong>Letras grandes</strong> (títulos 12pt, nombres 11pt)</li>
                    <li><strong>Encabezado personalizado</strong> con mes, año y total de festejados</li>
                    <li><strong>Páginas múltiples automáticas</strong> cuando se llenan las 4 columnas</li>
                    <li><strong>Organización por años de aniversario</strong> (de menor a mayor)</li>
                </ul>
            </div>
            
            <div class="section">
                <h2>🚀 Cómo Usar</h2>
                <ol>
                    <li>Prepara tu archivo Excel o CSV con los datos</li>
                    <li>Selecciona el mes y año para el encabezado</li>
                    <li>Haz clic en "Generar PDF"</li>
                    <li>Descarga el PDF generado</li>
                </ol>
                <p><a href="/">← Volver al generador</a></p>
            </div>
        </div>
    </body>
    </html>
    """

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)