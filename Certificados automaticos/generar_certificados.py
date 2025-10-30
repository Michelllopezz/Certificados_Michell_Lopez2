import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import os
import re
from datetime import datetime

# RUTAS
ruta_excel = "datos/datos (2).xlsx"
ruta_plantilla = "plantilla/Blue Minimalist Certificate Of Achievement (3).png"
carpeta_salida = "certificados"

os.makedirs(carpeta_salida, exist_ok=True)

# Leer Excel
df = pd.read_excel(ruta_excel)

# Función para formatear fecha
def formatear_fecha(valor_fecha=None):
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    if pd.notna(valor_fecha):
        try:
            dt = pd.to_datetime(valor_fecha)
            return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"
        except Exception:
            return str(valor_fecha)
    hoy = datetime.today()
    return f"{hoy.day} de {meses[hoy.month - 1]} de {hoy.year}"

# Función para limpiar nombres de archivo
def sanitizar_nombre(s):
    s = str(s).strip()
    s = re.sub(r'[\\/:"*?<>|]+', "", s)
    s = s.replace(" ", "_")
    return s

# Función que genera el PDF
def generar_certificado(nombre, codigo, proyecto, espacio, fecha_texto):
    if pd.isna(nombre) or nombre == "":
        return  # si no hay nombre, no genera nada

    nombre_file = sanitizar_nombre(nombre)
    ruta_salida = os.path.join(carpeta_salida, f"Certificado_{nombre_file}.pdf")

    c = canvas.Canvas(ruta_salida, pagesize=A4)
    width, height = A4

    plantilla = ImageReader(ruta_plantilla)
    c.drawImage(plantilla, 0, 0, width=width, height=height)

    # Texto
    c.setFont("Helvetica-Bold", 26)
    c.drawCentredString(width/2, 420, str(nombre))

    c.setFont("Helvetica", 12)
    c.drawCentredString(width/2, 385, f"Código: {codigo}")
    c.drawCentredString(width/2, 360, f"Proyecto: {proyecto}")
    c.drawCentredString(width/2, 335, f"Espacio académico: {espacio}")

    c.setFont("Helvetica-Oblique", 11)
    c.drawCentredString(width/2, 310, f"Fecha: {fecha_texto}")

    c.save()
    print(f"✅ Certificado generado: {ruta_salida}")

# Recorrer filas y generar un certificado por cada estudiante
for _, fila in df.iterrows():
    espacio = fila["Selecciona el espacio académico"]
    proyecto = fila["Nombre del Proyecto"]
    fecha_val = fila.get("Fecha", None)
    fecha_texto = formatear_fecha(fecha_val)

    # Recorre todos los posibles estudiantes (1, 2, 3, 4, etc.)
    for i in range(1, 10):  # aumenta el 10 si tienes más columnas
        col_nombre = f"Nombre completo del estudiante {i}"
        col_codigo = f"Código del estudiante {i}"

        if col_nombre in df.columns and col_codigo in df.columns:
            nombre = fila[col_nombre]
            codigo = fila[col_codigo]
            generar_certificado(nombre, codigo, proyecto, espacio, fecha_texto)

print("\n🎉 Todos los certificados fueron generados correctamente.")
