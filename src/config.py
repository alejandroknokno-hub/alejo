"""Configuración central de la aplicación de Control Documental.

Aquí se definen las rutas de los archivos, el nombre de la hoja de Excel y el
esquema (columnas) de los registros. Centralizar esto facilita el mantenimiento:
si necesitas agregar un campo nuevo, lo haces en un solo lugar.
"""
from __future__ import annotations

from pathlib import Path

# --- Rutas -----------------------------------------------------------------
# Carpeta raíz del proyecto (un nivel por encima de /src)
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"

# Archivo de Excel donde se guardan TODOS los registros de la jornada.
# Esta es la "conexión directa con Excel": leemos y escribimos sobre este .xlsx.
EXCEL_PATH = DATA_DIR / "control_documental.xlsx"
SHEET_NAME = "Registros"

# --- Esquema de datos (control documental) ---------------------------------
# Cada clave es el nombre interno de la columna; el valor es la etiqueta visible.
COLUMNS: dict[str, str] = {
    "fecha_hora": "Fecha y hora",
    "dia": "Día",
    "mes": "Mes",
    "anio": "Año",
    "codigo_vendedor": "Código de vendedor",
    "nombre": "Nombre",
    "cedula": "Cédula",
    "combinacion_contable": "Combinación contable",
    "valor": "Valor",
    "estado": "Estado",
    "informacion_relevante": "Información relevante",
    "registrado_por": "Registrado por",
}

# Orden de las columnas tal como aparecerán en el Excel y las tablas.
COLUMN_ORDER: list[str] = list(COLUMNS.keys())

# Estados posibles de un documento (útil para los indicadores de control).
ESTADOS = ["Pendiente", "En proceso", "Aprobado", "Rechazado"]

# Meses en español para mostrar de forma legible.
MESES_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

# --- Versión Oracle: monitoreo por captura de pantalla ---------------------
# Carpeta donde se guardan las capturas de pantalla del monitoreo.
CAPTURAS_DIR = DATA_DIR / "capturas"
# Bitácora de eventos (entradas, modificaciones y capturas) de cada sesión.
MONITOR_LOG = DATA_DIR / "monitor_log.csv"
MONITOR_COLUMNS = ["fecha_hora", "sesion", "tipo", "descripcion", "archivo"]
# Tipos de evento que se registran en la bitácora.
TIPOS_EVENTO = ["Captura", "Entrada", "Modificación", "Consulta", "Nota"]

# --- Versión Global: información del colaborador ----------------------------
GLOBAL_PATH = DATA_DIR / "global_colaboradores.xlsx"
GLOBAL_SHEET = "Colaboradores"

COLAB_COLUMNS: dict[str, str] = {
    "codigo_vendedor": "Código de vendedor",
    "cedula": "Cédula",
    "nombre": "Nombre",
    "cargo": "Cargo",
    "area": "Área",
    "correo": "Correo",
    "telefono": "Teléfono",
    "fecha_ingreso": "Fecha de ingreso",
    "estado": "Estado",
}
COLAB_ORDER: list[str] = list(COLAB_COLUMNS.keys())
ESTADOS_COLAB = ["Activo", "Inactivo", "Vacaciones", "Retirado"]
