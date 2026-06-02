"""Capa de acceso a datos: conexión directa con Excel (.xlsx).

Responsabilidades:
- Crear el archivo de Excel si no existe.
- Leer todos los registros como un DataFrame de pandas.
- Validar y agregar un nuevo registro capturado durante la jornada.
- Exportar/guardar de vuelta en Excel.

Se usa openpyxl como motor para leer/escribir .xlsx sin necesidad de tener
Excel instalado, por lo que funciona en cualquier sistema operativo.
"""
from __future__ import annotations

import re
from datetime import datetime

import pandas as pd

from . import config


# --- Inicialización --------------------------------------------------------
def asegurar_archivo() -> None:
    """Crea la carpeta /data y un Excel vacío con encabezados si no existen."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not config.EXCEL_PATH.exists():
        df_vacio = pd.DataFrame(columns=config.COLUMN_ORDER)
        df_vacio.to_excel(config.EXCEL_PATH, sheet_name=config.SHEET_NAME, index=False)


# --- Lectura ---------------------------------------------------------------
def leer_registros() -> pd.DataFrame:
    """Devuelve todos los registros del Excel como DataFrame.

    Garantiza que existan todas las columnas del esquema y que los tipos de
    fecha/numéricos sean correctos para la analítica.
    """
    asegurar_archivo()
    df = pd.read_excel(config.EXCEL_PATH, sheet_name=config.SHEET_NAME, engine="openpyxl")

    # Asegurar todas las columnas del esquema (por si el Excel es antiguo).
    for col in config.COLUMN_ORDER:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[config.COLUMN_ORDER]

    # Tipado para que los indicadores funcionen bien.
    if not df.empty:
        df["fecha_hora"] = pd.to_datetime(df["fecha_hora"], errors="coerce")
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
        for c in ("dia", "anio"):
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("Int64")
    return df


# --- Validación ------------------------------------------------------------
def validar_registro(registro: dict) -> list[str]:
    """Valida los campos obligatorios. Devuelve una lista de errores (vacía si OK)."""
    errores: list[str] = []

    if not str(registro.get("codigo_vendedor", "")).strip():
        errores.append("El código de vendedor es obligatorio.")

    if not str(registro.get("nombre", "")).strip():
        errores.append("El nombre es obligatorio.")

    cedula = str(registro.get("cedula", "")).strip()
    if not cedula:
        errores.append("La cédula es obligatoria.")
    elif not re.fullmatch(r"\d{5,15}", cedula):
        errores.append("La cédula debe contener solo números (entre 5 y 15 dígitos).")

    if not str(registro.get("combinacion_contable", "")).strip():
        errores.append("La combinación contable es obligatoria.")

    estado = registro.get("estado", "")
    if estado not in config.ESTADOS:
        errores.append(f"El estado debe ser uno de: {', '.join(config.ESTADOS)}.")

    return errores


# --- Escritura -------------------------------------------------------------
def agregar_registro(registro: dict, registrado_por: str = "") -> dict:
    """Agrega un registro nuevo al Excel.

    Completa automáticamente la fecha/hora y los campos día, mes y año a partir
    del momento de la captura (durante la jornada laboral).

    Devuelve {"ok": bool, "errores": list[str], "registro": dict}.
    """
    errores = validar_registro(registro)
    if errores:
        return {"ok": False, "errores": errores, "registro": registro}

    ahora = datetime.now()
    completo = {
        "fecha_hora": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "dia": ahora.day,
        "mes": config.MESES_ES[ahora.month - 1],
        "anio": ahora.year,
        "codigo_vendedor": str(registro.get("codigo_vendedor", "")).strip(),
        "nombre": str(registro.get("nombre", "")).strip(),
        "cedula": str(registro.get("cedula", "")).strip(),
        "combinacion_contable": str(registro.get("combinacion_contable", "")).strip(),
        "valor": float(registro.get("valor", 0) or 0),
        "estado": registro.get("estado", config.ESTADOS[0]),
        "informacion_relevante": str(registro.get("informacion_relevante", "")).strip(),
        "registrado_por": registrado_por.strip(),
    }

    df = leer_registros()
    df = pd.concat([df, pd.DataFrame([completo])], ignore_index=True)
    guardar_registros(df)
    return {"ok": True, "errores": [], "registro": completo}


def guardar_registros(df: pd.DataFrame) -> None:
    """Escribe el DataFrame completo de vuelta en el Excel (sobrescribe la hoja)."""
    asegurar_archivo()
    df = df[config.COLUMN_ORDER]
    df.to_excel(config.EXCEL_PATH, sheet_name=config.SHEET_NAME, index=False, engine="openpyxl")
