"""Consolidación de jornada: un único Excel acumulado con los documentos generados.

Todos los documentos de control que produce la IA se guardan en UN SOLO archivo
(`data/consolidacion_de_jornada.xlsx`), agregando una fila por cada documento.
No se separa por fecha: es un documento único que se va acumulando.
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

from . import config


def asegurar_archivo() -> None:
    """Crea el Excel consolidado con encabezados si aún no existe."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not config.CONSOLIDADO_PATH.exists():
        df = pd.DataFrame(columns=config.CONSOLIDADO_ORDER)
        df.to_excel(config.CONSOLIDADO_PATH, sheet_name=config.CONSOLIDADO_SHEET, index=False)


def leer_consolidado() -> pd.DataFrame:
    """Devuelve todas las entradas acumuladas en el consolidado."""
    asegurar_archivo()
    df = pd.read_excel(config.CONSOLIDADO_PATH, sheet_name=config.CONSOLIDADO_SHEET, engine="openpyxl")
    for col in config.CONSOLIDADO_ORDER:
        if col not in df.columns:
            df[col] = pd.NA
    return df[config.CONSOLIDADO_ORDER]


def guardar_documento(
    doc: dict,
    sesion: str = "",
    usuario: str = "",
    kpis: dict | None = None,
    actualizar: bool = False,
) -> dict:
    """Agrega (o actualiza) el documento de control en el Excel consolidado.

    - `actualizar=False`: agrega una fila nueva (acumula).
    - `actualizar=True`: reemplaza la última fila de la misma sesión (útil cuando
      el usuario regenera el documento tras responder aclaraciones).

    Devuelve {"ok": bool, "ruta": str, "entradas": int}.
    """
    df = leer_consolidado()
    ahora = datetime.now()
    kpis = kpis or {}

    fila = {
        "fecha_hora": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "sesion": sesion,
        "usuario": usuario,
        "registros_total": int(kpis.get("total_registros", 0)),
        "valor_total": float(kpis.get("valor_total", 0.0)),
        "titulo": doc.get("titulo", ""),
        "resumen": doc.get("resumen", ""),
        "preguntas_pendientes": len(doc.get("preguntas", [])),
        "documento_markdown": doc.get("documento_markdown", ""),
    }

    if actualizar and not df.empty and (df["sesion"] == sesion).any():
        ultimo = df[df["sesion"] == sesion].index[-1]
        df = df.drop(index=ultimo)

    df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
    df = df[config.CONSOLIDADO_ORDER]
    df.to_excel(config.CONSOLIDADO_PATH, sheet_name=config.CONSOLIDADO_SHEET, index=False, engine="openpyxl")

    return {"ok": True, "ruta": str(config.CONSOLIDADO_PATH), "entradas": len(df)}
