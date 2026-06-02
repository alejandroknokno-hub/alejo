"""Consolidación de jornada: un único Excel (con varias HOJAS) acumulado.

El archivo `data/consolidacion_de_jornada.xlsx` es un documento único (no separado
por fecha) con varias hojas:

- "Consolidación de jornada": una fila por cada documento de control generado (acumula).
- "Registros": detalle de los registros documentales (foto actual).
- "Colaboradores": ficha de los colaboradores (foto actual).
- "Bitácora": eventos de la sesión de monitoreo (foto actual).

La hoja de consolidación se acumula; las hojas de detalle se actualizan con el
estado vigente cada vez que se guarda un documento.
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

from . import config

# Nombres de las hojas de detalle.
HOJA_REGISTROS = "Registros"
HOJA_COLABORADORES = "Colaboradores"
HOJA_BITACORA = "Bitácora"


def asegurar_archivo() -> None:
    """Crea el Excel consolidado (solo la hoja principal) si aún no existe."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not config.CONSOLIDADO_PATH.exists():
        vacio = pd.DataFrame(columns=config.CONSOLIDADO_ORDER)
        _escribir_libro(vacio, None, None, None)


def leer_consolidado() -> pd.DataFrame:
    """Devuelve la hoja de consolidación (con nombres internos de columna)."""
    asegurar_archivo()
    df = pd.read_excel(config.CONSOLIDADO_PATH, sheet_name=config.CONSOLIDADO_SHEET, engine="openpyxl")
    # El archivo guarda etiquetas legibles: las mapeamos de vuelta a nombres internos.
    inverso = {v: k for k, v in config.CONSOLIDADO_COLUMNS.items()}
    df = df.rename(columns=inverso)
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
    df_registros: pd.DataFrame | None = None,
    df_colaboradores: pd.DataFrame | None = None,
    df_bitacora: pd.DataFrame | None = None,
) -> dict:
    """Agrega (o actualiza) el documento y refresca las hojas de detalle.

    - `actualizar=False`: agrega una fila nueva a la hoja de consolidación.
    - `actualizar=True`: reemplaza la última fila de la misma sesión (regeneración).

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

    _escribir_libro(df, df_registros, df_colaboradores, df_bitacora)
    return {"ok": True, "ruta": str(config.CONSOLIDADO_PATH), "entradas": len(df)}


# --- Escritura del libro con varias hojas ----------------------------------
def _hoja(writer, nombre: str, df: pd.DataFrame | None, etiquetas: dict, orden: list | None) -> None:
    """Escribe una hoja de detalle con columnas ordenadas y etiquetadas."""
    salida = (df.copy() if df is not None else pd.DataFrame(columns=orden or []))
    if orden:
        for col in orden:
            if col not in salida.columns:
                salida[col] = pd.NA
        salida = salida[orden]
    salida = salida.rename(columns=etiquetas)
    salida.to_excel(writer, sheet_name=nombre, index=False)


def _escribir_libro(
    df_consolidado: pd.DataFrame,
    df_registros: pd.DataFrame | None,
    df_colaboradores: pd.DataFrame | None,
    df_bitacora: pd.DataFrame | None,
) -> None:
    """Escribe el libro completo (hoja de consolidación + hojas de detalle)."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    consolidado = df_consolidado[config.CONSOLIDADO_ORDER].rename(columns=config.CONSOLIDADO_COLUMNS)

    with pd.ExcelWriter(config.CONSOLIDADO_PATH, engine="openpyxl") as writer:
        consolidado.to_excel(writer, sheet_name=config.CONSOLIDADO_SHEET, index=False)
        _hoja(writer, HOJA_REGISTROS, df_registros, config.COLUMNS, config.COLUMN_ORDER)
        _hoja(writer, HOJA_COLABORADORES, df_colaboradores, config.COLAB_COLUMNS, config.COLAB_ORDER)
        _hoja(writer, HOJA_BITACORA, df_bitacora, {}, None)
