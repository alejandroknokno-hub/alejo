"""Indicador SIRO: completa/enriquece el indicador y recalcula lo derivado.

Trabaja sobre la hoja "SIROS INFORMATICA" (creación en Oracle). Hace dos cosas:

1. **Enriquecer por cédula desde Global** (colaboradores): rellena código de
   vendedor, área, fecha de ingreso y nombre donde falten ('', NaN o 'N/A').

2. **Recalcular los campos derivados** (los que hoy se llenan a mano):
   - Días de creación / efectivos al ingreso / asignación de correo / cierre,
     como NETWORKDAYS (días hábiles, lunes a viernes, inclusivo).
   - Índice de efectividad (EFECTIVO/RETRASO) e índice TI según umbrales.
   - Rango de días según la hoja MEDIDAS.
   - Mes y año a partir de la fecha de creación del SIRO.

Las fórmulas se dedujeron del propio indicador y se validan contra él.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from . import config

# --- Nombres de columnas de la hoja SIROS INFORMATICA ----------------------
HOJA = "SIROS INFORMATICA"
C_ID = "ID"
C_FECHA_CREACION = "FECHA CREACION SIRO"
C_NOVEDAD = "NOVEDAD"
C_ESTADO = "ESTADO"
C_OBS_ESTADO = "OBSERVACION ESTADO"
C_FECHA_REV_RH = "FECHA DE REVISION RH"
C_FECHA_ASIG = "FECHA DE ASIGNACION CORREOS"
C_FECHA_CIERRE = "FECHA DE CIERRE TI"
C_CREADOR = "CREADOR DEL SIRO"
C_NOMBRE = "NOMBRE Y APELLIDOS DEL COLABORADOR"
C_CEDULA = "CEDULA"
C_FECHA_INGRESO = "FECHA INGRESO"
C_CODIGO = "CODIGO DEL VENDEDOR"
C_AREA = "AREA"
C_DIAS_EFECTIVOS = "DIAS EFECTIVOS AL INGRESO"
C_INDICE = "INDICE DE EFECTIVIDAD"
C_DIAS_CREACION = "DIAS DE CREACION"
C_RANGO_CREACION = "RANGO DE DIAS POR CREACION"
C_DIAS_ASIG = "DIAS DE ASIGNACION CORREO"
C_INDICE_TI = "INDICE DE EFECTIVIDAD TI"
C_RANGO_ASIG = "RANGO DE DIAS POR AS. CORREO"
C_DIAS_CIERRE = "DIAS DE CIERRE SIRO POR RESPUESTA"
C_MES = "MES SIRO"
C_ANIO = "AÑO"

# Umbrales de efectividad (deducidos del indicador).
UMBRAL_INGRESO = 5   # EFECTIVO si días efectivos al ingreso <= 5
UMBRAL_TI = 3        # EFECTIVO si días de asignación de correo <= 3

_VACIOS = {"", "n/a", "na", "nan", "none", "#n/d"}


# --- Cálculos base ---------------------------------------------------------
def networkdays(inicio, fin) -> int | None:
    """Días hábiles (lun-vie) entre dos fechas, inclusivo (estilo NETWORKDAYS)."""
    if inicio is None or fin is None or pd.isna(inicio) or pd.isna(fin):
        return None
    try:
        a = np.datetime64(pd.Timestamp(inicio).date())
        b = np.datetime64(pd.Timestamp(fin).date())
    except (ValueError, TypeError):
        return None
    if b < a:
        return None
    return int(np.busday_count(a, b + np.timedelta64(1, "D")))


def rango_creacion(dias) -> str | None:
    if dias is None or pd.isna(dias):
        return None
    d = int(dias)
    if d <= 5:
        return "Entre 0 a 5 Dias"
    if d <= 10:
        return "Entre 6 a 10 Dias"
    if d <= 15:
        return "Entre 11 a 15 Dias"
    return "Mayor a 15 Dias"


def rango_asignacion(dias) -> str | None:
    if dias is None or pd.isna(dias):
        return None
    d = int(dias)
    if d <= 3:
        return "Entre 1 a 3 Dias"
    if d <= 6:
        return "Entre 4 a 6 Dias"
    return "Mayor a 7 Dias"


def indice(dias, umbral: int) -> str | None:
    if dias is None or pd.isna(dias):
        return None
    return "EFECTIVO" if int(dias) <= umbral else "RETRASO"


def _es_vacio(v) -> bool:
    if v is None or (np.isscalar(v) and pd.isna(v)):
        return True
    return str(v).strip().lower() in _VACIOS


# --- Carga -----------------------------------------------------------------
def cargar_indicador(ruta_o_buffer) -> pd.DataFrame:
    """Lee la hoja SIROS INFORMATICA (tolera espacios en el nombre/encabezados)."""
    xls = pd.ExcelFile(ruta_o_buffer, engine="openpyxl")
    objetivo = next((h for h in xls.sheet_names if h.strip().upper() == HOJA), None)
    if objetivo is None:
        raise ValueError(f"No se encontró la hoja '{HOJA}' en el archivo.")
    df = xls.parse(objetivo)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# --- Enriquecer desde Global (por cédula) ----------------------------------
def enriquecer_con_global(df: pd.DataFrame, df_colaboradores: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Rellena código, área, fecha de ingreso y nombre por cédula desde Global."""
    df = df.copy()
    reporte = {"codigo": 0, "area": 0, "fecha_ingreso": 0, "nombre": 0}
    if df_colaboradores is None or df_colaboradores.empty or C_CEDULA not in df.columns:
        return df, reporte

    colab = df_colaboradores.copy()
    colab["cedula"] = colab["cedula"].astype("string").str.strip()
    indexado = colab.set_index("cedula")

    def _ced(v):
        s = str(v).strip()
        return s.split(".")[0] if s.endswith(".0") else s

    mapeo = {
        C_CODIGO: ("codigo_vendedor", "codigo"),
        C_AREA: ("area", "area"),
        C_FECHA_INGRESO: ("fecha_ingreso", "fecha_ingreso"),
        C_NOMBRE: ("nombre", "nombre"),
    }
    for col_destino, (col_global, clave) in mapeo.items():
        if col_destino not in df.columns or col_global not in indexado.columns:
            continue
        for i in df.index:
            if not _es_vacio(df.at[i, col_destino]):
                continue
            ced = _ced(df.at[i, C_CEDULA])
            if ced in indexado.index:
                valor = indexado.at[ced, col_global]
                if isinstance(valor, pd.Series):
                    valor = valor.iloc[0]
                if not _es_vacio(valor):
                    df.at[i, col_destino] = valor
                    reporte[clave] += 1
    return df, reporte


# --- Recalcular campos derivados -------------------------------------------
def calcular_derivados(df: pd.DataFrame) -> pd.DataFrame:
    """Recalcula días, índices, rangos, mes y año a partir de las fechas."""
    df = df.copy()
    for c in (C_FECHA_CREACION, C_FECHA_REV_RH, C_FECHA_ASIG, C_FECHA_CIERRE, C_FECHA_INGRESO):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for i in df.index:
        creacion = df.at[i, C_FECHA_CREACION] if C_FECHA_CREACION in df.columns else None
        rev = df.at[i, C_FECHA_REV_RH] if C_FECHA_REV_RH in df.columns else None
        asig = df.at[i, C_FECHA_ASIG] if C_FECHA_ASIG in df.columns else None
        cierre = df.at[i, C_FECHA_CIERRE] if C_FECHA_CIERRE in df.columns else None
        ingreso = df.at[i, C_FECHA_INGRESO] if C_FECHA_INGRESO in df.columns else None

        d_creacion = networkdays(creacion, rev)
        d_efectivos = networkdays(ingreso, rev)
        d_asig = networkdays(creacion, asig)
        d_cierre = networkdays(rev, cierre)

        _set(df, i, C_DIAS_CREACION, d_creacion)
        _set(df, i, C_DIAS_EFECTIVOS, d_efectivos)
        _set(df, i, C_DIAS_ASIG, d_asig)
        _set(df, i, C_DIAS_CIERRE, d_cierre)
        _set(df, i, C_INDICE, indice(d_efectivos, UMBRAL_INGRESO))
        _set(df, i, C_INDICE_TI, indice(d_asig, UMBRAL_TI))
        _set(df, i, C_RANGO_CREACION, rango_creacion(d_creacion))
        _set(df, i, C_RANGO_ASIG, rango_asignacion(d_asig))
        # Mes y año se derivan de la FECHA INGRESO (validado contra el indicador).
        if ingreso is not None and not pd.isna(ingreso):
            ts = pd.Timestamp(ingreso)
            _set(df, i, C_MES, config.MESES_ES[ts.month - 1])
            _set(df, i, C_ANIO, ts.year)
    return df


def _set(df, i, col, valor) -> None:
    if col in df.columns:
        df.at[i, col] = valor


# --- Orquestación ----------------------------------------------------------
def completar(ruta_o_buffer, df_colaboradores: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Carga el indicador, lo enriquece desde Global y recalcula lo derivado."""
    df = cargar_indicador(ruta_o_buffer)
    df = df.dropna(how="all")
    if C_ID in df.columns:
        df = df[df[C_ID].notna()]

    df, rep_global = enriquecer_con_global(df, df_colaboradores)
    df = calcular_derivados(df)

    reporte = {
        "filas": len(df),
        "rellenado_codigo": rep_global["codigo"],
        "rellenado_area": rep_global["area"],
        "rellenado_fecha_ingreso": rep_global["fecha_ingreso"],
        "rellenado_nombre": rep_global["nombre"],
    }
    return df, reporte
