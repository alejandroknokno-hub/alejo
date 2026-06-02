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
UMBRAL_NOVEDAD = 4   # EFECTIVO si días de creación de la novedad <= 4

_VACIOS = {"", "n/a", "na", "nan", "none", "#n/d"}

# --- Columnas de la hoja SIROS NOVEDADES -----------------------------------
HOJA_NOV = "SIROS NOVEDADES"
N_ID = "ID"
N_FECHA_CREACION = "FECHA CREACION"
N_FECHA_REVISION = "FECHA DE REVISION"
N_FECHA_RESP_TI = "FECHA DE RESPUESTA TI"
N_DIAS_CIERRE = "DIAS DE CIERRE"
N_RANGO_TI = "RANGO DE DIAS TI"
N_DIAS_CREACION = "DIAS DE CREACION"
N_RANGO = "RANGO DE DIAS"
N_INDICE = "INDICE DE EFECTIVIDAD RESPUESTA"
N_RAZON = "RAZON DEL RETRASO"
N_NOVEDAD = "NOVEDAD"
N_ESTADO = "ESTADO"
N_CODIGO = "CODIGO DE VENDEDOR NUEVO"
N_CREADOR = "CREADOR"
N_NOMBRE = "NOMBRE Y APELLIDOS"
N_CEDULA = "CEDULA"
N_AREA = "AREA"
N_OBS = "OBSERVACION"
N_MES = "MES"

# Subtipos de novedad (hojas pequeñas): mismas fechas y regla (<= 4).
HOJA_ACJEFE = "SIROS NOVEDADES AC.JEFE"
HOJA_PROM = "SIROS NOVEDADES PROM.INTERNAS"

# Orden de categorías para los resúmenes/gráficos.
ORDEN_RANGO_CREACION = ["Entre 0 a 5 Dias", "Entre 6 a 10 Dias", "Entre 11 a 15 Dias", "Mayor a 15 Dias"]
ORDEN_RANGO_TI_NOV = ["Entre 0 a 2 Dias", "Entre 3 a 5 Dias", "Mayor a 6 Dias"]


def _val(v):
    if v is None or (np.isscalar(v) and pd.isna(v)):
        return None
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return float(v)
    return v


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


def rango_cierre_novedad(dias) -> str | None:
    """Rango TI de novedades (bloque 3 de MEDIDAS): 0-2 / 3-5 / >6."""
    if dias is None or pd.isna(dias):
        return None
    d = int(dias)
    if d <= 2:
        return "Entre 0 a 2 Dias"
    if d <= 5:
        return "Entre 3 a 5 Dias"
    return "Mayor a 6 Dias"


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


# ===========================================================================
# NOVEDADES (traslados, cambio de jefe, promociones internas)
# ===========================================================================
def enriquecer_novedades(df: pd.DataFrame, df_colaboradores: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Rellena código de vendedor nuevo, área y nombre por cédula desde Global."""
    df = df.copy()
    reporte = {"codigo": 0, "area": 0, "nombre": 0}
    if df_colaboradores is None or df_colaboradores.empty or N_CEDULA not in df.columns:
        return df, reporte

    colab = df_colaboradores.copy()
    colab["cedula"] = colab["cedula"].astype("string").str.strip()
    indexado = colab.set_index("cedula")

    mapeo = {N_CODIGO: ("codigo_vendedor", "codigo"), N_AREA: ("area", "area"), N_NOMBRE: ("nombre", "nombre")}
    for col_destino, (col_global, clave) in mapeo.items():
        if col_destino not in df.columns or col_global not in indexado.columns:
            continue
        for i in df.index:
            if not _es_vacio(df.at[i, col_destino]):
                continue
            ced = str(df.at[i, N_CEDULA]).strip().split(".")[0]
            if ced in indexado.index:
                valor = indexado.at[ced, col_global]
                if isinstance(valor, pd.Series):
                    valor = valor.iloc[0]
                if not _es_vacio(valor):
                    df.at[i, col_destino] = valor
                    reporte[clave] += 1
    return df, reporte


def calcular_derivados_novedades(df: pd.DataFrame) -> pd.DataFrame:
    """Recalcula días, rangos, índice de efectividad respuesta y mes en NOVEDADES."""
    df = df.copy()
    for c in (N_FECHA_CREACION, N_FECHA_REVISION, N_FECHA_RESP_TI):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    for i in df.index:
        creacion = df.at[i, N_FECHA_CREACION] if N_FECHA_CREACION in df.columns else None
        rev = df.at[i, N_FECHA_REVISION] if N_FECHA_REVISION in df.columns else None
        resp = df.at[i, N_FECHA_RESP_TI] if N_FECHA_RESP_TI in df.columns else None

        d_creacion = networkdays(creacion, rev)
        d_cierre = networkdays(rev, resp)

        _set(df, i, N_DIAS_CREACION, d_creacion)
        _set(df, i, N_DIAS_CIERRE, d_cierre)
        _set(df, i, N_RANGO, rango_creacion(d_creacion))
        _set(df, i, N_RANGO_TI, rango_cierre_novedad(d_cierre) if d_cierre is not None else "#N/D")
        _set(df, i, N_INDICE, indice(d_creacion, UMBRAL_NOVEDAD))
        if creacion is not None and not pd.isna(creacion):
            _set(df, i, N_MES, config.MESES_ES[pd.Timestamp(creacion).month - 1])
    return df


def calcular_derivados_subtipo(df: pd.DataFrame) -> pd.DataFrame:
    """Para AC.JEFE / PROM.INTERNAS: días de creación e índice (regla <= 4)."""
    df = df.copy()
    for c in (N_FECHA_CREACION, N_FECHA_REVISION):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    if N_FECHA_CREACION not in df.columns:
        return df
    df = df[df[N_FECHA_CREACION].notna()]
    for i in df.index:
        d = networkdays(df.at[i, N_FECHA_CREACION], df.at[i, N_FECHA_REVISION])
        _set(df, i, N_DIAS_CREACION, d)
        _set(df, i, "INDICE DE EFECTIVIDAD", indice(d, UMBRAL_NOVEDAD))
    return df


# ===========================================================================
# Resúmenes para los GRÁFICOS
# ===========================================================================
def _conteo(serie: pd.Series, orden: list | None = None) -> pd.DataFrame:
    s = serie.dropna().astype(str).str.strip()
    s = s[~s.str.lower().isin(_VACIOS)]
    vc = s.value_counts()
    if orden:
        vc = vc.reindex(orden, fill_value=0)
    return vc.rename_axis("Categoría").reset_index(name="Cantidad")


def resumenes_informatica(df: pd.DataFrame) -> list[tuple]:
    return [
        ("SIROS por rango de creación", _conteo(df.get(C_RANGO_CREACION, pd.Series(dtype=str)), ORDEN_RANGO_CREACION), "bar"),
        ("SIROS por estado", _conteo(df.get(C_ESTADO, pd.Series(dtype=str))), "pie"),
        ("SIROS por área", _conteo(df.get(C_AREA, pd.Series(dtype=str))), "bar"),
    ]


def resumenes_novedades(df: pd.DataFrame) -> list[tuple]:
    return [
        ("Novedades por estado", _conteo(df.get(N_ESTADO, pd.Series(dtype=str))), "pie"),
        ("Índice de efectividad respuesta", _conteo(df.get(N_INDICE, pd.Series(dtype=str))), "pie"),
        ("Novedades por rango TI", _conteo(df.get(N_RANGO_TI, pd.Series(dtype=str)), ORDEN_RANGO_TI_NOV), "bar"),
    ]


# ===========================================================================
# Orquestación del libro completo + exportación con gráficos
# ===========================================================================
def completar_libro(ruta_o_buffer, df_colaboradores: pd.DataFrame) -> tuple[dict, dict]:
    """Procesa todas las hojas del indicador (informática + novedades + subtipos)."""
    xls = pd.ExcelFile(ruta_o_buffer, engine="openpyxl")
    nombres = {h.strip().upper(): h for h in xls.sheet_names}
    dfs: dict = {}
    reporte: dict = {}

    if HOJA in nombres:
        df = xls.parse(nombres[HOJA])
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        if C_ID in df.columns:
            df = df[df[C_ID].notna()]
        df, rg = enriquecer_con_global(df, df_colaboradores)
        df = calcular_derivados(df)
        dfs["informatica"] = df
        reporte["informatica_filas"] = len(df)
        reporte["informatica_rellenados"] = sum(rg.values())

    if HOJA_NOV in nombres:
        df = xls.parse(nombres[HOJA_NOV])
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        if N_ID in df.columns:
            df = df[df[N_ID].notna()]
        df, rgn = enriquecer_novedades(df, df_colaboradores)
        df = calcular_derivados_novedades(df)
        dfs["novedades"] = df
        reporte["novedades_filas"] = len(df)
        reporte["novedades_rellenados"] = sum(rgn.values())

    for clave, nombre_norm in (("ac_jefe", HOJA_ACJEFE), ("prom", HOJA_PROM)):
        if nombre_norm in nombres:
            df = xls.parse(nombres[nombre_norm])
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all")
            dfs[clave] = calcular_derivados_subtipo(df)

    return dfs, reporte


def exportar_libro(dfs: dict) -> bytes:
    """Escribe el indicador completado (todas las hojas) + GRÁFICOS con gráficos."""
    import io

    buf = io.BytesIO()
    hojas_datos = [
        ("informatica", HOJA), ("novedades", HOJA_NOV),
        ("ac_jefe", HOJA_ACJEFE), ("prom", HOJA_PROM),
    ]
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for clave, nombre in hojas_datos:
            if clave in dfs:
                dfs[clave].to_excel(writer, sheet_name=nombre[:31], index=False)
        if "informatica" in dfs:
            _hoja_graficos(writer.book, "GRAFICOS INFORMATICA", resumenes_informatica(dfs["informatica"]))
        if "novedades" in dfs:
            _hoja_graficos(writer.book, "GRAFICOS NOVEDADES", resumenes_novedades(dfs["novedades"]))
    return buf.getvalue()


def _hoja_graficos(wb, nombre: str, items: list[tuple]) -> None:
    """Crea una hoja con tablas de resumen y gráficos nativos (torta/barras)."""
    from openpyxl.chart import BarChart, PieChart, Reference

    ws = wb.create_sheet(nombre[:31])
    fila = 1
    ancla = 1
    for titulo, df2, tipo in items:
        ws.cell(row=fila, column=1, value=titulo)
        encabezado = fila + 1
        ws.cell(row=encabezado, column=1, value="Categoría")
        ws.cell(row=encabezado, column=2, value="Cantidad")
        r = encabezado + 1
        for _, registro in df2.iterrows():
            ws.cell(row=r, column=1, value=_val(registro["Categoría"]))
            ws.cell(row=r, column=2, value=_val(registro["Cantidad"]))
            r += 1
        ultima = r - 1
        if ultima >= encabezado + 1:
            grafico = PieChart() if tipo == "pie" else BarChart()
            grafico.title = titulo
            grafico.add_data(Reference(ws, min_col=2, min_row=encabezado, max_row=ultima), titles_from_data=True)
            grafico.set_categories(Reference(ws, min_col=1, min_row=encabezado + 1, max_row=ultima))
            ws.add_chart(grafico, f"E{ancla}")
            ancla += 16
        fila = ultima + 3
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 12
