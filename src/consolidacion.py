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

import numpy as np
import pandas as pd

from . import config, indicador_siro, indicadores, pivote_excel

# Nombres de las hojas de detalle.
HOJA_REGISTROS = "Registros"
HOJA_COLABORADORES = "Colaboradores"
HOJA_BITACORA = "Bitácora"
# Hojas analíticas.
HOJA_KPIS = "KPIs"
HOJA_PIVOTES = "Tablas dinámicas"
HOJA_DASHBOARD = "Dashboard"
HOJA_PIVOTE_INT = "Pivote interactivo"
# Hojas del indicador SIRO dentro de la consolidación.
HOJA_SIRO_GRAFICOS = "SIRO Graficos"
SIRO_HOJAS = {
    "informatica": "SIRO Informatica",
    "novedades": "SIRO Novedades",
    "ac_jefe": "SIRO Cambio de jefe",
    "prom": "SIRO Prom internas",
}


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
        # Hojas analíticas: KPIs, tablas dinámicas y dashboard con gráficos.
        _hojas_analiticas(writer.book, df_registros)
        # Hojas del indicador SIRO (si se guardó previamente) + sus gráficos.
        _hojas_siro(writer, _leer_cache_siro())
        # Hoja vacía donde se inyectará la tabla dinámica nativa (más abajo).
        writer.book.create_sheet(HOJA_PIVOTE_INT)

    # Inyecta la PivotTable interactiva (recalculada por Excel al abrir).
    # Si algo falla, el resto del libro queda intacto y válido.
    pivote_excel.inyectar_pivote(
        config.CONSOLIDADO_PATH, df_registros,
        hoja_datos=HOJA_REGISTROS, hoja_pivote=HOJA_PIVOTE_INT,
    )


# --- Hojas analíticas (KPIs, tablas dinámicas y dashboard) ------------------
def _val(v):
    """Convierte valores de pandas/numpy a tipos nativos para openpyxl."""
    if v is None or (np.isscalar(v) and pd.isna(v)):
        return None
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return float(v)
    return v


def _volcar_tabla(ws, df: pd.DataFrame, etiquetas: dict, fila: int, titulo: str) -> dict:
    """Escribe una tabla (título + encabezados + filas) y devuelve su ubicación."""
    ws.cell(row=fila, column=1, value=titulo)
    encabezado = fila + 1
    columnas = list(df.columns)
    for j, col in enumerate(columnas):
        ws.cell(row=encabezado, column=1 + j, value=etiquetas.get(col, col))
    primera = encabezado + 1
    r = primera
    for _, registro in df.iterrows():
        for j, col in enumerate(columnas):
            ws.cell(row=r, column=1 + j, value=_val(registro[col]))
        r += 1
    ultima = r - 1
    return {"encabezado": encabezado, "primera": primera, "ultima": ultima,
            "ncols": len(columnas), "siguiente": r + 2}


def _hojas_analiticas(wb, df_registros: pd.DataFrame | None) -> None:
    """Crea las hojas KPIs, Tablas dinámicas y Dashboard (con gráficos)."""
    from openpyxl.chart import BarChart, PieChart, Reference

    df = df_registros if df_registros is not None else pd.DataFrame(columns=config.COLUMN_ORDER)

    # --- Hoja KPIs ---
    k = indicadores.kpis_generales(df)
    ws_kpi = wb.create_sheet(HOJA_KPIS)
    ws_kpi.cell(row=1, column=1, value="Indicadores de control")
    ws_kpi.cell(row=2, column=1, value="Indicador")
    ws_kpi.cell(row=2, column=2, value="Valor")
    filas_kpi = [
        ("Registros totales", k["total_registros"]),
        ("Registros hoy", k["registros_hoy"]),
        ("Vendedores activos", k["vendedores_activos"]),
        ("Tasa de aprobación (%)", round(k["tasa_aprobacion"], 1)),
        ("Pendientes", k["pendientes"]),
        ("Valor total", k["valor_total"]),
        ("Valor de hoy", k["valor_hoy"]),
    ]
    for i, (nombre, valor) in enumerate(filas_kpi, start=3):
        ws_kpi.cell(row=i, column=1, value=nombre)
        ws_kpi.cell(row=i, column=2, value=_val(valor))

    fila_alertas = 3 + len(filas_kpi) + 1
    ws_kpi.cell(row=fila_alertas, column=1, value="Alertas de control")
    for i, alerta in enumerate(indicadores.alertas_control(df), start=fila_alertas + 1):
        ws_kpi.cell(row=i, column=1, value=alerta["tipo"].upper())
        ws_kpi.cell(row=i, column=2, value=alerta["mensaje"])
    ws_kpi.column_dimensions["A"].width = 26
    ws_kpi.column_dimensions["B"].width = 60

    # --- Hoja Tablas dinámicas (resúmenes que alimentan el dashboard) ---
    est = indicadores.por_estado(df)
    vend = indicadores.por_vendedor(df).head(10)
    comb = indicadores.por_combinacion(df).head(10)

    ws_piv = wb.create_sheet(HOJA_PIVOTES)
    etiquetas_piv = {
        "estado": "Estado", "cantidad": "Cantidad",
        "codigo_vendedor": "Código vendedor", "nombre": "Nombre",
        "registros": "Registros", "valor_total": "Valor total",
        "combinacion_contable": "Combinación contable",
    }
    pos_est = _volcar_tabla(ws_piv, est, etiquetas_piv, 1, "Registros por estado")
    pos_vend = _volcar_tabla(ws_piv, vend, etiquetas_piv, pos_est["siguiente"], "Registros por vendedor")
    pos_comb = _volcar_tabla(ws_piv, comb, etiquetas_piv, pos_vend["siguiente"], "Por combinación contable")
    ws_piv.column_dimensions["A"].width = 22
    ws_piv.column_dimensions["B"].width = 18

    # --- Hoja Dashboard (gráficos nativos de Excel) ---
    ws_dash = wb.create_sheet(HOJA_DASHBOARD)
    ws_dash.cell(row=1, column=1, value="Dashboard de la jornada")

    def _categorias(pos, col):
        return Reference(ws_piv, min_col=col, min_row=pos["primera"], max_row=pos["ultima"])

    def _datos(pos, col):
        return Reference(ws_piv, min_col=col, min_row=pos["encabezado"], max_row=pos["ultima"])

    # Torta: registros por estado.
    if pos_est["ultima"] >= pos_est["primera"]:
        torta = PieChart()
        torta.title = "Registros por estado"
        torta.add_data(_datos(pos_est, 2), titles_from_data=True)
        torta.set_categories(_categorias(pos_est, 1))
        ws_dash.add_chart(torta, "A3")

    # Barras: registros por vendedor.
    if pos_vend["ultima"] >= pos_vend["primera"]:
        barras_v = BarChart()
        barras_v.title = "Top vendedores (registros)"
        barras_v.add_data(_datos(pos_vend, 3), titles_from_data=True)
        barras_v.set_categories(_categorias(pos_vend, 1))
        barras_v.y_axis.title = "Registros"
        ws_dash.add_chart(barras_v, "J3")

    # Barras: por combinación contable.
    if pos_comb["ultima"] >= pos_comb["primera"]:
        barras_c = BarChart()
        barras_c.title = "Por combinación contable"
        barras_c.add_data(_datos(pos_comb, 2), titles_from_data=True)
        barras_c.set_categories(_categorias(pos_comb, 1))
        barras_c.y_axis.title = "Registros"
        ws_dash.add_chart(barras_c, "A20")


# --- Indicador SIRO dentro de la consolidación -----------------------------
def _guardar_cache_siro(dfs: dict) -> None:
    """Persiste las hojas del indicador SIRO completado en una caché."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(config.SIRO_CACHE_PATH, engine="openpyxl") as writer:
        for clave in SIRO_HOJAS:
            if clave in dfs and dfs[clave] is not None and not dfs[clave].empty:
                dfs[clave].to_excel(writer, sheet_name=clave, index=False)


def _leer_cache_siro() -> dict:
    """Lee la caché del indicador SIRO. Devuelve {} si no existe."""
    if not config.SIRO_CACHE_PATH.exists():
        return {}
    try:
        xls = pd.ExcelFile(config.SIRO_CACHE_PATH, engine="openpyxl")
        return {h: xls.parse(h) for h in xls.sheet_names}
    except Exception:
        return {}


def leer_siro() -> dict:
    """API pública: devuelve las hojas del indicador SIRO guardado (o {})."""
    return _leer_cache_siro()


def _hojas_siro(writer, dfs_siro: dict) -> None:
    """Escribe las hojas de datos del indicador SIRO + una hoja de gráficos SIRO."""
    if not dfs_siro:
        return
    for clave, nombre in SIRO_HOJAS.items():
        df = dfs_siro.get(clave)
        if df is not None and not df.empty:
            df.to_excel(writer, sheet_name=nombre[:31], index=False)

    items = []
    if "informatica" in dfs_siro:
        items += indicador_siro.resumenes_informatica(dfs_siro["informatica"])
    if "novedades" in dfs_siro:
        items += indicador_siro.resumenes_novedades(dfs_siro["novedades"])
    if items:
        indicador_siro._hoja_graficos(writer.book, HOJA_SIRO_GRAFICOS, items)


def guardar_siro_en_consolidado(
    dfs: dict,
    df_registros: pd.DataFrame | None = None,
    df_colaboradores: pd.DataFrame | None = None,
    df_bitacora: pd.DataFrame | None = None,
) -> dict:
    """Guarda el indicador SIRO completado dentro de la Consolidación de jornada.

    Persiste los datos SIRO y reescribe el libro consolidado (que ahora incluye
    las hojas SIRO y sus gráficos), conservando los documentos ya acumulados.
    """
    _guardar_cache_siro(dfs)
    consolidado = leer_consolidado()
    # _escribir_libro incluye las hojas SIRO (desde la caché) e inyecta la PivotTable.
    _escribir_libro(consolidado, df_registros, df_colaboradores, df_bitacora)
    hojas = [SIRO_HOJAS[k] for k in SIRO_HOJAS if k in dfs and not dfs[k].empty]
    return {"ok": True, "hojas": hojas, "graficos": HOJA_SIRO_GRAFICOS}
