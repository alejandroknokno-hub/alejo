"""Cálculo de indicadores de control y analítica sobre los registros.

Todas las funciones reciben un DataFrame (el de data_manager.leer_registros)
y devuelven valores o DataFrames listos para mostrar en el dashboard.
"""
from __future__ import annotations

from datetime import date

import pandas as pd

from . import config


def _solo_hoy(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra los registros cuya fecha sea la de hoy."""
    if df.empty:
        return df
    hoy = pd.Timestamp(date.today())
    return df[df["fecha_hora"].dt.normalize() == hoy]


def kpis_generales(df: pd.DataFrame) -> dict:
    """Devuelve los KPIs principales para las tarjetas del dashboard."""
    hoy = _solo_hoy(df)
    aprobados = int((df["estado"] == "Aprobado").sum()) if not df.empty else 0
    total = len(df)

    return {
        "total_registros": total,
        "registros_hoy": len(hoy),
        "vendedores_activos": int(df["codigo_vendedor"].nunique()) if not df.empty else 0,
        "valor_total": float(df["valor"].sum()) if not df.empty else 0.0,
        "valor_hoy": float(hoy["valor"].sum()) if not hoy.empty else 0.0,
        "pendientes": int((df["estado"] == "Pendiente").sum()) if not df.empty else 0,
        # % de cumplimiento = aprobados / total. Indicador de control clave.
        "tasa_aprobacion": (aprobados / total * 100) if total else 0.0,
    }


def por_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    """Resumen de registros y valor agrupado por vendedor."""
    if df.empty:
        return pd.DataFrame(columns=["codigo_vendedor", "nombre", "registros", "valor_total"])
    g = (
        df.groupby(["codigo_vendedor", "nombre"], dropna=False)
        .agg(registros=("cedula", "count"), valor_total=("valor", "sum"))
        .reset_index()
        .sort_values("registros", ascending=False)
    )
    return g


def por_estado(df: pd.DataFrame) -> pd.DataFrame:
    """Distribución de registros por estado documental."""
    if df.empty:
        return pd.DataFrame({"estado": config.ESTADOS, "cantidad": [0] * len(config.ESTADOS)})
    conteo = df["estado"].value_counts().reindex(config.ESTADOS, fill_value=0)
    return conteo.rename_axis("estado").reset_index(name="cantidad")


def por_combinacion(df: pd.DataFrame) -> pd.DataFrame:
    """Distribución por combinación contable."""
    if df.empty:
        return pd.DataFrame(columns=["combinacion_contable", "cantidad", "valor_total"])
    g = (
        df.groupby("combinacion_contable", dropna=False)
        .agg(cantidad=("cedula", "count"), valor_total=("valor", "sum"))
        .reset_index()
        .sort_values("cantidad", ascending=False)
    )
    return g


def tendencia_por_hora(df: pd.DataFrame) -> pd.DataFrame:
    """Cantidad de registros por hora del día de HOY (ritmo de la jornada)."""
    hoy = _solo_hoy(df)
    base = pd.DataFrame({"hora": range(24), "registros": [0] * 24})
    if hoy.empty:
        return base
    conteo = hoy.groupby(hoy["fecha_hora"].dt.hour).size()
    base["registros"] = base["hora"].map(conteo).fillna(0).astype(int)
    return base


def alertas_control(df: pd.DataFrame) -> list[dict]:
    """Genera alertas de control documental (calidad de los datos).

    Cada alerta es {"tipo": "warning|error|info", "mensaje": str}.
    """
    alertas: list[dict] = []
    if df.empty:
        return [{"tipo": "info", "mensaje": "Aún no hay registros capturados."}]

    # Cédulas duplicadas (posible doble captura).
    duplicados = df[df.duplicated("cedula", keep=False)]
    if not duplicados.empty:
        cedulas = ", ".join(sorted(duplicados["cedula"].astype(str).unique())[:10])
        alertas.append({
            "tipo": "warning",
            "mensaje": f"Cédulas con más de un registro: {cedulas}.",
        })

    # Documentos rechazados.
    rechazados = int((df["estado"] == "Rechazado").sum())
    if rechazados:
        alertas.append({"tipo": "error", "mensaje": f"{rechazados} documento(s) en estado Rechazado."})

    # Documentos pendientes acumulados.
    pendientes = int((df["estado"] == "Pendiente").sum())
    if pendientes:
        alertas.append({"tipo": "warning", "mensaje": f"{pendientes} documento(s) Pendiente(s) por gestionar."})

    # Registros sin información relevante.
    sin_info = int(df["informacion_relevante"].fillna("").str.strip().eq("").sum())
    if sin_info:
        alertas.append({"tipo": "info", "mensaje": f"{sin_info} registro(s) sin información relevante."})

    if not alertas:
        alertas.append({"tipo": "info", "mensaje": "Sin alertas: todos los controles en orden."})
    return alertas
