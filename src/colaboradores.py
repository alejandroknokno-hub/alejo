"""Versión Global: gestión de la información del colaborador.

Equivalente a tu aplicativo "Global": aquí se administra la ficha de cada
colaborador (código de vendedor, cédula, nombre, cargo, área, contacto, estado).

Se guarda en su propio Excel (data/global_colaboradores.xlsx) y se puede cruzar
con los registros de control documental a través de la cédula o el código de
vendedor para ver la actividad de cada persona.
"""
from __future__ import annotations

import re
from datetime import date

import pandas as pd

from . import config


# --- Inicialización --------------------------------------------------------
def asegurar_archivo() -> None:
    """Crea el Excel de colaboradores con encabezados si no existe."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not config.GLOBAL_PATH.exists():
        df = pd.DataFrame(columns=config.COLAB_ORDER)
        df.to_excel(config.GLOBAL_PATH, sheet_name=config.GLOBAL_SHEET, index=False)


# --- Lectura ---------------------------------------------------------------
def leer_colaboradores() -> pd.DataFrame:
    """Devuelve todos los colaboradores como DataFrame."""
    asegurar_archivo()
    df = pd.read_excel(config.GLOBAL_PATH, sheet_name=config.GLOBAL_SHEET, engine="openpyxl")
    for col in config.COLAB_ORDER:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[config.COLAB_ORDER]
    # La cédula como texto para no perder ceros ni convertir a notación científica.
    if not df.empty:
        df["cedula"] = df["cedula"].astype("string").str.strip()
        df["codigo_vendedor"] = df["codigo_vendedor"].astype("string").str.strip()
    return df


# --- Validación ------------------------------------------------------------
def validar(colab: dict, existentes: pd.DataFrame | None = None) -> list[str]:
    """Valida la ficha del colaborador. Devuelve lista de errores (vacía si OK)."""
    errores: list[str] = []

    cedula = str(colab.get("cedula", "")).strip()
    if not cedula:
        errores.append("La cédula es obligatoria.")
    elif not re.fullmatch(r"\d{5,15}", cedula):
        errores.append("La cédula debe contener solo números (5 a 15 dígitos).")

    if not str(colab.get("nombre", "")).strip():
        errores.append("El nombre es obligatorio.")

    if not str(colab.get("codigo_vendedor", "")).strip():
        errores.append("El código de vendedor es obligatorio.")

    correo = str(colab.get("correo", "")).strip()
    if correo and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", correo):
        errores.append("El correo no tiene un formato válido.")

    estado = colab.get("estado", "")
    if estado not in config.ESTADOS_COLAB:
        errores.append(f"El estado debe ser uno de: {', '.join(config.ESTADOS_COLAB)}.")

    # Evitar cédulas duplicadas.
    if existentes is not None and not existentes.empty and cedula:
        if cedula in existentes["cedula"].astype("string").str.strip().values:
            errores.append(f"Ya existe un colaborador con la cédula {cedula}.")

    return errores


# --- Escritura -------------------------------------------------------------
def agregar_colaborador(colab: dict) -> dict:
    """Agrega o actualiza la ficha de un colaborador (clave: cédula).

    Devuelve {"ok": bool, "errores": list[str], "accion": "creado"|"actualizado"}.
    """
    df = leer_colaboradores()
    cedula = str(colab.get("cedula", "")).strip()
    existe = (not df.empty) and (cedula in df["cedula"].astype("string").str.strip().values)

    # En actualización no se valida el duplicado de su propia cédula.
    errores = validar(colab, existentes=None if existe else df)
    if errores:
        return {"ok": False, "errores": errores, "accion": None}

    ficha = {
        "codigo_vendedor": str(colab.get("codigo_vendedor", "")).strip(),
        "cedula": cedula,
        "nombre": str(colab.get("nombre", "")).strip(),
        "cargo": str(colab.get("cargo", "")).strip(),
        "area": str(colab.get("area", "")).strip(),
        "correo": str(colab.get("correo", "")).strip(),
        "telefono": str(colab.get("telefono", "")).strip(),
        "fecha_ingreso": str(colab.get("fecha_ingreso", "") or date.today().isoformat()),
        "estado": colab.get("estado", config.ESTADOS_COLAB[0]),
    }

    if existe:
        # Quitamos la ficha previa y la reinsertamos (evita conflictos de tipos).
        df = df[df["cedula"].astype("string").str.strip() != cedula]
        accion = "actualizado"
    else:
        accion = "creado"

    df = pd.concat([df, pd.DataFrame([ficha])], ignore_index=True)
    guardar(df)
    return {"ok": True, "errores": [], "accion": accion}


def guardar(df: pd.DataFrame) -> None:
    """Escribe el DataFrame de colaboradores en el Excel."""
    asegurar_archivo()
    df = df[config.COLAB_ORDER]
    df.to_excel(config.GLOBAL_PATH, sheet_name=config.GLOBAL_SHEET, index=False, engine="openpyxl")


# --- Cruce con control documental ------------------------------------------
def actividad_por_colaborador(df_colab: pd.DataFrame, df_registros: pd.DataFrame) -> pd.DataFrame:
    """Cruza colaboradores con sus registros documentales (por cédula).

    Agrega columnas: cantidad de registros y valor total gestionado por persona.
    """
    base = df_colab.copy()
    if base.empty:
        return base

    if df_registros is None or df_registros.empty:
        base["registros"] = 0
        base["valor_gestionado"] = 0.0
        return base

    reg = df_registros.copy()
    reg["cedula"] = reg["cedula"].astype("string").str.strip()
    agg = (
        reg.groupby("cedula")
        .agg(registros=("cedula", "count"), valor_gestionado=("valor", "sum"))
        .reset_index()
    )
    base["cedula"] = base["cedula"].astype("string").str.strip()
    base = base.merge(agg, on="cedula", how="left")
    base["registros"] = base["registros"].fillna(0).astype(int)
    base["valor_gestionado"] = base["valor_gestionado"].fillna(0.0)
    return base
