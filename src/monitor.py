"""Versión Oracle: monitoreo del trabajo por captura de pantalla.

Objetivo: dejar evidencia y análisis de las ENTRADAS y MODIFICACIONES que haces
en tu aplicativo de trabajo (Oracle) durante la jornada.

Cómo funciona:
- `capturar_pantalla()` toma una foto de la pantalla y la guarda en /data/capturas.
- `registrar_evento()` anota en una bitácora (CSV) cada captura, entrada o modificación.
- `analizar_flujo()` revisa la bitácora de una "sesión de prueba" y describe cómo es
  tu flujo de trabajo (duración, ritmo, intervalos, tipos de actividad).

La captura usa la librería `mss` (multiplataforma). Si se ejecuta en un servidor
sin pantalla, las funciones lo informan con un mensaje claro en vez de fallar.
"""
from __future__ import annotations

import csv
from datetime import datetime

import pandas as pd

from . import config


# --- Disponibilidad --------------------------------------------------------
def captura_disponible() -> bool:
    """Indica si el entorno puede tomar capturas (hay pantalla y librería mss)."""
    try:
        import mss  # noqa: F401
    except ImportError:
        return False
    try:
        import mss as _mss

        with _mss.mss() as sct:
            return len(sct.monitors) > 0
    except Exception:
        # Entorno sin display (p. ej. servidor headless).
        return False


# --- Bitácora --------------------------------------------------------------
def registrar_evento(tipo: str, descripcion: str, sesion: str = "", archivo: str = "") -> None:
    """Anota un evento (Captura/Entrada/Modificación/...) en el log CSV."""
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    nuevo = not config.MONITOR_LOG.exists()
    with open(config.MONITOR_LOG, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if nuevo:
            writer.writerow(config.MONITOR_COLUMNS)
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            sesion.strip(),
            tipo,
            descripcion.strip(),
            archivo,
        ])


def leer_log(sesion: str | None = None) -> pd.DataFrame:
    """Devuelve la bitácora como DataFrame, opcionalmente filtrada por sesión."""
    if not config.MONITOR_LOG.exists():
        return pd.DataFrame(columns=config.MONITOR_COLUMNS)
    df = pd.read_csv(config.MONITOR_LOG)
    for col in config.MONITOR_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[config.MONITOR_COLUMNS]
    df["fecha_hora"] = pd.to_datetime(df["fecha_hora"], errors="coerce")
    if sesion:
        df = df[df["sesion"] == sesion]
    return df


# --- Captura ---------------------------------------------------------------
def capturar_pantalla(sesion: str = "", descripcion: str = "Captura automática") -> dict:
    """Toma una captura de pantalla y la registra en la bitácora.

    Devuelve {"ok": bool, "archivo": str | None, "error": str | None}.
    """
    if not captura_disponible():
        return {
            "ok": False,
            "archivo": None,
            "error": "La captura no está disponible en este entorno "
                     "(ejecuta la app en tu equipo de trabajo con pantalla y "
                     "con la librería 'mss' instalada).",
        }

    import mss
    import mss.tools

    config.CAPTURAS_DIR.mkdir(parents=True, exist_ok=True)
    ahora = datetime.now()
    nombre = f"captura_{ahora:%Y%m%d_%H%M%S_%f}.png"
    ruta = config.CAPTURAS_DIR / nombre

    with mss.mss() as sct:
        # monitors[0] = todos los monitores combinados (pantalla completa).
        pantalla = sct.monitors[0]
        imagen = sct.grab(pantalla)
        mss.tools.to_png(imagen.rgb, imagen.size, output=str(ruta))

    registrar_evento("Captura", descripcion, sesion=sesion, archivo=nombre)
    return {"ok": True, "archivo": str(ruta), "error": None}


def _ruta_ultima_captura(sesion: str | None = None) -> str | None:
    """Devuelve la ruta de la última captura de la sesión, o None si no hay."""
    recientes = capturas_recientes(sesion, limite=1)
    return recientes[0]["ruta"] if recientes else None


# --- Detección automática de cambios ---------------------------------------
def comparar_imagenes(ruta_a: str, ruta_b: str, tam: tuple[int, int] = (320, 180)) -> float:
    """Compara dos capturas y devuelve el % de pantalla que cambió (0-100).

    Las imágenes se reducen y pasan a escala de grises para que la comparación
    sea rápida y tolerante a pequeñas diferencias (cursor, antialias). Un píxel
    se considera "cambiado" si su brillo varía más de 25 niveles.
    """
    from PIL import Image
    import numpy as np

    a = Image.open(ruta_a).convert("L").resize(tam)
    b = Image.open(ruta_b).convert("L").resize(tam)
    arr_a = np.asarray(a, dtype="int16")
    arr_b = np.asarray(b, dtype="int16")
    diferencia = np.abs(arr_a - arr_b)
    cambiados = int(np.count_nonzero(diferencia > 25))
    return cambiados / diferencia.size * 100.0


def capturar_y_detectar(
    sesion: str = "",
    umbral_pct: float = 2.0,
    descripcion: str = "Captura automática",
) -> dict:
    """Toma una captura y detecta automáticamente si hubo un cambio respecto a la anterior.

    Si el cambio supera `umbral_pct`, registra un evento "Modificación" automático.

    Devuelve {"ok", "archivo", "cambio_detectado", "diferencia_pct", "primera", "error"}.
    """
    previa = _ruta_ultima_captura(sesion)
    res = capturar_pantalla(sesion=sesion, descripcion=descripcion)
    if not res["ok"]:
        return {**res, "cambio_detectado": False, "diferencia_pct": 0.0, "primera": False}

    nueva = res["archivo"]
    if previa is None:
        # Primera captura de la sesión: no hay con qué comparar.
        return {**res, "cambio_detectado": False, "diferencia_pct": 0.0, "primera": True}

    try:
        diff = comparar_imagenes(previa, nueva)
    except Exception as exc:  # pragma: no cover - depende de Pillow/IO
        return {**res, "cambio_detectado": False, "diferencia_pct": 0.0,
                "primera": False, "error": f"No se pudo comparar: {exc}"}

    detectado = diff >= umbral_pct
    if detectado:
        from pathlib import Path

        registrar_evento(
            "Modificación",
            f"Cambio detectado automáticamente ({diff:.1f}% de la pantalla)",
            sesion=sesion,
            archivo=Path(nueva).name,
        )
    return {
        "ok": True,
        "archivo": nueva,
        "cambio_detectado": detectado,
        "diferencia_pct": diff,
        "primera": False,
        "error": None,
    }


def capturas_recientes(sesion: str | None = None, limite: int = 12) -> list[dict]:
    """Lista las últimas capturas (ruta + fecha) para mostrar en la galería."""
    df = leer_log(sesion)
    df = df[df["tipo"] == "Captura"].dropna(subset=["archivo"])
    df = df[df["archivo"].astype(str).str.strip() != ""]
    df = df.sort_values("fecha_hora", ascending=False).head(limite)

    resultado = []
    for _, fila in df.iterrows():
        ruta = config.CAPTURAS_DIR / str(fila["archivo"])
        if ruta.exists():
            resultado.append({
                "ruta": str(ruta),
                "fecha_hora": fila["fecha_hora"],
                "descripcion": fila["descripcion"],
            })
    return resultado


# --- Análisis del flujo de trabajo -----------------------------------------
def analizar_flujo(sesion: str | None = None) -> dict:
    """Analiza la bitácora para describir cómo es el flujo de trabajo.

    Devuelve un resumen con métricas y un DataFrame de actividad por minuto,
    ideal para la "sesión de prueba" donde se observa el ritmo de trabajo.
    """
    df = leer_log(sesion).dropna(subset=["fecha_hora"]).sort_values("fecha_hora")

    if df.empty:
        return {
            "eventos": 0,
            "capturas": 0,
            "entradas": 0,
            "modificaciones": 0,
            "inicio": None,
            "fin": None,
            "duracion_min": 0.0,
            "intervalo_prom_seg": 0.0,
            "por_tipo": pd.DataFrame(columns=["tipo", "cantidad"]),
            "linea_tiempo": pd.DataFrame(columns=["minuto", "eventos"]),
            "resumen": "Sin eventos registrados en esta sesión.",
        }

    inicio, fin = df["fecha_hora"].min(), df["fecha_hora"].max()
    duracion_min = (fin - inicio).total_seconds() / 60.0

    # Intervalo promedio entre eventos (ritmo de trabajo).
    difs = df["fecha_hora"].diff().dropna().dt.total_seconds()
    intervalo_prom = float(difs.mean()) if not difs.empty else 0.0

    por_tipo = df["tipo"].value_counts().rename_axis("tipo").reset_index(name="cantidad")

    # Línea de tiempo: eventos por minuto.
    linea = (
        df.set_index("fecha_hora")
        .resample("1min")
        .size()
        .rename("eventos")
        .reset_index()
        .rename(columns={"fecha_hora": "minuto"})
    )

    entradas = int((df["tipo"] == "Entrada").sum())
    modificaciones = int((df["tipo"] == "Modificación").sum())
    capturas = int((df["tipo"] == "Captura").sum())

    resumen = (
        f"Se registraron {len(df)} eventos en {duracion_min:.1f} minutos "
        f"({entradas} entradas, {modificaciones} modificaciones, {capturas} capturas). "
        f"En promedio, una acción cada {intervalo_prom:.0f} segundos."
    )

    return {
        "eventos": len(df),
        "capturas": capturas,
        "entradas": entradas,
        "modificaciones": modificaciones,
        "inicio": inicio,
        "fin": fin,
        "duracion_min": duracion_min,
        "intervalo_prom_seg": intervalo_prom,
        "por_tipo": por_tipo,
        "linea_tiempo": linea,
        "resumen": resumen,
    }
