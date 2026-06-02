"""Grabador de sesión tipo "video" del flujo de trabajo completo.

Graba, paso a paso, TODO lo que ocurre durante tu jornada en el aplicativo:
- **Fotogramas de pantalla** tomados en intervalos cortos (como un video).
- **Lo que escribes** (texto del teclado), registrado en la misma bitácora con
  marca de tiempo, para reconstruir qué datos ingresaste y modificaste.
- Al final, los fotogramas se **ensamblan en un GIF** reproducible.

Importante (privacidad y control):
- La grabación es **opt-in**: la inicias y la detienes tú manualmente.
- Solo registra TU propia sesión, en TU equipo. Mientras está activa, la app
  muestra un indicador visible de "grabando".
- El registro de teclado usa `pynput`; en macOS puede pedir permisos de
  Accesibilidad la primera vez. Si `pynput` no está disponible, se graban los
  fotogramas igual (sin el texto).
"""
from __future__ import annotations

import threading
import time
from datetime import datetime

from . import config, monitor

# Estado global del grabador (uno por proceso). Persiste entre reruns de Streamlit.
_estado: dict = {
    "activo": False,
    "sesion": "",
    "hilo": None,
    "listener": None,
    "buffer": [],          # caracteres escritos aún no volcados a la bitácora
    "inicio": None,
    "intervalo": 2.0,
}


# --- Reconstrucción de texto (lógica pura, testeable) ----------------------
def texto_desde_teclas(teclas: list[str]) -> str:
    """Reconstruye el texto escrito a partir de una lista de teclas normalizadas.

    Tokens especiales: "SPACE", "BACKSPACE". El resto se asume como un carácter.
    """
    buffer: list[str] = []
    for t in teclas:
        if t == "BACKSPACE":
            if buffer:
                buffer.pop()
        elif t == "SPACE":
            buffer.append(" ")
        elif len(t) == 1:
            buffer.append(t)
    return "".join(buffer)


def _flush_texto(sesion: str) -> None:
    """Vuelca el texto acumulado en el buffer a la bitácora como un evento 'Texto'."""
    buf = _estado["buffer"]
    if buf:
        texto = "".join(buf).strip()
        if texto:
            monitor.registrar_evento("Texto", f"Escribió: {texto}", sesion=sesion)
        buf.clear()


# --- Hilos de captura ------------------------------------------------------
def _bucle_frames(sesion: str, intervalo: float) -> None:
    """Toma un fotograma cada `intervalo` segundos mientras la grabación esté activa."""
    while _estado["activo"]:
        monitor.capturar_pantalla(sesion=sesion, descripcion="Fotograma de grabación")
        time.sleep(max(0.2, intervalo))


def _iniciar_teclado(sesion: str):
    """Inicia el listener de teclado (pynput). Devuelve el listener o None."""
    try:
        from pynput import keyboard
        from pynput.keyboard import Key
    except Exception:
        return None

    def on_press(key):
        buf = _estado["buffer"]
        if key in (Key.enter, Key.tab):
            _flush_texto(sesion)            # un salto de línea / tab cierra el campo
        elif key == Key.space:
            buf.append(" ")
        elif key == Key.backspace:
            if buf:
                buf.pop()
        else:
            try:
                buf.append(key.char)        # carácter imprimible
            except AttributeError:
                pass                        # teclas especiales: se ignoran

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    return listener


# --- Control de la grabación -----------------------------------------------
def iniciar(sesion: str, intervalo_frames: float = 2.0) -> dict:
    """Inicia la grabación (fotogramas + teclado). Devuelve {'ok', 'error'}."""
    if _estado["activo"]:
        return {"ok": False, "error": "Ya hay una grabación en curso."}
    if not monitor.captura_disponible():
        return {"ok": False, "error": "No hay pantalla disponible para grabar en este entorno."}

    _estado.update(
        activo=True, sesion=sesion, buffer=[],
        inicio=datetime.now(), intervalo=intervalo_frames,
    )
    monitor.registrar_evento("Nota", "▶️ Inicio de grabación de sesión", sesion=sesion)

    hilo = threading.Thread(target=_bucle_frames, args=(sesion, intervalo_frames), daemon=True)
    hilo.start()
    _estado["hilo"] = hilo
    _estado["listener"] = _iniciar_teclado(sesion)
    return {"ok": True, "error": None}


def detener() -> dict:
    """Detiene la grabación y vuelca el último texto pendiente."""
    if not _estado["activo"]:
        return {"ok": False, "error": "No hay ninguna grabación activa."}

    sesion = _estado["sesion"]
    _estado["activo"] = False
    if _estado["listener"] is not None:
        try:
            _estado["listener"].stop()
        except Exception:
            pass
    _flush_texto(sesion)
    monitor.registrar_evento("Nota", "⏹️ Fin de grabación de sesión", sesion=sesion)

    _estado["listener"] = None
    _estado["hilo"] = None
    return {"ok": True, "error": None, "sesion": sesion}


def estado() -> dict:
    """Estado actual (seguro de serializar) para mostrar en la interfaz."""
    activo = _estado["activo"]
    inicio = _estado["inicio"]
    duracion = (datetime.now() - inicio).total_seconds() if (activo and inicio) else 0.0
    return {
        "activo": activo,
        "sesion": _estado["sesion"],
        "duracion_seg": duracion,
        "intervalo": _estado["intervalo"],
        "teclado": _estado["listener"] is not None,
    }


# --- Ensamblado del "video" (GIF) ------------------------------------------
def ensamblar_video(sesion: str, fps: int = 2) -> dict:
    """Une los fotogramas de la sesión en un GIF reproducible.

    Usa solo Pillow (sin ffmpeg). Devuelve {'ok', 'archivo', 'frames', 'error'}.
    """
    from PIL import Image

    # capturas_recientes devuelve de más nueva a más antigua; invertimos a orden cronológico.
    capturas = monitor.capturas_recientes(sesion, limite=100000)
    rutas = [c["ruta"] for c in reversed(capturas)]
    if not rutas:
        return {"ok": False, "archivo": None, "frames": 0,
                "error": "No hay fotogramas en esta sesión para ensamblar."}

    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    salida = config.DATA_DIR / f"grabacion_{sesion}.gif"

    imagenes = [Image.open(r).convert("RGB") for r in rutas]
    # Normalizamos todos los fotogramas al tamaño del primero.
    base = imagenes[0].size
    imagenes = [img if img.size == base else img.resize(base) for img in imagenes]

    duracion_ms = int(1000 / max(1, fps))
    imagenes[0].save(
        salida, save_all=True, append_images=imagenes[1:],
        duration=duracion_ms, loop=0, optimize=True,
    )
    return {"ok": True, "archivo": str(salida), "frames": len(rutas), "error": None}
