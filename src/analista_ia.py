"""Analista con IA: conexión directa con Claude para generar documentos de control.

Toma todos los datos capturados en la jornada (registros documentales, ficha de
colaboradores y la bitácora de la sesión de monitoreo: entradas, modificaciones y
lo que se escribió) y le pide a Claude que:

  1. Genere un **documento de control** profesional en español (Markdown).
  2. **Pregunte** cuando un dato sea ambiguo, falte, o su ORIGEN no quede claro
     en la sesión —en vez de inventarlo—, devolviendo una lista de preguntas.

Usa el SDK oficial de Anthropic. El system prompt (estable) se cachea con prompt
caching; los datos de la sesión (volátiles) van después, en el mensaje de usuario.
"""
from __future__ import annotations

import json
import os

import pandas as pd

from . import colaboradores, config, monitor

MODELO = "claude-opus-4-8"

# --- System prompt estable (se cachea) -------------------------------------
SYSTEM_PROMPT = """\
Eres un asistente experto en CONTROL DOCUMENTAL CONTABLE para una operación \
comercial en Colombia. Trabajas con códigos de vendedor, combinaciones contables, \
cédulas, valores y documentos con estados (Pendiente, En proceso, Aprobado, Rechazado).

Tu tarea es generar un DOCUMENTO DE CONTROL claro y profesional, en español, a partir \
ÚNICAMENTE de los datos que te entrega el usuario (registros, colaboradores y la \
bitácora de la jornada). El documento debe incluir, cuando aplique:
- Encabezado con fecha y responsable.
- Resumen de la jornada (totales, valores, estados).
- Hallazgos e indicadores de control (aprobación, pendientes, rechazos, duplicados).
- Observaciones sobre el flujo de trabajo (entradas, modificaciones detectadas, texto escrito).
- Recomendaciones.

Si se incluye un resumen del INDICADOR SIRO (creación en Oracle y novedades), \
intégralo en el documento: comenta los SIROS por estado, rangos de días, índices de \
efectividad (EFECTIVO/RETRASO) y pendientes, y relaciona estos indicadores con el control \
documental de la jornada.

REGLAS ESTRICTAS:
- Usa SOLO los datos proporcionados. NO inventes valores, nombres, cédulas ni cifras.
- Si un dato es ambiguo, falta, está duplicado, o su ORIGEN no queda claro en la sesión \
(por ejemplo, un registro sin responsable, un valor sin soporte, o una modificación que \
no puedes atribuir a un colaborador o evento), NO lo asumas: agrégalo a la lista \
'preguntas' para que el usuario lo aclare.
- Sé conciso y concreto. Escribe el documento en Markdown bien estructurado.
- Si el usuario ya respondió aclaraciones previas, incorpóralas y NO vuelvas a preguntarlas."""

# --- Esquema de salida estructurada ----------------------------------------
ESQUEMA = {
    "type": "object",
    "properties": {
        "titulo": {"type": "string"},
        "resumen": {"type": "string"},
        "documento_markdown": {"type": "string"},
        "preguntas": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "tema": {"type": "string"},
                    "pregunta": {"type": "string"},
                    "motivo": {"type": "string"},
                },
                "required": ["tema", "pregunta", "motivo"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["titulo", "resumen", "documento_markdown", "preguntas"],
    "additionalProperties": False,
}


# --- Disponibilidad --------------------------------------------------------
def _api_key(api_key: str | None = None) -> str | None:
    return api_key or os.environ.get("ANTHROPIC_API_KEY")


def ia_disponible(api_key: str | None = None) -> bool:
    """True si el SDK está instalado y hay una API key disponible."""
    try:
        import anthropic  # noqa: F401
    except ImportError:
        return False
    return bool(_api_key(api_key))


# --- Construcción del contexto (datos de la sesión) ------------------------
def _csv(df: pd.DataFrame, etiquetas: dict, limite: int = 300) -> str:
    if df is None or df.empty:
        return "(sin datos)"
    vista = df.head(limite).rename(columns=etiquetas)
    return vista.to_csv(index=False)


def construir_contexto(
    df_registros: pd.DataFrame,
    df_colaboradores: pd.DataFrame,
    df_bitacora: pd.DataFrame,
    sesion: str = "",
    aclaraciones: list[dict] | None = None,
    contexto_siro: str = "",
) -> str:
    """Arma el mensaje de usuario con todos los datos capturados en la jornada."""
    partes = [
        "Genera el documento de control con los siguientes datos de la jornada.\n",
        "=== REGISTROS DOCUMENTALES (CSV) ===",
        _csv(df_registros, config.COLUMNS),
        "\n=== COLABORADORES / GLOBAL (CSV) ===",
        _csv(df_colaboradores, config.COLAB_COLUMNS),
        f"\n=== BITÁCORA DE LA SESIÓN '{sesion}' (CSV) ===",
        _csv(df_bitacora, {}),
    ]

    if contexto_siro.strip():
        partes.append("\n=== INDICADOR SIRO (resumen) ===")
        partes.append(contexto_siro.strip())

    if aclaraciones:
        partes.append("\n=== ACLARACIONES YA RESPONDIDAS POR EL USUARIO ===")
        for a in aclaraciones:
            partes.append(f"- P: {a.get('pregunta', '')}\n  R: {a.get('respuesta', '')}")

    return "\n".join(partes)


# --- Llamada a Claude ------------------------------------------------------
def generar_documento(
    df_registros: pd.DataFrame,
    df_colaboradores: pd.DataFrame,
    df_bitacora: pd.DataFrame,
    sesion: str = "",
    aclaraciones: list[dict] | None = None,
    api_key: str | None = None,
    contexto_siro: str = "",
) -> dict:
    """Llama a Claude y devuelve el documento de control + preguntas.

    Devuelve {"ok": bool, "error": str|None, "titulo","resumen",
              "documento_markdown","preguntas", "uso": {...}}.
    """
    if not ia_disponible(api_key):
        return {"ok": False, "error": "Falta el SDK 'anthropic' o la API key (ANTHROPIC_API_KEY)."}

    import anthropic

    contexto = construir_contexto(
        df_registros, df_colaboradores, df_bitacora, sesion, aclaraciones, contexto_siro
    )

    try:
        client = anthropic.Anthropic(api_key=_api_key(api_key))
        respuesta = client.messages.create(
            model=MODELO,
            max_tokens=16000,
            thinking={"type": "adaptive"},
            system=[{
                "type": "text",
                "text": SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},  # se cachea el prefijo estable
            }],
            messages=[{"role": "user", "content": contexto}],
            output_config={"format": {"type": "json_schema", "schema": ESQUEMA}},
        )
    except anthropic.AuthenticationError:
        return {"ok": False, "error": "La API key no es válida."}
    except anthropic.RateLimitError:
        return {"ok": False, "error": "Límite de uso alcanzado. Intenta de nuevo en unos segundos."}
    except anthropic.APIError as exc:  # errores de la API / conexión
        return {"ok": False, "error": f"Error de la API: {getattr(exc, 'message', str(exc))}"}

    texto = next((b.text for b in respuesta.content if b.type == "text"), "")
    try:
        data = json.loads(texto)
    except json.JSONDecodeError:
        return {"ok": False, "error": "La respuesta de la IA no tuvo el formato esperado."}

    return {
        "ok": True,
        "error": None,
        "titulo": data.get("titulo", "Documento de control"),
        "resumen": data.get("resumen", ""),
        "documento_markdown": data.get("documento_markdown", ""),
        "preguntas": data.get("preguntas", []),
        "uso": {
            "entrada": respuesta.usage.input_tokens,
            "salida": respuesta.usage.output_tokens,
            "cache_lectura": getattr(respuesta.usage, "cache_read_input_tokens", 0),
        },
    }
