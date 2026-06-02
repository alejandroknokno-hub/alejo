"""Pruebas básicas de la lógica de datos e indicadores.

Ejecuta:  python -m pytest   (o)   python tests/test_app.py
No requiere Streamlit; valida el núcleo (validación, Excel e indicadores).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import colaboradores, config, data_manager, grabador, indicadores, monitor  # noqa: E402


def _registro_valido(**extra):
    base = {
        "codigo_vendedor": "V-001",
        "nombre": "Ana Pérez",
        "cedula": "1023456789",
        "combinacion_contable": "5105-01-001",
        "valor": 50000,
        "estado": "Aprobado",
        "informacion_relevante": "Soporte adjunto",
    }
    base.update(extra)
    return base


def test_validacion_detecta_campos_obligatorios():
    errores = data_manager.validar_registro({"codigo_vendedor": "", "cedula": "abc"})
    assert any("vendedor" in e.lower() for e in errores)
    assert any("cédula" in e.lower() or "cedula" in e.lower() for e in errores)


def test_validacion_cedula_solo_numeros():
    errores = data_manager.validar_registro(_registro_valido(cedula="12-AB"))
    assert any("cédula" in e.lower() or "cedula" in e.lower() for e in errores)


def test_registro_valido_pasa():
    assert data_manager.validar_registro(_registro_valido()) == []


def test_ciclo_completo_excel(tmp_path, monkeypatch):
    # Redirige el Excel a una ruta temporal para no tocar datos reales.
    excel = tmp_path / "test.xlsx"
    monkeypatch.setattr(config, "EXCEL_PATH", excel)
    monkeypatch.setattr(config, "DATA_DIR", tmp_path)

    res = data_manager.agregar_registro(_registro_valido(), registrado_por="tester")
    assert res["ok"], res["errores"]

    df = data_manager.leer_registros()
    assert len(df) == 1
    assert df.iloc[0]["codigo_vendedor"] == "V-001"
    # La fecha (día/mes/año) se completa automáticamente.
    assert df.iloc[0]["mes"] in config.MESES_ES

    kpis = indicadores.kpis_generales(df)
    assert kpis["total_registros"] == 1
    assert kpis["tasa_aprobacion"] == 100.0


def test_comparar_imagenes_detecta_diferencia(tmp_path):
    """La comparación da ~0% en imágenes iguales y alto % en imágenes distintas."""
    from PIL import Image

    iguales_a = tmp_path / "a.png"
    iguales_b = tmp_path / "b.png"
    distinta = tmp_path / "c.png"

    Image.new("RGB", (200, 120), (255, 255, 255)).save(iguales_a)
    Image.new("RGB", (200, 120), (255, 255, 255)).save(iguales_b)
    Image.new("RGB", (200, 120), (0, 0, 0)).save(distinta)

    assert monitor.comparar_imagenes(str(iguales_a), str(iguales_b)) < 1.0
    assert monitor.comparar_imagenes(str(iguales_a), str(distinta)) > 90.0


def _usar_temp(tmp_path, monkeypatch):
    """Redirige todas las rutas de datos a una carpeta temporal."""
    monkeypatch.setattr(config, "DATA_DIR", tmp_path)
    monkeypatch.setattr(config, "EXCEL_PATH", tmp_path / "registros.xlsx")
    monkeypatch.setattr(config, "GLOBAL_PATH", tmp_path / "colab.xlsx")
    monkeypatch.setattr(config, "MONITOR_LOG", tmp_path / "monitor_log.csv")
    monkeypatch.setattr(config, "CAPTURAS_DIR", tmp_path / "capturas")


# --- Versión Oracle: monitoreo ---------------------------------------------
def test_monitor_bitacora_y_flujo(tmp_path, monkeypatch):
    _usar_temp(tmp_path, monkeypatch)

    monitor.registrar_evento("Entrada", "Abro Oracle", sesion="s1")
    monitor.registrar_evento("Modificación", "Cambio valor", sesion="s1")
    monitor.registrar_evento("Entrada", "Otra sesión", sesion="s2")

    log_s1 = monitor.leer_log("s1")
    assert len(log_s1) == 2

    flujo = monitor.analizar_flujo("s1")
    assert flujo["eventos"] == 2
    assert flujo["entradas"] == 1
    assert flujo["modificaciones"] == 1
    assert "eventos" in flujo["resumen"].lower()


# --- Versión Global: colaboradores -----------------------------------------
def test_colaborador_crear_y_validar(tmp_path, monkeypatch):
    _usar_temp(tmp_path, monkeypatch)

    # Cédula inválida -> error.
    res = colaboradores.agregar_colaborador({
        "codigo_vendedor": "V-001", "cedula": "AB", "nombre": "X", "estado": "Activo",
    })
    assert not res["ok"]

    # Alta válida.
    res = colaboradores.agregar_colaborador({
        "codigo_vendedor": "V-001", "cedula": "1023456789", "nombre": "Ana",
        "cargo": "Asesora", "area": "Cartera", "estado": "Activo",
    })
    assert res["ok"] and res["accion"] == "creado"

    # Misma cédula -> actualiza, no duplica.
    res = colaboradores.agregar_colaborador({
        "codigo_vendedor": "V-001", "cedula": "1023456789", "nombre": "Ana María",
        "estado": "Activo",
    })
    assert res["ok"] and res["accion"] == "actualizado"

    df = colaboradores.leer_colaboradores()
    assert len(df) == 1
    assert df.iloc[0]["nombre"] == "Ana María"


def test_cruce_colaborador_con_registros(tmp_path, monkeypatch):
    _usar_temp(tmp_path, monkeypatch)

    colaboradores.agregar_colaborador({
        "codigo_vendedor": "V-001", "cedula": "1023456789", "nombre": "Ana", "estado": "Activo",
    })
    data_manager.agregar_registro(_registro_valido(cedula="1023456789"))

    df_colab = colaboradores.leer_colaboradores()
    df_reg = data_manager.leer_registros()
    cruce = colaboradores.actividad_por_colaborador(df_colab, df_reg)
    assert int(cruce.iloc[0]["registros"]) == 1


# --- Grabador tipo video ---------------------------------------------------
def test_texto_desde_teclas_reconstruye_y_corrige():
    teclas = list("hola") + ["BACKSPACE", "SPACE"] + list("mundo")
    # "hola" -> borra 'a' -> "hol" -> espacio -> "hol " -> "mundo"
    assert grabador.texto_desde_teclas(teclas) == "hol mundo"


def test_ensamblar_video_genera_gif(tmp_path, monkeypatch):
    _usar_temp(tmp_path, monkeypatch)
    from PIL import Image

    config.CAPTURAS_DIR.mkdir(parents=True, exist_ok=True)
    for i in range(3):
        nombre = f"frame_{i}.png"
        Image.new("RGB", (80, 60), (i * 80, 0, 0)).save(config.CAPTURAS_DIR / nombre)
        monitor.registrar_evento("Captura", "frame", sesion="vid", archivo=nombre)

    res = grabador.ensamblar_video("vid", fps=2)
    assert res["ok"], res["error"]
    assert res["frames"] == 3
    assert Path(res["archivo"]).exists()


if __name__ == "__main__":
    import subprocess

    raise SystemExit(subprocess.call(["python", "-m", "pytest", "-v", __file__]))
