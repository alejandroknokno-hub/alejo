"""Pruebas básicas de la lógica de datos e indicadores.

Ejecuta:  python -m pytest   (o)   python tests/test_app.py
No requiere Streamlit; valida el núcleo (validación, Excel e indicadores).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import config, data_manager, indicadores  # noqa: E402


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


if __name__ == "__main__":
    import subprocess

    raise SystemExit(subprocess.call(["python", "-m", "pytest", "-v", __file__]))
