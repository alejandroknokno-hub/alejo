"""Pruebas básicas de la lógica de datos e indicadores.

Ejecuta:  python -m pytest   (o)   python tests/test_app.py
No requiere Streamlit; valida el núcleo (validación, Excel e indicadores).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import (  # noqa: E402
    analista_ia,
    colaboradores,
    config,
    consolidacion,
    data_manager,
    grabador,
    indicador_siro,
    indicadores,
    monitor,
)


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
    monkeypatch.setattr(config, "CONSOLIDADO_PATH", tmp_path / "consolidacion_de_jornada.xlsx")
    monkeypatch.setattr(config, "SIRO_CACHE_PATH", tmp_path / "siro_cache.xlsx")


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


# --- Analista IA (sin red) -------------------------------------------------
def test_ia_no_disponible_sin_api_key(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    assert analista_ia.ia_disponible(api_key=None) is False


def test_generar_documento_sin_key_devuelve_error(monkeypatch):
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    df = data_manager.leer_registros()  # vacío, no importa
    res = analista_ia.generar_documento(df, df, df, sesion="s", api_key=None)
    assert res["ok"] is False
    assert "key" in res["error"].lower() or "anthropic" in res["error"].lower()


def test_construir_contexto_incluye_datos_y_aclaraciones():
    import pandas as pd

    registros = pd.DataFrame([{"codigo_vendedor": "V-1", "cedula": "123"}])
    ctx = analista_ia.construir_contexto(
        registros, registros.iloc[:0], registros.iloc[:0],
        sesion="sX",
        aclaraciones=[{"pregunta": "¿Origen del valor?", "respuesta": "Factura 001"}],
        contexto_siro="SIROS creación Oracle: 10 registros.",
    )
    assert "REGISTROS DOCUMENTALES" in ctx
    assert "sX" in ctx
    assert "Factura 001" in ctx
    assert "INDICADOR SIRO" in ctx and "10 registros" in ctx


def test_siro_resumen_texto():
    import pandas as pd
    inf = pd.DataFrame([
        {indicador_siro.C_ESTADO: "CERRADO", indicador_siro.C_AREA: "USEM",
         indicador_siro.C_RANGO_CREACION: "Entre 0 a 5 Dias"},
        {indicador_siro.C_ESTADO: "PENDIENTE", indicador_siro.C_AREA: "USEM",
         indicador_siro.C_RANGO_CREACION: "Mayor a 15 Dias"},
    ])
    texto = indicador_siro.resumen_texto({"informatica": inf})
    assert "2 registros" in texto
    assert "CERRADO" in texto and "USEM" in texto


# --- Consolidación de jornada (Excel único) --------------------------------
def test_consolidacion_unico_documento_acumula(tmp_path, monkeypatch):
    _usar_temp(tmp_path, monkeypatch)

    doc1 = {"titulo": "D1", "resumen": "r1", "documento_markdown": "# uno", "preguntas": []}
    consolidacion.guardar_documento(doc1, sesion="s1", usuario="u",
                                    kpis={"total_registros": 3, "valor_total": 100.0})
    doc2 = {"titulo": "D2", "resumen": "r2", "documento_markdown": "# dos",
            "preguntas": [{"tema": "t", "pregunta": "p", "motivo": "m"}]}
    consolidacion.guardar_documento(doc2, sesion="s2", usuario="u")

    df = consolidacion.leer_consolidado()
    assert len(df) == 2                      # acumula en un único archivo
    assert config.CONSOLIDADO_PATH.exists()  # un solo Excel, no separado por fecha

    # Regenerar la sesión s2 actualiza su última fila en vez de duplicar.
    doc2b = {"titulo": "D2b", "resumen": "r", "documento_markdown": "# dos v2", "preguntas": []}
    consolidacion.guardar_documento(doc2b, sesion="s2", usuario="u", actualizar=True)
    df = consolidacion.leer_consolidado()
    assert len(df) == 2
    assert (df["titulo"] == "D2b").any()
    assert not (df["titulo"] == "D2").any()


def test_consolidacion_tiene_varias_hojas(tmp_path, monkeypatch):
    import pandas as pd

    _usar_temp(tmp_path, monkeypatch)
    data_manager.agregar_registro(_registro_valido())
    colaboradores.agregar_colaborador({
        "codigo_vendedor": "V-001", "cedula": "1023456789", "nombre": "Ana", "estado": "Activo",
    })
    monitor.registrar_evento("Entrada", "abro oracle", sesion="s1")

    doc = {"titulo": "D", "resumen": "r", "documento_markdown": "# x", "preguntas": []}
    consolidacion.guardar_documento(
        doc, sesion="s1", usuario="u",
        df_registros=data_manager.leer_registros(),
        df_colaboradores=colaboradores.leer_colaboradores(),
        df_bitacora=monitor.leer_log("s1"),
    )

    hojas = pd.ExcelFile(config.CONSOLIDADO_PATH).sheet_names
    assert config.CONSOLIDADO_SHEET in hojas
    assert consolidacion.HOJA_REGISTROS in hojas
    assert consolidacion.HOJA_COLABORADORES in hojas
    assert consolidacion.HOJA_BITACORA in hojas

    # La hoja de registros trae el detalle (cédula del registro capturado).
    reg = pd.read_excel(config.CONSOLIDADO_PATH, sheet_name=consolidacion.HOJA_REGISTROS)
    assert "Cédula" in reg.columns
    assert len(reg) == 1


def test_consolidacion_hojas_analiticas_y_kpis(tmp_path, monkeypatch):
    import pandas as pd

    _usar_temp(tmp_path, monkeypatch)
    data_manager.agregar_registro(_registro_valido())  # estado Aprobado

    doc = {"titulo": "D", "resumen": "r", "documento_markdown": "# x", "preguntas": []}
    consolidacion.guardar_documento(
        doc, sesion="s1", df_registros=data_manager.leer_registros(),
    )

    hojas = pd.ExcelFile(config.CONSOLIDADO_PATH).sheet_names
    assert consolidacion.HOJA_KPIS in hojas
    assert consolidacion.HOJA_PIVOTES in hojas
    assert consolidacion.HOJA_DASHBOARD in hojas

    # La hoja KPIs refleja el indicador (1 registro aprobado -> 100% aprobación).
    kpi = pd.read_excel(config.CONSOLIDADO_PATH, sheet_name=consolidacion.HOJA_KPIS, header=None)
    texto = " ".join(kpi.fillna("").astype(str).values.ravel().tolist())
    assert "Tasa de aprobación" in texto
    assert "Registros totales" in texto


def test_siro_guardado_en_consolidacion(tmp_path, monkeypatch):
    import pandas as pd

    _usar_temp(tmp_path, monkeypatch)
    data_manager.agregar_registro(_registro_valido())

    inf = pd.DataFrame([{
        indicador_siro.C_ID: 1, indicador_siro.C_ESTADO: "CERRADO",
        indicador_siro.C_AREA: "USEM", indicador_siro.C_RANGO_CREACION: "Entre 0 a 5 Dias",
    }])
    nov = pd.DataFrame([{
        indicador_siro.N_ID: 2, indicador_siro.N_ESTADO: "CERRADO",
        indicador_siro.N_INDICE: "EFECTIVO", indicador_siro.N_RANGO_TI: "Entre 0 a 2 Dias",
    }])

    res = consolidacion.guardar_siro_en_consolidado(
        {"informatica": inf, "novedades": nov},
        df_registros=data_manager.leer_registros(),
    )
    assert res["ok"]

    hojas = pd.ExcelFile(config.CONSOLIDADO_PATH).sheet_names
    # Hojas SIRO + gráficos conviven con las hojas previas de la consolidación.
    assert "SIRO Informatica" in hojas
    assert "SIRO Novedades" in hojas
    assert consolidacion.HOJA_SIRO_GRAFICOS in hojas
    assert consolidacion.HOJA_REGISTROS in hojas
    assert config.CONSOLIDADO_SHEET in hojas


def test_consolidacion_pivote_interactivo_inyectado(tmp_path, monkeypatch):
    import zipfile
    import pandas as pd

    _usar_temp(tmp_path, monkeypatch)
    data_manager.agregar_registro(_registro_valido())

    consolidacion.guardar_documento(
        {"titulo": "D", "resumen": "r", "documento_markdown": "# x", "preguntas": []},
        sesion="s1", df_registros=data_manager.leer_registros(),
    )

    # La hoja existe y las partes OOXML de la tabla dinámica fueron inyectadas.
    with zipfile.ZipFile(config.CONSOLIDADO_PATH) as z:
        partes = z.namelist()
    assert "xl/pivotTables/pivotTable1.xml" in partes
    assert "xl/pivotCache/pivotCacheDefinition1.xml" in partes

    # El archivo sigue siendo un .xlsx válido y legible.
    assert consolidacion.HOJA_PIVOTE_INT in pd.ExcelFile(config.CONSOLIDADO_PATH).sheet_names


# --- Indicador SIRO --------------------------------------------------------
def test_siro_networkdays_y_rangos():
    from datetime import date
    # 2026-02-24 (mar) -> 2026-03-02 (lun) = 5 días hábiles inclusivo.
    assert indicador_siro.networkdays(date(2026, 2, 24), date(2026, 3, 2)) == 5
    assert indicador_siro.networkdays(date(2026, 3, 2), date(2026, 3, 2)) == 1
    assert indicador_siro.networkdays(None, date(2026, 3, 2)) is None
    assert indicador_siro.rango_creacion(5) == "Entre 0 a 5 Dias"
    assert indicador_siro.rango_creacion(11) == "Entre 11 a 15 Dias"
    assert indicador_siro.rango_creacion(40) == "Mayor a 15 Dias"
    assert indicador_siro.rango_asignacion(3) == "Entre 1 a 3 Dias"
    assert indicador_siro.indice(5, indicador_siro.UMBRAL_INGRESO) == "EFECTIVO"
    assert indicador_siro.indice(8, indicador_siro.UMBRAL_INGRESO) == "RETRASO"


def test_siro_calcular_derivados_fila_real():
    import pandas as pd
    # Fila real (JONATHAN POVEDA): creación 2/24, revisión 3/2, asignación 2/26,
    # cierre 3/3, ingreso 3/2 -> días 5/1/3/2, índices EFECTIVO/EFECTIVO.
    df = pd.DataFrame([{
        indicador_siro.C_FECHA_CREACION: "2026-02-24",
        indicador_siro.C_FECHA_REV_RH: "2026-03-02",
        indicador_siro.C_FECHA_ASIG: "2026-02-26",
        indicador_siro.C_FECHA_CIERRE: "2026-03-03",
        indicador_siro.C_FECHA_INGRESO: "2026-03-02",
        indicador_siro.C_DIAS_CREACION: None, indicador_siro.C_DIAS_EFECTIVOS: None,
        indicador_siro.C_DIAS_ASIG: None, indicador_siro.C_DIAS_CIERRE: None,
        indicador_siro.C_INDICE: None, indicador_siro.C_INDICE_TI: None,
        indicador_siro.C_RANGO_CREACION: None, indicador_siro.C_RANGO_ASIG: None,
        indicador_siro.C_MES: None, indicador_siro.C_ANIO: None,
    }])
    out = indicador_siro.calcular_derivados(df).iloc[0]
    assert out[indicador_siro.C_DIAS_CREACION] == 5
    assert out[indicador_siro.C_DIAS_EFECTIVOS] == 1
    assert out[indicador_siro.C_DIAS_ASIG] == 3
    assert out[indicador_siro.C_DIAS_CIERRE] == 2
    assert out[indicador_siro.C_INDICE] == "EFECTIVO"
    assert out[indicador_siro.C_RANGO_CREACION] == "Entre 0 a 5 Dias"
    assert out[indicador_siro.C_MES] == "Marzo"          # mes de la fecha de ingreso
    assert int(out[indicador_siro.C_ANIO]) == 2026


def test_siro_enriquecer_desde_global():
    import pandas as pd
    colab = pd.DataFrame([{
        "codigo_vendedor": "319EN09", "cedula": "1000755275", "nombre": "Juan Sepulveda",
        "cargo": "Asesor", "area": "USEM", "correo": "", "telefono": "",
        "fecha_ingreso": "2026-02-23", "estado": "Activo",
    }])
    df = pd.DataFrame([{
        indicador_siro.C_CEDULA: "1000755275",
        indicador_siro.C_CODIGO: "N/A", indicador_siro.C_AREA: "",
        indicador_siro.C_FECHA_INGRESO: None, indicador_siro.C_NOMBRE: "",
    }])
    out, rep = indicador_siro.enriquecer_con_global(df, colab)
    assert out.iloc[0][indicador_siro.C_CODIGO] == "319EN09"
    assert out.iloc[0][indicador_siro.C_AREA] == "USEM"
    assert rep["codigo"] == 1 and rep["area"] == 1 and rep["nombre"] == 1


def test_siro_novedades_derivados_y_rangos():
    import pandas as pd
    # Fila real: creación 2/20, revisión 3/19, respuesta TI 3/19 -> creación 20,
    # cierre 1; índice RETRASO (20 > 4); rango "Mayor a 15 Dias"; TI "Entre 0 a 2 Dias".
    df = pd.DataFrame([{
        indicador_siro.N_FECHA_CREACION: "2026-02-20",
        indicador_siro.N_FECHA_REVISION: "2026-03-19",
        indicador_siro.N_FECHA_RESP_TI: "2026-03-19",
        indicador_siro.N_DIAS_CREACION: None, indicador_siro.N_DIAS_CIERRE: None,
        indicador_siro.N_RANGO: None, indicador_siro.N_RANGO_TI: None,
        indicador_siro.N_INDICE: None, indicador_siro.N_MES: None,
    }])
    out = indicador_siro.calcular_derivados_novedades(df).iloc[0]
    assert out[indicador_siro.N_DIAS_CREACION] == 20
    assert out[indicador_siro.N_DIAS_CIERRE] == 1
    assert out[indicador_siro.N_INDICE] == "RETRASO"
    assert out[indicador_siro.N_RANGO] == "Mayor a 15 Dias"
    assert out[indicador_siro.N_RANGO_TI] == "Entre 0 a 2 Dias"
    assert out[indicador_siro.N_MES] == "Febrero"
    # Regla del índice de novedad: <= 4 es EFECTIVO.
    assert indicador_siro.rango_cierre_novedad(4) == "Entre 3 a 5 Dias"
    assert indicador_siro.indice(4, indicador_siro.UMBRAL_NOVEDAD) == "EFECTIVO"
    assert indicador_siro.indice(5, indicador_siro.UMBRAL_NOVEDAD) == "RETRASO"


def test_siro_exportar_libro_con_graficos():
    import io
    import zipfile
    import pandas as pd

    inf = pd.DataFrame([{
        indicador_siro.C_ID: 1, indicador_siro.C_ESTADO: "CERRADO",
        indicador_siro.C_AREA: "USEM", indicador_siro.C_RANGO_CREACION: "Entre 0 a 5 Dias",
    }])
    nov = pd.DataFrame([{
        indicador_siro.N_ID: 2, indicador_siro.N_ESTADO: "CERRADO",
        indicador_siro.N_INDICE: "EFECTIVO", indicador_siro.N_RANGO_TI: "Entre 0 a 2 Dias",
    }])
    data = indicador_siro.exportar_libro({"informatica": inf, "novedades": nov})
    assert zipfile.is_zipfile(io.BytesIO(data))
    import openpyxl
    hojas = openpyxl.load_workbook(io.BytesIO(data)).sheetnames
    assert "SIROS INFORMATICA" in hojas
    assert "SIROS NOVEDADES" in hojas
    assert "GRAFICOS INFORMATICA" in hojas
    assert "GRAFICOS NOVEDADES" in hojas


if __name__ == "__main__":
    import subprocess

    raise SystemExit(subprocess.call(["python", "-m", "pytest", "-v", __file__]))
