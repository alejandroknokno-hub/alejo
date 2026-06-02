"""Aplicación de Control Documental y Analítica en Tiempo Real.

Ejecuta:  streamlit run app.py

Funcionalidades:
  1. Captura de datos para control documental durante la jornada laboral.
  2. Conexión directa con Excel (lectura/escritura en data/control_documental.xlsx).
  3. Dashboard de analítica en tiempo real con refresco automático.
  4. Indicadores de control y alertas de calidad de datos.
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

from src import colaboradores, config, data_manager, indicadores, monitor

# El autorefresh es opcional; si no está instalado, la app sigue funcionando.
try:
    from streamlit_autorefresh import st_autorefresh
except ImportError:  # pragma: no cover
    st_autorefresh = None


st.set_page_config(
    page_title="Control Documental · Analítica en Tiempo Real",
    page_icon="📊",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Utilidades de presentación
# ---------------------------------------------------------------------------
def df_etiquetado(df: pd.DataFrame) -> pd.DataFrame:
    """Renombra las columnas internas a sus etiquetas legibles."""
    return df.rename(columns=config.COLUMNS)


def excel_en_memoria(df: pd.DataFrame) -> bytes:
    """Genera un .xlsx en memoria (registros documentales) para descarga."""
    return excel_en_memoria_generico(df, config.COLUMNS, config.SHEET_NAME)


def excel_en_memoria_generico(df: pd.DataFrame, etiquetas: dict, hoja: str) -> bytes:
    """Genera un .xlsx en memoria a partir de cualquier DataFrame y sus etiquetas."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.rename(columns=etiquetas).to_excel(writer, sheet_name=hoja, index=False)
    return buffer.getvalue()


def formato_moneda(valor: float) -> str:
    return f"$ {valor:,.0f}".replace(",", ".")


# ---------------------------------------------------------------------------
# Barra lateral
# ---------------------------------------------------------------------------
st.sidebar.title("📊 Control Documental")
st.sidebar.caption("Analítica en tiempo real conectada a Excel")

usuario = st.sidebar.text_input("👤 Usuario que registra", value="", placeholder="Tu nombre")

st.sidebar.divider()
tiempo_real = st.sidebar.toggle("🔄 Modo tiempo real (auto-refresco)", value=False)
intervalo = st.sidebar.slider("Intervalo de refresco (seg)", 5, 60, 15, disabled=not tiempo_real)

# --- Monitoreo (versión Oracle) en la barra lateral ---
st.sidebar.divider()
st.sidebar.markdown("### 🖥️ Monitoreo (Oracle)")
sesion_actual = st.sidebar.text_input("Sesión", value="sesion_prueba")
monitoreo_activo = st.sidebar.toggle(
    "📸 Capturar en cada refresco", value=False, disabled=not tiempo_real,
    help="Requiere el modo tiempo real activo. Toma una captura de pantalla en cada ciclo.",
)
umbral_cambio = st.sidebar.slider(
    "🎯 Sensibilidad de cambio (%)", 0.5, 20.0, 2.0, 0.5, disabled=not tiempo_real,
    help="Cambio mínimo de pantalla para registrar una modificación automática. Más bajo = más sensible.",
)

if tiempo_real and st_autorefresh is not None:
    ciclo = st_autorefresh(interval=intervalo * 1000, key="auto_refresh")
    # Captura automática + detección de cambios ligada al ciclo de refresco.
    if monitoreo_activo:
        res_cap = monitor.capturar_y_detectar(
            sesion=sesion_actual, umbral_pct=umbral_cambio,
            descripcion="Captura automática (tiempo real)",
        )
        if res_cap["ok"]:
            if res_cap["cambio_detectado"]:
                st.sidebar.warning(f"✏️ Cambio detectado: {res_cap['diferencia_pct']:.1f}% (captura #{ciclo})")
            elif res_cap.get("primera"):
                st.sidebar.caption(f"✅ Captura inicial #{ciclo} (base de comparación)")
            else:
                st.sidebar.caption(f"✅ Sin cambios · captura #{ciclo} ({res_cap['diferencia_pct']:.1f}%)")
        else:
            st.sidebar.caption("⚠️ Captura no disponible aquí")
elif tiempo_real and st_autorefresh is None:
    st.sidebar.warning("Instala 'streamlit-autorefresh' para el auto-refresco.")

st.sidebar.divider()
st.sidebar.caption(f"📁 Excel: `{config.EXCEL_PATH.name}`")
st.sidebar.caption(f"🕒 Última carga: {datetime.now().strftime('%H:%M:%S')}")


# ---------------------------------------------------------------------------
# Carga de datos
# ---------------------------------------------------------------------------
df = data_manager.leer_registros()

tab_captura, tab_dashboard, tab_monitor, tab_global, tab_registros = st.tabs(
    [
        "📝 Captura de datos",
        "📈 Analítica en tiempo real",
        "🖥️ Monitoreo (Oracle)",
        "👥 Global (Colaboradores)",
        "🗂️ Registros / Excel",
    ]
)


# ---------------------------------------------------------------------------
# TAB 1: Captura de datos
# ---------------------------------------------------------------------------
with tab_captura:
    st.subheader("Captura para control documental")
    st.caption("Los campos día, mes, año y hora se registran automáticamente al guardar.")

    with st.form("form_captura", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            codigo_vendedor = st.text_input("Código de vendedor *", placeholder="Ej: V-00123")
            nombre = st.text_input("Nombre *", placeholder="Nombre completo")
        with col2:
            cedula = st.text_input("Cédula *", placeholder="Solo números")
            combinacion_contable = st.text_input(
                "Combinación contable *", placeholder="Ej: 5105-01-001"
            )
        with col3:
            valor = st.number_input("Valor", min_value=0.0, step=1000.0, format="%.2f")
            estado = st.selectbox("Estado *", config.ESTADOS)

        informacion_relevante = st.text_area(
            "Información relevante", placeholder="Observaciones, soporte, novedad, etc."
        )

        enviado = st.form_submit_button("💾 Guardar registro", use_container_width=True)

    if enviado:
        resultado = data_manager.agregar_registro(
            {
                "codigo_vendedor": codigo_vendedor,
                "nombre": nombre,
                "cedula": cedula,
                "combinacion_contable": combinacion_contable,
                "valor": valor,
                "estado": estado,
                "informacion_relevante": informacion_relevante,
            },
            registrado_por=usuario,
        )
        if resultado["ok"]:
            st.success("✅ Registro guardado y sincronizado con Excel.")
            st.rerun()
        else:
            for err in resultado["errores"]:
                st.error(f"❌ {err}")


# ---------------------------------------------------------------------------
# TAB 2: Analítica en tiempo real
# ---------------------------------------------------------------------------
with tab_dashboard:
    st.subheader("Indicadores de control")

    k = indicadores.kpis_generales(df)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros hoy", k["registros_hoy"], delta=f"{k['total_registros']} en total")
    c2.metric("Vendedores activos", k["vendedores_activos"])
    c3.metric("Tasa de aprobación", f"{k['tasa_aprobacion']:.1f}%")
    c4.metric("Pendientes", k["pendientes"])

    c5, c6 = st.columns(2)
    c5.metric("Valor acumulado", formato_moneda(k["valor_total"]))
    c6.metric("Valor de hoy", formato_moneda(k["valor_hoy"]))

    st.divider()

    # Alertas de control
    st.markdown("#### 🚦 Alertas de control")
    for alerta in indicadores.alertas_control(df):
        getattr(st, alerta["tipo"] if alerta["tipo"] in ("error", "warning", "info") else "info")(
            alerta["mensaje"]
        )

    st.divider()

    if df.empty:
        st.info("Captura tu primer registro en la pestaña **Captura de datos** para ver la analítica.")
    else:
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("#### Registros por estado")
            est = indicadores.por_estado(df)
            fig = px.pie(est, names="estado", values="cantidad", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("#### Ritmo de la jornada (registros por hora)")
            tend = indicadores.tendencia_por_hora(df)
            fig = px.bar(tend, x="hora", y="registros")
            fig.update_layout(xaxis_title="Hora del día", yaxis_title="Registros")
            st.plotly_chart(fig, use_container_width=True)

        g3, g4 = st.columns(2)
        with g3:
            st.markdown("#### Top vendedores")
            vend = indicadores.por_vendedor(df).head(10)
            fig = px.bar(vend, x="codigo_vendedor", y="registros", hover_data=["nombre"])
            fig.update_layout(xaxis_title="Código de vendedor", yaxis_title="Registros")
            st.plotly_chart(fig, use_container_width=True)
        with g4:
            st.markdown("#### Por combinación contable")
            comb = indicadores.por_combinacion(df).head(10)
            fig = px.bar(comb, x="combinacion_contable", y="cantidad")
            fig.update_layout(xaxis_title="Combinación contable", yaxis_title="Registros")
            st.plotly_chart(fig, use_container_width=True)


# ---------------------------------------------------------------------------
# TAB 3: Monitoreo (versión Oracle) — capturas y análisis del flujo de trabajo
# ---------------------------------------------------------------------------
with tab_monitor:
    st.subheader("🖥️ Monitoreo del trabajo en Oracle")
    st.caption(
        "Captura tu pantalla y **detecta automáticamente** los cambios respecto a la "
        "captura anterior. Cuando el cambio supera la sensibilidad configurada, se "
        "registra solo como una *modificación*. Ajusta la sensibilidad en la barra lateral."
    )

    if not monitor.captura_disponible():
        st.warning(
            "⚠️ La captura de pantalla no está disponible en este entorno. "
            "Ejecuta la app en tu equipo de trabajo (con pantalla) y con la librería "
            "`mss` instalada. La bitácora de eventos sí funciona en todos lados."
        )

    # --- Acciones rápidas ---
    a1, a2, a3 = st.columns(3)
    if a1.button("📸 Capturar y detectar", use_container_width=True):
        res = monitor.capturar_y_detectar(
            sesion=sesion_actual, umbral_pct=umbral_cambio, descripcion="Captura manual",
        )
        if res["ok"]:
            if res["cambio_detectado"]:
                st.success(f"✏️ Cambio detectado automáticamente: {res['diferencia_pct']:.1f}% de la pantalla.")
            elif res.get("primera"):
                st.info("Captura inicial guardada (servirá de base para comparar).")
            else:
                st.info(f"Captura guardada. Sin cambios relevantes ({res['diferencia_pct']:.1f}%).")
        else:
            st.error(res["error"])

    with a2.popover("➕ Registrar evento", use_container_width=True):
        with st.form("form_evento", clear_on_submit=True):
            tipo_evt = st.selectbox("Tipo", config.TIPOS_EVENTO)
            desc_evt = st.text_input("Descripción", placeholder="Ej: Modifiqué la combinación contable X")
            if st.form_submit_button("Guardar evento"):
                monitor.registrar_evento(tipo_evt, desc_evt, sesion=sesion_actual)
                st.success("Evento registrado.")
                st.rerun()

    a3.metric("Sesión activa", sesion_actual)

    st.divider()

    # --- Análisis del flujo de trabajo (sesión de prueba) ---
    st.markdown("#### 🔎 Análisis del flujo de trabajo")
    flujo = monitor.analizar_flujo(sesion_actual)
    st.info(flujo["resumen"])

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Eventos", flujo["eventos"])
    m2.metric("Entradas", flujo["entradas"])
    m3.metric("Modificaciones", flujo["modificaciones"])
    m4.metric("Capturas", flujo["capturas"])

    m5, m6 = st.columns(2)
    m5.metric("Duración (min)", f"{flujo['duracion_min']:.1f}")
    m6.metric("Ritmo (seg/acción)", f"{flujo['intervalo_prom_seg']:.0f}")

    if not flujo["linea_tiempo"].empty:
        c7, c8 = st.columns(2)
        with c7:
            st.markdown("**Actividad por minuto**")
            fig = px.bar(flujo["linea_tiempo"], x="minuto", y="eventos")
            st.plotly_chart(fig, use_container_width=True)
        with c8:
            st.markdown("**Eventos por tipo**")
            fig = px.pie(flujo["por_tipo"], names="tipo", values="cantidad", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # --- Galería de capturas recientes ---
    st.markdown("#### 🖼️ Capturas recientes")
    capturas = monitor.capturas_recientes(sesion_actual, limite=9)
    if not capturas:
        st.caption("Aún no hay capturas en esta sesión.")
    else:
        cols = st.columns(3)
        for i, cap in enumerate(capturas):
            with cols[i % 3]:
                st.image(cap["ruta"], use_container_width=True,
                         caption=f"{cap['fecha_hora']:%H:%M:%S} · {cap['descripcion']}")

    # --- Bitácora completa ---
    with st.expander("📜 Ver bitácora completa de la sesión"):
        log = monitor.leer_log(sesion_actual)
        st.dataframe(log, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# TAB 4: Global — información del colaborador
# ---------------------------------------------------------------------------
with tab_global:
    st.subheader("👥 Global · Información del colaborador")
    st.caption("Administra la ficha de cada colaborador y crúzala con su actividad documental.")

    with st.expander("➕ Agregar / actualizar colaborador", expanded=False):
        with st.form("form_colab", clear_on_submit=True):
            g1, g2, g3 = st.columns(3)
            with g1:
                g_codigo = st.text_input("Código de vendedor *", placeholder="V-00123")
                g_cedula = st.text_input("Cédula *", placeholder="Solo números")
                g_nombre = st.text_input("Nombre *")
            with g2:
                g_cargo = st.text_input("Cargo")
                g_area = st.text_input("Área")
                g_correo = st.text_input("Correo")
            with g3:
                g_tel = st.text_input("Teléfono")
                g_ingreso = st.date_input("Fecha de ingreso")
                g_estado = st.selectbox("Estado *", config.ESTADOS_COLAB)

            if st.form_submit_button("💾 Guardar colaborador", use_container_width=True):
                res = colaboradores.agregar_colaborador({
                    "codigo_vendedor": g_codigo,
                    "cedula": g_cedula,
                    "nombre": g_nombre,
                    "cargo": g_cargo,
                    "area": g_area,
                    "correo": g_correo,
                    "telefono": g_tel,
                    "fecha_ingreso": g_ingreso.isoformat(),
                    "estado": g_estado,
                })
                if res["ok"]:
                    st.success(f"✅ Colaborador {res['accion']} correctamente.")
                    st.rerun()
                else:
                    for err in res["errores"]:
                        st.error(f"❌ {err}")

    df_colab = colaboradores.leer_colaboradores()

    if df_colab.empty:
        st.info("Aún no hay colaboradores registrados.")
    else:
        cruce = colaboradores.actividad_por_colaborador(df_colab, df)

        k1, k2, k3 = st.columns(3)
        k1.metric("Colaboradores", len(df_colab))
        k2.metric("Activos", int((df_colab["estado"] == "Activo").sum()))
        k3.metric("Con actividad documental", int((cruce["registros"] > 0).sum()))

        renombrar = {**config.COLAB_COLUMNS, "registros": "Registros", "valor_gestionado": "Valor gestionado"}
        st.dataframe(cruce.rename(columns=renombrar), use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Descargar colaboradores (Excel)",
            data=excel_en_memoria_generico(df_colab, config.COLAB_COLUMNS, config.GLOBAL_SHEET),
            file_name=f"global_colaboradores_{datetime.now():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ---------------------------------------------------------------------------
# TAB 5: Registros / Excel
# ---------------------------------------------------------------------------
with tab_registros:
    st.subheader("Registros capturados")

    if df.empty:
        st.info("No hay registros todavía.")
    else:
        # Filtros rápidos
        f1, f2, f3 = st.columns(3)
        filtro_estado = f1.multiselect("Filtrar por estado", config.ESTADOS)
        filtro_vendedor = f2.text_input("Filtrar por código de vendedor")
        filtro_cedula = f3.text_input("Filtrar por cédula")

        vista = df.copy()
        if filtro_estado:
            vista = vista[vista["estado"].isin(filtro_estado)]
        if filtro_vendedor:
            vista = vista[vista["codigo_vendedor"].astype(str).str.contains(filtro_vendedor, case=False, na=False)]
        if filtro_cedula:
            vista = vista[vista["cedula"].astype(str).str.contains(filtro_cedula, na=False)]

        st.caption(f"{len(vista)} de {len(df)} registros")
        st.dataframe(df_etiquetado(vista), use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Descargar Excel filtrado",
            data=excel_en_memoria(vista),
            file_name=f"control_documental_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
