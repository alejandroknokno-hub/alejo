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

from src import config, data_manager, indicadores

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
    """Genera un .xlsx en memoria para el botón de descarga."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_etiquetado(df).to_excel(writer, sheet_name=config.SHEET_NAME, index=False)
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

if tiempo_real and st_autorefresh is not None:
    st_autorefresh(interval=intervalo * 1000, key="auto_refresh")
elif tiempo_real and st_autorefresh is None:
    st.sidebar.warning("Instala 'streamlit-autorefresh' para el auto-refresco.")

st.sidebar.divider()
st.sidebar.caption(f"📁 Excel: `{config.EXCEL_PATH.name}`")
st.sidebar.caption(f"🕒 Última carga: {datetime.now().strftime('%H:%M:%S')}")


# ---------------------------------------------------------------------------
# Carga de datos
# ---------------------------------------------------------------------------
df = data_manager.leer_registros()

tab_captura, tab_dashboard, tab_registros = st.tabs(
    ["📝 Captura de datos", "📈 Analítica en tiempo real", "🗂️ Registros / Excel"]
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
# TAB 3: Registros / Excel
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
