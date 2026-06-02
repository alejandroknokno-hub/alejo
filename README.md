# 📊 Control Documental con Analítica en Tiempo Real

Aplicación en **Python** para **captura de datos de control documental** durante la
jornada laboral, con **conexión directa a Excel** y un **dashboard de indicadores de
control en tiempo real**.

Pensada para uso contable/comercial: códigos de vendedor, combinaciones contables,
nombre, cédula, valor, estado del documento e información relevante. La fecha (día,
mes, año y hora) se captura automáticamente al guardar cada registro.

---

## ✨ Funcionalidades

| Módulo | Descripción |
|---|---|
| 📝 **Captura de datos** | Formulario validado para registrar documentos durante la jornada. |
| 🔗 **Conexión con Excel** | Lee y escribe directamente en `data/control_documental.xlsx`. |
| 📈 **Analítica en tiempo real** | Dashboard con KPIs y gráficos que se refrescan automáticamente. |
| 🚦 **Indicadores de control** | Tasa de aprobación, pendientes, alertas de cédulas duplicadas, etc. |
| 🖥️ **Monitoreo (versión Oracle)** | Captura de pantalla de tus entradas/modificaciones + análisis del flujo de trabajo. |
| 👥 **Global (Colaboradores)** | Ficha de cada colaborador, cruzada con su actividad documental. |
| 🤖 **IA / Documentos (Claude)** | Conexión directa con Claude: analiza los datos y genera los documentos de control. |
| 📋 **Indicador SIRO** | Completa y recalcula tu indicador SIRO (creación Oracle) cruzando con Global. |
| ⬇️ **Exportación** | Descarga del Excel filtrado en cualquier momento. |

### 📋 Indicador SIRO — completar y recalcular

Sube tu indicador (hoja **SIROS INFORMATICA**) y la app:

- **Enriquece por cédula desde Global**: rellena *código de vendedor, área, fecha de
  ingreso y nombre* donde estén vacíos o en `N/A`.
- **Recalcula los campos derivados** (los que hoy llenas a mano), con fórmulas
  **deducidas y validadas contra tu propio archivo** (coincidencia 100 % en días,
  rangos e índices):
  - *Días* = `NETWORKDAYS` (días hábiles, lun–vie, inclusivo):
    creación→revisión RH, ingreso→revisión RH, creación→asignación correos, revisión RH→cierre TI.
  - *Índice de efectividad* = EFECTIVO si días efectivos al ingreso ≤ 5; *Índice TI* = EFECTIVO si días de asignación ≤ 3.
  - *Rangos* según la hoja MEDIDAS; *Mes/Año* desde la **fecha de ingreso**.
- Descargas el indicador completado en Excel.

Además, el **monitoreo** ahora permite elegir la **plataforma** (Oracle / SIRO / Global),
ya que las peticiones suelen hacerse por la plataforma SIRO y se capturan igual que Oracle.

### 🤖 IA / Documentos — Generación con Claude

La app se conecta directamente con **Claude (API de Anthropic)** para analizar todo lo
capturado en la jornada (registros, colaboradores y la bitácora de monitoreo) y **redactar
el documento de control** automáticamente.

- Punto clave: **cuando un dato u origen no queda claro en la sesión, la IA no lo inventa —
  te pregunta.** Respondes las dudas y pulsas *“Regenerar con mis respuestas”* para obtener
  una versión depurada.
- El documento se descarga en Markdown.
- **Consolidación de jornada**: cada documento generado se guarda automáticamente en
  **un único Excel** (`data/consolidacion_de_jornada.xlsx`) que **va acumulando todo,
  sin separarlo por fecha**, organizado **por hojas**:
  - *Consolidación de jornada* — una fila por cada documento (acumula).
  - *Registros* — detalle de los registros documentales.
  - *Colaboradores* — ficha de los colaboradores.
  - *Bitácora* — eventos de la sesión de monitoreo.
  - *KPIs* — indicadores de control (totales, tasa de aprobación, pendientes, valores) y alertas.
  - *Tablas dinámicas* — resúmenes por estado, por vendedor y por combinación contable.
  - *Dashboard* — gráficos nativos de Excel (torta por estado, barras por vendedor y por combinación).
  - *Pivote interactivo* — una **tabla dinámica nativa de Excel** (campos arrastrables).
    Se genera inyectando las partes OOXML de la PivotTable con `refreshOnLoad`, de modo que
    **Excel la recalcula al abrir** el archivo a partir de la hoja *Registros*. Por defecto
    cuenta registros por estado; en Excel puedes arrastrar cualquier otro campo (vendedor,
    combinación contable, mes, etc.).

  Al regenerar tras responder aclaraciones, se actualiza la última entrada de la sesión
  (no se duplica).
- Configura tu **API key** en la barra lateral (campo *API key de Anthropic*) o define la
  variable de entorno `ANTHROPIC_API_KEY`. La key se usa solo en tu sesión.
- Modelo: `claude-opus-4-8`, con *adaptive thinking*, *prompt caching* del prompt estable y
  salida estructurada (documento + preguntas).

### 🖥️ Versión Oracle — Monitoreo del trabajo

Pensada para dejar **evidencia y análisis** de lo que haces en tu aplicativo de
trabajo (Oracle):

- **🎥 Grabación tipo video (flujo completo paso a paso)**: con un botón inicias
  una grabación que toma **fotogramas** de la pantalla cada pocos segundos *y*
  registra **lo que escribes** (teclado) en la misma línea de tiempo. Al detener,
  ensambla los fotogramas en un **GIF** descargable. Reconstruye qué viste, qué
  datos había y qué anexaste, paso a paso.
  - Es **opt-in**: la inicias y la detienes tú, con un indicador `🔴 GRABANDO`
    visible mientras está activa. Solo registra tu propia sesión en tu equipo.
  - El registro de teclado usa `pynput` (en macOS puede pedir permisos de
    Accesibilidad la primera vez). Si no está disponible, igual graba los fotogramas.
- **Captura de pantalla** manual (botón) o automática (en cada refresco del modo
  tiempo real). Las imágenes se guardan en `data/capturas/`.
- **Detección automática de cambios**: cada captura se compara con la anterior y,
  si el cambio visual supera la *sensibilidad* configurada, se registra solo como
  una **modificación detectada**. No tienes que anotar nada a mano.
- **Bitácora de eventos**: registra entradas, modificaciones, consultas, notas y
  el texto escrito.
- **Sesión de prueba**: la app analiza tu flujo de trabajo (duración, ritmo,
  actividad por minuto y tipos de acción) para entender *cómo trabajas*.

> 📌 La captura de pantalla requiere ejecutar la app en **tu equipo con pantalla**
> (no funciona en un servidor sin display). La bitácora y el análisis funcionan
> en cualquier entorno.

### 👥 Versión Global — Información del colaborador

Equivalente a tu aplicativo "Global": administra la ficha de cada persona
(código de vendedor, cédula, nombre, cargo, área, contacto, estado) en
`data/global_colaboradores.xlsx` y la **cruza por cédula** con los registros de
control documental para ver cuántos documentos y qué valor gestiona cada quien.

---

## 🧱 Requisitos

- **Python 3.11 o superior**
- Visual Studio Code (recomendado) con la extensión *Python*

---

## 🚀 Instalación y ejecución (paso a paso)

Abre la carpeta del proyecto en **VS Code** y, en la terminal integrada:

```bash
# 1) Crear y activar un entorno virtual
python -m venv .venv

# Windows (PowerShell)
.venv\Scripts\Activate.ps1
# macOS / Linux
source .venv/bin/activate

# 2) Instalar dependencias
pip install -r requirements.txt

# 3) Ejecutar la aplicación
streamlit run app.py
```

Se abrirá automáticamente en el navegador (por defecto `http://localhost:8501`).

> 💡 En VS Code también puedes pulsar **F5** y elegir *“Streamlit: app.py”*
> (ya está configurado en `.vscode/launch.json`).

---

## 🗂️ Estructura del proyecto

```
alejo/
├── app.py                  # Interfaz Streamlit (captura + dashboard)
├── requirements.txt        # Dependencias
├── README.md
├── data/
│   └── control_documental.xlsx   # Excel generado con los registros
├── src/
│   ├── config.py           # Esquema de columnas, rutas, estados
│   ├── data_manager.py     # Lectura/escritura y validación contra Excel
│   ├── indicadores.py      # Cálculo de KPIs e indicadores de control
│   ├── monitor.py          # Versión Oracle: capturas, detección y análisis
│   ├── grabador.py         # Grabación tipo video: fotogramas + teclado → GIF
│   ├── colaboradores.py    # Versión Global: ficha del colaborador
│   ├── analista_ia.py      # Conexión con Claude: genera documentos de control
│   ├── consolidacion.py    # Excel único acumulado "Consolidación de jornada"
│   ├── pivote_excel.py     # Inyecta una tabla dinámica (PivotTable) nativa en el .xlsx
│   └── indicador_siro.py   # Completa/recalcula el indicador SIRO (creación Oracle)
└── tests/
    └── test_app.py         # Pruebas básicas
```

---

## 📋 Campos del registro

- **Código de vendedor** *(obligatorio)*
- **Nombre** *(obligatorio)*
- **Cédula** *(obligatorio, solo números 5–15 dígitos)*
- **Combinación contable** *(obligatorio)*
- **Valor**
- **Estado**: Pendiente · En proceso · Aprobado · Rechazado
- **Información relevante**
- **Día / Mes / Año / Fecha y hora** *(automáticos)*
- **Registrado por** *(usuario de la barra lateral)*

---

## 📈 Indicadores de control incluidos

- Registros del día y acumulados
- Vendedores activos
- **Tasa de aprobación** (% de documentos aprobados)
- Documentos pendientes / rechazados
- Valor acumulado y valor del día
- Distribución por estado y por combinación contable
- Ritmo de la jornada (registros por hora)
- Alertas: cédulas duplicadas, registros sin información, rechazos

---

## 🔧 Personalización

Para agregar o cambiar campos, edita el diccionario `COLUMNS` en
[`src/config.py`](src/config.py). El resto de la app se adapta automáticamente.
