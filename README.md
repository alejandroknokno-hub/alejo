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
| ⬇️ **Exportación** | Descarga del Excel filtrado en cualquier momento. |

### 🖥️ Versión Oracle — Monitoreo del trabajo

Pensada para dejar **evidencia y análisis** de lo que haces en tu aplicativo de
trabajo (Oracle):

- **Captura de pantalla** manual (botón) o automática (en cada refresco del modo
  tiempo real). Las imágenes se guardan en `data/capturas/`.
- **Bitácora de eventos**: registra entradas, modificaciones, consultas y notas.
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
│   ├── monitor.py          # Versión Oracle: capturas y análisis del flujo
│   └── colaboradores.py    # Versión Global: ficha del colaborador
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
