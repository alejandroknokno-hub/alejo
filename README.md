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
| ⬇️ **Exportación** | Descarga del Excel filtrado en cualquier momento. |

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
│   └── indicadores.py      # Cálculo de KPIs e indicadores de control
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
