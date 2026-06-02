"""Inyección de una TABLA DINÁMICA (PivotTable) nativa e interactiva en un .xlsx.

openpyxl no crea PivotTables, así que generamos las partes OOXML que Excel
necesita y las inyectamos dentro del archivo ya escrito:

- xl/pivotCache/pivotCacheDefinition1.xml  (+ sus rels y records)
- xl/pivotTables/pivotTable1.xml            (+ sus rels)
- Overrides en [Content_Types].xml
- <pivotCaches> en xl/workbook.xml (+ relación en workbook.xml.rels)
- relación de la hoja destino hacia la pivotTable

La caché se marca con `refreshOnLoad="1"` y `recordCount="0"`, de modo que Excel
**recalcula la tabla dinámica al abrir** el archivo a partir de la hoja de datos.
Así obtienes una tabla dinámica 100% interactiva (campos arrastrables) sin que
tengamos que precargar todos los registros en la caché.

La tabla dinámica por defecto cuenta registros por estado (fila = Estado,
valores = conteo de Cédula). En Excel puedes arrastrar cualquier otro campo.
"""
from __future__ import annotations

import io
import re
import zipfile

from openpyxl.utils import get_column_letter

from . import config

NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def inyectar_pivote(ruta, df_registros, hoja_datos: str, hoja_pivote: str) -> bool:
    """Inyecta una tabla dinámica en `ruta`. Devuelve True si lo logró.

    Es atómica: solo reescribe el archivo si todo salió bien; ante cualquier
    error deja el archivo original intacto y devuelve False.
    """
    try:
        with open(ruta, "rb") as fh:
            original = fh.read()

        partes: dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(original)) as zin:
            nombres = zin.namelist()
            for n in nombres:
                partes[n] = zin.read(n)

        # --- Localizar la hoja destino (su archivo sheetN.xml) ---
        workbook_xml = partes["xl/workbook.xml"].decode("utf-8")
        rels_xml = partes["xl/_rels/workbook.xml.rels"].decode("utf-8")

        m = re.search(r'<sheet[^>]*\bname="%s"[^>]*/>' % re.escape(hoja_pivote), workbook_xml)
        if not m:
            return False
        rid = re.search(r'r:id="([^"]+)"', m.group(0))
        if not rid:
            return False
        rid_hoja = rid.group(1)
        # El orden de atributos varía (Type/Target/Id), así que ubicamos primero
        # la etiqueta <Relationship> con ese Id y luego extraemos su Target.
        rel_tag = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*/>' % re.escape(rid_hoja), rels_xml)
        if not rel_tag:
            return False
        tgt = re.search(r'Target="([^"]+)"', rel_tag.group(0))
        if not tgt:
            return False
        destino = tgt.group(1).lstrip("/")          # /xl/worksheets/sheet8.xml -> xl/worksheets/sheet8.xml
        sheet_file = destino.split("/")[-1]
        sheet_rels = f"xl/worksheets/_rels/{sheet_file}.rels"

        # --- Dimensiones de la fuente de datos ---
        ncols = len(config.COLUMN_ORDER)
        ndatos = 0 if df_registros is None else len(df_registros)
        ult_col = get_column_letter(ncols)
        ult_fila = ndatos + 1                         # encabezado + filas
        ref = f"A1:{ult_col}{ult_fila}"

        idx_estado = config.COLUMN_ORDER.index("estado")
        idx_cedula = config.COLUMN_ORDER.index("cedula")

        # --- Construir las partes de la pivot ---
        partes["xl/pivotCache/pivotCacheDefinition1.xml"] = _cache_definition(ref, hoja_datos).encode("utf-8")
        partes["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = _rels_simple(
            "rId1",
            f"{NS_REL}/pivotCacheRecords",
            "pivotCacheRecords1.xml",
        ).encode("utf-8")
        partes["xl/pivotCache/pivotCacheRecords1.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<pivotCacheRecords xmlns="{NS_MAIN}" xmlns:r="{NS_REL}" count="0"/>'
        ).encode("utf-8")
        partes["xl/pivotTables/pivotTable1.xml"] = _pivot_table(idx_estado, idx_cedula).encode("utf-8")
        partes["xl/pivotTables/_rels/pivotTable1.xml.rels"] = _rels_simple(
            "rId1",
            f"{NS_REL}/pivotCacheDefinition",
            "../pivotCache/pivotCacheDefinition1.xml",
        ).encode("utf-8")

        # --- Content types ---
        ct = partes["[Content_Types].xml"].decode("utf-8")
        overrides = (
            '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
            '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
            '<Override PartName="/xl/pivotTables/pivotTable1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>'
        )
        partes["[Content_Types].xml"] = ct.replace("</Types>", overrides + "</Types>").encode("utf-8")

        # --- workbook.xml: <pivotCaches> + relación nueva ---
        rid_cache = _nuevo_rid(rels_xml)
        pivot_caches = f'<pivotCaches><pivotCache cacheId="1" r:id="{rid_cache}"/></pivotCaches>'
        partes["xl/workbook.xml"] = workbook_xml.replace("</workbook>", pivot_caches + "</workbook>").encode("utf-8")
        rel_cache = (
            f'<Relationship Id="{rid_cache}" '
            f'Type="{NS_REL}/pivotCacheDefinition" '
            f'Target="pivotCache/pivotCacheDefinition1.xml"/>'
        )
        partes["xl/_rels/workbook.xml.rels"] = rels_xml.replace(
            "</Relationships>", rel_cache + "</Relationships>"
        ).encode("utf-8")

        # --- Relación de la hoja destino -> pivotTable ---
        rel_pivot = (
            f'<Relationship Id="rId1" Type="{NS_REL}/pivotTable" '
            f'Target="../pivotTables/pivotTable1.xml"/>'
        )
        if sheet_rels in partes:
            actual = partes[sheet_rels].decode("utf-8")
            rid_nuevo = _nuevo_rid(actual)
            rel_pivot = rel_pivot.replace('Id="rId1"', f'Id="{rid_nuevo}"')
            partes[sheet_rels] = actual.replace(
                "</Relationships>", rel_pivot + "</Relationships>"
            ).encode("utf-8")
        else:
            partes[sheet_rels] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{NS_PKG_REL}">{rel_pivot}</Relationships>'
            ).encode("utf-8")

        # --- Reescritura atómica del .xlsx ---
        salida = io.BytesIO()
        with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
            for nombre, datos in partes.items():
                zout.writestr(nombre, datos)
        with open(ruta, "wb") as fh:
            fh.write(salida.getvalue())
        return True
    except Exception:
        return False


# --- Plantillas XML --------------------------------------------------------
def _nuevo_rid(rels_xml: str) -> str:
    ids = [int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def _rels_simple(rid: str, tipo: str, destino: str) -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_PKG_REL}">'
        f'<Relationship Id="{rid}" Type="{tipo}" Target="{destino}"/>'
        f'</Relationships>'
    )


def _cache_definition(ref: str, hoja_datos: str) -> str:
    campos = "".join(
        f'<cacheField name="{config.COLUMNS[c]}" numFmtId="0"><sharedItems/></cacheField>'
        for c in config.COLUMN_ORDER
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotCacheDefinition xmlns="{NS_MAIN}" xmlns:r="{NS_REL}" '
        f'r:id="rId1" refreshOnLoad="1" refreshedBy="ControlDocumental" refreshedDate="0" '
        f'createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">'
        f'<cacheSource type="worksheet"><worksheetSource ref="{ref}" sheet="{hoja_datos}"/></cacheSource>'
        f'<cacheFields count="{len(config.COLUMN_ORDER)}">{campos}</cacheFields>'
        f'</pivotCacheDefinition>'
    )


def _pivot_table(idx_estado: int, idx_cedula: int) -> str:
    n = len(config.COLUMN_ORDER)
    pivot_fields = []
    for i in range(n):
        if i == idx_estado:
            pivot_fields.append(
                '<pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
            )
        elif i == idx_cedula:
            pivot_fields.append('<pivotField dataField="1" showAll="0"/>')
        else:
            pivot_fields.append('<pivotField showAll="0"/>')
    campos = "".join(pivot_fields)
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<pivotTableDefinition xmlns="{NS_MAIN}" name="TablaDinamica1" cacheId="1" '
        f'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        f'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Valores" '
        f'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" '
        f'createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
        f'<location ref="A3:B20" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>'
        f'<pivotFields count="{n}">{campos}</pivotFields>'
        f'<rowFields count="1"><field x="{idx_estado}"/></rowFields>'
        f'<rowItems count="1"><i><x/></i></rowItems>'
        f'<colItems count="1"><i/></colItems>'
        f'<dataFields count="1">'
        f'<dataField name="Conteo de Cédula" fld="{idx_cedula}" subtotal="count" baseField="0" baseItem="0"/>'
        f'</dataFields>'
        f'<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" '
        f'showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        f'</pivotTableDefinition>'
    )
