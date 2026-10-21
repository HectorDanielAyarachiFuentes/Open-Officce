# -*- coding: utf-8 -*-
"""
calc_toolkit.py — LibreOffice Calc Automation & Audit Framework
==============================================================
Módulo oficial de automatización nativa de LibreOffice Calc (Python + UNO).
Diseñado para resolver prácticas, exámenes y automatizaciones en segundos,
garantizando 100% de cumplimiento con las reglas docentes de la cátedra:
- Motor nativo UNO con manejo robusto de named pipes y cierres limpios.
- Ordenamiento multicriterio a prueba de fallos silenciosos (TableSortField).
- Filtros estándar con DatabaseRanges y preservación automática de encabezados.
- Tablas dinámicas (DataPilot) y Pivot Charts dinámicos (loext:data-pilot-source).
- Auditoría automatizada de entrega (detección de Err:510, #¿NOMBRE?, filas colapsadas).
"""

import os
import sys
import time
import subprocess
import zipfile
import re
from typing import List, Tuple, Dict, Any, Optional

SOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
PYTHON_PATH = r"C:\Program Files\LibreOffice\program\python.exe"
DEFAULT_PIPE_NAME = "calc_toolkit_pipe"

try:
    import uno
    from com.sun.star.beans import PropertyValue
    from com.sun.star.table import TableSortField, CellRangeAddress, CellAddress
    from com.sun.star.sheet import TableFilterField, FilterOperator, GeneralFunction
    from com.sun.star.awt import Rectangle
except ImportError:
    pass  # Permite inspección de código fuera de Python UNO si fuera necesario


class CalcSession:
    """Administrador de sesión UNO conectado a LibreOffice headless."""

    def __init__(self, pipe_name: str = DEFAULT_PIPE_NAME):
        self.pipe_name = pipe_name
        self.context = None
        self.desktop = None
        self.process = None

    def start_or_connect(self) -> "CalcSession":
        """Inicia o reutiliza un proceso LibreOffice headless por named pipe."""
        import uno
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        url = f"uno:pipe,name={self.pipe_name};urp;StarOffice.ComponentContext"

        for attempt in range(3):
            try:
                self.context = resolver.resolve(url)
                self.desktop = self.context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.context
                )
                return self
            except Exception:
                if attempt == 0:
                    cmd = [
                        SOFFICE_PATH,
                        "--headless",
                        "--invisible",
                        "--nocrashreport",
                        "--nodefault",
                        "--nofirststartwizard",
                        "--nologo",
                        "--norestore",
                        f'--accept=pipe,name={self.pipe_name};urp;StarOffice.ComponentContext'
                    ]
                    self.process = subprocess.Popen(cmd)
                time.sleep(1.2)

        raise RuntimeError("No se pudo establecer conexión con LibreOffice mediante UNO.")

    def open_doc(self, filepath: str):
        """Abre un documento de cálculo (.ods)."""
        abs_path = os.path.abspath(filepath)
        file_url = uno.systemPathToFileUrl(abs_path)
        props = (
            PropertyValue(Name="Hidden", Value=True),
        )
        doc = self.desktop.loadComponentFromURL(file_url, "_blank", 0, props)
        if not doc:
            raise RuntimeError(f"No se pudo cargar el archivo: {filepath}")
        return doc

    def save_and_close(self, doc, filepath: Optional[str] = None):
        """Calcula, guarda y cierra el documento liberando bloqueos."""
        doc.calculateAll()
        if filepath:
            abs_path = os.path.abspath(filepath)
            dest_url = uno.systemPathToFileUrl(abs_path)
            doc.storeToURL(dest_url, ())
        else:
            doc.store()
        doc.close(True)


# ==============================================================================
# 1. GESTIÓN DE HOJAS Y ENCABEZADOS
# ==============================================================================

def get_or_copy_sheet(doc, target_name: str, source_name: Optional[str] = None):
    """Obtiene una hoja existente o la crea copiando de una fuente (conforme a cátedra)."""
    sheets = doc.getSheets()
    if sheets.hasByName(target_name):
        sheets.removeByName(target_name)
    
    if source_name and sheets.hasByName(source_name):
        sheets.copyByName(source_name, target_name, sheets.getCount())
    else:
        sheets.insertNewByName(target_name, sheets.getCount())
    return sheets.getByName(target_name)


def preserve_top_headers(sheet, header_row_index: int = 3):
    """
    Restaura la visibilidad de todas las filas previas al inicio de los datos (0 a header_row_index).
    Evita que Calc colapse u oculte los parámetros fijos (ej. $C$1, $C$2) y títulos de columna.
    """
    rows = sheet.getRows()
    for r in range(header_row_index + 1):
        rows.getByIndex(r).setPropertyValue("IsVisible", True)


# ==============================================================================
# 2. ORDENAMIENTO NATIVO MULTICRITERIO (TableSortField)
# ==============================================================================

def sort_range_native(
    sheet,
    start_col: int,
    start_row: int,
    end_col: int,
    end_row: int,
    keys: List[Tuple[int, bool]],
    contains_header: bool = True
):
    """
    Ordena un rango usando estrictamente com.sun.star.table.TableSortField y uno.Any.
    keys: Lista de tuplas (col_idx, is_ascending).
          Ej: [(1, True), (2, True)] -> Clave 1: Col B asc, Clave 2: Col C asc.
    """
    import uno
    rng = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    sort_desc = rng.createSortDescriptor()

    sort_fields = []
    for col_idx, is_asc in keys:
        sf = uno.createUnoStruct("com.sun.star.table.TableSortField")
        sf.Field = col_idx
        sf.IsAscending = is_asc
        sf.IsCaseSensitive = False
        sort_fields.append(sf)

    for p in sort_desc:
        if p.Name == "SortFields":
            p.Value = uno.Any("[]com.sun.star.table.TableSortField", tuple(sort_fields))
        elif p.Name == "ContainsHeader":
            p.Value = contains_header

    rng.sort(tuple(sort_desc))


# ==============================================================================
# 3. FILTROS ESTÁNDAR Y PERSISTENCIA DE DATABASE RANGES
# ==============================================================================

def apply_filter_standard(
    doc,
    sheet,
    db_name: str,
    range_coords: Tuple[int, int, int, int],
    conditions: List[Dict[str, Any]],
    preserve_header_rows: int = 4
):
    """
    Aplica un filtro estándar persistente mediante DatabaseRange y restaura encabezados.
    range_coords: (start_col, start_row, end_col, end_row).
    conditions: lista de dicts con:
        - field: int (columna 0-based)
        - op: FilterOperator (ej. EQUAL, GREATER_EQUAL, LESS_EQUAL)
        - is_numeric: bool
        - num_value: float (si es numérico)
        - str_value: str (si es texto)
        - connection: str ("AND" / "OR")
    """
    import uno
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.sheet import TableFilterField, FilterConnection

    sc, sr, ec, er = range_coords
    db_ranges = doc.DatabaseRanges
    
    # Crear o recuperar DatabaseRange
    if not db_ranges.hasByName(db_name):
        addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
        addr.Sheet = sheet.RangeAddress.Sheet
        addr.StartColumn = sc
        addr.StartRow = sr
        addr.EndColumn = ec
        addr.EndRow = er
        db_ranges.addNewByName(db_name, addr)
    
    dbr = db_ranges.getByName(db_name)
    filter_desc = dbr.getFilterDescriptor()
    filter_desc.setPropertyValue("ContainsHeader", True)
    
    filter_fields = []
    for cond in conditions:
        ff = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
        ff.Field = cond["field"]
        ff.Operator = cond["op"]
        ff.IsNumeric = cond.get("is_numeric", False)
        if ff.IsNumeric:
            ff.NumericValue = float(cond["num_value"])
        else:
            ff.StringValue = str(cond["str_value"])
        
        conn_str = cond.get("connection", "AND").upper()
        ff.Connection = FilterConnection.AND if conn_str == "AND" else FilterConnection.OR
        filter_fields.append(ff)

    filter_desc.setFilterFields(tuple(filter_fields))
    sheet.filter(filter_desc)
    dbr.refresh()

    # Preservar filas de título y parámetros superiores
    if preserve_header_rows > 0:
        preserve_top_headers(sheet, preserve_header_rows - 1)


# ==============================================================================
# 4. TABLAS DINÁMICAS (DataPilot) Y PIVOT CHARTS
# ==============================================================================

def create_datapilot(
    sheet,
    name: str,
    source_coords: Tuple[int, int, int, int],
    dest_cell: Tuple[int, int],
    row_fields: List[int],
    data_fields: List[Tuple[int, Any]]
):
    """
    Crea una Tabla Dinámica nativa en Calc.
    source_coords: (start_col, start_row, end_col, end_row)
    dest_cell: (dest_col, dest_row)
    data_fields: lista de (col_idx, GeneralFunction.COUNT | SUM | AVERAGE)
    """
    import uno
    from com.sun.star.table import CellRangeAddress, CellAddress
    from com.sun.star.sheet import DataPilotFieldOrientation

    dp_tables = sheet.getDataPilotTables()
    if dp_tables.hasByName(name):
        dp_tables.removeByName(name)

    dp_desc = dp_tables.createDataPilotDescriptor()
    sc, sr, ec, er = source_coords
    src_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    src_addr.Sheet = sheet.RangeAddress.Sheet
    src_addr.StartColumn = sc
    src_addr.StartRow = sr
    src_addr.EndColumn = ec
    src_addr.EndRow = er
    dp_desc.setSourceRange(src_addr)

    # Configurar filas
    for rf_idx in row_fields:
        f = dp_desc.getDataPilotFields().getByIndex(rf_idx)
        f.setPropertyValue("Orientation", DataPilotFieldOrientation.ROW)

    # Configurar campos de datos
    for df_idx, fn in data_fields:
        f = dp_desc.getDataPilotFields().getByIndex(df_idx)
        f.setPropertyValue("Orientation", DataPilotFieldOrientation.DATA)
        f.setPropertyValue("Function", fn)

    dest_addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
    dest_addr.Sheet = sheet.RangeAddress.Sheet
    dest_addr.Column = dest_cell[0]
    dest_addr.Row = dest_cell[1]

    dp_tables.insertNewByName(name, dest_addr, dp_desc)
    return dp_tables.getByName(name)


def create_pivot_chart_dynamic(
    doc,
    sheet,
    chart_name: str,
    rect_coords: Tuple[int, int, int, int],
    live_range_coords: Tuple[int, int, int, int],
    datapilot_name: str,
    title: str = "",
    x_axis_title: str = "",
    y_axis_title: str = ""
):
    """
    Crea un gráfico dinámico vinculado directamente al rango vivo del DataPilot
    (sin usar .setData()) con estilo oficial de cátedra y coordenadas visibles.
    rect_coords: (x, y, width, height) en 1/100 mm (ej. 38000, 5000, 16000, 10000).
    live_range_coords: (sc, sr, ec, er) excluyendo estrictamente 'Total general'.
    """
    import uno
    from com.sun.star.awt import Rectangle
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.drawing import FillStyle

    charts = sheet.getCharts()
    if charts.hasByName(chart_name):
        charts.removeByName(chart_name)

    rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    rect.X, rect.Y, rect.Width, rect.Height = rect_coords

    sc, sr, ec, er = live_range_coords
    chart_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    chart_addr.Sheet = sheet.RangeAddress.Sheet
    chart_addr.StartColumn = sc
    chart_addr.StartRow = sr
    chart_addr.EndColumn = ec
    chart_addr.EndRow = er

    # Inserción con rangos vivos de celda (primer True = contiene encabezados, segundo True = contiene etiquetas)
    charts.addNewByName(chart_name, rect, (chart_addr,), True, True)
    chart = charts.getByName(chart_name).getEmbeddedObject()

    # Formato visual docente
    chart.setPropertyValue("HasMainTitle", bool(title))
    if title:
        chart.Title.String = title

    diagram = chart.getFirstDiagram()
    if diagram:
        diagram.Wall.setPropertyValue("FillStyle", FillStyle.NONE)
        diagram.setPropertyValue("DataCaption", 1)  # Mostrar valor numérico sobre barras

    return charts.getByName(chart_name)


def inject_pivot_chart_loext(ods_path: str, datapilot_name: str):
    """
    Inyecta quirúrgicamente el atributo nativo loext:data-pilot-source en el gráfico
    dentro del archivo ODS guardado, activando las solapas interactivas de filtrado.
    Garantiza además que xlink:href='..' apunte a la planilla y no a un local-table.
    """
    abs_path = os.path.abspath(ods_path)
    tmp_path = abs_path + ".tmp.zip"

    with zipfile.ZipFile(abs_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith("content.xml") and item.filename != "content.xml":
                text = data.decode("utf-8")
                # Asegurar enlace ascendente
                text = text.replace('xlink:href="."', 'xlink:href=".."')
                
                # Inyectar namespace loext si falta
                if 'xmlns:loext=' not in text:
                    text = text.replace(
                        '<office:document-content',
                        '<office:document-content xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext/1.0"'
                    )
                # Inyectar atributo de tabla dinámica
                if f'loext:data-pilot-source="{datapilot_name}"' not in text:
                    text = re.sub(
                        r'<chart:chart(\s+[^>]*)>',
                        rf'<chart:chart\1 loext:data-pilot-source="{datapilot_name}">',
                        text,
                        count=1
                    )
                data = text.encode("utf-8")
            zout.writestr(item, data)

    os.replace(tmp_path, abs_path)


# ==============================================================================
# 5. AUDITORÍA AUTOMÁTICA DE CALIDAD Y ENTREGA
# ==============================================================================

def audit_workbook(doc) -> Dict[str, Any]:
    """
    Audita todo el libro de cálculo buscando errores de sintaxis o visualización:
    - Celdas con error: #¿NOMBRE?, #¡VALOR!, #REF!, Err:510, etc.
    - Filas de encabezado o parámetros colapsadas u ocultas.
    """
    issues = []
    sheets = doc.getSheets()
    error_patterns = ["#¿NOMBRE?", "#¡VALOR!", "#REF!", "#N/A", "Err:510", "Err:502", "Err:508"]

    for s_idx in range(sheets.getCount()):
        sheet = sheets.getByIndex(s_idx)
        sheet_name = sheet.getName()

        # Detectar dinámicamente la fila de encabezados
        # Si A4 o B4 tiene texto de cabecera (ej. Legajo, Apellido, Sección), el encabezado es fila 4 (índice 3).
        # Si A1 tiene texto (ej. Id, Nombre), el encabezado es fila 1 (índice 0).
        header_idx = 0
        for test_r in range(6):
            txt_a = sheet.getCellByPosition(0, test_r).getString().strip().lower()
            txt_b = sheet.getCellByPosition(1, test_r).getString().strip().lower()
            if any(k in txt_a or k in txt_b for k in ["legajo", "apellido", "nombre", "sección", "seccion", "escalafón", "escalafon"]):
                header_idx = test_r
                break

        # Verificar que ninguna fila desde 0 hasta el encabezado esté colapsada u oculta
        rows = sheet.getRows()
        for r in range(header_idx + 1):
            if not rows.getByIndex(r).getPropertyValue("IsVisible"):
                issues.append(f"Hoja '{sheet_name}': Fila {r+1} (encabezado/parámetro superior) está colapsada/oculta.")

        # Chequear celdas con errores (recorriendo cursor usado)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        max_col = min(cursor.RangeAddress.EndColumn + 1, 30)
        max_row = min(cursor.RangeAddress.EndRow + 1, 300)

        for r in range(max_row):
            for c in range(max_col):
                cell = sheet.getCellByPosition(c, r)
                txt = cell.getString().strip()
                for err in error_patterns:
                    if err in txt:
                        issues.append(f"Hoja '{sheet_name}', Celda ({c},{r}): Error de cálculo '{txt}'")

    return {
        "status": "APROBADO" if not issues else "OBSERVADO",
        "issues_count": len(issues),
        "issues": issues
    }
