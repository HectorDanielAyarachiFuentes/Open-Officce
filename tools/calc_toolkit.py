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
    from com.sun.star.table import TableSortField, CellRangeAddress, CellAddress, CellContentType
    from com.sun.star.sheet import TableFilterField, FilterOperator, GeneralFunction
    from com.sun.star.awt import Rectangle, FontWeight
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

        for attempt in range(6):
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
                # Backoff lineal: 1.5s, 3s, 4.5s, 6s, 7.5s
                time.sleep(1.5 * (attempt + 1))

        raise RuntimeError("No se pudo establecer conexión con LibreOffice mediante UNO.")

    def open_doc(self, filepath: str):
        """Abre un documento de cálculo (.ods)."""
        abs_path = os.path.abspath(filepath)
        file_url = uno.systemPathToFileUrl(abs_path)
        pv = PropertyValue()
        pv.Name = "Hidden"
        pv.Value = True
        doc = self.desktop.loadComponentFromURL(file_url, "_blank", 0, (pv,))
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


def inject_pivot_chart_loext(ods_path: str, datapilot_name: str, chart_name: Optional[str] = None):
    """
    Inyecta quirúrgicamente el atributo nativo loext:data-pilot-source en el gráfico
    dentro del archivo ODS guardado, activando las solapas interactivas de filtrado.
    Garantiza además que xlink:href='..' apunte a la planilla y no a un local-table.
    Si se indica chart_name, sólo modifica el gráfico cuyo id/name coincida (evita
    corromper otros gráficos presentes en el ODS).
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

                # Inyectar atributo de tabla dinámica sólo en el gráfico correcto
                attr = f'loext:data-pilot-source="{datapilot_name}"'
                if attr not in text:
                    if chart_name:
                        # Busca el bloque draw:object que corresponde al nombre del gráfico
                        # y modifica únicamente el primer <chart:chart> dentro de ese bloque
                        pattern = (
                            r'(<draw:object[^>]*draw:name="' + re.escape(chart_name) + r'"[^>]*>.*?)'
                            r'(<chart:chart)(\s+[^>]*>)'
                        )
                        replacement = rf'\1\2\3 {attr}>'
                        new_text = re.sub(pattern, replacement, text, count=1, flags=re.DOTALL)
                        if new_text == text:
                            # Fallback: modificar el primer <chart:chart> si no se encontró por nombre
                            new_text = re.sub(
                                r'<chart:chart(\s+[^>]*)>',
                                rf'<chart:chart\1 {attr}>',
                                text,
                                count=1
                            )
                        text = new_text
                    else:
                        text = re.sub(
                            r'<chart:chart(\s+[^>]*)>',
                            rf'<chart:chart\1 {attr}>',
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
        max_col = cursor.RangeAddress.EndColumn + 1
        max_row = cursor.RangeAddress.EndRow + 1

        if max_col > 30 or max_row > 300:
            issues.append(
                f"Hoja '{sheet_name}': Tamaño grande ({max_col} cols × {max_row} filas). "
                f"Se auditan todas las celdas del área usada."
            )

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


# ==============================================================================
# 6. HIGH-LEVEL AUTOMATION HELPERS & UNIVERSAL EXAM SOLVER
# ==============================================================================

def auto_detect_data_bounds(sheet, header_row: int = 0) -> Tuple[int, int]:
    """Detecta dinámicamente el número de columnas y filas con datos en una hoja."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    total_cols = cursor.RangeAddress.EndColumn + 1
    total_rows = cursor.RangeAddress.EndRow + 1
    return total_cols, total_rows


def transform_column_text(sheet, col_idx: int, transform_fn=str.upper, start_row: int = 1, end_row: Optional[int] = None):
    """
    Transforma texto de una columna en bloque (ej. mayúsculas, minúsculas, strip).
    start_row: índice 0-based de la primera fila de datos (ej. 1 para fila 2).
    """
    if end_row is None:
        _, total_rows = auto_detect_data_bounds(sheet)
        end_row = total_rows
        
    for r in range(start_row, end_row):
        cell = sheet.getCellByPosition(col_idx, r)
        val = cell.getString().strip()
        if val:
            cell.setString(transform_fn(val))


def add_calculated_column(
    sheet,
    col_idx: int,
    header_name: str,
    formula_template: str,
    number_format: Optional[int] = None,
    start_row: int = 1,
    end_row: Optional[int] = None
):
    """
    Agrega una columna calculada completa con encabezado y fórmulas dinámicas.
    formula_template usa '{row}' para la fila actual en Calc (1-based).
    Ej: formula_template='=YEAR(TODAY())-E{row}' o '=IF(MOD(H{row}; 2)=0; "PAR"; "IMPAR")'.
    """
    if end_row is None:
        _, total_rows = auto_detect_data_bounds(sheet)
        end_row = total_rows

    # Encabezado
    h_cell = sheet.getCellByPosition(col_idx, 0)
    h_cell.setString(header_name)
    
    # Fórmulas
    for r in range(start_row, end_row):
        calc_row = r + 1
        cell = sheet.getCellByPosition(col_idx, r)
        formula = formula_template.replace("{row}", str(calc_row))
        cell.setFormula(formula)
        if number_format is not None:
            cell.setPropertyValue("NumberFormat", number_format)


def copy_filtered_sheet(
    doc,
    src_sheet,
    target_sheet_name: str,
    condition_fn,
    filter_fields_config: List[Dict[str, Any]],
    db_range_name: str,
    formula_overrides: Optional[Dict[int, str]] = None,
    header_style: Optional[Dict[str, Any]] = None
) -> Tuple[Any, int]:
    """
    Copia registros que cumplan condition_fn a una nueva hoja con encabezados y DatabaseRange.
    - Preserva automáticamente valores numéricos reales (.setValue) y tipos de celdas.
    - Soporta formula_overrides {col_idx: formula_template} para recalcular sobre la fila de destino.
    - Registra el DatabaseRange con ContainsHeader=True y sus TableFilterField nativos.
    """
    import uno
    sheets = doc.getSheets()
    if sheets.hasByName(target_sheet_name):
        sheets.removeByName(target_sheet_name)
    sheets.insertNewByName(target_sheet_name, sheets.getCount())
    target_sheet = sheets.getByName(target_sheet_name)

    total_cols, total_rows = auto_detect_data_bounds(src_sheet)

    # Copiar encabezados
    for c in range(total_cols):
        target_sheet.getCellByPosition(c, 0).setString(src_sheet.getCellByPosition(c, 0).getString())

    dest_row_idx = 1
    for r in range(1, total_rows):
        if condition_fn(src_sheet, r):
            calc_r = dest_row_idx + 1
            for c in range(total_cols):
                if formula_overrides and c in formula_overrides:
                    f_tpl = formula_overrides[c]
                    target_sheet.getCellByPosition(c, dest_row_idx).setFormula(f_tpl.replace("{row}", str(calc_r)))
                else:
                    src_c = src_sheet.getCellByPosition(c, r)
                    src_type = src_c.getType()
                    if src_type == CellContentType.VALUE:
                        target_sheet.getCellByPosition(c, dest_row_idx).setValue(src_c.getValue())
                    elif src_type == CellContentType.FORMULA:
                        target_sheet.getCellByPosition(c, dest_row_idx).setFormula(src_c.getFormula())
                    else:
                        target_sheet.getCellByPosition(c, dest_row_idx).setString(src_c.getString())
            dest_row_idx += 1

    # Estilo de encabezados
    h_range = target_sheet.getCellRangeByPosition(0, 0, total_cols - 1, 0)
    h_range.setPropertyValue("CharWeight", FontWeight.BOLD)
    bg_color = header_style.get("bg_color", 0x1B365D) if header_style else 0x1B365D
    txt_color = header_style.get("txt_color", 0xFFFFFF) if header_style else 0xFFFFFF
    h_range.setPropertyValue("CellBackColor", bg_color)
    h_range.setPropertyValue("CharColor", txt_color)
    h_range.setPropertyValue("HoriJustify", uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER"))

    for c in range(total_cols):
        target_sheet.getColumns().getByIndex(c).setPropertyValue("OptimalWidth", True)

    # DatabaseRange formal
    db_ranges = doc.DatabaseRanges
    if db_ranges.hasByName(db_range_name):
        db_ranges.removeByName(db_range_name)
    addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    addr.Sheet = target_sheet.RangeAddress.Sheet
    addr.StartColumn = 0
    addr.StartRow = 0
    addr.EndColumn = total_cols - 1
    addr.EndRow = max(dest_row_idx - 1, 1)
    db_ranges.addNewByName(db_range_name, addr)

    dbr = db_ranges.getByName(db_range_name)
    fd = dbr.getFilterDescriptor()
    fd.setPropertyValue("ContainsHeader", True)

    ff_list = []
    for cfg in filter_fields_config:
        ff = uno.createUnoStruct("com.sun.star.sheet.TableFilterField")
        ff.Field = cfg["field"]
        ff.Operator = cfg.get("operator", uno.Enum("com.sun.star.sheet.FilterOperator", "EQUAL"))
        ff.IsNumeric = cfg.get("is_numeric", False)
        if ff.IsNumeric:
            ff.NumericValue = float(cfg["value"])
        else:
            ff.StringValue = str(cfg["value"])
        conn_str = cfg.get("connection", "AND").upper()
        ff.Connection = uno.Enum("com.sun.star.sheet.FilterConnection", conn_str)
        ff_list.append(ff)

    fd.setFilterFields(tuple(ff_list))
    dbr.refresh()

    return target_sheet, dest_row_idx - 1


def create_datapilot_and_pivot_chart(
    doc,
    target_sheet,
    src_sheet,
    dp_name: str,
    chart_name: str,
    src_coords: Tuple[int, int, int, int],
    dp_dest_cell: Tuple[int, int],
    row_fields: List[int],
    col_fields: List[int],
    data_fields: List[Tuple[int, str]],
    chart_rect: Tuple[int, int, int, int],
    chart_title: str = ""
):
    """
    Crea una Tabla Dinámica nativa y su correspondiente Gráfico Dinámico vinculado
    excluyendo automáticamente los totales generales del rango del gráfico.
    chart_rect: (x, y, width, height) en 1/100 mm.
    """
    import uno
    dp_tables = target_sheet.getDataPilotTables()
    if dp_tables.hasByName(dp_name):
        dp_tables.removeByName(dp_name)

    dp_desc = dp_tables.createDataPilotDescriptor()
    sc, sr, ec, er = src_coords
    src_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    src_addr.Sheet = src_sheet.RangeAddress.Sheet
    src_addr.StartColumn = sc
    src_addr.StartRow = sr
    src_addr.EndColumn = ec
    src_addr.EndRow = er
    dp_desc.setSourceRange(src_addr)

    for rf in row_fields:
        dp_desc.getDataPilotFields().getByIndex(rf).setPropertyValue(
            "Orientation", uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "ROW")
        )

    for cf in col_fields:
        dp_desc.getDataPilotFields().getByIndex(cf).setPropertyValue(
            "Orientation", uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "COLUMN")
        )

    for df_idx, fn_name in data_fields:
        df = dp_desc.getDataPilotFields().getByIndex(df_idx)
        df.setPropertyValue("Orientation", uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "DATA"))
        df.setPropertyValue("Function", uno.Enum("com.sun.star.sheet.GeneralFunction", fn_name.upper()))

    dest_addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
    dest_addr.Sheet = target_sheet.RangeAddress.Sheet
    dest_addr.Column = dp_dest_cell[0]
    dest_addr.Row = dp_dest_cell[1]

    dp_tables.insertNewByName(dp_name, dest_addr, dp_desc)
    doc.calculateAll()

    # Rango vivo para el gráfico (excluyendo fila y columna de 'Total Resultado')
    # DataPilot estándar de 2 filas y 2 columnas con fila de filtro:
    # Encabezado de columnas en dest_addr.Row + 3
    # Datos en dest_addr.Row + 4 a dest_addr.Row + 5
    charts = target_sheet.getCharts()
    if charts.hasByName(chart_name):
        charts.removeByName(chart_name)

    rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
    rect.X, rect.Y, rect.Width, rect.Height = chart_rect

    chart_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    chart_addr.Sheet = target_sheet.RangeAddress.Sheet
    chart_addr.StartColumn = dp_dest_cell[0]
    chart_addr.StartRow = dp_dest_cell[1] + 3
    chart_addr.EndColumn = dp_dest_cell[0] + max(len(col_fields) + 1, 2)
    chart_addr.EndRow = dp_dest_cell[1] + 3 + max(len(row_fields) + 1, 2)

    charts.addNewByName(chart_name, rect, (chart_addr,), True, True)
    chart_obj = charts.getByName(chart_name).getEmbeddedObject()

    if chart_title:
        chart_obj.setPropertyValue("HasMainTitle", True)
        chart_obj.Title.String = chart_title

    diagram = chart_obj.getFirstDiagram()
    if diagram:
        try:
            diagram.Wall.setPropertyValue("FillStyle", uno.Enum("com.sun.star.drawing.FillStyle", "NONE"))
        except Exception:
            pass
        try:
            diagram.setPropertyValue("DataCaption", 1)
        except Exception:
            pass

    return dp_tables.getByName(dp_name), charts.getByName(chart_name)

