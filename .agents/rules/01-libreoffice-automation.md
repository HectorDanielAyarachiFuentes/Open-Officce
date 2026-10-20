# Technical Automation Protocol — LibreOffice / OpenOffice
# File: .agents/rules/01-libreoffice-automation.md

# 1. Entorno de Ejecución y Motor Nativo
- **Ejecutable oficial:** `C:\Program Files\LibreOffice\program\soffice.exe`
- **Intérprete Python con UNO:** `C:\Program Files\LibreOffice\program\python.exe`
- **Protocolo de conexión:** Conectar siempre mediante named pipes o sockets locales:
  `soffice.exe --headless --accept="pipe,name=<nombre_pipe>;urp;StarOffice.ComponentContext"`
- **Prohibición de librerías genéricas e IAs convencionales:** Queda estrictamente prohibido usar librerías genéricas de Python como `pandas` o reescribir manualmente el XML interno con serializadores genéricos (`xml.etree.ElementTree`, `minidom`) que no respeten los 30+ namespaces de OpenDocument (ODF), ya que destruyen los rangos de base de datos, metadatos y corrompen las fórmulas (`Err:510`).

---

# 2. Protocolo para Filtros Estándar y Sensibilidad al Cursor
Para que los docentes o usuarios vean los criterios preconfigurados al abrir `Datos > Más filtros > Filtro estándar...`:
1. **Sensibilidad al Cursor (Contexto de Celda Activa):**
   - En LibreOffice Calc, el diálogo de filtro depende estrictamente de la celda activa. Si el cursor está fuera de la tabla, Calc abre el diálogo en blanco.
   - **Regla:** El cursor debe situarse obligatoriamente sobre cualquier celda interna de la tabla o fila de encabezados (ej. `A4`, `C4`, `A1`).
2. **Persistencia Dual en DatabaseRanges:**
   - Registrar siempre un rango de base de datos formal nombrado mediante `doc.DatabaseRanges.addNewByName("<Nombre_Rango>", cellRangeAddress)`.
   - Configurar descriptor: `fd.setPropertyValue("ContainsHeader", True)` y asignar campos `com.sun.star.sheet.TableFilterField`.
   - Invocar `dbr.refresh()` para vincularlo a la vista.
3. **Cobertura Ampliada en content.xml (__Anonymous_Sheet_DB__):**
   - Sincronizar en `content.xml` el rango anónimo interno (`__Anonymous_Sheet_DB__<Index>`) con sus condiciones `<table:filter-and>`, para que Calc asocie los filtros ante cualquier clic.
4. **Copia de Resultados a Hojas Nuevas:**
   - Si la consigna exige volcar el filtro en hoja nueva, usar `CopyOutputData = True` apuntando a la dirección de destino, asegurando que ambas hojas tengan su propio `DatabaseRange` registrado.
5. **Preservación Obligatoria de Filas Superiores (Encabezados y Parámetros):**
   - Al ejecutar un filtro a nivel de hoja, Calc evalúa las filas anteriores al rango de datos contra el criterio. Si la fila 4 es el encabezado y las filas 1 a 3 contienen títulos o parámetros (ej. `$C$1`, `$C$2`), Calc oculta esas filas (`IsVisible = False`), provocando que la tabla parezca "recortada" o sin su encabezado.
   - **Regla Estricta:** Inmediatamente después de aplicar cualquier filtro estándar a una hoja, restaurar de forma explícita la visibilidad de todas las filas previas a los registros:
     ```python
     # Si el encabezado está en la fila 4 (índice 3), restaurar filas 0 a 3:
     for r in range(header_row_index + 1):
         sheet.getRows().getByIndex(r).setPropertyValue("IsVisible", True)
     ```

---

# 3. Prevención de Err:510 (Serializadores XML)
- **Causa Raíz:** Serializar `content.xml` con librerías genéricas elimina prefijos de namespaces (`table:formula="of:=..."` se convierte en `=of:=...`), provocando `Err:510`.
- **Regla Estricta:** Reconstruir archivos siempre con el API UNO nativo (`soffice.exe --headless` + `python.exe`). Ajustes en `content.xml` deben limitarse a inyecciones o sustituciones quirúrgicas UTF-8 sin re-serializar el árbol completo.

---

# 4. Manejo de Fórmulas y Localización (API UNO en Inglés)
Al invocar `cell.setFormula(...)`, **suministrar siempre el nombre de función en inglés estándar**:
- `=COUNTIF(...)` → En español se muestra como `=CONTAR.SI(...)`.
- `=COUNTIFS(...)` → En español `=CONTAR.SI.CONJUNTO(...)`.
- `=SUM(...)` → `=SUMA(...)`.
- `=COUNTA(...)` → `=CONTARA(...)`.
- `=AVERAGEIF(...)` → `=PROMEDIO.SI(...)`.
- `=AVERAGE(...)` → `=PROMEDIO(...)`.
- `=IF(...)` → `=SI(...)`.
- `=MOD(...)` → `=RESIDUO(...)`.
- `=ISEVEN(...)` → `=ES.PAR(...)`.
- `=UPPER(...)` → `=MAYUSC(...)`.
- `=YEAR(TODAY())` → `=AÑO(HOY())`.
*Advertencia:* Pasar nombres en español directamente a `.setFormula(...)` genera `#¿NOMBRE?`.

---

# 5. Gráficos Estadísticos y Visualización
- **Prevención de Gráficos Vacíos en Headless:**
  - Invocar `doc.calculateAll()` antes de generar gráficos para evitar valores `NaN` (`2.22e-308`).
- **Frontend Embellecedor:**
  - Fondo transparente: `diagram.Wall.setPropertyValue("FillStyle", uno.Enum("com.sun.star.drawing.FillStyle", "NONE"))`.
  - Etiquetas de datos numéricas: `diagram.setPropertyValue("DataCaption", 1)`.
  - Títulos descriptivos claros y títulos en ejes X e Y.
- **Posicionamiento Visible en Pantalla:**
  - Posicionar gráficos en el campo visual inmediato: al lado de la tabla (`X = 11000` a `18000`, `Y = 1500`) o directamente debajo (`X = 36000` a `38000`, `Y = 4000` a `5000`). Nunca usar coordenadas desmedidas (`X > 100000`).

---

# 6. Tablas Dinámicas y Pivot Charts con Solapas de Filtrado
- **Obligatoriedad Absoluta de DataPilot Real:** En cualquier consigna donde se resuman o agrupen registros por categorías con gráfico comparativo, **es mandatorio generar una Tabla Dinámica nativa (`DataPilotTable`) y vincular el gráfico como Gráfico Dinámico (Pivot Chart)**. Queda prohibido simular con fórmulas estáticas.
- **Protocolo UNO Nativo:**
  1. `dp_desc = sheet.getDataPilotTables().createDataPilotDescriptor()`.
  2. `dp_desc.setSourceRange(cellRangeAddress)`.
  3. Campos: `Orientation = ROW`, datos con `Function = COUNT | SUM | AVERAGE`.
  4. Inserción: `sheet.getDataPilotTables().insertNewByName("<Nombre>", outAddr, dp_desc)`.
  5. Obtención de rango: Determinar `chart_rng` a partir del rango del DataPilot pero **excluyendo estrictamente la fila y columna de 'Total general'**.
  6. Inserción del Gráfico Dinámico:
     `sheet.getCharts().addNewByName(name, rect, (chart_rng,), True, True)`.
- **Prohibición Terminante de `chart.getData().setData(...)`:**
  - **Causa del fallo:** Ejecutar `.setData(...)` desconecta el gráfico de los datos vivos de la hoja y genera una tabla interna estática `<table:table table:name="local-table">` con `xlink:href="."`. Esto destruye el carácter dinámico del Pivot Chart.
  - **Regla:** El gráfico debe alimentarse exclusivamente del rango vivo de celdas registrado en su creación.
- **Inyección de Atributo Pivot Chart (`loext:data-pilot-source`):**
  - En el archivo `content.xml` interno del gráfico (`Object 1/content.xml`):
    - Asegurar que `<chart:chart>` mantenga `xlink:href=".."` (apuntando al documento contenedor).
    - Inyectar el atributo `loext:data-pilot-source="<NombreDataPilot>"`.
    - Esto activa en la interfaz de Calc los botones desplegables nativos de filtrado (`<NombreCampo> ▼` y `Datos ▼`).

---

# 7. Ordenamiento Nativo de Rangos (TableSortField con uno.Any)
- **Prohibición Estricta de `com.sun.star.util.SortField`:**
  En hojas de cálculo (`SheetCellRange`), el struct `com.sun.star.util.SortField` **falla silenciosamente** (no produce error pero no ordena las filas).
- **Uso Obligatorio de `com.sun.star.table.TableSortField`:**
  Para ordenar rangos en Calc, se debe instanciar `com.sun.star.table.TableSortField` y envolver la tupla en `uno.Any`:
```python
sort_desc = rng.createSortDescriptor()

# Clave 1: Ej. Apellido (Columna B -> Field = 1)
sf1 = uno.createUnoStruct("com.sun.star.table.TableSortField")
sf1.Field = 1
sf1.IsAscending = True
sf1.IsCaseSensitive = False

# Clave 2: Ej. Nombres (Columna C -> Field = 2)
sf2 = uno.createUnoStruct("com.sun.star.table.TableSortField")
sf2.Field = 2
sf2.IsAscending = True
sf2.IsCaseSensitive = False

# Asignación obligatoria mediante uno.Any
for prop in sort_desc:
    if prop.Name == "SortFields":
        prop.Value = uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2))
    elif prop.Name == "ContainsHeader":
        prop.Value = True

# Ejecución del ordenamiento
rng.sort(tuple(sort_desc))
```

---

# 8. Principio de Entregable Limpio (La Estrella es la Resolución .ODS)
- En `Resolucion/`, el foco excluyente es el archivo `.ods` (y `.csv` si lo pide la consigna).
- Prohibido generar informes de procesador de textos (`.odt`, `.pdf`, `.txt`) o carpetas redundantes dentro de `Resolucion/`.
- Justificaciones teóricas se resuelven en la propia planilla o en un único `.md` ligero.
