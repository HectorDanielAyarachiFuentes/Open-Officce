# Protocolo Maestro de Automatización y Resolución — LibreOffice

Este documento establece las directivas supremas para la automatización, resolución de trabajos prácticos y exámenes de **LibreOffice (Calc, Writer, Impress)** en este repositorio.

Las reglas detalladas operativas se encuentran modularizadas en la carpeta [`.agents/rules/`](file:///c:/Users/Ramoncito/.antigravity-ide/Open%20Officce/.agents/rules/):
- [01-libreoffice-automation.md](file:///c:/Users/Ramoncito/.antigravity-ide/Open%20Officce/.agents/rules/01-libreoffice-automation.md): Motor UNO nativo, DatabaseRanges, Filtros, Tablas Dinámicas y Pivot Charts.
- [02-criterios-catedra.md](file:///c:/Users/Ramoncito/.antigravity-ide/Open%20Officce/.agents/rules/02-criterios-catedra.md): Criterios de cátedra, sintaxis Calc, posicionamiento y evaluación docente.
- [03-guias-resolucion.md](file:///c:/Users/Ramoncito/.antigravity-ide/Open%20Officce/.agents/rules/03-guias-resolucion.md): Guías de resolución rápida de Práctica 01, 02, 03 y Final 2024.

---

## 1. Mandamientos Inquebrantables de Ejecución

1. **Entorno 100% Nativo LibreOffice (Python + UNO):**
   - Ejecutable: `C:\Program Files\LibreOffice\program\soffice.exe`
   - Intérprete: `C:\Program Files\LibreOffice\program\python.exe`
   - Conexión: named pipe local (`soffice.exe --headless --accept="pipe,name=...;urp;StarOffice.ComponentContext"`).
   - **Prohibido:** Usar librerías genéricas (`pandas`, `openpyxl`) o re-serializar `content.xml` con `xml.etree` (evita corrupción de namespaces y el error `Err:510`).

2. **Prohibición Terminante de Microsoft Excel:**
   - La cátedra evalúa exclusivamente en LibreOffice Calc. Fórmulas resueltas o guardadas desde Excel invalidan la entrega.

3. **Tablas Dinámicas (DataPilot) y Gráficos Dinámicos con Solapas Desplegables:**
   - Todo agrupamiento, resumen o categorización debe resolverse con `sheet.getDataPilotTables().insertNewByName(...)`. Prohibido simular con fórmulas manuales cuando se pide tabla dinámica.
   - Todo gráfico vinculado a una Tabla Dinámica debe ser un **Pivot Chart oficial**, inyectando en su XML el atributo nativo `loext:data-pilot-source="<NombreTablaDinamica>"`. Esto activa las solapas de filtrado interactivas (`Campo ▼`, `Datos ▼`).
   - **Prohibido usar `chart.getData().setData(...)`:** Dicho método desconecta el gráfico de la planilla y crea un `<table:table table:name="local-table">` estático con `xlink:href="."`. El gráfico DEBE vincularse pasando directamente la tupla de rango vivo `(dp_range,)` a `addNewByName(..., True, True)` con `xlink:href=".."`.
   - Jamás incluir la fila ni columna de "Total general" en los rangos del gráfico.

4. **Persistencia Dual de Filtros Estándar y Preservación de Encabezados (Filas 1 a 4):**
   - **Preservación de la parte superior:** Al aplicar filtros estándar en una hoja, Calc evalúa las filas de parámetros y encabezados y suele colapsarlas/ocultarlas (`IsVisible = False`). Es **estrictamente obligatorio** restaurar la visibilidad de todas las filas de título, parámetros y encabezados (ej. `for r in range(4): sheet.getRows().getByIndex(r).setPropertyValue("IsVisible", True)`). La parte superior de la tabla NUNCA debe faltar ni estar oculta.
   - Registrar siempre un `DatabaseRange` con `doc.DatabaseRanges.addNewByName(...)`.
   - Asignar `TableFilterField` y ejecutar `refresh()`.
   - Garantizar que el cursor activo en la hoja quede situado dentro del rango de datos o fila de encabezados para que Calc abra el diálogo con los filtros visibles.

5. **Ordenamiento Nativo — `TableSortField` Obligatorio con `uno.Any`:**
   - **Prohibición terminante de `com.sun.star.util.SortField`:** En rangos de hojas de cálculo (`SheetCellRange`), `com.sun.star.util.SortField` falla silenciosamente y NO ordena los datos.
   - Usar **exclusivamente** `uno.createUnoStruct("com.sun.star.table.TableSortField")`.
   - Al configurar `SortFields` en el descriptor de ordenamiento, envolver siempre con:
     `uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2))`
     (la asignación de tupla directa de Python falla en silencio).
   - Configurar claves múltiples respetando estrictamente la jerarquía solicitada (ej. Clave 1 = Apellido, Clave 2 = Nombres).

6. **Fórmulas y Localización:**
   - En el API UNO (`cell.setFormula(...)`), los nombres de función van siempre en **inglés estándar**: `=COUNTIF(...)`, `=COUNTIFS(...)`, `=SUM(...)`, `=AVERAGE(...)`, `=IF(...)`, `=MOD(...)`. Calc los traduce automáticamente al español en la interfaz.

7. **Principio de Entregable Limpio:**
   - La estrella exclusiva de la entrega es el archivo `.ods` resuelto.
   - Prohibido generar informes `.odt`, `.pdf` o `.txt` innecesarios dentro de las carpetas de resolución a menos que la consigna lo pida explícitamente. Las respuestas conceptuales van en una hoja de la planilla o en un único `.md` ligero.

---

## 2. Mapa Rápido de Prácticas
- **Práctica 02:** Fórmulas con referencias absolutas (`$C$1`, `$C$2`), filtros estándar en 2A, 2B, 2C con ordenamiento dual, DataPilot y Pivot Charts en 3A y 3B.
- **Práctica 03:** Encuesta sin columna de Timestamp (25 columnas), filtros estándar copiados a hojas nuevas con DatabaseRange propio, funciones condicionales y DataPilot de cruce Software vs Modalidad.
- **Final 2024:** Hojas `Consignas`, `Personas`, `Filtro`, `Tabla Dinamica` (DataPilot con función PROMEDIO y Pivot Chart vertical con solapas desplegables).
