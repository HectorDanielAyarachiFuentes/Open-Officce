# Reglas de Operación y Automatización Nativa — LibreOffice y OpenOffice

Esta directiva establece los protocolos, pautas pedagógicas y requisitos obligatorios para manipular, resolver trabajos prácticos y automatizar documentos de **LibreOffice Calc, Writer e Impress** desde el entorno del asistente.

---

## 1. Entorno de Ejecución y Motor Nativo
- **Ejecutable oficial:** `C:\Program Files\LibreOffice\program\soffice.exe`
- **Intérprete Python con UNO:** `C:\Program Files\LibreOffice\program\python.exe`
- **Protocolo de conexión:** Conectar siempre mediante named pipes o sockets locales:
  `soffice.exe --headless --accept="pipe,name=<nombre_pipe>;urp;StarOffice.ComponentContext"`
- **Prohibición de librerías genéricas e IAs convencionales:** Queda estrictamente prohibido usar librerías genéricas de Python como `pandas` o reescribir manualmente el XML interno con serializadores genéricos (`xml.etree.ElementTree`, `minidom`) que no respeten los 30+ namespaces de OpenDocument (ODF), ya que destruyen los rangos de base de datos, metadatos y corrompen las fórmulas (`Err:510`).

---

## 2. Protocolo Obligatorio para Filtros Estándar (Standard Filter) y Sensibilidad al Cursor
Para que los docentes o usuarios puedan ver los criterios preconfigurados al abrir el cuadro de diálogo **`Datos > Más filtros > Filtro estándar...`**:
1. **Sensibilidad al Cursor (Contexto de Celda Activa en la GUI):**
   - En la interfaz gráfica de LibreOffice Calc, el diálogo `Datos > Más filtros > Filtro estándar...` es estrictamente dependiente de la posición de la celda activa.
   - Si el usuario tiene seleccionada una celda **fuera de la tabla de datos** (por ejemplo, celdas de parámetros en filas 1 a 3 como `C2` con 'Descuentos 22,00%', títulos o celdas vacías), Calc interpreta la acción como la creación de un nuevo filtro en blanco para esa posición y abre el diálogo con todas las condiciones en `- ninguno -`.
   - **Regla de Operación:** Para visualizar o editar los criterios precargados en la GUI, el cursor **debe situarse obligatoriamente sobre cualquier celda de la tabla de datos o fila de encabezados** (ej. rango `A4:N108`).
2. **Persistencia Dual en `DatabaseRanges`:**
   - Registrar siempre un rango de base de datos formal nombrado mediante `doc.DatabaseRanges.addNewByName("<Nombre_Rango>", cellRangeAddress)`.
   - Configurar el descriptor con `fd = dbr.getFilterDescriptor()`, `fd.setPropertyValue("ContainsHeader", True)` y asignar los campos `com.sun.star.sheet.TableFilterField`.
   - Invocar `dbr.refresh()` para vincularlo a la vista de la hoja.
3. **Cobertura Ampliada en XML de Hoja (__Anonymous_Sheet_DB__):**
   - Como medida de blindaje técnico, sincronizar quirúrgicamente en `content.xml` el rango anónimo interno de la hoja (`__Anonymous_Sheet_DB__<Index>`) con cobertura expandida (`A1:XFD1048576`) y sus condiciones `<table:filter-and>`, de modo que Calc asocie los filtros incluso ante clics periféricos.
4. **Copia de Resultados a Hojas Nuevas:**
   - Si la consigna exige volcar el filtro en una hoja nueva (ej. `de 0 a 2 años`), usar `filter_desc.setPropertyValue("CopyOutputData", True)` apuntando a la dirección de celda de la nueva hoja, asegurando que tanto la hoja origen como la destino tengan su propio `DatabaseRange` registrado.

---

## 3. Prohibición de Serializadores XML Genéricos (Prevención de Err:510)
- **Causa Raíz de Err:510:** Al reescribir `content.xml` mediante `xml.etree.ElementTree` u otras herramientas que analizan y vuelven a serializar el árbol XML, se elimina el prefijo de namespace de las fórmulas de Calc (`table:formula="of:=..."` se serializa como `=of:=...`), provocando el error `Err:510 (Variable / Sintaxis ausente)` en Calc.
- **Regla Estricta:** Reconstruir archivos siempre mediante el API UNO nativo (`soffice.exe --headless` + `python.exe`). Cualquier ajuste en `content.xml` debe limitarse a sustitución quirúrgica de texto UTF-8 sobre `<table:database-ranges>` o inyección de atributos sin re-serializar el resto del documento.

---

## 4. Manejo de Fórmulas y Localización en Calc
- **Idioma de las Fórmulas en el API UNO:**
  - Al invocar `cell.setFormula(...)`, **siempre debe suministrarse el nombre de función en inglés estándar**:
    - `=COUNTIF(...)` → Se traduce y muestra como `=CONTAR.SI(...)` en Calc en español.
    - `=COUNTIFS(...)` → Se traduce y muestra como `=CONTAR.SI.CONJUNTO(...)`.
    - `=SUM(...)` → Se traduce y muestra como `=SUMA(...)`.
    - `=COUNTA(...)` → Se traduce y muestra como `=CONTARA(...)`.
    - `=AVERAGEIF(...)` → Se traduce y muestra como `=PROMEDIO.SI(...)`.
    - `=AVERAGE(...)` → Se traduce y muestra como `=PROMEDIO(...)`.
    - `=IF(...)` → Se traduce y muestra como `=SI(...)`.
    - `=MOD(...)` → Se traduce y muestra como `=RESIDUO(...)`.
  - *Advertencia:* Pasar funciones en español directamente a `.setFormula(...)` provoca el error de cálculo `#¿NOMBRE?`.

---

## 5. Gráficos Estadísticos y Visualización (Reglas de Renderizado y Estética)
- **Prevención de Gráficos Vacíos (Bug de Fórmulas en Headless):**
  - En modo headless, LibreOffice Calc no evalúa síncronamente las fórmulas recién insertadas antes de crear el gráfico, lo que genera valores `NaN` (`2.22e-308`), dejando el gráfico sin barras y con eje limitado a 12.
  - **Regla Obligatoria:** Invocar siempre `doc.calculateAll()` y poblar explícitamente el array de datos del gráfico mediante `chart.getData().setData(...)` con los valores numéricos calculados.
- **Frontend Embellecedor en Gráficos:**
  - Eliminar el fondo gris opaco por defecto: `diagram.Wall.setPropertyValue("FillStyle", uno.Enum("com.sun.star.drawing.FillStyle", "NONE"))`.
  - Activar etiquetas de datos numéricas sobre cada barra o sector: `diagram.setPropertyValue("DataCaption", 1)`.
  - Asignar paletas limpias y contrastadas (azul institucional `0x1A73E8`, verde `0x2E7D32`, teal, púrpura).
  - Títulos descriptivos con fuente de tamaño adecuado y títulos en ejes X e Y.
- **Cero Solapamiento y Posicionamiento Visible en Pantalla:**
  - Posicionar los gráficos **en el campo visual inmediato del usuario**, evitando márgenes desmedidos (ej. NUNCA colocar `rect.X = 110000`, ya que desplaza el gráfico más de 1 metro a la derecha y queda oculto fuera de pantalla).
  - Posicionamiento óptimo:
    - Inmediatamente a la derecha de la tabla: `rect.X = 11000` a `18000` (Col E a G), `rect.Y = 1500`.
    - O directamente debajo de la tabla: `rect.X = 38000` (Col P), `rect.Y = 5000` (Fila 9/10), abarcando columnas P a V.

---

## 6. Tablas Dinámicas (DataPilot) y Gráficos Dinámicos (Pivot Charts con Solapas de Filtrado)
- **Obligatoriedad de DataPilot Real:** Si una consigna solicita una Tabla Dinámica (o Piloto de Datos) y su correspondiente gráfico comparativo, **es mandatorio generar una Tabla Dinámica nativa** (`DataPilotTable`) y **vincular el gráfico como un Gráfico Dinámico real**. Queda estrictamente prohibido simular la tabla dinámica con fórmulas estáticas.
- **Protocolo de Creación en Python UNO:**
  1. Instanciar el descriptor: `dp_desc = sheet.getDataPilotTables().createDataPilotDescriptor()`.
  2. Asignar rango de origen: `dp_desc.setSourceRange(cellRangeAddress)`.
  3. Configurar campos:
     - Filas: `field.setPropertyValue("Orientation", uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "ROW"))`.
     - Datos: `field.setPropertyValue("Orientation", uno.Enum("com.sun.star.sheet.DataPilotFieldOrientation", "DATA"))`.
     - Función de agregación: `field.setPropertyValue("Function", uno.Enum("com.sun.star.sheet.GeneralFunction", "SUM" | "AVERAGE" | "COUNT"))`.
  4. Inserción: `sheet.getDataPilotTables().insertNewByName("<NombreTabla>", outCellAddress, dp_desc)`.
  5. Formato numérico: Formatear las celdas numéricas de la tabla dinámica generada (ej. `NumberFormat = 4` para moneda/decimales) para presentación visual impecable.
- **Activación de Gráficos Dinámicos con Solapas Interactivas ("Solapitas" de Filtrado):**
  - Para que LibreOffice Calc reconozca el gráfico como un **Gráfico Dinámico (Pivot Chart)** oficial y muestre las solapas desplegables interactivas en la interfaz (`<Campo> ▼`, `Datos ▼`):
    1. **Rango de Datos:** Asignar como rango de datos del gráfico el rango de salida completo de la Tabla Dinámica:
       `charts.addNewByName("<NombreGrafico>", rect, (dp_table.getOutputRange(),), True, True)`.
    2. **Atributo de Vinculación ODF:** Vincular el gráfico al motor DataPilot inyectando quirúrgicamente el atributo ODF nativo:
       `<chart:chart ... loext:data-pilot-source="<NombreTablaDinamica>">`
       dentro del archivo `<NombreGrafico>/content.xml` en el paquete ODS.
    3. **Snippet de Inyección Quirúrgica:**
       ```python
       import re, zipfile

       files_dict = {}
       with zipfile.ZipFile(ods_path, 'r') as z:
           for name in z.namelist():
               files_dict[name] = z.read(name)

       chart_xml_path = f"{chart_name}/content.xml"
       c_xml = files_dict[chart_xml_path].decode('utf-8')
       if f'loext:data-pilot-source="{dp_name}"' not in c_xml:
           pattern = r'(<chart:chart\b[^>]*)(>)'
           c_xml = re.sub(pattern, rf'\1 loext:data-pilot-source="{dp_name}"\2', c_xml)
           files_dict[chart_xml_path] = c_xml.encode('utf-8')

       with zipfile.ZipFile(ods_path, 'w', zipfile.ZIP_DEFLATED) as z:
           for name, data in files_dict.items():
               z.writestr(name, data)
       ```
    4. Esto activa el controlador de gráficos dinámicos de LibreOffice, renderizando automáticamente las solapas desplegables interactivas en la parte inferior y lateral del gráfico (`Vivienda ▼`, `Sección ▼`, `Datos ▼`).

---

## 7. Ordenamiento Nativo de Rangos en Calc (SortDescriptor y Envoltura uno.Any)
- **Causa del Fallo Silencioso:** Al configurar la propiedad `SortFields` de un `SortDescriptor` obtenido mediante `rng.createSortDescriptor()`, Python pasa una tupla de estructuras `TableSortField`. Si se asigna directamente como tupla estándar (`p.Value = (sf1, sf2)`), el puente UNO **ignora la asignación en silencio**, no produce ningún error pero deja la tabla sin ordenar.
- **Regla Estricta:** La asignación de `SortFields` debe envolverse obligatoriamente con el tipo de secuencia UNO explícito mediante `uno.Any`:
  ```python
  from com.sun.star.table import TableSortField

  sf1 = TableSortField()
  sf1.Field = 8  # Índice de columna relativo a la primera columna del rango (0 = columna inicial del rango)
  sf1.IsAscending = True
  sf1.IsCaseSensitive = False

  for p in sort_desc:
      if p.Name == "ContainsHeader":
          p.Value = True
      elif p.Name == "SortFields":
          p.Value = uno.Any("[]com.sun.star.table.TableSortField", (sf1,))
  rng.sort(tuple(sort_desc))
  ```
- **Indexación Relativa:** La propiedad `Field` es relativa a la primera columna del rango seleccionado (ej. si el rango es `A4:N108`, la columna `A` es índice 0; si el rango fuera `C4:N108`, la columna `C` sería índice 0).

---

## 8. Gestión de Bloqueo de Archivos en Windows (File Locking)
- En sistemas Windows, LibreOffice bloquea con bloqueo exclusivo (`[WinError 32]`) cualquier archivo abierto en su interfaz gráfica (`.~lock.<nombre>.ods#`).
- Si se detecta un error de permiso o bloqueo de archivo:
  1. Verificar si existen procesos activos de `soffice.exe` o `soffice.bin`.
  2. Advertir claramente al usuario que cierre la ventana de LibreOffice para liberar el archivo antes de intentar sobreescribirlo o sincronizarlo.

---

## 9. Criterios Pedagógicos, Buenas Prácticas y Pautas de Evaluación de Cátedra

Las docentes de la cátedra enfatizan una serie de reglas operativas, buenas prácticas y advertencias metodológicas determinantes para la aprobación de trabajos prácticos y parciales:

### 9.1. Uso Obligatorio de LibreOffice Calc (Software Libre)
- **Prohibición de Microsoft Excel:** La cátedra evalúa y califica específicamente bajo el entorno y la sintaxis nativa de **LibreOffice Calc**. Resolver o guardar archivos desde Excel invalida la entrega.
- **Incompatibilidad de Fórmulas:** Varias funciones difieren en nombre y comportamiento lógico entre ambas suites (por ejemplo, funciones condicionales y lógicas). Aunque una solución funcione en Excel, si falla o genera error en Calc, se considera reprobada.
- **Apertura de Archivos por Defecto:** Si el sistema operativo tiene instalados ambos programas, el doble clic puede abrir Excel por error. Debe hacerse siempre `Clic derecho > Abrir con > LibreOffice Calc`.

### 9.2. Organización y Gestión de Hojas (Pestañas)
- **Copiar antes de Modificar (No Mover):** Cada consigna requiere duplicar una hoja base (ej. `Personas`, `Encuestas`). Al acceder a `Mover o copiar hoja...`, **es obligatorio marcar la casilla "Copiar" y no "Mover"**, garantizando que la hoja original permanezca intacta con sus datos fuente.
- **Nombre Exacto de las Pestañas:** Es indispensable respetar rigurosamente la nomenclatura indicada en las consignas (ej. `1A`, `1B`, `2A`, `2B`, `2C`, `3A`, `3B`, `de 0 a 2 años`, etc.). Los docentes corrigen dirigiéndose directamente a la pestaña específica asignada; si el ejercicio se resolvió en la hoja equivocada o con nombre distorsionado, se califica negativamente.
- **Contexto de Hoja Activa:** Asegurarse de tener seleccionada y activa la pestaña correspondiente antes de aplicar filtros, ordenamientos o insertar columnas de cálculo.

### 9.3. Fórmulas, Sintaxis y Asistente de Funciones
- **El Signo Igual (`=`):** Es de uso obligatorio al inicio de cualquier celda de cálculo para que LibreOffice Calc la interprete como expresión evaluable y no como texto plano.
- **Referencias Absolutas vs. Relativas (`$`):**
  - **Parámetros fijos externos:** Cuando se vincula un valor fijo ubicado en una celda externa (montos fijos, porcentajes, tasas de descuento, incentivos), **debe fijarse obligatoriamente con el signo pesos** (ej. `$C$1`, `$C$2`) utilizando la tecla `F4`.
  - **Celdas variables por fila:** Las referencias a datos del registro individual (horas, legajo, escalafón) deben conservarse relativas (sin `$`) para que se adapten automáticamente al arrastrarse hacia abajo.
- **Cadenas de Texto entre Comillas Dobles (`" "`):** Todo texto dentro de una condición o argumento funcional (como `"F"`, `"M"`, `"Nivel 1"`, `"Sí"`, `"No"`) debe escribirse obligatoriamente entre comillas dobles. Su omisión genera inmediatamente los errores `#¿NOMBRE?` o `#¡VALOR!`.
- **Sensibilidad a Acentos y Tildes:** En funciones de búsqueda y conteo condicional (`CONTAR.SI`, `CONTAR.SI.CONJUNTO`), el criterio debe coincidir de forma idéntica con el texto original en los datos: `"Sí"` con tilde es tratado como un valor distinto de `"Si"` sin tilde.
- **Separadores Sintácticos (`:` vs. `;`):**
  - Los **dos puntos (`:`)** definen un rango continuo de celdas (ej. `A4:A108`, `G2:G178`).
  - El **punto y coma (`;`)** se utiliza obligatoriamente para separar argumentos funcionales o celdas individuales disjuntas.
- **Asistente para Funciones (`fx`):** Se recomienda utilizar el asistente (`Insertar > Función...` o `Ctrl + F2`) cuando existan dudas sobre la estructura de la función, verificando en la vista previa del diálogo que el resultado sea válido antes de presionar *Aceptar*.

### 9.4. Selección para Ordenar y Filtrar
- **Posicionarse dentro de la Tabla:** Antes de ejecutar `Datos > Ordenar...` o `Datos > Más filtros > Filtro estándar...`, el cursor debe posicionarse obligatoriamente sobre cualquier celda con datos de la tabla o en su fila de encabezados (nunca en celdas vacías o periféricas). Esto garantiza que Calc reconozca el rango de datos completo y conserve los encabezados de columna.
- **Filtro Automático vs. Filtro Estándar:**
  - **Filtro Automático:** Recomendado para selecciones rápidas o valores puntuales discretos.
  - **Filtro Estándar:** De uso obligatorio cuando la consigna solicita rangos numéricos continuos (ej. escalafón entre 6 y 10) o condiciones combinadas complejas con operadores lógicos (`Y` / `O`).

### 9.5. Tablas Dinámicas y Gráficos Estadísticos
- **Exclusión Estricta de Totales Generales:** Al seleccionar datos de una tabla o tabla dinámica para confeccionar un gráfico, **nunca se debe incluir la fila o columna de 'Total general' o resultados acumulados**. Incluir los totales duplica visualmente la escala y arruina la proporcionalidad de las barras o sectores.
- **Elementos Obligatorios del Gráfico:** Un gráfico no debe consistir en elementos vacíos; debe contar obligatoriamente con:
  1. Título descriptivo claro y temático.
  2. Nombres explícitos en los ejes (eje X horizontal y eje Y vertical).
  3. Etiquetas de datos numéricas sobre barras o sectores, formateadas con los decimales adecuados (generalmente 2 decimales para promedios o porcentajes).

### 9.6. Hábitos Operativos y Atajos Clave
- **Guardar Cambios Frecuentemente (`Ctrl + G`):** Guardar con periodicidad regular para proteger el progreso.
- **Navegación Rápida (`Ctrl + Flecha abajo`):** Desplazarse al instante hasta la última fila con datos sin tener que hacer scroll manual a lo largo de tablas extensas.
- **Autocompletado de Fórmulas:** Hacer doble clic en la esquina inferior derecha de la celda activa (cuadro de llenado) para autocompletar fórmulas hacia toda la columna de manera instantánea.
- **Resolución de Columnas Estrechas (`###`):** Si una columna visualiza numerales (`###`), no constituye un error de cálculo, sino ancho insuficiente de la celda; se resuelve haciendo doble clic en la línea divisoria del encabezado de columna o seleccionando *Ancho óptimo*.

---

## 10. Guía de Referencia Rápida — Práctica 02 (Planilla de Cálculo)
- **Hoja personas:**
  - `Sueldo` (Col L): `=K{row}*$C$1` (referencia absoluta al Monto fijo de `$1875` en `$C$1`).
  - `Descuento` (Col M): `=L{row}*$C$2` (referencia absoluta al `22,00%` en `$C$2`).
  - `Incentivo` (Col N): `=IF(K{row}>7; 1700; 0)`.
- **Hoja encuestas:**
  - `Calculo` (Col N): `=IF(MOD(A{row}; 2)=0; E{row}*2; E{row})`.
- **Hojas 1A y 1B (Ordenamiento múltiple):**
  - `1A`: Ordenar rango `A4:N108` por Localidad (Field 8) y Dirección (Field 6) con `uno.Any`.
  - `1B`: Ordenar rango `A1:N225` por Curso (Field 2), Sexo (Field 3) y Vivienda (Field 9) con `uno.Any`.
- **Hojas 2A, 2B y 2C (Filtros estándar):**
  - `2A`: Contaduría con escalafón 6 a 10. Ordenado nativamente según consigna docente explícita por Clave 1: Nombre (Field 2, Ascendente) y Clave 2: Apellido (Field 1, Ascendente) en `DatabaseRange` mediante `cells.sort(sd)` para persistencia dual de `<table:sort>` y `<table:filter>`. Cursor activo posicionado dentro de la tabla en celda C4 (`Nombres`).
  - `2B`: Contaduría con escalafón 6.
  - `2C`: Vivienda PROPIA y Electricidad SI.
- **Hojas 3A y 3B (Tablas Dinámicas y Gráficos Dinámicos con solapitas):**
  - `3A`: Tabla Dinámica de Vivienda (Alquila: 507, Propia: 611) + Gráfico Dinámico vertical ubicado en Col P a V (`X = 38000`, `Y = 5000`) con `loext:data-pilot-source="TablaDinamica_3A"` (`Vivienda ▼`, `Datos ▼`).
  - `3B`: Tabla Dinámica de Sueldo Promedio por Sección + Gráfico Dinámico vertical ubicado al costado en Col E (`X = 11000`, `Y = 1500`) con `loext:data-pilot-source="TablaDinamica_3B"` (`Sección ▼`, `Datos ▼`).

---

## 11. Guía de Referencia Rápida — Práctica 03 (Planilla de Cálculo)
- **Hoja Respuestas de formulario 1:**
  - `Punto 3`: Eliminación obligatoria de la columna A (`Timestamp` / `Marca temporal`), dejando 25 columnas de datos (Col A a Col Y) y 178 filas (177 respuestas encuestadas).
  - Registrar rango de base de datos formal: `doc.DatabaseRanges.addNewByName("Tabla_Encuesta_Base", cellRangeAddress)` con `ContainsHeader = True` y `refresh()`.
- **Hojas de Filtro Estándar (Copia de datos con DatabaseRange propio):**
  - `de 0 a 2 años` (Punto 4): Filtro `Antigüedad laboral: = 'De 0 a 2 años'` (58 registros copiados + encabezado). `DatabaseRange: Filtro_0_a_2_anos`.
  - `Más de 10 años` (Punto 5): Filtro `Antigüedad laboral: = 'Más de 10 años'` (49 registros copiados + encabezado). `DatabaseRange: Filtro_Mas_de_10_anos`.
  - `Más de 10 años y Linux` (Punto 5.a): Filtro compuesto con operador AND: `Antigüedad = 'Más de 10 años'` Y `SO = 'Linux'` (7 registros copiados + encabezado). `DatabaseRange: Filtro_Mas_de_10_anos_Linux`.
- **Cálculos Estadísticos con Fórmulas en Inglés (Puntos 7 y 8):**
  - `Punto 7` (Capacitación al momento del ingreso): `=COUNTIF('Respuestas de formulario 1'.G2:G178; "Sí")` -> 72 respondentes (41,14%).
  - `Punto 8` (Capacitación virtual previa): `=COUNTIF('Respuestas de formulario 1'.H2:H178; "Si")` -> 105 respondentes (60,00%).
  - `Punto 6 complementario` (Capacitación presencial previa): `=COUNTIF('Respuestas de formulario 1'.I2:I178; "Si")` -> 133 respondentes (76,00%).
- **Análisis Metodológico (Punto 9):**
  - Dictamen: **NO** estuvo bien formulada. La pregunta cualitativa abierta en un instrumento tabulado impide el cálculo con `=CONTAR.SI` o tablas dinámicas sin reclasificación manual; debió formularse como pregunta cerrada con escala tipo Likert o categorías predeterminadas.
- **Tablas Dinámicas y Gráficos Dinámicos con Solapas Desplegables (Puntos 10 y 11):**
  - `Punto 10` (Modalidades de Capacitación): Tabla Dinámica nativa `TD_Modalidades` en hoja `Tabla Dinámica - Modalidades` (Fila y Datos = Col Q) + Gráfico Dinámico vertical (`X = 14000`, `Y = 1500`) con `loext:data-pilot-source="TD_Modalidades"` inyectado quirúrgicamente en su XML (`Modalidad ▼`, `Datos ▼`).
    - Resultados: Virtual (68 - 38,86%), Presencial en puesto (46 - 26,29%), Semipresencial (40 - 22,86%), Presencial fuera (21 - 12,00%).
  - `Punto 11` (Cruce Software vs Modalidad): Tabla Dinámica nativa de contingencia `TD_Software_Modalidad` en hoja `Tabla Dinámica - Cruce Software` (Filas = Software Col U, Columnas = Modalidad Col Q, Datos = Recuento) + Gráfico Dinámico vertical (`X = 1500`, `Y = 14500`) con `loext:data-pilot-source="TD_Software_Modalidad"` (`Software ▼`, `Modalidad ▼`, `Datos ▼`).
    - Totales: Moodle (50), LibreOffice Writer (38), LibreOffice Calc (21), MediaWiki (24), Ninguna de los anteriores (42).
- **Diagnóstico y Propuesta Institucional (Punto 12):**
  - Sistemas Operativos: Windows 77,14% (135), Linux 10,86% (19), Otro 12,00% (21).
  - Disponibilidad: 84% cuenta con entre 2 y 5 horas semanales.
  - Propuesta: Plan formativo modular Moodle + Suite LibreOffice (Writer y Calc) bajo modalidad semipresencial / híbrida flexible de 3 horas semanales durante 6 semanas.
