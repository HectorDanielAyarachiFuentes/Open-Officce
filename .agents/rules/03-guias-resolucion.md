# Guías de Referencia Rápida por Trabajo Práctico y Exámenes
# File: .agents/rules/03-guias-resolucion.md

---

# 1. Práctica 01 (Planilla de Cálculo)
- **Hoja Inicio:** Copias de celdas y fórmulas, series numéricas y de fechas en rangos continuos, inserción y borrado de filas/columnas.
- **Hoja Punto3:** Totalizadores con CTRL + Flecha abajo:
  - Cantidad de filas: `=COUNTA(...)`.
  - Sueldo anual total: `=SUM(...)`.
  - Promedio categoría sin decimales: `=AVERAGE(...)` con `NumberFormat = 1`.
  - Cantidad de mujeres: `=COUNTIF(E4:E103; "F")`.
- **Hoja Punto4:** Columnas calculadas:
  - `MNOMBRE`: `=UPPER(...)`.
  - `Año Nac`: `=YEAR(...)`.
  - `Ascenso`: categoría + 2.
  - `Aumento`: salario * nuevo código / 2.
  - `Deducciones`: 10% del aumento.
  - `Nivel`: `=IF(categoria>3; "Nivel 2"; "Nivel 1")`.

---

# 2. Práctica 02 (Planilla de Cálculo)
- **Hoja personas:**
  - `Sueldo` (Col L): `=K{row}*$C$1` (referencia absoluta `$C$1` para monto fijo de `$1875`).
  - `Descuento` (Col M): `=L{row}*$C$2` (referencia absoluta `$C$2` para `22,00%`).
  - `Incentivo` (Col N): `=IF(K{row}>7; 1700; 0)`.
- **Hoja encuestas:**
  - `Calculo` (Col N): `=IF(MOD(A{row}; 2)=0; E{row}*2; E{row})`.
- **Hojas 1A y 1B (Ordenamiento múltiple con TableSortField):**
  - **Requisito crítico:** Usar siempre `uno.createUnoStruct("com.sun.star.table.TableSortField")` con `.IsAscending = True` (prohibido `com.sun.star.util.SortField` o `.SortAscending`, que fallan silenciosamente).
  - `1A`: Ordenar `A4:N108` por Localidad (Field 8) y Dirección (Field 6) con `uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2))`.
  - `1B`: Ordenar `A1:N225` por Curso (Field 2), Sexo (Field 3) y Vivienda (Field 9) con `uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2, sf3))`.
- **Hojas 2A, 2B y 2C (Filtros estándar y ordenamiento):**
  - `2A`: Contaduría con escalafón 6 a 10.
    - **Ordenamiento dual obligatorio:** Ordenar el rango de datos `A4:N109` antes o después de filtrar usando `com.sun.star.table.TableSortField`:
      - Clave 1: Apellido (`Field = 1`, `IsAscending = True`)
      - Clave 2: Nombres (`Field = 2`, `IsAscending = True`)
      - Envoltura: `uno.Any("[]com.sun.star.table.TableSortField", (sf1, sf2))`.
    - **Preservación de filas 1 a 4:** Ejecutar siempre `for r in range(4): sheet.getRows().getByIndex(r).setPropertyValue("IsVisible", True)` para evitar que el filtro colapse los encabezados y parámetros externos (`$C$1`, `$C$2`).
    - **Cursor:** Posicionar en celda interna de datos o encabezado (ej. `B4` o `C4`).
  - `2B`: Contaduría con escalafón 6. Restaurar visibilidad de filas 1 a 4 tras filtrar.
  - `2C`: Vivienda PROPIA y Electricidad SI. Restaurar visibilidad de filas 1 a 4 tras filtrar.
- **Hojas 3A y 3B (Tablas Dinámicas y Pivot Charts con solapitas):**
  - `3A`: Tabla Dinámica de Vivienda (Alquila: 507, Propia: 611) + Gráfico Dinámico vertical (`X = 38000`, `Y = 5000`) con `loext:data-pilot-source="TablaDinamica_3A"` (`Vivienda ▼`, `Datos ▼`).
  - `3B`: Tabla Dinámica de Sueldo Promedio por Sección + Gráfico Dinámico vertical (`X = 11000`, `Y = 1500`) con `loext:data-pilot-source="TablaDinamica_3B"` (`Sección ▼`, `Datos ▼`).

---

# 3. Práctica 03 (Planilla de Cálculo)
- **Hoja Respuestas de formulario 1:**
  - `Punto 3`: Eliminar columna A (`Timestamp`), dejando 25 columnas de datos (A a Y) y 178 filas.
  - `DatabaseRange`: `Tabla_Encuesta_Base` con `ContainsHeader = True` y `refresh()`.
- **Hojas de Filtro Estándar (Copia de datos con DatabaseRange propio):**
  - `de 0 a 2 años` (Punto 4): `Antigüedad laboral: = 'De 0 a 2 años'` (58 registros + encabezado). `DatabaseRange: Filtro_0_a_2_anos`.
  - `Más de 10 años` (Punto 5): `Antigüedad laboral: = 'Más de 10 años'` (49 registros + encabezado). `DatabaseRange: Filtro_Mas_de_10_anos`.
  - `Más de 10 años y Linux` (Punto 5.a): `Antigüedad = 'Más de 10 años'` Y `SO = 'Linux'` (7 registros + encabezado). `DatabaseRange: Filtro_Mas_de_10_anos_Linux`.
- **Cálculos Estadísticos con Fórmulas en Inglés (Puntos 7 y 8):**
  - `Punto 7` (Capacitación al ingreso): `=COUNTIF('Respuestas de formulario 1'.G2:G178; "Sí")` -> 72 (41,14%).
  - `Punto 8` (Capacitación virtual previa): `=COUNTIF('Respuestas de formulario 1'.H2:H178; "Si")` -> 105 (60,00%).
  - `Punto 6 comp` (Capacitación presencial previa): `=COUNTIF('Respuestas de formulario 1'.I2:I178; "Si")` -> 133 (76,00%).
- **Tablas Dinámicas y Pivot Charts con Solapas Desplegables (Puntos 10 y 11):**
  - `Punto 10`: TD nativa `TD_Modalidades` + Gráfico Dinámico vertical (`X = 14000`, `Y = 1500`) con `loext:data-pilot-source="TD_Modalidades"` (`Modalidad ▼`, `Datos ▼`).
  - `Punto 11`: TD nativa de contingencia `TD_Software_Modalidad` + Gráfico Dinámico vertical (`X = 1500`, `Y = 14500`) con `loext:data-pilot-source="TD_Software_Modalidad"` (`Software ▼`, `Modalidad ▼`, `Datos ▼`).

---

# 4. Final 2024 (Planilla de Cálculo)
- **Hoja encuesta:**
  - `edad` (Col F): `=YEAR(TODAY())-E{row}` formateada como entero.
  - `par-categoría` (Col L): `=ISEVEN(H{row})` (retorna `VERDADERO` / `FALSO`).
  - `Mayúscula` (Col M): `=UPPER(C{row})` (apellido en mayúsculas).
  - `Plus por cat` (Col N): `=IF(H{row}>=3; "Si corresponde"; " ")`.
  - `Tabla Dinámica (TD_Regimen en P3:Q8)`: Personal (51), Funcionarios (49), Total (100).
  - `Gráfico Dinámico (Pivot Chart)`: Ubicado en Col P a V (`X = 36000`, `Y = 5500`), vinculado a `TD_Regimen` con `loext:data-pilot-source="TD_Regimen"` (`Régimen ▼`, `Datos ▼`).
- **Hoja P7 (Filtro Automático):**
  - Duplicado exacto de `encuesta` registrado con `DatabaseRange: Filtro_P7` y `AutoFilter = True`.
- **Hojas P7-a, P7-b, P7-c (Filtros estándar con copia de datos y DatabaseRange formal):**
  - `P7-a`: `región = 'AN'` Y `régimen = 'Personal'` (9 registros + encabezado). `Datos_P7_a`.
  - `P7-b`: `edad < 30` Y `régimen = 'Funcionarios'` (6 registros + encabezado). `Datos_P7_b`.
  - `P7-c`: `régimen = 'Funcionarios'` Y `región = 'CL'` (3 registros + encabezado). `Datos_P7_c`.
- **Exportación CSV (Punto 10 - Looker Studio):**
  - `final-TIC2024.csv` exportado con 14 columnas completas y 101 filas en codificación UTF-8 limpia.

---

# 5. Final (Talentia HR — Transformación Digital y Capacitación)
- **Hoja prueba (Base de datos enriquecida):**
  - `apellido` (Col C): Texto normalizado a letras **MAYÚSCULAS** (`transform_column_text(sheet, 2)`).
  - `edad` (Col F): `=YEAR(TODAY())-E{row}` formateada como entero numérico.
  - `tipo_categoria` (Col L): `=IF(MOD(H{row}; 2)=0; "PAR"; "IMPAR")` (previene `Err:504`).
  - `modalidad_capacitacion` (Col M): `=IF(H{row}>=3; "Presencial / Taller Intensivo"; "Virtual Asincrónica")`.
  - `prioridad_formacion` (Col N): `=IF(H{row}>=3; "Alta Prioridad"; "Prioridad Estándar")`.
  - `DatabaseRange`: `Tabla_Nomina` (`A1:N101`) con `AutoFilter = True`.
- **Hojas de Filtros Independientes (`copy_filtered_sheet`):**
  - `Filtro a`: Funcionarios y región `CV` (4 registros + encabezado). `DatabaseRange: Filtro_a`.
  - `Filtro b`: Personal y región `CM` o `A` (2 registros + encabezado). `DatabaseRange: Filtro_b`.
- **Hoja Resumen:**
  - Tablas con `=COUNTIFS(...)` para Funcionarios sobre modalidad (20 presencial, 29 virtual) y Personal sobre prioridad (19 alta, 32 estándar).
  - Detalle curricular por nivel y DataPilot nativo `TD_Propuesta_Formativa` con Pivot Chart dinámico (`loext:data-pilot-source`).
  - Ubicación de entrega: carpeta `Final/Solucion/final.ods` y `Solucion/final.ods`.
