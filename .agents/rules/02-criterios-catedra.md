# Criterios Docentes, Buenas Prácticas y Pautas de Evaluación de Cátedra
# File: .agents/rules/02-criterios-catedra.md

Las docentes de la cátedra aplican las siguientes pautas determinantes para la corrección y aprobación de trabajos prácticos y parciales:

---

# 1. Uso Obligatorio de LibreOffice Calc (Software Libre)
- **Prohibición de Microsoft Excel:** La cátedra evalúa y corrige exclusivamente bajo el entorno y sintaxis de **LibreOffice Calc**. Resolver o guardar desde Excel invalida la entrega.
- **Incompatibilidad de Fórmulas:** Funciones lógicas y condicionales presentan diferencias de nomenclatura y sintaxis entre Excel y Calc. Si una solución falla en Calc, se considera reprobada independientemente de cómo funcione en Excel.
- **Apertura de Archivos:** Clic derecho sobre el `.ods` > *Abrir con > LibreOffice Calc* para evitar aperturas por defecto en Excel.

---

# 2. Organización y Gestión de Hojas (Pestañas)
- **Copiar antes de Modificar (No Mover):** Al duplicar una hoja base (ej. `Personas`, `Encuestas`), en el cuadro *Mover o copiar hoja...*, **marcar obligatoriamente la casilla "Copiar" y no "Mover"**, conservando intacta la hoja original con sus datos fuente.
- **Nombre Exacto de Pestañas:** Respetar de forma rigurosa la nomenclatura solicitada (ej. `1A`, `1B`, `2A`, `2B`, `2C`, `3A`, `3B`, `de 0 a 2 años`, `P7`, `P7-a`). Los docentes corrigen navegando directamente a la pestaña específica; resolver en la hoja equivocada o con nombre distorsionado se califica negativamente.
- **Contexto de Hoja Activa:** Estar siempre ubicado en la pestaña correspondiente antes de aplicar filtros, ordenamientos o fórmulas.

---

# 3. Fórmulas, Sintaxis y Asistente de Funciones
- **El Signo Igual (`=`):** Es de uso obligatorio al inicio de cualquier celda de cálculo para que LibreOffice Calc la interprete como expresión evaluable y no como texto plano.
- **Referencias Absolutas vs. Relativas (`$`):**
  - **Parámetros fijos externos:** Cuando se vincula un valor fijo ubicado en una celda externa (montos, porcentajes, tasas de descuento, incentivos), **debe fijarse obligatoriamente con el signo pesos** (ej. `$C$1`, `$C$2`) utilizando la tecla `F4`.
  - **Celdas variables por fila:** Las referencias a datos del registro individual (horas, legajo, escalafón) deben conservarse relativas (sin `$`) para que se adapten automáticamente al arrastrarse hacia abajo.
- **Cadenas de Texto entre Comillas Dobles (`" "`):** Todo texto dentro de una condición o argumento funcional (como `"F"`, `"M"`, `"Nivel 1"`, `"Sí"`, `"No"`, `"Personal"`) debe escribirse obligatoriamente entre comillas dobles. Su omisión genera inmediatamente los errores `#¿NOMBRE?` o `#¡VALOR!`.
- **Sensibilidad a Acentos y Tildes:** En funciones de búsqueda y conteo condicional (`CONTAR.SI`, `CONTAR.SI.CONJUNTO`), el criterio debe coincidir de forma idéntica con el texto original en los datos: `"Sí"` con tilde es tratado como un valor distinto de `"Si"` sin tilde.
- **Separadores Sintácticos (`:` vs. `;`):**
  - Los **dos puntos (`:`)** definen un rango continuo de celdas (ej. `A4:A108`, `G2:G178`).
  - El **punto y coma (`;`)** se utiliza obligatoriamente para separar argumentos funcionales o celdas individuales disjuntas.
- **Asistente para Funciones (`fx`):** Se recomienda utilizar el asistente (`Insertar > Función...` o `Ctrl + F2`) cuando existan dudas sobre la estructura de la función, verificando en la vista previa del diálogo que el resultado sea válido antes de presionar *Aceptar*.

---

# 4. Selección para Ordenar y Filtrar
- **Posicionarse dentro de la Tabla:** Antes de ejecutar `Datos > Ordenar...` o `Datos > Más filtros > Filtro estándar...`, el cursor debe posicionarse obligatoriamente sobre cualquier celda con datos de la tabla o en su fila de encabezados (nunca en celdas vacías o periféricas). Esto garantiza que Calc reconozca el rango de datos completo y conserve los encabezados de columna.
- **Filtro Automático vs. Filtro Estándar:**
  - **Filtro Automático:** Recomendado para selecciones rápidas o valores puntuales discretos.
  - **Filtro Estándar:** De uso obligatorio cuando la consigna solicita rangos numéricos continuos (ej. escalafón entre 6 y 10) o condiciones combinadas complejas con operadores lógicos (`Y` / `O`).

---

# 5. Tablas Dinámicas y Gráficos Estadísticos
- **Exclusión Estricta de Totales Generales:** Al seleccionar datos de una tabla o tabla dinámica para confeccionar un gráfico, **nunca se debe incluir la fila o columna de 'Total general' o resultados acumulados**. Incluir los totales duplica visualmente la escala y arruina la proporcionalidad de las barras o sectores.
- **Elementos Obligatorios del Gráfico:** Un gráfico debe contar obligatoriamente con:
  1. Título descriptivo claro y temático.
  2. Nombres explícitos en los ejes (eje X horizontal y eje Y vertical).
  3. Etiquetas de datos numéricas sobre barras o sectores, formateadas con los decimales adecuados (generalmente 2 decimales para promedios o porcentajes).

---

# 6. Hábitos Operativos y Atajos Clave
- **Guardar Cambios Frecuentemente (`Ctrl + G`):** Guardar con periodicidad regular para proteger el progreso.
- **Navegación Rápida (`Ctrl + Flecha abajo`):** Desplazarse al instante hasta la última fila con datos sin tener que hacer scroll manual a lo largo de tablas extensas.
- **Autocompletado de Fórmulas:** Hacer doble clic en la esquina inferior derecha de la celda activa (cuadro de llenado) para autocompletar fórmulas hacia toda la columna de manera instantánea.
- **Resolución de Columnas Estrechas (`###`):** Si una columna visualiza numerales (`###`), no constituye un error de cálculo, sino ancho insuficiente de la celda; se resuelve haciendo doble clic en la línea divisoria del encabezado de columna o seleccionando *Ancho óptimo*.
