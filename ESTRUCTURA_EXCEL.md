# Estructura del archivo `archivo/Lista de Precios Giro Vivo 2026.xlsx`

> Relevamiento en modo solo lectura (`openpyxl.load_workbook(..., data_only=False)`, sin guardar el libro). Fecha del relevamiento: 13/08/2026.

## Hojas del libro

- **Hoja1** (staging manual — entrada donde el usuario pega la lista del proveedor)
- Bicicletas MTB
- Bicicletas Ruta
- Bicicletas Gravel
- Bicicletas Paseo
- Bicicletas BMX-Infantiles
- Indumentaria
- Computadoras
- Cuadros
- Nutricional
- Accesorios
- Repuestos
- RegistroCambios (hoja de trazabilidad)
- NoActualizados (hoja de informe que regenera Tipo 1; ver sección propia)
- **Marcas Actualización Dinámica** (EXCLUIR SIEMPRE — solo la lee el control de capacidad de `verificar_formulas.py`; se modifica solo con autorización expresa)

## Detalle por hoja

### Hoja1 (staging manual)
- Encabezados (fila 1): `CODIGO`, `PRECIO` (columnas C, D, E de la fila de header están vacías/`None`)
- Filas con datos: 0 (actualmente vacía, a la espera de que el usuario pegue la lista del proveedor)
- `max_row` reportado por Excel: 2078 (rango usado históricamente, pero sin datos actuales)

### Hojas destino por marca/categoría
Las 11 hojas de destino (Bicicletas MTB, Bicicletas Ruta, Bicicletas Gravel, Bicicletas Paseo, Bicicletas BMX-Infantiles, Indumentaria, Computadoras, Cuadros, Nutricional, Accesorios, Repuestos) comparten **exactamente el mismo esquema de encabezados** (fila 1):

`FechaActualizacion`, `CodigoDux`, `NombreProducto`, `CodigoFactura`, `Marca`, `PrecioContado`, `PrecioSinIVA`, `PrecioConIVA`, `DolarReferencia`, `ActualizadoSegun`

y, solo en las hojas donde ya se asignó categoría, `Categoria` a continuación (ver más abajo).

- Las tablas de Excel `Table_1` … `Table_11` cubren solo de `FechaActualizacion` a `DolarReferencia` (A:I). `ActualizadoSegun` y `Categoria` quedan **fuera** de las tablas, a propósito: las tablas alimentan las fórmulas de `Marcas Actualización Dinámica` y no se amplían.
- `ActualizadoSegun` (columna J): lista desplegable `Manual` / `Dólar` / `Lista proveedor` en J2:J5000. Creada una sola vez con `scripts/crear_columna_actualizado.py`.
- `Categoria` (columna K): texto libre, sin lista desplegable por ahora. Se crea y se completa con `scripts/asignar_categoria.py` (ver sección propia).

Cantidad aproximada de filas con datos por hoja (conteo de filas con al menos una celda no vacía, a partir de la fila 2):

| Hoja | Filas con datos (aprox.) | max_row (rango usado) |
|---|---|---|
| Bicicletas MTB | 116 | 998 |
| Bicicletas Ruta | 38 | 1008 |
| Bicicletas Gravel | 10 | 998 |
| Bicicletas Paseo | 12 | 1000 |
| Bicicletas BMX-Infantiles | 43 | 1003 |
| Indumentaria | 184 | 1000 |
| Computadoras | 14 | 1000 |
| Cuadros | 12 | 1000 |
| Nutricional | 49 | 1000 |
| Accesorios | 644 | 986 |
| Repuestos | 309 | 985 |

**Observaciones importantes:**
- Algunas hojas usan **fórmulas** en `PrecioContado` y/o `PrecioSinIVA` (ej.: `='Bicicletas Gravel'!$H2-('Bicicletas Gravel'!$H2*15/100)`), referenciadas contra la columna `PrecioConIVA` (columna H) de la misma hoja. Esto se ve en Bicicletas Gravel, Bicicletas Paseo, Bicicletas BMX-Infantiles, Cuadros, Nutricional. Confirma la regla del proyecto de **no abrir con `data_only=True`** para no perder estas fórmulas al guardar.
- En `Bicicletas MTB` y `Bicicletas Ruta`, en cambio, `PrecioContado` y `PrecioSinIVA` aparecen como valores fijos (no fórmula) en las filas de muestra revisadas.
- `CodigoFactura` puede contener **múltiples códigos separados por coma** en una misma celda (ej.: `'1096025104,1096025105,1096025106,...'`), consistente con la regla de matching por split + strip + igualdad exacta normalizada.
- `CodigoFactura` puede venir vacío/`None` en algunas filas (ej.: fila 2 y 3 de `Cuadros`).
- `CodigoDux` mezcla tipos: valores numéricos (`597`, `715`) y strings (`'SPYSNIPER3'`, `'ZI-BRRFNR56-028'`) según la fila — confirma la advertencia de la regla 5 del proyecto sobre cuidado con tipos int/float/str al comparar códigos.
- `FechaActualizacion` mezcla texto en formato `dd/mm/aaaa` (mayoría de los casos) con al menos un valor tipo `datetime.datetime` sin convertir a texto (visto en `Bicicletas BMX-Infantiles`, fila 2: `datetime.datetime(2026, 7, 30, 0, 0)`). Esto contradice parcialmente la regla 4 del proyecto ("fechas siempre como texto dd/mm/aaaa") para ese registro puntual — dato existente en el archivo, no introducido por este relevamiento.
- En `Accesorios`, las primeras filas relevadas corresponden a lo que parecen ser **encabezados de subcategoría** (ej. `'ACEITES'`, `'ASIENTOS'` en la columna `NombreProducto`, con el resto de columnas vacías o en 0) — a tener en cuenta al filtrar filas de producto real vs. filas de agrupación visual.
- `NombreProducto` en `Nutricional` presenta problemas de codificación de caracteres especiales preexistentes en el archivo (ej. `amino�cidos` en vez de `aminoácidos`) — no es un artefacto de este relevamiento.

### Columna Categoria (`scripts/asignar_categoria.py`)
- Aplicada el 23/11/2026 a las 5 hojas de bicicletas (`--filtro "bicicleta" --valor "Bicicletas"`): columna **K** en Bicicletas MTB (128 filas), Ruta (41), Gravel (11), Paseo (12) y BMX-Infantiles (46). En total, 5 encabezados + 238 valores = 243 celdas y 243 filas de `RegistroCambios`. Backup previo: `backup/Lista de Precios Giro Vivo 2026_20261123_113302.xlsx`.
- Las demás hojas todavía no tienen la columna.
- Uso (desde `scripts/`):
  - `python asignar_categoria.py --filtro "<fragmento>" --valor "<Categoría>"` → inspección + dry-run (no escribe).
  - Mismo comando con `--execute` → backup + escritura.
  - `--filtro` se puede repetir y compara contra el nombre de la hoja sin importar mayúsculas ni acentos. `--sobrescribir` reemplaza celdas que ya tienen **otra** categoría (por defecto no se tocan y se informan).
- Si la hoja no tiene `Categoria`, se crea en la primera columna libre a la derecha del último encabezado, copiando el formato y el ancho del encabezado vecino. Si ya existe, se usa esa.
- Fila de producto = fila con `CodigoDux` no vacío. Las celdas que ya tienen el valor pedido se saltean, así que correrlo dos veces no cambia nada.
- Siempre excluye Hoja1, Marcas Actualización Dinámica, RegistroCambios, NoActualizados y Venzo Sin Actualizar, y solo procesa hojas con `CodigoDux` y `PrecioConIVA` en la fila 1.

### NoActualizados (hoja de informe)
- La regenera cada ejecución real de Tipo 1 (`actualizar_precios.py`) con las filas de las marcas del proveedor (línea `marcas:` de su `NOTAS.md`) que no recibieron precio, y su motivo.
- Encabezados en la **fila 4**: `Hoja`, `Fila`, `CodigoDux`, `NombreProducto`, `CodigoFactura`, `Marca`, `PrecioConIVA`, `FechaActualizacion`, `ActualizadoSegun`, `Motivo`.
- No se registra en `RegistroCambios` y todos los scripts la excluyen por nombre.

### Marcas Actualización Dinámica (EXCLUIR SIEMPRE)
- Encabezados no siguen el esquema estándar: `Marca`, `Fecha Más Antigua`, `Necesita Actualización`, (columna vacía), `PRODUCTOS DESACTUALIZADOS (más de 30 días) - Lista Dinámica`, y varias columnas adicionales con `ArrayFormula` (fórmulas matriciales) y sub-encabezados (`Producto`, `Marca`, `Fecha Última Actualización`, `Días Sin Actualizar`).
- `max_row`: 5000 (dimensión A1:M5000 desde el 22/11/2026).
- 6 fórmulas matriciales de rango fijo (`t="array"`), ampliadas el 22/11/2026 con `scripts/ampliar_marcas.py` y envueltas en `EXPAND(TAKE(…))` para que rellenen con "" en lugar de dar `#N/A`:

  | Celda | Rango | Contenido | Capacidad |
  |---|---|---|---|
  | J1 | J1:J5000 | `EXPAND(TAKE(VSTACK(Table_1[NombreProducto],…,Table_11[NombreProducto]),5000),5000,,"")` | 5.000 filas de tablas |
  | K1 | K1:K5000 | ídem sobre `[Marca]` | 5.000 |
  | L1 | L1:L5000 | ídem sobre `[FechaActualizacion]` | 5.000 |
  | M1 | M1:M5000 | `EXPAND(TAKE(MAP(ANCHORARRAY(L1),LAMBDA(…fecha de texto a número…)),5000),5000,,"")` | 5.000 |
  | A2 | A2:C201 | `EXPAND(TAKE(LET(…marca, fecha más antigua, "Sí"/"No"…),200),200,3,"")` | 200 marcas |
  | E3 | E3:H3000 | `EXPAND(TAKE(LET(…productos con más de 30 días…),2998),2998,4,"")` | 2.998 productos |

  El orden de las tablas en `VSTACK` es `Table_1` … `Table_11` (Repuestos al final). Las demás celdas de los rangos son celdas de derrame de esas matrices (en el XML, `<f ca="1"/>` o `<f />` según quién guardó el libro).
- Al guardar con openpyxl se conservan el texto y el rango de las 6 fórmulas; se pierden los valores cacheados y el atributo `ca`, y Excel recalcula al abrir (verificado sin errores).
- `verificar_formulas.py` controla en cada corrida (solo lectura) que las filas de las tablas, las marcas con fecha y los productos desactualizados no superen el 90 % de la capacidad, y que las 6 fórmulas mantengan `EXPAND(TAKE(`.
- Otros elementos: formato condicional en `A2:C600` (`$C2="Sí"`) y `E3:H3000` (`$E3<>""`, `$H3>90`); celda combinada `E1:H1`; panel inmovilizado en la fila 2; columnas J:M agrupadas (nivel 1).
- El autofiltro `A1:C41` (y su nombre definido `_xlnm._FilterDatabase` en `workbook.xml`) está **desactualizado**: no cubre todo el rango A2:C201. Se dejó así a propósito.
- No se escribe en esta hoja salvo con autorización expresa del usuario para esa tarea, editando solo el XML de la hoja (nunca con openpyxl).

### RegistroCambios (trazabilidad)
- Encabezados (formato único, 10 columnas): `Fecha`, `TipoOperacion`, `Hoja`, `Fila`, `NombreProducto`, `Marca`, `Columna`, `ValorAnterior`, `ValorNuevo`, `Detalle`
- Filas con datos: 0. La hoja se rehízo el 21/08/2026 con `scripts/migrar_registro.py`, que descartó las 1403 filas de historial previo.
- `max_row`: 1 (solo la fila de encabezados)
- El formato lo define `scripts/registro.py` y lo comparten los tres scripts (Tipo 1, 2 y 3). Antes de escribir, cada uno valida que los encabezados coincidan con esos 10 y aborta sin guardar si no.
- Granularidad: una fila por cada celda modificada. Tipo 1 escribe 2 filas por producto (`PrecioConIVA` + `FechaActualizacion`), o 3 si se indica el dólar opcional (`--dolar`) y la hoja tiene `DolarReferencia` (ese valor se escribe en las mismas filas actualizadas; las hojas sin la columna se informan con un aviso), Tipo 2 (asignar dólar) escribe 2 (`DolarReferencia` + `FechaActualizacion`), Tipo 3 escribe 1 y Tipo 4 (actualizar precio por dólar) escribe 3 (`PrecioConIVA` + `DolarReferencia` + `FechaActualizacion`). Si se elige no modificar la fecha, Tipo 2 y Tipo 4 escriben una fila menos.

**Historial previo (descartado del libro, preservado en `backup/`):** hasta la migración la hoja tenía los encabezados `Hoja`, `Tabla/Producto (SKU)`, `FechaActualizacion`, `Valor anterior DolarReferencia`, `Valor nuevo`, pero convivían debajo tres formatos de distinto ancho (4-6, 7 y 9 columnas) más un bloque de resumen manual, porque cada script definía sus propios encabezados y ninguno los reescribía. El contenido completo quedó en `backup/Lista de Precios Giro Vivo 2026_20260821_114842.xlsx`.

## Columnas relevantes (nombres exactos de encabezado)

**Hojas destino (11 hojas, esquema idéntico):**
- `FechaActualizacion`
- `CodigoDux`
- `NombreProducto`
- `CodigoFactura`
- `Marca`
- `PrecioContado`
- `PrecioSinIVA`
- `PrecioConIVA`
- `DolarReferencia`
- `ActualizadoSegun`
- `Categoria` (por ahora solo en las 5 hojas de bicicletas)

**Hoja1 (staging):**
- `CODIGO`
- `PRECIO`

**RegistroCambios (formato único, definido en `scripts/registro.py`):**
- `Fecha`
- `TipoOperacion` (`Precios`, `Dolar`, `PrecioDolar`, `Fecha`, `Formulas`, `Estructura`, `Eliminacion`, `Categoria`; `Eliminacion` lo usa `scripts/eliminar_filas_sin_fecha.py`, con una fila por fila eliminada y `Columna` = `Fila completa`; `Estructura` registra encabezados nuevos, con `Columna` = letra; `Categoria` lo usa `scripts/asignar_categoria.py`, una fila por celda, con `Columna` = `Categoria`)
- `Hoja`
- `Fila`
- `NombreProducto`
- `Marca`
- `Columna`
- `ValorAnterior`
- `ValorNuevo`
- `Detalle`

**Nota:** El CLAUDE.md del proyecto menciona `CodigoFac`/`CodigoFactura` y `DolarRef`/`DolarReferencia` como nombres alternativos posibles. En el archivo real, el nombre efectivo en las 11 hojas destino es siempre `CodigoFactura` y `DolarReferencia` (no se encontraron variantes `CodigoFac` ni `DolarRef` en ninguna hoja).
