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
- **Marcas Actualización Dinámica** (EXCLUIR SIEMPRE — no leer, no escribir)

## Detalle por hoja

### Hoja1 (staging manual)
- Encabezados (fila 1): `CODIGO`, `PRECIO` (columnas C, D, E de la fila de header están vacías/`None`)
- Filas con datos: 0 (actualmente vacía, a la espera de que el usuario pegue la lista del proveedor)
- `max_row` reportado por Excel: 2078 (rango usado históricamente, pero sin datos actuales)

### Hojas destino por marca/categoría
Las 11 hojas de destino (Bicicletas MTB, Bicicletas Ruta, Bicicletas Gravel, Bicicletas Paseo, Bicicletas BMX-Infantiles, Indumentaria, Computadoras, Cuadros, Nutricional, Accesorios, Repuestos) comparten **exactamente el mismo esquema de encabezados** (fila 1):

`FechaActualizacion`, `CodigoDux`, `NombreProducto`, `CodigoFactura`, `Marca`, `PrecioContado`, `PrecioSinIVA`, `PrecioConIVA`, `DolarReferencia`

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

### Marcas Actualización Dinámica (EXCLUIR SIEMPRE)
- Encabezados no siguen el esquema estándar: `Marca`, `Fecha Más Antigua`, `Necesita Actualización`, (columna vacía), `PRODUCTOS DESACTUALIZADOS (más de 30 días) - Lista Dinámica`, y varias columnas adicionales con `ArrayFormula` (fórmulas matriciales) y sub-encabezados (`Producto`, `Marca`, `Fecha Última Actualización`, `Días Sin Actualizar`).
- `max_row`: 3000. No se debe leer ni escribir en esta hoja bajo ninguna circunstancia, según CLAUDE.md.

### RegistroCambios (trazabilidad)
- Encabezados (formato único, 10 columnas): `Fecha`, `TipoOperacion`, `Hoja`, `Fila`, `NombreProducto`, `Marca`, `Columna`, `ValorAnterior`, `ValorNuevo`, `Detalle`
- Filas con datos: 0. La hoja se rehízo el 21/08/2026 con `scripts/migrar_registro.py`, que descartó las 1403 filas de historial previo.
- `max_row`: 1 (solo la fila de encabezados)
- El formato lo define `scripts/registro.py` y lo comparten los tres scripts (Tipo 1, 2 y 3). Antes de escribir, cada uno valida que los encabezados coincidan con esos 10 y aborta sin guardar si no.
- Granularidad: una fila por cada celda modificada. Tipo 1 escribe 2 filas por producto (`PrecioConIVA` + `FechaActualizacion`), Tipo 2 escribe 3 (suma `DolarReferencia`) y Tipo 3 escribe 1.

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

**Hoja1 (staging):**
- `CODIGO`
- `PRECIO`

**RegistroCambios (formato único, definido en `scripts/registro.py`):**
- `Fecha`
- `TipoOperacion`
- `Hoja`
- `Fila`
- `NombreProducto`
- `Marca`
- `Columna`
- `ValorAnterior`
- `ValorNuevo`
- `Detalle`

**Nota:** El CLAUDE.md del proyecto menciona `CodigoFac`/`CodigoFactura` y `DolarRef`/`DolarReferencia` como nombres alternativos posibles. En el archivo real, el nombre efectivo en las 11 hojas destino es siempre `CodigoFactura` y `DolarReferencia` (no se encontraron variantes `CodigoFac` ni `DolarRef` en ninguna hoja).
