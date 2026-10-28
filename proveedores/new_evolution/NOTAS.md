# Proveedor: New Evolution

## Fórmula de precio (Tipo 1 — actualización de precios)
PrecioConIVA = Precio (Hoja1) + 23% + 21% (compuesto, aplicado en cadena, no equivalente a sumar 44% directo)

Ejemplo: Precio en Hoja1 = 1000 → 1000 * 1.23 = 1230.00 → 1230.00 * 1.21 = 1488.30

## Reglas específicas de este proveedor
**Limpieza de espacios en códigos** (`limpiar_espacios_codigo: si`): antes del matching, se eliminan todos los espacios de cada código de la columna `Codigo` de Hoja1 (extremos, del medio y no separables `\xa0`). Las comas se conservan, así que los códigos múltiples se siguen separando por coma.
- Dry-run: no escribe nada, pero el matching ya usa los códigos limpios y el resumen muestra cuántos códigos tenían espacios, con ejemplos (antes → después).
- Ejecución real: escribe en Hoja1 los códigos limpios como texto y registra cada celda modificada en `RegistroCambios` (Hoja `Hoja1`, Columna `Codigo`), una fila por celda.
- Los valores numéricos o sin espacios no se tocan. Las celdas que solo tienen espacios tampoco.

Además, aplica el flujo estándar de Tipo 1:
- Match exacto de cada código individual de `Codigo` (Hoja1) contra los elementos de `CodigoFactura` en todas las hojas (excepto `Marcas Actualización Dinámica`).
- Escribe `PrecioConIVA` (valor fijo, 2 decimales) y `FechaActualizacion` (texto `dd/mm/aaaa`) en la fila donde se halló el código.
- Códigos de Hoja1 sin match → listado en columna E de Hoja1.
- Registra cada celda modificada en `RegistroCambios`.

## Parámetros para el script (uso interno — no modificar el formato de esta sección)
tipo: porcentaje_compuesto
pasos: 0.23, 0.21
limpiar_espacios_codigo: si
