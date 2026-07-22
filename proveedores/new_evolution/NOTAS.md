# Proveedor: New Evolution

## Fórmula de precio (Tipo 1 — actualización de precios)
PrecioConIVA = Precio (Hoja1) + 23% + 21% (compuesto, aplicado en cadena, no equivalente a sumar 44% directo)

Ejemplo: Precio en Hoja1 = 1000 → 1000 * 1.23 = 1230.00 → 1230.00 * 1.21 = 1488.30

## Reglas específicas de este proveedor
Ninguna adicional a las generales del `CLAUDE.md`. Aplica el flujo estándar de Tipo 1:
- Match exacto de cada código individual de `Codigo` (Hoja1) contra los elementos de `CodigoFactura` en todas las hojas (excepto `Marcas Actualización Dinámica`).
- Escribe `PrecioConIVA` (valor fijo, 2 decimales) y `FechaActualizacion` (texto `dd/mm/aaaa`) en la fila donde se halló el código.
- Códigos de Hoja1 sin match → listado en columna E de Hoja1.
- Registra cada celda modificada en `RegistroCambios`.

## Parámetros para el script (uso interno — no modificar el formato de esta sección)
tipo: porcentaje_compuesto
pasos: 0.23, 0.21
