# Proveedor: Nutremax

## Fórmula de precio (Tipo 1 — actualización de precios)
PrecioConIVA = Precio (Hoja1) + 18%

Ejemplo: Precio en Hoja1 = 1000 → PrecioConIVA = 1000 * 1.18 = 1180.00

## Reglas específicas de este proveedor
Ninguna adicional a las generales del `CLAUDE.md`. Aplica el flujo estándar de Tipo 1:
- Match exacto de cada código individual de `Codigo` (Hoja1) contra los elementos de `CodigoFactura` en todas las hojas (excepto `Marcas Actualización Dinámica`).
- Escribe `PrecioConIVA` (valor fijo, 2 decimales) y `FechaActualizacion` (texto `dd/mm/aaaa`) en la fila donde se halló el código.
- Códigos de Hoja1 sin match → listado en columna E de Hoja1.
- Registra cada celda modificada en `RegistroCambios`.

## Parámetros para el script (uso interno — no modificar el formato de esta sección)
tipo: porcentaje_simple
porcentaje: 0.18
