# Proveedor: Venzo

## Fórmula de precio (Tipo 1 — actualización de precios)
PrecioConIVA = Precio (Hoja1) − 7,01% − 6% − 6,03% + 81% (compuesto, aplicado en cadena: cada paso se aplica sobre el resultado del anterior, no se suman los porcentajes)

Equivale a: Precio × (1 − 0,0701) × (1 − 0,06) × (1 − 0,0603) × 1,81

El precio de Hoja1 ya incluye IVA, así que el resultado va directo a `PrecioConIVA`.

Ejemplo: Precio en Hoja1 = 242465
→ 242465 × 0,9299 = 225468,20
→ × 0,94 = 211940,11
→ × 0,9397 = 199160,12
→ × 1,81 = 360479,82 (PrecioConIVA)

Los porcentajes son fijos para Venzo y aplican solo a este proveedor.

## Reglas específicas de este proveedor
Ninguna adicional a las generales del `CLAUDE.md`. Aplica el flujo estándar de Tipo 1:
- Match exacto de cada código individual de `Codigo` (Hoja1) contra los elementos de `CodigoFactura` en todas las hojas (excepto `Marcas Actualización Dinámica`).
- Escribe `PrecioConIVA` (valor fijo, 2 decimales) y `FechaActualizacion` (texto `dd/mm/aaaa`) en la fila donde se halló el código.
- Códigos de Hoja1 sin match → listado en columna E de Hoja1.
- Registra cada celda modificada en `RegistroCambios`.

## Parámetros para el script (uso interno — no modificar el formato de esta sección)
tipo: porcentaje_compuesto
pasos: -0.0701, -0.06, -0.0603, 0.81
