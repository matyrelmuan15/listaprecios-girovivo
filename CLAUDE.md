# Proyecto: Lista de Precios Giro Vivo — Reglas generales

Este archivo aplica a **todos los proveedores** y a **ambos tipos de prompt** (actualización de precios y asignación de dólar). Las reglas específicas de cada proveedor viven en `proveedores/<proveedor>/NOTAS.md`.

## Archivo de trabajo
- Libro único y persistente: `archivo/Lista de Precios Giro Vivo 2026.xlsx`.
- `Hoja1`: hoja de entrada, donde el usuario pega manualmente la lista del proveedor antes de cada corrida.
- Resto de las hojas: hojas de destino por marca/categoría (columnas `CodigoFac`/`CodigoFactura`, `Marca`, `PrecioConIVA`, `DolarRef`/`DolarReferencia`, `FechaActualizacion`, etc.).
- El número de hojas y filas **no es fijo**: recorrer siempre todas las hojas dinámicamente, nunca asumir una lista cerrada.
- **Hoja excluida siempre:** `Marcas Actualización Dinámica`. No se lee ni se modifica en ningún prompt.
- Hoja de trazabilidad: `RegistroCambios` (se crea si no existe, antes de escribir cualquier cambio).

## Reglas obligatorias (aplican siempre)
1. **Nunca modificar el archivo real sin backup previo con timestamp** en `backup/`.
2. Columnas siempre localizadas **por nombre de encabezado exacto**, nunca por posición ni letra de columna.
3. Abrir el libro con `openpyxl.load_workbook()` **sin** `data_only=True`, para no destruir fórmulas preexistentes del libro al guardar.
4. Valores escritos siempre **fijos** (no fórmulas):
   - Precios: número (float, 2 decimales máximo).
   - Fechas: texto en formato `dd/mm/aaaa`.
5. Matching de códigos: split por coma + strip + comparación de **igualdad exacta normalizada** (cuidado con tipos int/float/str que pueden diferir entre columnas — nunca comparar por substring o coincidencia parcial).
6. Excluir siempre la hoja `Marcas Actualización Dinámica`.
7. Recorrer todas las hojas del libro **dinámicamente** en cada corrida — pueden agregarse categorías nuevas con el tiempo.
8. Todo cambio se registra en `RegistroCambios`, con **una fila por cada celda modificada** (si se tocan 2 columnas en la misma fila, son 2 registros). El total de filas de `RegistroCambios` debe coincidir exactamente con la cantidad de celdas modificadas.
9. Flujo obligatorio en **3 fases**:
   1. Inspección (sin escribir nada).
   2. Dry-run / simulación con resumen (sin escribir nada).
   3. Ejecución real — solo tras aprobación explícita del usuario.

## Decisión de arquitectura
No agrupar proveedores por fórmula compartida, aunque hoy coincida el porcentaje entre dos o más. Cada proveedor tiene su propia carpeta y `NOTAS.md` independiente, porque las fórmulas pueden cambiar de forma desacoplada con el tiempo.

## Estructura de carpetas
```
ListaPreciosGiroVivo/
├── CLAUDE.md
├── archivo/
│   └── Lista de Precios Giro Vivo 2026.xlsx
├── backup/
├── proveedores/
│   └── <proveedor>/
│       └── NOTAS.md
└── scripts/
    ├── actualizar_precios.py
    └── asignar_dolar.py   (pendiente — se agrega en la etapa de dólar)
```

## Estilo de comunicación
Formal y profesional, en español rioplatense (voseo), sin vulgarismos.
