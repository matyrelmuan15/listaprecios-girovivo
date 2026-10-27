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
   - El formato es **único y compartido** por los tres scripts, y vive en `scripts/registro.py`: `Fecha | TipoOperacion | Hoja | Fila | NombreProducto | Marca | Columna | ValorAnterior | ValorNuevo | Detalle`. Ningún script define encabezados propios.
   - Antes de escribir, cada script valida que los encabezados de la hoja coincidan con esos 10. Si no coinciden, aborta sin guardar en vez de apilar filas desalineadas.
9. Flujo obligatorio en **3 fases**:
   1. Inspección (sin escribir nada).
   2. Dry-run / simulación con resumen (sin escribir nada).
   3. Ejecución real — solo tras aprobación explícita del usuario.

## Fórmulas obligatorias
- `PrecioContado` = `='<Hoja>'!$H6-('<Hoja>'!$H6*15/100)` y `PrecioSinIVA` = `='<Hoja>'!$H6/1.21`, donde la letra sale del encabezado `PrecioConIVA`.
- Tipo 1 y Tipo 4 reponen estas fórmulas en las filas que actualizan (vía `formulas.py`) y lo registran en `RegistroCambios`.
- `verificar_formulas.py` revisa y repara todo el libro (y redondea `PrecioConIVA` a 2 decimales).

## Decisión de arquitectura
No agrupar proveedores por fórmula compartida, aunque hoy coincida el porcentaje entre dos o más. Cada proveedor tiene su propia carpeta y `NOTAS.md` independiente, porque las fórmulas pueden cambiar de forma desacoplada con el tiempo.

## Estructura de carpetas

Raíz real y única del proyecto:
`C:\Users\Maty y Katy\Desktop\ListaPreciosGiroVivo`

```
ListaPreciosGiroVivo/
├── CLAUDE.md
├── ESTRUCTURA_EXCEL.md
├── requirements.txt
├── abrir_lista_de_precios.bat
├── crear_acceso_directo.ps1
├── archivo/
│   └── Lista de Precios Giro Vivo 2026.xlsx
├── backup/
├── proveedores/
│   └── <proveedor>/
│       └── NOTAS.md
└── scripts/
    ├── actualizar_precios.py   (Tipo 1 — precios por match de códigos; con --dolar opcional también escribe DolarReferencia en las filas actualizadas)
    ├── asignar_dolar.py        (Tipo 2 — asigna DolarReferencia + fecha, NO recalcula precios)
    ├── actualizar_precio_dolar.py (Tipo 4 — recalcula PrecioConIVA = precio ÷ DolarReferencia actual × dólar nuevo)
    ├── actualizar_fecha.py     (Tipo 3 — solo FechaActualizacion)
    ├── registro.py             (formato único de RegistroCambios, compartido)
    ├── formulas.py             (verificación de fórmulas Contado/SinIVA, compartido)
    ├── dolar.py                (regla única de validación del valor de dólar, compartido)
    ├── verificar_formulas.py   (revisión y reparación de fórmulas en todo el libro)
    ├── migrar_registro.py      (migración de un solo uso de RegistroCambios)
    ├── ia_revision.py          (revisión automática de anomalías)
    ├── sincronizar_git.py      (sincronización con GitHub, usada por la interfaz)
    └── interfaz.py             (interfaz gráfica de escritorio)
```

## Sincronización con GitHub
La interfaz sincroniza el proyecto con `origin/main` (vía `scripts/sincronizar_git.py`) para trabajar en dos PC sin comandos manuales.
- **Baja** cambios (`git fetch` + `git pull --ff-only`) al abrir el programa, antes de cada vista previa y antes de cada "Confirmar y aplicar". Si en la confirmación llegaron cambios nuevos, se cancela y se exige una nueva vista previa. Si no hay conexión, se pregunta si continuar.
- **Sube** cambios (`git add -A` + commit + `git push`) solo después de una ejecución real terminada con código 0, y con el botón "Sincronizar ahora". Mensaje: `Auto: <operación> – <detalle> – dd/mm/aaaa HH:MM – PC <COMPUTERNAME>`. También sube commits pendientes de corridas anteriores sin conexión.
- **Nunca** `--force`, merge, rebase, stash ni reset automáticos. Si las dos PC tienen cambios distintos, el estado es "conflicto" y se resuelve a mano.
- Cada comando git: timeout de 60 s, `GIT_TERMINAL_PROMPT=0`, `GCM_INTERACTIVE=never`, `CREATE_NO_WINDOW` (sin ventanas de consola ni pedidos de contraseña).
- `.git/index.lock` se borra solo si tiene más de 2 minutos y no hay ningún `git.exe` corriendo; si no, se aborta.
- Si git no está instalado o la carpeta no es un repositorio, la sincronización se desactiva y el programa sigue funcionando.
- La sincronización **no** escribe en `RegistroCambios` (no son cambios de celdas).
- `scripts/config_interfaz.json` no se versiona: son preferencias de cada PC.
- Cada PC necesita la credencial de GitHub guardada (Git Credential Manager) y `user.name` / `user.email` configurados.

## Estilo de comunicación
Formal y profesional, en español rioplatense (voseo), sin vulgarismos.
