# Proyecto: Lista de Precios Giro Vivo — Reglas generales

Este archivo aplica a **todos los proveedores** y a **ambos tipos de prompt** (actualización de precios y asignación de dólar). Las reglas específicas de cada proveedor viven en `proveedores/<proveedor>/NOTAS.md`.

## Archivo de trabajo
- Libro único y persistente: `archivo/Lista de Precios Giro Vivo 2026.xlsx`.
- `Hoja1`: hoja de entrada, donde el usuario pega manualmente la lista del proveedor antes de cada corrida.
- Resto de las hojas: hojas de destino por marca/categoría (columnas `CodigoFac`/`CodigoFactura`, `Marca`, `PrecioConIVA`, `DolarRef`/`DolarReferencia`, `FechaActualizacion`, etc.).
- El número de hojas y filas **no es fijo**: recorrer siempre todas las hojas dinámicamente, nunca asumir una lista cerrada.
- **Hoja excluida siempre:** `Marcas Actualización Dinámica`. Ningún script de actualización la lee ni la modifica. Única excepción de lectura: el control de capacidad de `verificar_formulas.py` (solo lectura). Solo se modifica con autorización expresa del usuario para esa tarea (ver sección propia más abajo).
- Hoja de trazabilidad: `RegistroCambios` (se crea si no existe, antes de escribir cualquier cambio).

## Reglas obligatorias (aplican siempre)
1. **Nunca modificar el archivo real sin backup previo con timestamp** en `backup/`.
2. Columnas siempre localizadas **por nombre de encabezado exacto**, nunca por posición ni letra de columna.
3. Abrir el libro con `openpyxl.load_workbook()` **sin** `data_only=True`, para no destruir fórmulas preexistentes del libro al guardar.
4. Valores escritos siempre **fijos** (no fórmulas):
   - Precios: número (float, 2 decimales máximo).
   - Fechas: texto en formato `dd/mm/aaaa`.
5. Matching de códigos: split por coma + strip + comparación de **igualdad exacta normalizada** (cuidado con tipos int/float/str que pueden diferir entre columnas — nunca comparar por substring o coincidencia parcial).
6. Excluir siempre la hoja `Marcas Actualización Dinámica` (salvo la lectura del control de capacidad de `verificar_formulas.py`; modificarla requiere autorización expresa).
7. Recorrer todas las hojas del libro **dinámicamente** en cada corrida — pueden agregarse categorías nuevas con el tiempo.
8. Todo cambio se registra en `RegistroCambios`, con **una fila por cada celda modificada** (si se tocan 2 columnas en la misma fila, son 2 registros). El total de filas de `RegistroCambios` debe coincidir exactamente con la cantidad de celdas modificadas.
   - El formato es **único y compartido** por los tres scripts, y vive en `scripts/registro.py`: `Fecha | TipoOperacion | Hoja | Fila | NombreProducto | Marca | Columna | ValorAnterior | ValorNuevo | Detalle`. Ningún script define encabezados propios.
   - Antes de escribir, cada script valida que los encabezados de la hoja coincidan con esos 10. Si no coinciden, aborta sin guardar en vez de apilar filas desalineadas.
9. Flujo obligatorio en **3 fases**:
   1. Inspección (sin escribir nada).
   2. Dry-run / simulación con resumen (sin escribir nada).
   3. Ejecución real — solo tras aprobación explícita del usuario.
10. **Toda columna que se agregue a una hoja de destino tiene que quedar dentro de la tabla de esa hoja** (ver "Columnas dentro de las tablas").
11. **Nunca borrar archivos de `backup/` con comodines o patrones**; solo por nombre exacto y mostrando antes la lista de archivos a borrar.

## Columna ActualizadoSegun y hoja NoActualizados
- `ActualizadoSegun` (última columna de cada hoja de destino, con lista desplegable) guarda la **última** forma en que se actualizó el precio: `Lista proveedor` (la escribe Tipo 1), `Dólar` (la escribe Tipo 4) o `Manual` (la carga el usuario a mano). Tipo 2 y Tipo 3 no la tocan. Se escribe y registra solo si el valor cambia.
- La columna se creó una sola vez con `scripts/crear_columna_actualizado.py`. Lógica compartida en `scripts/actualizado_segun.py`.
- Cada `NOTAS.md` declara en `## Parámetros` la línea `marcas:` (valores de la columna `Marca` que trae ese proveedor, separados por coma).
- `NoActualizados`: se regenera en cada ejecución real de Tipo 1 con las filas del Excel de esas marcas que no recibieron precio, con su motivo. Nunca incluye códigos de Hoja1. Es hoja de informe (como la columna E de Hoja1): no se registra en `RegistroCambios`, tiene los encabezados en la fila 4 y todos los scripts la excluyen por nombre.

## Columnas dentro de las tablas
- **Toda columna que se agregue a una hoja de destino tiene que quedar dentro de la tabla (`Table_N`) de esa hoja.** Si queda afuera, al ordenar en Excel sus valores no se mueven con las filas, y no toma el estilo del encabezado (en Bicicletas MTB queda blanco sobre blanco).
- Las columnas se agregan siempre al final (nunca en el medio) y la tabla se amplía editando su XML con `scripts/tablas_xml.py` (no con openpyxl). Cada tabla ampliada se registra como `Estructura` (`Columna` = `Tabla <nombre>`, ref anterior → nueva).
- `scripts/ampliar_tablas.py` amplía las tablas que hayan quedado angostas y quita los filtros guardados (flujo de 3 fases, con verificación byte a byte de Marcas y apertura en Excel). Desde el 23/11/2026 todas las tablas incluyen `ActualizadoSegun` (y `Categoria` donde existe).

## Columna Categoria
- Columna de texto libre (sin lista desplegable por ahora) que se crea y completa con `scripts/asignar_categoria.py --filtro "<fragmento de hoja>" --valor "<Categoría>"` (dry-run por defecto, que arma y verifica el resultado en una carpeta temporal; `--execute` para escribir; `--sobrescribir` para pisar otra categoría; `--archivo` para probar sobre una copia).
- Se agrega en la primera columna libre al final y el mismo script amplía la tabla de la hoja para incluirla. Si ya está dentro y completa, no cambia nada. Fila de producto = `CodigoDux` no vacío.
- Registro: encabezado nuevo → `Estructura`; cada valor → `Categoria` (una fila por celda); cada tabla ampliada → `Estructura`. Conteo: filas de registro = celdas modificadas + tablas ampliadas.
- Aplicada el 23/11/2026 a las 11 hojas de destino (columna K, dentro de cada tabla): las 5 de bicicletas con `Bicicletas` y cada una de las otras 6 con su propio nombre (`--hoja "<Hoja>" --valor "<Hoja>"`). Para tomar una sola hoja usar `--hoja` (nombre exacto), no `--filtro` (coincidencia parcial).

## Fórmulas obligatorias
- `PrecioContado` = `='<Hoja>'!$H6-('<Hoja>'!$H6*15/100)` y `PrecioSinIVA` = `='<Hoja>'!$H6/1.21`, donde la letra sale del encabezado `PrecioConIVA`.
- Tipo 1 y Tipo 4 reponen estas fórmulas en las filas que actualizan (vía `formulas.py`) y lo registran en `RegistroCambios`.
- `verificar_formulas.py` revisa y repara todo el libro (y redondea `PrecioConIVA` a 2 decimales).

## Hoja Marcas Actualización Dinámica
- Se alimenta sola de las tablas `Table_1` … `Table_11` con 6 fórmulas matriciales de **rango fijo** (openpyxl no conserva las matrices dinámicas de Excel y fija el rango al guardar).
- Desde el 22/11/2026 (`scripts/ampliar_marcas.py`) las 6 están envueltas en `EXPAND(TAKE(…))`, de modo que rellenan con "" en lugar de dar `#N/A` y no pierden productos mientras no se supere la capacidad:
  - J1:J5000, K1:K5000, L1:L5000, M1:M5000 → **5.000** filas de tablas.
  - A2:C201 → **200** marcas.
  - E3:H3000 → **2.998** productos desactualizados.
- `verificar_formulas.py` hace en cada corrida un control **de solo lectura**: avisa "ampliar el rango de la hoja Marcas" si alguna medida supera el 90 % de su capacidad, o si alguna de las 6 fórmulas perdió `EXPAND(TAKE(`. No escribe en el libro ni en `RegistroCambios`.
- Ampliar o modificar la hoja requiere **autorización expresa** del usuario para esa tarea, se hace editando solo el XML de la hoja (nunca con openpyxl) y con el flujo de 3 fases.
- El autofiltro `A1:C41` de la hoja está desactualizado; se dejó como está a propósito (su nombre definido vive en `workbook.xml`).

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
    ├── actualizado_segun.py    (columna ActualizadoSegun y hoja NoActualizados, compartido)
    ├── crear_columna_actualizado.py (uso único: crea la columna ActualizadoSegun)
    ├── asignar_categoria.py    (crea/completa la columna Categoria en las hojas filtradas y amplía su tabla; registra "Estructura" y "Categoria")
    ├── ampliar_tablas.py       (amplía las tablas hasta el último encabezado y quita filtros guardados, editando el XML)
    ├── tablas_xml.py           (ampliación de tablas por XML + verificaciones openpyxl / byte a byte / Excel, compartido)
    ├── formulas.py             (verificación de fórmulas Contado/SinIVA, compartido)
    ├── dolar.py                (regla única de validación del valor de dólar, compartido)
    ├── verificar_formulas.py   (revisión y reparación de fórmulas en todo el libro + control de capacidad de Marcas, solo lectura)
    ├── ampliar_marcas.py       (amplía las 6 fórmulas de Marcas editando solo su XML; requiere autorización expresa)
    ├── normalizar_nombres.py   (normaliza NombreProducto de una hoja a "<PREFIJO> <MARCA> <MODELO>": --hoja y --prefijo; genera revision_nombres_<hoja>.xlsx para revisar y aplica solo Aplicar=SI, con MarcaPropuesta opcional; registra TipoOperacion "Nombre")
    ├── eliminar_filas_sin_fecha.py (elimina filas completas sin FechaActualizacion editando el XML; registra TipoOperacion "Eliminacion")
    ├── mover_fila.py         (mueve una fila de producto de una hoja a otra por NombreProducto + CodigoDux, editando el XML: copia por encabezado, repone fórmulas, amplía/achica tablas; registra "Movimiento", "Estructura" y "Eliminacion"; dry-run por defecto, --execute para escribir)
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
