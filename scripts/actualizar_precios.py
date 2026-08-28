"""
actualizar_precios.py
======================
Script genérico Tipo 1: actualización de precios por match de códigos.

MODO GUIADO (recomendado)
--------------------------
Si lo corrés sin argumentos:

    python actualizar_precios.py

el script te muestra un menú para elegir el proveedor con las flechas
↑ / ↓ y Enter, corre automáticamente la vista previa (dry-run) y al final
te pregunta, también con un menú de flechas, si confirmás aplicar los
cambios reales al archivo. No hace falta recordar nombres de proveedor
ni flags.

MODO POR LÍNEA DE COMANDOS (para automatizar o repetir una corrida exacta)
----------------------------------------------------------------------------
    python actualizar_precios.py <proveedor> [--execute]

Por defecto corre en modo DRY-RUN (no escribe nada, solo muestra el resumen
de lo que haría). Para ejecutar los cambios reales, agregar --execute.

Ejemplo:
    python actualizar_precios.py venzo
    python actualizar_precios.py venzo --execute

Reglas generales aplicadas: ver CLAUDE.md en la raíz del proyecto.
"""

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

import registro
from ia_revision import (
    formatear_items_marcados,
    imprimir_items_marcados_json,
    revisar_cambios_con_ia,
)

# ---------------------------------------------------------------------------
# Configuración de rutas (relativas a la raíz del proyecto)
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"
PROVEEDORES_DIR = PROJECT_ROOT / "proveedores"

HOJA_EXCLUIDA = "Marcas Actualización Dinámica"
HOJA_ENTRADA = "Hoja1"
HOJA_REGISTRO = registro.HOJA_REGISTRO

# Formato argentino válido: dígitos, agrupados opcionalmente de a 3 con punto
# de miles, y coma decimal opcional (ej: "2.607,75", "607,75", "2.607").
ARG_PRICE_PATTERN = re.compile(r"^\d{1,3}(\.\d{3})*(,\d+)?$")

# El formato de RegistroCambios es único y compartido: vive en registro.py.
# Este script no define encabezados propios.


# ---------------------------------------------------------------------------
# Menú interactivo por consola (flechas + Enter) — solo Windows
# ---------------------------------------------------------------------------
def listar_proveedores():
    """Lee dinámicamente las subcarpetas de proveedores/ que tengan NOTAS.md."""
    if not PROVEEDORES_DIR.exists():
        return []
    proveedores = []
    for carpeta in sorted(PROVEEDORES_DIR.iterdir(), key=lambda p: p.name.lower()):
        if carpeta.is_dir() and (carpeta / "NOTAS.md").exists():
            proveedores.append(carpeta.name)
    return proveedores


def menu_seleccionar(opciones, titulo="Seleccioná una opción"):
    """Menú de consola navegable con flechas ↑ / ↓ y Enter (requiere Windows).

    Devuelve el índice elegido, o None si se cancela con Esc.
    """
    try:
        import msvcrt
    except ImportError:
        # Si no estamos en Windows, se cae a un menú por número (respaldo).
        return _menu_seleccionar_numerico(opciones, titulo)

    seleccion = 0

    def dibujar():
        print("\n" + titulo)
        print("=" * len(titulo))
        for i, op in enumerate(opciones):
            marca = "  > " if i == seleccion else "    "
            print(f"{marca}{op}")
        print("\n(Flechas arriba/abajo para moverte, Enter para confirmar, Esc para cancelar)")

    dibujar()
    while True:
        tecla = msvcrt.getch()
        if tecla in (b"\xe0", b"\x00"):  # prefijo de tecla especial en Windows
            tecla2 = msvcrt.getch()
            if tecla2 == b"H":  # flecha arriba
                seleccion = (seleccion - 1) % len(opciones)
                dibujar()
            elif tecla2 == b"P":  # flecha abajo
                seleccion = (seleccion + 1) % len(opciones)
                dibujar()
        elif tecla == b"\r":  # Enter
            return seleccion
        elif tecla == b"\x1b":  # Esc
            return None


def _menu_seleccionar_numerico(opciones, titulo):
    """Respaldo simple (por si el script se corre fuera de Windows/consola real)."""
    print("\n" + titulo)
    print("=" * len(titulo))
    for i, op in enumerate(opciones, start=1):
        print(f"  {i}) {op}")
    while True:
        valor = input("Elegí un número: ").strip()
        if valor.isdigit() and 1 <= int(valor) <= len(opciones):
            return int(valor) - 1
        print("  Opción inválida, intentá de nuevo.")


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def normalize_code(value):
    """Normaliza un código para poder compararlo sin importar si viene
    como int, float o str (ej: 123, 123.0, '123', ' 123 ' deben ser iguales)."""
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    try:
        f = float(s)
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s


def find_header_col(ws, header_name, max_search_rows=3):
    """Busca en las primeras filas de la hoja una celda cuyo texto
    (normalizado) coincida exactamente con header_name. Devuelve el índice
    de columna (1-based) o None si no se encuentra."""
    target = header_name.strip().casefold()
    for row in ws.iter_rows(min_row=1, max_row=max_search_rows):
        for cell in row:
            if cell.value is not None and str(cell.value).strip().casefold() == target:
                return cell.column
    return None


def parse_notas_params(notas_path: Path) -> dict:
    """Extrae los parámetros de la sección '## Parámetros' del NOTAS.md."""
    if not notas_path.exists():
        raise FileNotFoundError(f"No se encontró NOTAS.md en: {notas_path}")

    lines = notas_path.read_text(encoding="utf-8").splitlines()
    params = {}
    in_params_section = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("## Parámetros"):
            in_params_section = True
            continue
        if in_params_section:
            if stripped.startswith("##"):
                break
            if ":" in stripped and stripped:
                key, _, value = stripped.partition(":")
                params[key.strip()] = value.strip()

    if not params:
        raise ValueError(
            f"No se encontraron parámetros en la sección '## Parámetros' de {notas_path}"
        )
    return params


def parse_precio(value) -> float:
    """Convierte el valor de precio de Hoja1 a float, aceptando tanto
    números planos (1234.56) como texto en formato argentino
    (punto de miles, coma decimal: '2.607,75').

    Si el valor no es un número plano y tampoco matchea el patrón de
    formato argentino válido, no se asume ninguna conversión: se levanta
    ValueError para que el llamador lo trate como formato no reconocido."""
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        raise ValueError(f"precio vacío: {value!r}")
    try:
        return float(s)
    except ValueError:
        pass
    if not ARG_PRICE_PATTERN.match(s):
        raise ValueError(f"formato de precio no reconocido: {value!r}")
    s_arg = s.replace(".", "").replace(",", ".")
    return float(s_arg)


def apply_formula(precio: float, params: dict) -> float:
    """Aplica la fórmula de precio del proveedor según los parámetros de su NOTAS.md."""
    tipo = params.get("tipo")

    if tipo == "porcentaje_simple":
        porcentaje = float(params["porcentaje"])
        resultado = precio * (1 + porcentaje)

    elif tipo == "porcentaje_compuesto":
        # pasos: "0.70, 0.21"  -> aplica cada porcentaje en cadena, no sumados
        pasos = [float(p.strip()) for p in params["pasos"].split(",")]
        resultado = precio
        for paso in pasos:
            resultado = resultado * (1 + paso)

    else:
        raise ValueError(f"Tipo de fórmula desconocido en NOTAS.md: '{tipo}'")

    return round(resultado, 2)


def backup_archivo():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{ARCHIVO_PATH.stem}_{timestamp}{ARCHIVO_PATH.suffix}"
    shutil.copy2(ARCHIVO_PATH, destino)
    return destino


# ---------------------------------------------------------------------------
# Lógica principal
# ---------------------------------------------------------------------------
def procesar(proveedor: str, execute: bool):
    notas_path = PROVEEDORES_DIR / proveedor / "NOTAS.md"
    params = parse_notas_params(notas_path)

    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")
        sys.exit(1)

    wb = openpyxl.load_workbook(ARCHIVO_PATH)  # sin data_only=True

    # Validación fail-fast del formato de RegistroCambios: se corre apenas
    # cargado el libro y ANTES del backup, para no dejar un backup inútil ni
    # apilar filas desalineadas bajo encabezados que no corresponden.
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)

    if HOJA_ENTRADA not in wb.sheetnames:
        print(f"ERROR: no se encontró la hoja de entrada '{HOJA_ENTRADA}'.")
        sys.exit(1)

    ws_hoja1 = wb[HOJA_ENTRADA]
    col_codigo = find_header_col(ws_hoja1, "Codigo")
    col_precio = find_header_col(ws_hoja1, "Precio")

    if col_codigo is None or col_precio is None:
        print("ERROR: no se encontraron las columnas 'Codigo' y/o 'Precio' en Hoja1.")
        sys.exit(1)

    # Construir diccionario código normalizado -> precio original (de Hoja1)
    codigos_hoja1 = {}  # normalized_code -> precio
    for row in ws_hoja1.iter_rows(min_row=2):
        celda_codigo = row[col_codigo - 1]
        celda_precio = row[col_precio - 1]
        if celda_codigo.value is None:
            continue
        norm = normalize_code(celda_codigo.value)
        if norm == "":
            continue
        codigos_hoja1[norm] = celda_precio.value

    # Pre-parseo de precios de Hoja1: separa los que tienen formato
    # reconocido (número plano o argentino válido) de los que no, para no
    # convertir a ciegas ni perderlos en silencio.
    precios_parsed = {}  # norm -> float
    avisos_formato = []  # lista de (norm, valor_original)
    for norm, valor_original in codigos_hoja1.items():
        try:
            precios_parsed[norm] = parse_precio(valor_original)
        except (TypeError, ValueError):
            avisos_formato.append((norm, valor_original))

    codigos_no_encontrados = set(precios_parsed.keys())
    detalle_corrida = (
        f"Proveedor: {proveedor} | Fórmula: "
        + ", ".join(f"{k}={v}" for k, v in params.items())
    )
    corrida = registro.RegistroCorrida(registro.TIPO_PRECIOS, detalle_corrida)
    dry_run_detalle = []  # lista de dicts con detalle original -> parseado (solo dry-run)
    cambios_precio = []  # lista para la revisión automática de anomalías (IA)
    resumen_por_hoja = {}  # hoja -> {"precios": n, "fechas": n}
    total_coincidencias = 0

    fecha_hoy = datetime.now().strftime("%d/%m/%Y")

    for sheet_name in wb.sheetnames:
        if sheet_name in (HOJA_ENTRADA, HOJA_EXCLUIDA, HOJA_REGISTRO):
            continue

        ws = wb[sheet_name]
        col_codfac = find_header_col(ws, "CodigoFactura") or find_header_col(ws, "CodigoFac")
        col_marca = find_header_col(ws, "Marca")
        col_precio_iva = find_header_col(ws, "PrecioConIVA")
        col_fecha = find_header_col(ws, "FechaActualizacion")
        col_nombre = find_header_col(ws, "NombreProducto")

        if col_codfac is None or col_precio_iva is None or col_fecha is None:
            # Esta hoja no es una hoja de destino válida para este prompt
            continue

        precios_mod = 0
        fechas_mod = 0

        for row in ws.iter_rows(min_row=2):
            celda_codfac = row[col_codfac - 1]
            if celda_codfac.value is None:
                continue

            elementos = [normalize_code(e) for e in str(celda_codfac.value).split(",")]
            elementos = [e for e in elementos if e]

            match = next((e for e in elementos if e in precios_parsed), None)
            if match is None:
                continue

            precio_original = codigos_hoja1[match]
            precio_parseado = precios_parsed[match]
            precio_nuevo = apply_formula(precio_parseado, params)

            fila_num = celda_codfac.row
            marca_valor = row[col_marca - 1].value if col_marca else None
            nombre_valor = row[col_nombre - 1].value if col_nombre else None

            celda_precio_iva = row[col_precio_iva - 1]
            celda_fecha = row[col_fecha - 1]

            valor_anterior_precio = celda_precio_iva.value
            valor_anterior_fecha = celda_fecha.value

            cambios_precio.append({
                "hoja": sheet_name,
                "codigo": match,
                "producto": nombre_valor,
                "valor_anterior": valor_anterior_precio,
                "valor_nuevo": precio_nuevo,
            })

            # Una fila de registro por cada celda modificada (regla 8):
            # PrecioConIVA y FechaActualizacion.
            corrida.agregar(
                hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                marca=marca_valor, columna="PrecioConIVA",
                valor_anterior=valor_anterior_precio, valor_nuevo=precio_nuevo,
            )
            corrida.agregar(
                hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                marca=marca_valor, columna="FechaActualizacion",
                valor_anterior=valor_anterior_fecha, valor_nuevo=fecha_hoy,
            )

            if not execute:
                dry_run_detalle.append({
                    "Hoja": sheet_name, "Fila": fila_num, "Codigo": match,
                    "Original": precio_original, "Parseado": precio_parseado,
                })

            if execute:
                celda_precio_iva.value = precio_nuevo
                celda_fecha.value = fecha_hoy

            precios_mod += 1
            fechas_mod += 1
            total_coincidencias += 1
            codigos_no_encontrados.discard(match)

        if precios_mod > 0:
            resumen_por_hoja[sheet_name] = {"precios": precios_mod, "fechas": fechas_mod}

    # ---- Columna E de Hoja1: códigos sin match ----
    no_encontrados_ordenados = sorted(codigos_no_encontrados)
    if execute:
        # Limpiar columna E existente (a partir de la fila 2) antes de reescribir
        max_row_actual = ws_hoja1.max_row
        for r in range(2, max_row_actual + 1):
            ws_hoja1.cell(row=r, column=5).value = None
        for i, codigo in enumerate(no_encontrados_ordenados, start=2):
            ws_hoja1.cell(row=i, column=5).value = codigo

    # ---- RegistroCambios ----
    # Verificación de conteo obligatoria (regla 8): 2 celdas por match
    # (PrecioConIVA + FechaActualizacion). Se corre antes de escribir, y si
    # falla se aborta sin guardar.
    celdas_modificadas = total_coincidencias * 2
    try:
        corrida.verificar_conteo(celdas_modificadas)
        if execute:
            corrida.escribir(wb)
    except registro.RegistroError as e:
        print(f"\n{e}")
        sys.exit(1)

    # ---------------------------------------------------------------
    # Resumen
    # ---------------------------------------------------------------
    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Proveedor: {proveedor} | Modo: {modo} ===\n")
    print(f"Fórmula aplicada: {params}")
    print(f"Total de coincidencias encontradas: {total_coincidencias}")
    print(f"Total de precios modificados: {total_coincidencias}")
    print(f"Total de fechas modificadas: {total_coincidencias}")
    print("\nDesglose por hoja:")
    if resumen_por_hoja:
        for hoja, datos in resumen_por_hoja.items():
            print(f"  - {hoja}: {datos['precios']} precios, {datos['fechas']} fechas")
    else:
        print("  (sin coincidencias en ninguna hoja)")
    print(f"\nCódigos de Hoja1 sin encontrar: {len(no_encontrados_ordenados)}")
    if no_encontrados_ordenados:
        print(f"  {no_encontrados_ordenados}")

    avisos_formato_ordenados = sorted(avisos_formato, key=lambda item: item[0])
    print(f"\nAVISO: formato de precio no reconocido: {len(avisos_formato_ordenados)}")
    if avisos_formato_ordenados:
        for norm, valor_original in avisos_formato_ordenados:
            print(f"  - Código '{norm}': valor original {valor_original!r}")

    # ---- Revisión automática de posibles anomalías (estadística + IA) ----
    print("\n--- Revisión automática de posibles anomalías ---")
    analisis_ia, aviso_ia = revisar_cambios_con_ia(
        cambios_precio,
        contexto_extra=(
            f"Actualización de precios (Tipo 1) para el proveedor '{proveedor}'. "
            f"Fórmula: {params}."
        ),
    )
    if aviso_ia:
        print(aviso_ia)
    if analisis_ia:
        print(analisis_ia)
    detalle_marcados = formatear_items_marcados(cambios_precio)
    if detalle_marcados:
        print(detalle_marcados)
    imprimir_items_marcados_json(cambios_precio)

    if not execute:
        registro.imprimir_previsualizacion(corrida)

    if not execute and dry_run_detalle:
        print("\nDetalle de filas a modificar (Original -> Parseado):")
        for item in dry_run_detalle:
            print(
                f"  - [{item['Hoja']}] fila {item['Fila']} código '{item['Codigo']}': "
                f"Original: {item['Original']!r} -> Parseado: {item['Parseado']}"
            )

    if execute:
        backup_path = backup_archivo()
        print(f"\nBackup creado en: {backup_path}")
        wb.save(ARCHIVO_PATH)
        print(f"Archivo guardado: {ARCHIVO_PATH}")
        print(f"Filas agregadas a {HOJA_REGISTRO}: {len(corrida)}")
    else:
        print(
            "\nEsto fue una simulación (dry-run). No se modificó el archivo ni se "
            "creó backup."
        )

    return total_coincidencias


def modo_guiado():
    """Flujo guiado por menú: elegir proveedor con flechas, correr dry-run
    automáticamente y preguntar con otro menú si se confirma la ejecución real."""
    proveedores = listar_proveedores()
    if not proveedores:
        sys.exit(
            "ERROR: no se encontró ningún proveedor con NOTAS.md dentro de "
            f"'{PROVEEDORES_DIR}'."
        )

    idx = menu_seleccionar(proveedores, "Elegí el proveedor")
    if idx is None:
        print("\nOperación cancelada. No se modificó nada.")
        return
    proveedor = proveedores[idx]

    print(f"\nCorriendo vista previa (dry-run) para '{proveedor}'...")
    coincidencias = procesar(proveedor, execute=False)

    if coincidencias == 0:
        print("\nNo hay ninguna coincidencia para modificar. No hace falta ejecutar nada.")
        return

    opciones = [
        "Sí, aplicar estos cambios reales al archivo",
        "No, salir sin modificar nada",
    ]
    idx_confirm = menu_seleccionar(
        opciones, "¿Confirmás aplicar estos cambios al archivo real?"
    )
    if idx_confirm == 0:
        print(f"\nAplicando cambios reales para '{proveedor}'...")
        procesar(proveedor, execute=True)
    else:
        print("\nNo se aplicó ningún cambio. Podés volver a correr el script cuando quieras.")


def main():
    parser = argparse.ArgumentParser(
        description="Actualiza precios (Tipo 1) para un proveedor según su NOTAS.md."
    )
    parser.add_argument(
        "proveedor",
        nargs="?",
        default=None,
        help=(
            "Nombre de la carpeta del proveedor en proveedores/. "
            "Si se omite, se abre el menú guiado."
        ),
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Ejecuta los cambios reales (por defecto corre en modo dry-run).",
    )
    args = parser.parse_args()

    if args.proveedor is None:
        # Sin argumentos: modo guiado con menú de flechas.
        modo_guiado()
        return

    # Con argumentos explícitos: comportamiento clásico por línea de comandos.
    procesar(args.proveedor, args.execute)


if __name__ == "__main__":
    main()
