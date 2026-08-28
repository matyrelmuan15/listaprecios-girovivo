#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
exportar_datos.py
-----------------
Script genérico Tipo 4 del proyecto "Lista de Precios Giro Vivo".

Exporta el código y el precio de cada producto del libro principal a un
archivo .xls con el formato exacto que espera el sistema externo:

    Hoja "Hoja Principal"
    Fila 1  -> encabezados: CODIGO | IMPORTE
    Fila 2+ -> un producto por fila

Mapeo de columnas:
    CodigoDux     -> CODIGO
    PrecioConIVA  -> IMPORTE

Reglas seguidas (según CLAUDE.md del proyecto):
  - El libro principal se abre en modo SOLO LECTURA y NUNCA se guarda:
    esta operación no modifica absolutamente nada del archivo original.
  - Por lo mismo, no se escribe ninguna fila en RegistroCambios (no hay
    ninguna celda modificada que registrar).
  - Columnas localizadas siempre por nombre de encabezado exacto (fila 1),
    nunca por posición ni letra de columna.
  - Se abre el libro SIN data_only=True, para no depender de valores
    cacheados por Excel.
  - Hojas excluidas SIEMPRE (no se leen): "Hoja1", "Marcas Actualización
    Dinámica" y "RegistroCambios".
  - Se recorren todas las demás hojas dinámicamente: si mañana se agrega
    una categoría nueva, entra sola sin tocar el código.
  - Antes de sobrescribir una exportación anterior, se guarda una copia con
    timestamp en exportacion/backups/.
  - Flujo obligatorio en 2 pasos: primero DRY-RUN (sin --execute), que
    muestra el detalle por hoja y una muestra sin escribir nada; recién con
    --execute se genera el archivo real.

Criterios de inclusión de una fila:
  - CodigoDux con contenido (las filas sin código se omiten).
  - PrecioConIVA numérico y mayor que cero (esto descarta filas vacías y
    las filas de agrupación visual, como los encabezados de subcategoría
    de la hoja "Accesorios", que vienen vacías o en 0).

Tipos de dato: CodigoDux se exporta TAL CUAL viene del libro. Los códigos
numéricos se escriben como número y los alfanuméricos como texto, sin
conversiones ni normalizaciones.

USO
---
  Vista previa (dry-run, no escribe nada):
    python exportar_datos.py

  Generar el archivo real (recién después de revisar el dry-run):
    python exportar_datos.py --execute
"""

import argparse
import datetime
import shutil
import sys
from pathlib import Path

import xlwt
from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"

EXPORT_DIR = PROJECT_ROOT / "exportacion"
EXPORT_BACKUP_DIR = EXPORT_DIR / "backups"
EXPORT_NOMBRE = "exportacion_giro_vivo.xls"
EXPORT_PATH = EXPORT_DIR / EXPORT_NOMBRE

# Hojas que no se leen bajo ninguna circunstancia.
HOJAS_EXCLUIDAS = {"Hoja1", "Marcas Actualización Dinámica", "RegistroCambios"}

# Columnas de origen (nombre de encabezado exacto en el libro principal).
COL_CODIGO_ORIGEN = "CodigoDux"
COL_PRECIO_ORIGEN = "PrecioConIVA"
COL_NOMBRE_ORIGEN = "NombreProducto"

# Formato de salida exigido por el sistema externo.
HOJA_DESTINO = "Hoja Principal"
ENCABEZADOS_DESTINO = ["CODIGO", "IMPORTE"]

# Límite duro del formato .xls antiguo (BIFF8): 65.536 filas por hoja.
LIMITE_FILAS_XLS = 65536

MUESTRA_FILAS = 5


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def encontrar_columnas(ws, requeridas, max_search_rows=3):
    """Busca encabezados exactos (comparación normalizada) en las primeras
    filas de la hoja. Devuelve {nombre_requerido: numero_de_columna}."""
    col_map = {}
    for row in ws.iter_rows(min_row=1, max_row=max_search_rows):
        for cell in row:
            if cell.value is None:
                continue
            nombre = str(cell.value).strip()
            for req in requeridas:
                if nombre.casefold() == req.casefold():
                    col_map.setdefault(req, cell.column)
    return col_map


def normalizar_codigo(valor):
    """Devuelve el CodigoDux tal cual, solo recortando espacios sobrantes si
    es texto. No convierte tipos: los números siguen siendo números."""
    if valor is None:
        return None
    if isinstance(valor, str):
        limpio = valor.strip()
        return limpio if limpio else None
    return valor


def parse_precio(valor):
    """Devuelve el precio como float si es un número válido y positivo.
    Devuelve None en cualquier otro caso (vacío, texto, fórmula, cero o
    negativo). Acepta texto numérico con coma decimal por si alguna celda
    quedó cargada como cadena."""
    if valor is None:
        return None
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        numero = float(valor)
    elif isinstance(valor, str):
        texto = valor.strip()
        if not texto or texto.startswith("="):
            return None
        texto = texto.replace("$", "").replace(" ", "")
        # Formato local: separador de miles "." y decimal ","
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        try:
            numero = float(texto)
        except ValueError:
            return None
    else:
        return None

    if numero <= 0:
        return None
    return round(numero, 2)


def es_formula(valor):
    return isinstance(valor, str) and valor.startswith("=")


# --------------------------------------------------------------------------
# Lectura del libro principal (solo lectura, nunca se guarda)
# --------------------------------------------------------------------------
def recolectar_datos():
    """Recorre dinámicamente las hojas de destino y devuelve
    (filas_exportables, reporte_por_hoja, hojas_salteadas)."""
    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")
        sys.exit(1)

    try:
        wb = load_workbook(ARCHIVO_PATH, data_only=False, read_only=True)
    except PermissionError:
        print(
            "ERROR: no se pudo leer el archivo. ¿Está abierto en Excel? "
            "Cerralo e intentá de nuevo."
        )
        sys.exit(1)
    except Exception as e:
        print(f"ERROR al leer el archivo: {e}")
        sys.exit(1)

    filas = []
    reporte = []
    salteadas = []

    try:
        for ws in wb.worksheets:
            if ws.title in HOJAS_EXCLUIDAS:
                continue

            requeridas = [COL_CODIGO_ORIGEN, COL_PRECIO_ORIGEN, COL_NOMBRE_ORIGEN]
            col_map = encontrar_columnas(ws, requeridas)

            faltantes = [
                c for c in (COL_CODIGO_ORIGEN, COL_PRECIO_ORIGEN) if c not in col_map
            ]
            if faltantes:
                salteadas.append({"hoja": ws.title, "faltantes": faltantes})
                continue

            c_codigo = col_map[COL_CODIGO_ORIGEN]
            c_precio = col_map[COL_PRECIO_ORIGEN]
            c_nombre = col_map.get(COL_NOMBRE_ORIGEN)

            incluidas = 0
            sin_codigo = 0
            sin_precio = 0
            precio_no_numerico = 0

            for row_idx, fila in enumerate(
                ws.iter_rows(min_row=2, values_only=True), start=2
            ):
                valor_codigo = fila[c_codigo - 1] if len(fila) >= c_codigo else None
                valor_precio = fila[c_precio - 1] if len(fila) >= c_precio else None
                valor_nombre = (
                    fila[c_nombre - 1] if c_nombre and len(fila) >= c_nombre else None
                )

                codigo = normalizar_codigo(valor_codigo)
                precio = parse_precio(valor_precio)

                # Fila completamente vacía: no cuenta como descarte.
                if codigo is None and precio is None and not valor_nombre:
                    continue

                if precio is None:
                    if es_formula(valor_precio) or (
                        valor_precio is not None and str(valor_precio).strip() != ""
                    ):
                        precio_no_numerico += 1
                    else:
                        sin_precio += 1
                    continue

                if codigo is None:
                    sin_codigo += 1
                    continue

                incluidas += 1
                filas.append(
                    {
                        "codigo": codigo,
                        "precio": precio,
                        "hoja": ws.title,
                        "fila": row_idx,
                        "nombre": valor_nombre,
                    }
                )

            reporte.append(
                {
                    "hoja": ws.title,
                    "incluidas": incluidas,
                    "sin_codigo": sin_codigo,
                    "sin_precio": sin_precio,
                    "precio_no_numerico": precio_no_numerico,
                }
            )
    finally:
        wb.close()

    return filas, reporte, salteadas


def detectar_duplicados(filas):
    """Devuelve {codigo_normalizado: [filas...]} para los códigos que
    aparecen más de una vez. El sistema externo suele rechazar códigos
    repetidos, así que conviene verlos antes de subir el archivo."""
    vistos = {}
    for item in filas:
        clave = str(item["codigo"]).strip().casefold()
        vistos.setdefault(clave, []).append(item)
    return {k: v for k, v in vistos.items() if len(v) > 1}


# --------------------------------------------------------------------------
# Escritura del archivo .xls
# --------------------------------------------------------------------------
def backup_exportacion_anterior():
    """Si ya existe una exportación previa, la copia con timestamp a
    exportacion/backups/ antes de sobrescribirla. Devuelve la ruta de la
    copia, o None si no había archivo anterior."""
    if not EXPORT_PATH.exists():
        return None
    EXPORT_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = EXPORT_BACKUP_DIR / f"{EXPORT_PATH.stem}_{timestamp}{EXPORT_PATH.suffix}"
    shutil.copy2(EXPORT_PATH, destino)
    return destino


def escribir_xls(filas, destino):
    """Genera el archivo .xls con el formato exacto que pide el sistema
    externo. El contenido se reescribe completo en cada corrida."""
    libro = xlwt.Workbook(encoding="utf-8")
    hoja = libro.add_sheet(HOJA_DESTINO)

    estilo_encabezado = xlwt.easyxf("font: bold on")
    estilo_importe = xlwt.easyxf(num_format_str="0.00")

    for col, titulo in enumerate(ENCABEZADOS_DESTINO):
        hoja.write(0, col, titulo, estilo_encabezado)

    for i, item in enumerate(filas, start=1):
        # CodigoDux se escribe tal cual: xlwt respeta el tipo de Python
        # (int/float -> número, str -> texto).
        hoja.write(i, 0, item["codigo"])
        hoja.write(i, 1, item["precio"], estilo_importe)

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    libro.save(str(destino))


# --------------------------------------------------------------------------
# Reporte por consola
# --------------------------------------------------------------------------
def imprimir_reporte(filas, reporte, salteadas, duplicados):
    print("=" * 70)
    print("EXPORTAR DATOS - Tipo 4 (solo lectura, no modifica el libro)")
    print(f"Origen : {ARCHIVO_PATH}")
    print(f"Destino: {EXPORT_PATH}")
    print(f"Formato: hoja '{HOJA_DESTINO}' | columnas {' | '.join(ENCABEZADOS_DESTINO)}")
    print("=" * 70)

    for item in salteadas:
        print(f"\n[SALTEADA] Hoja '{item['hoja']}': faltan columnas {item['faltantes']}")

    total_incluidas = 0
    total_sin_codigo = 0
    total_sin_precio = 0
    total_no_numerico = 0

    for item in reporte:
        total_incluidas += item["incluidas"]
        total_sin_codigo += item["sin_codigo"]
        total_sin_precio += item["sin_precio"]
        total_no_numerico += item["precio_no_numerico"]

        if (
            item["incluidas"] == 0
            and item["sin_codigo"] == 0
            and item["precio_no_numerico"] == 0
        ):
            continue

        print(f"\nHoja '{item['hoja']}':")
        print(f"  Productos a exportar: {item['incluidas']}")
        if item["sin_precio"]:
            print(f"  Omitidos por no tener PrecioConIVA: {item['sin_precio']}")
        if item["precio_no_numerico"]:
            print(
                f"  Omitidos por PrecioConIVA no numérico o en cero: "
                f"{item['precio_no_numerico']}"
            )
        if item["sin_codigo"]:
            print(
                f"  [ATENCIÓN] Omitidos por no tener CodigoDux (tenían precio): "
                f"{item['sin_codigo']}"
            )

    print("\n" + "-" * 70)
    print("TOTALES")
    print(f"  Productos que se van a exportar: {total_incluidas}")
    print(f"  Omitidos sin PrecioConIVA: {total_sin_precio}")
    print(f"  Omitidos con PrecioConIVA no numérico o en cero: {total_no_numerico}")
    print(f"  Omitidos con precio pero SIN CodigoDux: {total_sin_codigo}")
    print("-" * 70)

    if total_sin_codigo:
        print(
            "\n[ATENCIÓN] Hay productos con precio cargado que no tienen CodigoDux "
            "y no se van a exportar. Detalle:"
        )
        mostrados = 0
        for item in reporte:
            if item["sin_codigo"]:
                print(f"  - Hoja '{item['hoja']}': {item['sin_codigo']} producto(s)")
                mostrados += 1
            if mostrados >= 10:
                break

    if duplicados:
        print(f"\n[ATENCIÓN] Hay {len(duplicados)} código(s) repetido(s) en la exportación:")
        for clave, items in list(duplicados.items())[:10]:
            ubicaciones = ", ".join(f"{i['hoja']}!fila {i['fila']}" for i in items)
            print(f"  - {items[0]['codigo']!r} aparece {len(items)} veces: {ubicaciones}")
        if len(duplicados) > 10:
            print(f"  ... y {len(duplicados) - 10} más.")

    if filas:
        print(f"\nMuestra de las primeras {min(MUESTRA_FILAS, len(filas))} filas del archivo:")
        print(f"  {'CODIGO':<24} {'IMPORTE':>12}   (origen)")
        for item in filas[:MUESTRA_FILAS]:
            origen = f"{item['hoja']}!fila {item['fila']}"
            print(f"  {str(item['codigo']):<24} {item['precio']:>12.2f}   {origen}")


# --------------------------------------------------------------------------
# Flujo principal
# --------------------------------------------------------------------------
def procesar(execute):
    filas, reporte, salteadas = recolectar_datos()
    duplicados = detectar_duplicados(filas)

    imprimir_reporte(filas, reporte, salteadas, duplicados)

    if not filas:
        print("\nNo hay ningún producto para exportar. No se generó ningún archivo.")
        return

    if len(filas) + 1 > LIMITE_FILAS_XLS:
        print(
            f"\nERROR: la exportación tiene {len(filas)} filas y el formato .xls "
            f"admite como máximo {LIMITE_FILAS_XLS - 1}. No se generó ningún archivo."
        )
        sys.exit(1)

    if not execute:
        print("\n[DRY-RUN] No se generó ningún archivo. Revisá el detalle de arriba.")
        print("Si está todo correcto, volvé a correr el mismo comando agregando --execute.")
        return

    copia = backup_exportacion_anterior()
    if copia:
        print(f"\nCopia de la exportación anterior guardada en: {copia}")
    else:
        print("\nNo había exportación anterior: no hizo falta guardar copia.")

    try:
        escribir_xls(filas, EXPORT_PATH)
    except PermissionError:
        print(
            f"\nERROR: no se pudo escribir en {EXPORT_PATH}. "
            "¿Está abierto en Excel? Cerralo e intentá de nuevo."
        )
        sys.exit(1)

    print(f"\nExportación generada: {EXPORT_PATH}")
    print(f"Productos exportados: {len(filas)}")
    print("El libro principal no fue modificado (operación de solo lectura).")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Exporta CodigoDux -> CODIGO y PrecioConIVA -> IMPORTE del libro "
            "principal a un archivo .xls con el formato del sistema externo (Tipo 4)."
        )
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Genera el archivo real (por defecto corre en modo dry-run).",
    )
    args = parser.parse_args()
    procesar(args.execute)


if __name__ == "__main__":
    main()
