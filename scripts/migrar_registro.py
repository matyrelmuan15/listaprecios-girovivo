#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
migrar_registro.py
-------------------
Script de UN SOLO USO: rehace la hoja "RegistroCambios" con el formato
único de 10 columnas definido en registro.py.

Por qué existe
--------------
La hoja quedó con tres formatos conviviendo bajo los 5 encabezados
originales: el historial viejo, las filas de 7 columnas de
actualizar_precios.py / actualizar_fecha.py, y las de 9 columnas de
asignar_dolar.py. Como la hoja ya existía, ninguno de los scripts
reescribía los encabezados, así que los datos quedaban corridos bajo
títulos que no correspondían.

Qué hace
--------
Con --execute: backup con timestamp en backup/, elimina la hoja existente
y crea una nueva, vacía, con los 10 encabezados del formato único.

El historial anterior NO se conserva dentro del libro. Queda íntegro en el
backup con timestamp que se genera antes de tocar nada, que es una copia
completa del .xlsx: si más adelante hace falta rescatar algún dato viejo,
está ahí.

USO
---
  Vista previa (no escribe nada, es el modo por defecto):
      python migrar_registro.py

  Migración real (recién después de revisar la vista previa):
      python migrar_registro.py --execute

Reglas seguidas (CLAUDE.md):
  - Nunca se modifica el archivo real sin backup previo con timestamp.
  - Se abre el libro SIN data_only=True, para no destruir fórmulas.
  - Flujo obligatorio: primero dry-run, después ejecución.
"""

import argparse
import datetime
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

import registro

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"

# Una fila con esta cantidad de celdas con dato (o menos) no es un registro
# real: son los resúmenes manuales y las filas sueltas del historial viejo.
ANCHO_FILA_ATIPICA = 2


def hacer_backup(archivo_path):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo_path.stem}_{timestamp}{archivo_path.suffix}"
    shutil.copy2(archivo_path, destino)
    return destino


def ancho_efectivo(valores):
    """Cantidad de celdas con dato de una fila (ignora las vacías)."""
    return sum(1 for v in valores if v is not None and str(v).strip() != "")


def formatear_valores(valores):
    """Recorta las celdas vacías de la derecha para que el listado se lea."""
    recortada = list(valores)
    while recortada and (recortada[-1] is None or str(recortada[-1]).strip() == ""):
        recortada.pop()
    return " | ".join("(vacío)" if v is None else str(v) for v in recortada)


def inspeccionar(ws):
    """Relevamiento de la hoja actual, sin escribir nada."""
    filas = list(ws.iter_rows(min_row=1, values_only=True))
    encabezados = list(filas[0]) if filas else []

    distribucion = {}
    filas_datos = 0
    atipicas = []

    for numero, valores in enumerate(filas[1:], start=2):
        ancho = ancho_efectivo(valores)
        distribucion[ancho] = distribucion.get(ancho, 0) + 1
        if ancho > 0:
            filas_datos += 1
        if 0 < ancho <= ANCHO_FILA_ATIPICA:
            atipicas.append((numero, ancho, valores))

    return {
        "encabezados": encabezados,
        "filas_datos": filas_datos,
        "distribucion": distribucion,
        "atipicas": atipicas,
        "max_row": ws.max_row,
        "max_column": ws.max_column,
    }


def imprimir_dry_run(info):
    print("\n--- Estado actual de la hoja ---")
    print(f"  Dimensiones: max_row={info['max_row']}, max_column={info['max_column']}")
    print(f"  Encabezados actuales ({len(info['encabezados'])} columnas):")
    for i, h in enumerate(info["encabezados"], start=1):
        print(f"    {i:>2}. {'(vacío)' if h is None else h}")

    print(f"\n  Filas con datos: {info['filas_datos']}")
    print("  Ancho real de las filas (celdas con dato -> cantidad de filas):")
    for ancho in sorted(info["distribucion"]):
        cantidad = info["distribucion"][ancho]
        etiqueta = {
            0: "filas totalmente vacías",
            1: "filas sueltas",
            2: "resumen manual",
        }.get(ancho, "")
        if 5 <= ancho <= 6:
            etiqueta = "historial viejo (DolarReferencia)"
        elif ancho == 7:
            etiqueta = "actualizar_precios.py / actualizar_fecha.py"
        elif ancho == 9:
            etiqueta = "asignar_dolar.py"
        elif ancho in (3, 4):
            etiqueta = etiqueta or "historial viejo"
        sufijo = f"   <- {etiqueta}" if etiqueta else ""
        print(f"    {ancho:>2} celdas: {cantidad:>5} filas{sufijo}")

    print("\n--- Filas de resumen manual / atípicas ---")
    print(
        f"  (filas con {ANCHO_FILA_ATIPICA} celdas con dato o menos: no son "
        "registros de cambios reales)"
    )
    if info["atipicas"]:
        for numero, ancho, valores in info["atipicas"]:
            print(f"    fila {numero:>5} ({ancho} celdas): {formatear_valores(valores)}")
    else:
        print("    (ninguna)")

    print("\n--- Formato único que va a quedar ---")
    for i, h in enumerate(registro.REGISTRO_HEADERS, start=1):
        print(f"    {i:>2}. {h}")


def migrar(archivo_path, execute):
    if not archivo_path.exists():
        sys.exit(f"ERROR: no se encontró el archivo de trabajo en: {archivo_path}")

    modo = "EJECUCIÓN" if execute else "DRY-RUN (no se escribe nada)"
    print("=" * 70)
    print(f"MIGRACIÓN DE '{registro.HOJA_REGISTRO}' AL FORMATO ÚNICO | Modo: {modo}")
    print("=" * 70)
    print(f"Archivo: {archivo_path}")

    wb = load_workbook(archivo_path)  # sin data_only=True

    existia = registro.HOJA_REGISTRO in wb.sheetnames
    filas_a_descartar = 0

    if not existia:
        print(
            f"\nAVISO: la hoja '{registro.HOJA_REGISTRO}' no existe en el libro. "
            "Se va a crear vacía con los 10 encabezados del formato único."
        )
        print("\n--- Formato único que va a quedar ---")
        for i, h in enumerate(registro.REGISTRO_HEADERS, start=1):
            print(f"    {i:>2}. {h}")
    else:
        ws = wb[registro.HOJA_REGISTRO]
        info = inspeccionar(ws)
        filas_a_descartar = info["filas_datos"]
        imprimir_dry_run(info)

    print("\n" + "-" * 70)
    print("RESUMEN")
    if existia:
        print(f"  Filas de datos que se van a DESCARTAR: {filas_a_descartar}")
        print("  La hoja se elimina y se recrea vacía, solo con encabezados.")
    else:
        print("  Filas de datos que se van a descartar: 0 (la hoja no existía).")
    print(
        "  El historial completo queda preservado en el backup con timestamp "
        "que se genera antes de tocar el archivo."
    )
    print("-" * 70)

    if not execute:
        print(
            "\n[DRY-RUN] No se modificó nada, no se creó backup. "
            "Si está todo correcto, volvé a correr con --execute."
        )
        return

    destino_backup = hacer_backup(archivo_path)
    print(f"\nBackup creado en: {destino_backup}")

    if existia:
        del wb[registro.HOJA_REGISTRO]
        print(f"Hoja '{registro.HOJA_REGISTRO}' eliminada.")

    ws_nueva = wb.create_sheet(registro.HOJA_REGISTRO)
    registro.escribir_encabezados(ws_nueva)
    print(
        f"Hoja '{registro.HOJA_REGISTRO}' recreada con "
        f"{len(registro.REGISTRO_HEADERS)} encabezados."
    )

    # Chequeo de cierre: la hoja nueva tiene que pasar la validación que van
    # a correr los tres scripts.
    registro.validar_encabezados(wb)

    wb.save(archivo_path)
    print(f"Archivo guardado: {archivo_path}")
    print("\nMigración completada. Los tres scripts ya pueden escribir en la hoja.")


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Rehace la hoja RegistroCambios con el formato único de 10 "
            "columnas. Por defecto corre en dry-run."
        )
    )
    parser.add_argument(
        "--archivo",
        default=None,
        help=f"Ruta al libro Excel. Si se omite, se usa: {ARCHIVO_PATH}",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Aplica la migración real (backup + eliminar + recrear la hoja).",
    )
    args = parser.parse_args()

    archivo_path = Path(args.archivo) if args.archivo else ARCHIVO_PATH
    migrar(archivo_path, args.execute)


if __name__ == "__main__":
    main()
