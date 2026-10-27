#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verificar_formulas.py
----------------------
Script Tipo 5 del proyecto "Lista de Precios Giro Vivo".

Revisa TODAS las hojas de destino y deja PrecioContado y PrecioSinIVA con su
fórmula correcta en cada fila que tenga PrecioConIVA:

    PrecioContado = ='<Hoja>'!$H6-('<Hoja>'!$H6*15/100)
    PrecioSinIVA  = ='<Hoja>'!$H6/1.21

También redondea a 2 decimales los PrecioConIVA que tengan más.

Reglas seguidas (CLAUDE.md):
  - Backup con timestamp antes de guardar.
  - Columnas por nombre de encabezado, nunca por letra.
  - Libro abierto SIN data_only=True.
  - Excluye "Marcas Actualización Dinámica", Hoja1 y RegistroCambios.
  - Una fila de RegistroCambios por celda modificada, con verificación de
    conteo.
  - Flujo en fases: sin --execute solo muestra la simulación.

USO
    python verificar_formulas.py              (simulación y pregunta si aplicar)
    python verificar_formulas.py --dry-run    (solo simulación)
    python verificar_formulas.py --execute    (aplica sin preguntar)
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

import formulas
import registro

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"


def backup_archivo():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{ARCHIVO_PATH.stem}_{timestamp}{ARCHIVO_PATH.suffix}"
    shutil.copy2(ARCHIVO_PATH, destino)
    return destino


def procesar(execute):
    if not ARCHIVO_PATH.exists():
        sys.exit(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")

    wb = openpyxl.load_workbook(ARCHIVO_PATH)  # sin data_only=True
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    cambios, para_revisar = formulas.revisar(wb)
    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Verificación de fórmulas | Modo: {modo} ===")
    formulas.imprimir_resumen(cambios, para_revisar)

    if not cambios:
        print("\nNo hay nada para corregir. No se hizo backup ni se escribió nada.")
        return 0

    corrida = registro.RegistroCorrida(
        registro.TIPO_FORMULAS,
        "Verificación de fórmulas: PrecioContado = ConIVA - 15 %, "
        "PrecioSinIVA = ConIVA / 1,21, PrecioConIVA a 2 decimales",
    )
    formulas.registrar(corrida, cambios)

    try:
        corrida.verificar_conteo(len(cambios))
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    if not execute:
        registro.imprimir_previsualizacion(corrida)
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo ni se creó backup.")
        return len(cambios)

    formulas.aplicar(wb, cambios)
    try:
        corrida.escribir(wb)
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    backup_path = backup_archivo()
    print(f"\nBackup creado en: {backup_path}")
    try:
        wb.save(ARCHIVO_PATH)
    except PermissionError:
        sys.exit(
            "\nERROR: no se pudo guardar porque el archivo está abierto en Excel. "
            "Cerralo y volvé a correr el script. El archivo real no se modificó."
        )
    print(f"Archivo guardado: {ARCHIVO_PATH}")
    print(f"Celdas corregidas: {len(cambios)} | Filas agregadas a {registro.HOJA_REGISTRO}: {len(corrida)}")
    return len(cambios)


def main():
    parser = argparse.ArgumentParser(
        description="Verifica y repara las fórmulas de PrecioContado y PrecioSinIVA."
    )
    grupo = parser.add_mutually_exclusive_group()
    grupo.add_argument("--execute", action="store_true", help="Aplica los cambios sin preguntar.")
    grupo.add_argument("--dry-run", action="store_true", help="Solo muestra la simulación.")
    args = parser.parse_args()

    if args.execute:
        procesar(execute=True)
        return
    pendientes = procesar(execute=False)
    if args.dry_run or not pendientes:
        return
    respuesta = input("\n¿Confirmás aplicar estas correcciones al archivo real? (s/n): ").strip().lower()
    if respuesta in ("s", "si", "sí"):
        procesar(execute=True)
    else:
        print("\nNo se aplicó ningún cambio.")


if __name__ == "__main__":
    main()
