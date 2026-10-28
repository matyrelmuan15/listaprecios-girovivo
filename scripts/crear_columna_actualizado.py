#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
crear_columna_actualizado.py
-----------------------------
Script de UN SOLO USO. Agrega la columna "ActualizadoSegun" (con lista
desplegable Manual / Dólar / Lista proveedor) al final de cada hoja de
destino. La hoja manual "Venzo Sin Actualizar" se conserva (el usuario la
borra a mano); queda reemplazada por la hoja automática "NoActualizados".

    python crear_columna_actualizado.py            -> vista previa (no escribe)
    python crear_columna_actualizado.py --execute  -> backup + cambios reales

Las hojas que ya tienen la columna se saltean, así que correrlo dos veces no
duplica nada. Cada encabezado creado se registra en RegistroCambios
(TipoOperacion "Estructura"). Las filas existentes quedan vacías: el usuario
las completa a mano.
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

import actualizado_segun as act
import registro

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"

HOJAS_NO_DESTINO = {
    "Hoja1", "Marcas Actualización Dinámica", registro.HOJA_REGISTRO,
    act.HOJA_NO_ACTUALIZADOS, *act.HOJAS_REEMPLAZADAS,
}
ENCABEZADOS_DESTINO = ("CodigoFactura", "PrecioConIVA", "FechaActualizacion")


def es_hoja_destino(ws):
    nombres = {str(c.value).strip().casefold() for c in ws[1] if c.value is not None}
    return all(e.casefold() in nombres for e in ENCABEZADOS_DESTINO)


def backup_archivo(archivo):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo.stem}_{ts}{archivo.suffix}"
    shutil.copy2(archivo, destino)
    return destino


def procesar(execute, archivo=ARCHIVO_PATH):
    if not archivo.exists():
        sys.exit(f"ERROR: no se encontró el archivo de trabajo en: {archivo}")
    wb = openpyxl.load_workbook(archivo)  # sin data_only=True
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        sys.exit(f"\nERROR: {e}")

    corrida = registro.RegistroCorrida(
        registro.TIPO_ESTRUCTURA,
        "Nueva columna ActualizadoSegun (Manual / Dólar / Lista proveedor)",
    )
    a_crear, ya_tienen = [], []
    for ws in wb.worksheets:
        if ws.title in HOJAS_NO_DESTINO or not es_hoja_destino(ws):
            continue
        if act.buscar_columna(ws):
            ya_tienen.append(ws.title)
            continue
        col = act.ultima_columna_encabezado(ws) + 1
        a_crear.append((ws, col))
        corrida.agregar(hoja=ws.title, fila=1, nombre_producto=None, marca=None,
                        columna=get_column_letter(col), valor_anterior=None,
                        valor_nuevo=act.COL_ACTUALIZADO)


    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Crear columna {act.COL_ACTUALIZADO} | Modo: {modo} ===\n")
    print(f"Hojas donde se agrega la columna: {len(a_crear)}")
    for ws, col in a_crear:
        print(f"  - {ws.title}: columna {get_column_letter(col)} "
              f"(después de '{ws.cell(row=1, column=col - 1).value}')")
    if ya_tienen:
        print(f"Hojas que ya la tienen (se saltean): {', '.join(ya_tienen)}")
    print(f"Lista desplegable: {', '.join(act.VALORES)} (filas 2 a {act.FILAS_VALIDACION})")
    print("Las filas existentes quedan vacías.")

    if not a_crear:
        print("\nNo hay nada para cambiar.")
        return 0

    if not execute:
        registro.imprimir_previsualizacion(corrida)
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo ni se creó backup.")
        return len(a_crear)

    for ws, _ in a_crear:
        act.crear_columna(ws)
    try:
        corrida.verificar_conteo(len(a_crear))
        corrida.escribir(wb)
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    backup = backup_archivo(archivo)
    print(f"\nBackup creado en: {backup}")
    wb.save(archivo)
    print(f"Archivo guardado: {archivo}")
    print(f"Filas agregadas a {registro.HOJA_REGISTRO}: {len(corrida)}")
    return len(a_crear)


def main():
    p = argparse.ArgumentParser(description="Crea la columna ActualizadoSegun (uso único).")
    p.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    args = p.parse_args()
    procesar(args.execute)


if __name__ == "__main__":
    main()
