#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
asignar_categoria.py
---------------------
Asigna un texto fijo en la columna "Categoria" de las hojas de destino cuyo
nombre contenga alguno de los filtros indicados (sin importar mayúsculas ni
acentos).

    python asignar_categoria.py --filtro "bicicleta" --valor "Bicicletas"
        -> inspección + dry-run (no escribe nada)
    python asignar_categoria.py --filtro "bicicleta" --valor "Bicicletas" --execute
        -> backup + cambios reales

Reglas:
  - Si la hoja no tiene "Categoria", se crea el encabezado en la primera
    columna libre a la derecha del último encabezado (nunca se insertan
    columnas en el medio), copiando el formato y el ancho del encabezado
    vecino.
  - Si la hoja ya tiene "Categoria", se usa esa columna.
  - Fila de producto = fila con CodigoDux no vacío.
  - Celdas que ya tienen el valor pedido: se saltean (correrlo dos veces no
    cambia nada). Celdas con OTRA categoría: no se pisan salvo con
    --sobrescribir; se informan.
  - Valores fijos (texto), nunca fórmulas. Sin lista desplegable.
  - Registro en RegistroCambios, una fila por celda modificada:
    encabezado nuevo -> TipoOperacion "Estructura" (Columna = letra);
    valor de categoría -> TipoOperacion "Categoria" (Columna = "Categoria").
"""

import argparse
import shutil
import sys
import unicodedata
from copy import copy
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

COL_CATEGORIA = "Categoria"
COL_CODIGO_DUX = "CodigoDux"
COL_NOMBRE = "NombreProducto"
COL_MARCA = "Marca"
ENCABEZADOS_DESTINO = (COL_CODIGO_DUX, "PrecioConIVA")

HOJAS_EXCLUIDAS = {
    "Hoja1", "Marcas Actualización Dinámica", registro.HOJA_REGISTRO,
    act.HOJA_NO_ACTUALIZADOS, *act.HOJAS_REEMPLAZADAS,
}


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def normalizar(texto):
    """Sin acentos, sin espacios sobrantes y en minúsculas."""
    if texto is None:
        return ""
    sin_acentos = unicodedata.normalize("NFKD", str(texto))
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return sin_acentos.strip().casefold()


def es_vacio(valor):
    return valor is None or str(valor).strip() == ""


def mapa_encabezados(ws):
    """{encabezado_normalizado: número de columna} de la fila 1."""
    mapa = {}
    for celda in ws[1]:
        if not es_vacio(celda.value):
            mapa.setdefault(normalizar(celda.value), celda.column)
    return mapa


def hoja_coincide(titulo, filtros):
    t = normalizar(titulo)
    return any(f in t for f in filtros)


def backup_archivo(archivo):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo.stem}_{ts}{archivo.suffix}"
    shutil.copy2(archivo, destino)
    return destino


# --------------------------------------------------------------------------
# Planificación (no escribe nada)
# --------------------------------------------------------------------------
def planificar_hoja(ws, valor, sobrescribir):
    enc = mapa_encabezados(ws)
    col = enc.get(normalizar(COL_CATEGORIA))
    crear = col is None
    if crear:
        col = act.ultima_columna_encabezado(ws) + 1

    c_dux = enc[normalizar(COL_CODIGO_DUX)]
    c_nombre = enc.get(normalizar(COL_NOMBRE))
    c_marca = enc.get(normalizar(COL_MARCA))

    a_escribir, ya_ok, distintas, ocupadas_sin_dux = [], 0, [], []
    for r in range(2, ws.max_row + 1):
        actual = ws.cell(row=r, column=col).value
        if es_vacio(ws.cell(row=r, column=c_dux).value):
            if not crear and not es_vacio(actual):
                ocupadas_sin_dux.append((r, actual))
            continue
        if not es_vacio(actual) and str(actual).strip() == valor:
            ya_ok += 1
            continue
        if not es_vacio(actual) and not sobrescribir:
            distintas.append((r, actual))
            continue
        a_escribir.append({
            "fila": r,
            "anterior": None if es_vacio(actual) else actual,
            "nombre": ws.cell(row=r, column=c_nombre).value if c_nombre else None,
            "marca": ws.cell(row=r, column=c_marca).value if c_marca else None,
        })

    return {
        "ws": ws, "col": col, "crear": crear, "a_escribir": a_escribir,
        "ya_ok": ya_ok, "distintas": distintas,
        "ocupadas_sin_dux": ocupadas_sin_dux,
    }


# --------------------------------------------------------------------------
# Escritura
# --------------------------------------------------------------------------
def crear_encabezado(ws, col):
    nueva = ws.cell(row=1, column=col)
    nueva.value = COL_CATEGORIA
    if col > 1:
        modelo = ws.cell(row=1, column=col - 1)
        if modelo.has_style:
            nueva._style = copy(modelo._style)
        ancho = ws.column_dimensions[get_column_letter(col - 1)].width
        ws.column_dimensions[get_column_letter(col)].width = max(ancho or 0, 18)


def escribir_valor(ws, fila, col, valor, copiar_estilo):
    celda = ws.cell(row=fila, column=col)
    celda.value = valor
    if copiar_estilo and col > 1:
        vecina = ws.cell(row=fila, column=col - 1)
        if vecina.has_style:
            celda._style = copy(vecina._style)


# --------------------------------------------------------------------------
# Proceso principal
# --------------------------------------------------------------------------
def procesar(filtros, valor, execute, sobrescribir, archivo=ARCHIVO_PATH):
    if not archivo.exists():
        sys.exit(f"ERROR: no se encontró el archivo de trabajo en: {archivo}")
    valor = valor.strip()
    if not valor:
        sys.exit("ERROR: --valor no puede estar vacío.")
    filtros = [normalizar(f) for f in filtros if normalizar(f)]
    if not filtros:
        sys.exit("ERROR: indicá al menos un --filtro no vacío.")

    wb = openpyxl.load_workbook(archivo)  # sin data_only=True
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        sys.exit(f"\nERROR: {e}")

    # ---------------- Fase 1: inspección ----------------
    planes, sin_esquema = [], []
    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS or not hoja_coincide(ws.title, filtros):
            continue
        enc = mapa_encabezados(ws)
        if not all(normalizar(e) in enc for e in ENCABEZADOS_DESTINO):
            sin_esquema.append(ws.title)
            continue
        planes.append(planificar_hoja(ws, valor, sobrescribir))

    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Asignar Categoria | Modo: {modo} ===")
    print(f"Filtro de hojas: {', '.join(filtros)} | Valor: '{valor}'"
          f" | Sobrescribir otras categorías: {'sí' if sobrescribir else 'no'}\n")

    print("--- Fase 1: Inspección ---")
    if not planes:
        print("No se encontró ninguna hoja de destino que coincida con el filtro.")
    for p in planes:
        ws, letra = p["ws"], get_column_letter(p["col"])
        if p["crear"]:
            vecino = ws.cell(row=1, column=p["col"] - 1).value
            donde = f"columna {letra} (nueva, después de '{vecino}')"
        else:
            donde = f"columna {letra} (ya existe)"
        print(f"  - {ws.title}: {donde} | filas a marcar: {len(p['a_escribir'])}"
              f" | ya con '{valor}': {p['ya_ok']}"
              f" | con otra categoría (no se tocan): {len(p['distintas'])}")
        for r, v in p["distintas"][:5]:
            print(f"      fila {r}: '{v}'")
        if p["ocupadas_sin_dux"]:
            print(f"      Aviso: {len(p['ocupadas_sin_dux'])} celda(s) de Categoria con "
                  f"valor en filas sin CodigoDux (no se tocan).")
    if sin_esquema:
        print(f"  Hojas que coinciden pero no son de destino (se ignoran): "
              f"{', '.join(sin_esquema)}")

    # ---------------- Fase 2: dry-run / resumen ----------------
    corrida_enc = registro.RegistroCorrida(
        registro.TIPO_ESTRUCTURA, f"Nueva columna {COL_CATEGORIA}")
    corrida_val = registro.RegistroCorrida(
        registro.TIPO_CATEGORIA, f"Asignación de categoría '{valor}'",
        fecha=corrida_enc.fecha)

    for p in planes:
        ws, letra = p["ws"], get_column_letter(p["col"])
        if p["crear"]:
            corrida_enc.agregar(hoja=ws.title, fila=1, nombre_producto=None,
                                marca=None, columna=letra, valor_anterior=None,
                                valor_nuevo=COL_CATEGORIA)
        for f in p["a_escribir"]:
            corrida_val.agregar(hoja=ws.title, fila=f["fila"],
                                nombre_producto=f["nombre"], marca=f["marca"],
                                columna=COL_CATEGORIA, valor_anterior=f["anterior"],
                                valor_nuevo=valor)

    n_enc, n_val = len(corrida_enc), len(corrida_val)
    total = n_enc + n_val
    print("\n--- Fase 2: Resumen ---")
    print(f"  Hojas procesadas: {len(planes)}")
    print(f"  Encabezados nuevos (TipoOperacion '{registro.TIPO_ESTRUCTURA}'): {n_enc}")
    print(f"  Celdas de categoría (TipoOperacion '{registro.TIPO_CATEGORIA}'): {n_val}")
    print(f"  Total de celdas a modificar = filas de RegistroCambios: {total}")

    if total == 0:
        print("\nNo hay nada para cambiar.")
        return 0

    if not execute:
        registro.imprimir_previsualizacion(corrida_enc)
        registro.imprimir_previsualizacion(corrida_val)
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo "
              "ni se creó backup.")
        return total

    # ---------------- Fase 3: ejecución ----------------
    modificadas = 0
    for p in planes:
        ws, col = p["ws"], p["col"]
        if p["crear"]:
            crear_encabezado(ws, col)
            modificadas += 1
        for f in p["a_escribir"]:
            escribir_valor(ws, f["fila"], col, valor, copiar_estilo=p["crear"])
            modificadas += 1

    corrida_enc.filas.extend(corrida_val.filas)
    try:
        corrida_enc.verificar_conteo(modificadas)
        corrida_enc.escribir(wb)
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    backup = backup_archivo(archivo)
    print(f"\nBackup creado en: {backup}")
    wb.save(archivo)
    print(f"Archivo guardado: {archivo}")
    print(f"Celdas modificadas: {modificadas} | Filas agregadas a "
          f"{registro.HOJA_REGISTRO}: {len(corrida_enc)}")
    return modificadas


def main():
    p = argparse.ArgumentParser(
        description="Asigna un valor fijo en la columna Categoria de las hojas filtradas.")
    p.add_argument("--filtro", action="append", required=True,
                   help="Fragmento del nombre de hoja (se puede repetir). "
                        "Sin distinguir mayúsculas ni acentos.")
    p.add_argument("--valor", required=True, help="Texto de la categoría a escribir.")
    p.add_argument("--sobrescribir", action="store_true",
                   help="Reemplaza celdas que ya tienen otra categoría.")
    p.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    args = p.parse_args()
    procesar(args.filtro, args.valor, args.execute, args.sobrescribir)


if __name__ == "__main__":
    main()
