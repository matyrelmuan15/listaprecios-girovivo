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

Además, en cada corrida (simulación o ejecución) hace un control de SOLO
LECTURA de la hoja "Marcas Actualización Dinámica": toma la capacidad de los
rangos de J1, A2 y E3 y avisa si las filas de las tablas, las marcas con fecha
o los productos desactualizados superan el 90 %, o si alguna de las 6
fórmulas perdió el envoltorio EXPAND(TAKE(. No escribe nada.

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


# ---------------------------------------------------------------------------
# Control de capacidad de "Marcas Actualización Dinámica" (SOLO LECTURA)
# ---------------------------------------------------------------------------
# Las 6 fórmulas matriciales de esa hoja tienen rango fijo (openpyxl no
# conserva las matrices dinámicas). Desde el 22/11/2026 están envueltas en
# EXPAND(TAKE(...)) con capacidad 5.000 filas (J:M), 200 marcas (A2) y
# 2.998 productos (E3). Este control solo LEE la hoja y las tablas: no modifica
# ninguna celda ni escribe en RegistroCambios.
HOJA_MARCAS = "Marcas Actualización Dinámica"
MARCAS_MAESTRAS = ("J1", "K1", "L1", "M1", "A2", "E3")
MARCAS_UMBRAL = 0.90
MARCAS_DIAS = 30


def _filas_de_ref(ref):
    """Cantidad de filas de un rango tipo 'A2:C201'."""
    import re
    m = re.fullmatch(r"\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?", ref or "")
    if not m:
        return None
    ini = int(m.group(1))
    fin = int(m.group(2) or m.group(1))
    return fin - ini + 1


def _fecha_serial(valor):
    """Replica la columna M de Marcas: número > 0, fecha, o texto d/m/aaaa.
    Devuelve el número de serie de Excel o None si no es una fecha válida."""
    import datetime as _dt
    import re
    base = _dt.datetime(1899, 12, 30)
    if isinstance(valor, bool) or valor is None:
        return None
    if isinstance(valor, _dt.datetime):
        return (valor - base).days
    if isinstance(valor, _dt.date):
        return (_dt.datetime(valor.year, valor.month, valor.day) - base).days
    if isinstance(valor, (int, float)):
        return int(valor) if valor > 0 else None
    m = re.fullmatch(r"\s*(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{2,4})\s*", str(valor))
    if not m:
        return None
    try:
        d = _dt.datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None
    return (d - base).days


def controlar_marcas(wb):
    """Revisa que la hoja Marcas tenga capacidad suficiente. Solo lectura.

    Devuelve la lista de avisos (vacía si todo está bien) y la imprime."""
    import re
    from datetime import date
    from openpyxl.worksheet.formula import ArrayFormula

    print("\n--- Control de capacidad: hoja Marcas (solo lectura) ---")
    avisos = []
    if HOJA_MARCAS not in wb.sheetnames:
        avisos.append(f"No existe la hoja '{HOJA_MARCAS}'.")
        for a in avisos:
            print(f"  AVISO: {a}")
        return avisos
    ws = wb[HOJA_MARCAS]

    # 1) Fórmulas maestras: matriciales y con EXPAND(TAKE(
    capacidad = {}
    texto_j1 = ""
    for celda in MARCAS_MAESTRAS:
        valor = ws[celda].value
        if not isinstance(valor, ArrayFormula):
            avisos.append(f"{celda} no es una fórmula matricial: la lista dinámica no funcionará.")
            continue
        texto = valor.text or ""
        if celda == "J1":
            texto_j1 = texto
        limpio = texto.replace("_xlfn.", "").replace("_xlws.", "").lstrip("=")
        if not limpio.startswith("EXPAND(TAKE("):
            avisos.append(f"{celda} no tiene EXPAND(TAKE(: volvió al formato de tamaño fijo "
                          "(aparecerán #N/A o se perderán productos). Ver scripts/ampliar_marcas.py.")
        capacidad[celda] = _filas_de_ref(valor.ref)

    # 2) Uso actual, calculado desde las tablas que alimenta J1
    nombres_tabla = list(dict.fromkeys(re.findall(r"(Table_\d+)\[", texto_j1)))
    tablas = {}
    for hoja in wb.worksheets:
        for tabla in hoja.tables.values():
            tablas[tabla.name] = (hoja, tabla.ref)
    faltantes = [n for n in nombres_tabla if n not in tablas]
    if faltantes:
        avisos.append(f"J1 usa tablas que no existen en el libro: {faltantes}")

    hoy = (date.today() - date(1899, 12, 30)).days
    filas = 0
    marcas = set()
    desactualizados = 0
    for nombre in nombres_tabla:
        if nombre not in tablas:
            continue
        hoja, ref = tablas[nombre]
        filas_tabla = list(hoja[ref])
        enc = [c.value for c in filas_tabla[0]]
        filas += len(filas_tabla) - 1
        try:
            i_nom = enc.index("NombreProducto")
            i_mar = enc.index("Marca")
            i_fec = enc.index("FechaActualizacion")
        except ValueError:
            avisos.append(f"{nombre} ({hoja.title}) no tiene NombreProducto/Marca/FechaActualizacion.")
            continue
        for fila in filas_tabla[1:]:
            marca = fila[i_mar].value
            serial = _fecha_serial(fila[i_fec].value)
            if not isinstance(marca, str) or serial is None:
                continue
            marcas.add(marca)
            if isinstance(fila[i_nom].value, str) and hoy - serial > MARCAS_DIAS:
                desactualizados += 1

    # 3) Comparación contra la capacidad de cada rango
    medidas = [
        ("Filas de las tablas (J:M)", filas, capacidad.get("J1")),
        ("Marcas con fecha (A2)", len(marcas), capacidad.get("A2")),
        ("Productos desactualizados (E3)", desactualizados, capacidad.get("E3")),
    ]
    for etiqueta, uso, cap in medidas:
        if not cap:
            print(f"  {etiqueta}: {uso} (capacidad desconocida)")
            continue
        pct = uso / cap * 100
        print(f"  {etiqueta}: {uso} de {cap} ({pct:.1f} %)")
        if uso > cap * MARCAS_UMBRAL:
            avisos.append(f"{etiqueta}: {uso} de {cap} ({pct:.1f} %) supera el "
                          f"{MARCAS_UMBRAL:.0%}: ampliar el rango de la hoja Marcas "
                          "(scripts/ampliar_marcas.py, con autorización expresa).")

    if avisos:
        for a in avisos:
            print(f"  AVISO: {a}")
    else:
        print("  Sin avisos.")
    return avisos


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
    controlar_marcas(wb)  # solo lectura: no altera cambios, registro ni libro

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
