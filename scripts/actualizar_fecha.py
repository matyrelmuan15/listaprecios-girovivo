#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizar_fecha.py
--------------------
Script genérico Tipo 3 del proyecto "Lista de Precios Giro Vivo".

Actualiza ÚNICAMENTE la columna "FechaActualizacion" para una o varias
marcas, sin tocar ningún precio (PrecioConIVA, PrecioContado, PrecioSinIVA)
ni DolarReferencia. Pensado para dejar constancia de que la lista de una
marca fue revisada este mes, aunque el proveedor no haya cambiado precios.

Fecha a escribir (dos modos, uno u otro):
  --hoy               : usa la fecha de hoy automáticamente.
  --fecha dd/mm/aaaa  : usa una fecha específica elegida por el usuario.

Reglas seguidas (según CLAUDE.md del proyecto):
  - Nunca se modifica el archivo real sin backup previo con timestamp.
  - Columnas localizadas siempre por nombre de encabezado exacto (fila 1).
  - Se abre el libro SIN data_only=True, para no destruir fórmulas.
  - Fecha escrita siempre como texto fijo dd/mm/aaaa (no fórmula).
  - Se excluye siempre la hoja "Marcas Actualización Dinámica", además de
    "Hoja1" y "RegistroCambios".
  - Se recorren todas las hojas del libro dinámicamente.
  - Flujo obligatorio en 2 pasos: primero DRY-RUN (sin --execute), que
    muestra el detalle por hoja y un ejemplo sin escribir nada; recién con
    --execute se aplica el backup y los cambios reales.
  - Cada celda modificada se registra en RegistroCambios (Columna =
    "FechaActualizacion"), con verificación de conteo obligatoria.

USO
---
  Modo por flags, fecha de hoy:
    python actualizar_fecha.py --marca "Bontey" --hoy

  Modo por flags, fecha específica:
    python actualizar_fecha.py --marca "Bontey,Giant" --fecha 15/08/2026

  Ejecución real (recién después de revisar el dry-run):
    python actualizar_fecha.py --marca "Bontey" --hoy --execute

  Modo interactivo (sin argumentos, pregunta todo por teclado):
    python actualizar_fecha.py
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

HOJA_EXCLUIDA = "Marcas Actualización Dinámica"
HOJA_ENTRADA = "Hoja1"
# El formato de RegistroCambios es único y compartido: vive en registro.py.
HOJA_REGISTRO = registro.HOJA_REGISTRO

COL_MARCA = "Marca"
COL_FECHA = "FechaActualizacion"
COL_NOMBRE_CANDIDATOS = ["NombreProducto", "Producto", "Nombre"]


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def parse_fecha(valor):
    if valor is None:
        return None
    if isinstance(valor, datetime.datetime):
        return valor.date()
    if isinstance(valor, datetime.date):
        return valor
    texto = str(valor).strip()
    if not texto:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def normalizar_marca(valor):
    return str(valor).strip() if valor is not None else ""


def encontrar_columnas(ws, requeridas, max_search_rows=3):
    """Busca encabezados exactos (normalizados) en las primeras filas de la
    hoja. Devuelve (dict {nombre_requerido: numero_de_columna}, col_codigo,
    col_nombre)."""
    col_map = {}
    nombre_col = None
    codigo_col = None
    for row in ws.iter_rows(min_row=1, max_row=max_search_rows):
        for cell in row:
            if cell.value is None:
                continue
            nombre = str(cell.value).strip()
            for req in requeridas:
                if nombre.casefold() == req.casefold():
                    col_map[req] = cell.column
            for cand in COL_NOMBRE_CANDIDATOS:
                if nombre.casefold() == cand.casefold():
                    nombre_col = cell.column
            if nombre.casefold() in ("codigofactura", "codigofac"):
                codigo_col = cell.column
    return col_map, codigo_col, nombre_col


def listar_marcas():
    """Recorre todas las hojas de destino y devuelve la lista ordenada de
    marcas únicas encontradas en la columna 'Marca'. Pensado para que la
    interfaz gráfica arme la lista de selección sin hardcodear nombres.
    Cualquier problema para abrir el archivo (incluido tenerlo abierto en
    Excel) se informa por stderr con código de salida distinto de cero."""
    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}", file=sys.stderr)
        sys.exit(1)

    try:
        wb = load_workbook(ARCHIVO_PATH, data_only=False, read_only=True)
    except PermissionError:
        print(
            "ERROR: no se pudo leer el archivo. ¿Está abierto en Excel? Cerralo e intentá de nuevo.",
            file=sys.stderr,
        )
        sys.exit(1)
    except Exception as e:
        print(f"ERROR al leer el archivo: {e}", file=sys.stderr)
        sys.exit(1)

    marcas = set()
    for ws in wb.worksheets:
        if ws.title in (HOJA_ENTRADA, HOJA_EXCLUIDA, HOJA_REGISTRO, "NoActualizados", "Venzo Sin Actualizar"):
            continue
        col_map, _codigo, _nombre = encontrar_columnas(ws, [COL_MARCA])
        if COL_MARCA not in col_map:
            continue
        c_marca = col_map[COL_MARCA]
        for fila in ws.iter_rows(min_row=2, min_col=c_marca, max_col=c_marca,
                                  values_only=True):
            valor = normalizar_marca(fila[0])
            if valor:
                marcas.add(valor)

    wb.close()
    return sorted(marcas, key=lambda m: m.casefold())


def backup_archivo():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{ARCHIVO_PATH.stem}_{timestamp}{ARCHIVO_PATH.suffix}"
    shutil.copy2(ARCHIVO_PATH, destino)
    return destino


# --------------------------------------------------------------------------
# Lógica principal
# --------------------------------------------------------------------------
def procesar(marcas_filtro, fecha_nueva_texto, execute, interactivo=False):
    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")
        sys.exit(1)

    if parse_fecha(fecha_nueva_texto) is None:
        print(f"ERROR: fecha inválida '{fecha_nueva_texto}'. Usá formato dd/mm/aaaa.")
        sys.exit(1)

    wb = load_workbook(ARCHIVO_PATH, data_only=False)

    # Validación fail-fast del formato de RegistroCambios: se corre apenas
    # cargado el libro y ANTES del backup, para no dejar un backup inútil ni
    # apilar filas desalineadas bajo encabezados que no corresponden.
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)

    requeridas = [COL_MARCA, COL_FECHA]
    reporte = []
    cambios_a_aplicar = []
    ejemplo_global = None

    for ws in wb.worksheets:
        if ws.title in (HOJA_ENTRADA, HOJA_EXCLUIDA, HOJA_REGISTRO, "NoActualizados", "Venzo Sin Actualizar"):
            continue

        col_map, col_codigo, nombre_col = encontrar_columnas(ws, requeridas)
        if not all(req in col_map for req in requeridas):
            faltantes = [r for r in requeridas if r not in col_map]
            reporte.append({"hoja": ws.title, "saltada": True, "faltantes": faltantes})
            continue

        c_marca = col_map[COL_MARCA]
        c_fecha = col_map[COL_FECHA]

        candidatas = 0
        ya_actualizadas = 0
        ejemplo_hoja = None

        for row_idx in range(2, ws.max_row + 1):
            marca_actual = ws.cell(row=row_idx, column=c_marca).value
            if normalizar_marca(marca_actual) not in marcas_filtro:
                continue

            fecha_actual_raw = ws.cell(row=row_idx, column=c_fecha).value
            nombre_actual = ws.cell(row=row_idx, column=nombre_col).value if nombre_col else ""
            codigo_actual = ws.cell(row=row_idx, column=col_codigo).value if col_codigo else ""

            # Si ya tiene exactamente esa fecha como texto, no hace falta tocarla.
            if isinstance(fecha_actual_raw, str) and fecha_actual_raw.strip() == fecha_nueva_texto:
                ya_actualizadas += 1
                continue

            candidatas += 1
            cambio = {
                "hoja": ws.title,
                "fila": row_idx,
                "codigo": codigo_actual,
                "marca": marca_actual,
                "nombre_producto": nombre_actual,
                "fecha_anterior": fecha_actual_raw,
                "fecha_nueva": fecha_nueva_texto,
                "col_fecha": c_fecha,
            }
            cambios_a_aplicar.append(cambio)
            if ejemplo_hoja is None:
                ejemplo_hoja = cambio
            if ejemplo_global is None:
                ejemplo_global = cambio

        reporte.append({
            "hoja": ws.title, "saltada": False,
            "candidatas": candidatas, "ya_actualizadas": ya_actualizadas,
            "ejemplo": ejemplo_hoja,
        })

    imprimir_reporte(reporte, marcas_filtro, fecha_nueva_texto)

    if not cambios_a_aplicar:
        print("\nNo hay ninguna fila para modificar. No se hizo backup ni se escribió nada.")
        return

    # --- Filas de registro: una por cada CELDA modificada (regla 8) ---
    # En Tipo 3 es una sola por producto: FechaActualizacion.
    corrida = registro.RegistroCorrida(
        registro.TIPO_FECHA,
        f"Marcas: {', '.join(marcas_filtro)} | Fecha aplicada: {fecha_nueva_texto}",
    )
    for cambio in cambios_a_aplicar:
        corrida.agregar(
            hoja=cambio["hoja"], fila=cambio["fila"],
            nombre_producto=cambio["nombre_producto"], marca=cambio["marca"],
            columna=COL_FECHA, valor_anterior=cambio["fecha_anterior"],
            valor_nuevo=cambio["fecha_nueva"],
        )

    # Se muestra antes de pedir la confirmación, para poder revisar cómo
    # quedaría el registro antes de decidir.
    registro.imprimir_previsualizacion(corrida)

    confirmado = execute
    if not confirmado and interactivo:
        respuesta = input("\n¿Confirmás para aplicar estos cambios? (s/n): ").strip().lower()
        confirmado = respuesta in ("s", "si", "sí", "y", "yes")

    if not confirmado:
        print("\n[DRY-RUN] No se escribió nada. Revisá el detalle de arriba.")
        print("Si está todo correcto, volvé a correr el mismo comando agregando --execute.")
        return

    # --- Backup antes de tocar el archivo real ---
    destino_backup = backup_archivo()
    print(f"\nBackup creado en: {destino_backup}")

    # --- Aplicar cambios celda por celda (solo FechaActualizacion) ---
    for cambio in cambios_a_aplicar:
        ws = wb[cambio["hoja"]]
        ws.cell(row=cambio["fila"], column=cambio["col_fecha"]).value = cambio["fecha_nueva"]

    # --- Registrar en RegistroCambios ---
    # Verificación de conteo obligatoria (regla 8): 1 celda modificada por
    # match. Si no coincide se aborta sin guardar.
    try:
        corrida.verificar_conteo(len(cambios_a_aplicar))
        corrida.escribir(wb)
    except registro.RegistroError as e:
        print(f"\n{e}")
        sys.exit(1)

    wb.save(ARCHIVO_PATH)

    print(f"\nCambios aplicados: {len(cambios_a_aplicar)} filas modificadas.")
    print(f"Se registraron {len(corrida)} filas nuevas en '{HOJA_REGISTRO}'.")
    print("Archivo guardado correctamente.")


def imprimir_reporte(reporte, marcas_filtro, fecha_nueva_texto):
    print("=" * 70)
    print("ACTUALIZAR FECHA - Tipo 3 (sin tocar precios)")
    print(f"Marca(s): {', '.join(marcas_filtro)}")
    print(f"Fecha a asignar: {fecha_nueva_texto}")
    print("=" * 70)

    total_candidatas = 0
    total_ya_actualizadas = 0

    for item in reporte:
        if item["saltada"]:
            print(f"\n[SALTEADA] Hoja '{item['hoja']}': faltan columnas {item['faltantes']}")
            continue
        if item["candidatas"] == 0 and item["ya_actualizadas"] == 0:
            continue

        print(f"\nHoja '{item['hoja']}':")
        print(f"  Filas a actualizar: {item['candidatas']}")
        print(f"  Ya tenían esa fecha (sin cambios): {item['ya_actualizadas']}")
        total_candidatas += item["candidatas"]
        total_ya_actualizadas += item["ya_actualizadas"]

        if item["ejemplo"]:
            e = item["ejemplo"]
            print("  Ejemplo:")
            print(f"    Producto: {e['nombre_producto']!r} (código {e['codigo']!r})")
            print(f"    FechaActualizacion anterior: {e['fecha_anterior']!r}")
            print(f"    -> FechaActualizacion nueva: {e['fecha_nueva']}")

    print("\n" + "-" * 70)
    print("TOTALES")
    print(f"  Filas que se van a modificar: {total_candidatas}")
    print(f"  Filas ya al día (sin cambios): {total_ya_actualizadas}")
    print("-" * 70)


# --------------------------------------------------------------------------
# Modo interactivo
# --------------------------------------------------------------------------
def pedir_no_vacio(mensaje):
    while True:
        valor = input(mensaje).strip()
        if valor:
            return valor
        print("  (no puede quedar vacío)")


def pedir_fecha(mensaje):
    while True:
        valor = input(mensaje).strip()
        if parse_fecha(valor) is not None:
            return valor
        print("  Formato inválido. Usá dd/mm/aaaa (ej: 15/08/2026).")


def modo_interactivo():
    print("=" * 70)
    print("ACTUALIZAR FECHA - Tipo 3")
    print("=" * 70)

    marcas_texto = pedir_no_vacio("Marca(s) a actualizar (si son varias, separalas con coma): ")
    marcas_filtro = [normalizar_marca(m) for m in marcas_texto.split(",") if m.strip()]

    print("\n¿Qué fecha querés asignar?")
    print("  1) La fecha de hoy")
    print("  2) Una fecha específica")
    opcion = ""
    while opcion not in ("1", "2"):
        opcion = input("Elegí una opción (1/2): ").strip()

    if opcion == "1":
        fecha_nueva_texto = datetime.date.today().strftime("%d/%m/%Y")
    else:
        fecha_nueva_texto = pedir_fecha("Fecha a asignar (dd/mm/aaaa): ")

    print()
    return marcas_filtro, fecha_nueva_texto


def main():
    parser = argparse.ArgumentParser(
        description="Actualiza solo FechaActualizacion (Tipo 3) para una o varias marcas."
    )
    parser.add_argument("--marca", help="Marca o marcas separadas por coma")
    fecha_group = parser.add_mutually_exclusive_group()
    fecha_group.add_argument("--hoy", action="store_true", help="Usa la fecha de hoy")
    fecha_group.add_argument("--fecha", help="Fecha específica dd/mm/aaaa")
    parser.add_argument(
        "--execute", action="store_true",
        help="Ejecuta los cambios reales (por defecto corre en modo dry-run).",
    )
    parser.add_argument(
        "--listar-marcas", action="store_true",
        help="Imprime una marca por línea y sale. Uso interno de la interfaz gráfica.",
    )
    args = parser.parse_args()

    if args.listar_marcas:
        for marca in listar_marcas():
            print(marca)
        return

    interactivo = False
    if not args.marca or (not args.hoy and not args.fecha):
        marcas_filtro, fecha_nueva_texto = modo_interactivo()
        interactivo = True
    else:
        marcas_filtro = [normalizar_marca(m) for m in args.marca.split(",") if m.strip()]
        fecha_nueva_texto = (
            datetime.date.today().strftime("%d/%m/%Y") if args.hoy else args.fecha
        )

    if not marcas_filtro:
        sys.exit("ERROR: falta indicar al menos una marca.")

    procesar(marcas_filtro, fecha_nueva_texto, args.execute, interactivo=interactivo)


if __name__ == "__main__":
    main()
