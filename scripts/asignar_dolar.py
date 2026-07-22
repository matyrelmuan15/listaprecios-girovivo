#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
asignar_dolar.py
-----------------
Script genérico Tipo 2 del proyecto "Lista de Precios Giro Vivo".

Recalcula "PrecioConIVA" a partir de un nuevo valor de dólar y actualiza
"DolarReferencia" y "FechaActualizacion", según tres modos posibles:

  - marca        : filtra por columna "Marca" (una o varias, separadas por coma)
  - fecha         : filtra por rango de "FechaActualizacion" (dd/mm/aaaa)
  - marca_fecha   : combina ambos filtros (deben cumplirse los dos)

Fórmula aplicada fila por fila:
  nuevo PrecioConIVA = (PrecioConIVA actual / DolarReferencia actual) * PRECIO_ACTUALIZADO

Reglas seguidas (según CLAUDE.md del proyecto):
  - Nunca se modifica el archivo real sin backup previo con timestamp.
  - Columnas localizadas siempre por nombre de encabezado exacto (fila 1 de cada hoja).
  - Se abre el libro SIN data_only=True, para no destruir fórmulas preexistentes.
  - Valores escritos siempre fijos (no fórmulas): precio como número (2 decimales
    máx.), DolarReferencia como número, fecha como texto dd/mm/aaaa.
  - Se excluye siempre la hoja "Marcas Actualización Dinámica".
  - Se recorren todas las hojas del libro dinámicamente.
  - Flujo obligatorio en 2 pasos: primero DRY-RUN (sin --confirmar), que muestra
    el detalle por hoja y un ejemplo de cálculo sin escribir nada; recién con
    --confirmar se aplica el backup y los cambios reales.

USO
---
  1) Vista previa (no escribe nada):
     python asignar_dolar.py --archivo "archivo/Lista de Precios Giro Vivo 2026.xlsx" \
         --modo marca --marca "Bontey" --precio-actualizado 1350

  2) Ejecución real (recién después de revisar la vista previa):
     python asignar_dolar.py --archivo "archivo/Lista de Precios Giro Vivo 2026.xlsx" \
         --modo marca --marca "Bontey" --precio-actualizado 1350 --confirmar

  Modo fecha:
     python asignar_dolar.py --archivo "..." --modo fecha \
         --fecha-desde 01/06/2026 --fecha-hasta 30/06/2026 --precio-actualizado 1350

  Modo marca_fecha:
     python asignar_dolar.py --archivo "..." --modo marca_fecha \
         --marca "Bontey,OtraMarca" --fecha-desde 01/06/2026 --fecha-hasta 30/06/2026 \
         --precio-actualizado 1350
"""

import argparse
import datetime
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

SHEET_EXCLUDE = "Marcas Actualización Dinámica"
SHEET_REGISTRO = "RegistroCambios"

# Encabezados que buscamos en la fila 1 de cada hoja (por nombre exacto, sin
# distinguir mayúsculas/minúsculas ni espacios al principio/final).
COL_PRECIO = "PrecioConIVA"
COL_DOLAR = "DolarReferencia"
COL_MARCA = "Marca"
COL_FECHA = "FechaActualizacion"
COL_NOMBRE_CANDIDATOS = ["NombreProducto", "Producto", "Nombre"]


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------

def parse_fecha(valor):
    """Intenta interpretar 'valor' como fecha dd/mm/aaaa. Devuelve date o None."""
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


def es_vacio_o_cero(valor):
    if valor is None:
        return True
    if isinstance(valor, str):
        texto = valor.strip().replace(",", ".")
        if texto == "":
            return True
        try:
            return float(texto) == 0.0
        except ValueError:
            return True  # no numérico -> lo tratamos como inválido/vacío
    try:
        return float(valor) == 0.0
    except (TypeError, ValueError):
        return True


def a_float(valor):
    if isinstance(valor, str):
        valor = valor.strip().replace(",", ".")
    return float(valor)


def encontrar_columnas(ws, requeridas):
    """Busca encabezados exactos (normalizados) en la fila 1 de la hoja.
    Devuelve dict {nombre_requerido: numero_de_columna} solo con las encontradas."""
    col_map = {}
    nombre_col = None
    for cell in ws[1]:
        if cell.value is None:
            continue
        nombre = str(cell.value).strip()
        for req in requeridas:
            if nombre.lower() == req.lower():
                col_map[req] = cell.column
        for cand in COL_NOMBRE_CANDIDATOS:
            if nombre.lower() == cand.lower():
                nombre_col = cell.column
    return col_map, nombre_col


def normalizar_marca(valor):
    return str(valor).strip() if valor is not None else ""


def hacer_backup(archivo_path):
    archivo_path = Path(archivo_path)
    backup_dir = archivo_path.parent.parent / "backup"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = backup_dir / f"{archivo_path.stem}_{timestamp}{archivo_path.suffix}"
    shutil.copy2(archivo_path, destino)
    return destino


# --------------------------------------------------------------------------
# Lógica principal
# --------------------------------------------------------------------------

def procesar(args, interactivo=False):
    modo = args.modo
    requeridas = [COL_PRECIO, COL_DOLAR, COL_FECHA]
    if modo in ("marca", "marca_fecha"):
        requeridas.append(COL_MARCA)

    marcas_filtro = []
    if args.marca:
        marcas_filtro = [normalizar_marca(m) for m in args.marca.split(",") if m.strip()]

    fecha_desde = fecha_hasta = None
    if modo in ("fecha", "marca_fecha"):
        fecha_desde = parse_fecha(args.fecha_desde)
        fecha_hasta = parse_fecha(args.fecha_hasta)
        if fecha_desde is None or fecha_hasta is None:
            sys.exit("ERROR: --fecha-desde y --fecha-hasta son obligatorias y deben "
                      "tener formato dd/mm/aaaa en modo 'fecha' o 'marca_fecha'.")

    if modo in ("marca", "marca_fecha") and not marcas_filtro:
        sys.exit("ERROR: --marca es obligatoria (una o varias separadas por coma) "
                  "en modo 'marca' o 'marca_fecha'.")

    precio_actualizado = args.precio_actualizado
    hoy_texto = datetime.date.today().strftime("%d/%m/%Y")

    wb = load_workbook(args.archivo, data_only=False)

    reporte = []          # detalle por hoja para el dry-run
    cambios_a_aplicar = []  # lista de dicts con todo lo necesario para escribir + registrar
    ejemplo_global = None

    for ws in wb.worksheets:
        if ws.title == SHEET_EXCLUDE or ws.title == SHEET_REGISTRO:
            continue

        col_map, nombre_col = encontrar_columnas(ws, requeridas)
        if not all(req in col_map for req in requeridas):
            faltantes = [r for r in requeridas if r not in col_map]
            reporte.append({
                "hoja": ws.title, "saltada": True, "faltantes": faltantes,
            })
            continue

        c_precio = col_map[COL_PRECIO]
        c_dolar = col_map[COL_DOLAR]
        c_fecha = col_map[COL_FECHA]
        c_marca = col_map.get(COL_MARCA)

        candidatas = 0
        omitidas_dolar = 0
        omitidas_fecha_invalida = 0
        fuera_de_rango = 0
        ejemplo_hoja = None

        for row_idx in range(2, ws.max_row + 1):
            precio_actual = ws.cell(row=row_idx, column=c_precio).value
            dolar_actual = ws.cell(row=row_idx, column=c_dolar).value
            fecha_actual_raw = ws.cell(row=row_idx, column=c_fecha).value
            marca_actual = ws.cell(row=row_idx, column=c_marca).value if c_marca else None
            nombre_actual = ws.cell(row=row_idx, column=nombre_col).value if nombre_col else ""

            # Saltear filas totalmente vacías en la columna clave de precio
            if precio_actual is None and dolar_actual is None and not marca_actual and not fecha_actual_raw:
                continue

            # --- Filtro de marca ---
            if modo in ("marca", "marca_fecha"):
                if normalizar_marca(marca_actual) not in marcas_filtro:
                    continue

            # --- Filtro de fecha ---
            fecha_valida = None
            if modo in ("fecha", "marca_fecha"):
                fecha_valida = parse_fecha(fecha_actual_raw)
                if fecha_valida is None:
                    omitidas_fecha_invalida += 1
                    continue
                if not (fecha_desde <= fecha_valida <= fecha_hasta):
                    fuera_de_rango += 1
                    continue

            candidatas += 1

            # --- Chequeo DolarReferencia vacía o cero ---
            if es_vacio_o_cero(dolar_actual):
                omitidas_dolar += 1
                continue

            try:
                precio_f = a_float(precio_actual)
                dolar_f = a_float(dolar_actual)
            except (TypeError, ValueError):
                omitidas_dolar += 1
                continue

            nuevo_precio = round((precio_f / dolar_f) * precio_actualizado, 2)

            cambio = {
                "hoja": ws.title,
                "fila": row_idx,
                "nombre_producto": nombre_actual,
                "marca": marca_actual,
                "fecha_anterior": fecha_actual_raw,
                "precio_anterior": precio_f,
                "dolar_anterior": dolar_f,
                "precio_nuevo": nuevo_precio,
                "dolar_nuevo": precio_actualizado,
                "fecha_nueva": hoy_texto,
                "col_precio": c_precio,
                "col_dolar": c_dolar,
                "col_fecha": c_fecha,
            }
            cambios_a_aplicar.append(cambio)

            if ejemplo_hoja is None:
                ejemplo_hoja = cambio
            if ejemplo_global is None:
                ejemplo_global = cambio

        reporte.append({
            "hoja": ws.title,
            "saltada": False,
            "candidatas": candidatas,
            "omitidas_dolar": omitidas_dolar,
            "omitidas_fecha_invalida": omitidas_fecha_invalida,
            "fuera_de_rango": fuera_de_rango,
            "ejemplo": ejemplo_hoja,
        })

    imprimir_reporte(modo, reporte, precio_actualizado, marcas_filtro, fecha_desde, fecha_hasta)

    if not cambios_a_aplicar:
        print("\nNo hay ninguna fila para modificar. No se hizo backup ni se escribió nada.")
        return

    confirmado = args.confirmar
    if not confirmado and interactivo:
        respuesta = input("\n¿Confirmás para aplicar estos cambios? (s/n): ").strip().lower()
        confirmado = respuesta in ("s", "si", "sí", "y", "yes")

    if not confirmado:
        print("\n[DRY-RUN] No se escribió nada. Revisá el detalle de arriba.")
        print("Si está todo correcto, volvé a correr el mismo comando agregando --confirmar.")
        return

    # --- Backup antes de tocar el archivo real ---
    destino_backup = hacer_backup(args.archivo)
    print(f"\nBackup creado en: {destino_backup}")

    # --- Aplicar cambios celda por celda ---
    for cambio in cambios_a_aplicar:
        ws = wb[cambio["hoja"]]
        ws.cell(row=cambio["fila"], column=cambio["col_precio"]).value = cambio["precio_nuevo"]
        ws.cell(row=cambio["fila"], column=cambio["col_dolar"]).value = cambio["dolar_nuevo"]
        ws.cell(row=cambio["fila"], column=cambio["col_fecha"]).value = cambio["fecha_nueva"]

    # --- Registrar en RegistroCambios ---
    ws_reg = wb[SHEET_REGISTRO] if SHEET_REGISTRO in wb.sheetnames else wb.create_sheet(SHEET_REGISTRO)
    if ws_reg.max_row == 1 and ws_reg.cell(row=1, column=1).value is None:
        encabezados = construir_encabezados_registro(modo)
        for i, h in enumerate(encabezados, start=1):
            ws_reg.cell(row=1, column=i).value = h

    fila_libre = ws_reg.max_row + 1
    for cambio in cambios_a_aplicar:
        valores = construir_fila_registro(modo, cambio, precio_actualizado)
        for i, v in enumerate(valores, start=1):
            ws_reg.cell(row=fila_libre, column=i).value = v
        fila_libre += 1

    wb.save(args.archivo)

    print(f"\nCambios aplicados: {len(cambios_a_aplicar)} filas modificadas.")
    print(f"Se registraron {len(cambios_a_aplicar)} filas nuevas en '{SHEET_REGISTRO}'.")
    print("Archivo guardado correctamente.")


def construir_encabezados_registro(modo):
    base_inicio = ["Hoja", "NombreProducto"]
    if modo == "marca":
        medio = ["Marca"]
    elif modo == "fecha":
        medio = ["FechaActualizacion anterior"]
    else:  # marca_fecha
        medio = ["Marca", "FechaActualizacion anterior"]
    fin = ["PrecioConIVA anterior", "DolarReferencia anterior",
           "Multiplicador (PRECIO ACTUALIZADO)", "PrecioConIVA nuevo",
           "DolarReferencia nuevo", "FechaActualizacion aplicada"]
    return base_inicio + medio + fin


def construir_fila_registro(modo, cambio, precio_actualizado):
    base_inicio = [cambio["hoja"], cambio["nombre_producto"]]
    if modo == "marca":
        medio = [cambio["marca"]]
    elif modo == "fecha":
        medio = [cambio["fecha_anterior"]]
    else:
        medio = [cambio["marca"], cambio["fecha_anterior"]]
    fin = [cambio["precio_anterior"], cambio["dolar_anterior"], precio_actualizado,
           cambio["precio_nuevo"], cambio["dolar_nuevo"], cambio["fecha_nueva"]]
    return base_inicio + medio + fin


def imprimir_reporte(modo, reporte, precio_actualizado, marcas_filtro, fecha_desde, fecha_hasta):
    print("=" * 70)
    print(f"MODO: {modo}")
    if marcas_filtro:
        print(f"Marca(s): {', '.join(marcas_filtro)}")
    if fecha_desde:
        print(f"Rango de fechas: {fecha_desde.strftime('%d/%m/%Y')} a {fecha_hasta.strftime('%d/%m/%Y')}")
    print(f"PRECIO ACTUALIZADO (nuevo DolarReferencia): {precio_actualizado}")
    print("=" * 70)

    total_candidatas = 0
    total_omitidas_dolar = 0
    total_omitidas_fecha = 0
    total_fuera_rango = 0

    for item in reporte:
        if item["saltada"]:
            print(f"\n[SALTEADA] Hoja '{item['hoja']}': faltan columnas {item['faltantes']}")
            continue

        print(f"\nHoja '{item['hoja']}':")
        print(f"  Candidatas a modificar: {item['candidatas']}")
        print(f"  Omitidas por DolarReferencia vacía/cero: {item['omitidas_dolar']}")
        if modo in ("fecha", "marca_fecha"):
            print(f"  Omitidas por FechaActualizacion vacía/ilegible: {item['omitidas_fecha_invalida']}")
        if modo == "marca_fecha":
            print(f"  De la(s) marca(s), fuera del rango de fechas: {item['fuera_de_rango']}")

        total_candidatas += item["candidatas"]
        total_omitidas_dolar += item["omitidas_dolar"]
        total_omitidas_fecha += item["omitidas_fecha_invalida"]
        total_fuera_rango += item["fuera_de_rango"]

        if item["ejemplo"]:
            e = item["ejemplo"]
            print("  Ejemplo de cálculo:")
            print(f"    PrecioConIVA anterior: {e['precio_anterior']}")
            print(f"    DolarReferencia actual (divisor): {e['dolar_anterior']}")
            print(f"    Multiplicador (PRECIO ACTUALIZADO): {precio_actualizado}")
            print(f"    -> PrecioConIVA nuevo: {e['precio_nuevo']}")
            print(f"    -> DolarReferencia nuevo: {e['dolar_nuevo']}")
            if modo in ("fecha", "marca_fecha"):
                print(f"    FechaActualizacion anterior: {e['fecha_anterior']}")
            print(f"    -> FechaActualizacion nueva: {e['fecha_nueva']}")

    print("\n" + "-" * 70)
    print("TOTALES")
    print(f"  Filas que se van a modificar: {total_candidatas - total_omitidas_dolar - total_omitidas_fecha}")
    print(f"  Omitidas por DolarReferencia vacía/cero: {total_omitidas_dolar}")
    if modo in ("fecha", "marca_fecha"):
        print(f"  Omitidas por FechaActualizacion vacía/ilegible: {total_omitidas_fecha}")
    if modo == "marca_fecha":
        print(f"  Fuera del rango de fechas: {total_fuera_rango}")
    print("-" * 70)


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
        print("  Formato inválido. Usá dd/mm/aaaa (ej: 04/06/2026).")


def pedir_precio(mensaje):
    while True:
        valor = input(mensaje).strip().replace(",", ".")
        try:
            return float(valor)
        except ValueError:
            print("  Ingresá un número válido (ej: 1350 o 1350.50).")


def completar_interactivamente(args):
    """Completa los campos faltantes de 'args' preguntando por teclado.
    Devuelve True si el modo es interactivo (para pedir confirmación al final
    en la misma corrida, sin necesidad de --confirmar)."""
    interactivo = args.modo is None

    if not interactivo:
        return False

    print("=" * 70)
    print("ASIGNAR VALOR DÓLAR - Tipo 2")
    print("=" * 70)

    if not args.archivo:
        args.archivo = pedir_no_vacio("Ruta del archivo Excel: ")

    print("\n¿Qué tipo de actualización querés hacer?")
    print("  1) Por marca")
    print("  2) Por fecha")
    print("  3) Por marca y fecha al mismo tiempo")
    opcion = ""
    while opcion not in ("1", "2", "3"):
        opcion = input("Elegí una opción (1/2/3): ").strip()
    args.modo = {"1": "marca", "2": "fecha", "3": "marca_fecha"}[opcion]

    if args.modo in ("marca", "marca_fecha") and not args.marca:
        args.marca = pedir_no_vacio(
            "Marca(s) (si son varias, separalas con coma): ")

    if args.modo in ("fecha", "marca_fecha"):
        if not args.fecha_desde:
            args.fecha_desde = pedir_fecha("Fecha desde (dd/mm/aaaa): ")
        if not args.fecha_hasta:
            args.fecha_hasta = pedir_fecha("Fecha hasta (dd/mm/aaaa): ")

    if args.precio_actualizado is None:
        args.precio_actualizado = pedir_precio(
            "Precio actualizado (nuevo valor de dólar): ")

    print()
    return True


def main():
    parser = argparse.ArgumentParser(description="Tipo 2 - Asignar valor dólar")
    parser.add_argument("--archivo", help="Ruta al libro Excel")
    parser.add_argument("--modo", choices=["marca", "fecha", "marca_fecha"],
                         help="Si se omite, el script pregunta interactivamente.")
    parser.add_argument("--marca", help="Marca o marcas separadas por coma (requerido en modo marca/marca_fecha)")
    parser.add_argument("--fecha-desde", help="dd/mm/aaaa (requerido en modo fecha/marca_fecha)")
    parser.add_argument("--fecha-hasta", help="dd/mm/aaaa (requerido en modo fecha/marca_fecha)")
    parser.add_argument("--precio-actualizado", type=float,
                         help="Nuevo valor de dólar (multiplicador)")
    parser.add_argument("--confirmar", action="store_true",
                         help="Aplica los cambios de verdad. Sin este flag, solo muestra el dry-run "
                              "(salvo en modo interactivo, que pregunta al final).")
    args = parser.parse_args()

    interactivo = completar_interactivamente(args)

    # Validaciones básicas también para uso por flags (mensajes claros en vez
    # de que argparse tire error genérico de "required").
    if not args.archivo:
        sys.exit("ERROR: falta --archivo")
    if not args.modo:
        sys.exit("ERROR: falta --modo (marca | fecha | marca_fecha)")
    if args.precio_actualizado is None:
        sys.exit("ERROR: falta --precio-actualizado")
    if not Path(args.archivo).exists():
        sys.exit(f"ERROR: no se encontró el archivo '{args.archivo}'")

    procesar(args, interactivo=interactivo)


if __name__ == "__main__":
    main()