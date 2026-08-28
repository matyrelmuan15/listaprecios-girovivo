#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
asignar_dolar.py
-----------------
Script genérico Tipo 2 del proyecto "Lista de Precios Giro Vivo".

Recalcula "PrecioConIVA" a partir de un nuevo valor de dólar y actualiza
"DolarReferencia" y "FechaActualizacion", según tres modos posibles:

  - marca        : filtra por columna "Marca" (una o varias)
  - fecha         : filtra por rango de "FechaActualizacion" (dd/mm/aaaa)
  - marca_fecha   : combina ambos filtros (deben cumplirse los dos)

Fórmula aplicada fila por fila:
  nuevo PrecioConIVA = (PrecioConIVA actual / DolarReferencia actual) * PRECIO_ACTUALIZADO

MODO GUIADO (recomendado)
--------------------------
Si lo corrés sin ningún argumento:

    python asignar_dolar.py

el script:
  1) Te pregunta el tipo de actualización con un menú de flechas.
  2) Si corresponde, te muestra las marcas que existen realmente en el
     archivo (leídas del Excel) para que tildes una o varias con la barra
     espaciadora, en vez de tipear el nombre exacto.
  3) Te pide fecha(s) y el nuevo valor de dólar por teclado (con validación).
  4) Corre el dry-run, te muestra el resumen, y al final te pregunta con
     un menú de flechas si confirmás aplicar los cambios reales.

Reglas seguidas (según CLAUDE.md del proyecto):
  - Nunca se modifica el archivo real sin backup previo con timestamp.
  - Columnas localizadas siempre por nombre de encabezado exacto (fila 1 de cada hoja).
  - Se abre el libro SIN data_only=True para escribir, para no destruir fórmulas
    preexistentes.
  - Valores escritos siempre fijos (no fórmulas): precio como número (2 decimales
    máx.), DolarReferencia como número, fecha como texto dd/mm/aaaa.
  - Se excluye siempre la hoja "Marcas Actualización Dinámica".
  - Se recorren todas las hojas del libro dinámicamente.
  - Flujo obligatorio en 2 pasos: primero DRY-RUN (sin --confirmar), que muestra
    el detalle por hoja y un ejemplo de cálculo sin escribir nada; recién con
    --confirmar (o confirmando en el menú guiado) se aplica el backup y los
    cambios reales.

USO POR LÍNEA DE COMANDOS (para automatizar o repetir una corrida exacta)
----------------------------------------------------------------------------
  1) Vista previa (no escribe nada):
     python asignar_dolar.py --modo marca --marca "Bontey" --precio-actualizado 1350

  2) Ejecución real (recién después de revisar la vista previa):
     python asignar_dolar.py --modo marca --marca "Bontey" --precio-actualizado 1350 --confirmar

  Modo fecha:
     python asignar_dolar.py --modo fecha \
         --fecha-desde 01/06/2026 --fecha-hasta 30/06/2026 --precio-actualizado 1350

  Modo marca_fecha:
     python asignar_dolar.py --modo marca_fecha \
         --marca "Bontey,OtraMarca" --fecha-desde 01/06/2026 --fecha-hasta 30/06/2026 \
         --precio-actualizado 1350
"""

import argparse
import datetime
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

import registro
from ia_revision import (
    formatear_items_marcados,
    imprimir_items_marcados_json,
    revisar_cambios_con_ia,
)

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"

SHEET_EXCLUDE = "Marcas Actualización Dinámica"
# El formato de RegistroCambios es único y compartido: vive en registro.py.
SHEET_REGISTRO = registro.HOJA_REGISTRO

# Encabezados que buscamos en la fila 1 de cada hoja (por nombre exacto, sin
# distinguir mayúsculas/minúsculas ni espacios al principio/final).
COL_PRECIO = "PrecioConIVA"
COL_DOLAR = "DolarReferencia"
COL_MARCA = "Marca"
COL_FECHA = "FechaActualizacion"
COL_NOMBRE_CANDIDATOS = ["NombreProducto", "Producto", "Nombre"]


# --------------------------------------------------------------------------
# Menús interactivos por consola (flechas + Enter / Espacio) — solo Windows
# --------------------------------------------------------------------------
def menu_seleccionar(opciones, titulo="Seleccioná una opción"):
    """Menú de selección única navegable con flechas ↑ / ↓ y Enter.
    Devuelve el índice elegido, o None si se cancela con Esc."""
    try:
        import msvcrt
    except ImportError:
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
        if tecla in (b"\xe0", b"\x00"):
            tecla2 = msvcrt.getch()
            if tecla2 == b"H":
                seleccion = (seleccion - 1) % len(opciones)
                dibujar()
            elif tecla2 == b"P":
                seleccion = (seleccion + 1) % len(opciones)
                dibujar()
        elif tecla == b"\r":
            return seleccion
        elif tecla == b"\x1b":
            return None


def menu_multi_seleccionar(opciones, titulo="Marcá una o más opciones"):
    """Menú de selección múltiple (checkboxes) navegable con flechas,
    Espacio para tildar/destildar y Enter para confirmar.
    Devuelve una lista de índices elegidos (ordenada), o None si se cancela."""
    try:
        import msvcrt
    except ImportError:
        return _menu_multi_numerico(opciones, titulo)

    seleccion = 0
    marcadas = set()

    def dibujar(mensaje_extra=""):
        print("\n" + titulo)
        print("=" * len(titulo))
        for i, op in enumerate(opciones):
            check = "[x]" if i in marcadas else "[ ]"
            cursor = ">" if i == seleccion else " "
            print(f" {cursor} {check} {op}")
        print(
            "\n(Flechas arriba/abajo, Espacio para tildar/destildar, "
            "Enter para confirmar, Esc para cancelar)"
        )
        if mensaje_extra:
            print(mensaje_extra)

    dibujar()
    while True:
        tecla = msvcrt.getch()
        if tecla in (b"\xe0", b"\x00"):
            tecla2 = msvcrt.getch()
            if tecla2 == b"H":
                seleccion = (seleccion - 1) % len(opciones)
                dibujar()
            elif tecla2 == b"P":
                seleccion = (seleccion + 1) % len(opciones)
                dibujar()
        elif tecla == b" ":
            if seleccion in marcadas:
                marcadas.discard(seleccion)
            else:
                marcadas.add(seleccion)
            dibujar()
        elif tecla == b"\r":
            if not marcadas:
                dibujar("  Tenés que tildar al menos una marca con la barra espaciadora.")
                continue
            return sorted(marcadas)
        elif tecla == b"\x1b":
            return None


def _menu_seleccionar_numerico(opciones, titulo):
    print("\n" + titulo)
    print("=" * len(titulo))
    for i, op in enumerate(opciones, start=1):
        print(f"  {i}) {op}")
    while True:
        valor = input("Elegí un número: ").strip()
        if valor.isdigit() and 1 <= int(valor) <= len(opciones):
            return int(valor) - 1
        print("  Opción inválida, intentá de nuevo.")


def _menu_multi_numerico(opciones, titulo):
    print("\n" + titulo)
    print("=" * len(titulo))
    for i, op in enumerate(opciones, start=1):
        print(f"  {i}) {op}")
    while True:
        valor = input("Elegí uno o más números separados por coma: ").strip()
        partes = [p.strip() for p in valor.split(",") if p.strip()]
        if partes and all(p.isdigit() and 1 <= int(p) <= len(opciones) for p in partes):
            return sorted({int(p) - 1 for p in partes})
        print("  Entrada inválida, intentá de nuevo.")


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


def listar_marcas(archivo_path):
    """Lee el Excel en modo solo lectura y devuelve la lista ordenada de
    marcas distintas presentes en la columna 'Marca' de cualquier hoja
    (excepto la excluida y RegistroCambios). No modifica ni guarda nada."""
    wb = load_workbook(archivo_path, data_only=True, read_only=True)
    marcas = set()
    try:
        for ws in wb.worksheets:
            if ws.title in (SHEET_EXCLUDE, SHEET_REGISTRO):
                continue
            col_marca = None
            primera_fila = next(ws.iter_rows(min_row=1, max_row=1), None)
            if primera_fila is None:
                continue
            for cell in primera_fila:
                if cell.value is not None and str(cell.value).strip().lower() == COL_MARCA.lower():
                    col_marca = cell.column
                    break
            if col_marca is None:
                continue
            for row in ws.iter_rows(min_row=2):
                if col_marca - 1 >= len(row):
                    continue
                valor = row[col_marca - 1].value
                if valor is not None and str(valor).strip():
                    marcas.add(str(valor).strip())
    finally:
        wb.close()
    return sorted(marcas)


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

    # Validación fail-fast del formato de RegistroCambios: se corre apenas
    # cargado el libro y ANTES del backup, para no dejar un backup inútil ni
    # apilar filas desalineadas bajo encabezados que no corresponden.
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        sys.exit(f"\nERROR: {e}")

    reporte = []          # detalle por hoja para el dry-run
    cambios_a_aplicar = []  # lista de dicts con todo lo necesario para escribir + registrar
    ejemplo_global = None

    for ws in wb.worksheets:
        if ws.title == SHEET_EXCLUDE or ws.title == SHEET_REGISTRO:
            continue

        # "Marca" se localiza siempre, aunque el modo no filtre por ella:
        # el registro necesita la marca de la fila en los tres modos.
        a_buscar = list(dict.fromkeys(requeridas + [COL_MARCA]))
        col_map, nombre_col = encontrar_columnas(ws, a_buscar)
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

    # --- Filas de registro: una por cada CELDA modificada (regla 8) ---
    # En Tipo 2 son 3 por producto: PrecioConIVA, DolarReferencia y
    # FechaActualizacion. Se arman también en dry-run, para poder mostrarlas.
    corrida = registro.RegistroCorrida(
        registro.TIPO_DOLAR,
        construir_detalle(modo, marcas_filtro, fecha_desde, fecha_hasta, precio_actualizado),
    )
    for cambio in cambios_a_aplicar:
        comunes = {
            "hoja": cambio["hoja"],
            "fila": cambio["fila"],
            "nombre_producto": cambio["nombre_producto"],
            "marca": cambio["marca"],
        }
        corrida.agregar(columna=COL_PRECIO,
                        valor_anterior=cambio["precio_anterior"],
                        valor_nuevo=cambio["precio_nuevo"], **comunes)
        corrida.agregar(columna=COL_DOLAR,
                        valor_anterior=cambio["dolar_anterior"],
                        valor_nuevo=cambio["dolar_nuevo"], **comunes)
        corrida.agregar(columna=COL_FECHA,
                        valor_anterior=cambio["fecha_anterior"],
                        valor_nuevo=cambio["fecha_nueva"], **comunes)

    if not cambios_a_aplicar:
        print("\nNo hay ninguna fila para modificar. No se hizo backup ni se escribió nada.")
        imprimir_items_marcados_json([])
        return

    # ---- Revisión automática de posibles anomalías (estadística + IA) ----
    cambios_para_ia = [
        {
            "hoja": c["hoja"],
            "codigo": c["fila"],
            "producto": c["nombre_producto"],
            "valor_anterior": c["precio_anterior"],
            "valor_nuevo": c["precio_nuevo"],
        }
        for c in cambios_a_aplicar
    ]
    print("\n--- Revisión automática de posibles anomalías ---")
    analisis_ia, aviso_ia = revisar_cambios_con_ia(
        cambios_para_ia,
        contexto_extra=(
            f"Asignación de nuevo valor de dólar (Tipo 2). Modo: {modo}. "
            f"Nuevo valor de dólar (DolarReferencia): {precio_actualizado}."
        ),
    )
    if aviso_ia:
        print(aviso_ia)
    if analisis_ia:
        print(analisis_ia)
    detalle_marcados = formatear_items_marcados(cambios_para_ia)
    if detalle_marcados:
        print(detalle_marcados)
    imprimir_items_marcados_json(cambios_para_ia)

    # Se muestra antes de pedir la confirmación, para que en el modo guiado
    # se pueda revisar cómo quedaría el registro antes de decidir.
    registro.imprimir_previsualizacion(corrida)

    confirmado = args.confirmar
    if not confirmado and interactivo:
        idx_confirm = menu_seleccionar(
            ["Sí, aplicar estos cambios reales", "No, salir sin modificar nada"],
            "¿Confirmás aplicar estos cambios al archivo real?",
        )
        confirmado = (idx_confirm == 0)

    if not confirmado:
        print("\n[DRY-RUN] No se escribió nada. Revisá el detalle de arriba.")
        if not interactivo:
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
    # Verificación de conteo obligatoria (regla 8): 3 celdas por producto
    # (PrecioConIVA + DolarReferencia + FechaActualizacion). Si no coincide
    # se aborta sin guardar, así que el archivo real queda intacto.
    try:
        corrida.verificar_conteo(len(cambios_a_aplicar) * 3)
        corrida.escribir(wb)
    except registro.RegistroError as e:
        sys.exit(f"\n{e}")

    wb.save(args.archivo)

    print(f"\nCambios aplicados: {len(cambios_a_aplicar)} filas modificadas.")
    print(f"Se registraron {len(corrida)} filas nuevas en '{SHEET_REGISTRO}'.")
    print("Archivo guardado correctamente.")


def construir_detalle(modo, marcas_filtro, fecha_desde, fecha_hasta, precio_actualizado):
    """Contexto de la corrida para la columna Detalle del registro."""
    partes = [f"Modo: {modo}"]
    partes.append(
        "Marcas: " + (", ".join(marcas_filtro) if marcas_filtro else "(no aplica)")
    )
    if fecha_desde and fecha_hasta:
        partes.append(
            f"Rango: {fecha_desde.strftime('%d/%m/%Y')} a "
            f"{fecha_hasta.strftime('%d/%m/%Y')}"
        )
    else:
        partes.append("Rango: (no aplica)")
    partes.append(f"Nuevo DolarReferencia: {precio_actualizado}")
    return " | ".join(partes)


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


class _ArgsGuiados:
    """Objeto simple para reutilizar procesar() desde el modo guiado."""
    pass


def modo_guiado():
    """Flujo guiado 100% por menú (modo por flechas, marcas por casilleros,
    fecha/precio por teclado con validación) reutilizando procesar()."""
    if not ARCHIVO_PATH.exists():
        sys.exit(f"ERROR: no se encontró el archivo en '{ARCHIVO_PATH}'.")

    print("=" * 70)
    print("ASIGNAR VALOR DÓLAR - Tipo 2")
    print("=" * 70)

    opciones_modo = ["Por marca", "Por fecha", "Por marca y fecha al mismo tiempo"]
    idx_modo = menu_seleccionar(opciones_modo, "¿Qué tipo de actualización querés hacer?")
    if idx_modo is None:
        print("\nOperación cancelada. No se modificó nada.")
        return
    modo = ["marca", "fecha", "marca_fecha"][idx_modo]

    marca_texto = None
    if modo in ("marca", "marca_fecha"):
        marcas_disponibles = listar_marcas(ARCHIVO_PATH)
        if not marcas_disponibles:
            sys.exit("ERROR: no se encontraron marcas en el archivo (columna 'Marca' vacía o ausente).")
        idxs = menu_multi_seleccionar(
            marcas_disponibles, "Elegí una o más marcas (Espacio para tildar, Enter para confirmar)"
        )
        if idxs is None:
            print("\nOperación cancelada. No se modificó nada.")
            return
        marca_texto = ",".join(marcas_disponibles[i] for i in idxs)

    fecha_desde = fecha_hasta = None
    if modo in ("fecha", "marca_fecha"):
        print()
        fecha_desde = pedir_fecha("Fecha desde (dd/mm/aaaa): ")
        fecha_hasta = pedir_fecha("Fecha hasta (dd/mm/aaaa): ")

    print()
    precio_actualizado = pedir_precio("Precio actualizado (nuevo valor de dólar): ")

    args = _ArgsGuiados()
    args.archivo = str(ARCHIVO_PATH)
    args.modo = modo
    args.marca = marca_texto
    args.fecha_desde = fecha_desde
    args.fecha_hasta = fecha_hasta
    args.precio_actualizado = precio_actualizado
    args.confirmar = False

    procesar(args, interactivo=True)


def main():
    parser = argparse.ArgumentParser(description="Tipo 2 - Asignar valor dólar")
    parser.add_argument(
        "--archivo",
        default=None,
        help=(
            "Ruta al libro Excel. Si se omite, se usa el archivo único del "
            f"proyecto: {ARCHIVO_PATH}"
        ),
    )
    parser.add_argument("--modo", choices=["marca", "fecha", "marca_fecha"],
                         help="Si se omite junto con el resto de los flags, se abre el menú guiado.")
    parser.add_argument("--marca", help="Marca o marcas separadas por coma (requerido en modo marca/marca_fecha)")
    parser.add_argument("--fecha-desde", help="dd/mm/aaaa (requerido en modo fecha/marca_fecha)")
    parser.add_argument("--fecha-hasta", help="dd/mm/aaaa (requerido en modo fecha/marca_fecha)")
    parser.add_argument("--precio-actualizado", type=float,
                         help="Nuevo valor de dólar (multiplicador)")
    parser.add_argument("--confirmar", action="store_true",
                         help="Aplica los cambios de verdad. Sin este flag, solo muestra el dry-run "
                              "(salvo en el menú guiado, que pregunta al final).")
    args = parser.parse_args()

    if args.modo is None:
        # Sin --modo (y en general sin ningún flag): menú guiado completo.
        modo_guiado()
        return

    if not args.archivo:
        args.archivo = str(ARCHIVO_PATH)

    if args.precio_actualizado is None:
        sys.exit("ERROR: falta --precio-actualizado")
    if not Path(args.archivo).exists():
        sys.exit(f"ERROR: no se encontró el archivo '{args.archivo}'")

    procesar(args, interactivo=False)


if __name__ == "__main__":
    main()
