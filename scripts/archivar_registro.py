#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
archivar_registro.py
---------------------
Archiva el historial de la hoja RegistroCambios en un libro aparte, para que
el libro principal quede más liviano, y deja RegistroCambios solo con los
encabezados de registro.py más una fila "Estructura" que deja constancia del
archivo.

  1. Copia todas las filas de RegistroCambios (encabezados incluidos, mismo
     orden y mismos valores) a:
        archivo/RegistroCambios_historico_<aaaammdd primera>_a_<aaaammdd última>.xlsx
     Las fechas salen de la columna Fecha de la primera y la última fila de
     datos. Las celdas que en el libro principal son fórmulas vivas (los
     ValorNuevo de "Formulas", p. ej. ='Bicicletas MTB'!$H2/1.21) se guardan
     en el histórico como TEXTO con la misma fórmula: el histórico no tiene
     esas hojas y como fórmula darían #REF!.
  2. En el libro principal edita SOLO el XML de la hoja RegistroCambios
     (queda la fila 1 tal cual + 1 fila "Estructura", dimension y vista
     reiniciadas) y quita xl/calcChain.xml (Excel lo regenera al abrir; ya
     no tendría sentido porque apunta a las fórmulas borradas). Todo lo
     demás, incluida "Marcas Actualización Dinámica", se copia byte a byte.

Registro: esta tarea NO se registra celda por celda (sería registrar el
borrado del propio registro). Es la única excepción a la regla 8 del
CLAUDE.md: deja UNA fila TipoOperacion "Estructura" con la cantidad de filas
archivadas y el nombre del histórico.

El aviso del archivo va en esa fila y NO en una celda suelta (por ejemplo
L1): registro.validar_encabezados() aborta si hay texto en la fila 1 a la
derecha de la columna J.

Verificaciones (siempre, antes de tocar el archivo real):
  1. Histórico (openpyxl): misma cantidad de filas, primera y última fila
     iguales y, además, todas las celdas iguales a las del original.
  2. Libro nuevo (openpyxl): RegistroCambios = encabezados + fila Estructura,
     validar_encabezados() pasa, y todas las demás hojas con valores
     idénticos al original.
  3. Byte a byte: Marcas y toda parte no modificada quedan idénticas.
  4. Excel (COM): el libro nuevo y el histórico abren sin pedido de reparación.

USO (desde scripts/)
    python archivar_registro.py            -> inspección + dry-run sobre una copia en una carpeta temporal
    python archivar_registro.py --execute  -> lo mismo + backup + histórico + reemplazo del archivo real
    --archivo RUTA                         -> procesa otro libro (por ejemplo, una copia de prueba)
    --sin-excel                            -> omite la verificación 4 (solo para pruebas)
"""

import argparse
import html
import re
import shutil
import sys
import tempfile
import warnings
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

import registro
from tablas_xml import (HOJA_MARCAS, Abortar, abrir_en_excel, agregar_registro,
                        mapa_partes, sha, valor_comparable, verificar_en_excel)

warnings.filterwarnings("ignore")

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"
CALC_CHAIN = "xl/calcChain.xml"

# Elementos de la hoja que el motor no sabe ajustar: si aparecen, se aborta.
NO_CONTEMPLADOS = ("autoFilter", "tableParts", "mergeCells", "conditionalFormatting",
                   "dataValidations", "hyperlinks", "extLst", "sortState", "pane")


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def kb(n):
    return f"{n / 1024:,.0f} KB".replace(",", ".")


def fecha_aaaammdd(valor, donde):
    if isinstance(valor, datetime):
        return valor.strftime("%Y%m%d")
    try:
        return datetime.strptime(str(valor).strip()[:10], "%d/%m/%Y").strftime("%Y%m%d")
    except (TypeError, ValueError):
        raise Abortar(f"no se pudo leer la fecha de la {donde} fila de RegistroCambios: {valor!r}")


def leer_registro(archivo):
    """Todas las filas de RegistroCambios (valores tal como los lee openpyxl,
    sin data_only: las fórmulas vienen como texto '=...')."""
    wb = openpyxl.load_workbook(archivo, read_only=True)
    if registro.HOJA_REGISTRO not in wb.sheetnames:
        raise Abortar(f"no existe la hoja '{registro.HOJA_REGISTRO}'.")
    ws = wb[registro.HOJA_REGISTRO]
    filas = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return filas


def valores_hojas(archivo, excluir=()):
    """{hoja: [filas]} de todas las hojas, para comparar valores."""
    wb = openpyxl.load_workbook(archivo, read_only=True)
    res = {ws.title: [tuple(valor_comparable(v) for v in r) for r in ws.iter_rows(values_only=True)]
           for ws in wb.worksheets if ws.title not in excluir}
    orden = wb.sheetnames
    wb.close()
    return res, orden


# --------------------------------------------------------------------------
# Histórico
# --------------------------------------------------------------------------
def crear_historico(filas, destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = registro.HOJA_REGISTRO
    for r, fila in enumerate(filas, start=1):
        for c, v in enumerate(fila, start=1):
            if v is None:
                continue
            celda = ws.cell(row=r, column=c, value=v)
            if isinstance(v, str) and v.startswith("="):
                celda.data_type = "s"  # fórmula del original -> texto
    for c in range(1, len(filas[0]) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.column_dimensions["E"].width = 53.6
    ws.column_dimensions["J"].width = 60
    wb.save(destino)


def verificar_historico(filas, destino):
    wb = openpyxl.load_workbook(destino, read_only=True)
    ws = wb[registro.HOJA_REGISTRO]
    leidas = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    ancho = len(filas[0])
    leidas = [tuple(list(r) + [None] * (ancho - len(r)))[:ancho] for r in leidas]
    errores = []
    if len(leidas) != len(filas):
        errores.append(f"filas: histórico {len(leidas)} vs original {len(filas)}")
    if leidas[:1] != filas[:1]:
        errores.append("los encabezados no coinciden")
    if leidas[1:2] != filas[1:2]:
        errores.append("la primera fila de datos no coincide")
    if leidas[-1:] != filas[-1:]:
        errores.append("la última fila no coincide")
    distintas = [i + 1 for i, (a, b) in enumerate(zip(filas, leidas)) if a != b]
    if distintas:
        errores.append(f"{len(distintas)} filas con valores distintos (p. ej. {distintas[:5]})")
    return not errores, errores, len(leidas)


# --------------------------------------------------------------------------
# Libro principal (XML)
# --------------------------------------------------------------------------
def vaciar_hoja_registro(xml, fila_estructura):
    for tag in NO_CONTEMPLADOS:
        if re.search(rf"<{tag}\b", xml):
            raise Abortar(f"la hoja RegistroCambios tiene <{tag}>; el script no lo contempla.")
    if "<sheetData>" not in xml or "</sheetData>" not in xml:
        raise Abortar("la hoja RegistroCambios no tiene un bloque sheetData editable.")
    cab, resto = xml.split("<sheetData>", 1)
    datos, pie = resto.split("</sheetData>", 1)
    m = re.match(r'<row\b[^>]*?\br="1"[^>]*>.*?</row>', datos, re.S)
    if not m:
        raise Abortar("no se encontró la fila 1 de RegistroCambios en el XML.")
    # Dimension y vista: la hoja queda de 2 filas, con la vista arriba de todo.
    cab = re.sub(r'(<dimension\b[^>]*?\bref=")[^"]*(")', r"\g<1>A1:J1\g<2>", cab, count=1)
    cab = re.sub(r'\s+topLeftCell="[^"]*"', "", cab)
    cab = re.sub(r"<selection\b[^>]*/>", '<selection activeCell="A1" sqref="A1"/>', cab)
    xml = cab + "<sheetData>" + m.group(0) + "</sheetData>" + pie
    return agregar_registro(xml, [fila_estructura])


def construir_libro(origen, salida, fila_estructura):
    z = zipfile.ZipFile(origen)
    partes = mapa_partes(z)
    for req in (HOJA_MARCAS, registro.HOJA_REGISTRO):
        if req not in partes:
            raise Abortar(f"falta la hoja '{req}' en el libro.")
    if partes[registro.HOJA_REGISTRO]["tablas"]:
        raise Abortar("RegistroCambios tiene una tabla de Excel; el script no la contempla.")
    p_reg = partes[registro.HOJA_REGISTRO]["parte"]
    xml_reg, desde, hasta = vaciar_hoja_registro(z.read(p_reg).decode("utf-8"), fila_estructura)
    if (desde, hasta) != (2, 2):
        raise Abortar(f"la fila Estructura quedó en {desde}-{hasta} (se esperaba la fila 2).")
    modificadas = {p_reg: xml_reg.encode("utf-8")}

    quitar = set()
    if CALC_CHAIN in z.namelist():
        quitar.add(CALC_CHAIN)
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rels2 = re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', "", rels)
        ct = z.read("[Content_Types].xml").decode("utf-8")
        ct2 = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        if rels2 == rels or ct2 == ct:
            raise Abortar("no se encontró la referencia a calcChain.xml para quitarla.")
        modificadas["xl/_rels/workbook.xml.rels"] = rels2.encode("utf-8")
        modificadas["[Content_Types].xml"] = ct2.encode("utf-8")

    with zipfile.ZipFile(salida, "w") as zout:
        for item in z.infolist():
            if item.filename in quitar:
                continue
            datos = modificadas.get(item.filename, z.read(item.filename))
            nuevo = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            nuevo.compress_type, nuevo.external_attr = item.compress_type, item.external_attr
            zout.writestr(nuevo, datos)

    # Verificación byte a byte
    parte_marcas = partes[HOJA_MARCAS]["parte"]
    with zipfile.ZipFile(salida) as zo:
        esperadas = sorted(n for n in z.namelist() if n not in quitar)
        conjunto_ok = sorted(zo.namelist()) == esperadas
        marcas_ok = sha(zo.read(parte_marcas)) == sha(z.read(parte_marcas))
        alteradas = [n for n in esperadas if n not in modificadas and zo.read(n) != z.read(n)]
    z.close()
    ok = conjunto_ok and marcas_ok and not alteradas
    det = (f"Marcas idéntica byte a byte: {'sí' if marcas_ok else 'NO'} | "
           f"partes sin tocar alteradas: {alteradas or 'ninguna'} | "
           f"conjunto de partes esperado: {'sí' if conjunto_ok else 'NO'} | "
           f"modificadas: {', '.join(sorted(modificadas))} | quitadas: {', '.join(sorted(quitar)) or 'ninguna'}")
    return ok, det


def verificar_libro(orig_vals, orden_orig, salida, fila_estructura):
    errores = []
    nuevos, orden = valores_hojas(salida)
    if orden != orden_orig:
        errores.append(f"el orden de hojas cambió: {orden}")
    for hoja, filas in orig_vals.items():
        if nuevos.get(hoja) != filas:
            errores.append(f"la hoja '{hoja}' tiene valores distintos")
    reg = nuevos.get(registro.HOJA_REGISTRO, [])
    esperado = [tuple(registro.REGISTRO_HEADERS),
                tuple(registro._valor_para_celda(fila_estructura[h]) for h in registro.REGISTRO_HEADERS)]
    if reg != esperado:
        errores.append(f"RegistroCambios no quedó con encabezados + fila Estructura: {reg[:3]}")
    wb = openpyxl.load_workbook(salida)
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        errores.append(f"validar_encabezados falla: {e}")
    if registro.hoja_vacia(wb[registro.HOJA_REGISTRO]):
        errores.append("RegistroCambios quedó vacía")
    wb.close()
    return not errores, errores


# --------------------------------------------------------------------------
# Proceso principal
# --------------------------------------------------------------------------
def procesar(execute, archivo, con_excel=True):
    archivo = Path(archivo).resolve()
    if not archivo.exists():
        raise Abortar(f"no se encontró el archivo de trabajo en: {archivo}")
    sha_origen = sha(archivo.read_bytes())
    tam_origen = archivo.stat().st_size
    modo = "EJECUCIÓN" if execute else "DRY-RUN (sobre una copia; no se modifica el archivo real)"
    print(f"\n=== Archivar historial de {registro.HOJA_REGISTRO} | Modo: {modo} ===\n")

    carpeta = Path(tempfile.mkdtemp(prefix="archivar_registro_"))
    copia = carpeta / ("ORIGEN_" + archivo.name)
    shutil.copy2(archivo, copia)
    if sha(copia.read_bytes()) != sha_origen:
        raise Abortar("la copia de trabajo no coincide con el archivo real.")

    # ---------------- Fase 1: inspección ----------------
    print("--- Fase 1: Inspección ---")
    filas = leer_registro(copia)
    if not filas or [registro._normalizar_encabezado(v) for v in filas[0][:10]] != \
            [registro._normalizar_encabezado(h) for h in registro.REGISTRO_HEADERS]:
        raise Abortar(f"RegistroCambios no tiene los encabezados de registro.py: {filas[:1]}")
    wb_chk = openpyxl.load_workbook(copia, read_only=True)
    try:
        registro.validar_encabezados(wb_chk)
    finally:
        wb_chk.close()
    datos = filas[1:]
    if not datos:
        print("RegistroCambios no tiene filas para archivar.")
        return 0
    vacias = sum(1 for r in datos if all(v is None or str(v).strip() == "" for v in r))
    formulas_vivas = sum(1 for r in datos for v in r if isinstance(v, str) and v.startswith("="))
    primera, ultima = datos[0][0], datos[-1][0]
    nombre_hist = (f"RegistroCambios_historico_{fecha_aaaammdd(primera, 'primera')}"
                   f"_a_{fecha_aaaammdd(ultima, 'última')}.xlsx")
    destino_hist = archivo.parent / nombre_hist

    with zipfile.ZipFile(copia) as z:
        partes = mapa_partes(z)
        p_reg = partes[registro.HOJA_REGISTRO]["parte"]
        xml_reg = z.read(p_reg).decode("utf-8")
        info_reg = z.getinfo(p_reg)
        tablas = partes[registro.HOJA_REGISTRO]["tablas"]
        wbx = z.read("xl/workbook.xml").decode("utf-8")
    idx_reg = list(partes).index(registro.HOJA_REGISTRO)
    nombres_reg = [html.unescape(n) for n in re.findall(r"<definedName\b[^>]*>[^<]*</definedName>", wbx)
                   if f'localSheetId="{idx_reg}"' in n or registro.HOJA_REGISTRO in n]

    print(f"Libro: {archivo}")
    print(f"  Tamaño actual: {kb(tam_origen)} ({tam_origen:,} bytes)".replace(",", "."))
    print(f"  Parte XML de RegistroCambios: {p_reg} | {kb(info_reg.compress_size)} comprimida, "
          f"{kb(info_reg.file_size)} sin comprimir")
    print(f"  Filas: {len(filas):,} (1 de encabezados + {len(datos):,} de datos; vacías: {vacias})".replace(",", "."))
    print(f"  Primera fila de datos: Fecha {primera} | {datos[0][1]} | {datos[0][2]}")
    print(f"  Última fila de datos:  Fecha {ultima} | {datos[-1][1]} | {datos[-1][2]}")
    print(f"  Celdas con fórmula viva dentro del registro: {formulas_vivas} "
          "(en el histórico se guardan como texto)")
    print(f"  Tabla de Excel en RegistroCambios: {', '.join(tablas) if tablas else 'no'}")
    print(f"  Autofiltro en la hoja: {'sí' if '<autoFilter' in xml_reg else 'no'} | "
          f"nombres definidos de la hoja: {nombres_reg or 'ninguno'}")
    print(f"  Archivo histórico: {destino_hist}")
    if destino_hist.exists():
        raise Abortar(f"ya existe {destino_hist.name}; no se pisa. Revisalo antes de volver a correr.")

    # ---------------- Fase 2: armado y verificaciones ----------------
    fecha = registro.timestamp_corrida()
    cantidad = len(datos)
    fila_estructura = registro.nueva_fila(
        fecha, registro.TIPO_ESTRUCTURA, registro.HOJA_REGISTRO, None, None, None,
        f"Historial (filas 2 a {len(filas)})",
        f"{cantidad} filas ({primera} a {ultima})",
        f"Archivado en {nombre_hist}",
        f"Archivo del historial: {cantidad} filas de RegistroCambios archivadas en "
        f"archivo/{nombre_hist}; en este libro quedan solo los encabezados.")

    print("\n--- Fase 2: Simulación ---")
    hist_tmp = carpeta / nombre_hist
    salida = carpeta / archivo.name
    crear_historico(filas, hist_tmp)
    ok_bytes, det_bytes = construir_libro(copia, salida, fila_estructura)
    print(f"  Histórico simulado: {hist_tmp} ({kb(hist_tmp.stat().st_size)})")
    print(f"  Libro simulado:     {salida}")
    tam_nuevo = salida.stat().st_size
    print(f"  Tamaño del libro: {kb(tam_origen)} -> {kb(tam_nuevo)} "
          f"({100 * (tam_origen - tam_nuevo) / tam_origen:.0f} % menos)")
    print(f"\n--- Fila que queda en '{registro.HOJA_REGISTRO}' (fila 2) ---")
    print("    - " + " | ".join("(vacío)" if fila_estructura[h] is None else str(fila_estructura[h])
                                for h in registro.REGISTRO_HEADERS))

    print("\n--- Verificación 1: histórico (openpyxl) ---")
    ok1, err1, n_hist = verificar_historico(filas, hist_tmp)
    print(f"  Filas: original {len(filas)} | histórico {n_hist} | "
          "primera y última fila, y todas las celdas: " + ("iguales" if ok1 else "DISTINTAS"))
    print("  OK" if ok1 else "\n".join(f"  ERROR: {e}" for e in err1))

    print("\n--- Verificación 2: libro nuevo (openpyxl) ---")
    orig_vals, orden_orig = valores_hojas(copia, excluir={registro.HOJA_REGISTRO})
    ok2, err2 = verificar_libro(orig_vals, orden_orig, salida, fila_estructura)
    print(f"  Demás hojas ({len(orig_vals)}) con valores idénticos, RegistroCambios = encabezados + "
          f"fila Estructura, validar_encabezados OK")
    print("  OK" if ok2 else "\n".join(f"  ERROR: {e}" for e in err2))

    print("\n--- Verificación 3: byte a byte ---")
    print(f"  {'OK' if ok_bytes else 'ERROR'} | {det_bytes}")

    ok4 = True
    if con_excel:
        print("\n--- Verificación 4: Excel ---")
        ok_l, res_l, lin_l, _ = verificar_en_excel(salida, carpeta, [])
        print(f"  Libro nuevo: {'abre sin pedido de reparación' if res_l.get('ok') else 'NO abre'}"
              f" | tablas: {len(res_l.get('tablas') or [])}")
        if not ok_l:
            print("\n".join(l for l in lin_l if "ERROR" in l))
        res_h = abrir_en_excel(hist_tmp, carpeta)
        print(f"  Histórico:   {'abre sin pedido de reparación' if res_h.get('ok') else 'NO abre: ' + str(res_h.get('error'))}")
        ok4 = ok_l and res_h.get("ok")
        print(f"  {'OK' if ok4 else 'ERROR'}")
    else:
        print("\n--- Verificación 4: Excel --- (omitida con --sin-excel)")

    if not (ok1 and ok2 and ok_bytes and ok4):
        raise Abortar(f"alguna verificación falló; no se toca el archivo real. Carpeta: {carpeta}")

    if not execute:
        print(f"\nEsto fue una simulación (dry-run). No se modificó el archivo real ni se creó backup.")
        print(f"Carpeta temporal: {carpeta}")
        return cantidad

    # ---------------- Fase 3: ejecución ----------------
    if sha(archivo.read_bytes()) != sha_origen:
        raise Abortar("el archivo real cambió durante la corrida; volvé a correr el script.")
    if destino_hist.exists():
        raise Abortar(f"apareció {destino_hist.name} durante la corrida; no se pisa.")
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    backup = BACKUP_DIR / f"{archivo.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{archivo.suffix}"
    shutil.copy2(archivo, backup)
    if sha(backup.read_bytes()) != sha_origen:
        raise Abortar(f"el backup no coincide con el original: {backup}")
    print(f"\nBackup creado en: {backup}")

    shutil.copyfile(hist_tmp, destino_hist)
    if sha(destino_hist.read_bytes()) != sha(hist_tmp.read_bytes()):
        destino_hist.unlink()
        raise Abortar("el histórico copiado no coincide con el verificado; no se tocó el libro.")
    print(f"Histórico creado: {destino_hist}")
    try:
        shutil.copyfile(salida, archivo)
    except PermissionError:
        destino_hist.unlink()
        raise Abortar("no se pudo reemplazar el libro porque está abierto en Excel. Cerralo y "
                      "volvé a correr el script. El libro no se modificó y se quitó el histórico recién creado.")
    if sha(archivo.read_bytes()) != sha(salida.read_bytes()):
        raise Abortar(f"el libro guardado no coincide con el verificado. Restaurá el backup: {backup}")
    print(f"Libro reemplazado: {archivo} ({kb(archivo.stat().st_size)})")
    print(f"Filas archivadas: {cantidad} | Filas en {registro.HOJA_REGISTRO}: encabezados + 1 (Estructura)")
    return cantidad


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Archiva el historial de RegistroCambios en un libro aparte.")
    ap.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    ap.add_argument("--archivo", type=Path, default=ARCHIVO_PATH,
                    help="Libro a procesar (por defecto, el archivo real).")
    ap.add_argument("--sin-excel", action="store_true", help="Omite la verificación en Excel.")
    args = ap.parse_args()
    try:
        procesar(args.execute, args.archivo, con_excel=not args.sin_excel)
    except (Abortar, registro.RegistroError) as e:
        sys.exit(f"\nABORTADO: {e}")


if __name__ == "__main__":
    main()
