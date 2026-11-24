#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
normalizar_nombres.py
----------------------
Normaliza la columna NombreProducto de UNA hoja de destino al formato

    <PREFIJO> <MARCA> <MODELO>        (todo en mayúsculas)

por ejemplo "BICICLETA MTB GIANT ANTHEM ADVANCED SL 1".

USO (desde scripts/)
    python normalizar_nombres.py --hoja "Bicicletas MTB" --prefijo "BICICLETA MTB"
        -> Fases 1 y 2: inspección + genera la planilla de revisión
           (revision_nombres_mtb.xlsx en la raíz). NO toca el libro.
    python normalizar_nombres.py --hoja "Bicicletas MTB" --prefijo "BICICLETA MTB" --aplicar
        -> dry-run de la Fase 3: lee la planilla de revisión y muestra qué se
           aplicaría. NO toca el libro.
    ... --aplicar --execute
        -> Fase 3 real: backup + escritura + RegistroCambios.
    --revision RUTA  -> otra planilla de revisión (por defecto
                        revision_nombres_<última palabra de la hoja>.xlsx)
    --archivo RUTA   -> procesa otro libro (por ejemplo, una copia de prueba)

Reglas de la propuesta
  - Columnas por nombre de encabezado: CodigoDux, NombreProducto, Marca.
  - Fila de producto = CodigoDux no vacío.
  - Se limpian espacios (tabs, dobles, al principio/final) y se pasa a
    mayúsculas.
  - Se quita la primera palabra del prefijo si ya está al principio
    (BICICLETA), y toda aparición de las demás palabras del prefijo (MTB),
    porque se reponen en su lugar.
  - Se quita la marca (columna Marca) donde aparezca y se la ubica justo
    después del prefijo. Si no estaba, se agrega.
  - No se borra ninguna otra palabra (rodado, grupo, color, talle, año).
  - Las filas cuyo nombre ya es exactamente el propuesto no se tocan.

Filas dudosas (Dudoso = SI, Aplicar = NO; nunca se corrigen solas)
  - Sin Marca.
  - E-bikes (E-BIKE, EBIKE, ELECTRICA) y kits (KIT en el nombre o en CodigoDux).
  - Posibles errores de tipeo: primera palabra parecida pero distinta de la
    del prefijo (BICLCETA), o palabras muy parecidas a otra que aparece en
    los nombres del libro (ADVACED ~ ADVANCED).
  - Nombre que menciona otra marca de la misma hoja distinta de la de la fila.
  - Nombre con descripción propia ("BICICLETA ...") que no menciona la marca
    de la fila (posible marca mal cargada).
  - Nombre con otra categoría (GRAVEL, RUTA, PASEO, BMX, URBANA...).

Fase 3
  - Solo filas con Aplicar = SI. NombrePropuesto puede haber sido editado.
  - Columna opcional MarcaPropuesta: si tiene valor y difiere de la Marca
    actual, también se escribe la Marca (una fila de registro aparte,
    Columna "Marca"). Conteo: registro = celdas de nombre + celdas de marca.
  - Antes de escribir verifica que cada fila siga teniendo el mismo
    CodigoDux y NombreActual que en la planilla; si algo cambió, aborta.
  - openpyxl sin data_only; backup con timestamp; una fila de registro
    (TipoOperacion "Nombre", Columna "NombreProducto") por celda modificada,
    con verificación de conteo. Se guarda primero en una copia temporal, se
    relee y se compara contra el original (todas las hojas, incluida
    "Marcas Actualización Dinámica", que no se modifica) antes de reemplazar
    el archivo real.
"""

import argparse
import difflib
import re
import shutil
import sys
import tempfile
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import actualizado_segun as act
import registro
import tablas_xml as tx

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"

HOJA_MARCAS = "Marcas Actualización Dinámica"
HOJAS_EXCLUIDAS = {
    "Hoja1", HOJA_MARCAS, registro.HOJA_REGISTRO,
    act.HOJA_NO_ACTUALIZADOS, *act.HOJAS_REEMPLAZADAS,
}

COL_DUX = "CodigoDux"
COL_NOMBRE = "NombreProducto"
COL_MARCA = "Marca"

TIPO_NOMBRE = registro.TIPO_NOMBRE

COL_REV_MARCA = "MarcaPropuesta"  # vacía = no cambiar la Marca
REVISION_HEADERS = ["Fila", "CodigoDux", "Marca", "NombreActual",
                    "NombrePropuesto", COL_REV_MARCA, "Motivo", "Dudoso", "Aplicar"]
# Al leer, las columnas se buscan por nombre; MarcaPropuesta es opcional
# (planillas viejas, como revision_nombres_mtb.xlsx, no la tienen en ese orden).
REVISION_REQUERIDAS = ["Fila", "CodigoDux", "NombreActual", "NombrePropuesto",
                       "Dudoso", "Aplicar"]
HOJA_PROPUESTA = "Propuesta"  # formato viejo de una sola hoja
HOJA_SIN_CAMBIOS = "SinCambios"
HOJA_INFO = "Info"

PALABRAS_EBIKE = {"E-BIKE", "EBIKE", "E-BIKES", "ELECTRICA", "ELÉCTRICA"}
PALABRAS_CATEGORIA = {"GRAVEL", "RUTA", "PASEO", "BMX", "URBANA", "INFANTIL",
                      "MTB", "TRIATLON"}
# Primeras palabras que se parecen a BICICLETA pero no son tipeos: otro tipo
# de producto. Se conservan en el modelo y la fila queda dudosa.
OTROS_PRODUCTOS = {"TRICICLETA", "MONOCICLO", "TRICICLO"}
UMBRAL_TIPEO = 0.86


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def normalizar(texto):
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.strip().casefold()


def es_vacio(v):
    return v is None or str(v).strip() == ""


def limpiar(texto):
    """Mayúsculas y espacios simples (tabs y NBSP incluidos)."""
    return " ".join(str(texto).replace("\xa0", " ").split()).upper()


def tokens(texto):
    return limpiar(texto).split() if not es_vacio(texto) else []


def columnas(ws, requeridas):
    mapa = {}
    for celda in ws[1]:
        if not es_vacio(celda.value):
            mapa.setdefault(normalizar(celda.value), celda.column)
    faltan = [c for c in requeridas if normalizar(c) not in mapa]
    if faltan:
        raise SystemExit(f"ERROR: la hoja '{ws.title}' no tiene las columnas {faltan}.")
    return {c: mapa[normalizar(c)] for c in requeridas}


def quitar_secuencia(toks, seq):
    """Quita todas las apariciones de la secuencia de tokens seq."""
    if not seq:
        return toks, 0
    out, i, n = [], 0, 0
    while i < len(toks):
        if toks[i:i + len(seq)] == seq:
            i += len(seq)
            n += 1
        else:
            out.append(toks[i])
            i += 1
    return out, n


def es_palabra(t):
    return bool(re.fullmatch(r"[A-ZÑ]+", t))


def palabra(t):
    """Token sin la puntuación pegada: '(ILUSION' -> 'ILUSION'."""
    return t.strip("()[]{},.;:/-+")


def hojas_destino(wb):
    return [ws for ws in wb.worksheets if ws.title not in HOJAS_EXCLUIDAS]


def slug_hoja(hoja):
    return normalizar(hoja.split()[-1]).replace("/", "-")


# --------------------------------------------------------------------------
# Contexto del libro: vocabulario y marcas (solo lectura)
# --------------------------------------------------------------------------
def contexto_libro(wb):
    vocab, marcas = Counter(), set()
    for ws in hojas_destino(wb):
        enc = {normalizar(c.value): c.column for c in ws[1] if not es_vacio(c.value)}
        c_nom, c_mar = enc.get(normalizar(COL_NOMBRE)), enc.get(normalizar(COL_MARCA))
        if not c_nom:
            continue
        for r in range(2, ws.max_row + 1):
            vocab.update(p for p in map(palabra, tokens(ws.cell(r, c_nom).value))
                         if es_palabra(p))
            if c_mar and not es_vacio(ws.cell(r, c_mar).value):
                marcas.add(limpiar(ws.cell(r, c_mar).value))
    return vocab, marcas


def solo_cambia_terminacion(a, b):
    """CARBON/CARBONO, FRENO/FRENOS, COMPLETA/COMPLETO: variantes de género o
    número, no tipeos."""
    corto, largo = sorted((a, b), key=len)
    if len(largo) - len(corto) == 1 and largo.startswith(corto):
        return True
    return len(a) == len(b) and a[:-1] == b[:-1]


def posibles_tipeos(toks, vocab, excluir):
    hallados = []
    for t in map(palabra, toks):
        if len(t) < 5 or not es_palabra(t) or t in excluir:
            continue
        for cand, n in vocab.items():
            if cand == t or len(cand) < 5 or n < vocab[t]:
                continue
            if solo_cambia_terminacion(t, cand):
                continue
            if difflib.SequenceMatcher(None, t, cand).ratio() >= UMBRAL_TIPEO:
                hallados.append(f"{t} ~ {cand}")
    return hallados


# --------------------------------------------------------------------------
# Propuesta de una fila
# --------------------------------------------------------------------------
def proponer(nombre, marca, codigo_dux, prefijo_toks, vocab, marcas_libro,
             codigo_factura=None):
    motivos, dudas = [], []
    original = "" if nombre is None else str(nombre)
    toks = tokens(original)
    marca_toks = tokens(marca)

    if limpiar(original) != original:
        motivos.append("espacios/mayúsculas normalizados")

    # Primera palabra del prefijo (BICICLETA), con detección de tipeo.
    primera = prefijo_toks[0]
    if toks and toks[0] == primera:
        toks = toks[1:]
    elif toks and toks[0] in OTROS_PRODUCTOS:
        dudas.append(f"no es {primera}: {toks[0]} (¿otra hoja u otro prefijo?)")
        motivos.append(f"se agrega {primera} y se conserva {toks[0]} en el modelo")
    elif toks and difflib.SequenceMatcher(None, toks[0], primera).ratio() >= 0.8:
        dudas.append(f"posible error de tipeo: {toks[0]} ~ {primera}")
        motivos.append(f"{toks[0]} reemplazado por {primera}")
        toks = toks[1:]
    else:
        motivos.append(f"se agrega {primera}")

    # Demás palabras del prefijo (MTB): se quitan donde estén.
    tenia_prefijo_ok = limpiar(original).startswith(" ".join(prefijo_toks) + " ")
    for p in prefijo_toks[1:]:
        # "BICICLETA DE RUTA ..." -> el "DE" solo conecta la categoría.
        toks, n_de = quitar_secuencia(toks, ["DE", p])
        toks, n = quitar_secuencia(toks, [p])
        n += n_de
        if n_de:
            motivos.append(f"'DE {p}' -> {p} en el prefijo")
        if not tenia_prefijo_ok:
            motivos.append(f"{p} movido al prefijo" if n else f"se agrega {p}")
        elif n > 1:
            motivos.append(f"{p} repetido quitado")

    # Marca.
    if marca_toks:
        toks, n = quitar_secuencia(toks, marca_toks)
        if n == 0:
            motivos.append(f"se agrega la marca {' '.join(marca_toks)}")
            if limpiar(original).startswith(primera) or (toks and toks[0] == primera):
                dudas.append("el nombre ya trae descripción pero no menciona la "
                             "marca de la fila (¿marca mal cargada?)")
        else:
            motivos.append("marca movida después del prefijo")
            if n > 1:
                motivos.append("marca repetida quitada")
    else:
        dudas.append("fila sin Marca")

    # Dudas por contenido.
    todos = set(tokens(original))
    if todos & PALABRAS_EBIKE:
        dudas.append("e-bike")
    if "KIT" in todos or re.search(r"(^|[^A-Z])KIT([^A-Z]|$)", limpiar(codigo_dux or "")):
        dudas.append("posible kit (KIT en el nombre o en CodigoDux)")
    if re.search(r"KID|INFANTIL", limpiar(codigo_factura or "")):
        dudas.append("posible infantil (KID/INFANTIL en CodigoFactura)")
    otras_cat = (todos & PALABRAS_CATEGORIA) - set(prefijo_toks)
    if otras_cat:
        dudas.append(f"menciona otra categoría: {', '.join(sorted(otras_cat))}")
    otras_marcas = sorted(m for m in marcas_libro
                          if m != " ".join(marca_toks) and m in todos and len(m) >= 3)
    if otras_marcas:
        dudas.append(f"menciona otra marca: {', '.join(otras_marcas)}")
    tipeos = posibles_tipeos(toks, vocab, excluir=set(marca_toks) | marcas_libro)
    if tipeos:
        dudas.append("posible error de tipeo: " + "; ".join(tipeos))

    propuesto = " ".join(prefijo_toks + marca_toks + toks)
    return propuesto, motivos, dudas


# --------------------------------------------------------------------------
# Fases 1 y 2
# --------------------------------------------------------------------------
COL_FACTURA = "CodigoFactura"  # opcional: solo para detectar infantiles


def validar_hoja(wb, hoja):
    if hoja not in wb.sheetnames or hoja in HOJAS_EXCLUIDAS:
        raise SystemExit(f"ERROR: la hoja '{hoja}' no existe o está excluida.")
    return wb[hoja]


def inspeccionar(wb, hoja, prefijo, vocab):
    ws = validar_hoja(wb, hoja)
    cols = columnas(ws, [COL_DUX, COL_NOMBRE, COL_MARCA])
    enc = {normalizar(c.value): c.column for c in ws[1] if not es_vacio(c.value)}
    c_fac = enc.get(normalizar(COL_FACTURA))
    prefijo_toks = tokens(prefijo)
    # Marcas de esta hoja: las de otras hojas incluyen marcas de componentes
    # (SHIMANO, ELITE...) que aparecen legítimamente en los nombres.
    marcas_libro = {limpiar(ws.cell(r, cols[COL_MARCA]).value)
                    for r in range(2, ws.max_row + 1)
                    if not es_vacio(ws.cell(r, cols[COL_MARCA]).value)}

    filas, vacias = [], 0
    for r in range(2, ws.max_row + 1):
        dux = ws.cell(r, cols[COL_DUX]).value
        if es_vacio(dux):
            vacias += 1
            continue
        nombre = ws.cell(r, cols[COL_NOMBRE]).value
        marca = ws.cell(r, cols[COL_MARCA]).value
        factura = ws.cell(r, c_fac).value if c_fac else None
        propuesto, motivos, dudas = proponer(nombre, marca, str(dux), prefijo_toks,
                                             vocab, marcas_libro, factura)
        filas.append({
            "fila": r, "dux": dux, "marca": marca, "actual": nombre,
            "propuesto": propuesto, "cambia": propuesto != nombre,
            "motivos": motivos, "dudas": dudas,
        })
    return ws, filas, vacias


def nombre_pestana(hoja):
    """Nombre de la pestaña de la planilla de revisión para una hoja."""
    return hoja[:31]


def escribir_revision(ruta, lotes):
    """lotes: lista de (hoja, prefijo, filas). Una pestaña por hoja, más
    SinCambios (todas las hojas) e Info."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    azul = PatternFill("solid", fgColor="305496")
    i_motivo = REVISION_HEADERS.index("Motivo")
    i_dudoso = REVISION_HEADERS.index("Dudoso")
    letra_aplicar = get_column_letter(REVISION_HEADERS.index("Aplicar") + 1)

    for hoja, _prefijo, filas in lotes:
        ws = wb.create_sheet(nombre_pestana(hoja))
        ws.append(REVISION_HEADERS)
        for f in (x for x in filas if x["cambia"]):
            motivo = "; ".join(f["motivos"])
            if f["dudas"]:
                motivo = "DUDOSO: " + " | ".join(f["dudas"]) + " || " + motivo
            ws.append([f["fila"], str(f["dux"]), f["marca"], f["actual"],
                       f["propuesto"], None, motivo,
                       "SI" if f["dudas"] else "NO", "NO" if f["dudas"] else "SI"])
        for celda in ws[1]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = azul
        for fila in ws.iter_rows(min_row=2):
            if fila[i_dudoso].value == "SI":
                for c in fila:
                    c.fill = amarillo
            fila[i_motivo].alignment = Alignment(wrap_text=True, vertical="top")
        for col, ancho in enumerate((6, 16, 12, 55, 60, 15, 70, 8, 8), start=1):
            ws.column_dimensions[get_column_letter(col)].width = ancho
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        if ws.max_row > 1:
            dv = DataValidation(type="list", formula1='"SI,NO"', allow_blank=False)
            dv.add(f"{letra_aplicar}2:{letra_aplicar}{ws.max_row}")
            ws.add_data_validation(dv)

    ws2 = wb.create_sheet(HOJA_SIN_CAMBIOS)
    ws2.append(["Hoja", "Fila", "CodigoDux", "Marca", "NombreProducto"])
    for hoja, _prefijo, filas in lotes:
        for f in (x for x in filas if not x["cambia"]):
            ws2.append([hoja, f["fila"], str(f["dux"]), f["marca"], f["actual"]])
    for celda in ws2[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = azul
    for col, ancho in zip("ABCDE", (20, 6, 16, 12, 60)):
        ws2.column_dimensions[col].width = ancho

    ws3 = wb.create_sheet(HOJA_INFO)
    ws3.append(["Generado", datetime.now().strftime("%d/%m/%Y %H:%M")])
    for hoja, prefijo, _filas in lotes:
        ws3.append(["Hoja", hoja, "Prefijo", prefijo, "Pestaña", nombre_pestana(hoja)])
    ws3.append(["Instrucciones", "Podés editar NombrePropuesto, MarcaPropuesta "
                "(vacía = no cambiar la Marca) y Aplicar (SI/NO). No modifiques "
                "Fila, CodigoDux ni NombreActual."])
    ws3.column_dimensions["B"].width = 30
    ws3.column_dimensions["D"].width = 25
    wb.save(ruta)


def fase_propuesta(args, pares, ruta_rev):
    wb = openpyxl.load_workbook(args.archivo)
    vocab, _ = contexto_libro(wb)
    lotes = []
    print("=== FASE 1 — Inspección (sin escribir) ===")
    for hoja, prefijo in pares:
        ws, filas, vacias = inspeccionar(wb, hoja, prefijo, vocab)
        cambian = [f for f in filas if f["cambia"]]
        dudosas = [f for f in cambian if f["dudas"]]
        lotes.append((hoja, limpiar(prefijo), filas))
        print(f"\n--- '{hoja}' (prefijo '{limpiar(prefijo)}') ---")
        print(f"  Filas de producto (CodigoDux):        {len(filas)}"
              f"   (filas vacías: {vacias})")
        print(f"  Ya cumplen el formato (no se tocan):  {len(filas) - len(cambian)}")
        print(f"  Hay que cambiar:                      {len(cambian)}")
        print(f"     - no dudosas (Aplicar=SI):         {len(cambian) - len(dudosas)}")
        print(f"     - dudosas    (Aplicar=NO):         {len(dudosas)}")
        if dudosas:
            print("  Dudosas:")
            for f in dudosas:
                print(f"    fila {f['fila']:>3}: {limpiar(f['actual'])}")
                print(f"               -> {' | '.join(f['dudas'])}")

    escribir_revision(ruta_rev, lotes)
    print(f"\n=== FASE 2 — Propuesta generada: {ruta_rev} ===")
    print("  El libro NO se modificó. Revisá la planilla y luego corré el mismo "
          "comando con --aplicar (dry-run) y después --aplicar --execute.")


# --------------------------------------------------------------------------
# Fase 3
# --------------------------------------------------------------------------
def leer_revision(ruta, hojas):
    """Devuelve {hoja: [filas]}. Busca la pestaña con el nombre de cada hoja
    (o 'Propuesta', formato viejo de una sola hoja). Columnas por nombre de
    encabezado; MarcaPropuesta es opcional (vacía = no tocar la Marca)."""
    if not ruta.exists():
        raise SystemExit(f"ERROR: no existe la planilla de revisión {ruta}.")
    wb = openpyxl.load_workbook(ruta, data_only=True)  # planilla propia, no el libro
    resultado = {}
    for hoja in hojas:
        pestana = nombre_pestana(hoja)
        if pestana in wb.sheetnames:
            ws = wb[pestana]
        elif len(hojas) == 1 and HOJA_PROPUESTA in wb.sheetnames:
            ws = wb[HOJA_PROPUESTA]
        else:
            raise SystemExit(f"ERROR: la planilla no tiene la pestaña '{pestana}'.")
        enc = {}
        for celda in ws[1]:
            if es_vacio(celda.value):
                continue
            clave = normalizar(celda.value)
            if clave in enc:
                raise SystemExit(f"ERROR: encabezado repetido en '{ws.title}': {celda.value!r}")
            enc[clave] = celda.column
        faltan = [h for h in REVISION_REQUERIDAS if normalizar(h) not in enc]
        if faltan:
            raise SystemExit(f"ERROR: a la pestaña '{ws.title}' le faltan las columnas {faltan}.")
        usadas = REVISION_REQUERIDAS + [COL_REV_MARCA]
        filas = []
        for r in range(2, ws.max_row + 1):
            fila = {h: (ws.cell(r, enc[normalizar(h)]).value if normalizar(h) in enc else None)
                    for h in usadas}
            if all(es_vacio(v) for v in fila.values()):
                continue
            filas.append(fila)
        resultado[hoja] = filas
    return resultado


def preparar_hoja(ws, hoja, prefijo, revision):
    """Arma los cambios de una hoja. Devuelve (cambios, omitidas, errores, avisos)."""
    cols = columnas(ws, [COL_DUX, COL_NOMBRE, COL_MARCA])
    cambios, omitidas, errores, avisos, vistas = [], [], [], [], set()
    for f in revision:
        aplicar = limpiar(f["Aplicar"] or "")
        try:
            r = int(f["Fila"])
        except (TypeError, ValueError):
            errores.append(f"{hoja}: Fila inválida: {f['Fila']!r}")
            continue
        if r in vistas:
            errores.append(f"{hoja}: fila {r} repetida en la planilla")
            continue
        vistas.add(r)
        if aplicar not in ("SI", "NO"):
            errores.append(f"{hoja} fila {r}: Aplicar debe ser SI o NO (tiene {f['Aplicar']!r})")
            continue
        if aplicar == "NO":
            omitidas.append((r, f["NombreActual"], "Aplicar = NO"
                             + (" (dudosa)" if limpiar(f["Dudoso"] or "") == "SI" else "")))
            continue
        dux_libro = ws.cell(r, cols[COL_DUX]).value
        nombre_libro = ws.cell(r, cols[COL_NOMBRE]).value
        if str(dux_libro).strip() != str(f["CodigoDux"]).strip() or \
                (nombre_libro or "") != (f["NombreActual"] or ""):
            errores.append(f"{hoja} fila {r}: el libro cambió desde la propuesta "
                           f"(CodigoDux/NombreProducto no coinciden)")
            continue
        nuevo = "" if f["NombrePropuesto"] is None else str(f["NombrePropuesto"]).strip()
        if not nuevo:
            errores.append(f"{hoja} fila {r}: NombrePropuesto vacío con Aplicar = SI")
            continue
        marca_libro = ws.cell(r, cols[COL_MARCA]).value
        marca_nueva = None if es_vacio(f[COL_REV_MARCA]) else str(f[COL_REV_MARCA]).strip()
        cambia_marca = marca_nueva is not None and marca_nueva != marca_libro
        if nuevo == nombre_libro and not cambia_marca:
            omitidas.append((r, nombre_libro, "nombre y marca propuestos iguales a los actuales"))
            continue
        if not limpiar(nuevo).startswith(prefijo + " "):
            avisos.append(f"{hoja} fila {r}: '{nuevo}' no empieza con '{prefijo}' (se aplica igual)")
        marca_final = marca_nueva if cambia_marca else marca_libro
        base = {"hoja": hoja, "fila": r, "nombre": nuevo, "marca": marca_final}
        if nuevo != nombre_libro:
            cambios.append({**base, "columna": COL_NOMBRE, "col": cols[COL_NOMBRE],
                            "anterior": nombre_libro, "nuevo": nuevo})
        if cambia_marca:
            cambios.append({**base, "columna": COL_MARCA, "col": cols[COL_MARCA],
                            "anterior": marca_libro, "nuevo": marca_nueva})
    return cambios, omitidas, errores, avisos


def fase_aplicar(args, pares, ruta_rev):
    revision = leer_revision(ruta_rev, [h for h, _ in pares])
    wb = openpyxl.load_workbook(args.archivo)  # sin data_only (regla 3)
    registro.validar_encabezados(wb)

    fecha = registro.timestamp_corrida()
    corridas, cambios, omitidas, errores, avisos = [], [], [], [], []
    for hoja, prefijo_arg in pares:
        ws = validar_hoja(wb, hoja)
        prefijo = limpiar(prefijo_arg)
        c, o, e, a = preparar_hoja(ws, hoja, prefijo, revision[hoja])
        cambios += c
        omitidas += [(hoja, *x) for x in o]
        errores += e
        avisos += a
        # Una corrida por hoja (mismo timestamp) para que el Detalle diga su prefijo.
        corrida = registro.RegistroCorrida(
            TIPO_NOMBRE, f"Normalización de nombres – prefijo '{prefijo}' – {ruta_rev.name}",
            fecha=fecha)
        for x in c:
            corrida.agregar(hoja=hoja, fila=x["fila"], nombre_producto=x["nombre"],
                            marca=x["marca"], columna=x["columna"],
                            valor_anterior=x["anterior"], valor_nuevo=x["nuevo"])
        corridas.append(corrida)

    print(f"=== FASE 3 — {'EJECUCIÓN' if args.execute else 'DRY-RUN'} desde {ruta_rev.name} ===")
    if errores:
        print("\nERRORES (no se escribe nada):")
        for e in errores:
            print(f"  - {e}")
        print("Si el libro cambió, volvé a generar la propuesta.")
        sys.exit(1)

    n_reg = sum(len(c) for c in corridas)
    n_nombre = sum(c["columna"] == COL_NOMBRE for c in cambios)
    print("\n  Resumen por hoja:")
    for hoja, _ in pares:
        ch = [c for c in cambios if c["hoja"] == hoja]
        nn = sum(c["columna"] == COL_NOMBRE for c in ch)
        print(f"    {hoja:<20} filas {len({c['fila'] for c in ch}):>3}   "
              f"celdas {len(ch):>3} ({nn} {COL_NOMBRE} + {len(ch) - nn} {COL_MARCA})   "
              f"omitidas {sum(o[0] == hoja for o in omitidas):>3}")
    print(f"    {'TOTAL':<20} filas {len({(c['hoja'], c['fila']) for c in cambios}):>3}   "
          f"celdas {len(cambios):>3} ({n_nombre} {COL_NOMBRE} + {len(cambios) - n_nombre} "
          f"{COL_MARCA})   filas de RegistroCambios: {n_reg}")
    print()
    for c in cambios:
        print(f"    {c['hoja']} fila {c['fila']:>3} [{c['columna']}]: "
              f"{limpiar(c['anterior'] or '(vacío)')}\n             -> {c['nuevo']}")
    if omitidas:
        print("\n  Omitidas:")
        for hoja, r, nombre, motivo in omitidas:
            print(f"    {hoja} fila {r:>3}: {limpiar(nombre or '')}  [{motivo}]")
    for a in avisos:
        print(f"  AVISO: {a}")
    for corrida in corridas:
        if len(corrida):
            registro.imprimir_previsualizacion(corrida, limite=2)

    if not args.execute:
        print("\n[DRY-RUN] No se escribió nada. Para aplicar, agregá --execute.")
        return
    if not cambios:
        print("\nNo hay cambios para aplicar.")
        return

    for c in cambios:
        wb[c["hoja"]].cell(c["fila"], c["col"]).value = c["nuevo"]
    try:
        registro.verificar_conteo([f for c in corridas for f in c.filas], len(cambios))
        filas_reg_antes = (wb[registro.HOJA_REGISTRO].max_row
                           if registro.HOJA_REGISTRO in wb.sheetnames else 1)
        for corrida in corridas:
            if len(corrida):
                corrida.escribir(wb)
    except registro.RegistroError as e:
        print(f"\n{e}")
        sys.exit(1)

    with tempfile.TemporaryDirectory() as tmp:
        salida = Path(tmp) / args.archivo.name
        wb.save(salida)
        verificar(args.archivo, salida, cambios, filas_reg_antes, n_reg)
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = BACKUP_DIR / f"{args.archivo.stem}_{ts}{args.archivo.suffix}"
        shutil.copy2(args.archivo, backup)
        print(f"\nBackup creado en: {backup}")
        try:
            shutil.copyfile(salida, args.archivo)
        except PermissionError:
            raise SystemExit("ERROR: no se pudo reemplazar el libro (¿está abierto en "
                             "Excel?). No se modificó nada; cerralo y repetí.")

    print(f"\nCambios aplicados: {len(cambios)} celdas "
          f"({n_nombre} {COL_NOMBRE} + {len(cambios) - n_nombre} {COL_MARCA}).")
    print(f"Filas agregadas a {registro.HOJA_REGISTRO}: {n_reg} "
          f"(coincide con las celdas modificadas).")


def verificar(original, salida, cambios, filas_reg_antes, n_reg):
    """Relee la copia guardada y la compara celda por celda con el original."""
    wo = openpyxl.load_workbook(original)
    ns = openpyxl.load_workbook(salida)
    esperados = {(c["hoja"], c["fila"], c["col"]): c["nuevo"] for c in cambios}
    problemas = []
    if wo.sheetnames != ns.sheetnames:
        problemas.append("cambió la lista de hojas")
    for nombre in wo.sheetnames:
        if nombre == registro.HOJA_REGISTRO or nombre not in ns.sheetnames:
            continue
        a, b = wo[nombre], ns[nombre]
        for r in range(1, max(a.max_row, b.max_row) + 1):
            for c in range(1, max(a.max_column, b.max_column) + 1):
                va, vb = a.cell(r, c).value, b.cell(r, c).value
                if (nombre, r, c) in esperados:
                    if vb != esperados[(nombre, r, c)]:
                        problemas.append(f"{nombre}!{get_column_letter(c)}{r} no quedó con el valor nuevo")
                elif tx.valor_comparable(va) != tx.valor_comparable(vb):
                    problemas.append(f"{nombre}!{get_column_letter(c)}{r} cambió inesperadamente")
        if [t.ref for t in a.tables.values()] != [t.ref for t in b.tables.values()]:
            problemas.append(f"cambió una tabla de '{nombre}'")
    reg = ns[registro.HOJA_REGISTRO]
    if reg.max_row - filas_reg_antes != n_reg:
        problemas.append(f"RegistroCambios sumó {reg.max_row - filas_reg_antes} filas, "
                         f"se esperaban {n_reg}")
    if problemas:
        print("\nVERIFICACIÓN FALLIDA (no se reemplazó el libro):")
        for p in problemas[:30]:
            print(f"  - {p}")
        sys.exit(1)
    print("\nVerificación OK: solo cambiaron las celdas previstas y el registro; "
          f"'{HOJA_MARCAS}' y demás hojas, idénticas.")


# --------------------------------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--hoja", required=True, action="append",
                    help="nombre exacto de la hoja (repetible, en pareja con --prefijo)")
    ap.add_argument("--prefijo", required=True, action="append",
                    help='por ejemplo "BICICLETA MTB" (uno por cada --hoja, en el mismo orden)')
    ap.add_argument("--revision", type=Path, help="planilla de revisión")
    ap.add_argument("--archivo", type=Path, default=ARCHIVO_PATH)
    ap.add_argument("--aplicar", action="store_true", help="Fase 3 (dry-run sin --execute)")
    ap.add_argument("--execute", action="store_true", help="con --aplicar: escribe el libro")
    args = ap.parse_args()
    if args.execute and not args.aplicar:
        ap.error("--execute solo se usa junto con --aplicar")
    if len(args.hoja) != len(args.prefijo):
        ap.error("tiene que haber un --prefijo por cada --hoja")
    if len(set(args.hoja)) != len(args.hoja):
        ap.error("hay una --hoja repetida")
    pares = list(zip(args.hoja, args.prefijo))
    if args.revision:
        ruta_rev = args.revision
    elif len(pares) == 1:
        ruta_rev = PROJECT_ROOT / f"revision_nombres_{slug_hoja(args.hoja[0])}.xlsx"
    else:
        ap.error("con varias hojas indicá la planilla con --revision")
    if args.aplicar:
        fase_aplicar(args, pares, ruta_rev)
    else:
        fase_propuesta(args, pares, ruta_rev)


if __name__ == "__main__":
    main()
