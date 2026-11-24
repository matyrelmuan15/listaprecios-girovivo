#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mover_fila.py
--------------
Mueve UNA fila de producto de una hoja de destino a otra, editando solo el
XML del .xlsx (la hoja "Marcas Actualización Dinámica" y toda parte que no
haga falta tocar se copian byte a byte).

Qué hace:
  1. Ubica la fila de origen por NombreProducto + CodigoDux (igualdad exacta
     normalizada; tiene que haber exactamente una).
  2. La copia en la hoja de destino, columna por columna según el nombre del
     encabezado, como última fila de producto (CodigoDux no vacío). Toma el
     estilo de cada celda de la fila de producto anterior de la hoja de
     destino. PrecioContado y PrecioSinIVA se escriben con las fórmulas
     obligatorias (formulas.py) apuntando a la hoja y fila de destino.
  3. Si la fila nueva queda fuera de la tabla de destino, amplía la tabla
     (tablas_xml.transformar_tabla).
  4. Elimina la fila completa de origen con el motor de
     eliminar_filas_sin_fecha.py (corre filas, fórmulas, dimension y la ref
     de la tabla). Las validaciones de datos "abiertas" (que llegan más allá
     de la última fila de la hoja, p. ej. J2:J5000) se dejan como estaban.
  5. RegistroCambios: una fila "Movimiento" por celda escrita en destino,
     una "Estructura" si se amplía la tabla de destino y una "Eliminacion"
     por la fila borrada en origen.

Verificaciones (sobre la copia, antes de tocar el archivo real):
  - byte a byte: Marcas idéntica y partes no modificadas idénticas;
  - openpyxl: fila nueva, corrimiento de la hoja de origen, resto de las
    hojas sin cambios, refs de tablas, validaciones, fórmulas y registro;
  - apertura en Excel sin pedido de reparación y refs vistas por Excel.

USO
  python mover_fila.py --origen "Bicicletas MTB" --destino "Bicicletas Gravel" \
      --nombre "BICICLETA GRAVEL ZION AVRA" --dux 295            (dry-run)
  ... --execute                                                     (real)
  --archivo RUTA   trabaja sobre otro .xlsx (por ejemplo una copia)
"""

import argparse
import datetime
import html
import re
import shutil
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path

warnings.filterwarnings("ignore")
SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import openpyxl  # noqa: E402
from openpyxl.utils import get_column_letter, range_boundaries  # noqa: E402

import eliminar_filas_sin_fecha as ef  # noqa: E402
import formulas  # noqa: E402
import registro  # noqa: E402
import tablas_xml as tx  # noqa: E402

PROJECT_ROOT = SCRIPTS.parent
ARCHIVO = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"
TIPO_MOVIMIENTO = getattr(registro, "TIPO_MOVIMIENTO", "Movimiento")

COL_DUX, COL_NOMBRE, COL_MARCA = "CodigoDux", "NombreProducto", "Marca"
COL_IVA, COL_CONTADO, COL_SINIVA = formulas.COL_CONIVA, formulas.COL_CONTADO, formulas.COL_SINIVA

Abortar = tx.Abortar
L = get_column_letter
n = tx.norm


# --------------------------------------------------------------------------
# 1. Análisis (solo lectura)
# --------------------------------------------------------------------------
def encabezados(ws):
    """{nombre normalizado: (columna, nombre)} de la fila 1."""
    enc = {}
    for c in ws[1]:
        if ef.vacia(c.value):
            continue
        k = n(c.value)
        if k in enc:
            raise Abortar(f"[{ws.title}] encabezado repetido: {c.value!r}")
        enc[k] = (c.column, str(c.value).strip())
    return enc


def unica_tabla(ws):
    if len(ws.tables) != 1:
        raise Abortar(f"[{ws.title}] se esperaba exactamente una tabla y hay {len(ws.tables)}.")
    return next(iter(ws.tables.values()))


def analizar(archivo, hoja_o, hoja_d, nombre, dux):
    wb = openpyxl.load_workbook(archivo)  # sin data_only=True; nunca se guarda
    registro.validar_encabezados(wb)
    if hoja_o == hoja_d:
        raise Abortar("La hoja de origen y la de destino son la misma.")
    for h in (hoja_o, hoja_d):
        if h not in wb.sheetnames:
            raise Abortar(f"No existe la hoja {h!r}.")
        if h in tx.HOJAS_EXCLUIDAS:
            raise Abortar(f"La hoja {h!r} está excluida.")
    wo, wd = wb[hoja_o], wb[hoja_d]
    eo, ed = encabezados(wo), encabezados(wd)
    for enc, h in ((eo, hoja_o), (ed, hoja_d)):
        faltan = [c for c in (COL_DUX, COL_NOMBRE, COL_IVA, COL_CONTADO, COL_SINIVA) if n(c) not in enc]
        if faltan:
            raise Abortar(f"[{h}] faltan encabezados: {faltan}")
    solo_o = [eo[k][1] for k in eo if k not in ed]
    solo_d = [ed[k][1] for k in ed if k not in eo]

    # Fila de origen: NombreProducto + CodigoDux exactos
    c_nom, c_dux = eo[n(COL_NOMBRE)][0], eo[n(COL_DUX)][0]
    candidatas = [r for r in range(2, wo.max_row + 1)
                  if n(wo.cell(r, c_nom).value) == n(nombre) and n(wo.cell(r, c_dux).value) == n(dux)]
    if len(candidatas) != 1:
        raise Abortar(f"[{hoja_o}] se esperaba una fila con NombreProducto={nombre!r} y "
                      f"CodigoDux={dux!r}; se encontraron {len(candidatas)}: {candidatas}")
    fila_o = candidatas[0]
    valores = {k: wo.cell(fila_o, col).value for k, (col, _) in eo.items()}
    perdidos = [eo[k][1] for k in eo if k not in ed and not ef.vacia(valores[k])]
    if perdidos:
        raise Abortar(f"Columnas con valor en origen que no existen en destino: {perdidos}")

    precio = valores[n(COL_IVA)]
    if not formulas._es_numero(precio) or precio == 0:
        raise Abortar(f"La fila de origen no tiene un PrecioConIVA numérico válido ({precio!r}).")
    if round(precio, 2) != precio:
        raise Abortar(f"PrecioConIVA con más de 2 decimales ({precio}); corré verificar_formulas.py antes.")
    letra_o = L(eo[n(COL_IVA)][0])
    form_o = {COL_CONTADO: formulas.formula_contado(hoja_o, letra_o, fila_o),
              COL_SINIVA: formulas.formula_siniva(hoja_o, letra_o, fila_o)}
    avisos = []
    for col_f, esperado in form_o.items():
        actual = valores[n(col_f)]
        if formulas._normalizar_formula(actual) != formulas._normalizar_formula(esperado):
            avisos.append(f"En origen, {col_f} no tenía la fórmula esperada ({actual!r}); "
                          "en destino se escribe igual la fórmula obligatoria.")

    # Destino: sin duplicados, última fila de producto + 1, fila libre
    cd_dux = ed[n(COL_DUX)][0]
    dups = [r for r in range(2, wd.max_row + 1) if n(wd.cell(r, cd_dux).value) == n(dux)]
    if dups:
        raise Abortar(f"[{hoja_d}] ya existe el CodigoDux {dux!r} en las filas {dups}.")
    productos = [r for r in range(2, wd.max_row + 1) if not ef.vacia(wd.cell(r, cd_dux).value)]
    fila_d = (max(productos) if productos else 1) + 1
    ocupadas = [wd.cell(fila_d, c).coordinate for c in range(1, wd.max_column + 1)
                if not ef.vacia(wd.cell(fila_d, c).value)]
    if ocupadas:
        raise Abortar(f"[{hoja_d}] la fila {fila_d} no está vacía: {ocupadas}")

    # Tablas
    td, to = unica_tabla(wd), unica_tabla(wo)
    c1, r1, c2, r2 = range_boundaries(td.ref)
    if r1 != 1 or r2 < fila_d - 1:
        raise Abortar(f"[{hoja_d}] la tabla {td.name} ({td.ref}) no cubre las filas de producto.")
    if max(c for c, _ in ed.values()) > c2:
        raise Abortar(f"[{hoja_d}] hay encabezados fuera de la tabla {td.name}; corré ampliar_tablas.py.")
    ref_d_nueva = td.ref if r2 >= fila_d else f"{L(c1)}{r1}:{L(c2)}{fila_d}"
    o1, q1, o2, q2 = range_boundaries(to.ref)
    if not (q1 < fila_o <= q2):
        raise Abortar(f"[{hoja_o}] la fila {fila_o} no está dentro de la tabla {to.name} ({to.ref}).")
    ref_o_nueva = f"{L(o1)}{q1}:{L(o2)}{q2 - 1}"

    # Celdas a escribir en destino
    letra_d = L(ed[n(COL_IVA)][0])
    celdas = []
    for k, (col, nom_col) in sorted(ed.items(), key=lambda x: x[1][0]):
        if k == n(COL_CONTADO):
            v = formulas.formula_contado(hoja_d, letra_d, fila_d)
        elif k == n(COL_SINIVA):
            v = formulas.formula_siniva(hoja_d, letra_d, fila_d)
        else:
            v = valores.get(k)
        if ef.vacia(v):
            continue
        celdas.append({"col": col, "columna": nom_col, "valor": v, "origen": valores.get(k)})

    return {
        "wb": wb, "hoja_o": hoja_o, "hoja_d": hoja_d, "fila_o": fila_o, "fila_d": fila_d,
        "eo": eo, "ed": ed, "solo_o": solo_o, "solo_d": solo_d, "valores": valores,
        "letra_o": letra_o, "letra_d": letra_d, "celdas": celdas, "avisos": avisos,
        "nombre": valores[n(COL_NOMBRE)], "marca": valores.get(n(COL_MARCA)), "dux": valores[n(COL_DUX)],
        "tabla_d": td.name, "ref_d": td.ref, "ref_d_nueva": ref_d_nueva,
        "cols_d": [tc.name for tc in td.tableColumns],
        "tabla_o": to.name, "ref_o": to.ref, "ref_o_nueva": ref_o_nueva,
        "cols_o": [tc.name for tc in to.tableColumns],
    }


# --------------------------------------------------------------------------
# 2. Registro
# --------------------------------------------------------------------------
def armar_registro(p, fecha=None):
    fecha = fecha or registro.timestamp_corrida()
    filas = []
    det = f"Fila movida desde '{p['hoja_o']}' fila {p['fila_o']} (CodigoDux {p['dux']})"
    for c in p["celdas"]:
        extra = ""
        if c["columna"] in (COL_CONTADO, COL_SINIVA):
            extra = f" | fórmula repuesta para la fila de destino (en origen: {c['origen']})"
        filas.append(registro.nueva_fila(fecha, TIPO_MOVIMIENTO, p["hoja_d"], p["fila_d"], p["nombre"],
                                         p["marca"], c["columna"], None, c["valor"], det + extra))
    if p["ref_d_nueva"] != p["ref_d"]:
        filas.append(registro.nueva_fila(
            fecha, registro.TIPO_ESTRUCTURA, p["hoja_d"], 1, None, None, f"Tabla {p['tabla_d']}",
            p["ref_d"], p["ref_d_nueva"],
            f"Ampliación de tabla para incluir la fila {p['fila_d']} movida desde '{p['hoja_o']}'"))
    v = p["valores"]
    marca = "(sin marca)" if ef.vacia(p["marca"]) else p["marca"]
    datos = " | ".join(f"{p['eo'][k][1]}: {v[k]}" for k in p["eo"]
                       if k not in (n(COL_NOMBRE), n(COL_DUX), n(COL_MARCA), n(COL_CONTADO), n(COL_SINIVA)))
    filas.append(registro.nueva_fila(
        fecha, registro.TIPO_ELIMINACION, p["hoja_o"], p["fila_o"], p["nombre"], p["marca"],
        "Fila completa", f"CodigoDux: {p['dux']} | Marca: {marca}", "(fila eliminada)",
        f"Fila movida a '{p['hoja_d']}' fila {p['fila_d']} | {datos} | "
        f"Tabla {p['tabla_o']}: {p['ref_o']} -> {p['ref_o_nueva']}"))
    esperado = len(p["celdas"]) + (p["ref_d_nueva"] != p["ref_d"]) + 1
    registro.verificar_conteo(filas, esperado)
    return filas


# --------------------------------------------------------------------------
# 3. Transformaciones XML
# --------------------------------------------------------------------------
def celda_xml(ref, valor, estilo):
    s = f' s="{estilo}"' if estilo else ""
    if isinstance(valor, str) and valor.startswith("="):
        return f'<c r="{ref}"{s}><f>{html.escape(valor[1:], quote=False)}</f><v /></c>'
    if isinstance(valor, bool) or isinstance(valor, (datetime.date, datetime.datetime)):
        raise Abortar(f"{ref}: tipo de valor no contemplado ({valor!r}).")
    if isinstance(valor, (int, float)):
        return f'<c r="{ref}"{s} t="n"><v>{valor}</v></c>'
    texto = html.escape(str(valor), quote=False)
    esp = ' xml:space="preserve"' if texto != texto.strip() else ""
    return f'<c r="{ref}"{s} t="inlineStr"><is><t{esp}>{texto}</t></is></c>'


def estilos_fila(datos, fila):
    m = re.search(rf'<row r="{fila}"[^>]*?(?:/>|>(.*?)</row>)', datos, re.S)
    if not m or not m.group(1):
        return {}
    return {letra: s for letra, s in
            re.findall(r'<c r="([A-Z]+)\d+"(?:\s+s="(\d+)")?', m.group(1))}


def escribir_fila(xml, p):
    cab, resto = xml.split("<sheetData>", 1)
    datos, pie = resto.split("</sheetData>", 1)
    fila = p["fila_d"]
    estilos = estilos_fila(datos, fila - 1) if fila > 2 else {}
    cont = "".join(celda_xml(f"{L(c['col'])}{fila}", c["valor"], estilos.get(L(c["col"])))
                   for c in p["celdas"])
    m = re.search(rf'<row r="{fila}"([^>]*?)(?:/>|>(.*?)</row>)', datos, re.S)
    if m:
        if m.group(2) and m.group(2).strip():
            raise Abortar(f"[{p['hoja_d']}] la fila {fila} tiene celdas en el XML.")
        datos = datos[:m.start()] + f'<row r="{fila}"{m.group(1).rstrip()}>{cont}</row>' + datos[m.end():]
    else:
        nueva = f'<row r="{fila}">{cont}</row>'
        sig = next((mm for mm in re.finditer(r'<row r="(\d+)"', datos) if int(mm.group(1)) > fila), None)
        datos = datos[:sig.start()] + nueva + datos[sig.start():] if sig else datos + nueva

    def dim_sub(mm):
        return mm.group(1) + str(max(int(mm.group(2)), fila)) + mm.group(3)
    cab = re.sub(r'(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)(")', dim_sub, cab, count=1)
    return cab + "<sheetData>" + datos + "</sheetData>" + pie, estilos


def restaurar_validaciones_abiertas(nuevo, original):
    """Si todas las validaciones de la hoja llegan más allá de su última fila
    (p. ej. J2:J5000), se deja el bloque original en vez del corrido."""
    m_o = re.search(r"<dataValidations\b.*?</dataValidations>", original, re.S)
    if not m_o:
        return nuevo, False
    ultima = max(int(x) for x in re.findall(r'<row r="(\d+)"', original))
    partes = [pp for s in re.findall(r'<dataValidation\b[^>]*?\bsqref="([^"]+)"', m_o.group(0))
              for pp in s.split()]
    if not all(int(re.search(r"(\d+)$", pp).group(1)) > ultima for pp in partes):
        return nuevo, False
    m_n = re.search(r"<dataValidations\b.*?</dataValidations>", nuevo, re.S)
    return nuevo[:m_n.start()] + m_o.group(0) + nuevo[m_n.end():], True


def construir(archivo, salida, p, filas_reg):
    z = zipfile.ZipFile(archivo)
    if "xl/calcChain.xml" in z.namelist():
        raise Abortar("El libro tiene xl/calcChain.xml; este script no lo contempla.")
    partes = tx.mapa_partes(z)
    pd_, po_ = partes[p["hoja_d"]], partes[p["hoja_o"]]
    mod = {}

    xml_d, estilos = escribir_fila(z.read(pd_["parte"]).decode("utf-8"), p)
    mod[pd_["parte"]] = xml_d.encode("utf-8")
    if p["ref_d_nueva"] != p["ref_d"]:
        parte_t = pd_["tablas"][p["tabla_d"]]
        plan = {"ref_nueva": p["ref_d_nueva"], "filtros": [], "nuevas": [],
                "columnas_finales": p["cols_d"], "tabla": p["tabla_d"]}
        mod[parte_t] = tx.transformar_tabla(z.read(parte_t).decode("utf-8"), plan).encode("utf-8")

    orig_o = z.read(po_["parte"]).decode("utf-8")
    xml_o, stats = ef.transformar_hoja(orig_o, p["hoja_o"], {p["fila_o"]}, p["letra_o"])
    xml_o, val_ok = restaurar_validaciones_abiertas(xml_o, orig_o)
    mod[po_["parte"]] = xml_o.encode("utf-8")
    parte_to = po_["tablas"][p["tabla_o"]]
    txml_o = ef.transformar_tabla(z.read(parte_to).decode("utf-8"), {p["fila_o"]})
    ref_vista = re.search(r'<table\b[^>]*?\bref="([^"]+)"', txml_o).group(1)
    if ref_vista != p["ref_o_nueva"]:
        raise Abortar(f"La tabla {p['tabla_o']} quedaría {ref_vista} (esperada {p['ref_o_nueva']}).")
    mod[parte_to] = txml_o.encode("utf-8")

    p_reg = partes[registro.HOJA_REGISTRO]["parte"]
    xml_reg, desde, hasta = tx.agregar_registro(z.read(p_reg).decode("utf-8"), filas_reg)
    mod[p_reg] = xml_reg.encode("utf-8")

    with zipfile.ZipFile(salida, "w") as zout:
        for item in z.infolist():
            nuevo = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            nuevo.compress_type, nuevo.external_attr = item.compress_type, item.external_attr
            zout.writestr(nuevo, mod.get(item.filename, z.read(item.filename)))
    return z, partes, mod, desde, hasta, {"estilos": estilos, "corrimiento": stats,
                                          "validaciones_sin_correr": val_ok}


# --------------------------------------------------------------------------
# 4. Verificación openpyxl
# --------------------------------------------------------------------------
def verificar_openpyxl(p, salida, filas_reg, desde, base_formulas):
    wo_, errores = p["wb"], []
    wb = openpyxl.load_workbook(salida)
    hd, ho, fd, fo = p["hoja_d"], p["hoja_o"], p["fila_d"], p["fila_o"]
    vc = tx.valor_comparable

    def fila(ws, r, ncols):
        return [vc(ws.cell(r, c).value) for c in range(1, ncols + 1)]

    # Destino
    a, b = wo_[hd], wb[hd]
    nc = max(a.max_column, b.max_column)
    for r in range(1, fd):
        if fila(a, r, nc) != fila(b, r, nc):
            errores.append(f"[{hd}] fila {r} cambió")
    esperada = [None] * nc
    for c in p["celdas"]:
        esperada[c["col"] - 1] = c["valor"]
    if fila(b, fd, nc) != esperada:
        errores.append(f"[{hd}] fila {fd}: {fila(b, fd, nc)} != {esperada}")
    if b.max_row != max(a.max_row, fd):
        errores.append(f"[{hd}] max_row {b.max_row} (esperado {max(a.max_row, fd)})")

    # Origen: filas anteriores iguales, posteriores corridas una fila arriba
    a, b = wo_[ho], wb[ho]
    nc = max(a.max_column, b.max_column)
    for r in range(1, fo):
        if fila(a, r, nc) != fila(b, r, nc):
            errores.append(f"[{ho}] fila {r} cambió")
    gens = (formulas.formula_contado, formulas.formula_siniva)
    for r in range(fo, a.max_row):
        esperado = []
        for v in fila(a, r + 1, nc):
            for g in gens:
                if isinstance(v, str) and formulas._normalizar_formula(v) == \
                        formulas._normalizar_formula(g(ho, p["letra_o"], r + 1)):
                    v = g(ho, p["letra_o"], r)
                    break
            esperado.append(v)
        if fila(b, r, nc) != esperado:
            errores.append(f"[{ho}] fila {r} no coincide con la fila original {r + 1}")
            break
    if b.max_row != a.max_row - 1:
        errores.append(f"[{ho}] max_row {b.max_row} (esperado {a.max_row - 1})")

    # Resto de las hojas sin cambios; tablas; validaciones
    for ws in wb.worksheets:
        if ws.title not in (hd, ho, registro.HOJA_REGISTRO):
            if [tuple(map(vc, f)) for f in wo_[ws.title].iter_rows(values_only=True)] != \
               [tuple(map(vc, f)) for f in ws.iter_rows(values_only=True)]:
                errores.append(f"[{ws.title}] cambiaron valores")
        for t in ws.tables.values():
            esp = {p["tabla_d"]: p["ref_d_nueva"], p["tabla_o"]: p["ref_o_nueva"]}.get(
                t.name, wo_[ws.title].tables[t.name].ref if t.name in wo_[ws.title].tables else None)
            if t.ref != esp:
                errores.append(f"{t.name}: ref {t.ref} (esperada {esp})")
            c1, r1, c2, _ = range_boundaries(t.ref)
            if [tc.name for tc in t.tableColumns] != [ws.cell(r1, c).value for c in range(c1, c2 + 1)]:
                errores.append(f"{t.name}: columnas distintas de los encabezados")
        if ws.title != registro.HOJA_REGISTRO:
            antes = [(str(dv.sqref), dv.formula1) for dv in wo_[ws.title].data_validations.dataValidation]
            ahora = [(str(dv.sqref), dv.formula1) for dv in ws.data_validations.dataValidation]
            if antes != ahora:
                errores.append(f"[{ws.title}] validaciones {antes} -> {ahora}")

    # Fórmulas obligatorias en todo el libro: igual que antes
    cambios, revisar = formulas.revisar(wb)
    if (len(cambios), len(revisar)) != base_formulas:
        errores.append(f"formulas.revisar: {len(cambios)} cambios / {len(revisar)} para revisar "
                       f"(antes {base_formulas})")

    # RegistroCambios
    reg = wb[registro.HOJA_REGISTRO]
    if reg.max_row != desde - 1 + len(filas_reg):
        errores.append(f"RegistroCambios: {reg.max_row} filas (esperadas {desde - 1 + len(filas_reg)})")
    for i, f in enumerate(filas_reg):
        leidas = [reg.cell(desde + i, j).value for j in range(1, 11)]
        esperadas = [registro._valor_para_celda(f[h]) for h in registro.REGISTRO_HEADERS]
        if leidas != esperadas:
            errores.append(f"RegistroCambios fila {desde + i}: {leidas} != {esperadas}")
    return not errores, errores, wb


# --------------------------------------------------------------------------
# 5. Programa
# --------------------------------------------------------------------------
def backup_archivo(archivo):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo.stem}_{ts}{archivo.suffix}"
    shutil.copy2(archivo, destino)
    return destino


def mostrar(v):
    return "(vacío)" if v is None else repr(v)


def correr(archivo, hoja_o, hoja_d, nombre, dux, execute):
    sha_origen = tx.sha(archivo.read_bytes())
    p = analizar(archivo, hoja_o, hoja_d, nombre, dux)
    base = formulas.revisar(p["wb"])
    base_formulas = (len(base[0]), len(base[1]))

    print("=== Fase 1: Inspección ===")
    print(f"Archivo: {archivo}")
    print(f"Origen:  [{hoja_o}] fila {p['fila_o']} | CodigoDux {p['dux']!r} | {p['nombre']}")
    print(f"Destino: [{hoja_d}] fila {p['fila_d']} (última fila de producto + 1)")
    print(f"Columnas solo en origen: {p['solo_o'] or 'ninguna'} | solo en destino: {p['solo_d'] or 'ninguna'}")
    print("\nValores (origen -> destino):")
    escritas = {c["columna"]: c for c in p["celdas"]}
    for k, (col, nom_col) in sorted(p["eo"].items(), key=lambda x: x[1][0]):
        dest = escritas.get(nom_col)
        cd = p["ed"].get(k)
        donde = f"{L(cd[0])}{p['fila_d']}" if cd else "(no existe)"
        print(f"  {nom_col:18} {L(col)}{p['fila_o']}: {mostrar(p['valores'][k]):62} -> {donde}: "
              f"{mostrar(dest['valor']) if dest else '(no se escribe, vacío)'}")
    print(f"\nTabla destino {p['tabla_d']}: {p['ref_d']} -> {p['ref_d_nueva']}"
          + ("" if p["ref_d_nueva"] != p["ref_d"] else " (sin cambios)"))
    print(f"Tabla origen  {p['tabla_o']}: {p['ref_o']} -> {p['ref_o_nueva']}")
    for a in p["avisos"]:
        print(f"AVISO: {a}")
    print(f"Fórmulas del libro antes: {base_formulas[0]} a corregir / {base_formulas[1]} para revisar")

    filas_reg = armar_registro(p)
    n_mov = sum(1 for f in filas_reg if f["TipoOperacion"] == TIPO_MOVIMIENTO)
    n_est = sum(1 for f in filas_reg if f["TipoOperacion"] == registro.TIPO_ESTRUCTURA)
    n_eli = sum(1 for f in filas_reg if f["TipoOperacion"] == registro.TIPO_ELIMINACION)
    print(f"\nRegistroCambios: {len(filas_reg)} filas = {n_mov} '{TIPO_MOVIMIENTO}' + "
          f"{n_est} '{registro.TIPO_ESTRUCTURA}' + {n_eli} '{registro.TIPO_ELIMINACION}'")

    print("\n=== Fase 2: Simulación sobre copia temporal ===")
    carpeta = Path(tempfile.mkdtemp(prefix="mover_fila_"))
    salida = carpeta / archivo.name
    z, partes, mod, desde, hasta, info = construir(archivo, salida, p, filas_reg)
    print(f"Copia simulada: {salida}")
    print(f"Partes modificadas: {sorted(mod)}")
    print(f"Estilos tomados de la fila {p['fila_d'] - 1} de destino: {info['estilos']}")
    print(f"Corrimiento en origen: {info['corrimiento']} | validaciones abiertas sin correr: "
          f"{info['validaciones_sin_correr']}")
    print(f"RegistroCambios: filas {desde}-{hasta}")
    registro.imprimir_previsualizacion(type("C", (), {"filas": filas_reg, "__len__": lambda s: len(filas_reg)})(),
                                       limite=len(filas_reg))

    print("\n--- Verificación 1: byte a byte ---")
    ok1, det1 = tx.verificar_bytes(z, salida, mod, partes)
    print(f"  {'OK' if ok1 else 'ERROR'} | {det1}")

    print("\n--- Verificación 2: openpyxl ---")
    ok2, err2, wb2 = verificar_openpyxl(p, salida, filas_reg, desde, base_formulas)
    print("  Fila nueva, corrimiento, otras hojas, tablas, validaciones, fórmulas y registro: "
          + ("OK" if ok2 else "ERROR"))
    for e in err2:
        print(f"  ERROR: {e}")
    ws = wb2[hoja_d]
    print(f"  [{hoja_d}] fila {p['fila_d']}: " + " | ".join(
        f"{L(c)}={ws.cell(p['fila_d'], c).value}" for c in range(1, ws.max_column + 1)
        if ws.cell(p['fila_d'], c).value is not None))
    wso = wb2[hoja_o]
    print(f"  [{hoja_o}] fila {p['fila_o']} ahora: {wso.cell(p['fila_o'], 3).value} | "
          f"F={wso.cell(p['fila_o'], 6).value}")

    print("\n--- Verificación 3: Excel ---")
    planes = [{"tabla": p["tabla_d"], "ref_nueva": p["ref_d_nueva"], "nuevas": [],
               "columnas_finales": p["cols_d"]},
              {"tabla": p["tabla_o"], "ref_nueva": p["ref_o_nueva"], "nuevas": [],
               "columnas_finales": p["cols_o"]}]
    ok3, res, lin, intentos = tx.verificar_en_excel(salida, carpeta, planes, exigir_sin_filtros=False,
                                                    requeridas={p["tabla_d"], p["tabla_o"]})
    print(("  Excel lo abrió sin pedido de reparación." if res.get("ok") else "  Excel NO pudo abrirlo.")
          + (f" (intentos: {intentos})" if intentos > 1 else ""))
    for linea in lin:
        if linea.startswith("  ERROR") or linea.startswith("Excel:") or \
                f" {p['tabla_d']} " in linea or f" {p['tabla_o']} " in linea:
            print(linea)
    if res.get("ok"):
        tablas = res["tablas"] if isinstance(res["tablas"], list) else [res["tablas"]]
        for t in tablas:
            if t["tabla"] in (p["tabla_d"], p["tabla_o"]) and t["visibles"] != t["filas"]:
                ok3 = False
                print(f"  ERROR: {t['tabla']} tiene filas ocultas ({t['visibles']}/{t['filas']})")
    print(f"  {'OK' if ok3 else 'ERROR'}")
    z.close()

    if not (ok1 and ok2 and ok3):
        raise Abortar("alguna verificación falló; no se toca el archivo real.")
    if not execute:
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo real ni se creó backup.")
        return

    print("\n=== Fase 3: Ejecución ===")
    if tx.sha(archivo.read_bytes()) != sha_origen:
        raise Abortar("el archivo real cambió durante la corrida; volvé a correr el script.")
    backup = backup_archivo(archivo)
    print(f"Backup creado en: {backup}")
    try:
        shutil.copyfile(salida, archivo)
    except PermissionError:
        raise Abortar("no se pudo reemplazar el archivo porque está abierto en Excel. "
                      "Cerralo y volvé a correr el script. El archivo real no se modificó.")
    if tx.sha(archivo.read_bytes()) != tx.sha(salida.read_bytes()):
        raise Abortar(f"el archivo guardado no coincide con el verificado. Restaurá el backup: {backup}")
    with zipfile.ZipFile(backup) as zb, zipfile.ZipFile(archivo) as zr:
        parte_m = tx.mapa_partes(zb)[tx.HOJA_MARCAS]["parte"]
        igual = zb.read(parte_m) == zr.read(parte_m)
    print(f"Archivo reemplazado: {archivo}")
    print(f"Marcas idéntica byte a byte contra el backup: {'sí' if igual else 'NO'}")
    print(f"Filas agregadas a {registro.HOJA_REGISTRO}: {len(filas_reg)} ({desde}-{hasta})")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Mueve una fila de producto de una hoja a otra.")
    ap.add_argument("--origen", required=True, help="Hoja de origen (nombre exacto).")
    ap.add_argument("--destino", required=True, help="Hoja de destino (nombre exacto).")
    ap.add_argument("--nombre", required=True, help="NombreProducto de la fila a mover.")
    ap.add_argument("--dux", required=True, help="CodigoDux de la fila a mover.")
    ap.add_argument("--archivo", default=str(ARCHIVO))
    ap.add_argument("--execute", action="store_true", help="Reemplaza el archivo real (con backup).")
    a = ap.parse_args()
    try:
        correr(Path(a.archivo), a.origen, a.destino, a.nombre, a.dux, a.execute)
    except (Abortar, ef.Abortar, registro.RegistroError) as e:
        sys.exit(f"ABORTADO: {e}")


if __name__ == "__main__":
    main()
