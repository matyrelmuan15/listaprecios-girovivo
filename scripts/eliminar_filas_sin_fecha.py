# -*- coding: utf-8 -*-
"""
Elimina filas sin FechaActualizacion editando el .xlsx a nivel XML.

Por qué XML y no openpyxl: openpyxl reescribe la hoja "Marcas Actualización
Dinámica" (pierde valores cacheados, el atributo ca y convierte las 1.332
celdas de derrame <f ca="1"/> en <f/>). Acá esa hoja, y toda parte que no
haga falta tocar, se copia byte a byte.

Partes que SÍ se modifican:
  - xl/worksheets/sheetN.xml de las hojas con filas a eliminar
    (filas, números de fila, fórmulas por fila, dimension, validaciones).
  - xl/tables/tableN.xml de esas hojas (ref, autoFilter, sortState).
  - hoja RegistroCambios (se agregan filas, como texto en línea).
  - xl/calcChain.xml se quita (Excel lo regenera solo al abrir).

USO
  python eliminar_filas_sin_fecha.py ORIGEN SALIDA [--exigir-nombre] [--json informe.json]
  (siempre escribe en SALIDA; nunca pisa ORIGEN)
"""
import argparse, hashlib, html, json, re, sys, warnings, zipfile
from pathlib import Path

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))
import openpyxl  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
import formulas  # noqa: E402
import registro  # noqa: E402

HOJA_MARCAS = "Marcas Actualización Dinámica"
EXCLUIDAS = {"Hoja1", HOJA_MARCAS, "RegistroCambios", "NoActualizados"}
TIPO_ELIMINACION = getattr(registro, "TIPO_ELIMINACION", "Eliminacion")
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class Abortar(Exception):
    pass


def vacia(v):
    return v is None or (isinstance(v, str) and not v.strip())


def col_por_nombre(ws, nombre):
    obj = nombre.strip().casefold()
    for c in ws[1]:
        if c.value is not None and str(c.value).strip().casefold() == obj:
            return c.column
    return None


def sha(b):
    return hashlib.sha256(b).hexdigest()


# ---------------------------------------------------------------------------
# 1. Análisis con openpyxl (solo lectura, sin data_only)
# ---------------------------------------------------------------------------
def analizar(origen, exigir_nombre, conservar=frozenset()):
    wb = openpyxl.load_workbook(origen)
    registro.validar_encabezados(wb)
    plan, vacias, info = {}, {}, {}
    for ws in wb.worksheets:
        if ws.title in EXCLUIDAS:
            continue
        cols = {n: col_por_nombre(ws, n) for n in
                ("FechaActualizacion", "CodigoDux", "NombreProducto", "Marca", "CodigoFactura",
                 "PrecioConIVA", "PrecioContado", "PrecioSinIVA", "DolarReferencia", "ActualizadoSegun")}
        if cols["FechaActualizacion"] is None or (cols["NombreProducto"] is None and cols["CodigoDux"] is None):
            continue
        def v(r, n):
            return ws.cell(r, cols[n]).value if cols[n] else None
        filas = []
        for r in range(2, ws.max_row + 1):
            if not vacia(v(r, "FechaActualizacion")):
                continue
            nom, dux = v(r, "NombreProducto"), v(r, "CodigoDux")
            if all(vacia(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)):
                vacias.setdefault(ws.title, []).append(r)
                continue
            if vacia(nom) and (exigir_nombre or vacia(dux)):
                continue
            if (ws.title, r) in conservar:
                continue
            filas.append({
                "fila": r, "dux": None if vacia(dux) else dux, "nombre": nom,
                "marca": v(r, "Marca"), "codfac": v(r, "CodigoFactura"),
                "coniva": v(r, "PrecioConIVA"), "dolar": v(r, "DolarReferencia"),
                "actualizado": v(r, "ActualizadoSegun"),
            })
        info[ws.title] = {"max_row": ws.max_row, "letra_iva": get_column_letter(cols["PrecioConIVA"]),
                          "tablas": {t.name: t.ref for t in ws.tables.values()}}
        if filas:
            plan[ws.title] = filas
    return wb, plan, vacias, info


# ---------------------------------------------------------------------------
# 2. Utilidades de corrimiento
# ---------------------------------------------------------------------------
def nueva_fila(r, borradas_ordenadas):
    import bisect
    return r - bisect.bisect_left(borradas_ordenadas, r)


def correr_rango(ref, borr):
    """A1:B10 -> rango corrido. Devuelve None si el rango queda vacío."""
    m = re.fullmatch(r"(\$?[A-Z]+\$?)(\d+)(?::(\$?[A-Z]+\$?)(\d+))?", ref)
    if not m:
        raise Abortar(f"Rango no reconocido: {ref}")
    c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    if c2 is None:
        if r1 in borr:
            return None
        return f"{c1}{nueva_fila(r1, borr)}"
    r2 = int(r2)
    import bisect
    n1 = r1 - bisect.bisect_left(borr, r1)          # borradas < r1
    n2 = r2 - bisect.bisect_right(borr, r2)         # borradas <= r2
    if n2 < n1:
        return None
    return f"{c1}{n1}:{c2}{n2}"


def correr_sqref(sqref, borr):
    partes = [correr_rango(p, borr) for p in sqref.split()]
    return " ".join(p for p in partes if p)


# ---------------------------------------------------------------------------
# 3. Transformación de una hoja
# ---------------------------------------------------------------------------
RE_ROW = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.S)


def transformar_hoja(xml, hoja, borradas, letra_iva):
    borr = sorted(borradas)
    if re.search(r"<(conditionalFormatting|mergeCells|hyperlinks|extLst)\b", xml):
        raise Abortar(f"[{hoja}] tiene formato condicional / combinadas / hipervínculos / extLst: "
                      "el motor no los contempla, se aborta.")
    stats = {"eliminadas": 0, "corridas": 0, "formulas_repuestas": 0}

    def fila_sub(m):
        r = int(m.group(1))
        txt = m.group(0)
        if r in borradas:
            stats["eliminadas"] += 1
            return ""
        nr = nueva_fila(r, borr)
        if nr == r:
            return txt
        stats["corridas"] += 1
        txt = re.sub(r'^<row r="\d+"', f'<row r="{nr}"', txt)
        txt = re.sub(r'(<c r="[A-Z]+)' + str(r) + '"', lambda mm: f'{mm.group(1)}{nr}"', txt)

        def f_sub(fm):
            attrs, cuerpo = fm.group(1), html.unescape(fm.group(2))
            if attrs.strip():
                raise Abortar(f"[{hoja}] fila {r}: fórmula con atributos ({attrs}); no contemplado.")
            for gen in (formulas.formula_contado, formulas.formula_siniva):
                if formulas._normalizar_formula("=" + cuerpo) == formulas._normalizar_formula(gen(hoja, letra_iva, r)):
                    stats["formulas_repuestas"] += 1
                    nuevo = gen(hoja, letra_iva, nr)[1:]
                    return f"<f>{html.escape(nuevo, quote=False)}</f>"
            raise Abortar(f"[{hoja}] fila {r}: fórmula no reconocida «{cuerpo}». Se aborta sin escribir.")
        txt = re.sub(r"<f([^>]*)>(.*?)</f>", f_sub, txt)
        if re.search(r"<f[^>]*/>", txt):
            raise Abortar(f"[{hoja}] fila {r}: fórmula compartida/vacía no contemplada.")
        return txt

    cab, resto = xml.split("<sheetData>", 1) if "<sheetData>" in xml else (None, None)
    if cab is None:
        raise Abortar(f"[{hoja}] sin <sheetData>.")
    datos, pie = resto.split("</sheetData>", 1)
    datos = RE_ROW.sub(fila_sub, datos)
    if stats["eliminadas"] != len(borradas):
        raise Abortar(f"[{hoja}] se esperaban {len(borradas)} filas a eliminar y se encontraron {stats['eliminadas']}.")
    # dimension
    def dim_sub(m):
        return f'<dimension ref="{correr_rango(m.group(1), borr)}"/>'
    cab = re.sub(r'<dimension ref="([^"]+)"/>', dim_sub, cab)
    # validaciones de datos
    pie = re.sub(r'(<dataValidation\b[^>]*?\bsqref=")([^"]+)(")',
                 lambda m: m.group(1) + correr_sqref(m.group(2), borr) + m.group(3), pie)
    return cab + "<sheetData>" + datos + "</sheetData>" + pie, stats


def transformar_tabla(xml, borr):
    borr = sorted(borr)
    def ref_sub(m):
        return m.group(1) + correr_rango(m.group(2), borr) + '"'
    return re.sub(r'(\bref=")([^"]+)"', ref_sub, xml)


# ---------------------------------------------------------------------------
# 4. RegistroCambios: filas nuevas como texto en línea
# ---------------------------------------------------------------------------
def celda_xml(ref, valor):
    if valor is None:
        return ""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}"><v>{valor}</v></c>'
    s = html.escape(str(valor), quote=False)
    esp = ' xml:space="preserve"' if s != s.strip() else ""
    return f'<c r="{ref}" t="inlineStr"><is><t{esp}>{s}</t></is></c>'


def agregar_registro(xml, filas_reg):
    ultima = max(int(x) for x in re.findall(r'<row r="(\d+)"', xml))
    nuevas = []
    for i, fila in enumerate(filas_reg, start=ultima + 1):
        celdas = "".join(celda_xml(f"{get_column_letter(j)}{i}", registro._valor_para_celda(fila[h]))
                         for j, h in enumerate(registro.REGISTRO_HEADERS, start=1))
        nuevas.append(f'<row r="{i}" spans="1:10">{celdas}</row>')
    xml = xml.replace("</sheetData>", "".join(nuevas) + "</sheetData>", 1)
    xml = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>',
                 lambda m: f'<dimension ref="A1:{m.group(1)}{ultima + len(filas_reg)}"/>', xml)
    return xml, ultima + 1, ultima + len(filas_reg)


# ---------------------------------------------------------------------------
# 5. Mapa hoja -> parte XML y tablas
# ---------------------------------------------------------------------------
def mapa_partes(z):
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    rid2t.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    hojas = {}
    for nombre, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+"[^>]*r:id="(rId\d+)"', wbx):
        nombre = html.unescape(nombre)
        parte = "xl/" + rid2t[rid].lstrip("/").replace("xl/", "")
        relp = parte.replace("worksheets/", "worksheets/_rels/") + ".rels"
        tablas = []
        if relp in z.namelist():
            for t in re.findall(r'Target="\.\./tables/([^"]+)"', z.read(relp).decode("utf-8")):
                tablas.append("xl/tables/" + t)
        hojas[nombre] = {"parte": parte, "tablas": tablas}
    return hojas


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("origen"); ap.add_argument("salida")
    ap.add_argument("--exigir-nombre", action="store_true",
                    help="Solo elimina filas con NombreProducto (criterio del relevamiento: 240).")
    ap.add_argument("--json")
    ap.add_argument("--conservar", action="append", default=[],
                    help="Hoja:fila que NO se elimina aunque cumpla la regla (repetible).")
    ap.add_argument("--esperado", type=int, help="Total esperado; si no coincide, aborta.")
    ap.add_argument("--fecha-corrida")
    a = ap.parse_args()
    origen, salida = Path(a.origen), Path(a.salida)
    if origen.resolve() == salida.resolve():
        sys.exit("ERROR: la salida no puede ser el mismo archivo de origen.")

    conservar = {(c.rsplit(":", 1)[0], int(c.rsplit(":", 1)[1])) for c in a.conservar}
    wb, plan, vacias, info = analizar(origen, a.exigir_nombre, conservar)
    total = sum(len(v) for v in plan.values())
    if a.esperado is not None and total != a.esperado:
        raise Abortar(f"Total {total} distinto del esperado {a.esperado}: {dict((h, len(v)) for h, v in plan.items())}")

    corrida = registro.RegistroCorrida(TIPO_ELIMINACION,
        "Eliminación de filas sin FechaActualizacion (con NombreProducto o CodigoDux)"
        + (" – criterio: exige NombreProducto" if a.exigir_nombre else ""), fecha=a.fecha_corrida)
    for hoja, filas in plan.items():
        for f in filas:
            marca = "(sin marca)" if vacia(f["marca"]) else f["marca"]
            ident = f"CodigoDux: {f['dux']}" if f["dux"] is not None else "ENCABEZADO"
            datos = (f"Fila original {f['fila']} | CodigoFactura: {f['codfac']} | PrecioConIVA: {f['coniva']}"
                     f" | DolarReferencia: {f['dolar']} | ActualizadoSegun: {f['actualizado']}")
            corrida.filas.append(registro.nueva_fila(
                corrida.fecha, corrida.tipo_operacion, hoja, f["fila"], f["nombre"], f["marca"],
                "Fila completa", f"{ident} | Marca: {marca}", "(fila eliminada)",
                corrida.detalle + " | " + datos))
    corrida.verificar_conteo(total)

    zin = zipfile.ZipFile(origen)
    partes = mapa_partes(zin)
    modificadas, stats = {}, {}
    for hoja, filas in plan.items():
        borr = {f["fila"] for f in filas}
        p = partes[hoja]["parte"]
        xml, st = transformar_hoja(zin.read(p).decode("utf-8"), hoja, borr, info[hoja]["letra_iva"])
        modificadas[p] = xml.encode("utf-8"); stats[hoja] = st
        for t in partes[hoja]["tablas"]:
            modificadas[t] = transformar_tabla(zin.read(t).decode("utf-8"), borr).encode("utf-8")
    p_reg = partes["RegistroCambios"]["parte"]
    xml_reg, desde, hasta = agregar_registro(zin.read(p_reg).decode("utf-8"), corrida.filas)
    modificadas[p_reg] = xml_reg.encode("utf-8")

    # calcChain: se quita (parte, relación y content type)
    quitar = {"xl/calcChain.xml"}
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    modificadas["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*Target="calcChain.xml"/>', "", rels).encode("utf-8")
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    modificadas["[Content_Types].xml"] = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', "", ct).encode("utf-8")

    with zipfile.ZipFile(salida, "w") as zout:
        for item in zin.infolist():
            if item.filename in quitar:
                continue
            data = modificadas.get(item.filename, zin.read(item.filename))
            nuevo = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            nuevo.compress_type, nuevo.external_attr = item.compress_type, item.external_attr
            zout.writestr(nuevo, data)

    parte_marcas = partes[HOJA_MARCAS]["parte"]
    with zipfile.ZipFile(salida) as zchk:
        igual = sha(zchk.read(parte_marcas)) == sha(zin.read(parte_marcas))
        sin_tocar = [n for n in zin.namelist() if n not in quitar and n not in modificadas
                     and zchk.read(n) != zin.read(n)]
    if not igual or sin_tocar:
        salida.unlink()
        raise Abortar(f"Verificación byte a byte falló (Marcas igual={igual}, otras partes alteradas={sin_tocar}).")

    informe = {"total": total, "por_hoja": {h: len(v) for h, v in plan.items()}, "stats": stats,
               "vacias": vacias, "registro_filas": [desde, hasta], "marcas_byte_identica": igual,
               "partes_modificadas": sorted(modificadas), "partes_quitadas": sorted(quitar),
               "plan": plan, "registro": corrida.filas}
    if a.json:
        json.dump(informe, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print(f"Filas a eliminar: {total} | {informe['por_hoja']}")
    print(f"Stats: {stats}")
    print(f"RegistroCambios: filas {desde}-{hasta} | Marcas byte a byte idéntica: {igual}")
    print(f"Partes modificadas: {informe['partes_modificadas']} | quitadas: {sorted(quitar)}")
    print(f"Salida: {salida}")


if __name__ == "__main__":
    try:
        main()
    except (Abortar, registro.RegistroError) as e:
        sys.exit(f"ABORTADO: {e}")
