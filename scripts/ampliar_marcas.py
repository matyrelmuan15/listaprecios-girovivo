# -*- coding: utf-8 -*-
"""
Hace tolerantes al tamaño las 6 fórmulas matriciales de "Marcas Actualización
Dinámica", editando SOLO xl/worksheets/<hoja Marcas>.xml (el resto del paquete se
copia byte a byte).

Cambios:
  J1, K1, L1  ref -> X1:X5000   EXPAND(TAKE(VSTACK(...),5000),5000,,"")
  M1          ref -> M1:M5000   EXPAND(TAKE(MAP(...),5000),5000,,"")
  A2          ref -> A2:C201    EXPAND(TAKE(LET(...),200),200,3,"")
  E3          ref -> E3:H3000   EXPAND(TAKE(LET(...),2998),2998,4,"")
  - Se conservan los atributos de cada celda maestra (estilo, tipo, ca="1").
  - Las 1.332 celdas de derrame <f ca="1"/> no se tocan.
  - Celdas con #N/A cacheado dentro de J:M pasan a "" cacheado (es lo que
    devuelve la fórmula nueva), para que Excel muestre el valor correcto aun
    antes de recalcular.
  - <dimension> se amplía a A1:M5000.
USO: python ampliar_marcas.py ORIGEN SALIDA
"""
import hashlib, html, re, sys, zipfile
sys.stdout.reconfigure(encoding="utf-8")

HOJA = "Marcas Actualización Dinámica"
N_AUX, N_MARCAS, N_PROD = 5000, 200, 2998

ESPERADO_INICIO = {  # celda maestra -> prefijo del texto actual (verificación)
    "J1": "_xlfn.VSTACK(Table_1[NombreProducto]",
    "K1": "_xlfn.VSTACK(Table_1[Marca]",
    "L1": "_xlfn.VSTACK(Table_1[FechaActualizacion]",
    "M1": "_xlfn.MAP(_xlfn.ANCHORARRAY(L1)",
    "A2": "_xlfn.LET(_xlpm.inc,ISTEXT(_xlfn.ANCHORARRAY(K1))",
    "E3": "_xlfn.LET(_xlpm.inc,ISTEXT(_xlfn.ANCHORARRAY(J1))",
}
NUEVO = {
    "J1": ("J1:J5000", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_AUX}),{N_AUX},,"")'),
    "K1": ("K1:K5000", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_AUX}),{N_AUX},,"")'),
    "L1": ("L1:L5000", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_AUX}),{N_AUX},,"")'),
    "M1": ("M1:M5000", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_AUX}),{N_AUX},,"")'),
    "A2": ("A2:C201", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_MARCAS}),{N_MARCAS},3,"")'),
    "E3": ("E3:H3000", lambda t: f'_xlfn.EXPAND(_xlfn.TAKE({t},{N_PROD}),{N_PROD},4,"")'),
}


class Abortar(Exception):
    pass


def parte_de_hoja(z):
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m = re.search(r'<sheet name="' + re.escape(HOJA) + r'" sheetId="\d+"[^>]*r:id="(rId\d+)"', wbx)
    if not m:
        raise Abortar("No se encontró la hoja Marcas en workbook.xml")
    t = re.search(r'<Relationship [^>]*Id="' + m.group(1) + r'"[^>]*Target="([^"]+)"', rels) or \
        re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="' + m.group(1) + '"', rels)
    return "xl/" + t.group(1).lstrip("/").replace("xl/", "")


def transformar(xml):
    hechos = {}

    def sub_maestra(m):
        ref, attrs_c, attrs_f, texto = m.group(1), m.group(2), m.group(3), m.group(4)
        if ref not in NUEVO:
            return m.group(0)
        if not html.unescape(texto).startswith(ESPERADO_INICIO[ref]):
            raise Abortar(f"{ref}: el texto actual no es el esperado; no se modifica nada.")
        if "_xlfn.EXPAND(" in texto:
            raise Abortar(f"{ref}: ya parece modificada (contiene EXPAND).")
        nuevo_ref, envolver = NUEVO[ref]
        attrs_f2, n = re.subn(r'\bref="[^"]+"', f'ref="{nuevo_ref}"', attrs_f)
        if n != 1 or 't="array"' not in attrs_f:
            raise Abortar(f"{ref}: no es una fórmula matricial con ref.")
        hechos[ref] = nuevo_ref
        return f'<c r="{ref}"{attrs_c}><f{attrs_f2}>{envolver(texto)}</f>'

    xml2 = re.sub(r'<c r="([A-Z]+\d+)"([^>]*)><f([^>]*t="array"[^>]*)>(.*?)</f>', sub_maestra, xml)
    if set(hechos) != set(NUEVO):
        raise Abortar(f"Se esperaban 6 maestras y se modificaron {sorted(hechos)}")

    # #N/A cacheados en J:M (filas 2..5000) -> "" cacheado
    na = 0

    def sub_na(m):
        nonlocal na
        col, fila = m.group(1), int(m.group(2))
        if col in "JKLM" and 1 <= fila <= N_AUX and m.group(4) == "#N/A":
            na += 1
            return f'<c r="{col}{fila}"{m.group(3)} t="str"><v></v></c>'
        raise Abortar(f"Error cacheado inesperado en {col}{fila}: {m.group(4)}")
    xml2 = re.sub(r'<c r="([A-Z]+)(\d+)"([^>]*?) t="e"><v>([^<]*)</v></c>', sub_na, xml2)

    xml2, n = re.subn(r'<dimension ref="A1:M3000"/>', '<dimension ref="A1:M5000"/>', xml2)
    if n != 1:
        raise Abortar("dimension inesperada")
    return xml2, hechos, na


def main():
    origen, salida = sys.argv[1], sys.argv[2]
    zin = zipfile.ZipFile(origen)
    parte = parte_de_hoja(zin)
    xml = zin.read(parte).decode("utf-8")
    antes_f = re.findall(r"<f[^>]*/>|<f[^>]*>.*?</f>", xml)
    xml2, hechos, na = transformar(xml)
    despues_f = re.findall(r"<f[^>]*/>|<f[^>]*>.*?</f>", xml2)
    if len(antes_f) != len(despues_f):
        raise Abortar("cambió la cantidad de fórmulas")
    distintas = [i for i, (a, b) in enumerate(zip(antes_f, despues_f)) if a != b]
    if len(distintas) != 6:
        raise Abortar(f"se esperaban 6 fórmulas distintas y hay {len(distintas)}")
    with zipfile.ZipFile(salida, "w") as zout:
        for it in zin.infolist():
            data = xml2.encode("utf-8") if it.filename == parte else zin.read(it.filename)
            ni = zipfile.ZipInfo(it.filename, date_time=it.date_time)
            ni.compress_type, ni.external_attr = it.compress_type, it.external_attr
            zout.writestr(ni, data)
    with zipfile.ZipFile(salida) as zc:
        otras = [n for n in zin.namelist() if n != parte and zc.read(n) != zin.read(n)]
        if otras or sorted(zc.namelist()) != sorted(zin.namelist()):
            raise Abortar(f"partes alteradas: {otras}")
    print(f"Parte modificada: {parte} | maestras: {hechos} | #N/A cacheados -> \"\": {na}")
    print(f"<f> antes {len(antes_f)} después {len(despues_f)} | resto del paquete byte a byte idéntico")


if __name__ == "__main__":
    try:
        main()
    except Abortar as e:
        sys.exit(f"ABORTADO: {e}")
