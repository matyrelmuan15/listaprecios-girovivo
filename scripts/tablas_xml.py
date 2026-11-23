#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tablas_xml.py
--------------
Módulo compartido para ampliar tablas de Excel (Table_N) y quitar filtros
guardados editando SOLO el XML del .xlsx, más las tres verificaciones del
resultado. Lo usan ampliar_tablas.py y asignar_categoria.py.

Regla del proyecto: toda columna de una hoja de destino tiene que quedar
dentro de la tabla de esa hoja.

  analizar(...)           plan por tabla: ref nueva, columnas que entran,
                          filtros con criterio y filas ocultas por filtro.
  construir(...)          copia el .xlsx parte por parte aplicando el plan
                          (tablas, filas ocultas y, opcionalmente, filas de
                          RegistroCambios como texto en línea).
  verificar_openpyxl(...) refs, columnas, filtros, valores y registro.
  verificar_bytes(...)    Marcas y las partes no modificadas, byte a byte.
  abrir_en_excel(...) +
  evaluar_excel(...)      apertura en Excel sin reparación, tablas, filtros
                          y formato visible de los encabezados.
"""

import hashlib
import html
import json
import re
import subprocess
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula

import registro

HOJA_MARCAS = "Marcas Actualización Dinámica"
HOJAS_EXCLUIDAS = {"Hoja1", HOJA_MARCAS, registro.HOJA_REGISTRO,
                   "NoActualizados", "Venzo Sin Actualizar"}
TIMEOUT_EXCEL = 300  # segundos

# Elementos que convierten un filterColumn en un criterio real de filtro.
CRITERIOS = ("<filters", "<customFilters", "<top10", "<dynamicFilter",
             "<colorFilter", "<iconFilter")


class Abortar(Exception):
    pass


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def sha(datos):
    return hashlib.sha256(datos).hexdigest()


def norm(v):
    return "" if v is None else str(v).strip().casefold()


def valor_comparable(v):
    if isinstance(v, ArrayFormula):
        return ("ArrayFormula", v.ref, v.text)
    return v


def attr(tag, nombre):
    m = re.search(rf'\b{nombre}="([^"]*)"', tag)
    return html.unescape(m.group(1)) if m else None


def mapa_partes(z):
    """{hoja: {"parte": xl/worksheets/sheetN.xml, "tablas": {nombre: xl/tables/tableN.xml}}}"""
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for tag in re.findall(r"<Relationship\b[^>]*>", rels):
        rid2t[attr(tag, "Id")] = attr(tag, "Target")
    hojas = {}
    for tag in re.findall(r"<sheet\b[^>]*>", wbx):
        nombre = attr(tag, "name")
        rid = re.search(r'r:id="([^"]+)"', tag).group(1)
        destino = rid2t[rid]
        parte = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
        relp = parte.replace("worksheets/", "worksheets/_rels/") + ".rels"
        tablas = {}
        if relp in z.namelist():
            for t in re.findall(r'Target="([^"]*tables/[^"]+)"', z.read(relp).decode("utf-8")):
                tparte = t.lstrip("/") if t.startswith("/") else "xl/tables/" + t.split("tables/")[-1]
                txml = z.read(tparte).decode("utf-8")
                tablas[attr(re.search(r"<table\b[^>]*>", txml).group(0), "name")] = tparte
        hojas[nombre] = {"parte": parte, "tablas": tablas}
    return hojas


def describir_criterio(nombre_col, interior):
    """Texto legible del criterio de un filterColumn."""
    m = re.search(r"<filters\b([^>]*)>(.*?)</filters>|<filters\b([^>]*)/>", interior, re.S)
    if m:
        cuerpo = m.group(2) or ""
        vals = [html.unescape(v) for v in re.findall(r'<filter\b[^>]*\bval="([^"]*)"', cuerpo)]
        if attr("<x " + (m.group(1) or m.group(3) or "") + ">", "blank") in ("1", "true"):
            vals.append("(vacías)")
        return f"{nombre_col} = " + " o ".join(vals), vals
    m = re.search(r"<customFilters\b([^>]*)>(.*?)</customFilters>", interior, re.S)
    if m:
        y = attr("<x " + m.group(1) + ">", "and") in ("1", "true")
        ops = {"equal": "=", "notEqual": "<>", "greaterThan": ">", "greaterThanOrEqual": ">=",
               "lessThan": "<", "lessThanOrEqual": "<="}
        partes = []
        for tag in re.findall(r"<customFilter\b[^>]*>", m.group(2)):
            op = ops.get(attr(tag, "operator") or "equal", attr(tag, "operator"))
            partes.append(f'{nombre_col} {op} "{attr(tag, "val")}"')
        return (" y " if y else " o ").join(partes), None
    return f"{nombre_col}: {interior.strip()}", None


# --------------------------------------------------------------------------
# 1. Análisis (solo lectura)
# --------------------------------------------------------------------------
def analizar(archivo, z, partes, hojas=None, quitar_filtros=True):
    """Arma el plan por tabla.

    hojas:          si se indica, solo se analizan esas hojas.
    quitar_filtros: si es False, los filtros guardados se dejan como están
                    (no se planifica quitarlos ni mostrar filas).
    Devuelve (wb, planes, avisos); wb queda abierto solo para lectura.
    """
    wb = openpyxl.load_workbook(archivo)  # sin data_only=True; nunca se guarda
    registro.validar_encabezados(wb)

    planes, avisos = [], []
    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS or (hojas is not None and ws.title not in hojas):
            continue
        xml_hoja = z.read(partes[ws.title]["parte"]).decode("utf-8")
        ocultas_hoja = set()
        for tag in re.findall(r"<row\b[^>]*>", xml_hoja):
            if attr(tag, "hidden") in ("1", "true"):
                ocultas_hoja.add(int(attr(tag, "r")))
        if re.search(r"<autoFilter\b[^>]*>.*?<filterColumn", xml_hoja, re.S):
            avisos.append(f"{ws.title}: tiene un autofiltro de hoja (fuera de tabla) con criterios; "
                          "no se toca, revisalo a mano.")

        for t in ws.tables.values():
            c1, r1, c2, r2 = range_boundaries(t.ref)
            parte_t = partes[ws.title]["tablas"].get(t.name)
            if parte_t is None:
                raise Abortar(f"No se encontró el XML de la tabla {t.name} ({ws.title}).")
            txml = z.read(parte_t).decode("utf-8")

            # Columnas actuales de la tabla vs encabezados de la hoja
            enc = {c.column: c.value for c in ws[r1]
                   if c.value is not None and str(c.value).strip() != ""}
            actuales = [tc.name for tc in t.tableColumns]
            esperados = [ws.cell(row=r1, column=c).value for c in range(c1, c2 + 1)]
            if [norm(a) for a in actuales] != [norm(e) for e in esperados]:
                raise Abortar(f"{ws.title} / {t.name}: las columnas de la tabla {actuales} no "
                              f"coinciden con los encabezados de la hoja {esperados}.")
            ultima = max(enc) if enc else c2
            nuevas = []
            for col in range(c2 + 1, ultima + 1):
                nombre = enc.get(col)
                if nombre is None:
                    raise Abortar(f"{ws.title}: la celda {get_column_letter(col)}{r1} está vacía entre "
                                  f"la tabla {t.name} y el último encabezado; no se amplía.")
                nuevas.append((col, str(nombre).strip()))
            todos = [norm(a) for a in actuales] + [norm(n) for _, n in nuevas]
            if len(set(todos)) != len(todos):
                raise Abortar(f"{ws.title} / {t.name}: hay encabezados repetidos: {todos}")
            ref_nueva = f"{get_column_letter(c1)}{r1}:{get_column_letter(max(c2, ultima))}{r2}"

            # Filtros guardados (criterios)
            filtros = []
            m_af = re.search(r"<autoFilter\b[^>]*?/>|<autoFilter\b[^>]*>.*?</autoFilter>", txml, re.S)
            if m_af and quitar_filtros:
                for fc in re.finditer(r"<filterColumn\b([^>]*?)(?:/>|>(.*?)</filterColumn>)",
                                      m_af.group(0), re.S):
                    interior = fc.group(2) or ""
                    if not any(k in interior for k in CRITERIOS):
                        continue
                    col_id = int(attr("<x " + fc.group(1) + ">", "colId"))
                    nombre_col = actuales[col_id] if col_id < len(actuales) else f"col {col_id}"
                    texto, vals = describir_criterio(nombre_col, interior)
                    filtros.append({"xml": fc.group(0), "columna": nombre_col,
                                    "col": c1 + col_id, "criterio": texto, "vals": vals})

            ocultas = sorted(r for r in ocultas_hoja if r1 < r <= r2)
            a_mostrar = ocultas if filtros else []
            if ocultas and not filtros and quitar_filtros:
                avisos.append(f"{ws.title} / {t.name}: {len(ocultas)} filas ocultas sin filtro "
                              "activo (ocultas a mano); no se tocan.")
            # Coherencia: para filtros de igualdad, las ocultas deberían no cumplir el criterio
            for f in filtros:
                if f["vals"] is None:
                    f["coherencia"] = "n/d (criterio personalizado)"
                    continue
                vals = {norm(v) for v in f["vals"] if v != "(vacías)"}
                cumplen = [r for r in ocultas
                           if norm(ws.cell(row=r, column=f["col"]).value) in vals]
                f["coherencia"] = ("todas las ocultas no cumplen el criterio" if not cumplen
                                   else f"{len(cumplen)} ocultas SÍ cumplen el criterio")

            if ref_nueva != t.ref or filtros or a_mostrar:
                planes.append({
                    "hoja": ws.title, "tabla": t.name, "parte_tabla": parte_t,
                    "parte_hoja": partes[ws.title]["parte"],
                    "ref": t.ref, "ref_nueva": ref_nueva, "nuevas": nuevas,
                    "columnas_finales": actuales + [n for _, n in nuevas],
                    "filtros": filtros, "a_mostrar": a_mostrar,
                    "boton_filtro": m_af is not None,
                })
    return wb, planes, avisos


# --------------------------------------------------------------------------
# 2. Transformaciones XML
# --------------------------------------------------------------------------
def transformar_tabla(txml, plan):
    nuevo = re.sub(r'(<table\b[^>]*?\bref=")[^"]+(")',
                   lambda m: m.group(1) + plan["ref_nueva"] + m.group(2), txml, count=1)

    def af_sub(m):
        bloque = m.group(0)
        for f in plan["filtros"]:
            bloque = bloque.replace(f["xml"], "", 1)
        bloque = re.sub(r'(<autoFilter\b[^>]*?\bref=")[^"]+(")',
                        lambda x: x.group(1) + plan["ref_nueva"] + x.group(2), bloque, count=1)
        vacio = re.match(r"(<autoFilter\b[^>]*?)\s*>\s*</autoFilter>$", bloque, re.S)
        return vacio.group(1) + " />" if vacio else bloque

    nuevo = re.sub(r"<autoFilter\b[^>]*?/>|<autoFilter\b[^>]*>.*?</autoFilter>",
                   af_sub, nuevo, count=1, flags=re.S)

    if plan["nuevas"]:
        ids = [int(i) for i in re.findall(r'<tableColumn\b[^>]*?\bid="(\d+)"', nuevo)]
        prox = max(ids) + 1
        extra = "".join(f'<tableColumn id="{prox + i}" name="{html.escape(n, quote=True)}" />'
                        for i, (_, n) in enumerate(plan["nuevas"]))
        total = len(plan["columnas_finales"])
        nuevo = re.sub(r'(<tableColumns\b[^>]*?\bcount=")\d+(")',
                       lambda m: m.group(1) + str(total) + m.group(2), nuevo, count=1)
        nuevo = nuevo.replace("</tableColumns>", extra + "</tableColumns>", 1)
    for f in plan["filtros"]:
        if f["xml"] in nuevo:
            raise Abortar(f"No se pudo quitar el filtro de {plan['tabla']}.")
    return nuevo


def mostrar_filas(xml, filas):
    filas = set(filas)
    hechas = []

    def row_sub(m):
        tag = m.group(0)
        r = int(attr(tag, "r"))
        if r in filas and attr(tag, "hidden") in ("1", "true"):
            hechas.append(r)
            return re.sub(r'\s+hidden="(?:1|true)"', "", tag, count=1)
        return tag

    nuevo = re.sub(r"<row\b[^>]*>", row_sub, xml)
    if sorted(hechas) != sorted(filas):
        raise Abortar(f"Se esperaban {len(filas)} filas para mostrar y se encontraron {len(hechas)}.")
    return nuevo


def celda_xml(ref, valor):
    if valor is None:
        return ""
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}" t="n"><v>{valor}</v></c>'
    s = html.escape(str(valor), quote=False)
    esp = ' xml:space="preserve"' if s != s.strip() else ""
    return f'<c r="{ref}" t="inlineStr"><is><t{esp}>{s}</t></is></c>'


def agregar_registro(xml, filas_reg):
    if "</sheetData>" not in xml:
        raise Abortar("La hoja RegistroCambios no tiene un bloque sheetData editable.")
    ultima = max(int(x) for x in re.findall(r'<row\b[^>]*?\br="(\d+)"', xml))
    nuevas = []
    for i, fila in enumerate(filas_reg, start=ultima + 1):
        celdas = "".join(celda_xml(f"{get_column_letter(j)}{i}",
                                   registro._valor_para_celda(fila[h]))
                         for j, h in enumerate(registro.REGISTRO_HEADERS, start=1))
        nuevas.append(f'<row r="{i}">{celdas}</row>')
    xml = xml.replace("</sheetData>", "".join(nuevas) + "</sheetData>", 1)
    hasta = ultima + len(filas_reg)
    xml = re.sub(r'(<dimension\b[^>]*?\bref="A1:[A-Z]+)\d+(")',
                 lambda m: f"{m.group(1)}{hasta}{m.group(2)}", xml, count=1)
    return xml, ultima + 1, hasta


def construir(archivo, salida, z, planes, filas_reg, partes):
    modificadas = {}
    for p in planes:
        modificadas[p["parte_tabla"]] = transformar_tabla(
            z.read(p["parte_tabla"]).decode("utf-8"), p).encode("utf-8")
        if p["a_mostrar"]:
            base = modificadas.get(p["parte_hoja"], z.read(p["parte_hoja"])).decode("utf-8")
            modificadas[p["parte_hoja"]] = mostrar_filas(base, p["a_mostrar"]).encode("utf-8")
    p_reg = partes[registro.HOJA_REGISTRO]["parte"]
    xml_reg, desde, hasta = agregar_registro(z.read(p_reg).decode("utf-8"), filas_reg)
    modificadas[p_reg] = xml_reg.encode("utf-8")

    with zipfile.ZipFile(salida, "w") as zout:
        for item in z.infolist():
            datos = modificadas.get(item.filename, z.read(item.filename))
            nuevo = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            nuevo.compress_type, nuevo.external_attr = item.compress_type, item.external_attr
            zout.writestr(nuevo, datos)
    return modificadas, desde, hasta


# --------------------------------------------------------------------------
# 3. Verificaciones
# --------------------------------------------------------------------------
def verificar_bytes(z, salida, modificadas, partes):
    parte_marcas = partes[HOJA_MARCAS]["parte"]
    with zipfile.ZipFile(salida) as zo:
        if sorted(zo.namelist()) != sorted(z.namelist()):
            return False, "el conjunto de partes del .xlsx cambió"
        marcas_ok = sha(zo.read(parte_marcas)) == sha(z.read(parte_marcas))
        alteradas = [n for n in z.namelist()
                     if n not in modificadas and zo.read(n) != z.read(n)]
    ok = marcas_ok and not alteradas
    return ok, (f"Marcas idéntica byte a byte: {'sí' if marcas_ok else 'NO'} | "
                f"partes sin tocar alteradas: {alteradas or 'ninguna'} | "
                f"partes modificadas: {len(modificadas)}")


def verificar_openpyxl(wb_orig, salida, planes, filas_reg, desde, hojas=None,
                       exigir_sin_filtros=True):
    """Controla el resultado con openpyxl.

    En todas las hojas: cada tabla tiene la ref esperada (la nueva si estaba
    en el plan, la de antes si no) y sus columnas coinciden con los
    encabezados. En las hojas del alcance (todas, o las de `hojas`): ningún
    encabezado de la fila 1 queda fuera de la tabla y, si
    exigir_sin_filtros, no quedan criterios de filtro."""
    errores = []
    wb = openpyxl.load_workbook(salida)
    por_tabla = {p["tabla"]: p for p in planes}
    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS:
            continue
        en_alcance = hojas is None or ws.title in hojas
        ultimo_enc = max((c.column for c in ws[1]
                          if c.value is not None and str(c.value).strip() != ""), default=0)
        for t in ws.tables.values():
            p = por_tabla.get(t.name)
            esperado = p["ref_nueva"] if p else wb_orig[ws.title].tables[t.name].ref
            if t.ref != esperado:
                errores.append(f"{t.name}: ref {t.ref} (esperada {esperado})")
            c1, r1, c2, r2 = range_boundaries(t.ref)
            nombres = [tc.name for tc in t.tableColumns]
            celdas = [ws.cell(row=r1, column=c).value for c in range(c1, c2 + 1)]
            if nombres != celdas:
                errores.append(f"{t.name}: columnas {nombres} != encabezados {celdas}")
            if not en_alcance:
                continue
            if r1 == 1 and ultimo_enc > c2:
                fuera = [ws.cell(row=1, column=c).value for c in range(c2 + 1, ultimo_enc + 1)]
                errores.append(f"{t.name}: encabezados fuera de la tabla: {fuera}")
            if exigir_sin_filtros and t.autoFilter is not None and t.autoFilter.filterColumn:
                errores.append(f"{t.name}: todavía tiene criterios de filtro")
            if p and p["a_mostrar"]:
                ocultas = [r for r in p["a_mostrar"] if ws.row_dimensions[r].hidden]
                if ocultas:
                    errores.append(f"{t.name}: siguen ocultas {len(ocultas)} filas")
    # Valores de todas las hojas (salvo RegistroCambios) sin cambios
    for ws in wb.worksheets:
        if ws.title == registro.HOJA_REGISTRO:
            continue
        a = [tuple(map(valor_comparable, f)) for f in wb_orig[ws.title].iter_rows(values_only=True)]
        b = [tuple(map(valor_comparable, f)) for f in ws.iter_rows(values_only=True)]
        if a != b:
            errores.append(f"{ws.title}: cambiaron valores de celdas")
    # RegistroCambios
    reg = wb[registro.HOJA_REGISTRO]
    if reg.max_row != desde - 1 + len(filas_reg):
        errores.append(f"RegistroCambios: {reg.max_row} filas (esperadas {desde - 1 + len(filas_reg)})")
    for i, fila in enumerate(filas_reg):
        leidas = [reg.cell(row=desde + i, column=j).value for j in range(1, 11)]
        esperadas = [registro._valor_para_celda(fila[h]) for h in registro.REGISTRO_HEADERS]
        if leidas != esperadas:
            errores.append(f"RegistroCambios fila {desde + i}: {leidas} != {esperadas}")
            break
    return not errores, errores


PS_EXCEL = r"""
param([string]$Ruta, [string]$PidFile)
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'
$x = New-Object -ComObject Excel.Application
$p = [uint32]0
[void][W.U]::GetWindowThreadProcessId([IntPtr]$x.Hwnd, [ref]$p)
Set-Content -Path $PidFile -Value $p
$x.Visible = $false
$x.DisplayAlerts = $true
$x.AskToUpdateLinks = $false
$res = @{ ok = $false }
try {
  $wb = $x.Workbooks.Open($Ruta, 0, $true)
  $tablas = @()
  foreach ($ws in $wb.Worksheets) {
    foreach ($lo in $ws.ListObjects) {
      $cols = @(); foreach ($c in $lo.ListColumns) { $cols += $c.Name }
      $enc = @()
      $fila = $lo.HeaderRowRange.Row
      $ultima = $ws.Cells($fila, $ws.Columns.Count).End(-4159).Column
      $desde = $lo.HeaderRowRange.Column
      if ($ultima -lt ($desde + $lo.ListColumns.Count - 1)) { $ultima = $desde + $lo.ListColumns.Count - 1 }
      foreach ($cell in $ws.Range($ws.Cells($fila, $desde), $ws.Cells($fila, $ultima)).Cells) {
        $enc += @{ c = $cell.Address($false, $false); fondo = $cell.DisplayFormat.Interior.Color;
                   letra = $cell.DisplayFormat.Font.Color; negrita = $cell.DisplayFormat.Font.Bold }
      }
      $filtro = $false
      try { if ($lo.AutoFilter -ne $null) { $filtro = [bool]$lo.AutoFilter.FilterMode } } catch {}
      $filas = 0; $visibles = 0
      if ($lo.DataBodyRange -ne $null) {
        $filas = $lo.DataBodyRange.Rows.Count
        try { $visibles = $lo.DataBodyRange.Columns(1).SpecialCells(12).Count } catch { $visibles = 0 }
      }
      $tablas += @{ hoja = $ws.Name; tabla = $lo.Name; ref = $lo.Range.Address($false, $false);
                    columnas = $cols; encabezados = $enc; filtroActivo = $filtro;
                    filas = $filas; visibles = $visibles; boton = [bool]$lo.ShowAutoFilter }
    }
  }
  $wb.Close($false)
  $res = @{ ok = $true; tablas = $tablas }
} catch {
  $res = @{ ok = $false; error = "$_" }
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
$res | ConvertTo-Json -Depth 6 -Compress
"""


def abrir_en_excel(ruta, carpeta):
    """Abre el libro en Excel (solo lectura, oculto). Si Excel pide reparar,
    el diálogo bloquea la apertura y se corta por timeout."""
    ps = Path(carpeta) / "verificar_excel.ps1"
    ps.write_text(PS_EXCEL, encoding="utf-8-sig")
    pid_file = Path(carpeta) / "excel.pid"
    if pid_file.exists():
        pid_file.unlink()
    try:
        r = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps),
             "-Ruta", str(ruta), "-PidFile", str(pid_file)],
            capture_output=True, timeout=TIMEOUT_EXCEL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        if pid_file.exists():
            pid = pid_file.read_text().strip()
            subprocess.run(["taskkill", "/PID", pid, "/F"], capture_output=True)
        return {"ok": False, "error": f"Excel no terminó de abrir en {TIMEOUT_EXCEL} s "
                                      "(posible pedido de reparación)."}
    salida = r.stdout.decode("utf-8", errors="replace").strip()
    try:
        return json.loads(salida.splitlines()[-1])
    except (ValueError, IndexError):
        return {"ok": False, "error": "Respuesta inesperada de Excel: "
                + (salida or r.stderr.decode("utf-8", errors="replace"))[:500]}


def color_hex(bgr):
    """Excel devuelve los colores como entero BGR; se muestran como #RRGGBB."""
    try:
        v = int(bgr)
    except (TypeError, ValueError):
        return str(bgr)
    return f"#{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


def evaluar_excel(res, planes, antes=False, exigir_sin_filtros=True):
    """Devuelve (ok, líneas de detalle).

    Para cada tabla ampliada compara el formato visible (fondo, letra y
    negrita) de los encabezados nuevos con el de la columna anterior, y marca
    como error cualquier encabezado con letra del mismo color que el fondo.
    exigir_sin_filtros=False admite filtros activos (cuando no se pidió
    quitarlos)."""
    if not res.get("ok"):
        return False, [f"Excel: {res.get('error')}"]
    por_tabla = {p["tabla"]: p for p in planes}
    lineas, errores = [], []
    tablas = res["tablas"] if isinstance(res["tablas"], list) else [res["tablas"]]
    for t in tablas:
        enc = t["encabezados"] if isinstance(t["encabezados"], list) else [t["encabezados"]]
        cols = t["columnas"] if isinstance(t["columnas"], list) else [t["columnas"]]
        por_letra = {re.match(r"[A-Z]+", e["c"]).group(0): e for e in enc}
        invisibles = [e["c"] for e in enc if e["fondo"] == e["letra"]]
        p = por_tabla.get(t["tabla"])
        if p and p["nuevas"]:
            letra_ref = get_column_letter(p["nuevas"][0][0] - 1)
            mostrar = [letra_ref] + [get_column_letter(c) for c, _ in p["nuevas"]]
        else:
            letra_ref = None
            mostrar = list(por_letra)[-3:]
        distintos = []
        if letra_ref in por_letra:
            r = por_letra[letra_ref]
            clave = (r["fondo"], r["letra"], r["negrita"])
            distintos = [l for l in mostrar[1:] if l in por_letra and
                         (por_letra[l]["fondo"], por_letra[l]["letra"], por_letra[l]["negrita"]) != clave]
        linea = (f"  {t['hoja']:27} {t['tabla']:9} ref={t['ref']:9} filas={t['filas']:3} "
                 f"visibles={t['visibles']:3} filtroActivo={'sí' if t['filtroActivo'] else 'no'} "
                 f"| últimas columnas: {', '.join(cols[-3:])}")
        if invisibles:
            linea += f" | ENCABEZADO INVISIBLE: {', '.join(invisibles)}"
        lineas.append(linea)
        detalle = [f"{por_letra[l]['c']} fondo {color_hex(por_letra[l]['fondo'])} "
                   f"letra {color_hex(por_letra[l]['letra'])}{' negrita' if por_letra[l]['negrita'] else ''}"
                   for l in mostrar if l in por_letra]
        lineas.append(f"  {'':27} encabezados: " + " | ".join(detalle))
        if antes:
            continue
        if p and t["ref"] != p["ref_nueva"]:
            errores.append(f"{t['tabla']}: Excel ve ref {t['ref']} (esperada {p['ref_nueva']})")
        if p and cols != p["columnas_finales"]:
            errores.append(f"{t['tabla']}: Excel ve columnas {cols}")
        if exigir_sin_filtros and (t["filtroActivo"] or t["visibles"] != t["filas"]):
            errores.append(f"{t['tabla']}: filtro activo o filas ocultas en Excel")
        if invisibles:
            errores.append(f"{t['tabla']}: encabezados invisibles {invisibles}")
        if distintos:
            errores.append(f"{t['tabla']}: {distintos} no se ven igual que {letra_ref}1")
    return not errores, lineas + [f"  ERROR: {e}" for e in errores]
