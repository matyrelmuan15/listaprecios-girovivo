#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ordenar_validacion_nombres.py
------------------------------
Dos tareas de orden estructural (no tocan precios ni valores de celdas):

1) Empareja la validación de datos (lista desplegable) de la columna
   ActualizadoSegun en cada hoja de destino a un único rango continuo
   <col>2:<col>5000 (criterio fijo, para que las filas nuevas que se agreguen
   desde Excel tengan la lista). Si alguna tabla supera la fila 5000, aborta
   en la inspección. La columna se busca por nombre de encabezado. Solo se
   reescribe el atributo sqref de esa validación; la lista no cambia. Las
   hojas que ya tienen el rango correcto no se tocan ni se registran.

2) Borra de xl/workbook.xml los nombres definidos "SegmentaciónDeDatos_Marca…"
   que apuntan a #N/A (restos de segmentaciones perdidas). Antes confirma, en
   solo lectura, que ninguna otra parte del libro los mencione (fórmulas,
   tablas, validaciones, otros nombres) y que no queden slicers,
   slicerCaches ni relaciones de segmentación. Si encuentra algo, aborta.
   Ningún otro nombre definido se toca.

Se edita SOLO el XML (hojas de destino, workbook.xml y RegistroCambios);
el resto, incluida "Marcas Actualización Dinámica", se copia byte a byte.
Reutiliza tablas_xml.py (construir, verificaciones y apertura en Excel).

Registro (TipoOperacion "Estructura"):
  - una fila por hoja con validación ajustada: Columna "Validación
    <encabezado>", rango anterior -> nuevo;
  - una fila por nombre borrado: Hoja "(Libro)", Columna "Nombre definido
    <nombre>", contenido anterior -> "(borrado)".

USO (desde scripts/)
    python ordenar_validacion_nombres.py            -> inspección + dry-run
    python ordenar_validacion_nombres.py --execute  -> + backup + reemplazo real
    --archivo RUTA                                  -> procesa otro libro
"""

import argparse
import html
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

import registro
from ampliar_tablas import ARCHIVO_PATH, backup_archivo
from tablas_xml import (HOJA_MARCAS, HOJAS_EXCLUIDAS, Abortar, abrir_en_excel, attr,
                        construir, mapa_partes, sha, verificar_bytes, verificar_en_excel,
                        verificar_openpyxl)

ENCABEZADO = "ActualizadoSegun"
# Criterio fijo: la lista cubre de la fila 2 a la 5000 en todas las hojas, en un
# solo rango continuo, para que las filas nuevas agregadas desde Excel la tengan.
FILA_FINAL = 5000
VALORES_LISTA = {"Manual", "Dólar", "Lista proveedor"}
PREFIJO_NOMBRE = "SegmentaciónDeDatos_Marca"
RE_NOMBRE = re.compile(r"^SegmentaciónDeDatos_Marca\d*$")
# Rastros de segmentaciones que no deberían existir en ninguna parte.
RASTROS = ("slicer", "timeline", "x14:slicerList", "x15:slicerCaches")


def sqref_celdas(sqref):
    """Conjunto de (col, fila) que cubre un sqref (puede tener varios rangos)."""
    celdas = set()
    for rango in sqref.split():
        c1, r1, c2, r2 = range_boundaries(rango)
        for c in range(c1, c2 + 1):
            for r in range(r1, r2 + 1):
                celdas.add((c, r))
    return celdas


# --------------------------------------------------------------------------
# Tarea 1: validaciones
# --------------------------------------------------------------------------
def planificar_validaciones(wb, z, partes):
    planes = []
    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS or not ws.tables:
            continue
        if len(ws.tables) != 1:
            raise Abortar(f"{ws.title}: tiene {len(ws.tables)} tablas; se esperaba una.")
        t = list(ws.tables.values())[0]
        c1, r1, c2, r2 = range_boundaries(t.ref)
        col = next((c.column for c in ws[r1] if c.value is not None
                    and str(c.value).strip() == ENCABEZADO), None)
        if col is None or not (c1 <= col <= c2):
            raise Abortar(f"{ws.title}: no se encontró '{ENCABEZADO}' dentro de la tabla {t.name}.")
        letra = get_column_letter(col)
        if r2 > FILA_FINAL:
            raise Abortar(f"{ws.title}: la tabla {t.name} llega a la fila {r2}, más allá de "
                          f"{FILA_FINAL}; hay que ampliar el criterio antes de seguir.")
        nuevo = f"{letra}{r1 + 1}:{letra}{FILA_FINAL}"

        parte = partes[ws.title]["parte"]
        xml = z.read(parte).decode("utf-8")
        if "x14:dataValidation" in xml:
            raise Abortar(f"{ws.title}: tiene validaciones x14 (extLst); no previstas.")
        tags = re.findall(r"<dataValidation\b[^>]*>.*?</dataValidation>", xml, re.S)
        candidatas = []
        for bloque in tags:
            tag = re.match(r"<dataValidation\b[^>]*>", bloque).group(0)
            celdas = sqref_celdas(attr(tag, "sqref"))
            if any(c == col for c, _ in celdas):
                candidatas.append((bloque, tag, celdas))
        if len(candidatas) != 1:
            raise Abortar(f"{ws.title}: {len(candidatas)} validaciones tocan la columna {letra}; "
                          "se esperaba exactamente una.")
        bloque, tag, celdas = candidatas[0]
        if any(c != col for c, _ in celdas):
            raise Abortar(f"{ws.title}: la validación de {letra} también cubre otras columnas.")
        f1 = re.search(r"<formula1>(.*?)</formula1>", bloque, re.S)
        valores = set(html.unescape(f1.group(1)).strip('"').split(",")) if f1 else set()
        if attr(tag, "type") != "list" or valores != VALORES_LISTA:
            raise Abortar(f"{ws.title}: la validación de {letra} no es la lista esperada: {valores}")
        # Valores cargados en la columna más allá de la fila final quedarían sin validar.
        fuera = [r for r in range(FILA_FINAL + 1, ws.max_row + 1)
                 if ws.cell(row=r, column=col).value not in (None, "")]
        if fuera:
            raise Abortar(f"{ws.title}: hay valores en {letra} después de la fila {FILA_FINAL} "
                          f"(filas {fuera[:10]}).")
        planes.append({
            "hoja": ws.title, "tabla": t.name, "parte": parte, "letra": letra,
            "sqref": attr(tag, "sqref"), "sqref_nuevo": nuevo, "tag": tag,
            "filas_datos": r2 - r1, "ultima_fila": r2,
            "cambia": attr(tag, "sqref") != nuevo,
        })
    return planes


def aplicar_validacion(xml, plan):
    tag_nuevo = re.sub(r'(\bsqref=")[^"]*(")',
                       lambda m: m.group(1) + plan["sqref_nuevo"] + m.group(2), plan["tag"], count=1)
    if xml.count(plan["tag"]) != 1:
        raise Abortar(f"{plan['hoja']}: la etiqueta de validación no es única en el XML.")
    return xml.replace(plan["tag"], tag_nuevo, 1)


# --------------------------------------------------------------------------
# Tarea 2: nombres definidos
# --------------------------------------------------------------------------
def planificar_nombres(z):
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    nombres = []
    for m in re.finditer(r"<definedName\b([^>]*)>(.*?)</definedName>", wbx, re.S):
        nombre = attr("<x " + m.group(1) + ">", "name")
        nombres.append({"xml": m.group(0), "nombre": nombre,
                        "contenido": html.unescape(m.group(2))})
    a_borrar = [n for n in nombres if RE_NOMBRE.match(n["nombre"] or "")]
    if len(a_borrar) != 8:
        raise Abortar(f"se esperaban 8 nombres '{PREFIJO_NOMBRE}…' y hay {len(a_borrar)}.")
    for n in a_borrar:
        if n["contenido"].strip() != "#N/A":
            raise Abortar(f"el nombre {n['nombre']} no apunta a #N/A: {n['contenido']}")
    otros = [n for n in nombres if n not in a_borrar]

    # Control de referencias (solo lectura): fuera de esos 8 elementos, nadie
    # puede mencionar los nombres ni haber rastros de segmentaciones.
    hallazgos = []
    claves = [PREFIJO_NOMBRE, html.escape(PREFIJO_NOMBRE),
              PREFIJO_NOMBRE.replace("ó", "&#243;"), PREFIJO_NOMBRE.replace("ó", "&#xF3;"),
              "SegmentaciónDeDatos", "Segmentaci&#243;nDeDatos"]
    for parte in z.namelist():
        texto = z.read(parte).decode("utf-8", errors="replace")
        if parte == "xl/workbook.xml":
            for n in a_borrar:
                texto = texto.replace(n["xml"], "", 1)
        for k in claves:
            if k.casefold() in texto.casefold():
                hallazgos.append(f"{parte}: menciona '{k}'")
        for k in RASTROS:
            if k.casefold() in texto.casefold() or k.casefold() in parte.casefold():
                hallazgos.append(f"{parte}: contiene '{k}'")
    return wbx, a_borrar, otros, sorted(set(hallazgos))


def aplicar_nombres(wbx, a_borrar):
    for n in a_borrar:
        if wbx.count(n["xml"]) != 1:
            raise Abortar(f"el nombre {n['nombre']} no aparece una sola vez en workbook.xml.")
        wbx = wbx.replace(n["xml"], "", 1)
    if re.search(r"<definedNames\b[^>]*>\s*</definedNames>", wbx):
        wbx = re.sub(r"<definedNames\b[^>]*>\s*</definedNames>", "", wbx)
    return wbx


# --------------------------------------------------------------------------
# Verificaciones propias
# --------------------------------------------------------------------------
def verificar_propias(salida, planes_v, a_borrar, otros):
    errores = []
    wb = openpyxl.load_workbook(salida)
    for p in planes_v:
        dvs = wb[p["hoja"]].data_validations.dataValidation
        vistos = [str(dv.sqref) for dv in dvs]
        if vistos != [p["sqref_nuevo"]]:
            errores.append(f"{p['hoja']}: validaciones {vistos} (esperada {p['sqref_nuevo']})")
    nombres = set(wb.defined_names.keys())
    for ws in wb.worksheets:
        nombres |= {f"{ws.title}!{k}" for k in ws.defined_names.keys()}
    quedan = [n["nombre"] for n in a_borrar if n["nombre"] in nombres]
    if quedan:
        errores.append(f"siguen los nombres {quedan}")
    with zipfile.ZipFile(salida) as zo:
        wbx = zo.read("xl/workbook.xml").decode("utf-8")
    for n in otros:
        if n["xml"] not in wbx:
            errores.append(f"cambió o desapareció el nombre {n['nombre']}")
    return not errores, errores


def verificar_excel_extra(res, planes_v, a_borrar):
    if not res.get("ok"):
        return False, [f"Excel: {res.get('error')}"]
    errores, lineas = [], []
    val = res.get("validaciones") or {}
    for p in planes_v:
        visto = val.get(p["hoja"])
        lineas.append(f"  {p['hoja']:27} validación en Excel: {visto}")
        if visto != p["sqref_nuevo"]:
            errores.append(f"{p['hoja']}: Excel ve la validación en {visto} (esperada {p['sqref_nuevo']})")
    nombres = res.get("nombres") or []
    nombres = nombres if isinstance(nombres, list) else [nombres]
    lineas.append(f"  Nombres en Excel: {nombres}")
    quedan = [n["nombre"] for n in a_borrar if n["nombre"] in nombres]
    if quedan:
        errores.append(f"Excel todavía ve {quedan}")
    return not errores, lineas + [f"  ERROR: {e}" for e in errores]


# --------------------------------------------------------------------------
# Proceso principal
# --------------------------------------------------------------------------
def procesar(execute, archivo=ARCHIVO_PATH):
    if not archivo.exists():
        raise Abortar(f"no se encontró el archivo de trabajo en: {archivo}")
    sha_origen = sha(archivo.read_bytes())
    z = zipfile.ZipFile(archivo)
    partes = mapa_partes(z)
    for req in (HOJA_MARCAS, registro.HOJA_REGISTRO):
        if req not in partes:
            raise Abortar(f"falta la hoja '{req}' en el libro.")

    modo = "EJECUCIÓN" if execute else "DRY-RUN (no se modifica el archivo real)"
    print(f"\n=== Validación ActualizadoSegun y nombres de segmentación | Modo: {modo} ===\n")

    # ---------------- Fase 1: inspección ----------------
    wb_orig = openpyxl.load_workbook(archivo)  # sin data_only=True; nunca se guarda
    registro.validar_encabezados(wb_orig)
    planes_v = planificar_validaciones(wb_orig, z, partes)
    wbx, a_borrar, otros, hallazgos = planificar_nombres(z)

    print("--- Fase 1: Inspección ---")
    print(f"Tarea 1: validación de '{ENCABEZADO}' en {len(planes_v)} hojas")
    print(f"  {'Hoja':27} {'Tabla':9} {'Filas datos':>11} {'Última fila':>11}  "
          f"{'Rango actual':22} -> Rango propuesto")
    for p in planes_v:
        print(f"  {p['hoja']:27} {p['tabla']:9} {p['filas_datos']:>11} {p['ultima_fila']:>11}  "
              f"{p['sqref']:22} -> "
              f"{p['sqref_nuevo']}{'' if p['cambia'] else '  (sin cambio)'}")
    print(f"\nTarea 2: nombres definidos a borrar ({len(a_borrar)})")
    for n in a_borrar:
        print(f"  - {n['nombre']:30} = {n['contenido']}")
    print(f"  Nombres que se conservan sin tocar: "
          + ", ".join(f"{n['nombre']} = {n['contenido']}" for n in otros))
    print("  Control de referencias (fórmulas, tablas, validaciones, otros nombres, "
          "slicers, slicerCaches, relaciones, tipos de contenido):")
    if hallazgos:
        for h in hallazgos:
            print(f"    HALLAZGO: {h}")
        raise Abortar("hay referencias a las segmentaciones; no se borra nada. Revisalo a mano.")
    print(f"    ninguna referencia en las {len(z.namelist())} partes del archivo.")

    # ---------------- Fase 2: registro + libro simulado ----------------
    fecha = registro.timestamp_corrida()
    cambios_v = [p for p in planes_v if p["cambia"]]
    filas_reg = []
    for p in cambios_v:
        filas_reg.append(registro.nueva_fila(
            fecha, registro.TIPO_ESTRUCTURA, p["hoja"], 1, None, None,
            f"Validación {ENCABEZADO}", p["sqref"], p["sqref_nuevo"],
            f"Lista desplegable emparejada a {p['letra']}2:{p['letra']}{FILA_FINAL} "
            f"(criterio único; {p['tabla']} llega hoy a la fila {p['ultima_fila']})"))
    for n in a_borrar:
        filas_reg.append(registro.nueva_fila(
            fecha, registro.TIPO_ESTRUCTURA, "(Libro)", None, None, None,
            f"Nombre definido {n['nombre']}", n["contenido"], "(borrado)",
            "Nombre de segmentación perdida, sin referencias en el libro"))
    registro.verificar_conteo(filas_reg, len(cambios_v) + len(a_borrar))

    print("\n--- Fase 2: Resumen ---")
    print(f"  Validaciones ajustadas: {len(cambios_v)} | Nombres borrados: {len(a_borrar)}")
    print(f"  Filas de RegistroCambios: {len(filas_reg)} "
          f"(= {len(cambios_v)} validaciones + {len(a_borrar)} nombres)")
    print(f"\n--- Vista previa de '{registro.HOJA_REGISTRO}' ---")
    for fila in filas_reg:
        print("    - " + " | ".join("(vacío)" if fila[h] is None else str(fila[h])
                                    for h in registro.REGISTRO_HEADERS))
    if not filas_reg:
        print("\nNo hay nada para cambiar.")
        return 0

    extra = {}
    for p in cambios_v:
        extra[p["parte"]] = aplicar_validacion(z.read(p["parte"]).decode("utf-8"), p).encode("utf-8")
    extra["xl/workbook.xml"] = aplicar_nombres(wbx, a_borrar).encode("utf-8")

    carpeta = Path(tempfile.mkdtemp(prefix="ordenar_validacion_"))
    salida = carpeta / archivo.name
    modificadas, desde, hasta = construir(archivo, salida, z, [], filas_reg, partes, extra=extra)
    print(f"\nLibro simulado: {salida}")
    print(f"Partes modificadas: {', '.join(sorted(modificadas))}")

    # ---------------- Verificaciones ----------------
    print("\n--- Verificación 1: openpyxl (tablas, valores, registro, validaciones, nombres) ---")
    ok1a, err1a = verificar_openpyxl(wb_orig, salida, [], filas_reg, desde)
    ok1b, err1b = verificar_propias(salida, planes_v, a_borrar, otros)
    ok1 = ok1a and ok1b
    print("  OK" if ok1 else "\n".join(f"  ERROR: {e}" for e in err1a + err1b))
    print(f"  RegistroCambios: filas {desde} a {hasta}")

    print("\n--- Verificación 2: byte a byte ---")
    ok2, det2 = verificar_bytes(z, salida, modificadas, partes)
    print(f"  {'OK' if ok2 else 'ERROR'} | {det2}")

    print("\n--- Verificación 3: Excel ---")
    requeridas = [p["tabla"] for p in planes_v]
    ok3a, res, lin, intentos = verificar_en_excel(salida, carpeta, [], requeridas=requeridas)
    print(("  Excel abrió el libro simulado sin pedido de reparación."
           if res.get("ok") else "  Excel NO pudo abrir el libro simulado.")
          + (f" (intentos: {intentos})" if intentos > 1 else ""))
    ok3b, lin_b = verificar_excel_extra(res, planes_v, a_borrar)
    print("\n".join(lin_b))
    errs = [l for l in lin if l.startswith("  ERROR")]
    if errs:
        print("\n".join(errs))
    ok3 = ok3a and ok3b
    print(f"  {'OK' if ok3 else 'ERROR'}")

    z.close()
    if not (ok1 and ok2 and ok3):
        raise Abortar("alguna verificación falló; no se toca el archivo real.")

    if not execute:
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo real ni se creó backup.")
        return len(filas_reg)

    # ---------------- Fase 3: ejecución ----------------
    if sha(archivo.read_bytes()) != sha_origen:
        raise Abortar("el archivo real cambió durante la corrida; volvé a correr el script.")
    backup = backup_archivo(archivo)
    print(f"\nBackup creado en: {backup}")
    try:
        shutil.copyfile(salida, archivo)
    except PermissionError:
        raise Abortar("no se pudo reemplazar el archivo porque está abierto en Excel. "
                      "Cerralo y volvé a correr el script. El archivo real no se modificó.")
    if sha(archivo.read_bytes()) != sha(salida.read_bytes()):
        raise Abortar(f"el archivo guardado no coincide con el verificado. Restaurá el backup: {backup}")
    print(f"Archivo reemplazado: {archivo}")
    print(f"Validaciones ajustadas: {len(cambios_v)} | Nombres borrados: {len(a_borrar)} | "
          f"Filas agregadas a {registro.HOJA_REGISTRO}: {len(filas_reg)}")
    return len(filas_reg)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Empareja la validación de ActualizadoSegun y "
                                             "borra los nombres de segmentaciones rotas.")
    ap.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    ap.add_argument("--archivo", type=Path, default=ARCHIVO_PATH,
                    help="Libro a procesar (por defecto, el archivo real).")
    args = ap.parse_args()
    try:
        procesar(args.execute, args.archivo)
    except (Abortar, registro.RegistroError) as e:
        sys.exit(f"\nABORTADO: {e}")


if __name__ == "__main__":
    main()
