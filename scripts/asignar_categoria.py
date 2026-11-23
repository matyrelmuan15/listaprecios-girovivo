#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
asignar_categoria.py
---------------------
Asigna un texto fijo en la columna "Categoria" de las hojas de destino cuyo
nombre contenga alguno de los filtros indicados (sin importar mayúsculas ni
acentos), y deja esa columna DENTRO de la tabla de Excel de la hoja.

    python asignar_categoria.py --filtro "bicicleta" --valor "Bicicletas"
        -> inspección + dry-run (arma y verifica el resultado en una carpeta
           temporal; no toca el archivo real)
    python asignar_categoria.py --filtro "bicicleta" --valor "Bicicletas" --execute
        -> lo mismo + backup + reemplazo del archivo real
    python asignar_categoria.py --hoja "Cuadros" --valor "Cuadros"
        -> solo la hoja con ese nombre exacto (sin coincidencias parciales)
    --archivo RUTA  -> procesa otro libro (por ejemplo, una copia de prueba)

Reglas:
  - Si la hoja no tiene "Categoria", se crea el encabezado en la primera
    columna libre a la derecha del último encabezado (nunca se insertan
    columnas en el medio), copiando el formato y el ancho del encabezado
    vecino.
  - Si la hoja ya tiene "Categoria", se usa esa columna.
  - Regla del proyecto: toda columna de una hoja de destino queda dentro de
    la tabla. Si "Categoria" queda fuera de la tabla de la hoja (porque se
    acaba de crear o porque ya estaba afuera), la tabla se amplía. Si ya está
    dentro y todas las filas tienen el valor, no cambia nada.
  - Fila de producto = fila con CodigoDux no vacío.
  - Celdas que ya tienen el valor pedido: se saltean. Celdas con OTRA
    categoría: no se pisan salvo con --sobrescribir; se informan.
  - Valores fijos (texto), nunca fórmulas. Sin lista desplegable.
  - Los filtros guardados de las tablas no se tocan (eso es ampliar_tablas.py).

Cómo se arma el resultado (siempre en una carpeta temporal):
  1. openpyxl escribe encabezado y valores, y registra en RegistroCambios
     "Estructura" (encabezado, Columna = letra) y "Categoria" (cada valor,
     Columna = "Categoria"). Se guarda una copia intermedia.
  2. Sobre esa copia, tablas_xml amplía por XML solo las tablas de las hojas
     filtradas y registra una fila "Estructura" por tabla ampliada
     (Columna "Tabla <nombre>", ref anterior -> nueva).
  3. Verificaciones: openpyxl (refs, columnas, valores, registro y que las
     demás hojas no cambien respecto del original), byte a byte contra la
     copia intermedia (Marcas incluida) y apertura en Excel.
  El conteo verifica: filas de registro = celdas modificadas + tablas ampliadas.
"""

import argparse
import shutil
import sys
import tempfile
import unicodedata
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

import actualizado_segun as act
import registro
import tablas_xml as tx

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"

COL_CATEGORIA = "Categoria"
COL_CODIGO_DUX = "CodigoDux"
COL_NOMBRE = "NombreProducto"
COL_MARCA = "Marca"
ENCABEZADOS_DESTINO = (COL_CODIGO_DUX, "PrecioConIVA")

HOJAS_EXCLUIDAS = {
    "Hoja1", "Marcas Actualización Dinámica", registro.HOJA_REGISTRO,
    act.HOJA_NO_ACTUALIZADOS, *act.HOJAS_REEMPLAZADAS,
}


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def normalizar(texto):
    """Sin acentos, sin espacios sobrantes y en minúsculas."""
    if texto is None:
        return ""
    sin_acentos = unicodedata.normalize("NFKD", str(texto))
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return sin_acentos.strip().casefold()


def es_vacio(valor):
    return valor is None or str(valor).strip() == ""


def mapa_encabezados(ws):
    """{encabezado_normalizado: número de columna} de la fila 1."""
    mapa = {}
    for celda in ws[1]:
        if not es_vacio(celda.value):
            mapa.setdefault(normalizar(celda.value), celda.column)
    return mapa


def hoja_coincide(titulo, filtros, exactas=()):
    """--filtro: el fragmento aparece en el nombre. --hoja: el nombre es
    exactamente ese (sin importar mayúsculas ni acentos)."""
    t = normalizar(titulo)
    return any(f in t for f in filtros) or t in exactas


def backup_archivo(archivo):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo.stem}_{ts}{archivo.suffix}"
    shutil.copy2(archivo, destino)
    return destino


def tabla_de_hoja(ws):
    """(tabla, c1, r1, c2, r2) de la tabla que empieza en la fila 1, o None."""
    for t in ws.tables.values():
        c1, r1, c2, r2 = range_boundaries(t.ref)
        if r1 == 1:
            return t, c1, r1, c2, r2
    return None


# --------------------------------------------------------------------------
# Planificación (no escribe nada)
# --------------------------------------------------------------------------
def planificar_hoja(ws, valor, sobrescribir):
    enc = mapa_encabezados(ws)
    col = enc.get(normalizar(COL_CATEGORIA))
    crear = col is None
    if crear:
        col = act.ultima_columna_encabezado(ws) + 1

    c_dux = enc[normalizar(COL_CODIGO_DUX)]
    c_nombre = enc.get(normalizar(COL_NOMBRE))
    c_marca = enc.get(normalizar(COL_MARCA))

    a_escribir, ya_ok, distintas, ocupadas_sin_dux = [], 0, [], []
    for r in range(2, ws.max_row + 1):
        actual = ws.cell(row=r, column=col).value
        if es_vacio(ws.cell(row=r, column=c_dux).value):
            if not crear and not es_vacio(actual):
                ocupadas_sin_dux.append((r, actual))
            continue
        if not es_vacio(actual) and str(actual).strip() == valor:
            ya_ok += 1
            continue
        if not es_vacio(actual) and not sobrescribir:
            distintas.append((r, actual))
            continue
        a_escribir.append({
            "fila": r,
            "anterior": None if es_vacio(actual) else actual,
            "nombre": ws.cell(row=r, column=c_nombre).value if c_nombre else None,
            "marca": ws.cell(row=r, column=c_marca).value if c_marca else None,
        })

    # Tabla de la hoja: ¿Categoria queda adentro?
    info = tabla_de_hoja(ws)
    tabla = ref_nueva = None
    if info:
        t, c1, r1, c2, r2 = info
        tabla = t.name
        if col > c2:
            ref_nueva = f"{get_column_letter(c1)}{r1}:{get_column_letter(col)}{r2}"

    return {
        "ws": ws, "col": col, "crear": crear, "a_escribir": a_escribir,
        "ya_ok": ya_ok, "distintas": distintas,
        "ocupadas_sin_dux": ocupadas_sin_dux,
        "tabla": tabla, "ref": info[0].ref if info else None, "ref_nueva": ref_nueva,
    }


# --------------------------------------------------------------------------
# Escritura (siempre sobre el libro en memoria; se guarda en una copia)
# --------------------------------------------------------------------------
def crear_encabezado(ws, col):
    nueva = ws.cell(row=1, column=col)
    nueva.value = COL_CATEGORIA
    if col > 1:
        modelo = ws.cell(row=1, column=col - 1)
        if modelo.has_style:
            nueva._style = copy(modelo._style)
        ancho = ws.column_dimensions[get_column_letter(col - 1)].width
        ws.column_dimensions[get_column_letter(col)].width = max(ancho or 0, 18)


def escribir_valor(ws, fila, col, valor, copiar_estilo):
    celda = ws.cell(row=fila, column=col)
    celda.value = valor
    if copiar_estilo and col > 1:
        vecina = ws.cell(row=fila, column=col - 1)
        if vecina.has_style:
            celda._style = copy(vecina._style)


def comparar_con_original(wb_orig, salida, hojas_cambiadas):
    """Las hojas que no se procesan (Marcas incluida) deben quedar con los
    mismos valores y fórmulas que en el archivo original."""
    wb = openpyxl.load_workbook(salida)
    errores = []
    for ws in wb_orig.worksheets:
        if ws.title in hojas_cambiadas or ws.title == registro.HOJA_REGISTRO:
            continue
        a = [tuple(map(tx.valor_comparable, f)) for f in ws.iter_rows(values_only=True)]
        b = [tuple(map(tx.valor_comparable, f)) for f in wb[ws.title].iter_rows(values_only=True)]
        if a != b:
            errores.append(f"{ws.title}: cambiaron valores respecto del original")
    return not errores, errores


# --------------------------------------------------------------------------
# Proceso principal
# --------------------------------------------------------------------------
def procesar(filtros, valor, execute, sobrescribir, archivo=ARCHIVO_PATH, hojas=()):
    archivo = Path(archivo)
    if not archivo.exists():
        raise tx.Abortar(f"no se encontró el archivo de trabajo en: {archivo}")
    valor = valor.strip()
    if not valor:
        raise tx.Abortar("--valor no puede estar vacío.")
    if hojas:
        nombres = {normalizar(n) for n in openpyxl.load_workbook(archivo, read_only=True).sheetnames}
        faltan = [h for h in hojas if normalizar(h) not in nombres]
        if faltan:
            raise tx.Abortar(f"no existe ninguna hoja llamada {faltan}.")
    filtros = [normalizar(f) for f in (filtros or []) if normalizar(f)]
    exactas = {normalizar(h) for h in (hojas or []) if normalizar(h)}
    if not filtros and not exactas:
        raise tx.Abortar("indicá al menos un --filtro o una --hoja.")

    sha_origen = tx.sha(archivo.read_bytes())
    wb_orig = openpyxl.load_workbook(archivo)  # referencia para comparar; nunca se guarda
    wb = openpyxl.load_workbook(archivo)       # sin data_only=True; se guarda solo en una copia
    registro.validar_encabezados(wb)

    # ---------------- Fase 1: inspección ----------------
    planes, sin_esquema = [], []
    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS or not hoja_coincide(ws.title, filtros, exactas):
            continue
        enc = mapa_encabezados(ws)
        if not all(normalizar(e) in enc for e in ENCABEZADOS_DESTINO):
            sin_esquema.append(ws.title)
            continue
        planes.append(planificar_hoja(ws, valor, sobrescribir))

    modo = "EJECUCIÓN" if execute else "DRY-RUN (no se modifica el archivo real)"
    print(f"\n=== Asignar Categoria | Modo: {modo} ===")
    print(f"Archivo: {archivo}")
    print(f"Hojas: {', '.join([f'contiene {f!r}' for f in filtros] + [f'= {h!r}' for h in sorted(exactas)])}"
          f" | Valor: '{valor}'"
          f" | Sobrescribir otras categorías: {'sí' if sobrescribir else 'no'}\n")

    print("--- Fase 1: Inspección ---")
    if not planes:
        print("No se encontró ninguna hoja de destino que coincida con el filtro.")
    for p in planes:
        ws, letra = p["ws"], get_column_letter(p["col"])
        if p["crear"]:
            vecino = ws.cell(row=1, column=p["col"] - 1).value
            donde = f"columna {letra} (nueva, después de '{vecino}')"
        else:
            donde = f"columna {letra} (ya existe)"
        if p["tabla"] is None:
            tabla = "sin tabla de Excel"
        elif p["ref_nueva"]:
            tabla = f"tabla {p['tabla']} {p['ref']} -> {p['ref_nueva']}"
        else:
            tabla = f"tabla {p['tabla']} {p['ref']} (ya la incluye)"
        print(f"  - {ws.title}: {donde} | {tabla} | filas a marcar: {len(p['a_escribir'])}"
              f" | ya con '{valor}': {p['ya_ok']}"
              f" | con otra categoría (no se tocan): {len(p['distintas'])}")
        for r, v in p["distintas"][:5]:
            print(f"      fila {r}: '{v}'")
        if p["ocupadas_sin_dux"]:
            print(f"      Aviso: {len(p['ocupadas_sin_dux'])} celda(s) de Categoria con "
                  f"valor en filas sin CodigoDux (no se tocan).")
        if p["tabla"] is None:
            print("      Aviso: la hoja no tiene tabla de Excel; no hay nada que ampliar.")
    if sin_esquema:
        print(f"  Hojas que coinciden pero no son de destino (se ignoran): "
              f"{', '.join(sin_esquema)}")

    # ---------------- Fase 2: registro de celdas ----------------
    corrida_enc = registro.RegistroCorrida(
        registro.TIPO_ESTRUCTURA, f"Nueva columna {COL_CATEGORIA}")
    corrida_val = registro.RegistroCorrida(
        registro.TIPO_CATEGORIA, f"Asignación de categoría '{valor}'",
        fecha=corrida_enc.fecha)
    for p in planes:
        ws, letra = p["ws"], get_column_letter(p["col"])
        if p["crear"]:
            corrida_enc.agregar(hoja=ws.title, fila=1, nombre_producto=None,
                                marca=None, columna=letra, valor_anterior=None,
                                valor_nuevo=COL_CATEGORIA)
        for f in p["a_escribir"]:
            corrida_val.agregar(hoja=ws.title, fila=f["fila"],
                                nombre_producto=f["nombre"], marca=f["marca"],
                                columna=COL_CATEGORIA, valor_anterior=f["anterior"],
                                valor_nuevo=valor)

    a_ampliar = [p for p in planes if p["ref_nueva"]]
    n_enc, n_val, n_tab = len(corrida_enc), len(corrida_val), len(a_ampliar)
    total = n_enc + n_val + n_tab
    print("\n--- Fase 2: Resumen ---")
    print(f"  Hojas procesadas: {len(planes)}")
    print(f"  Encabezados nuevos (TipoOperacion '{registro.TIPO_ESTRUCTURA}'): {n_enc}")
    print(f"  Celdas de categoría (TipoOperacion '{registro.TIPO_CATEGORIA}'): {n_val}")
    print(f"  Tablas ampliadas (TipoOperacion '{registro.TIPO_ESTRUCTURA}'): {n_tab}")
    print(f"  Total de cambios = filas de RegistroCambios: {total}")

    if total == 0:
        print("\nNo hay nada para cambiar.")
        return 0

    # Etapa 1 (openpyxl): celdas + registro de celdas -> copia intermedia
    modificadas_celdas = 0
    for p in planes:
        ws, col = p["ws"], p["col"]
        if p["crear"]:
            crear_encabezado(ws, col)
            modificadas_celdas += 1
        for f in p["a_escribir"]:
            escribir_valor(ws, f["fila"], col, valor, copiar_estilo=p["crear"])
            modificadas_celdas += 1
    corrida_enc.filas.extend(corrida_val.filas)
    corrida_enc.verificar_conteo(modificadas_celdas)
    corrida_enc.escribir(wb)

    carpeta = Path(tempfile.mkdtemp(prefix="asignar_categoria_"))
    intermedio = carpeta / "intermedio.xlsx"
    wb.save(intermedio)

    # Etapa 2 (XML): ampliar tablas de las hojas procesadas
    z = zipfile.ZipFile(intermedio)
    partes = tx.mapa_partes(z)
    hojas_proc = {p["ws"].title for p in planes}
    wb_int, planes_t, avisos = tx.analizar(intermedio, z, partes, hojas=hojas_proc,
                                           quitar_filtros=False)
    esperadas = {p["tabla"]: p["ref_nueva"] for p in a_ampliar}
    obtenidas = {t["tabla"]: t["ref_nueva"] for t in planes_t}
    if esperadas != obtenidas:
        raise tx.Abortar(f"la ampliación de tablas no coincide con lo previsto: "
                         f"esperado {esperadas}, calculado {obtenidas}")
    filas_tab = [registro.nueva_fila(
        corrida_enc.fecha, registro.TIPO_ESTRUCTURA, t["hoja"], 1, None, None,
        f"Tabla {t['tabla']}", t["ref"], t["ref_nueva"],
        "Ampliación de tabla: se incorporan " + ", ".join(n for _, n in t["nuevas"]))
        for t in planes_t]
    registro.verificar_conteo(corrida_enc.filas + filas_tab, modificadas_celdas + len(planes_t))

    salida = carpeta / archivo.name
    modificadas, desde, hasta = tx.construir(intermedio, salida, z, planes_t, filas_tab, partes)
    print(f"\nResultado simulado: {salida}")

    print(f"\n--- Vista previa de '{registro.HOJA_REGISTRO}' ---")
    todas = corrida_enc.filas + filas_tab
    muestra = corrida_enc.filas[:n_enc] + corrida_val.filas[:3] + filas_tab
    for fila in muestra:
        print("    - " + " | ".join("(vacío)" if fila[h] is None else str(fila[h])
                                    for h in registro.REGISTRO_HEADERS))
    if len(todas) > len(muestra):
        print(f"    ... y {len(todas) - len(muestra)} filas más ({len(todas)} en total).")

    # ---------------- Verificaciones ----------------
    print("\n--- Verificación 1: openpyxl ---")
    ok1, err1 = tx.verificar_openpyxl(wb_int, salida, planes_t, filas_tab, desde,
                                      hojas=hojas_proc, exigir_sin_filtros=False)
    ok1b, err1b = comparar_con_original(wb_orig, salida, hojas_proc)
    filas_reg_total = hasta - (wb_orig[registro.HOJA_REGISTRO].max_row)
    ok1c = filas_reg_total == total
    print(f"  Tablas, columnas, valores y registro: {'OK' if ok1 else 'ERROR'}")
    print(f"  Hojas no procesadas (Marcas incluida) iguales al original: {'OK' if ok1b else 'ERROR'}")
    print(f"  Filas nuevas en RegistroCambios: {filas_reg_total} (esperadas {total}): "
          f"{'OK' if ok1c else 'ERROR'}")
    for e in err1 + err1b:
        print(f"  ERROR: {e}")

    print("\n--- Verificación 2: byte a byte (etapa XML contra la copia intermedia) ---")
    ok2, det2 = tx.verificar_bytes(z, salida, modificadas, partes)
    print(f"  {'OK' if ok2 else 'ERROR'} | {det2}")

    print("\n--- Verificación 3: Excel ---")
    nombres = {t["tabla"] for t in planes_t} | {p["tabla"] for p in planes if p["tabla"]}
    ok3, res, lin, intentos = tx.verificar_en_excel(salida, carpeta, planes_t,
                                                    exigir_sin_filtros=False, requeridas=nombres)
    print(("  Excel lo abrió sin pedido de reparación." if res.get("ok")
           else "  Excel NO pudo abrirlo.") + (f" (intentos: {intentos})" if intentos > 1 else ""))
    mostrar_sig = False
    for linea in lin:
        if linea.startswith("  ERROR") or linea.startswith("Excel:"):
            print(linea)
        elif "encabezados:" in linea:
            if mostrar_sig:
                print(linea)
        else:
            mostrar_sig = any(f" {n} " in linea for n in nombres)
            if mostrar_sig:
                print(linea)
    print(f"  {'OK' if ok3 else 'ERROR'}")

    z.close()
    if not (ok1 and ok1b and ok1c and ok2 and ok3):
        raise tx.Abortar("alguna verificación falló; no se toca el archivo real.")

    if not execute:
        print("\nEsto fue una simulación (dry-run). No se modificó el archivo real ni se creó backup.")
        return total

    # ---------------- Fase 3: ejecución ----------------
    if tx.sha(archivo.read_bytes()) != sha_origen:
        raise tx.Abortar("el archivo real cambió durante la corrida; volvé a correr el script.")
    backup = backup_archivo(archivo)
    print(f"\nBackup creado en: {backup}")
    try:
        shutil.copyfile(salida, archivo)
    except PermissionError:
        raise tx.Abortar("no se pudo reemplazar el archivo porque está abierto en Excel. "
                         "Cerralo y volvé a correr el script. El archivo real no se modificó.")
    if tx.sha(archivo.read_bytes()) != tx.sha(salida.read_bytes()):
        raise tx.Abortar(f"el archivo guardado no coincide con el verificado. Restaurá el backup: {backup}")
    print(f"Archivo reemplazado: {archivo}")
    print(f"Celdas modificadas: {modificadas_celdas} | Tablas ampliadas: {len(planes_t)} | "
          f"Filas agregadas a {registro.HOJA_REGISTRO}: {total}")
    return total


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(
        description="Asigna un valor fijo en la columna Categoria de las hojas filtradas.")
    p.add_argument("--filtro", action="append",
                   help="Fragmento del nombre de hoja (se puede repetir). "
                        "Sin distinguir mayúsculas ni acentos.")
    p.add_argument("--hoja", action="append",
                   help="Nombre exacto de la hoja (se puede repetir). Sin distinguir "
                        "mayúsculas ni acentos; no alcanza a otras hojas por coincidencia parcial.")
    p.add_argument("--valor", required=True, help="Texto de la categoría a escribir.")
    p.add_argument("--sobrescribir", action="store_true",
                   help="Reemplaza celdas que ya tienen otra categoría.")
    p.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    p.add_argument("--archivo", type=Path, default=ARCHIVO_PATH,
                   help="Libro a procesar (por defecto, el archivo real).")
    args = p.parse_args()
    try:
        procesar(args.filtro, args.valor, args.execute, args.sobrescribir, args.archivo, args.hoja)
    except (tx.Abortar, registro.RegistroError) as e:
        sys.exit(f"\nABORTADO: {e}")


if __name__ == "__main__":
    main()
