#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ampliar_tablas.py
------------------
1) Amplía hacia la derecha cada tabla de Excel (Table_N) que haya quedado más
   angosta que los encabezados de su hoja (por ejemplo ActualizadoSegun y
   Categoria creadas después de la tabla), para que esas columnas se ordenen
   y filtren junto con el resto y tomen el estilo de encabezado de la tabla.
2) Quita los criterios de los filtros guardados en las tablas y vuelve a
   mostrar las filas que esos filtros ocultaban. Los botones de filtro del
   encabezado quedan como están (no se agregan ni se quitan).

Se hace editando SOLO el XML de las partes necesarias del .xlsx:
  - xl/tables/tableN.xml de las tablas que cambian (ref, autoFilter,
    tableColumns),
  - la hoja cuyas filas ocultas por filtro se vuelven a mostrar
    (solo se quita el atributo hidden de esas filas),
  - la hoja RegistroCambios (filas nuevas como texto en línea).
Todo lo demás, incluida "Marcas Actualización Dinámica", se copia byte a byte.
openpyxl se usa solo para LEER (sin data_only=True) y para verificar.

Registro (TipoOperacion "Estructura"), una fila por cada cambio:
  - una por tabla ampliada:  Columna "Tabla <nombre>", ref anterior -> nueva;
  - una por filtro quitado:  Columna "Filtro <tabla> [<columna>]",
    criterio anterior -> "(sin filtro)".
La verificación de conteo compara esas filas con tablas ampliadas + filtros
quitados.

Verificaciones (siempre, antes de tocar el archivo real):
  1. openpyxl: refs, columnas de cada tabla = encabezados, sin filtros, sin
     filas ocultas por filtro, valores de todas las hojas idénticos y filas
     nuevas de RegistroCambios correctas.
  2. Byte a byte: Marcas y toda parte no modificada quedan idénticas.
  3. Excel (COM): el libro abre sin pedido de reparación, las tablas tienen la
     ref y las columnas esperadas, no hay filtros activos ni filas ocultas, y
     los encabezados nuevos se ven igual que el anterior (fondo y letra).

La lógica vive en tablas_xml.py (compartida con asignar_categoria.py).

USO (desde scripts/)
    python ampliar_tablas.py            -> inspección + dry-run (no toca el archivo real)
    python ampliar_tablas.py --execute  -> lo mismo + backup + reemplazo del archivo real
    --archivo RUTA                      -> procesa otro libro (por ejemplo, una copia de prueba)
"""

import argparse
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl.utils import get_column_letter

import registro
from tablas_xml import (HOJA_MARCAS, Abortar, abrir_en_excel, analizar, construir,
                        evaluar_excel, mapa_partes, sha, verificar_bytes,
                        verificar_en_excel, verificar_openpyxl)

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"

# --------------------------------------------------------------------------
# Proceso principal
# --------------------------------------------------------------------------
def backup_archivo(archivo):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{archivo.stem}_{ts}{archivo.suffix}"
    shutil.copy2(archivo, destino)
    return destino


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
    print(f"\n=== Ampliar tablas y quitar filtros guardados | Modo: {modo} ===\n")

    # ---------------- Fase 1: inspección ----------------
    wb_orig, planes, avisos = analizar(archivo, z, partes)
    print("--- Fase 1: Inspección ---")
    ampliadas = [p for p in planes if p["ref_nueva"] != p["ref"]]
    filtros = [(p, f) for p in planes for f in p["filtros"]]
    print(f"Tablas a ampliar: {len(ampliadas)}")
    for p in ampliadas:
        print(f"  - {p['hoja']:27} {p['tabla']:9} {p['ref']:9} -> {p['ref_nueva']:9} "
              f"entran: {', '.join(f'{get_column_letter(c)} {n}' for c, n in p['nuevas'])}"
              f"{'' if p['boton_filtro'] else '  (sin botones de filtro; no se agregan)'}")
    print(f"\nFiltros guardados a quitar: {len(filtros)}")
    for p, f in filtros:
        print(f"  - {p['hoja']} / {p['tabla']}: {f['criterio']} | filas que se vuelven a mostrar: "
              f"{len(p['a_mostrar'])} | {f['coherencia']}")
    for a in avisos:
        print(f"  Aviso: {a}")

    # ---------------- Fase 2: registro + libro simulado ----------------
    fecha = registro.timestamp_corrida()
    filas_reg = []
    for p in ampliadas:
        filas_reg.append(registro.nueva_fila(
            fecha, registro.TIPO_ESTRUCTURA, p["hoja"], 1, None, None, f"Tabla {p['tabla']}",
            p["ref"], p["ref_nueva"],
            "Ampliación de tabla: se incorporan " + ", ".join(n for _, n in p["nuevas"])))
    for p, f in filtros:
        filas_reg.append(registro.nueva_fila(
            fecha, registro.TIPO_ESTRUCTURA, p["hoja"], 1, None, None,
            f"Filtro {p['tabla']} [{f['columna']}]", f["criterio"], "(sin filtro)",
            f"Filtro guardado quitado | filas que se vuelven a mostrar: {len(p['a_mostrar'])}"))
    registro.verificar_conteo(filas_reg, len(ampliadas) + len(filtros))

    print("\n--- Fase 2: Resumen ---")
    print(f"  Tablas ampliadas: {len(ampliadas)} | Filtros quitados: {len(filtros)} | "
          f"Filas que se vuelven a mostrar: {sum(len(p['a_mostrar']) for p in planes)}")
    print(f"  Filas de RegistroCambios: {len(filas_reg)} "
          f"(= {len(ampliadas)} tablas + {len(filtros)} filtros)")
    print(f"\n--- Vista previa de '{registro.HOJA_REGISTRO}' ---")
    for fila in filas_reg:
        print("    - " + " | ".join("(vacío)" if fila[h] is None else str(fila[h])
                                    for h in registro.REGISTRO_HEADERS))

    if not filas_reg:
        print("\nNo hay nada para cambiar.")
        return 0

    carpeta = Path(tempfile.mkdtemp(prefix="ampliar_tablas_"))
    salida = carpeta / archivo.name
    modificadas, desde, hasta = construir(archivo, salida, z, planes, filas_reg, partes)
    print(f"\nLibro simulado: {salida}")
    print(f"Partes modificadas: {', '.join(sorted(modificadas))}")

    # ---------------- Verificaciones ----------------
    print("\n--- Verificación 1: openpyxl (refs, columnas, filtros, valores, registro) ---")
    ok1, err1 = verificar_openpyxl(wb_orig, salida, planes, filas_reg, desde)
    print("  OK" if ok1 else "\n".join(f"  ERROR: {e}" for e in err1))
    print(f"  RegistroCambios: filas {desde} a {hasta}")

    print("\n--- Verificación 2: byte a byte ---")
    ok2, det2 = verificar_bytes(z, salida, modificadas, partes)
    print(f"  {'OK' if ok2 else 'ERROR'} | {det2}")

    print("\n--- Verificación 3: Excel ---")
    copia_antes = carpeta / ("ANTES_" + archivo.name)
    shutil.copy2(archivo, copia_antes)
    _, lin_antes = evaluar_excel(abrir_en_excel(copia_antes, carpeta), planes, antes=True)
    print("  Antes (archivo actual):")
    print("\n".join(lin_antes))
    ok3, res_despues, lin, intentos = verificar_en_excel(salida, carpeta, planes)
    print(("  Después (libro simulado): Excel lo abrió sin pedido de reparación."
           if res_despues.get("ok") else "  Después (libro simulado): Excel NO pudo abrirlo.")
          + (f" (intentos: {intentos})" if intentos > 1 else ""))
    print("\n".join(lin))
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
    print(f"Tablas ampliadas: {len(ampliadas)} | Filtros quitados: {len(filtros)} | "
          f"Filas agregadas a {registro.HOJA_REGISTRO}: {len(filas_reg)}")
    return len(filas_reg)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Amplía las tablas y quita los filtros guardados.")
    ap.add_argument("--execute", action="store_true", help="Aplica los cambios reales.")
    ap.add_argument("--archivo", type=Path, default=ARCHIVO_PATH,
                    help="Libro a procesar (por defecto, el archivo real). Útil para probar sobre una copia.")
    args = ap.parse_args()
    try:
        procesar(args.execute, args.archivo)
    except (Abortar, registro.RegistroError) as e:
        sys.exit(f"\nABORTADO: {e}")


if __name__ == "__main__":
    main()
