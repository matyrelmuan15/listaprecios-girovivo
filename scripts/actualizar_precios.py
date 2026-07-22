"""
actualizar_precios.py
======================
Script genérico Tipo 1: actualización de precios por match de códigos.

Uso:
    python actualizar_precios.py <proveedor> [--execute]

Por defecto corre en modo DRY-RUN (no escribe nada, solo muestra el resumen
de lo que haría). Para ejecutar los cambios reales, agregar --execute.

Ejemplo:
    python actualizar_precios.py venzo
    python actualizar_precios.py venzo --execute

Reglas generales aplicadas: ver CLAUDE.md en la raíz del proyecto.
"""

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Configuración de rutas (relativas a la raíz del proyecto)
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
ARCHIVO_PATH = PROJECT_ROOT / "archivo" / "Lista de Precios Giro Vivo 2026.xlsx"
BACKUP_DIR = PROJECT_ROOT / "backup"
PROVEEDORES_DIR = PROJECT_ROOT / "proveedores"

HOJA_EXCLUIDA = "Marcas Actualización Dinámica"
HOJA_ENTRADA = "Hoja1"
HOJA_REGISTRO = "RegistroCambios"

# Formato argentino válido: dígitos, agrupados opcionalmente de a 3 con punto
# de miles, y coma decimal opcional (ej: "2.607,75", "607,75", "2.607").
ARG_PRICE_PATTERN = re.compile(r"^\d{1,3}(\.\d{3})*(,\d+)?$")

REGISTRO_HEADERS = [
    "Hoja", "Fila", "Codigo", "Marca", "Columna", "ValorAnterior", "ValorNuevo"
]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def normalize_code(value):
    """Normaliza un código para poder compararlo sin importar si viene
    como int, float o str (ej: 123, 123.0, '123', ' 123 ' deben ser iguales)."""
    if value is None:
        return ""
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else str(value)
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    try:
        f = float(s)
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s


def find_header_col(ws, header_name, max_search_rows=3):
    """Busca en las primeras filas de la hoja una celda cuyo texto
    (normalizado) coincida exactamente con header_name. Devuelve el índice
    de columna (1-based) o None si no se encuentra."""
    target = header_name.strip().casefold()
    for row in ws.iter_rows(min_row=1, max_row=max_search_rows):
        for cell in row:
            if cell.value is not None and str(cell.value).strip().casefold() == target:
                return cell.column
    return None


def parse_notas_params(notas_path: Path) -> dict:
    """Extrae los parámetros de la sección '## Parámetros' del NOTAS.md."""
    if not notas_path.exists():
        raise FileNotFoundError(f"No se encontró NOTAS.md en: {notas_path}")

    lines = notas_path.read_text(encoding="utf-8").splitlines()
    params = {}
    in_params_section = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("## Parámetros"):
            in_params_section = True
            continue
        if in_params_section:
            if stripped.startswith("##"):
                break
            if ":" in stripped and stripped:
                key, _, value = stripped.partition(":")
                params[key.strip()] = value.strip()

    if not params:
        raise ValueError(
            f"No se encontraron parámetros en la sección '## Parámetros' de {notas_path}"
        )
    return params


def parse_precio(value) -> float:
    """Convierte el valor de precio de Hoja1 a float, aceptando tanto
    números planos (1234.56) como texto en formato argentino
    (punto de miles, coma decimal: '2.607,75').

    Si el valor no es un número plano y tampoco matchea el patrón de
    formato argentino válido, no se asume ninguna conversión: se levanta
    ValueError para que el llamador lo trate como formato no reconocido."""
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        raise ValueError(f"precio vacío: {value!r}")
    try:
        return float(s)
    except ValueError:
        pass
    if not ARG_PRICE_PATTERN.match(s):
        raise ValueError(f"formato de precio no reconocido: {value!r}")
    s_arg = s.replace(".", "").replace(",", ".")
    return float(s_arg)


def apply_formula(precio: float, params: dict) -> float:
    """Aplica la fórmula de precio del proveedor según los parámetros de su NOTAS.md."""
    tipo = params.get("tipo")

    if tipo == "porcentaje_simple":
        porcentaje = float(params["porcentaje"])
        resultado = precio * (1 + porcentaje)

    elif tipo == "porcentaje_compuesto":
        # pasos: "0.70, 0.21"  -> aplica cada porcentaje en cadena, no sumados
        pasos = [float(p.strip()) for p in params["pasos"].split(",")]
        resultado = precio
        for paso in pasos:
            resultado = resultado * (1 + paso)

    else:
        raise ValueError(f"Tipo de fórmula desconocido en NOTAS.md: '{tipo}'")

    return round(resultado, 2)


def backup_archivo():
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"{ARCHIVO_PATH.stem}_{timestamp}{ARCHIVO_PATH.suffix}"
    shutil.copy2(ARCHIVO_PATH, destino)
    return destino


# ---------------------------------------------------------------------------
# Lógica principal
# ---------------------------------------------------------------------------
def procesar(proveedor: str, execute: bool):
    notas_path = PROVEEDORES_DIR / proveedor / "NOTAS.md"
    params = parse_notas_params(notas_path)

    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")
        sys.exit(1)

    wb = openpyxl.load_workbook(ARCHIVO_PATH)  # sin data_only=True

    if HOJA_ENTRADA not in wb.sheetnames:
        print(f"ERROR: no se encontró la hoja de entrada '{HOJA_ENTRADA}'.")
        sys.exit(1)

    ws_hoja1 = wb[HOJA_ENTRADA]
    col_codigo = find_header_col(ws_hoja1, "Codigo")
    col_precio = find_header_col(ws_hoja1, "Precio")

    if col_codigo is None or col_precio is None:
        print("ERROR: no se encontraron las columnas 'Codigo' y/o 'Precio' en Hoja1.")
        sys.exit(1)

    # Construir diccionario código normalizado -> precio original (de Hoja1)
    codigos_hoja1 = {}  # normalized_code -> precio
    for row in ws_hoja1.iter_rows(min_row=2):
        celda_codigo = row[col_codigo - 1]
        celda_precio = row[col_precio - 1]
        if celda_codigo.value is None:
            continue
        norm = normalize_code(celda_codigo.value)
        if norm == "":
            continue
        codigos_hoja1[norm] = celda_precio.value

    # Pre-parseo de precios de Hoja1: separa los que tienen formato
    # reconocido (número plano o argentino válido) de los que no, para no
    # convertir a ciegas ni perderlos en silencio.
    precios_parsed = {}  # norm -> float
    avisos_formato = []  # lista de (norm, valor_original)
    for norm, valor_original in codigos_hoja1.items():
        try:
            precios_parsed[norm] = parse_precio(valor_original)
        except (TypeError, ValueError):
            avisos_formato.append((norm, valor_original))

    codigos_no_encontrados = set(precios_parsed.keys())
    registro_rows = []  # lista de dicts con REGISTRO_HEADERS
    dry_run_detalle = []  # lista de dicts con detalle original -> parseado (solo dry-run)
    resumen_por_hoja = {}  # hoja -> {"precios": n, "fechas": n}
    total_coincidencias = 0

    fecha_hoy = datetime.now().strftime("%d/%m/%Y")

    for sheet_name in wb.sheetnames:
        if sheet_name in (HOJA_ENTRADA, HOJA_EXCLUIDA, HOJA_REGISTRO):
            continue

        ws = wb[sheet_name]
        col_codfac = find_header_col(ws, "CodigoFactura") or find_header_col(ws, "CodigoFac")
        col_marca = find_header_col(ws, "Marca")
        col_precio_iva = find_header_col(ws, "PrecioConIVA")
        col_fecha = find_header_col(ws, "FechaActualizacion")

        if col_codfac is None or col_precio_iva is None or col_fecha is None:
            # Esta hoja no es una hoja de destino válida para este prompt
            continue

        precios_mod = 0
        fechas_mod = 0

        for row in ws.iter_rows(min_row=2):
            celda_codfac = row[col_codfac - 1]
            if celda_codfac.value is None:
                continue

            elementos = [normalize_code(e) for e in str(celda_codfac.value).split(",")]
            elementos = [e for e in elementos if e]

            match = next((e for e in elementos if e in precios_parsed), None)
            if match is None:
                continue

            precio_original = codigos_hoja1[match]
            precio_parseado = precios_parsed[match]
            precio_nuevo = apply_formula(precio_parseado, params)

            fila_num = celda_codfac.row
            marca_valor = row[col_marca - 1].value if col_marca else None

            celda_precio_iva = row[col_precio_iva - 1]
            celda_fecha = row[col_fecha - 1]

            valor_anterior_precio = celda_precio_iva.value
            valor_anterior_fecha = celda_fecha.value

            registro_rows.append({
                "Hoja": sheet_name, "Fila": fila_num, "Codigo": match,
                "Marca": marca_valor, "Columna": "PrecioConIVA",
                "ValorAnterior": valor_anterior_precio, "ValorNuevo": precio_nuevo,
            })
            registro_rows.append({
                "Hoja": sheet_name, "Fila": fila_num, "Codigo": match,
                "Marca": marca_valor, "Columna": "FechaActualizacion",
                "ValorAnterior": valor_anterior_fecha, "ValorNuevo": fecha_hoy,
            })

            if not execute:
                dry_run_detalle.append({
                    "Hoja": sheet_name, "Fila": fila_num, "Codigo": match,
                    "Original": precio_original, "Parseado": precio_parseado,
                })

            if execute:
                celda_precio_iva.value = precio_nuevo
                celda_fecha.value = fecha_hoy

            precios_mod += 1
            fechas_mod += 1
            total_coincidencias += 1
            codigos_no_encontrados.discard(match)

        if precios_mod > 0:
            resumen_por_hoja[sheet_name] = {"precios": precios_mod, "fechas": fechas_mod}

    # ---- Columna E de Hoja1: códigos sin match ----
    no_encontrados_ordenados = sorted(codigos_no_encontrados)
    if execute:
        # Limpiar columna E existente (a partir de la fila 2) antes de reescribir
        max_row_actual = ws_hoja1.max_row
        for r in range(2, max_row_actual + 1):
            ws_hoja1.cell(row=r, column=5).value = None
        for i, codigo in enumerate(no_encontrados_ordenados, start=2):
            ws_hoja1.cell(row=i, column=5).value = codigo

    # ---- RegistroCambios ----
    if execute:
        if HOJA_REGISTRO not in wb.sheetnames:
            ws_reg = wb.create_sheet(HOJA_REGISTRO)
            ws_reg.append(REGISTRO_HEADERS)
        else:
            ws_reg = wb[HOJA_REGISTRO]

        for r in registro_rows:
            ws_reg.append([r[h] for h in REGISTRO_HEADERS])

        # Verificación de conteo obligatoria
        celdas_modificadas = total_coincidencias * 2  # precio + fecha por cada match
        if len(registro_rows) != celdas_modificadas:
            print(
                "ERROR DE INTEGRIDAD: la cantidad de filas de registro "
                f"({len(registro_rows)}) no coincide con la cantidad de celdas "
                f"modificadas ({celdas_modificadas}). No se guarda el archivo."
            )
            sys.exit(1)

    # ---------------------------------------------------------------
    # Resumen
    # ---------------------------------------------------------------
    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Proveedor: {proveedor} | Modo: {modo} ===\n")
    print(f"Fórmula aplicada: {params}")
    print(f"Total de coincidencias encontradas: {total_coincidencias}")
    print(f"Total de precios modificados: {total_coincidencias}")
    print(f"Total de fechas modificadas: {total_coincidencias}")
    print("\nDesglose por hoja:")
    if resumen_por_hoja:
        for hoja, datos in resumen_por_hoja.items():
            print(f"  - {hoja}: {datos['precios']} precios, {datos['fechas']} fechas")
    else:
        print("  (sin coincidencias en ninguna hoja)")
    print(f"\nCódigos de Hoja1 sin encontrar: {len(no_encontrados_ordenados)}")
    if no_encontrados_ordenados:
        print(f"  {no_encontrados_ordenados}")

    avisos_formato_ordenados = sorted(avisos_formato, key=lambda item: item[0])
    print(f"\nAVISO: formato de precio no reconocido: {len(avisos_formato_ordenados)}")
    if avisos_formato_ordenados:
        for norm, valor_original in avisos_formato_ordenados:
            print(f"  - Código '{norm}': valor original {valor_original!r}")

    if not execute and dry_run_detalle:
        print("\nDetalle de filas a modificar (Original -> Parseado):")
        for item in dry_run_detalle:
            print(
                f"  - [{item['Hoja']}] fila {item['Fila']} código '{item['Codigo']}': "
                f"Original: {item['Original']!r} -> Parseado: {item['Parseado']}"
            )

    if execute:
        backup_path = backup_archivo()
        print(f"\nBackup creado en: {backup_path}")
        wb.save(ARCHIVO_PATH)
        print(f"Archivo guardado: {ARCHIVO_PATH}")
        print(f"Filas agregadas a RegistroCambios: {len(registro_rows)}")
    else:
        print(
            "\nEsto fue una simulación (dry-run). No se modificó el archivo ni se "
            "creó backup. Para ejecutar los cambios reales, correr de nuevo con --execute."
        )


def main():
    parser = argparse.ArgumentParser(
        description="Actualiza precios (Tipo 1) para un proveedor según su NOTAS.md."
    )
    parser.add_argument("proveedor", help="Nombre de la carpeta del proveedor en proveedores/")
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Ejecuta los cambios reales (por defecto corre en modo dry-run).",
    )
    args = parser.parse_args()
    procesar(args.proveedor, args.execute)


if __name__ == "__main__":
    main()
