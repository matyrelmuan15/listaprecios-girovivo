"""
actualizar_precios.py
======================
Script genérico Tipo 1: actualización de precios por match de códigos.

MODO GUIADO (recomendado)
--------------------------
Si lo corrés sin argumentos:

    python actualizar_precios.py

el script te muestra un menú para elegir el proveedor con las flechas
↑ / ↓ y Enter, corre automáticamente la vista previa (dry-run) y al final
te pregunta, también con un menú de flechas, si confirmás aplicar los
cambios reales al archivo. No hace falta recordar nombres de proveedor
ni flags.

MODO POR LÍNEA DE COMANDOS (para automatizar o repetir una corrida exacta)
----------------------------------------------------------------------------
    python actualizar_precios.py <proveedor> [--dolar VALOR] [--execute]

Por defecto corre en modo DRY-RUN (no escribe nada, solo muestra el resumen
de lo que haría). Para ejecutar los cambios reales, agregar --execute.

--dolar es opcional: si se indica (ej. 1350, 1350,50 o 1350.50), además de
PrecioConIVA y FechaActualizacion se escribe ese valor en DolarReferencia de
las mismas filas actualizadas por match de código. Las hojas sin columna
DolarReferencia se actualizan igual (precio y fecha) y se informan con un
AVISO. Sin --dolar, DolarReferencia no se toca. El valor sigue la regla
única de dolar.py (compartida con los demás scripts): la coma se toma como punto decimal, el punto es
siempre decimal (no hay separador de miles: 1.350 es 1.35) y se escribe tal
cual, sin redondear.

Ejemplo:
    python actualizar_precios.py venzo
    python actualizar_precios.py venzo --dolar 1350,50
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

import actualizado_segun as act
import formulas
import registro
from asignar_dolar import COL_DOLAR
from dolar import parse_dolar
from ia_revision import (
    formatear_items_marcados,
    imprimir_items_marcados_json,
    revisar_cambios_con_ia,
)

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
HOJA_REGISTRO = registro.HOJA_REGISTRO
# Hojas de informe que nunca son destino (además de Hoja1, la excluida y el registro).
HOJAS_INFORME = (act.HOJA_NO_ACTUALIZADOS,) + act.HOJAS_REEMPLAZADAS

# El formato de RegistroCambios es único y compartido: vive en registro.py.
# Este script no define encabezados propios.


# ---------------------------------------------------------------------------
# Menú interactivo por consola (flechas + Enter) — solo Windows
# ---------------------------------------------------------------------------
def listar_proveedores():
    """Lee dinámicamente las subcarpetas de proveedores/ que tengan NOTAS.md."""
    if not PROVEEDORES_DIR.exists():
        return []
    proveedores = []
    for carpeta in sorted(PROVEEDORES_DIR.iterdir(), key=lambda p: p.name.lower()):
        if carpeta.is_dir() and (carpeta / "NOTAS.md").exists():
            proveedores.append(carpeta.name)
    return proveedores


def menu_seleccionar(opciones, titulo="Seleccioná una opción"):
    """Menú de consola navegable con flechas ↑ / ↓ y Enter (requiere Windows).

    Devuelve el índice elegido, o None si se cancela con Esc.
    """
    try:
        import msvcrt
    except ImportError:
        # Si no estamos en Windows, se cae a un menú por número (respaldo).
        return _menu_seleccionar_numerico(opciones, titulo)

    seleccion = 0

    def dibujar():
        print("\n" + titulo)
        print("=" * len(titulo))
        for i, op in enumerate(opciones):
            marca = "  > " if i == seleccion else "    "
            print(f"{marca}{op}")
        print("\n(Flechas arriba/abajo para moverte, Enter para confirmar, Esc para cancelar)")

    dibujar()
    while True:
        tecla = msvcrt.getch()
        if tecla in (b"\xe0", b"\x00"):  # prefijo de tecla especial en Windows
            tecla2 = msvcrt.getch()
            if tecla2 == b"H":  # flecha arriba
                seleccion = (seleccion - 1) % len(opciones)
                dibujar()
            elif tecla2 == b"P":  # flecha abajo
                seleccion = (seleccion + 1) % len(opciones)
                dibujar()
        elif tecla == b"\r":  # Enter
            return seleccion
        elif tecla == b"\x1b":  # Esc
            return None


def _menu_seleccionar_numerico(opciones, titulo):
    """Respaldo simple (por si el script se corre fuera de Windows/consola real)."""
    print("\n" + titulo)
    print("=" * len(titulo))
    for i, op in enumerate(opciones, start=1):
        print(f"  {i}) {op}")
    while True:
        valor = input("Elegí un número: ").strip()
        if valor.isdigit() and 1 <= int(valor) <= len(opciones):
            return int(valor) - 1
        print("  Opción inválida, intentá de nuevo.")


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


def quitar_espacios_codigo(value):
    """Quita TODOS los espacios de un código de Hoja1 (extremos, del medio y
    no separables). Las comas quedan intactas, así que un código múltiple
    'A 1, B 2' pasa a 'A1,B2' y se sigue separando igual por coma.

    Devuelve el código limpio, o None si no hay nada que cambiar (valores no
    texto, sin espacios, o que quedarían vacíos)."""
    if not isinstance(value, str):
        return None
    limpio = re.sub(r"[\s  ]", "", value)
    if limpio == value or limpio == "":
        return None
    return limpio


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


def _validar_grupos_miles(partes, original):
    """Valida que los grupos separados por el separador de miles sean
    correctos: el primero de 1 a 3 dígitos y los siguientes de 3 exactos."""
    if not (1 <= len(partes[0]) <= 3) or any(len(g) != 3 for g in partes[1:]):
        raise ValueError(f"formato de precio no reconocido: {original!r}")


def parse_precio(value) -> float:
    """Convierte el precio de Hoja1 a float, sin importar cómo venga escrito.

    Acepta:
      - Números reales de Excel (int/float): se usan tal cual.
      - Texto con símbolo de moneda o espacios: '$ 2.607,75', 'ARS 2607,75'.
      - Formato argentino: '2.607,75', '2607,75', '3.030.000'.
      - Formato anglosajón: '2,607.75', '2607.75', '3,030,000'.

    Cómo decide cuál es el separador decimal:
      - Si hay punto y coma, el que aparece último es el decimal.
      - Si un mismo separador aparece varias veces (1.234.567), es de miles.
      - Si aparece una sola vez: con exactamente 3 dígitos después se toma
        como miles (los precios no llevan 3 decimales); si no, es decimal.

    Si después de limpiar queda algo que no es un número, levanta ValueError
    para que la fila se informe como "formato no reconocido" y no se toque."""
    if isinstance(value, bool):
        raise ValueError(f"formato de precio no reconocido: {value!r}")
    if isinstance(value, (int, float)):
        return float(value)

    original = value
    s = str(value)
    # Quitar símbolos de moneda y espacios (incluidos los no separables).
    s = re.sub(r"(?i)u\$s|us\$|ars|usd|\$", "", s)
    s = re.sub(r"[\s\u00a0\u202f']", "", s)
    if not s:
        raise ValueError(f"precio vacío: {original!r}")
    if not re.fullmatch(r"\d[\d.,]*\d|\d", s):
        raise ValueError(f"formato de precio no reconocido: {original!r}")

    tiene_punto = "." in s
    tiene_coma = "," in s

    if tiene_punto and tiene_coma:
        decimal = "." if s.rfind(".") > s.rfind(",") else ","
        miles = "," if decimal == "." else "."
        entero, _, decimales = s.rpartition(decimal)
        if decimal in entero:
            raise ValueError(f"formato de precio no reconocido: {original!r}")
        partes = entero.split(miles)
        _validar_grupos_miles(partes, original)
        return float("".join(partes) + "." + decimales)

    if tiene_punto or tiene_coma:
        sep = "." if tiene_punto else ","
        partes = s.split(sep)
        if len(partes) > 2:
            _validar_grupos_miles(partes, original)
            return float("".join(partes))
        entero, decimales = partes
        if len(decimales) == 3 and entero != "0":
            return float(entero + decimales)
        return float(entero + "." + decimales)

    return float(s)


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
def procesar(proveedor: str, execute: bool, dolar=None):
    # El dólar opcional se valida antes de abrir el libro: si es inválido se
    # aborta sin tocar nada.
    if dolar is not None:
        try:
            dolar = parse_dolar(dolar)
        except ValueError as e:
            print(f"ERROR: {e} No se abrió ni modificó el archivo.")
            sys.exit(1)

    notas_path = PROVEEDORES_DIR / proveedor / "NOTAS.md"
    params = parse_notas_params(notas_path)
    # Regla opcional por proveedor: no es parte de la fórmula de precio, así
    # que se saca de params para no alterar el texto de la fórmula.
    limpiar_espacios = params.pop("limpiar_espacios_codigo", "").strip().casefold() in ("si", "sí")
    # Marcas del proveedor (parámetro "marcas:" del NOTAS.md): definen qué
    # filas del Excel deberían haberse actualizado, para la hoja NoActualizados.
    marcas_texto = params.pop("marcas", "").strip()
    marcas_prov = act.parse_marcas(marcas_texto)

    if not ARCHIVO_PATH.exists():
        print(f"ERROR: no se encontró el archivo de trabajo en: {ARCHIVO_PATH}")
        sys.exit(1)

    wb = openpyxl.load_workbook(ARCHIVO_PATH)  # sin data_only=True

    # Validación fail-fast del formato de RegistroCambios: se corre apenas
    # cargado el libro y ANTES del backup, para no dejar un backup inútil ni
    # apilar filas desalineadas bajo encabezados que no corresponden.
    try:
        registro.validar_encabezados(wb)
    except registro.RegistroError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)

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
    # Si el proveedor tiene limpiar_espacios_codigo: si, el matching usa ya
    # los códigos sin espacios (también en el dry-run).
    codigos_hoja1 = {}  # normalized_code -> precio
    limpiezas_codigo = []  # lista de (celda, valor_anterior, valor_limpio)
    for row in ws_hoja1.iter_rows(min_row=2):
        celda_codigo = row[col_codigo - 1]
        celda_precio = row[col_precio - 1]
        if celda_codigo.value is None:
            continue
        valor_codigo = celda_codigo.value
        if limpiar_espacios:
            limpio = quitar_espacios_codigo(valor_codigo)
            if limpio is not None:
                limpiezas_codigo.append((celda_codigo, valor_codigo, limpio))
                valor_codigo = limpio
        norm = normalize_code(valor_codigo)
        if norm == "":
            continue
        codigos_hoja1[norm] = celda_precio.value

    # Pre-parseo de precios de Hoja1: normaliza cualquier formato habitual
    # (número de Excel, argentino, anglosajón, con $ o espacios) y separa los
    # que no se pueden interpretar, para no convertir a ciegas ni perderlos.
    precios_parsed = {}  # norm -> float
    avisos_formato = []  # lista de (norm, valor_original)
    for norm, valor_original in codigos_hoja1.items():
        try:
            precios_parsed[norm] = parse_precio(valor_original)
        except (TypeError, ValueError):
            avisos_formato.append((norm, valor_original))

    codigos_no_encontrados = set(precios_parsed.keys())
    detalle_corrida = (
        f"Proveedor: {proveedor} | Fórmula: "
        + ", ".join(f"{k}={v}" for k, v in params.items())
    )
    if dolar is not None:
        detalle_corrida += f" | Dólar: {dolar}"
    if limpiar_espacios:
        detalle_corrida += " | Limpieza de espacios en códigos de Hoja1"
    corrida = registro.RegistroCorrida(registro.TIPO_PRECIOS, detalle_corrida)

    # Una fila de registro por cada código de Hoja1 limpiado (regla 8).
    for celda_codigo, anterior, limpio in limpiezas_codigo:
        corrida.agregar(
            hoja=HOJA_ENTRADA, fila=celda_codigo.row, nombre_producto=None,
            marca=None, columna="Codigo", valor_anterior=anterior, valor_nuevo=limpio,
        )
        if execute:
            celda_codigo.value = limpio
    dry_run_detalle = []  # lista de dicts con detalle original -> parseado (solo dry-run)
    cambios_precio = []  # lista para la revisión automática de anomalías (IA)
    resumen_por_hoja = {}  # hoja -> {"precios": n, "fechas": n, "dolares": n}
    total_coincidencias = 0
    total_dolares = 0  # matches en los que se escribió DolarReferencia
    hojas_sin_dolar = []  # hojas con matches pero sin columna DolarReferencia
    filas_tocadas = {}  # hoja -> set(filas), para verificar fórmulas
    precios_nuevos = {}  # (hoja, fila) -> PrecioConIVA que va a quedar
    total_actualizado_segun = 0  # celdas ActualizadoSegun que cambian a "Lista proveedor"
    hojas_sin_actualizado = []  # hojas con matches pero sin columna ActualizadoSegun
    no_actualizados = []  # filas de las marcas del proveedor que no recibieron precio
    codigos_formato_invalido = {norm for norm, _ in avisos_formato}

    fecha_hoy = datetime.now().strftime("%d/%m/%Y")

    for sheet_name in wb.sheetnames:
        if sheet_name in (HOJA_ENTRADA, HOJA_EXCLUIDA, HOJA_REGISTRO) + HOJAS_INFORME:
            continue

        ws = wb[sheet_name]
        col_codfac = find_header_col(ws, "CodigoFactura") or find_header_col(ws, "CodigoFac")
        col_marca = find_header_col(ws, "Marca")
        col_precio_iva = find_header_col(ws, "PrecioConIVA")
        col_fecha = find_header_col(ws, "FechaActualizacion")
        col_nombre = find_header_col(ws, "NombreProducto")
        col_dolar = find_header_col(ws, COL_DOLAR) if dolar is not None else None
        col_codigodux = find_header_col(ws, "CodigoDux")
        col_actualizado = act.buscar_columna(ws)

        if col_codfac is None or col_precio_iva is None or col_fecha is None:
            # Esta hoja no es una hoja de destino válida para este prompt
            continue

        precios_mod = 0
        fechas_mod = 0
        dolares_mod = 0

        for row in ws.iter_rows(min_row=2):
            celda_codfac = row[col_codfac - 1]
            elementos = []
            if celda_codfac.value is not None:
                elementos = [normalize_code(e) for e in str(celda_codfac.value).split(",")]
                elementos = [e for e in elementos if e]

            match = next((e for e in elementos if e in precios_parsed), None)
            if match is None:
                # Fila de una marca de este proveedor que no recibe precio:
                # va a la hoja NoActualizados con el motivo.
                marca_fila = row[col_marca - 1].value if col_marca else None
                if marcas_prov and act.marca_coincide(marca_fila, marcas_prov):
                    if not elementos:
                        motivo = act.MOTIVO_SIN_CODIGO
                    elif any(e in codigos_formato_invalido for e in elementos):
                        motivo = act.MOTIVO_FORMATO
                    else:
                        motivo = act.MOTIVO_NO_EN_LISTA
                    no_actualizados.append({
                        "Hoja": sheet_name,
                        "Fila": celda_codfac.row,
                        "CodigoDux": row[col_codigodux - 1].value if col_codigodux else None,
                        "NombreProducto": row[col_nombre - 1].value if col_nombre else None,
                        "CodigoFactura": celda_codfac.value,
                        "Marca": marca_fila,
                        "PrecioConIVA": row[col_precio_iva - 1].value,
                        "FechaActualizacion": row[col_fecha - 1].value,
                        "ActualizadoSegun": (
                            row[col_actualizado - 1].value if col_actualizado else None
                        ),
                        "Motivo": motivo,
                    })
                continue

            precio_original = codigos_hoja1[match]
            precio_parseado = precios_parsed[match]
            precio_nuevo = apply_formula(precio_parseado, params)

            fila_num = celda_codfac.row
            marca_valor = row[col_marca - 1].value if col_marca else None
            nombre_valor = row[col_nombre - 1].value if col_nombre else None

            celda_precio_iva = row[col_precio_iva - 1]
            celda_fecha = row[col_fecha - 1]

            valor_anterior_precio = celda_precio_iva.value
            valor_anterior_fecha = celda_fecha.value

            cambios_precio.append({
                "hoja": sheet_name,
                "codigo": match,
                "producto": nombre_valor,
                "valor_anterior": valor_anterior_precio,
                "valor_nuevo": precio_nuevo,
            })

            # Una fila de registro por cada celda modificada (regla 8):
            # PrecioConIVA y FechaActualizacion.
            corrida.agregar(
                hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                marca=marca_valor, columna="PrecioConIVA",
                valor_anterior=valor_anterior_precio, valor_nuevo=precio_nuevo,
            )
            corrida.agregar(
                hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                marca=marca_valor, columna="FechaActualizacion",
                valor_anterior=valor_anterior_fecha, valor_nuevo=fecha_hoy,
            )

            # DolarReferencia solo si se indicó --dolar y la hoja tiene la
            # columna: una fila de registro más por esa celda.
            celda_dolar = row[col_dolar - 1] if col_dolar else None
            if celda_dolar is not None:
                corrida.agregar(
                    hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                    marca=marca_valor, columna=COL_DOLAR,
                    valor_anterior=celda_dolar.value, valor_nuevo=dolar,
                )

            # ActualizadoSegun = "Lista proveedor" (solo si la hoja tiene la
            # columna y el valor cambia: si ya decía eso, no es una celda
            # modificada y no se registra).
            celda_actualizado = row[col_actualizado - 1] if col_actualizado else None
            escribir_actualizado = (
                celda_actualizado is not None and celda_actualizado.value != act.VALOR_LISTA
            )
            if escribir_actualizado:
                corrida.agregar(
                    hoja=sheet_name, fila=fila_num, nombre_producto=nombre_valor,
                    marca=marca_valor, columna=act.COL_ACTUALIZADO,
                    valor_anterior=celda_actualizado.value, valor_nuevo=act.VALOR_LISTA,
                )
                total_actualizado_segun += 1

            if not execute:
                dry_run_detalle.append({
                    "Hoja": sheet_name, "Fila": fila_num, "Codigo": match,
                    "Original": precio_original, "Parseado": precio_parseado,
                })

            if execute:
                celda_precio_iva.value = precio_nuevo
                celda_fecha.value = fecha_hoy
                if celda_dolar is not None:
                    celda_dolar.value = dolar
                if escribir_actualizado:
                    celda_actualizado.value = act.VALOR_LISTA

            filas_tocadas.setdefault(sheet_name, set()).add(fila_num)
            precios_nuevos[(sheet_name, fila_num)] = precio_nuevo

            precios_mod += 1
            fechas_mod += 1
            if celda_dolar is not None:
                dolares_mod += 1
                total_dolares += 1
            total_coincidencias += 1
            codigos_no_encontrados.discard(match)

        if precios_mod > 0:
            resumen_por_hoja[sheet_name] = {
                "precios": precios_mod, "fechas": fechas_mod, "dolares": dolares_mod,
            }
            if dolar is not None and col_dolar is None:
                hojas_sin_dolar.append(sheet_name)
            if col_actualizado is None:
                hojas_sin_actualizado.append(sheet_name)

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
    # Verificación de conteo obligatoria (regla 8): 2 celdas por match
    # (PrecioConIVA + FechaActualizacion), más 1 por cada match en que se
    # escribió DolarReferencia y 1 por cada código de Hoja1 limpiado de
    # espacios. Se corre antes de escribir, y si falla se
    # aborta sin guardar.
    # Además se verifican las fórmulas de PrecioContado y PrecioSinIVA de las
    # filas actualizadas: si alguna perdió la fórmula, se repone y se registra
    # (una fila de registro por celda, igual que el resto).
    cambios_formulas, _ = formulas.revisar(
        wb, filas=filas_tocadas, precios_nuevos=precios_nuevos, redondear_coniva=False
    )
    formulas.registrar(corrida, cambios_formulas)
    if execute:
        formulas.aplicar(wb, cambios_formulas)

    celdas_modificadas = (
        total_coincidencias * 2 + total_dolares + len(cambios_formulas) + len(limpiezas_codigo)
        + total_actualizado_segun
    )
    try:
        corrida.verificar_conteo(celdas_modificadas)
        if execute:
            corrida.escribir(wb)
    except registro.RegistroError as e:
        print(f"\n{e}")
        sys.exit(1)

    # ---------------------------------------------------------------
    # Resumen
    # ---------------------------------------------------------------
    modo = "EJECUCIÓN" if execute else "DRY-RUN (sin escribir cambios)"
    print(f"\n=== Proveedor: {proveedor} | Modo: {modo} ===\n")
    print(f"Fórmula aplicada: {params}")
    if dolar is not None:
        print(f"Dólar asignado: {dolar}")
    print(f"Total de coincidencias encontradas: {total_coincidencias}")
    print(f"Total de precios modificados: {total_coincidencias}")
    print(f"Total de fechas modificadas: {total_coincidencias}")
    if dolar is not None:
        print(f"Total de DolarReferencia modificados: {total_dolares}")
    print("\nDesglose por hoja:")
    if resumen_por_hoja:
        for hoja, datos in resumen_por_hoja.items():
            linea = f"  - {hoja}: {datos['precios']} precios, {datos['fechas']} fechas"
            if dolar is not None:
                linea += f", {datos['dolares']} dólares"
            print(linea)
    else:
        print("  (sin coincidencias en ninguna hoja)")
    if hojas_sin_dolar:
        print(
            f"\nAVISO: no se pudo asignar el dólar en {len(hojas_sin_dolar)} hoja(s) "
            f"sin columna '{COL_DOLAR}' (precio y fecha sí se actualizan):"
        )
        for hoja in hojas_sin_dolar:
            print(f"  - {hoja}")
    if limpiar_espacios:
        print(f"\nCódigos de Hoja1 con espacios (se limpian): {len(limpiezas_codigo)}")
        for _, anterior, limpio in limpiezas_codigo[:10]:
            print(f"  - {anterior!r} -> {limpio!r}")
        if len(limpiezas_codigo) > 10:
            print(f"  ... y {len(limpiezas_codigo) - 10} más.")

    if hojas_sin_actualizado:
        print(
            f"\nAVISO: {len(hojas_sin_actualizado)} hoja(s) sin columna "
            f"'{act.COL_ACTUALIZADO}' (precio y fecha sí se actualizan). "
            "Corré crear_columna_actualizado.py para agregarla:"
        )
        for hoja in hojas_sin_actualizado:
            print(f"  - {hoja}")
    print(f"\nCeldas {act.COL_ACTUALIZADO} que pasan a '{act.VALOR_LISTA}': {total_actualizado_segun}")

    # ---- Productos del proveedor que no se actualizaron ----
    print("\n=== PRODUCTOS NO ACTUALIZADOS ===")
    if not marcas_prov:
        print(
            f"AVISO: el NOTAS.md de '{proveedor}' no tiene el parámetro 'marcas:', "
            f"así que no se puede armar la hoja '{act.HOJA_NO_ACTUALIZADOS}'."
        )
    else:
        print(f"Marcas del proveedor: {marcas_texto}")
        print(f"Filas de esas marcas que NO recibieron precio: {len(no_actualizados)}")
        for item in no_actualizados:
            print(
                f"  - [{item['Hoja']}] fila {item['Fila']} | {item['NombreProducto']} "
                f"| código: {item['CodigoFactura']} | {item['Motivo']}"
            )
        if execute:
            act.escribir_hoja_no_actualizados(
                wb, proveedor, marcas_texto, fecha_hoy, no_actualizados
            )
            print(f"Detalle guardado en la hoja '{act.HOJA_NO_ACTUALIZADOS}'.")
        else:
            print(f"(Al aplicar, este listado se guarda en la hoja '{act.HOJA_NO_ACTUALIZADOS}'.)")

    print(f"\nCódigos de Hoja1 sin encontrar: {len(no_encontrados_ordenados)}")
    if no_encontrados_ordenados:
        print(f"  {no_encontrados_ordenados}")

    avisos_formato_ordenados = sorted(avisos_formato, key=lambda item: item[0])
    print(f"\nAVISO: formato de precio no reconocido: {len(avisos_formato_ordenados)}")
    if avisos_formato_ordenados:
        for norm, valor_original in avisos_formato_ordenados:
            print(f"  - Código '{norm}': valor original {valor_original!r}")

    formulas.imprimir_resumen(
        cambios_formulas, [], titulo="Fórmulas de contado y sin IVA en las filas actualizadas"
    )

    # ---- Revisión automática de posibles anomalías (estadística + IA) ----
    print("\n--- Revisión automática de posibles anomalías ---")
    analisis_ia, aviso_ia = revisar_cambios_con_ia(
        cambios_precio,
        contexto_extra=(
            f"Actualización de precios (Tipo 1) para el proveedor '{proveedor}'. "
            f"Fórmula: {params}."
        ),
    )
    if aviso_ia:
        print(aviso_ia)
    if analisis_ia:
        print(analisis_ia)
    detalle_marcados = formatear_items_marcados(cambios_precio)
    if detalle_marcados:
        print(detalle_marcados)
    imprimir_items_marcados_json(cambios_precio)

    if not execute:
        registro.imprimir_previsualizacion(corrida)

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
        print(f"Filas agregadas a {HOJA_REGISTRO}: {len(corrida)}")
    else:
        print(
            "\nEsto fue una simulación (dry-run). No se modificó el archivo ni se "
            "creó backup."
        )

    return total_coincidencias


def modo_guiado():
    """Flujo guiado por menú: elegir proveedor con flechas, correr dry-run
    automáticamente y preguntar con otro menú si se confirma la ejecución real."""
    proveedores = listar_proveedores()
    if not proveedores:
        sys.exit(
            "ERROR: no se encontró ningún proveedor con NOTAS.md dentro de "
            f"'{PROVEEDORES_DIR}'."
        )

    idx = menu_seleccionar(proveedores, "Elegí el proveedor")
    if idx is None:
        print("\nOperación cancelada. No se modificó nada.")
        return
    proveedor = proveedores[idx]

    print(f"\nCorriendo vista previa (dry-run) para '{proveedor}'...")
    coincidencias = procesar(proveedor, execute=False)

    if coincidencias == 0:
        print("\nNo hay ninguna coincidencia para modificar. No hace falta ejecutar nada.")
        return

    opciones = [
        "Sí, aplicar estos cambios reales al archivo",
        "No, salir sin modificar nada",
    ]
    idx_confirm = menu_seleccionar(
        opciones, "¿Confirmás aplicar estos cambios al archivo real?"
    )
    if idx_confirm == 0:
        print(f"\nAplicando cambios reales para '{proveedor}'...")
        procesar(proveedor, execute=True)
    else:
        print("\nNo se aplicó ningún cambio. Podés volver a correr el script cuando quieras.")


def main():
    parser = argparse.ArgumentParser(
        description="Actualiza precios (Tipo 1) para un proveedor según su NOTAS.md."
    )
    parser.add_argument(
        "proveedor",
        nargs="?",
        default=None,
        help=(
            "Nombre de la carpeta del proveedor en proveedores/. "
            "Si se omite, se abre el menú guiado."
        ),
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Ejecuta los cambios reales (por defecto corre en modo dry-run).",
    )
    parser.add_argument(
        "--dolar",
        default=None,
        help=(
            "Opcional. Valor de dólar a escribir en DolarReferencia de las filas "
            "actualizadas (acepta coma o punto decimal, ej. 1350,50)."
        ),
    )
    args = parser.parse_args()

    if args.proveedor is None:
        # Sin argumentos: modo guiado con menú de flechas.
        modo_guiado()
        return

    # Con argumentos explícitos: comportamiento clásico por línea de comandos.
    procesar(args.proveedor, args.execute, dolar=args.dolar)


if __name__ == "__main__":
    main()
