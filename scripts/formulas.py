#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
formulas.py
------------
Módulo compartido que verifica (y, si se pide, repara) las fórmulas de las
columnas PrecioContado y PrecioSinIVA en todas las hojas de destino.

Regla de negocio:
    PrecioContado = PrecioConIVA - 15 %
        ='<Hoja>'!$H6-('<Hoja>'!$H6*15/100)
    PrecioSinIVA  = PrecioConIVA / 1,21
        ='<Hoja>'!$H6/1.21
    (la letra de columna sale siempre del encabezado PrecioConIVA, nunca se
    asume la H).

Además detecta PrecioConIVA con más de 2 decimales (valor fijo) y propone
redondearlo, porque la regla 4 del CLAUDE.md pide 2 decimales máximo.

Qué NO toca:
  - Filas sin PrecioConIVA o con PrecioConIVA en 0: no hay precio del cual
    calcular. Las que tienen 0 pero sí un contado/sin IVA cargado se informan
    aparte, para revisarlas a mano.
  - Hojas sin los tres encabezados, y las hojas excluidas siempre.

Uso desde otros scripts:

    import formulas
    cambios, revisar = formulas.revisar(wb)                  # todo el libro
    cambios, revisar = formulas.revisar(wb, filas={"Accesorios": {507}},
                                        precios_nuevos={("Accesorios", 507): 1234.5},
                                        redondear_coniva=False)
    formulas.aplicar(wb, cambios)
"""

import re
from collections import Counter, OrderedDict

from openpyxl.utils import get_column_letter

COL_CONIVA = "PrecioConIVA"
COL_CONTADO = "PrecioContado"
COL_SINIVA = "PrecioSinIVA"
COL_NOMBRE = "NombreProducto"
COL_MARCA = "Marca"

DESCUENTO_CONTADO_PCT = 15
DIVISOR_IVA = "1.21"

HOJAS_EXCLUIDAS = {"Marcas Actualización Dinámica", "Hoja1", "RegistroCambios",
                   "NoActualizados", "Venzo Sin Actualizar"}

MOTIVO_VALOR_FIJO = "valor fijo en lugar de fórmula"
MOTIVO_VACIA = "celda vacía"
MOTIVO_DISTINTA = "fórmula distinta a la esperada"
MOTIVO_REDONDEO = "más de 2 decimales"


# ---------------------------------------------------------------------------
# Construcción de fórmulas
# ---------------------------------------------------------------------------
def _ref_hoja(nombre):
    """Nombre de hoja tal como lo escribe Excel en una fórmula."""
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", nombre):
        return nombre
    return "'" + nombre.replace("'", "''") + "'"


def formula_contado(hoja, letra, fila):
    ref = f"{_ref_hoja(hoja)}!${letra}{fila}"
    return f"={ref}-({ref}*{DESCUENTO_CONTADO_PCT}/100)"


def formula_siniva(hoja, letra, fila):
    return f"={_ref_hoja(hoja)}!${letra}{fila}/{DIVISOR_IVA}"


def _normalizar_formula(valor):
    return re.sub(r"\s+", "", str(valor)).casefold()


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _buscar_columna(ws, nombre, max_filas=3):
    objetivo = nombre.strip().casefold()
    for fila in ws.iter_rows(min_row=1, max_row=max_filas):
        for celda in fila:
            if celda.value is not None and str(celda.value).strip().casefold() == objetivo:
                return celda.column, celda.row
    return None, None


def _es_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _vacia(v):
    return v is None or (isinstance(v, str) and not v.strip())


def _es_formula(v):
    return isinstance(v, str) and v.startswith("=")


# ---------------------------------------------------------------------------
# Revisión
# ---------------------------------------------------------------------------
def revisar(wb, filas=None, precios_nuevos=None, redondear_coniva=True):
    """Recorre las hojas de destino y devuelve (cambios, para_revisar).

    filas           : dict {hoja: set(filas)} para limitar la revisión a esas
                      filas (lo usan los scripts Tipo 1 y 2). None = todo.
    precios_nuevos  : dict {(hoja, fila): precio} con el PrecioConIVA que va a
                      quedar después de la corrida (para el dry-run).
    redondear_coniva: si True, propone redondear PrecioConIVA a 2 decimales.

    Cada cambio es un dict con: hoja, fila, columna, col, nombre_producto,
    marca, valor_anterior, valor_nuevo, motivo.
    """
    precios_nuevos = precios_nuevos or {}
    cambios = []
    para_revisar = []

    for ws in wb.worksheets:
        if ws.title in HOJAS_EXCLUIDAS:
            continue
        if filas is not None and ws.title not in filas:
            continue

        c_iva, fila_enc = _buscar_columna(ws, COL_CONIVA)
        c_contado, _ = _buscar_columna(ws, COL_CONTADO)
        c_siniva, _ = _buscar_columna(ws, COL_SINIVA)
        if c_iva is None or c_contado is None or c_siniva is None:
            continue
        c_nombre, _ = _buscar_columna(ws, COL_NOMBRE)
        c_marca, _ = _buscar_columna(ws, COL_MARCA)
        letra = get_column_letter(c_iva)

        if filas is not None:
            rango = sorted(f for f in filas[ws.title] if f > fila_enc)
        else:
            rango = range(fila_enc + 1, ws.max_row + 1)

        for r in rango:
            precio = precios_nuevos.get((ws.title, r), ws.cell(r, c_iva).value)
            nombre = ws.cell(r, c_nombre).value if c_nombre else None
            marca = ws.cell(r, c_marca).value if c_marca else None
            base = {"hoja": ws.title, "fila": r, "nombre_producto": nombre, "marca": marca}

            if _vacia(precio):
                continue
            if _es_numero(precio) and precio == 0:
                v_contado = ws.cell(r, c_contado).value
                v_siniva = ws.cell(r, c_siniva).value
                if (_es_numero(v_contado) and v_contado) or (_es_numero(v_siniva) and v_siniva):
                    para_revisar.append({**base, "contado": v_contado, "siniva": v_siniva})
                continue

            for nombre_col, col, esperado in (
                (COL_CONTADO, c_contado, formula_contado(ws.title, letra, r)),
                (COL_SINIVA, c_siniva, formula_siniva(ws.title, letra, r)),
            ):
                actual = ws.cell(r, col).value
                if _es_formula(actual) and _normalizar_formula(actual) == _normalizar_formula(esperado):
                    continue
                if _vacia(actual):
                    motivo = MOTIVO_VACIA
                elif _es_formula(actual):
                    motivo = MOTIVO_DISTINTA
                else:
                    motivo = MOTIVO_VALOR_FIJO
                cambios.append({**base, "columna": nombre_col, "col": col,
                                "valor_anterior": actual, "valor_nuevo": esperado,
                                "motivo": motivo})

            if redondear_coniva and _es_numero(precio) and round(precio, 2) != precio:
                cambios.append({**base, "columna": COL_CONIVA, "col": c_iva,
                                "valor_anterior": precio, "valor_nuevo": round(precio, 2),
                                "motivo": MOTIVO_REDONDEO})

    return cambios, para_revisar


def aplicar(wb, cambios):
    """Escribe los cambios en memoria (el guardado lo hace el script)."""
    for c in cambios:
        wb[c["hoja"]].cell(row=c["fila"], column=c["col"]).value = c["valor_nuevo"]


def registrar(corrida, cambios):
    """Agrega una fila de RegistroCambios por cada celda reparada."""
    for c in cambios:
        corrida.agregar(hoja=c["hoja"], fila=c["fila"],
                        nombre_producto=c["nombre_producto"], marca=c["marca"],
                        columna=c["columna"], valor_anterior=c["valor_anterior"],
                        valor_nuevo=c["valor_nuevo"])


# ---------------------------------------------------------------------------
# Reporte
# ---------------------------------------------------------------------------
def imprimir_resumen(cambios, para_revisar, titulo="Verificación de fórmulas", ejemplos=3):
    print(f"\n--- {titulo} ---")
    if not cambios:
        print("  Todas las fórmulas de PrecioContado y PrecioSinIVA están correctas.")
    else:
        por_hoja = OrderedDict()
        for c in cambios:
            por_hoja.setdefault(c["hoja"], Counter())[(c["columna"], c["motivo"])] += 1
        print(f"  Celdas a corregir: {len(cambios)}")
        for hoja, cont in por_hoja.items():
            print(f"  - {hoja}:")
            for (columna, motivo), n in sorted(cont.items()):
                print(f"      {columna:<14} {motivo}: {n}")
        print("  Ejemplos:")
        for c in cambios[:ejemplos]:
            print(f"    [{c['hoja']}] fila {c['fila']} {c['columna']}: "
                  f"{c['valor_anterior']!r} -> {c['valor_nuevo']!r}")
    if para_revisar:
        print(f"\n  AVISO: {len(para_revisar)} filas tienen PrecioConIVA en 0 pero "
              "un contado o sin IVA cargado. No se tocan; revisalas a mano:")
        for p in para_revisar[:20]:
            print(f"    [{p['hoja']}] fila {p['fila']} {p['nombre_producto']!r}: "
                  f"Contado={p['contado']!r} SinIVA={p['siniva']!r}")
        if len(para_revisar) > 20:
            print(f"    ... y {len(para_revisar) - 20} más.")
