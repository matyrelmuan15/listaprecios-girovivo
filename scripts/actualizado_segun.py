#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
actualizado_segun.py
---------------------
Módulo compartido para la columna "ActualizadoSegun" y la hoja
"NoActualizados".

ActualizadoSegun guarda la ÚLTIMA forma en que se actualizó el precio de cada
fila. Valores posibles (lista desplegable en el Excel):

    Lista proveedor   la escribe actualizar_precios.py (Tipo 1)
    Dólar             la escribe actualizar_precio_dolar.py (Tipo 4)
    Manual            la carga el usuario a mano

asignar_dolar.py (Tipo 2) y actualizar_fecha.py (Tipo 3) no tocan la columna,
porque no cambian el precio.

La columna se crea UNA sola vez con crear_columna_actualizado.py. Si una hoja
no la tiene, los scripts igual actualizan el precio y avisan.

NoActualizados se regenera en cada ejecución real de Tipo 1: lista las filas
del Excel de las marcas del proveedor (parámetro "marcas:" de su NOTAS.md) que
no recibieron precio en esa corrida. Es una hoja de informe, como la columna E
de Hoja1: no se registra en RegistroCambios. Sus encabezados van en la fila 4
para que ningún script la confunda con una hoja de destino, y además todos la
excluyen por nombre.
"""

from copy import copy

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

COL_ACTUALIZADO = "ActualizadoSegun"
VALOR_LISTA = "Lista proveedor"
VALOR_DOLAR = "Dólar"
VALOR_MANUAL = "Manual"
VALORES = (VALOR_MANUAL, VALOR_DOLAR, VALOR_LISTA)

HOJA_NO_ACTUALIZADOS = "NoActualizados"
# Hoja manual anterior que NoActualizados reemplaza (el usuario la borra a mano;
# mientras exista, todos los scripts la excluyen).
HOJAS_REEMPLAZADAS = ("Venzo Sin Actualizar",)

FILAS_VALIDACION = 5000  # alcance de la lista desplegable

ENCABEZADOS_NO_ACTUALIZADOS = [
    "Hoja", "Fila", "CodigoDux", "NombreProducto", "CodigoFactura", "Marca",
    "PrecioConIVA", "FechaActualizacion", "ActualizadoSegun", "Motivo",
]

MOTIVO_SIN_CODIGO = "Sin CodigoFactura en el Excel"
MOTIVO_NO_EN_LISTA = "Código no incluido en la lista del proveedor (Hoja1)"
MOTIVO_FORMATO = "Código en Hoja1 con precio de formato no reconocido"


def _norm(valor):
    return "" if valor is None else str(valor).strip().casefold()


def buscar_columna(ws):
    """Número de columna (1-based) de ActualizadoSegun en la fila 1, o None."""
    for celda in ws[1]:
        if _norm(celda.value) == _norm(COL_ACTUALIZADO):
            return celda.column
    return None


def ultima_columna_encabezado(ws):
    ultima = 0
    for celda in ws[1]:
        if celda.value is not None and str(celda.value).strip() != "":
            ultima = celda.column
    return ultima


def crear_columna(ws):
    """Agrega el encabezado ActualizadoSegun después del último encabezado,
    copiando su formato, y una lista desplegable con los 3 valores.
    Devuelve el número de columna creada."""
    ultima = ultima_columna_encabezado(ws)
    col = ultima + 1
    nueva = ws.cell(row=1, column=col)
    nueva.value = COL_ACTUALIZADO
    if ultima:
        modelo = ws.cell(row=1, column=ultima)
        if modelo.has_style:
            nueva._style = copy(modelo._style)
        ancho = ws.column_dimensions[get_column_letter(ultima)].width
        ws.column_dimensions[get_column_letter(col)].width = max(ancho or 0, 18)
    agregar_validacion(ws, col)
    return col


def agregar_validacion(ws, col):
    letra = get_column_letter(col)
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(VALORES) + '"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Valor no válido",
        error="Elegí Manual, Dólar o Lista proveedor.",
    )
    dv.add(f"{letra}2:{letra}{FILAS_VALIDACION}")
    ws.add_data_validation(dv)


def parse_marcas(texto):
    """'GIANT, LIV, CADEX' -> {'giant', 'liv', 'cadex'} (normalizado)."""
    if not texto:
        return set()
    return {_norm(m) for m in texto.split(",") if m.strip()}


def marca_coincide(valor_marca, marcas_norm):
    return _norm(valor_marca) in marcas_norm


def escribir_hoja_no_actualizados(wb, proveedor, marcas_texto, fecha, filas):
    """Borra y vuelve a crear la hoja NoActualizados con el detalle de la
    corrida. `filas` es una lista de dicts con las claves de
    ENCABEZADOS_NO_ACTUALIZADOS."""
    if HOJA_NO_ACTUALIZADOS in wb.sheetnames:
        del wb[HOJA_NO_ACTUALIZADOS]
    ws = wb.create_sheet(HOJA_NO_ACTUALIZADOS)
    ws.cell(row=1, column=1).value = f"Productos sin actualizar — proveedor {proveedor}"
    ws.cell(row=2, column=1).value = f"Fecha de la corrida: {fecha}"
    ws.cell(row=2, column=4).value = f"Marcas del proveedor: {marcas_texto}"
    ws.cell(row=3, column=1).value = f"Total: {len(filas)}"
    from openpyxl.styles import Font
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    for i, titulo in enumerate(ENCABEZADOS_NO_ACTUALIZADOS, start=1):
        c = ws.cell(row=4, column=i)
        c.value = titulo
        c.font = Font(bold=True)
    for r, fila in enumerate(filas, start=5):
        for i, titulo in enumerate(ENCABEZADOS_NO_ACTUALIZADOS, start=1):
            ws.cell(row=r, column=i).value = fila.get(titulo)
    anchos = {"A": 22, "B": 6, "C": 16, "D": 55, "E": 30, "F": 16,
              "G": 14, "H": 18, "I": 16, "J": 50}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho
    ws.freeze_panes = "A5"
    return ws
