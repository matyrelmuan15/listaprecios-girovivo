#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
registro.py
------------
Módulo compartido para la hoja de trazabilidad "RegistroCambios".

Los tres scripts del proyecto (Tipo 1 precios, Tipo 2 dólar, Tipo 3 fecha)
escriben en la MISMA hoja. Antes, cada uno definía sus propios encabezados,
con anchos distintos, y como la hoja ya existía ninguno los reescribía: los
datos terminaban corridos bajo títulos que no correspondían.

Este módulo es la única fuente de verdad del formato. Ningún script debe
volver a definir encabezados de registro por su cuenta.

Formato único (10 columnas, en este orden exacto):

    Fecha            fecha y hora de la corrida, texto dd/mm/aaaa hh:mm
    TipoOperacion    "Precios" (Tipo 1) | "Dolar" (Tipo 2) | "Fecha" (Tipo 3)
    Hoja             nombre de la hoja destino
    Fila             número de fila modificada
    NombreProducto   valor de la columna NombreProducto de esa fila
    Marca            valor de la columna Marca de esa fila
    Columna          columna modificada (PrecioConIVA / DolarReferencia /
                     FechaActualizacion)
    ValorAnterior    valor previo de esa celda puntual
    ValorNuevo       valor nuevo de esa celda puntual
    Detalle          contexto de la corrida (proveedor y fórmula en Tipo 1;
                     modo, marcas y rango de fechas en Tipo 2 y 3)

Granularidad (regla 8 del CLAUDE.md): una fila de registro por cada CELDA
modificada. Tipo 1 -> 2 filas por producto. Tipo 2 -> 3. Tipo 3 -> 1.

Uso típico desde un script:

    import registro

    wb = openpyxl.load_workbook(ARCHIVO_PATH)   # sin data_only=True
    registro.validar_encabezados(wb)            # fail-fast, ANTES del backup

    corrida = registro.RegistroCorrida(registro.TIPO_PRECIOS, detalle)
    corrida.agregar(hoja=..., fila=..., nombre_producto=..., marca=...,
                    columna="PrecioConIVA", valor_anterior=..., valor_nuevo=...)
    corrida.verificar_conteo(celdas_modificadas)
    corrida.escribir(wb)
"""

import datetime

HOJA_REGISTRO = "RegistroCambios"

REGISTRO_HEADERS = [
    "Fecha",
    "TipoOperacion",
    "Hoja",
    "Fila",
    "NombreProducto",
    "Marca",
    "Columna",
    "ValorAnterior",
    "ValorNuevo",
    "Detalle",
]

# Valores admitidos para la columna TipoOperacion.
TIPO_PRECIOS = "Precios"
TIPO_DOLAR = "Dolar"
TIPO_FECHA = "Fecha"


class RegistroError(Exception):
    """Error base de la trazabilidad. Siempre debe abortar sin guardar."""


class RegistroFormatoError(RegistroError):
    """La hoja RegistroCambios existe pero no tiene el formato único."""


class RegistroConteoError(RegistroError):
    """Las filas de registro no coinciden con las celdas modificadas."""


# ---------------------------------------------------------------------------
# Utilidades internas
# ---------------------------------------------------------------------------
def _normalizar_encabezado(valor):
    """Normaliza un encabezado para compararlo sin importar mayúsculas ni
    espacios al principio/final."""
    if valor is None:
        return ""
    return str(valor).strip().casefold()


def _valor_para_celda(valor):
    """Convierte fechas a texto dd/mm/aaaa (regla 4 del CLAUDE.md).

    El libro real mezcla fechas como texto con algún datetime suelto; sin
    esta conversión el registro guardaría tipos distintos para la misma
    columna."""
    if isinstance(valor, datetime.datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, datetime.date):
        return valor.strftime("%d/%m/%Y")
    return valor


def hoja_vacia(ws):
    """True si la hoja no tiene ni siquiera la fila de encabezados."""
    if ws.max_row > 1:
        return False
    return all(celda.value is None for celda in ws[1])


def timestamp_corrida():
    """Fecha y hora de la corrida, en el formato de la columna Fecha."""
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M")


def escribir_encabezados(ws):
    for i, encabezado in enumerate(REGISTRO_HEADERS, start=1):
        ws.cell(row=1, column=i).value = encabezado


# ---------------------------------------------------------------------------
# Validación fail-fast
# ---------------------------------------------------------------------------
def validar_encabezados(wb):
    """Verifica que la hoja RegistroCambios tenga exactamente los 10
    encabezados del formato único.

    Se llama apenas cargado el libro y ANTES del backup, para no dejar un
    backup inútil ni apilar filas desalineadas bajo títulos que no
    corresponden. Si la hoja no existe, o existe pero está vacía, no hay
    nada que validar: se creará con el formato correcto al escribir.

    Levanta RegistroFormatoError con un mensaje accionable si no coincide.
    """
    if HOJA_REGISTRO not in wb.sheetnames:
        return

    ws = wb[HOJA_REGISTRO]
    if hoja_vacia(ws):
        return

    cantidad = len(REGISTRO_HEADERS)
    actuales = [ws.cell(row=1, column=i).value for i in range(1, cantidad + 1)]
    coincide = [_normalizar_encabezado(v) for v in actuales] == [
        _normalizar_encabezado(h) for h in REGISTRO_HEADERS
    ]

    # Columnas de más a la derecha de las 10 esperadas: también invalidan.
    sobrantes = []
    for i in range(cantidad + 1, ws.max_column + 1):
        valor = ws.cell(row=1, column=i).value
        if valor is not None and str(valor).strip():
            sobrantes.append(str(valor).strip())

    if coincide and not sobrantes:
        return

    encontrados = [("" if v is None else str(v)) for v in actuales] + sobrantes
    raise RegistroFormatoError(
        f"La hoja '{HOJA_REGISTRO}' no tiene el formato único esperado.\n"
        f"  Esperado ({cantidad} columnas): {REGISTRO_HEADERS}\n"
        f"  Encontrado:                     {encontrados}\n"
        "No se escribe nada para no apilar filas desalineadas.\n"
        "Corré 'python migrar_registro.py' para ver el dry-run de la "
        "migración, y luego 'python migrar_registro.py --execute' para "
        "rehacer la hoja con el formato único."
    )


def obtener_hoja_registro(wb):
    """Devuelve la hoja de registro lista para escribir.

    La crea con encabezados si no existe o si está vacía. Si existe con
    datos, valida el formato antes de devolverla.
    """
    if HOJA_REGISTRO not in wb.sheetnames:
        ws = wb.create_sheet(HOJA_REGISTRO)
        escribir_encabezados(ws)
        return ws

    ws = wb[HOJA_REGISTRO]
    if hoja_vacia(ws):
        escribir_encabezados(ws)
        return ws

    validar_encabezados(wb)
    return ws


# ---------------------------------------------------------------------------
# Construcción y escritura de filas
# ---------------------------------------------------------------------------
def nueva_fila(fecha, tipo_operacion, hoja, fila, nombre_producto, marca,
               columna, valor_anterior, valor_nuevo, detalle):
    """Arma un dict con las 10 claves del formato único."""
    return {
        "Fecha": fecha,
        "TipoOperacion": tipo_operacion,
        "Hoja": hoja,
        "Fila": fila,
        "NombreProducto": nombre_producto,
        "Marca": marca,
        "Columna": columna,
        "ValorAnterior": valor_anterior,
        "ValorNuevo": valor_nuevo,
        "Detalle": detalle,
    }


def escribir_filas(ws_reg, filas):
    """Agrega las filas al final de la hoja, siempre en el orden de
    REGISTRO_HEADERS."""
    for fila in filas:
        ws_reg.append([_valor_para_celda(fila[h]) for h in REGISTRO_HEADERS])


def verificar_conteo(filas, celdas_modificadas):
    """Regla 8 del CLAUDE.md: una fila de registro por celda modificada.

    Levanta RegistroConteoError si no coinciden, para abortar sin guardar.
    """
    if len(filas) != celdas_modificadas:
        raise RegistroConteoError(
            "ERROR DE INTEGRIDAD: la cantidad de filas de registro "
            f"({len(filas)}) no coincide con la cantidad de celdas "
            f"modificadas ({celdas_modificadas}). No se guarda el archivo."
        )


def imprimir_previsualizacion(corrida, limite=5):
    """Muestra en el dry-run cómo quedarían las filas de registro.

    Sirve para verificar a simple vista que las 10 columnas se arman bien
    antes de escribir nada en el libro.
    """
    print(f"\n--- Vista previa de '{HOJA_REGISTRO}' ---")
    print(f"  Filas de registro que se agregarían: {len(corrida)}")
    if not corrida.filas:
        print("  (ninguna)")
        return

    print(f"  Columnas ({len(REGISTRO_HEADERS)}): {' | '.join(REGISTRO_HEADERS)}")
    for fila in corrida.filas[:limite]:
        valores = [_valor_para_celda(fila[h]) for h in REGISTRO_HEADERS]
        print("    - " + " | ".join(
            "(vacío)" if v is None else str(v) for v in valores
        ))
    if len(corrida.filas) > limite:
        print(f"    ... y {len(corrida.filas) - limite} filas más.")


class RegistroCorrida:
    """Acumula las filas de registro de una corrida y las escribe al final.

    Fija de una sola vez la Fecha, el TipoOperacion y el Detalle, que son
    idénticos para todas las filas de la misma corrida.
    """

    def __init__(self, tipo_operacion, detalle, fecha=None):
        self.tipo_operacion = tipo_operacion
        self.detalle = detalle
        self.fecha = fecha or timestamp_corrida()
        self.filas = []

    def __len__(self):
        return len(self.filas)

    def agregar(self, hoja, fila, nombre_producto, marca, columna,
                valor_anterior, valor_nuevo):
        self.filas.append(
            nueva_fila(
                fecha=self.fecha,
                tipo_operacion=self.tipo_operacion,
                hoja=hoja,
                fila=fila,
                nombre_producto=nombre_producto,
                marca=marca,
                columna=columna,
                valor_anterior=valor_anterior,
                valor_nuevo=valor_nuevo,
                detalle=self.detalle,
            )
        )

    def verificar_conteo(self, celdas_modificadas):
        verificar_conteo(self.filas, celdas_modificadas)

    def escribir(self, wb):
        """Escribe las filas acumuladas. Revalida el formato por las dudas."""
        ws_reg = obtener_hoja_registro(wb)
        escribir_filas(ws_reg, self.filas)
        return ws_reg
