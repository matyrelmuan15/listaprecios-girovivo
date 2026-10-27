"""
dolar.py — Regla única de validación del valor de dólar ingresado por el usuario.

Este módulo es la ÚNICA fuente de la regla del dólar en el proyecto. La usan
actualizar_precios.py (--dolar), asignar_dolar.py y actualizar_precio_dolar.py
(--precio-actualizado y modo guiado) e interfaz.py. Ningún otro script debe
redefinirla: si la regla cambia, se cambia acá.

Regla:
  - Se quitan los espacios al principio y al final (strip).
  - La coma pasa a punto decimal y se convierte con float().
  - El punto es siempre decimal (no hay separador de miles: '1.350' es 1.35).
  - No se redondea: el valor se usa tal cual.
  - Tiene que ser un número finito (se rechazan nan e inf) y mayor a cero.

No confundir con la lectura de celdas del Excel (a_float y similares), que es
otra cosa y no pasa por acá.
"""

import math


def validar_dolar(texto):
    """Devuelve (valor, error). Si es válido, error es None y valor es el
    float sin redondear; si no, valor es None y error es el mensaje para el
    usuario."""
    texto = "" if texto is None else str(texto).strip()
    if not texto:
        return None, "Ingresá el valor de dólar."
    try:
        valor = float(texto.replace(",", "."))
    except ValueError:
        valor = None
    if valor is None or not math.isfinite(valor):
        return None, "No es un número válido (ej. 1515 o 1515,50)."
    if valor <= 0:
        return None, "El valor de dólar tiene que ser mayor a cero."
    return valor, None


def parse_dolar(texto) -> float:
    """Igual que validar_dolar, pero devuelve el float directamente y levanta
    ValueError con el mensaje de error si el valor es inválido."""
    valor, error = validar_dolar(texto)
    if error is not None:
        raise ValueError(error)
    return valor
