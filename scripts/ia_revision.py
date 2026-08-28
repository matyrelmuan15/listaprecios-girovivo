#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ia_revision.py
---------------
Módulo compartido por actualizar_precios.py y asignar_dolar.py para hacer
una revisión automática de posibles anomalías en los cambios de precio de
una corrida (dry-run o ejecución), sin necesidad de chequear a mano cada
una de las cientos de filas modificadas.

Cómo funciona (2 pasos, para no depender de la IA para todo y no gastar
de más):

  1) Detección estadística en Python (gratis, instantánea): se calcula el
     % de variación de cada precio respecto a su valor anterior y se marcan
     como "sospechosos" los que:
       - tienen una variación absoluta mayor al umbral configurado, o
       - el precio nuevo queda en cero o negativo.

  2) Solo si hay ítems marcados, se arma un resumen compacto (estadísticas
     generales + la lista de ítems marcados, no todas las filas) y se le
     pide a Claude, vía la API de Anthropic, una lectura en lenguaje simple:
     si algo parece un error de carga de datos versus una suba/baja grande
     pero razonable.

Requisitos:
  - Variable de entorno ANTHROPIC_API_KEY configurada con una API key válida
    de https://console.anthropic.com. Si no está configurada, o si falla la
    llamada (sin internet, error de la API, etc.), la función devuelve un
    aviso y el script sigue funcionando igual: la revisión por IA es un
    complemento, nunca un requisito para poder usar los scripts.
  - Conexión a internet hacia api.anthropic.com.

Costo: cada corrida que tenga ítems marcados hace UNA sola llamada a la API
(no una por fila), usando créditos de esa cuenta de Anthropic (independiente
de una suscripción a Claude.ai o Claude Code).
"""

import json
import os
import urllib.error
import urllib.request

MODELO_IA = "claude-sonnet-5"
MAX_ITEMS_A_LA_IA = 40          # tope de ítems marcados que se le mandan a la IA
UMBRAL_VARIACION_PCT = 40.0     # % de variación a partir del cual se marca un cambio
TIMEOUT_SEGUNDOS = 30

# Marcadores que delimitan, en una única línea de stdout, el bloque JSON con
# los ítems marcados en formato estructurado. La interfaz gráfica (interfaz.py)
# busca esta línea para armar una tabla y la filtra del cuadro de texto de
# salida; el resto de la salida (formatear_items_marcados, análisis de la IA,
# etc.) sigue imprimiéndose igual que antes.
MARCADOR_ITEMS_JSON_INICIO = "##ITEMS_MARCADOS_JSON##"
MARCADOR_ITEMS_JSON_FIN = "##FIN_ITEMS_MARCADOS_JSON##"


def _pct_variacion(anterior, nuevo):
    try:
        anterior_f = float(anterior)
        nuevo_f = float(nuevo)
    except (TypeError, ValueError):
        return None
    if anterior_f == 0:
        return None
    return (nuevo_f - anterior_f) / anterior_f * 100


def _detectar_sospechosos(cambios, umbral_pct=UMBRAL_VARIACION_PCT):
    """cambios: lista de dicts con al menos 'valor_anterior' y 'valor_nuevo'.
    Devuelve (lista_marcados, estadisticas_generales)."""
    marcados = []
    variaciones = []

    for c in cambios:
        pct = _pct_variacion(c.get("valor_anterior"), c.get("valor_nuevo"))
        if pct is not None:
            variaciones.append(pct)

        motivos = []
        try:
            nuevo_f = float(c.get("valor_nuevo"))
            if nuevo_f <= 0:
                motivos.append("el precio nuevo queda en cero o negativo")
        except (TypeError, ValueError):
            motivos.append("el precio nuevo no es un número reconocible")

        if pct is not None and abs(pct) >= umbral_pct:
            motivos.append(f"variación de {pct:+.1f}% respecto al valor anterior")

        if motivos:
            marcado = dict(c)
            marcado["_motivo"] = "; ".join(motivos)
            marcado["_pct_variacion"] = pct
            marcados.append(marcado)

    estadisticas = {
        "total_cambios_evaluados": len(cambios),
        "total_marcados_como_atipicos": len(marcados),
    }
    if variaciones:
        variaciones_ordenadas = sorted(variaciones)
        n = len(variaciones_ordenadas)
        estadisticas["variacion_pct_promedio"] = round(sum(variaciones_ordenadas) / n, 1)
        estadisticas["variacion_pct_mediana"] = round(variaciones_ordenadas[n // 2], 1)
        estadisticas["variacion_pct_minima"] = round(variaciones_ordenadas[0], 1)
        estadisticas["variacion_pct_maxima"] = round(variaciones_ordenadas[-1], 1)

    return marcados, estadisticas


def _armar_prompt(marcados, estadisticas, contexto_extra):
    items_a_enviar = marcados[:MAX_ITEMS_A_LA_IA]
    lineas = []
    for c in items_a_enviar:
        producto = c.get("producto") or "(sin nombre registrado)"
        lineas.append(
            f"- Hoja: {c.get('hoja')} | Código/Fila: {c.get('codigo')} | "
            f"Producto: {producto} | Anterior: {c.get('valor_anterior')} | "
            f"Nuevo: {c.get('valor_nuevo')} | Motivo detectado: {c.get('_motivo')}"
        )

    nota_recorte = ""
    if len(marcados) > MAX_ITEMS_A_LA_IA:
        nota_recorte = (
            f"\n(Se muestran los primeros {MAX_ITEMS_A_LA_IA} de "
            f"{len(marcados)} ítems marcados en total.)"
        )

    return f"""Contexto de la corrida: {contexto_extra}

Estadísticas generales de todos los cambios de esta corrida (no solo los marcados):
{json.dumps(estadisticas, ensure_ascii=False, indent=2)}

Ítems marcados automáticamente como estadísticamente atípicos (variación de
precio grande o valor inválido) dentro de una lista de precios de bicicletas
y equipamiento ciclístico:
{chr(10).join(lineas)}{nota_recorte}

Analizá esta lista y respondé en español rioplatense, formal, en no más de
5 líneas: ¿ves algo que parezca un error de carga de datos (por ejemplo, un
cero de más o de menos, un valor que no tiene sentido para el rubro) versus
algo que simplemente es una suba o baja grande pero razonable? Priorizá los
casos más graves primero, agrupando si varios comparten el mismo problema.
Si todo parece razonable, decilo directamente sin inventar problemas."""


def _llamar_api(api_key, prompt_texto):
    cuerpo = json.dumps({
        "model": MODELO_IA,
        "max_tokens": 700,
        "messages": [{"role": "user", "content": prompt_texto}],
    }).encode("utf-8")

    solicitud = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=cuerpo,
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        method="POST",
    )

    with urllib.request.urlopen(solicitud, timeout=TIMEOUT_SEGUNDOS) as respuesta:
        datos = json.loads(respuesta.read().decode("utf-8"))

    texto = "".join(
        bloque.get("text", "")
        for bloque in datos.get("content", [])
        if bloque.get("type") == "text"
    )
    return texto.strip() or "(La IA no devolvió texto en la respuesta.)"


def revisar_cambios_con_ia(cambios, contexto_extra=""):
    """Punto de entrada principal.

    cambios: lista de dicts, cada uno con al menos:
        - "hoja": nombre de la hoja
        - "codigo": código o identificador de fila
        - "producto": nombre del producto (puede ser None)
        - "valor_anterior": PrecioConIVA antes del cambio
        - "valor_nuevo": PrecioConIVA después del cambio

    Devuelve una tupla (analisis, aviso):
        - analisis: texto con la lectura de la IA (o el mensaje "sin
          observaciones" si no se marcó nada), o None si no se pudo generar.
        - aviso: texto explicando por qué se salteó la revisión (sin API key,
          error de red, etc.), o None si todo salió bien.

    Esta función NUNCA lanza una excepción hacia el llamador: cualquier
    problema se traduce en un 'aviso' para que el script principal siga
    funcionando con normalidad.
    """
    if not cambios:
        return "Sin cambios para revisar.", None

    try:
        marcados, estadisticas = _detectar_sospechosos(cambios)
    except Exception as e:  # nunca debe frenar la corrida principal
        return None, f"No se pudo completar la detección de anomalías ({e})."

    if not marcados:
        return (
            "Sin observaciones: no se detectaron cambios de precio "
            "estadísticamente atípicos en esta corrida.",
            None,
        )

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None, (
            f"Se detectaron {len(marcados)} cambio(s) atípico(s), pero no se "
            "configuró la variable de entorno ANTHROPIC_API_KEY: se salteó la "
            "lectura por IA. Revisá manualmente los ítems marcados abajo."
        )

    try:
        prompt_texto = _armar_prompt(marcados, estadisticas, contexto_extra)
        analisis = _llamar_api(api_key, prompt_texto)
        return analisis, None
    except urllib.error.HTTPError as e:
        return None, (
            f"La API de Anthropic devolvió un error ({e.code}). Se salteó la "
            "lectura por IA; revisá manualmente los ítems marcados abajo."
        )
    except Exception as e:
        return None, (
            f"No se pudo contactar a la API de Anthropic ({e}). Se salteó la "
            "lectura por IA; revisá manualmente los ítems marcados abajo."
        )


def formatear_items_marcados(cambios, umbral_pct=UMBRAL_VARIACION_PCT):
    """Devuelve el texto formateado de los ítems marcados, para imprimir
    siempre (haya o no respuesta de la IA), así el usuario ve la lista
    concreta de filas a revisar."""
    marcados, _ = _detectar_sospechosos(cambios, umbral_pct)
    if not marcados:
        return ""
    lineas = ["\nÍtems marcados para revisión manual (variación grande o valor inválido):"]
    for c in marcados:
        producto = c.get("producto") or "(sin nombre registrado)"
        lineas.append(
            f"  - [{c.get('hoja')}] {c.get('codigo')} — {producto}: "
            f"{c.get('valor_anterior')} -> {c.get('valor_nuevo')} "
            f"({c.get('_motivo')})"
        )
    return "\n".join(lineas)


def obtener_items_marcados(cambios, umbral_pct=UMBRAL_VARIACION_PCT):
    """Devuelve los mismos ítems que formatear_items_marcados(), pero como
    lista de diccionarios estructurados (hoja, codigo, producto,
    valor_anterior, valor_nuevo, motivo) en vez de texto, para que quien
    llame los pueda mostrar como tabla."""
    marcados, _ = _detectar_sospechosos(cambios, umbral_pct)
    return [
        {
            "hoja": c.get("hoja"),
            "codigo": c.get("codigo"),
            "producto": c.get("producto"),
            "valor_anterior": c.get("valor_anterior"),
            "valor_nuevo": c.get("valor_nuevo"),
            "motivo": c.get("_motivo"),
        }
        for c in marcados
    ]


def imprimir_items_marcados_json(cambios, umbral_pct=UMBRAL_VARIACION_PCT):
    """Imprime SIEMPRE (incluso con lista vacía) una línea con los ítems
    marcados en JSON, delimitada por MARCADOR_ITEMS_JSON_INICIO/FIN, para que
    la interfaz gráfica arme su tabla y pueda distinguir 'sin observaciones'
    de 'todavía no corrió'. Se imprime en una sola línea (JSON compacto, sin
    saltos internos) para que la interfaz la pueda leer con un simple
    line-by-line de stdout."""
    items = obtener_items_marcados(cambios, umbral_pct)
    print(
        MARCADOR_ITEMS_JSON_INICIO
        + json.dumps(items, ensure_ascii=False, default=str)
        + MARCADOR_ITEMS_JSON_FIN
    )
