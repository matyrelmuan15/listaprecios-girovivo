#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
interfaz.py
===========
Interfaz de escritorio (Tkinter, sin dependencias extra) para el proyecto
"Lista de Precios Giro Vivo". Pensada para uso propio, sin necesidad de
escribir comandos en PowerShell.

Cómo funciona
-------------
Esta ventana NO reimplementa la lógica de negocio: arma el comando correcto
y corre `actualizar_precios.py` / `asignar_dolar.py` como un proceso hijo
(igual que si los tipearas en la terminal), y muestra la salida en vivo en
un cuadro de texto. Así toda la lógica (fórmulas, matching de códigos,
backups, RegistroCambios, revisión por IA) sigue viviendo en un solo lugar:
los scripts. Si mañana cambiás algo ahí, la interfaz no necesita tocarse.

Estructura (para seguir afinando en Claude Code)
-------------------------------------------------
- InterfazGiroVivo: clase principal, controla la ventana y la navegación
  entre pantallas (Inicio / Actualizar precios / Asignar dólar).
- Cada pantalla es un método "construir_pantalla_*" que arma un ttk.Frame.
- ejecutar_comando_async(): corre un comando en un hilo aparte y va
  volcando cada línea de salida al cuadro de texto, sin congelar la
  ventana. Usa una cola (queue.Queue) porque Tkinter no es thread-safe:
  el hilo de fondo solo escribe en la cola, y el bucle principal
  (root.after) es el único que toca los widgets.
- listar_proveedores() y listar_marcas() se importan directamente de los
  scripts existentes (son funciones puras, no ejecutan nada raro al
  importarlas).

Requisitos: Python 3.8+ con Tkinter (viene incluido en la instalación
estándar de Python para Windows).
"""

import datetime
import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox, ttk

SCRIPT_DIR = Path(__file__).resolve().parent
PYTHON = sys.executable

# Enlaces fijos de este proyecto (ajustar acá si cambian de lugar).
URL_PROYECTO_CLAUDE = "https://claude.ai/cowork/project/019ffbd7-3313-756a-8ea9-a08cb158e185"
CARPETA_PROVEEDORES = r"C:\Users\Maty y Katy\Mi unidad\Proveedores GV"
ARCHIVO_EXCEL_PRINCIPAL = (
    r"C:\Users\Maty y Katy\Desktop\ListaPreciosGiroVivo\archivo\Lista de Precios Giro Vivo 2026.xlsx"
)
NOMBRE_HOJA_ENTRADA = "Hoja1"

ICONO_PATH = SCRIPT_DIR / "icono.ico"
CONFIG_PATH = SCRIPT_DIR / "config_interfaz.json"

sys.path.insert(0, str(SCRIPT_DIR))
from actualizar_precios import listar_proveedores  # noqa: E402
from dolar import parse_dolar, validar_dolar  # noqa: E402
from asignar_dolar import ARCHIVO_PATH as ARCHIVO_DOLAR, listar_marcas  # noqa: E402
from ia_revision import MARCADOR_ITEMS_JSON_FIN, MARCADOR_ITEMS_JSON_INICIO  # noqa: E402

PATRON_ITEMS_JSON = re.compile(
    re.escape(MARCADOR_ITEMS_JSON_INICIO) + r"(.*)" + re.escape(MARCADOR_ITEMS_JSON_FIN)
)


# ---------------------------------------------------------------------------
# Configuración local (recordar la última selección del usuario)
# ---------------------------------------------------------------------------
def cargar_config():
    """Lee scripts/config_interfaz.json. Si no existe o está corrompido,
    devuelve un diccionario vacío sin romper la app."""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError, OSError, UnicodeDecodeError):
        return {}


def guardar_config(datos):
    """Guarda la config local. Si falla (por ejemplo sin permisos de
    escritura), no interrumpe el flujo de la interfaz."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except OSError:
        pass

# ---------------------------------------------------------------------------
# Paleta (inspirada en la identidad de marca: negro + verde lima)
# ---------------------------------------------------------------------------
COLOR_FONDO = "#121212"
COLOR_TARJETA = "#1E1E1E"
COLOR_TARJETA_HOVER = "#262626"
COLOR_ACENTO = "#B0FF2A"
COLOR_ACENTO_TEXTO = "#121212"
COLOR_TEXTO = "#F2F2F2"
COLOR_TEXTO_SECUNDARIO = "#A8A8A8"
COLOR_BORDE = "#333333"
COLOR_ALERTA = "#FFB020"
COLOR_ERROR = "#FF5C5C"
FUENTE_BASE = ("Segoe UI", 10)
FUENTE_TITULO = ("Segoe UI", 16, "bold")
FUENTE_SUBTITULO = ("Segoe UI", 10)
FUENTE_BOTON_GRANDE = ("Segoe UI", 11, "bold")
FUENTE_MONO = ("Consolas", 9)


class InterfazGiroVivo:
    def __init__(self, root):
        self.root = root
        self.root.title("Lista de precios — Giro Vivo")
        self.root.geometry("900x640")
        self.root.minsize(820, 560)
        self.root.configure(bg=COLOR_FONDO)

        if ICONO_PATH.exists():
            try:
                self.root.iconbitmap(str(ICONO_PATH))
            except tk.TclError:
                pass

        self.config_interfaz = cargar_config()

        self._configurar_estilos()

        # Contenedor donde se intercambian las pantallas
        self.contenedor = tk.Frame(self.root, bg=COLOR_FONDO)
        self.contenedor.pack(fill="both", expand=True)

        self.pantalla_actual = None
        self.proceso_en_curso = False

        self.mostrar_inicio()

    # ------------------------------------------------------------------
    # Estilos ttk
    # ------------------------------------------------------------------
    def _configurar_estilos(self):
        estilo = ttk.Style()
        estilo.theme_use("clam")

        estilo.configure("TFrame", background=COLOR_FONDO)
        estilo.configure("Tarjeta.TFrame", background=COLOR_TARJETA)

        estilo.configure(
            "TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE_BASE
        )
        estilo.configure(
            "Titulo.TLabel", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE_TITULO
        )
        estilo.configure(
            "Subtitulo.TLabel",
            background=COLOR_FONDO,
            foreground=COLOR_TEXTO_SECUNDARIO,
            font=FUENTE_SUBTITULO,
        )
        estilo.configure(
            "Tarjeta.TLabel", background=COLOR_TARJETA, foreground=COLOR_TEXTO, font=FUENTE_BASE
        )
        estilo.configure(
            "TarjetaSecundario.TLabel",
            background=COLOR_TARJETA,
            foreground=COLOR_TEXTO_SECUNDARIO,
            font=FUENTE_BASE,
        )

        estilo.configure(
            "Acento.TButton",
            background=COLOR_ACENTO,
            foreground=COLOR_ACENTO_TEXTO,
            font=FUENTE_BOTON_GRANDE,
            borderwidth=0,
            focuscolor=COLOR_ACENTO,
            padding=10,
        )
        estilo.map("Acento.TButton", background=[("active", "#9EE626"), ("disabled", "#5A6B33")])

        estilo.configure(
            "Normal.TButton",
            background=COLOR_TARJETA,
            foreground=COLOR_TEXTO,
            font=FUENTE_BASE,
            borderwidth=1,
            padding=8,
        )
        estilo.map("Normal.TButton", background=[("active", COLOR_TARJETA_HOVER)])

        estilo.configure(
            "TCombobox",
            fieldbackground=COLOR_TARJETA,
            background=COLOR_TARJETA,
            foreground=COLOR_TEXTO,
            arrowcolor=COLOR_TEXTO,
        )
        # El tema "clam" ignora fieldbackground/foreground en estado "readonly"
        # y usa sus propios colores por default, casi invisibles sobre COLOR_FONDO.
        # Hay que fijarlos explícitamente para ese estado (cerrado y con el
        # desplegable abierto).
        estilo.map(
            "TCombobox",
            fieldbackground=[("readonly", COLOR_TARJETA)],
            foreground=[("readonly", COLOR_TEXTO)],
            selectbackground=[("readonly", COLOR_TARJETA)],
            selectforeground=[("readonly", COLOR_TEXTO)],
        )
        self.root.option_add("*TCombobox*Listbox.background", COLOR_TARJETA)
        self.root.option_add("*TCombobox*Listbox.foreground", COLOR_TEXTO)
        self.root.option_add("*TCombobox*Listbox.selectBackground", COLOR_ACENTO)
        self.root.option_add("*TCombobox*Listbox.selectForeground", COLOR_ACENTO_TEXTO)
        estilo.configure("TEntry", fieldbackground=COLOR_TARJETA, foreground=COLOR_TEXTO)
        estilo.configure(
            "TRadiobutton", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE_BASE
        )
        estilo.map("TRadiobutton", background=[("active", COLOR_FONDO)])
        estilo.configure(
            "TCheckbutton", background=COLOR_FONDO, foreground=COLOR_TEXTO, font=FUENTE_BASE
        )
        estilo.map("TCheckbutton", background=[("active", COLOR_FONDO)])

    # ------------------------------------------------------------------
    # Navegación entre pantallas
    # ------------------------------------------------------------------
    def _cambiar_pantalla(self, constructor):
        if self.pantalla_actual is not None:
            self.pantalla_actual.destroy()
        self.pantalla_actual = constructor(self.contenedor)
        self.pantalla_actual.pack(fill="both", expand=True)

    def mostrar_inicio(self):
        self._cambiar_pantalla(self._construir_pantalla_inicio)

    # ------------------------------------------------------------------
    # Pantalla: inicio
    # ------------------------------------------------------------------
    def _construir_pantalla_inicio(self, padre):
        marco = ttk.Frame(padre, padding=24)

        ttk.Label(marco, text="Lista de precios Giro Vivo", style="Titulo.TLabel").pack(
            anchor="w"
        )
        ttk.Label(marco, text="Elegí qué querés hacer", style="Subtitulo.TLabel").pack(
            anchor="w", pady=(2, 20)
        )

        fila_tarjetas = ttk.Frame(marco)
        fila_tarjetas.pack(fill="x")
        fila_tarjetas.columnconfigure(0, weight=1)
        fila_tarjetas.columnconfigure(1, weight=1)

        self._tarjeta_opcion(
            fila_tarjetas,
            fila=0,
            columna=0,
            titulo="Actualizar precios",
            descripcion="Pegás la lista de un proveedor y se actualizan los precios en el Excel.",
            accion=self.mostrar_actualizar_precios,
        )
        self._tarjeta_opcion(
            fila_tarjetas,
            fila=0,
            columna=1,
            titulo="Asignar dólar",
            descripcion="Cambia solo el dólar de referencia (y la fecha). No recalcula precios.",
            accion=self.mostrar_asignar_dolar,
        )
        self._tarjeta_opcion(
            fila_tarjetas,
            fila=1,
            columna=0,
            titulo="Actualizar fecha",
            descripcion="Deja constancia de revisión: actualiza solo la fecha de una o varias marcas, sin tocar precios.",
            accion=self.mostrar_actualizar_fecha,
        )
        self._tarjeta_opcion(
            fila_tarjetas,
            fila=1,
            columna=1,
            titulo="Exportar datos",
            descripcion="Genera el archivo .xls con código y precio de cada producto, listo para subir al sistema.",
            accion=self.mostrar_exportar_datos,
        )
        self._tarjeta_opcion(
            fila_tarjetas,
            fila=2,
            columna=0,
            titulo="Actualizar precio por dólar",
            descripcion="Recalcula el precio con IVA según el nuevo dólar y actualiza el dólar de referencia.",
            accion=self.mostrar_actualizar_precio_dolar,
        )

        fila_enlaces = ttk.Frame(marco)
        fila_enlaces.pack(fill="x", pady=(24, 0))
        self._enlace(
            fila_enlaces, "Ver proyecto en Claude", self._abrir_proyecto_claude
        ).pack(side="left")
        self._enlace(
            fila_enlaces, "Abrir carpeta de proveedores", self._abrir_carpeta_proveedores
        ).pack(side="left", padx=(24, 0))
        self._enlace(
            fila_enlaces, "Abrir Excel (Hoja1)", self._abrir_excel_hoja1
        ).pack(side="left", padx=(24, 0))

        return marco

    def _enlace(self, padre, texto, accion):
        etiqueta = tk.Label(
            padre, text=texto, bg=COLOR_FONDO, fg=COLOR_ACENTO, font=("Segoe UI", 9, "underline"),
            cursor="hand2",
        )
        etiqueta.bind("<Button-1>", lambda e: accion())
        return etiqueta

    def _abrir_proyecto_claude(self):
        webbrowser.open(URL_PROYECTO_CLAUDE)

    def _abrir_carpeta_proveedores(self):
        try:
            os.startfile(CARPETA_PROVEEDORES)  # noqa: S606 — abre el Explorador de Windows
        except FileNotFoundError:
            messagebox.showerror(
                "No se encontró la carpeta",
                f"No se pudo abrir:\n{CARPETA_PROVEEDORES}\n\n"
                "Revisá que la ruta siga siendo la misma (por ejemplo, si cambió el nombre "
                "de usuario de Windows o la unidad de Google Drive).",
            )
        except Exception as e:
            messagebox.showerror("No se pudo abrir la carpeta", str(e))

    def _abrir_excel_hoja1(self):
        """Abre el libro de precios activando puntualmente 'Hoja1', usando
        automatización COM de Excel (pywin32). Import diferido: si pywin32
        no está instalado, no rompe la app entera, solo este enlace."""
        if not Path(ARCHIVO_EXCEL_PRINCIPAL).exists():
            messagebox.showerror(
                "No se encontró el archivo",
                f"No se pudo abrir:\n{ARCHIVO_EXCEL_PRINCIPAL}\n\n"
                "Revisá que la ruta siga siendo la misma (por ejemplo, si cambió el nombre "
                "de usuario de Windows o se movió el archivo).",
            )
            return

        try:
            import win32com.client
            import pywintypes
        except ImportError:
            if messagebox.askyesno(
                "Falta un componente opcional",
                "Para abrir el Excel directamente en la pestaña 'Hoja1' hace falta instalar "
                "la librería 'pywin32' (podés hacerlo con: pip install pywin32).\n\n"
                "¿Querés que abra el archivo igual, sin garantizar que quede en esa pestaña?",
            ):
                try:
                    os.startfile(ARCHIVO_EXCEL_PRINCIPAL)  # noqa: S606
                except Exception as e:
                    messagebox.showerror("No se pudo abrir el archivo", str(e))
            return

        ruta_absoluta = os.path.abspath(ARCHIVO_EXCEL_PRINCIPAL)
        try:
            try:
                excel_app = win32com.client.GetObject(Class="Excel.Application")
            except pywintypes.com_error:
                excel_app = win32com.client.Dispatch("Excel.Application")

            libro = None
            for wb in excel_app.Workbooks:
                try:
                    if os.path.normcase(wb.FullName) == os.path.normcase(ruta_absoluta):
                        libro = wb
                        break
                except Exception:
                    continue
            if libro is None:
                libro = excel_app.Workbooks.Open(ruta_absoluta)

            excel_app.Visible = True

            try:
                libro.Sheets(NOMBRE_HOJA_ENTRADA).Activate()
            except pywintypes.com_error:
                messagebox.showwarning(
                    "No se encontró la pestaña",
                    f"El libro se abrió, pero no se encontró la pestaña '{NOMBRE_HOJA_ENTRADA}' "
                    "(puede haber sido borrada o renombrada).",
                )

            try:
                excel_app.Activate()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("No se pudo abrir el Excel", str(e))

    def _tarjeta_opcion(self, padre, columna, titulo, descripcion, accion, fila=0):
        tarjeta = tk.Frame(padre, bg=COLOR_TARJETA, padx=18, pady=18)
        padx = (0, 8) if columna == 0 else (8, 0)
        pady = (0, 8) if fila == 0 else (8, 0)
        tarjeta.grid(row=fila, column=columna, sticky="nsew", padx=padx, pady=pady)

        tk.Label(
            tarjeta, text=titulo, bg=COLOR_TARJETA, fg=COLOR_TEXTO,
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")
        tk.Label(
            tarjeta, text=descripcion, bg=COLOR_TARJETA, fg=COLOR_TEXTO_SECUNDARIO,
            font=FUENTE_BASE, wraplength=320, justify="left",
        ).pack(anchor="w", pady=(6, 14))

        ttk.Button(tarjeta, text="Elegir", style="Acento.TButton", command=accion).pack(
            anchor="w"
        )

    # ------------------------------------------------------------------
    # Pantalla: actualizar precios
    # ------------------------------------------------------------------
    def mostrar_actualizar_precios(self):
        self._cambiar_pantalla(self._construir_pantalla_actualizar_precios)

    def _construir_pantalla_actualizar_precios(self, padre):
        marco = ttk.Frame(padre, padding=20)

        self._encabezado_pantalla(marco, "Actualizar precios")

        fila_proveedor = ttk.Frame(marco)
        fila_proveedor.pack(fill="x", pady=(4, 12))
        ttk.Label(fila_proveedor, text="Proveedor:").pack(side="left", padx=(0, 8))

        proveedores = listar_proveedores()
        proveedor_guardado = self.config_interfaz.get("proveedor")
        if proveedor_guardado not in proveedores:
            proveedor_guardado = proveedores[0] if proveedores else ""
        self.var_proveedor = tk.StringVar(value=proveedor_guardado)
        combo = ttk.Combobox(
            fila_proveedor, textvariable=self.var_proveedor, values=proveedores,
            state="readonly", width=30,
        )
        combo.pack(side="left")
        if not proveedores:
            ttk.Label(
                marco,
                text="No se encontró ningún proveedor con NOTAS.md en la carpeta 'proveedores/'.",
                foreground=COLOR_ERROR,
            ).pack(anchor="w", pady=(0, 8))

        # Dólar opcional. No se guarda en la configuración: cambia seguido y
        # arrancar vacío evita asignar por error un valor viejo.
        fila_dolar = ttk.Frame(marco)
        fila_dolar.pack(fill="x", pady=(0, 12))
        ttk.Label(fila_dolar, text="Valor dólar (opcional):").pack(side="left", padx=(0, 8))
        self.var_dolar_ap = tk.StringVar()
        ttk.Entry(fila_dolar, textvariable=self.var_dolar_ap, width=14).pack(side="left")
        ttk.Label(
            fila_dolar,
            text=("Si lo completás (ej. 1515 o 1515,50), se asigna a DolarReferencia "
                  "de los productos actualizados. Dejalo vacío para no tocar el dólar."),
            style="Subtitulo.TLabel", wraplength=420, justify="left",
        ).pack(side="left", padx=(10, 0))

        # Firma (proveedor + dólar) del último dry-run que terminó bien.
        # "Confirmar y aplicar" solo se habilita si la firma actual coincide.
        self.ap_dryrun_firma = None
        self.var_proveedor.trace_add("write", self._ap_al_cambiar_parametro)
        self.var_dolar_ap.trace_add("write", self._ap_al_cambiar_parametro)

        fila_botones = ttk.Frame(marco)
        fila_botones.pack(fill="x", pady=(0, 10))
        self.btn_ap_preview = ttk.Button(
            fila_botones, text="Vista previa (dry-run)", style="Normal.TButton",
            command=self._ap_correr_preview,
        )
        self.btn_ap_preview.pack(side="left", padx=(0, 8))
        self.btn_ap_confirmar = ttk.Button(
            fila_botones, text="Confirmar y aplicar", style="Acento.TButton",
            command=self._ap_correr_confirmar, state="disabled",
        )
        self.btn_ap_confirmar.pack(side="left")

        self.texto_ap = self._crear_area_salida(marco)
        self.tabla_ap = self._crear_tabla_items_marcados(marco)

        return marco

    def _ap_al_cambiar_parametro(self, *_):
        """Cualquier cambio de proveedor o dólar invalida el dry-run previo."""
        self.ap_dryrun_firma = None
        if not self.proceso_en_curso:
            self.btn_ap_confirmar.configure(state="disabled")

    def _ap_firma_actual(self):
        return (self.var_proveedor.get().strip(), self.var_dolar_ap.get().strip())

    def _ap_dolar_validado(self):
        """Devuelve (ok, dolar). dolar es None si el campo está vacío, o el
        valor normalizado con punto decimal. Si es inválido muestra el error."""
        texto = self.var_dolar_ap.get().strip()
        if not texto:
            return True, None
        try:
            return True, str(parse_dolar(texto))
        except ValueError as e:
            mensaje = str(e)
            messagebox.showerror("Valor de dólar inválido", mensaje[:1].upper() + mensaje[1:])
            return False, None

    def _ap_armar_comando(self, proveedor, dolar, execute):
        cmd = [PYTHON, str(SCRIPT_DIR / "actualizar_precios.py"), proveedor]
        if dolar is not None:
            cmd += ["--dolar", dolar]
        if execute:
            cmd += ["--execute"]
        return cmd

    def _ap_correr_preview(self):
        proveedor = self.var_proveedor.get().strip()
        if not proveedor:
            messagebox.showerror("Falta el proveedor", "Elegí un proveedor antes de continuar.")
            return
        ok, dolar = self._ap_dolar_validado()
        if not ok:
            return
        self.config_interfaz["proveedor"] = proveedor
        guardar_config(self.config_interfaz)
        self.btn_ap_confirmar.configure(state="disabled")
        firma = self._ap_firma_actual()
        self.ap_dryrun_firma = None
        cmd = self._ap_armar_comando(proveedor, dolar, execute=False)
        self._ejecutar_comando_async(
            cmd, self.texto_ap, [self.btn_ap_preview, self.btn_ap_confirmar],
            tabla=self.tabla_ap,
            on_finish=lambda codigo: self._ap_al_terminar_preview(codigo, firma),
        )

    def _ap_al_terminar_preview(self, codigo, firma):
        # _finalizar_proceso() rehabilita los botones a ciegas; si el dry-run
        # falló o se cambió el proveedor/dólar mientras corría, no vale.
        valido = codigo == 0 and firma == self._ap_firma_actual()
        self.ap_dryrun_firma = firma if valido else None
        self.btn_ap_confirmar.configure(state="normal" if valido else "disabled")

    def _ap_correr_confirmar(self):
        proveedor = self.var_proveedor.get().strip()
        if not proveedor:
            return
        if self.ap_dryrun_firma != self._ap_firma_actual():
            messagebox.showerror(
                "Falta la vista previa",
                "Corré primero la vista previa con este mismo proveedor y valor de dólar.",
            )
            self.btn_ap_confirmar.configure(state="disabled")
            return
        ok, dolar = self._ap_dolar_validado()
        if not ok:
            return
        texto_dolar = f" y asignar dólar {dolar} a DolarReferencia" if dolar is not None else ""
        if not messagebox.askyesno(
            "Confirmar cambios reales",
            f"Esto va a modificar el archivo real para el proveedor '{proveedor}'"
            f"{texto_dolar} (con backup previo). ¿Confirmás?",
        ):
            return
        self.config_interfaz["proveedor"] = proveedor
        guardar_config(self.config_interfaz)
        cmd = self._ap_armar_comando(proveedor, dolar, execute=True)
        self._ejecutar_comando_async(
            cmd, self.texto_ap, [self.btn_ap_preview, self.btn_ap_confirmar],
            tabla=self.tabla_ap,
            on_finish=lambda codigo: messagebox.showinfo(
                "Listo", "Cambios aplicados correctamente."
            ) if codigo == 0 else messagebox.showerror(
                "Hubo un problema", "Revisá el detalle en el cuadro de salida."
            ),
        )

    # ------------------------------------------------------------------
    # Pantalla: asignar dólar
    # ------------------------------------------------------------------
    def mostrar_asignar_dolar(self):
        self._cambiar_pantalla(
            lambda padre: self._construir_pantalla_asignar_dolar(padre, recalcular=False)
        )

    def mostrar_actualizar_precio_dolar(self):
        self._cambiar_pantalla(
            lambda padre: self._construir_pantalla_asignar_dolar(padre, recalcular=True)
        )

    def _construir_pantalla_asignar_dolar(self, padre, recalcular=False):
        """Pantalla compartida por 'Asignar dólar' (Tipo 2, no recalcula
        precios) y 'Actualizar precio por dólar' (Tipo 4, recalcula
        PrecioConIVA). Solo cambian el script, el título y la configuración
        guardada; los campos son los mismos."""
        marco = ttk.Frame(padre, padding=20)

        if recalcular:
            self.ad_script = "actualizar_precio_dolar.py"
            self.ad_config_clave = "actualizar_precio_dolar"
            titulo = "Actualizar precio por dólar"
            aclaracion = ("Recalcula PrecioConIVA: precio actual ÷ DolarReferencia actual × "
                          "nuevo dólar, y escribe el nuevo dólar en DolarReferencia.")
        else:
            self.ad_script = "asignar_dolar.py"
            self.ad_config_clave = "asignar_dolar"
            titulo = "Asignar dólar"
            aclaracion = ("Solo escribe el nuevo valor en DolarReferencia (y la fecha). "
                          "PrecioConIVA no se modifica.")

        self._encabezado_pantalla(marco, titulo)
        ttk.Label(marco, text=aclaracion, style="Subtitulo.TLabel").pack(
            anchor="w", pady=(0, 6)
        )

        config_ad = self.config_interfaz.get(self.ad_config_clave, {})
        modo_guardado = config_ad.get("modo")
        if modo_guardado not in ("marca", "fecha", "marca_fecha"):
            modo_guardado = "marca"

        # Firma (modo + marcas + fechas + precio) del último dry-run que
        # terminó bien. "Confirmar y aplicar" solo se habilita si la firma
        # actual coincide con esta; cualquier cambio de parámetro la anula.
        self.ad_dryrun_firma = None
        self.ad_marcas_cargando = False
        self.ad_marcas_guardadas = set(config_ad.get("marcas", []))

        # --- Tipo de actualización ---
        self.var_modo = tk.StringVar(value=modo_guardado)
        fila_modo = ttk.Frame(marco)
        fila_modo.pack(fill="x", pady=(4, 12))
        ttk.Label(fila_modo, text="Tipo de actualización:").pack(side="left", padx=(0, 8))
        for valor, texto in (
            ("marca", "Por marca"), ("fecha", "Por fecha"), ("marca_fecha", "Marca y fecha"),
        ):
            ttk.Radiobutton(
                fila_modo, text=texto, value=valor, variable=self.var_modo,
                command=self._ad_al_cambiar_modo,
            ).pack(side="left", padx=(0, 10))

        # --- Zona de parámetros -------------------------------------------
        # Contenedor único, empaquetado UNA sola vez y siempre antes de los
        # botones. Adentro se usa grid: ocultar con grid_remove() conserva la
        # fila asignada, así que volver a mostrar un campo lo devuelve a su
        # lugar. Con pack() no pasaba eso: cada re-empaquetado mandaba el
        # campo al final de la lista de hermanos, o sea debajo del cuadro de
        # salida y de la tabla de revisión.
        self.marco_params = ttk.Frame(marco)
        self.marco_params.pack(fill="x")
        self.marco_params.columnconfigure(0, weight=1)

        # Fila 0: nuevo valor de dólar (visible en los tres modos)
        marco_precio = ttk.Frame(self.marco_params)
        marco_precio.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(marco_precio, text="Nuevo valor de dólar:").pack(side="left", padx=(0, 8))
        self.var_precio = tk.StringVar(value=config_ad.get("precio_actualizado", ""))
        ttk.Entry(marco_precio, textvariable=self.var_precio, width=14).pack(side="left")

        # Fila 1: fecha a registrar en FechaActualizacion (visible en los tres
        # modos). Arranca siempre en HOY; no se guarda en la configuración.
        hoy_reg = datetime.date.today()
        marco_fecha_reg = ttk.Frame(self.marco_params)
        marco_fecha_reg.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(marco_fecha_reg, text="Fecha a registrar:").pack(side="left", padx=(0, 8))
        self.var_reg_dia = tk.StringVar(value=f"{hoy_reg.day:02d}")
        self.var_reg_mes = tk.StringVar(value=f"{hoy_reg.month:02d}")
        self.var_reg_anio = tk.StringVar(value=str(hoy_reg.year))
        self.combos_fecha_reg = self._ad_combos_fecha(
            marco_fecha_reg, self.var_reg_dia, self.var_reg_mes, self.var_reg_anio
        )
        self.var_no_tocar_fecha = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            marco_fecha_reg, text="No modificar la fecha de actualización",
            variable=self.var_no_tocar_fecha, command=self._ad_al_cambiar_no_tocar_fecha,
        ).pack(side="left", padx=(24, 0))

        # Fila 2: marcas
        self.marco_marcas = ttk.Frame(self.marco_params)
        self.marco_marcas.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        fila_marcas_header = ttk.Frame(self.marco_marcas)
        fila_marcas_header.pack(fill="x", pady=(0, 4))
        ttk.Label(fila_marcas_header, text="Marca(s) (podés elegir varias):").pack(side="left")
        self.btn_ad_actualizar_marcas = ttk.Button(
            fila_marcas_header, text="Actualizar lista", style="Normal.TButton",
            command=self._ad_cargar_marcas,
        )
        self.btn_ad_actualizar_marcas.pack(side="right")

        self.lbl_ad_estado_marcas = tk.Label(
            self.marco_marcas, text="Cargando marcas…", bg=COLOR_FONDO,
            fg=COLOR_TEXTO_SECUNDARIO, font=FUENTE_SUBTITULO, anchor="w",
            justify="left", wraplength=820,
        )
        self.lbl_ad_estado_marcas.pack(anchor="w", fill="x", pady=(0, 4))

        contenedor_marcas = tk.Frame(self.marco_marcas, bg=COLOR_BORDE)
        contenedor_marcas.pack(fill="x")
        self.lista_marcas = tk.Listbox(
            contenedor_marcas, selectmode=tk.MULTIPLE, height=5, bg=COLOR_TARJETA,
            fg=COLOR_TEXTO, selectbackground=COLOR_ACENTO, selectforeground=COLOR_ACENTO_TEXTO,
            highlightthickness=0, borderwidth=0, exportselection=False,
        )
        scroll_marcas = ttk.Scrollbar(contenedor_marcas, command=self.lista_marcas.yview)
        self.lista_marcas.configure(yscrollcommand=scroll_marcas.set)
        self.lista_marcas.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scroll_marcas.pack(side="right", fill="y")
        self.lista_marcas.bind("<<ListboxSelect>>", self._ad_al_cambiar_parametro)

        # Fila 3: rango de fechas (dos tercias de Combobox día/mes/año)
        # Default: desde el 1 del mes de ayer, hasta AYER. No hasta hoy:
        # por defecto el script sella FechaActualizacion con la fecha
        # de hoy, así que un rango que incluya hoy vuelve a capturar en la
        # corrida siguiente las mismas filas que acaba de actualizar.
        hoy = datetime.date.today()
        ayer = hoy - datetime.timedelta(days=1)
        desde_default = ayer.replace(day=1)

        self.marco_fechas = ttk.Frame(self.marco_params)
        self.marco_fechas.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        self.var_desde_dia = tk.StringVar(
            value=config_ad.get("desde_dia") or f"{desde_default.day:02d}")
        self.var_desde_mes = tk.StringVar(
            value=config_ad.get("desde_mes") or f"{desde_default.month:02d}")
        self.var_desde_anio = tk.StringVar(
            value=config_ad.get("desde_anio") or str(desde_default.year))
        self.var_hasta_dia = tk.StringVar(
            value=config_ad.get("hasta_dia") or f"{ayer.day:02d}")
        self.var_hasta_mes = tk.StringVar(
            value=config_ad.get("hasta_mes") or f"{ayer.month:02d}")
        self.var_hasta_anio = tk.StringVar(
            value=config_ad.get("hasta_anio") or str(ayer.year))

        ttk.Label(self.marco_fechas, text="Fecha desde:").pack(side="left", padx=(0, 8))
        self._ad_combos_fecha(
            self.marco_fechas, self.var_desde_dia, self.var_desde_mes, self.var_desde_anio
        )
        ttk.Label(self.marco_fechas, text="Fecha hasta:").pack(side="left", padx=(24, 8))
        self._ad_combos_fecha(
            self.marco_fechas, self.var_hasta_dia, self.var_hasta_mes, self.var_hasta_anio
        )

        # --- Resumen del filtro + errores de validación ---
        self.lbl_ad_resumen = tk.Label(
            marco, text="", bg=COLOR_FONDO, fg=COLOR_TEXTO_SECUNDARIO,
            font=FUENTE_SUBTITULO, anchor="w", justify="left", wraplength=820,
        )
        self.lbl_ad_resumen.pack(fill="x", anchor="w", pady=(2, 0))

        self.lbl_ad_error = tk.Label(
            marco, text="", bg=COLOR_FONDO, fg=COLOR_ERROR,
            font=FUENTE_SUBTITULO, anchor="w", justify="left", wraplength=820,
        )
        self.lbl_ad_error.pack(fill="x", anchor="w", pady=(2, 6))

        # --- Botones (siempre después de todos los parámetros) ---
        fila_botones = ttk.Frame(marco)
        fila_botones.pack(fill="x", pady=(0, 10))
        self.btn_ad_preview = ttk.Button(
            fila_botones, text="Vista previa (dry-run)", style="Normal.TButton",
            command=self._ad_correr_preview, state="disabled",
        )
        self.btn_ad_preview.pack(side="left", padx=(0, 8))
        self.btn_ad_confirmar = ttk.Button(
            fila_botones, text="Confirmar y aplicar", style="Acento.TButton",
            command=self._ad_correr_confirmar, state="disabled",
        )
        self.btn_ad_confirmar.pack(side="left")

        # --- Salida y tabla: siempre al final ---
        self.texto_ad = self._crear_area_salida(marco)
        self.tabla_ad = self._crear_tabla_items_marcados(marco)

        # Los trace se enganchan recién acá, con todos los widgets ya creados:
        # el callback toca los botones y el rótulo de resumen.
        for var in (
            self.var_precio,
            self.var_desde_dia, self.var_desde_mes, self.var_desde_anio,
            self.var_hasta_dia, self.var_hasta_mes, self.var_hasta_anio,
            self.var_reg_dia, self.var_reg_mes, self.var_reg_anio,
        ):
            var.trace_add("write", self._ad_al_cambiar_parametro)

        self._ad_actualizar_visibilidad()
        self._ad_cargar_marcas()
        return marco

    def _ad_combos_fecha(self, padre, var_dia, var_mes, var_anio):
        """Tres Combobox readonly día / mes / año, con el mismo criterio que
        la pantalla de Tipo 3. Se empaquetan con side='left' sobre 'padre'."""
        hoy = datetime.date.today()
        dias = [f"{d:02d}" for d in range(1, 32)]
        meses = [f"{m:02d}" for m in range(1, 13)]
        anios = [str(a) for a in range(hoy.year - 1, hoy.year + 6)]

        combos = []
        for var, valores, ancho in ((var_dia, dias, 3), (var_mes, meses, 3), (var_anio, anios, 5)):
            if combos:
                ttk.Label(padre, text="/").pack(side="left", padx=2)
            combo = ttk.Combobox(
                padre, textvariable=var, values=valores, state="readonly",
                width=ancho, justify="center",
            )
            combo.pack(side="left")
            combos.append(combo)
        return tuple(combos)

    # -- Carga de marcas disponibles --
    def _ad_cargar_marcas(self):
        """Relee las marcas del Excel. Distingue tres desenlaces en el rótulo
        de estado: marcas encontradas, archivo legible pero sin marcas, y
        error de lectura (el caso típico: el libro abierto en Excel)."""
        seleccionadas = set(self._ad_marcas_seleccionadas())
        if seleccionadas:
            self.ad_marcas_guardadas = seleccionadas

        self.ad_marcas_cargando = True
        self.ad_dryrun_firma = None
        self.lista_marcas.delete(0, tk.END)
        self.lbl_ad_estado_marcas.configure(text="Cargando marcas…", fg=COLOR_TEXTO_SECUNDARIO)
        self.btn_ad_actualizar_marcas.configure(state="disabled")
        self.root.update_idletasks()

        marcas = []
        try:
            marcas = listar_marcas(ARCHIVO_DOLAR)
        except Exception as e:
            self.lbl_ad_estado_marcas.configure(
                text=(
                    f"No se pudo leer el Excel ({type(e).__name__}): {e}\n"
                    "Si lo tenés abierto en Excel, cerralo y tocá «Actualizar lista»."
                ),
                fg=COLOR_ERROR,
            )
        else:
            if marcas:
                self.lbl_ad_estado_marcas.configure(
                    text=f"{len(marcas)} marca(s) disponibles.", fg=COLOR_TEXTO_SECUNDARIO
                )
            else:
                self.lbl_ad_estado_marcas.configure(
                    text=(
                        "El Excel se leyó bien, pero no se encontró ninguna marca "
                        "en la columna 'Marca' de ninguna hoja."
                    ),
                    fg=COLOR_ALERTA,
                )

        for i, m in enumerate(marcas):
            self.lista_marcas.insert(tk.END, m)
            if m in self.ad_marcas_guardadas:
                self.lista_marcas.selection_set(i)

        self.ad_marcas_cargando = False
        self.btn_ad_actualizar_marcas.configure(state="normal")
        self._ad_al_cambiar_parametro()

    # -- Estado de la pantalla --
    def _ad_marcas_seleccionadas(self):
        return [self.lista_marcas.get(i) for i in self.lista_marcas.curselection()]

    def _ad_al_cambiar_no_tocar_fecha(self):
        """Con la casilla tildada, los combos de 'Fecha a registrar' quedan
        deshabilitados: el script no va a escribir FechaActualizacion."""
        estado = "disabled" if self.var_no_tocar_fecha.get() else "readonly"
        for combo in self.combos_fecha_reg:
            combo.configure(state=estado)
        self._ad_al_cambiar_parametro()

    def _ad_fecha_texto(self, cual):
        """'desde', 'hasta' o 'reg' -> 'dd/mm/aaaa', tal como lo espera el script."""
        return "{}/{}/{}".format(
            getattr(self, f"var_{cual}_dia").get(),
            getattr(self, f"var_{cual}_mes").get(),
            getattr(self, f"var_{cual}_anio").get(),
        )

    def _ad_fecha_valida(self, cual):
        """Devuelve date, o None si la combinación no existe (31/02, por
        ejemplo: los Combobox permiten armarla)."""
        try:
            return datetime.datetime.strptime(self._ad_fecha_texto(cual), "%d/%m/%Y").date()
        except ValueError:
            return None

    def _ad_actualizar_visibilidad(self):
        modo = self.var_modo.get()
        if modo in ("marca", "marca_fecha"):
            self.marco_marcas.grid()
        else:
            self.marco_marcas.grid_remove()
        if modo in ("fecha", "marca_fecha"):
            self.marco_fechas.grid()
        else:
            self.marco_fechas.grid_remove()

    def _ad_al_cambiar_modo(self):
        self._ad_actualizar_visibilidad()
        self._ad_al_cambiar_parametro()

    def _ad_al_cambiar_parametro(self, *_):
        """Punto único de refresco: invalida el dry-run previo, reescribe el
        resumen y recalcula el estado de los botones. Lo llaman los trace_add
        de las StringVar, el <<ListboxSelect>> de las marcas y los radios."""
        self.ad_dryrun_firma = None
        self._ad_actualizar_resumen()
        self._ad_actualizar_estado_botones()

    def _ad_firma_actual(self):
        modo = self.var_modo.get()
        marcas = tuple(self._ad_marcas_seleccionadas()) if modo in ("marca", "marca_fecha") else ()
        fechas = (
            (self._ad_fecha_texto("desde"), self._ad_fecha_texto("hasta"))
            if modo in ("fecha", "marca_fecha") else ()
        )
        fecha_reg = None if self.var_no_tocar_fecha.get() else self._ad_fecha_texto("reg")
        return (modo, marcas, fechas, self.var_precio.get().strip().replace(",", "."), fecha_reg)

    def _ad_validar(self):
        """Devuelve (ok, mensaje). El mensaje va al rótulo rojo de la pantalla."""
        modo = self.var_modo.get()

        _, error = validar_dolar(self.var_precio.get())
        if error is not None:
            return False, error

        if modo in ("marca", "marca_fecha") and not self._ad_marcas_seleccionadas():
            return False, "Elegí al menos una marca de la lista."

        if modo in ("fecha", "marca_fecha"):
            desde = self._ad_fecha_valida("desde")
            hasta = self._ad_fecha_valida("hasta")
            if desde is None:
                return False, "La 'fecha desde' no existe (revisá el día para ese mes y año)."
            if hasta is None:
                return False, "La 'fecha hasta' no existe (revisá el día para ese mes y año)."
            if desde > hasta:
                return False, "La 'fecha desde' tiene que ser anterior o igual a la 'fecha hasta'."

        if not self.var_no_tocar_fecha.get() and self._ad_fecha_valida("reg") is None:
            return False, "La 'fecha a registrar' no existe (revisá el día para ese mes y año)."

        return True, ""

    def _ad_actualizar_resumen(self):
        modo = self.var_modo.get()
        partes = []
        if modo in ("fecha", "marca_fecha"):
            partes.append(
                f"Filtro: {self._ad_fecha_texto('desde')} a {self._ad_fecha_texto('hasta')}"
            )
        if modo in ("marca", "marca_fecha"):
            marcas = self._ad_marcas_seleccionadas()
            if not marcas:
                detalle = "ninguna seleccionada"
            elif len(marcas) <= 8:
                detalle = ", ".join(marcas)
            else:
                detalle = ", ".join(marcas[:8]) + f" (+{len(marcas) - 8} más)"
            partes.append(f"Marcas: {detalle}")
        if self.var_no_tocar_fecha.get():
            partes.append("Fecha de actualización: sin cambios")
        else:
            partes.append(f"Fecha a registrar: {self._ad_fecha_texto('reg')}")
        self.lbl_ad_resumen.configure(text=" — ".join(partes))

    def _ad_actualizar_estado_botones(self):
        ok, _ = self._ad_validar()
        ocupado = self.proceso_en_curso or self.ad_marcas_cargando
        self.btn_ad_preview.configure(state="normal" if ok and not ocupado else "disabled")
        listo = (
            ok and not ocupado
            and self.ad_dryrun_firma is not None
            and self.ad_dryrun_firma == self._ad_firma_actual()
        )
        self.btn_ad_confirmar.configure(state="normal" if listo else "disabled")

    # -- Armado y ejecución del comando --
    def _ad_armar_comando(self, confirmar):
        ok, mensaje = self._ad_validar()
        if not ok:
            self.lbl_ad_error.configure(text=mensaje)
            return None
        self.lbl_ad_error.configure(text="")

        modo = self.var_modo.get()
        precio_texto = self.var_precio.get().strip().replace(",", ".")
        cmd = [PYTHON, str(SCRIPT_DIR / self.ad_script), "--modo", modo]

        marcas = []
        if modo in ("marca", "marca_fecha"):
            marcas = self._ad_marcas_seleccionadas()
            cmd += ["--marca", ",".join(marcas)]

        if modo in ("fecha", "marca_fecha"):
            cmd += [
                "--fecha-desde", self._ad_fecha_texto("desde"),
                "--fecha-hasta", self._ad_fecha_texto("hasta"),
            ]

        cmd += ["--precio-actualizado", precio_texto]

        if self.var_no_tocar_fecha.get():
            cmd += ["--no-tocar-fecha"]
        else:
            cmd += ["--fecha-registro", self._ad_fecha_texto("reg")]

        if confirmar:
            cmd += ["--confirmar"]

        self.config_interfaz[self.ad_config_clave] = {
            "modo": modo,
            "marcas": marcas,
            "precio_actualizado": precio_texto,
            "desde_dia": self.var_desde_dia.get(),
            "desde_mes": self.var_desde_mes.get(),
            "desde_anio": self.var_desde_anio.get(),
            "hasta_dia": self.var_hasta_dia.get(),
            "hasta_mes": self.var_hasta_mes.get(),
            "hasta_anio": self.var_hasta_anio.get(),
        }
        guardar_config(self.config_interfaz)

        return cmd

    def _ad_botones_bloqueables(self):
        return [self.btn_ad_preview, self.btn_ad_confirmar, self.btn_ad_actualizar_marcas]

    def _ad_correr_preview(self):
        cmd = self._ad_armar_comando(confirmar=False)
        if cmd is None:
            return
        firma = self._ad_firma_actual()
        self.ad_dryrun_firma = None
        self._ejecutar_comando_async(
            cmd, self.texto_ad, self._ad_botones_bloqueables(),
            tabla=self.tabla_ad,
            on_finish=lambda codigo: self._ad_al_terminar_preview(codigo, firma),
        )

    def _ad_al_terminar_preview(self, codigo, firma):
        # _finalizar_proceso() rehabilita todos los botones a ciegas; esto
        # corre después y deja el estado real. Se compara contra la firma
        # capturada al lanzar: si tocaste algún parámetro mientras corría,
        # el dry-run ya no vale.
        self.ad_dryrun_firma = firma if codigo == 0 else None
        self._ad_actualizar_estado_botones()

    def _ad_correr_confirmar(self):
        if self.ad_dryrun_firma != self._ad_firma_actual():
            self.lbl_ad_error.configure(
                text="Corré primero la vista previa con estos mismos parámetros."
            )
            self._ad_actualizar_estado_botones()
            return
        cmd = self._ad_armar_comando(confirmar=True)
        if cmd is None:
            return
        if not messagebox.askyesno(
            "Confirmar cambios reales",
            "Esto va a modificar el archivo real (con backup previo). ¿Confirmás?",
        ):
            return
        self._ejecutar_comando_async(
            cmd, self.texto_ad, self._ad_botones_bloqueables(),
            tabla=self.tabla_ad,
            on_finish=self._ad_al_terminar_confirmar,
        )

    def _ad_al_terminar_confirmar(self, codigo):
        # Tras aplicar, el dry-run deja de ser válido: las filas ya cambiaron.
        self.ad_dryrun_firma = None
        self._ad_actualizar_estado_botones()
        if codigo == 0:
            messagebox.showinfo("Listo", "Cambios aplicados correctamente.")
        else:
            messagebox.showerror("Hubo un problema", "Revisá el detalle en el cuadro de salida.")

    # ------------------------------------------------------------------
    # Pantalla: actualizar fecha (Tipo 3)
    # ------------------------------------------------------------------
    def mostrar_actualizar_fecha(self):
        self._cambiar_pantalla(self._construir_pantalla_actualizar_fecha)

    def _construir_pantalla_actualizar_fecha(self, padre):
        marco = ttk.Frame(padre, padding=20)

        self._encabezado_pantalla(marco, "Actualizar fecha")
        ttk.Label(
            marco,
            text="Deja constancia de que se revisó una marca este mes, sin tocar precios.",
            style="Subtitulo.TLabel",
        ).pack(anchor="w", pady=(0, 14))

        config_af = self.config_interfaz.get("actualizar_fecha", {})
        self.af_dryrun_ok = False
        self.af_marcas_cargando = False
        self.af_marcas_guardadas = set(config_af.get("marcas", []))

        # --- Marca(s): lista de selección múltiple, cargada por subprocess ---
        fila_marcas_header = ttk.Frame(marco)
        fila_marcas_header.pack(fill="x", pady=(0, 4))
        ttk.Label(fila_marcas_header, text="Marca(s) (podés elegir varias):").pack(side="left")
        self.btn_af_actualizar_marcas = ttk.Button(
            fila_marcas_header, text="Actualizar lista", style="Normal.TButton",
            command=self._af_cargar_marcas_async,
        )
        self.btn_af_actualizar_marcas.pack(side="right")

        self.lbl_af_estado_marcas = tk.Label(
            marco, text="Cargando marcas…", bg=COLOR_FONDO, fg=COLOR_TEXTO_SECUNDARIO,
            font=FUENTE_SUBTITULO, anchor="w",
        )
        self.lbl_af_estado_marcas.pack(anchor="w", pady=(0, 4))

        contenedor_marcas = tk.Frame(marco, bg=COLOR_BORDE)
        contenedor_marcas.pack(fill="x", pady=(0, 12))
        self.lista_af_marcas = tk.Listbox(
            contenedor_marcas, selectmode="extended", height=7, bg=COLOR_TARJETA,
            fg=COLOR_TEXTO, selectbackground=COLOR_ACENTO, selectforeground=COLOR_ACENTO_TEXTO,
            highlightthickness=0, borderwidth=0, exportselection=False,
        )
        scroll_af_marcas = ttk.Scrollbar(contenedor_marcas, command=self.lista_af_marcas.yview)
        self.lista_af_marcas.configure(yscrollcommand=scroll_af_marcas.set)
        self.lista_af_marcas.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scroll_af_marcas.pack(side="right", fill="y")
        self.lista_af_marcas.bind("<<ListboxSelect>>", self._af_on_seleccion_marcas)

        # --- Modo de fecha: radios + selector día/mes/año a la derecha ---
        modo_af_guardado = config_af.get("modo")
        if modo_af_guardado not in ("hoy", "especifica"):
            modo_af_guardado = "hoy"
        self.var_af_modo = tk.StringVar(value=modo_af_guardado)

        hoy = datetime.date.today()
        dias = [f"{d:02d}" for d in range(1, 32)]
        meses = [f"{m:02d}" for m in range(1, 13)]
        anios = [str(a) for a in range(hoy.year - 1, hoy.year + 6)]

        self.var_af_dia = tk.StringVar(value=config_af.get("fecha_dia") or f"{hoy.day:02d}")
        self.var_af_mes = tk.StringVar(value=config_af.get("fecha_mes") or f"{hoy.month:02d}")
        self.var_af_anio = tk.StringVar(value=config_af.get("fecha_anio") or str(hoy.year))

        fila_modo_af = ttk.Frame(marco)
        fila_modo_af.pack(fill="x", pady=(0, 12))
        ttk.Radiobutton(
            fila_modo_af, text="Fecha de hoy", value="hoy", variable=self.var_af_modo,
            command=self._af_actualizar_visibilidad,
        ).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(
            fila_modo_af, text="Fecha específica:", value="especifica", variable=self.var_af_modo,
            command=self._af_actualizar_visibilidad,
        ).pack(side="left", padx=(0, 8))

        self.combo_af_dia = ttk.Combobox(
            fila_modo_af, textvariable=self.var_af_dia, values=dias, state="disabled",
            width=3, justify="center",
        )
        self.combo_af_dia.pack(side="left")
        ttk.Label(fila_modo_af, text="/").pack(side="left", padx=2)
        self.combo_af_mes = ttk.Combobox(
            fila_modo_af, textvariable=self.var_af_mes, values=meses, state="disabled",
            width=3, justify="center",
        )
        self.combo_af_mes.pack(side="left")
        ttk.Label(fila_modo_af, text="/").pack(side="left", padx=2)
        self.combo_af_anio = ttk.Combobox(
            fila_modo_af, textvariable=self.var_af_anio, values=anios, state="disabled",
            width=5, justify="center",
        )
        self.combo_af_anio.pack(side="left")

        fila_botones = ttk.Frame(marco)
        fila_botones.pack(fill="x", pady=(2, 10))
        self.btn_af_preview = ttk.Button(
            fila_botones, text="Vista previa (dry-run)", style="Normal.TButton",
            command=self._af_correr_preview, state="disabled",
        )
        self.btn_af_preview.pack(side="left", padx=(0, 8))
        self.btn_af_confirmar = ttk.Button(
            fila_botones, text="Confirmar y aplicar", style="Acento.TButton",
            command=self._af_correr_confirmar, state="disabled",
        )
        self.btn_af_confirmar.pack(side="left")

        self.texto_af = self._crear_area_salida(marco)

        self._af_actualizar_visibilidad()
        self._af_cargar_marcas_async()
        return marco

    def _af_actualizar_visibilidad(self):
        estado = "readonly" if self.var_af_modo.get() == "especifica" else "disabled"
        self.combo_af_dia.configure(state=estado)
        self.combo_af_mes.configure(state=estado)
        self.combo_af_anio.configure(state=estado)

    # -- Carga de marcas disponibles (subprocess: actualizar_fecha.py --listar-marcas) --
    def _af_cargar_marcas_async(self):
        if self.af_marcas_cargando:
            return
        self.af_marcas_cargando = True
        self.af_dryrun_ok = False
        seleccionadas_actuales = {
            self.lista_af_marcas.get(i) for i in self.lista_af_marcas.curselection()
        }
        if seleccionadas_actuales:
            self.af_marcas_guardadas = seleccionadas_actuales
        self.btn_af_actualizar_marcas.configure(state="disabled")
        self.lista_af_marcas.delete(0, tk.END)
        self.lbl_af_estado_marcas.configure(text="Cargando marcas…", fg=COLOR_TEXTO_SECUNDARIO)
        self._af_actualizar_estado_botones()

        cola = queue.Queue()
        cmd = [PYTHON, str(SCRIPT_DIR / "actualizar_fecha.py"), "--listar-marcas"]

        def trabajo():
            try:
                env = os.environ.copy()
                env["PYTHONUTF8"] = "1"
                env["PYTHONIOENCODING"] = "utf-8"
                resultado = subprocess.run(
                    cmd, cwd=str(SCRIPT_DIR), capture_output=True, text=True,
                    encoding="utf-8", errors="replace", env=env,
                )
                cola.put(("fin", resultado.returncode, resultado.stdout, resultado.stderr))
            except Exception as e:
                cola.put(("excepcion", str(e)))

        threading.Thread(target=trabajo, daemon=True).start()

        def sondear():
            try:
                item = cola.get_nowait()
            except queue.Empty:
                self.root.after(80, sondear)
                return
            self.af_marcas_cargando = False
            self.btn_af_actualizar_marcas.configure(state="normal")
            if item[0] == "excepcion":
                self._af_mostrar_error_marcas(item[1])
            else:
                _, codigo, stdout, stderr = item
                if codigo != 0:
                    mensaje = stderr.strip() or stdout.strip() or f"El script terminó con código {codigo}."
                    self._af_mostrar_error_marcas(mensaje)
                else:
                    marcas = [linea.strip() for linea in stdout.splitlines() if linea.strip()]
                    self._af_llenar_lista_marcas(marcas)

        sondear()

    def _af_llenar_lista_marcas(self, marcas):
        self.lista_af_marcas.delete(0, tk.END)
        for i, m in enumerate(marcas):
            self.lista_af_marcas.insert(tk.END, m)
            if m in self.af_marcas_guardadas:
                self.lista_af_marcas.selection_set(i)
        if marcas:
            self.lbl_af_estado_marcas.configure(
                text=f"{len(marcas)} marca(s) disponibles.", fg=COLOR_TEXTO_SECUNDARIO
            )
        else:
            self.lbl_af_estado_marcas.configure(
                text="No se encontraron marcas en el archivo.", fg=COLOR_ALERTA
            )
        self._af_actualizar_estado_botones()

    def _af_mostrar_error_marcas(self, mensaje):
        self.lista_af_marcas.delete(0, tk.END)
        self.lbl_af_estado_marcas.configure(
            text="No se pudo cargar la lista de marcas (ver detalle abajo).", fg=COLOR_ERROR
        )
        self.texto_af.insert(tk.END, f"\nERROR al listar marcas: {mensaje}\n", "error")
        self.texto_af.see(tk.END)
        self._af_actualizar_estado_botones()

    def _af_on_seleccion_marcas(self, event=None):
        self.af_dryrun_ok = False
        self._af_actualizar_estado_botones()

    def _af_actualizar_estado_botones(self):
        hay_seleccion = bool(self.lista_af_marcas.curselection())
        ocupado = self.proceso_en_curso or self.af_marcas_cargando
        self.btn_af_preview.configure(state="normal" if hay_seleccion and not ocupado else "disabled")
        self.btn_af_confirmar.configure(
            state="normal" if hay_seleccion and self.af_dryrun_ok and not ocupado else "disabled"
        )

    # -- Armado y ejecución del comando --
    def _af_armar_comando(self, confirmar):
        marcas = [self.lista_af_marcas.get(i) for i in self.lista_af_marcas.curselection()]
        if not marcas:
            messagebox.showerror(
                "Falta la marca", "Elegí al menos una marca de la lista."
            )
            return None

        modo = self.var_af_modo.get()
        cmd = [PYTHON, str(SCRIPT_DIR / "actualizar_fecha.py"), "--marca", ",".join(marcas)]

        dia, mes, anio = self.var_af_dia.get(), self.var_af_mes.get(), self.var_af_anio.get()
        if modo == "hoy":
            cmd += ["--hoy"]
        else:
            fecha_texto = f"{dia}/{mes}/{anio}"
            if not self._af_fecha_valida(fecha_texto):
                messagebox.showerror(
                    "Fecha inválida",
                    "Elegí una fecha real (el día no existe en ese mes/año).",
                )
                return None
            cmd += ["--fecha", fecha_texto]

        if confirmar:
            cmd += ["--execute"]

        self.config_interfaz["actualizar_fecha"] = {
            "marcas": marcas, "modo": modo,
            "fecha_dia": dia, "fecha_mes": mes, "fecha_anio": anio,
        }
        guardar_config(self.config_interfaz)

        return cmd

    @staticmethod
    def _af_fecha_valida(texto):
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", texto):
            return False
        try:
            datetime.datetime.strptime(texto, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def _af_correr_preview(self):
        cmd = self._af_armar_comando(confirmar=False)
        if cmd is None:
            return
        self.af_dryrun_ok = False
        self._ejecutar_comando_async(
            cmd, self.texto_af,
            [self.btn_af_preview, self.btn_af_confirmar, self.btn_af_actualizar_marcas],
            on_finish=self._af_al_terminar_preview,
        )

    def _af_al_terminar_preview(self, codigo):
        self.af_dryrun_ok = (codigo == 0)
        self._af_actualizar_estado_botones()

    def _af_correr_confirmar(self):
        cmd = self._af_armar_comando(confirmar=True)
        if cmd is None:
            return
        if not messagebox.askyesno(
            "Confirmar cambios reales",
            "Esto va a modificar el archivo real (con backup previo). ¿Confirmás?",
        ):
            return
        self._ejecutar_comando_async(
            cmd, self.texto_af,
            [self.btn_af_preview, self.btn_af_confirmar, self.btn_af_actualizar_marcas],
            on_finish=self._af_al_terminar_confirmar,
        )

    def _af_al_terminar_confirmar(self, codigo):
        self.af_dryrun_ok = False
        self._af_actualizar_estado_botones()
        if codigo == 0:
            messagebox.showinfo("Listo", "Cambios aplicados correctamente.")
        else:
            messagebox.showerror("Hubo un problema", "Revisá el detalle en el cuadro de salida.")

    # ------------------------------------------------------------------
    # Pantalla: exportar datos (Tipo 4)
    # ------------------------------------------------------------------
    def mostrar_exportar_datos(self):
        self._cambiar_pantalla(self._construir_pantalla_exportar_datos)

    def _construir_pantalla_exportar_datos(self, padre):
        marco = ttk.Frame(padre, padding=20)

        self._encabezado_pantalla(marco, "Exportar datos")

        ttk.Label(
            marco,
            text=(
                "Genera el archivo .xls con el código y el precio de cada producto, "
                "en el formato que pide el sistema externo."
            ),
            style="Subtitulo.TLabel",
        ).pack(anchor="w", pady=(0, 2))
        ttk.Label(
            marco,
            text=(
                "Se exportan CodigoDux -> CODIGO y PrecioConIVA -> IMPORTE de todas las "
                "hojas de marca/categoría. Esta operación es de solo lectura: no modifica "
                "el libro principal."
            ),
            style="Subtitulo.TLabel",
            wraplength=760,
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        fila_botones = ttk.Frame(marco)
        fila_botones.pack(fill="x", pady=(0, 10))
        self.btn_ex_preview = ttk.Button(
            fila_botones, text="Vista previa (dry-run)", style="Normal.TButton",
            command=self._ex_correr_preview,
        )
        self.btn_ex_preview.pack(side="left", padx=(0, 8))
        self.btn_ex_confirmar = ttk.Button(
            fila_botones, text="Generar archivo", style="Acento.TButton",
            command=self._ex_correr_confirmar, state="disabled",
        )
        self.btn_ex_confirmar.pack(side="left", padx=(0, 8))
        self._enlace(
            fila_botones, "Abrir carpeta de exportación", self._abrir_carpeta_exportacion
        ).pack(side="left", padx=(12, 0))

        self.texto_ex = self._crear_area_salida(marco)
        return marco

    def _abrir_carpeta_exportacion(self):
        carpeta = SCRIPT_DIR.parent / "exportacion"
        if not carpeta.exists():
            messagebox.showinfo(
                "Todavía no existe",
                "La carpeta de exportación se crea la primera vez que generás el archivo.",
            )
            return
        try:
            os.startfile(str(carpeta))  # noqa: S606 — abre el Explorador de Windows
        except Exception as e:
            messagebox.showerror("No se pudo abrir la carpeta", str(e))

    def _ex_correr_preview(self):
        self.btn_ex_confirmar.configure(state="disabled")
        cmd = [PYTHON, str(SCRIPT_DIR / "exportar_datos.py")]
        self._ejecutar_comando_async(
            cmd, self.texto_ex, [self.btn_ex_preview, self.btn_ex_confirmar],
            on_finish=lambda codigo: self.btn_ex_confirmar.configure(
                state="normal" if codigo == 0 else "disabled"
            ),
        )

    def _ex_correr_confirmar(self):
        if not messagebox.askyesno(
            "Confirmar exportación",
            "Esto va a reemplazar el archivo de exportación anterior "
            "(se guarda una copia en exportacion/backups). ¿Confirmás?",
        ):
            return
        cmd = [PYTHON, str(SCRIPT_DIR / "exportar_datos.py"), "--execute"]
        self._ejecutar_comando_async(
            cmd, self.texto_ex, [self.btn_ex_preview, self.btn_ex_confirmar],
            on_finish=self._ex_al_terminar_confirmar,
        )

    def _ex_al_terminar_confirmar(self, codigo):
        self.btn_ex_confirmar.configure(state="disabled")
        if codigo == 0:
            messagebox.showinfo("Listo", "Archivo de exportación generado correctamente.")
        else:
            messagebox.showerror("Hubo un problema", "Revisá el detalle en el cuadro de salida.")

    # ------------------------------------------------------------------
    # Utilidades comunes de UI
    # ------------------------------------------------------------------
    def _encabezado_pantalla(self, marco, titulo):
        fila = ttk.Frame(marco)
        fila.pack(fill="x", pady=(0, 14))
        ttk.Button(fila, text="< Volver", style="Normal.TButton", command=self.mostrar_inicio).pack(
            side="left"
        )
        ttk.Label(fila, text=titulo, style="Titulo.TLabel").pack(side="left", padx=(14, 0))

    def _crear_area_salida(self, marco):
        contenedor = tk.Frame(marco, bg=COLOR_BORDE)
        contenedor.pack(fill="both", expand=True, pady=(4, 0))
        texto = tk.Text(
            contenedor, bg="#0A0A0A", fg=COLOR_TEXTO, insertbackground=COLOR_TEXTO,
            font=FUENTE_MONO, wrap="word", relief="flat", padx=10, pady=8,
        )
        scroll = ttk.Scrollbar(contenedor, command=texto.yview)
        texto.configure(yscrollcommand=scroll.set)
        texto.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        scroll.pack(side="right", fill="y")

        texto.tag_configure("alerta", foreground=COLOR_ALERTA)
        texto.tag_configure("error", foreground=COLOR_ERROR)
        return texto

    def _crear_tabla_items_marcados(self, marco):
        """Tabla (ttk.Treeview) con los ítems marcados por la revisión
        automática de una corrida (variación de precio grande o valor
        inválido). Se llena a partir del bloque JSON que emiten
        actualizar_precios.py / asignar_dolar.py al final de cada corrida."""
        contenedor = ttk.Frame(marco)
        contenedor.pack(fill="x", pady=(10, 0))

        ttk.Label(contenedor, text="Ítems marcados para revisión:", style="Subtitulo.TLabel").pack(
            anchor="w", pady=(0, 4)
        )

        marco_tabla = ttk.Frame(contenedor)
        marco_tabla.pack(fill="x")

        columnas = ("hoja", "codigo", "producto", "anterior", "nuevo", "motivo")
        encabezados = {
            "hoja": "Hoja", "codigo": "Código", "producto": "Producto",
            "anterior": "Anterior", "nuevo": "Nuevo", "motivo": "Motivo",
        }
        anchos = {"hoja": 90, "codigo": 70, "producto": 150, "anterior": 70, "nuevo": 70, "motivo": 220}

        tabla = ttk.Treeview(marco_tabla, columns=columnas, show="headings", height=6)
        for col in columnas:
            tabla.heading(col, text=encabezados[col])
            tabla.column(col, width=anchos[col], anchor="w")
        scroll_tabla = ttk.Scrollbar(marco_tabla, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=scroll_tabla.set)

        etiqueta_estado = ttk.Label(contenedor, text="Todavía no hay resultados.", style="Subtitulo.TLabel")

        return {
            "marco_tabla": marco_tabla,
            "tabla": tabla,
            "scroll": scroll_tabla,
            "estado": etiqueta_estado,
        }

    def _reiniciar_tabla_items(self, widgets):
        """Limpia la tabla y muestra un estado de 'esperando resultado' al
        arrancar una corrida nueva, para no dejar la tabla de la corrida
        anterior a la vista mientras corre la nueva."""
        widgets["tabla"].delete(*widgets["tabla"].get_children())
        widgets["marco_tabla"].pack_forget()
        widgets["estado"].configure(text="Corriendo… esperando el resultado de esta corrida.")
        widgets["estado"].pack(anchor="w")

    def _llenar_tabla_items(self, widgets, items):
        """Vuelca la lista de ítems marcados (dicts con hoja/codigo/producto/
        valor_anterior/valor_nuevo/motivo) recibida en el bloque JSON de la
        corrida. Lista vacía -> mensaje 'Sin observaciones', nunca una tabla
        vacía."""
        widgets["tabla"].delete(*widgets["tabla"].get_children())
        widgets["estado"].pack_forget()
        widgets["marco_tabla"].pack_forget()

        if not items:
            widgets["estado"].configure(text="Sin observaciones.")
            widgets["estado"].pack(anchor="w")
            return

        for it in items:
            widgets["tabla"].insert("", tk.END, values=(
                it.get("hoja", ""),
                it.get("codigo", ""),
                it.get("producto") or "",
                it.get("valor_anterior", ""),
                it.get("valor_nuevo", ""),
                it.get("motivo", ""),
            ))
        widgets["marco_tabla"].pack(fill="x")
        widgets["tabla"].pack(side="left", fill="both", expand=True)
        widgets["scroll"].pack(side="right", fill="y")

    # ------------------------------------------------------------------
    # Ejecución de comandos en segundo plano (no congela la ventana)
    # ------------------------------------------------------------------
    def _ejecutar_comando_async(self, cmd, widget_texto, botones, tabla=None, on_finish=None):
        if self.proceso_en_curso:
            return
        self.proceso_en_curso = True
        for b in botones:
            b.configure(state="disabled")
        if tabla is not None:
            self._reiniciar_tabla_items(tabla)

        widget_texto.delete("1.0", tk.END)
        widget_texto.insert(tk.END, f"$ {' '.join(cmd[1:])}\n\n")

        cola = queue.Queue()

        def trabajo_en_hilo():
            try:
                env = os.environ.copy()
                env["PYTHONUTF8"] = "1"
                env["PYTHONIOENCODING"] = "utf-8"
                proceso = subprocess.Popen(
                    cmd, cwd=str(SCRIPT_DIR), stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                    errors="replace", bufsize=1, env=env,
                )
                for linea in proceso.stdout:
                    cola.put(("linea", linea))
                proceso.wait()
                cola.put(("fin", proceso.returncode))
            except Exception as e:
                cola.put(("excepcion", str(e)))

        threading.Thread(target=trabajo_en_hilo, daemon=True).start()

        def sondear():
            try:
                while True:
                    tipo, dato = cola.get_nowait()
                    if tipo == "linea":
                        coincidencia = PATRON_ITEMS_JSON.search(dato)
                        if coincidencia:
                            if tabla is not None:
                                try:
                                    items = json.loads(coincidencia.group(1))
                                except json.JSONDecodeError:
                                    items = []
                                self._llenar_tabla_items(tabla, items)
                            continue  # no se muestra el JSON crudo en el cuadro de texto

                        etiqueta = None
                        if "ERROR" in dato.upper():
                            etiqueta = "error"
                        elif "AVISO" in dato.upper() or "atípic" in dato.lower():
                            etiqueta = "alerta"
                        widget_texto.insert(tk.END, dato, etiqueta)
                        widget_texto.see(tk.END)
                    elif tipo == "fin":
                        self._finalizar_proceso(botones)
                        if on_finish:
                            on_finish(dato)
                        return
                    elif tipo == "excepcion":
                        widget_texto.insert(tk.END, f"\nERROR al ejecutar: {dato}\n", "error")
                        self._finalizar_proceso(botones)
                        return
            except queue.Empty:
                pass
            self.root.after(80, sondear)

        sondear()

    def _finalizar_proceso(self, botones):
        self.proceso_en_curso = False
        for b in botones:
            b.configure(state="normal")


def main():
    root = tk.Tk()
    InterfazGiroVivo(root)
    root.mainloop()


if __name__ == "__main__":
    main()
