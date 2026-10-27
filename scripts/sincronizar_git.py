#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sincronizar_git.py
==================
Sincronización del proyecto con GitHub, para trabajar en dos PC sin hacer
git add / commit / push / pull a mano. Lo usa interfaz.py; no depende de
nada fuera de la biblioteca estándar: llama a `git` por subprocess.

Reglas (ver también la sección "Sincronización con GitHub" de CLAUDE.md)
-----------------------------------------------------------------------
- Cada comando git corre desde la raíz del proyecto, con timeout de 60 s,
  con GIT_TERMINAL_PROMPT=0 y GCM_INTERACTIVE=never (nunca queda esperando
  una contraseña) y, en Windows, con CREATE_NO_WINDOW (la interfaz corre
  con pythonw y no deben aparecer ventanas de consola).
- Nunca se usa --force, ni merge, ni rebase, ni stash, ni reset. Si esta PC
  y GitHub tienen cambios distintos, se informa "conflicto" y se resuelve
  a mano.
- .git/index.lock: si existe, tiene más de 2 minutos y no hay ningún
  git.exe corriendo, se borra y se reintenta una vez. Si no, se aborta.
- La sincronización no escribe nada en RegistroCambios: no son cambios de
  celdas, son versiones de archivos.

Cada función pública devuelve un ResultadoSync (estado + mensaje en
español) y nunca lanza excepciones: la interfaz tiene que poder seguir
funcionando aunque git falle.
"""

import datetime
import os
import subprocess
import time
from dataclasses import dataclass, field
from pathlib import Path

RAIZ_PROYECTO = Path(__file__).resolve().parent.parent
REMOTO = "origin"
RAMA = "main"
TIMEOUT_SEGUNDOS = 60
ANTIGUEDAD_MINIMA_LOCK = 120  # segundos

# Solo existe en Windows; en otro sistema vale 0 (sin efecto).
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Textos de git (en inglés: se fuerza LC_ALL=C) que permiten clasificar fallas.
PATRONES_AUTENTICACION = (
    "authentication failed", "could not read username", "could not read password",
    "terminal prompts disabled", "cannot prompt", "interactivity",
    "returned error: 401", "returned error: 403", "repository not found",
    "permission denied (publickey)",
)
PATRONES_SIN_CONEXION = (
    "could not resolve host", "failed to connect", "unable to access",
    "timed out", "connection refused", "network is unreachable",
    "could not connect", "connection was reset", "connection reset",
    "unable to look up",
)
PATRONES_RECHAZO_PUSH = (
    "[rejected]", "fetch first", "non-fast-forward", "updates were rejected",
)
PATRONES_ARCHIVO_EN_USO = (
    "unable to unlink", "permission denied", "invalid argument",
    "being used by another process", "unable to create file",
)
PATRONES_SIN_IDENTIDAD = ("please tell me who you are", "user.email", "user.name")

MENSAJE_EXCEL_ABIERTO = "Cerrá el Excel y tocá Sincronizar ahora."


@dataclass
class ResultadoSync:
    """estado: ok / sin_cambios / sin_conexion / conflicto / error / desactivado.
    mensaje: texto en español para mostrar al usuario.
    trajo_cambios: True si bajar_cambios() actualizó archivos de esta PC.
    archivos: archivos que se bajaron o que están en conflicto.
    pendiente_subir: True si bajar_cambios() vio cambios de esta PC sin subir.
    detalle: salida técnica de git (para el área de salida, si hace falta)."""
    estado: str
    mensaje: str
    trajo_cambios: bool = False
    archivos: list = field(default_factory=list)
    pendiente_subir: bool = False
    detalle: str = ""


class _GitNoEncontrado(Exception):
    pass


class _GitTimeout(Exception):
    def __init__(self, args):
        super().__init__(" ".join(args))
        self.args_git = args


class _LockOcupado(Exception):
    pass


# ---------------------------------------------------------------------------
# Ejecución de git
# ---------------------------------------------------------------------------
def _git(raiz, *args, entrada=None):
    """Corre `git <args>` en `raiz` y devuelve el CompletedProcess.
    Lanza _GitNoEncontrado si git no está instalado y _GitTimeout si tarda
    más de TIMEOUT_SEGUNDOS."""
    env = os.environ.copy()
    env["GIT_TERMINAL_PROMPT"] = "0"
    env["GCM_INTERACTIVE"] = "never"
    env["LC_ALL"] = "C"  # mensajes de git en inglés, para poder clasificarlos
    opciones = {"input": entrada} if entrada is not None else {"stdin": subprocess.DEVNULL}
    try:
        return subprocess.run(
            ["git", "-c", "core.quotepath=false", *args],
            cwd=str(raiz), capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=TIMEOUT_SEGUNDOS, env=env,
            creationflags=CREATE_NO_WINDOW, **opciones,
        )
    except FileNotFoundError as e:
        raise _GitNoEncontrado() from e
    except subprocess.TimeoutExpired as e:
        raise _GitTimeout(args) from e


def _hay_git_corriendo():
    """True si hay algún git.exe en ejecución (tasklist, sin ventana)."""
    if os.name != "nt":
        return False
    try:
        salida = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq git.exe", "/NH", "/FO", "CSV"],
            capture_output=True, timeout=15, creationflags=CREATE_NO_WINDOW,
        ).stdout.decode("latin-1", errors="replace")
    except (OSError, subprocess.TimeoutExpired):
        return True  # ante la duda, no se borra el lock
    return '"git.exe"' in salida.lower()


def _git_con_lock(raiz, *args, entrada=None):
    """Como _git(), para comandos que escriben el índice. Si fallan por un
    .git/index.lock abandonado (más de 2 minutos y ningún git.exe
    corriendo), lo borra y reintenta una vez. Si el lock parece estar en
    uso, lanza _LockOcupado con un mensaje claro."""
    resultado = _git(raiz, *args, entrada=entrada)
    if resultado.returncode == 0 or "index.lock" not in resultado.stderr:
        return resultado

    lock = Path(raiz) / ".git" / "index.lock"
    if lock.exists():
        antiguedad = time.time() - lock.stat().st_mtime
        if antiguedad < ANTIGUEDAD_MINIMA_LOCK:
            raise _LockOcupado(
                "Otro programa (por ejemplo VS Code) está usando git en este momento "
                "(.git/index.lock es reciente). Esperá un par de minutos y tocá "
                "Sincronizar ahora."
            )
        if _hay_git_corriendo():
            raise _LockOcupado(
                "Hay otro proceso git en ejecución y existe .git/index.lock. Cerrá los "
                "programas que usen git (por ejemplo VS Code) y tocá Sincronizar ahora."
            )
        try:
            lock.unlink()
        except OSError as e:
            raise _LockOcupado(f"No se pudo borrar .git/index.lock abandonado: {e}") from e
    return _git(raiz, *args, entrada=entrada)


# ---------------------------------------------------------------------------
# Consultas
# ---------------------------------------------------------------------------
def verificar_disponible(raiz=RAIZ_PROYECTO):
    """Devuelve None si se puede sincronizar, o un ResultadoSync con estado
    'desactivado' explicando por qué no (git no instalado, la carpeta no es
    un repositorio, otra rama, sin remoto)."""
    try:
        top = _git(raiz, "rev-parse", "--show-toplevel")
        if top.returncode != 0 or Path(top.stdout.strip()).resolve() != Path(raiz).resolve():
            return ResultadoSync(
                "desactivado",
                "Sincronización desactivada: la carpeta del proyecto no es un repositorio git.",
            )
        rama = _git(raiz, "rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
        if rama != RAMA:
            return ResultadoSync(
                "desactivado",
                f"Sincronización desactivada: el repositorio está en la rama '{rama}' "
                f"y la sincronización solo trabaja sobre '{RAMA}'.",
            )
        if _git(raiz, "remote", "get-url", REMOTO).returncode != 0:
            return ResultadoSync(
                "desactivado",
                f"Sincronización desactivada: el repositorio no tiene el remoto '{REMOTO}'.",
            )
    except _GitNoEncontrado:
        return ResultadoSync(
            "desactivado",
            "Sincronización desactivada: git no está instalado en esta PC.",
        )
    return None


def _contar_diferencias(raiz):
    """(commits locales sin subir, commits de GitHub que faltan acá),
    respecto de la última copia conocida de origin/main."""
    salida = _git(raiz, "rev-list", "--left-right", "--count", f"HEAD...{REMOTO}/{RAMA}")
    if salida.returncode != 0:
        raise RuntimeError(salida.stderr.strip() or "No se pudo comparar con GitHub.")
    locales, remotos = salida.stdout.split()
    return int(locales), int(remotos)


def _cambios_sin_commitear(raiz):
    """Lista de archivos con cambios locales sin commitear (incluye nuevos)."""
    salida = _git(raiz, "status", "--porcelain")
    if salida.returncode != 0:
        raise RuntimeError(salida.stderr.strip() or "No se pudo leer el estado del repositorio.")
    return [linea[3:].strip().strip('"') for linea in salida.stdout.splitlines() if linea.strip()]


def _nombres_diff(raiz, rango):
    salida = _git(raiz, "diff", "--name-only", rango)
    return [l for l in salida.stdout.splitlines() if l.strip()] if salida.returncode == 0 else []


def _archivos_bloqueados(raiz, archivos):
    """Archivos que no se pueden abrir para escritura (en Windows, el caso
    típico es el .xlsx abierto en Excel). Abrir en 'r+b' no modifica nada."""
    bloqueados = []
    for relativo in archivos:
        ruta = Path(raiz) / relativo
        if not ruta.is_file():
            continue
        try:
            with open(ruta, "r+b"):
                pass
        except OSError:
            bloqueados.append(relativo)
    return bloqueados


def _contiene(texto, patrones):
    texto = texto.lower()
    return any(p in texto for p in patrones)


def _resultado_falla_remota(salida, accion):
    """Clasifica la falla de un fetch / pull / push."""
    stderr = (salida.stderr or "").strip()
    if _contiene(stderr, PATRONES_AUTENTICACION):
        return ResultadoSync(
            "error",
            f"GitHub rechazó las credenciales al {accion}. Revisá la credencial guardada "
            "de GitHub en esta PC (Administrador de credenciales de Windows).",
            detalle=stderr,
        )
    if _contiene(stderr, PATRONES_SIN_CONEXION):
        return ResultadoSync(
            "sin_conexion", "Sin conexión: trabajando con la versión local.", detalle=stderr
        )
    return ResultadoSync("error", f"git falló al {accion}.", detalle=stderr)


def _lista_corta(archivos, maximo=5):
    if len(archivos) <= maximo:
        return ", ".join(archivos)
    return ", ".join(archivos[:maximo]) + f" (+{len(archivos) - maximo} más)"


def mensaje_commit(operacion, detalle=""):
    """'Auto: <operación> – <detalle> – dd/mm/aaaa HH:MM – PC <COMPUTERNAME>'."""
    partes = [f"Auto: {operacion}"]
    if detalle:
        partes.append(detalle)
    partes.append(datetime.datetime.now().strftime("%d/%m/%Y %H:%M"))
    partes.append(f"PC {os.environ.get('COMPUTERNAME', 'desconocida')}")
    return " – ".join(partes)


# ---------------------------------------------------------------------------
# Operaciones públicas
# ---------------------------------------------------------------------------
def bajar_cambios(raiz=RAIZ_PROYECTO):
    """git fetch y, si GitHub tiene algo nuevo y esta PC no tiene cambios
    propios, git pull --ff-only. Nunca combina nada automáticamente."""
    try:
        return _bajar_cambios(raiz)
    except _GitNoEncontrado:
        return ResultadoSync(
            "desactivado", "Sincronización desactivada: git no está instalado en esta PC."
        )
    except _GitTimeout as e:
        if e.args_git and e.args_git[0] in ("fetch", "pull"):
            return ResultadoSync(
                "sin_conexion",
                f"Sin conexión: GitHub no respondió en {TIMEOUT_SEGUNDOS} segundos; "
                "trabajando con la versión local.",
            )
        return ResultadoSync("error", f"git tardó más de {TIMEOUT_SEGUNDOS} s en «git {e}».")
    except _LockOcupado as e:
        return ResultadoSync("error", str(e))
    except Exception as e:  # nunca debe romper la interfaz
        return ResultadoSync("error", f"Error inesperado al bajar cambios: {e}")


def _bajar_cambios(raiz):
    no_disponible = verificar_disponible(raiz)
    if no_disponible is not None:
        return no_disponible

    fetch = _git(raiz, "fetch", REMOTO)
    if fetch.returncode != 0:
        return _resultado_falla_remota(fetch, "consultar GitHub")

    locales, remotos = _contar_diferencias(raiz)
    sin_commitear = _cambios_sin_commitear(raiz)

    if remotos == 0:
        pendientes = []
        if locales:
            pendientes.append(f"{locales} commit(s) de esta PC sin subir")
        if sin_commitear:
            pendientes.append(f"{len(sin_commitear)} archivo(s) con cambios sin subir")
        mensaje = "Sincronizado con GitHub: no hay cambios nuevos para bajar."
        if pendientes:
            mensaje += " Pendiente de subir: " + "; ".join(pendientes) + "."
        return ResultadoSync("sin_cambios", mensaje, pendiente_subir=bool(pendientes))

    remotos_archivos = _nombres_diff(raiz, f"HEAD...{REMOTO}/{RAMA}")
    if locales or sin_commitear:
        locales_archivos = sin_commitear or _nombres_diff(raiz, f"{REMOTO}/{RAMA}...HEAD")
        causa = (
            "esta PC tiene commits que GitHub no tiene, y GitHub tiene commits que esta PC no tiene"
            if locales else
            "esta PC tiene cambios sin subir y mientras tanto GitHub recibió cambios de la otra PC"
        )
        return ResultadoSync(
            "conflicto",
            "Conflicto: hay cambios en las dos PC. No se hizo nada automático; hay que "
            "resolverlo a mano antes de seguir aplicando cambios.",
            archivos=sorted(set(locales_archivos) | set(remotos_archivos)),
            detalle=(
                f"Causa: {causa}.\n"
                f"Cambios de esta PC: {_lista_corta(locales_archivos, 10) or '(ninguno)'}\n"
                f"Cambios en GitHub: {_lista_corta(remotos_archivos, 10) or '(ninguno)'}"
            ),
        )

    bloqueados = _archivos_bloqueados(raiz, _nombres_diff(raiz, f"HEAD..{REMOTO}/{RAMA}"))
    if bloqueados:
        return ResultadoSync(
            "error",
            f"Hay cambios nuevos en GitHub, pero hay archivos en uso. {MENSAJE_EXCEL_ABIERTO}",
            archivos=bloqueados,
            detalle="Archivos en uso: " + ", ".join(bloqueados),
        )

    pull = _git_con_lock(raiz, "pull", "--ff-only", REMOTO, RAMA)
    if pull.returncode != 0:
        if _contiene(pull.stderr, PATRONES_ARCHIVO_EN_USO) and not _contiene(
            pull.stderr, PATRONES_AUTENTICACION
        ):
            return ResultadoSync(
                "error",
                f"No se pudieron actualizar los archivos porque alguno está en uso. "
                f"{MENSAJE_EXCEL_ABIERTO}",
                detalle=pull.stderr.strip(),
            )
        return _resultado_falla_remota(pull, "bajar los cambios")

    return ResultadoSync(
        "ok",
        f"Se bajaron {remotos} cambio(s) de GitHub: {_lista_corta(remotos_archivos)}.",
        trajo_cambios=True,
        archivos=remotos_archivos,
    )


def subir_cambios(mensaje, raiz=RAIZ_PROYECTO):
    """git add -A; si hay algo nuevo, commit con `mensaje`; después push si
    hay commits sin subir (incluidos los de corridas anteriores sin
    conexión). Si GitHub rechaza el push, el commit queda en esta PC."""
    try:
        return _subir_cambios(raiz, mensaje)
    except _GitNoEncontrado:
        return ResultadoSync(
            "desactivado", "Sincronización desactivada: git no está instalado en esta PC."
        )
    except _GitTimeout as e:
        if e.args_git and e.args_git[0] == "push":
            return ResultadoSync(
                "sin_conexion",
                f"Sin conexión: GitHub no respondió en {TIMEOUT_SEGUNDOS} segundos. Los "
                "cambios quedaron guardados en esta PC y se suben en la próxima sincronización.",
            )
        return ResultadoSync("error", f"git tardó más de {TIMEOUT_SEGUNDOS} s en «git {e}».")
    except _LockOcupado as e:
        return ResultadoSync("error", str(e))
    except Exception as e:  # nunca debe romper la interfaz
        return ResultadoSync("error", f"Error inesperado al subir cambios: {e}")


def _subir_cambios(raiz, mensaje):
    no_disponible = verificar_disponible(raiz)
    if no_disponible is not None:
        return no_disponible

    add = _git_con_lock(raiz, "add", "-A")
    if add.returncode != 0:
        return ResultadoSync("error", "git falló al preparar los cambios (git add).",
                             detalle=add.stderr.strip())

    diff = _git(raiz, "diff", "--cached", "--quiet")
    if diff.returncode not in (0, 1):
        return ResultadoSync("error", "No se pudo leer qué cambios hay para subir.",
                             detalle=diff.stderr.strip())

    archivos = []
    if diff.returncode == 1:
        archivos = _nombres_diff(raiz, "--cached")
        commit = _git_con_lock(raiz, "commit", "-F", "-", entrada=mensaje)
        if commit.returncode != 0:
            if _contiene(commit.stderr, PATRONES_SIN_IDENTIDAD):
                return ResultadoSync(
                    "error",
                    "git no sabe con qué nombre firmar los cambios en esta PC. Configuralo "
                    "una vez en PowerShell con: git config --global user.name \"Tu nombre\" "
                    "y git config --global user.email \"tu@correo\".",
                    detalle=commit.stderr.strip(),
                )
            return ResultadoSync("error", "git falló al guardar los cambios (git commit).",
                                 detalle=(commit.stderr or commit.stdout).strip())

    locales, _ = _contar_diferencias(raiz)
    if locales == 0:
        return ResultadoSync("sin_cambios", "No había cambios para subir a GitHub.")

    push = _git(raiz, "push", REMOTO, RAMA)
    if push.returncode != 0:
        stderr = push.stderr.strip()
        if _contiene(stderr, PATRONES_SIN_CONEXION) and not _contiene(stderr, PATRONES_AUTENTICACION):
            return ResultadoSync(
                "sin_conexion",
                "Sin conexión: los cambios quedaron guardados en esta PC y se suben en "
                "la próxima sincronización.",
                archivos=archivos, detalle=stderr,
            )
        if _contiene(stderr, PATRONES_RECHAZO_PUSH):
            return ResultadoSync(
                "conflicto",
                "Conflicto: los cambios quedaron guardados en esta PC, pero GitHub tiene "
                "algo más nuevo (de la otra PC). Hay que resolverlo a mano; no se forzó nada.",
                archivos=archivos, detalle=stderr,
            )
        resultado = _resultado_falla_remota(push, "subir los cambios")
        resultado.mensaje += " Los cambios quedaron guardados en esta PC."
        resultado.archivos = archivos
        return resultado

    detalle_archivos = f": {_lista_corta(archivos)}" if archivos else ""
    return ResultadoSync(
        "ok",
        f"Se subieron {locales} commit(s) a GitHub{detalle_archivos}.",
        archivos=archivos,
    )
