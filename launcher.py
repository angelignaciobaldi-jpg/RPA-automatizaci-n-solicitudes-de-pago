# -*- coding: utf-8 -*-
"""
Lanzador (bootstrapper) del RPA de Solicitudes de Pago.

Al iniciar, descarga del repositorio público de GitHub la versión MÁS RECIENTE
del código (.py) y la ejecuta. Así la aplicación se mantiene actualizada con
solo hacer 'git push' al repo, SIN tener que reconstruir el ejecutable.

El navegador (Chromium), Python y las dependencias siguen empaquetados dentro
de este .exe. Si no hay internet, usa la última copia descargada o, en su
defecto, la copia de respaldo empaquetada (carpeta 'codigo_base').
"""

import os
import sys
import json
import shutil
import urllib.request

# --- Imports "ancla" -------------------------------------------------------
# El código real (app.py / sipp_rpa.py) se descarga e importa en tiempo de
# ejecución, así que PyInstaller no "ve" esas dependencias por análisis
# estático. Las importamos aquí para que SÍ queden empaquetadas en el .exe.
import csv          # noqa: F401
import re           # noqa: F401
import time         # noqa: F401
import logging      # noqa: F401
import threading    # noqa: F401
import queue        # noqa: F401
import unicodedata  # noqa: F401
import datetime     # noqa: F401
import tkinter      # noqa: F401
from tkinter import ttk, filedialog, messagebox  # noqa: F401
try:
    from playwright.sync_api import sync_playwright  # noqa: F401
except Exception:
    pass
# ---------------------------------------------------------------------------

REPO = "angelignaciobaldi-jpg/RPA-automatizaci-n-solicitudes-de-pago"
RAMA = "main"
API_URL = f"https://api.github.com/repos/{REPO}/contents?ref={RAMA}"
RAW_URL = f"https://raw.githubusercontent.com/{REPO}/{RAMA}"

# Módulos que se ejecutan (respaldo si la API de GitHub no responde).
MODULOS = ["app.py", "sipp_rpa.py", "config.py", "mapeos.py", "validar_csv.py"]


def ruta_recurso(rel):
    """Ruta a un recurso empaquetado (usa sys._MEIPASS en el .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def carpeta_codigo():
    """Carpeta local (escribible) donde vive el código que se ejecuta."""
    base = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
                        "RPA_SIPP", "codigo")
    os.makedirs(base, exist_ok=True)
    return base


def _http_get(url, timeout=20):
    req = urllib.request.Request(url, headers={"User-Agent": "RPA-SIPP-Launcher"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def _listar_py_del_repo():
    """Lista los .py de la raíz del repo (para detectar archivos NUEVOS)."""
    datos = json.loads(_http_get(API_URL).decode("utf-8"))
    return [it["name"] for it in datos
            if it.get("type") == "file" and it.get("name", "").endswith(".py")]


def actualizar_desde_github(dest):
    """Descarga los .py del repo a 'dest'. Devuelve cuántos cambiaron.
    Lanza excepción si no se pudo conectar a GitHub."""
    try:
        archivos = _listar_py_del_repo()
    except Exception:
        archivos = list(MODULOS)  # sin API: al menos baja los módulos conocidos
    if not archivos:
        archivos = list(MODULOS)
    cambiados = 0
    errores = 0
    for nombre in archivos:
        try:
            contenido = _http_get(f"{RAW_URL}/{nombre}")
            ruta = os.path.join(dest, nombre)
            previo = open(ruta, "rb").read() if os.path.exists(ruta) else None
            if previo != contenido:
                with open(ruta, "wb") as f:
                    f.write(contenido)
                cambiados += 1
        except Exception:
            errores += 1
    # Si TODO falló (sin internet), avisa al llamador para usar respaldo.
    if errores and errores == len(archivos):
        raise RuntimeError("No se pudo descargar el código desde GitHub.")
    return cambiados


def copiar_respaldo(dest):
    """Copia la copia empaquetada ('codigo_base') a 'dest' (respaldo offline)."""
    base = ruta_recurso("codigo_base")
    if not os.path.isdir(base):
        return
    for nombre in os.listdir(base):
        try:
            shutil.copy2(os.path.join(base, nombre), os.path.join(dest, nombre))
        except Exception:
            pass


def _ejecutar_app(dest):
    if dest not in sys.path:
        sys.path.insert(0, dest)
    for m in ("app", "sipp_rpa", "config", "mapeos", "validar_csv"):
        sys.modules.pop(m, None)
    import app
    app.App().mainloop()


def main():
    dest = carpeta_codigo()
    cambiados = 0
    try:
        cambiados = actualizar_desde_github(dest)
    except Exception:
        cambiados = 0  # sin internet o GitHub caído: seguimos con lo que haya

    # Si no hay código (1ª vez sin internet), usa el respaldo empaquetado.
    if not os.path.exists(os.path.join(dest, "app.py")):
        copiar_respaldo(dest)

    os.environ["RPA_ACTUALIZADOS"] = str(cambiados)

    try:
        _ejecutar_app(dest)
    except Exception:
        # El código descargado falló (descarga corrupta / commit roto):
        # restaura el respaldo empaquetado y reintenta.
        copiar_respaldo(dest)
        os.environ["RPA_ACTUALIZADOS"] = "0"
        _ejecutar_app(dest)


if __name__ == "__main__":
    main()
