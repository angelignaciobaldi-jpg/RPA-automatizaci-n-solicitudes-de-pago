# -*- coding: utf-8 -*-
"""
Valida el CSV con la MISMA lógica que usa el robot, sin abrir el navegador.
Reporta problemas de mapeo, CLABE, monto y carátulas faltantes.

Uso:  python validar_csv.py
"""

import os
import config
import mapeos
from sipp_rpa import leer_csv, campo, buscar_caratula, normalizar


def limpiar_monto(valor):
    """Quita $, comas y espacios -> número plano. Devuelve (limpio, ok)."""
    t = (valor or "").replace("$", "").replace(",", "").strip()
    try:
        float(t)
        return t, True
    except ValueError:
        return t, False


def main():
    filas = leer_csv(config.ARCHIVO_CSV)
    print(f"Registros leídos: {len(filas)}\n")

    hay_caratulas = os.path.isdir(config.CARPETA_CARATULAS)
    if not hay_caratulas:
        print(f"AVISO: aún no existe la carpeta de carátulas: {config.CARPETA_CARATULAS}\n")

    problemas = 0
    for i, fila in enumerate(filas, 2):  # fila 2 = primer registro en Excel
        nombre = campo(fila, "EX-COLABORADOR (DESCRIPCIÓN)", "NOMBRE DE CUENTA")
        errs = []

        # Mapeos
        for dic, col, etiqueta in [
            (mapeos.EMPRESAS, "EMPRESA", "EMPRESA"),
            (mapeos.SUCURSALES, "SUCURSAL", "SUCURSAL"),
            (mapeos.BANCOS, "BANCOS", "BANCO"),
            (mapeos.MONEDAS, "MONEDA", "MONEDA"),
            (mapeos.TRANSFERENCIAS, "TIPO DE TRANSFERENCIA", "TRANSFERENCIA"),
        ]:
            try:
                mapeos.traducir(dic, campo(fila, col), etiqueta)
            except ValueError as e:
                errs.append(str(e))

        # CLABE
        clabe = campo(fila, "CLAVE INTERBANCARIA")
        if not (clabe.isdigit() and len(clabe) == 18):
            errs.append(f"CLABE inválida ('{clabe}', longitud {len(clabe)}; debe ser 18 dígitos)")

        # Monto
        monto_raw = campo(fila, "MONTO")
        monto, ok = limpiar_monto(monto_raw)
        if not monto_raw:
            errs.append("MONTO vacío")
        elif not ok:
            errs.append(f"MONTO no numérico tras limpiar ('{monto_raw}' -> '{monto}')")

        # Carátula
        if hay_caratulas and not buscar_caratula(nombre):
            errs.append("Sin archivo de carátula que coincida con el nombre")

        if errs:
            problemas += 1
            print(f"[Fila {i}] {nombre}")
            for e in errs:
                print(f"    - {e}")

    print(f"\n=== {problemas} fila(s) con observaciones de {len(filas)} ===")


if __name__ == "__main__":
    main()
