# -*- coding: utf-8 -*-
"""
Mapeo entre los valores del CSV y el TEXTO EXACTO que aparece en cada lista
desplegable del formulario de SIPP.

Importante: las listas de SIPP usan el plugin "chosen" y se seleccionan por el
TEXTO visible de la opción, no por un id. Por eso aquí traducimos el valor del
CSV al texto tal cual aparece en el sistema.

Si algún valor del CSV no está en estos diccionarios, el robot lo registrará
como error en el log para que lo agregues aquí.
"""

# CSV "EMPRESA"  ->  texto en la lista "Empresa" del formulario
EMPRESAS = {
    "ABASTECEDORA": "Abastecedora",
    "ASAMAZ": "Asamaz",
    "ASKE": "Aske",
    "PETROPLAZAS": "Petroplazas",
    "PETROSMART": "Petro Smart",
}

# CSV "SUCURSAL"  ->  texto en la lista "Sucursal"
SUCURSALES = {
    "CORPORATIVO": "Corporativo",
}

# CSV "BANCOS"  ->  texto en la lista "Banco" (modal Cuenta Bancaria)
BANCOS = {
    "BANAMEX": "BANAMEX",
    "BBVA": "BBVA BANCOMER",
    "SANTANDER": "SANTANDER",
    "HSBC": "HSBC",
    "BANORTE": "BANORTE",
    "BANCOPPEL": "BANCOPPEL",
    "SPIN BY OXXO": "Spin by OXXO",
}

# CSV "MONEDA"  ->  texto en la lista "Moneda"
MONEDAS = {
    "PESOS (MXN)": "Pesos (MXN)",
    "DOLAR (USD)": "Dolar (USD)",
    "EURO": "Euro",
}

# CSV "TIPO DE TRANSFERENCIA"  ->  texto en la lista "Tipo Transferencia"
TRANSFERENCIAS = {
    "SPEI": "SPEI",
    "TEF": "TEF",
    "MISMO BANCO": "Mismo Banco",
}


def traducir(diccionario, valor, nombre_campo):
    """Devuelve el texto del sistema para un valor del CSV (sin distinguir
    mayúsculas/espacios). Lanza ValueError si no existe el mapeo."""
    clave = (valor or "").strip().upper()
    if clave in diccionario:
        return diccionario[clave]
    raise ValueError(
        f"No hay mapeo para {nombre_campo}='{valor}'. "
        f"Agrégalo en mapeos.py. Opciones conocidas: {list(diccionario.values())}"
    )
