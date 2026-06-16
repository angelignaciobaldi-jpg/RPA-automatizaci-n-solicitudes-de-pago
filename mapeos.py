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

# CSV/OCR "BANCOS"  ->  texto EXACTO en la lista "Banco" de SIPP (modal Cuenta
# Bancaria). Incluye los nombres que produce el OCR (BBVA, Citibanamex, Spin...).
BANCOS = {
    "BANAMEX": "BANAMEX",
    "CITIBANAMEX": "BANAMEX",
    "BBVA": "BBVA BANCOMER",
    "BBVA BANCOMER": "BBVA BANCOMER",
    "SANTANDER": "SANTANDER",
    "HSBC": "HSBC",
    "BANORTE": "BANORTE",
    "BANCOPPEL": "BANCOPPEL",
    "SPIN BY OXXO": "Spin by OXXO",
    "SPIN": "Spin by OXXO",
    "BANREGIO": "BANREGIO",
    "AZTECA": "AZTECA",
    "BANCO AZTECA": "AZTECA",
    "INBURSA": "INBURSA",
    "SCOTIABANK": "SCOTIABANK",
    "BAJIO": "BAJIO",
    "BANCO DEL BAJIO": "BAJIO",
    "CIBANCO": "CIBanco",
    "AFIRME": "AFIRME",
    "INVEX": "INVEX",
    "BANSI": "BANSI",
    "MIFEL": "MIFEL",
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
