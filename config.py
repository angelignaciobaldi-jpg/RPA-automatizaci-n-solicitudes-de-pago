# -*- coding: utf-8 -*-
"""
Configuración del RPA de Solicitudes de Pago (SIPP - Petroil).

Edita aquí las rutas, credenciales y banderas de seguridad.
NO subas este archivo con la contraseña real a ningún repositorio.
"""

import os

# --------------------------------------------------------------------------- #
#  SITIO
# --------------------------------------------------------------------------- #
# Ambiente de trabajo: "PRUEBAS" (stage) o "PRODUCCION".
# Cambia SOLO esta línea para alternar.
AMBIENTE = "PRUEBAS"

URLS = {
    "PRODUCCION": "https://sipp.petroil.com.mx/login.html",
    "PRUEBAS": "https://preprod.sipp.petroil.dev/login.html",
}
URL_LOGIN = URLS[AMBIENTE]

# Credenciales: NO se guardan en el código (por seguridad, para no distribuirlas
# en el .exe). La aplicación las pide al operador cada vez. Para la línea de
# comandos / scripts, se pueden definir variables de entorno SIPP_USUARIO y
# SIPP_CONTRASENA; si no, el CLI las pedirá por teclado.
USUARIO = os.environ.get("SIPP_USUARIO", "")
CONTRASENA = os.environ.get("SIPP_CONTRASENA", "")

# --------------------------------------------------------------------------- #
#  ARCHIVOS
# --------------------------------------------------------------------------- #
# Carpeta base = donde está este archivo.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# CSV con los datos de los pagos.
ARCHIVO_CSV = os.path.join(BASE_DIR, "BASE DE DATOS LISTADO 1.csv")

# Carpeta donde están los archivos de CARÁTULA (uno por colaborador,
# nombrados por el nombre del colaborador, ej. "ANGEL MORENO AYALA.pdf").
CARPETA_CARATULAS = os.path.join(BASE_DIR, "CARATULAS")
# Alternativa: lista de archivos de carátula sueltos (rutas) que el operador
# adjunta uno por uno en la app. Se busca en la carpeta Y en esta lista.
ARCHIVOS_CARATULAS = []

# Extensiones válidas de carátula (en orden de preferencia).
EXT_CARATULA = [".pdf", ".jpg", ".jpeg", ".png"]

# Carpeta donde se guardan los logs y capturas de error.
CARPETA_LOGS = os.path.join(BASE_DIR, "logs")

# --------------------------------------------------------------------------- #
#  BANDERAS DE SEGURIDAD / EJECUCIÓN
# --------------------------------------------------------------------------- #
# Si True, abre Chrome visible (recomendado mientras probamos).
NAVEGADOR_VISIBLE = True

# Cuántos registros procesar. Pon 1 para la primera prueba; None = todos.
MAX_REGISTROS = 1

# Si True, el robot llena TODO el formulario pero NO da clic en "Guardar":
# se detiene para que tú revises a mano. Déjalo en True hasta validar.
PAUSAR_ANTES_DE_GUARDAR = True

# Si True, tras guardar y adjuntar el Vo.Bo., también da clic en
# "Solicitar Autorización" (envía la solicitud a aprobación).
SOLICITAR_AUTORIZACION = True

# Tiempo máximo de espera por elemento (milisegundos). 60s para tolerar
# conexiones/equipos lentos y los catálogos grandes de PRODUCCIÓN.
TIMEOUT_MS = 60000

# --------------------------------------------------------------------------- #
#  VALORES FIJOS DEL FLUJO (ya definidos contigo)
# --------------------------------------------------------------------------- #
TIPO_PAGO = "Pago Extraordinario"          # solicitudPago.ID_TIPOPAGO
TIPO_BENEFICIARIO = "Acreedor"             # solicitudPago.ID_TIPOPAGOEXTRAORDINARIO
BENEFICIARIO_NO_REGISTRADO = True          # marcar checkbox "No Registrado"

# Forma de Pago: Cheque / Transferencia / Efectivo / Linea de Captura
FORMA_DE_PAGO = "Transferencia"
# Tipo de Gasto: Deducible / No Deducible / Deducible SF / Contribución
TIPO_DE_GASTO = "No Deducible"
# Concepto de pago donde va el importe (pestaña 'Conceptos de Pago').
# OJO: el catálogo de conceptos difiere entre PRODUCCION y PRUEBAS(stage).
CONCEPTO_PAGO = "PAGO PTU"
# Correo a registrar SOLO cuando el acreedor es nuevo (no está en el catálogo).
CORREO_ACREEDOR = "acreedoresdeudores@petroil.com.mx"

# Carpeta de documentos Vo.Bo. del Depto. de Compras (uno por colaborador,
# nombrado por el nombre del colaborador, igual que las carátulas).
CARPETA_VOBO = os.path.join(BASE_DIR, "VOBO")
# Alternativa: lista de archivos de Vo.Bo. sueltos (rutas) adjuntados uno a uno.
ARCHIVOS_VOBO = []

# --------------------------------------------------------------------------- #
#  CONFIGURACIÓN DE SESIÓN (pantalla posterior al login)
# --------------------------------------------------------------------------- #
# Tras el login, SIPP pide elegir Empresa + Plaza(Sucursal) y dar Guardar.
CONFIGURAR_SESION = True

# Empresa de la sesión: SIEMPRE ASKE / Corporativo (el usuario configura su
# sesión en ASKE; la empresa/sucursal de cada solicitud se toma del CSV).
EMPRESA_SESION = "ASKE"

# Plaza/Sucursal de la sesión.
SUCURSAL_SESION = "Corporativo"
