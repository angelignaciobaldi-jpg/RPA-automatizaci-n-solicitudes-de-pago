# -*- coding: utf-8 -*-
"""
Puente OCR -> RPA
=================
Conecta el Extractor Bancario (OCR) con el RPA de Solicitudes de Pago:
  1) Corre el OCR sobre una lista de documentos (carátulas/estados de cuenta).
  2) Convierte cada resultado del OCR en una "fila" con el formato que consume
     el RPA (mismas claves normalizadas que usa sipp_rpa.campo()).

El OCR extrae: beneficiario, banco, CLABE y RFC. El MONTO, la EMPRESA y la
SUCURSAL NO vienen en la carátula bancaria; los completa el operador en la app.
La carátula de cada fila ES el propio documento analizado (se guarda su ruta en
la clave especial "_CARATULA").
"""

import os
import re

import extractor_bancario as ocr
import sipp_rpa

# Monto tipo "8,606.90" / "15,595.34" / "464.36".
_MONTO_RE = r"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{2})"


def _a_float(texto):
    try:
        return round(float(texto.replace("$", "").replace(" ", "").replace(",", "")), 2)
    except Exception:
        return None


def extraer_monto_recibo(texto):
    """Extrae el TOTAL NETO a pagar de un recibo de utilidades / nómina (CFDI).
    Estrategia robusta: TOTAL = SUB TOTAL − DESCUENTOS, validado contra el
    'TOTAL' explícito. (La 'cantidad con letra' suele traer errores de OCR.)
    Devuelve un string '1234.56' o '' si no lo encuentra."""
    t = texto or ""

    def buscar(etq):
        m = re.search(etq + r"[^\d$]{0,12}\$?\s*" + _MONTO_RE, t, re.I)
        return _a_float(m.group(1)) if m else None

    subtotal = buscar(r"sub\s*total")
    descuentos = buscar(r"descuentos?")
    total_exp = None
    for m in re.finditer(r"total[^\d$]{0,10}\$?\s*" + _MONTO_RE, t, re.I):
        if "sub" in t[max(0, m.start() - 6):m.start()].lower():
            continue
        total_exp = _a_float(m.group(1))   # el último (cerca de FIRMA/abajo)

    neto = None
    if subtotal is not None and descuentos is not None:
        neto = round(subtotal - descuentos, 2)
    if neto is not None and total_exp is not None:
        elegido = total_exp if abs(total_exp - neto) < 0.05 else neto
    else:
        elegido = total_exp if total_exp is not None else neto
    return f"{elegido:.2f}" if elegido is not None else ""

N = sipp_rpa.normalizar

# Extensiones de documentos que el OCR puede leer.
EXTENSIONES_OCR = tuple(ocr.SUPPORTED_EXTENSIONS)

# Moneda por defecto (la carátula no la trae).
MONEDA_DEFECTO = "Pesos (MXN)"


def extraer_documentos(rutas, on_progreso=None):
    """Corre el OCR sobre 'rutas' (carátulas + opcionalmente constancias SAT).
    Devuelve la lista de ExtractionResult de las carátulas (las constancias se
    usan solo para enriquecer/corregir nombre y RFC).

    on_progreso(i, total, nombre_archivo) se llama por cada documento.
    """
    rutas = [r for r in rutas if r and os.path.isfile(r)]
    total = len(rutas)
    resultados = []
    constancias = []
    for i, ruta in enumerate(rutas, 1):
        if on_progreso:
            on_progreso(i, total, os.path.basename(ruta))
        try:
            texto, _tipo, _notas = ocr.extract_text(ruta)
            if ocr._is_constancia(texto):
                constancias.append(ocr._parse_constancia(texto, ruta))
            else:
                resultados.append(ocr.parse_document(ruta))
        except Exception as e:
            # Resultado vacío marcado para revisión (no se cae el lote).
            r = ocr.ExtractionResult(
                file_path=ruta, source_type="?", extracted_text="",
                beneficiary_name="", account_number="", clabe="",
                bank_name="", rfc="", clabe_is_valid=False,
                notes=[f"Error al leer: {e}"])
            resultados.append(r)
    # Enriquecer con las constancias fiscales (corrige nombre/RFC por RFC o
    # similitud de nombre).
    if constancias:
        for r in resultados:
            try:
                ocr._enrich_with_constancias(r, constancias)
            except Exception:
                pass
    return resultados


# RFC genéricos (público en general): equivalen a "sin RFC", se sustituyen.
_RFC_GENERICOS = {"XAXX010101000", "XEXX010101000"}

# Confusiones típicas del OCR, corregidas según la POSICIÓN dentro del RFC:
#   - posiciones de LETRAS (las primeras 3-4): dígito -> letra parecida
#   - posiciones de la FECHA (6 dígitos): letra -> dígito parecido
_OCR_A_LETRA = {"0": "O", "1": "I", "5": "S", "8": "B", "6": "G", "2": "Z", "4": "A"}
_OCR_A_DIGITO = {"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "G": "6",
                 "Z": "2", "Q": "0", "D": "0", "A": "4"}


def _arreglar_rfc(rfc):
    """Limpia y corrige por estructura un RFC leído por OCR. Devuelve el RFC en
    mayúsculas si queda con forma válida (RFC físico/moral), o '' si no."""
    rfc = re.sub(r"[^A-Z0-9Ñ&]", "", (rfc or "").upper())
    if len(rfc) not in (12, 13):
        return ""
    ini = len(rfc) - 9                       # 4 letras (física) o 3 (moral)
    letras = "".join(_OCR_A_LETRA.get(c, c) for c in rfc[:ini])
    fecha = "".join(_OCR_A_DIGITO.get(c, c) for c in rfc[ini:ini + 6])
    homoclave = rfc[ini + 6:]                # no se toca (es alfanumérica)
    arreglado = letras + fecha + homoclave
    if re.match(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{2,3}$", arreglado):
        return arreglado
    return ""


def _datos_recibo(texto):
    """De la sección COLABORADOR de un recibo CFDI saca (nombre, codigo, rfc).
    Estructura: 'No. 2917 NOMBRE BRANDON ... MORALES TIPO COMPROBANTE / RFC ...
    MAMB040709P45 NUM. SEG. SOCIAL ...'. El nombre va entre 'NOMBRE' y la
    siguiente etiqueta; el RFC, justo antes de 'NUM. SEG(URO) SOCIAL'."""
    t = texto or ""
    nombre = codigo = rfc = ""
    # Código: "No. 2917 NOMBRE" (tolera basura del OCR entre 'No.' y el número,
    # ej. "No. «68 NOMBRE").
    m = re.search(r"No\.?\s*[^\d\n]{0,4}(\d{1,6})\s+NOMBRE", t, re.I)
    if m:
        codigo = m.group(1)
    # Nombre: lo que sigue a 'NOMBRE' hasta la etiqueta 'TIPO COMPROBANTE'. En
    # escaneos malos esa etiqueta sale garabateada (ej. "TEOICOMEROBENTE"), así
    # que también cortamos en fin de línea / RFC / CURP. El operador puede
    # corregirlo y, si queda vacío, el llamador usa el nombre del archivo.
    m = re.search(
        r"NOMBRE\s+([A-ZÑÁÉÍÓÚ.\- ]{4,60}?)\s+(?:TIPO|RFC|CURP|COMPROB|T[A-Z]*COM|\n)",
        t, re.I)
    if m:
        nombre = re.sub(r"\s+", " ", m.group(1)).strip().upper()
    # RFC: el token (12-13 car.) inmediatamente ANTES de 'NUM. SEG. SOCIAL'
    # (el NSS siempre le sigue). Respaldo: tras la etiqueta 'RFC'.
    m = re.search(r"([A-Z0-9Ñ&]{12,13})\s+NUM[.\s]*SEG", t, re.I)
    if not m:
        m = re.search(r"\bRFC\b[\s:]*([A-Z0-9Ñ&]{12,13})\b", t, re.I)
    if m:
        rfc = _arreglar_rfc(m.group(1))
    return nombre, codigo, rfc


def extraer_datos_vobo(rutas, on_progreso=None):
    """OCR de los recibos de utilidades (Vo.Bo.): saca el MONTO neto + el
    nombre y código del colaborador (sección COLABORADOR del CFDI).
    Devuelve {clave_archivo_normalizado: {'monto','nombre','codigo'}}."""
    indice = {}
    rutas = [r for r in (rutas or []) if r and os.path.isfile(r)]
    total = len(rutas)
    for i, ruta in enumerate(rutas, 1):
        if on_progreso:
            on_progreso(i, total, os.path.basename(ruta))
        try:
            texto, _t, _n = ocr.extract_text(ruta)
        except Exception:
            continue
        monto = extraer_monto_recibo(texto)
        nombre, codigo, rfc = _datos_recibo(texto)
        base = os.path.splitext(os.path.basename(ruta))[0].strip()
        indice[sipp_rpa.normalizar(base)] = {
            "monto": monto, "nombre": nombre, "codigo": codigo, "rfc": rfc,
            # Respaldo: el archivo está nombrado por el colaborador.
            "nombre_archivo": base.upper()}
    return indice


def aplicar_datos_vobo(filas, indice):
    """Aplica a cada fila los datos de su recibo: rellena el MONTO (si falta),
    corrige el NOMBRE del colaborador y rellena su CÓDIGO (si falta). Empareja
    por nombre de archivo de la carátula o por el nombre del colaborador.
    Devuelve (montos_aplicados, nombres_aplicados, rfc_aplicados)."""
    if not indice:
        return 0, 0, 0
    N = sipp_rpa.normalizar
    n_monto = n_nombre = n_rfc = 0
    for f in filas:
        datos = None
        car = f.get("_CARATULA")
        if car:
            datos = indice.get(N(os.path.splitext(os.path.basename(car))[0]))
        if datos is None:
            nombre_f = sipp_rpa.campo(f, "EX-COLABORADOR (DESCRIPCIÓN)",
                                      "NOMBRE DE CUENTA")
            if nombre_f:
                datos = indice.get(N(nombre_f))
        if not datos:
            continue
        if datos.get("monto") and not sipp_rpa.limpiar_monto(
                sipp_rpa.campo_monto(f)):
            f[N("MONTO A PAGAR")] = datos["monto"]
            n_monto += 1
        # Nombre del colaborador: el de la sección COLABORADOR del recibo; si el
        # OCR no lo pudo leer, el del nombre del archivo (que es el colaborador).
        nombre = datos.get("nombre") or datos.get("nombre_archivo")
        if nombre:
            f[N("EX-COLABORADOR (DESCRIPCION)")] = nombre
            f[N("NOMBRE DE CUENTA")] = nombre
            n_nombre += 1
        if datos.get("codigo") and not sipp_rpa.campo(f, "NUM. COLABORADOR"):
            f[N("NUM. COLABORADOR")] = datos["codigo"]
        # RFC: solo si la carátula no trajo uno real (vacío o genérico XAXX/XEXX).
        if datos.get("rfc"):
            rfc_actual = sipp_rpa.campo(f, "RFC").strip().upper()
            if (not rfc_actual) or (rfc_actual in _RFC_GENERICOS):
                f[N("RFC")] = datos["rfc"]
                n_rfc += 1
    return n_monto, n_nombre, n_rfc


def resultado_a_fila(result, empresa="", sucursal="", monto="",
                     moneda=MONEDA_DEFECTO):
    """Convierte un ExtractionResult del OCR en una 'fila' del RPA.
    El nombre y RFC oficiales (de la Constancia Fiscal) tienen prioridad."""
    nombre = (getattr(result, "csf_name", "") or result.beneficiary_name or "").strip()
    rfc = (getattr(result, "csf_rfc", "") or result.rfc or "").strip()
    fila = {
        N("EX-COLABORADOR (DESCRIPCION)"): nombre,
        N("NOMBRE DE CUENTA"): nombre,
        N("EMPRESA"): empresa,
        N("SUCURSAL"): sucursal,
        N("BANCOS"): result.bank_name,
        N("CLAVE INTERBANCARIA"): result.clabe,
        N("RFC"): rfc,
        N("MONTO A PAGAR"): str(monto or ""),
        N("MONEDA"): moneda,
        # Marca interna: la carátula de esta fila es el documento analizado.
        "_CARATULA": result.file_path,
        # Datos informativos del OCR (para la tabla / validación).
        "_OCR_ESTADO": result.status,
        "_OCR_NOTAS": result.report_notes,
        "_OCR_CLABE_OK": result.clabe_is_valid,
    }
    return fila
