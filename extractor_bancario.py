from __future__ import annotations

import dataclasses
import csv
import difflib
import os
import re
import sys
import threading
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    from pypdf import PdfReader  # type: ignore
except Exception:
    try:
        from PyPDF2 import PdfReader  # type: ignore
    except Exception:
        PdfReader = None  # type: ignore

try:
    import fitz  # type: ignore
except Exception:
    fitz = None  # type: ignore

try:
    from PIL import Image, ImageOps, ImageFilter  # type: ignore
except Exception:
    Image = None  # type: ignore
    ImageOps = None  # type: ignore
    ImageFilter = None  # type: ignore

try:
    import pillow_heif  # type: ignore

    pillow_heif.register_heif_opener()
except Exception:
    pillow_heif = None  # type: ignore

try:
    import pytesseract  # type: ignore
except Exception:
    pytesseract = None  # type: ignore

try:
    from openpyxl import Workbook  # type: ignore
except Exception:
    Workbook = None  # type: ignore

TESSERACT_DEFAULT_PATHS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tesseract-OCR", "tesseract.exe"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "tesseract.exe"),
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".tif",
    ".tiff",
    ".heic",
    ".heif",
}

BANK_CODE_MAP = {
    "002": "Citibanamex",
    "012": "BBVA",
    "014": "Santander",
    "021": "HSBC",
    "030": "Banco del Bajio",
    "044": "Scotiabank",
    "058": "Banregio",
    "059": "Invex",
    "060": "Bansi",
    "072": "Banorte",
    "127": "Banco Azteca",
    "137": "BanCoppel",
    "143": "CIBanco",
    "036": "Inbursa",
    "042": "Mifel",
    "062": "Afirme",
    "728": "Spin",  # NVIO Pagos Mexico (Spin by OXXO y otras fintech)
}

BANK_NAME_ALIASES = [
    ("citibanamex", "Citibanamex"),
    ("banamex", "Citibanamex"),
    ("bbva", "BBVA"),
    ("bancomer", "BBVA"),
    ("santander", "Santander"),
    ("hsbc", "HSBC"),
    ("banco del bajio", "Banco del Bajio"),
    ("bajio", "Banco del Bajio"),
    ("scotiabank", "Scotiabank"),
    ("banregio", "Banregio"),
    ("invex", "Invex"),
    ("bansi", "Bansi"),
    ("banorte", "Banorte"),
    ("banco azteca", "Banco Azteca"),
    ("azteca", "Banco Azteca"),
    ("cibanco", "CIBanco"),
    ("autofin", "Banco Autofin"),
    ("bancoppel", "BanCoppel"),
    ("coppel", "BanCoppel"),
    ("inbursa", "Inbursa"),
    ("afirme", "Afirme"),
    ("spin", "Spin"),
    ("compropago", "Spin"),
    ("mercantil del norte", "Banorte"),
]

# Etiqueta para "no indicar banco" (se detecta solo). bank_hint == "" significa auto.
AUTO_BANK_LABEL = "Deteccion automatica"

# Bancos que el usuario puede elegir al cargar las caratulas.
SELECTABLE_BANKS = [
    AUTO_BANK_LABEL,
    "BBVA",
    "Citibanamex",
    "Banorte",
    "Santander",
    "HSBC",
    "Scotiabank",
    "Banregio",
    "Banco Azteca",
    "BanCoppel",
    "Inbursa",
    "Afirme",
    "Invex",
    "Bansi",
    "Banco del Bajio",
    "CIBanco",
    "Otro",
]

# Prefijo CLABE esperado por banco (para avisar si la seleccion no coincide).
BANK_EXPECTED_CLABE_PREFIX = {name: code for code, name in BANK_CODE_MAP.items()}

# RFC de las propias instituciones financieras. Aparecen en el encabezado de los
# estados de cuenta y NO deben confundirse con el RFC del titular.
BANK_RFCS = {
    "BRM940216EQ6",  # Banco Regional (Banregio)
    "COM131212AI3",  # Compropago (Spin by OXXO)
    "BBA830831LJ2",  # BBVA Mexico
    "BNM840515VB1",  # Banco Nacional de Mexico (Citibanamex)
    "BMN930209927",  # Banco Mercantil del Norte (Banorte)
    "HMI950125KG8",  # HSBC Mexico
    "BMS9007158J9",  # Banco Santander Mexico
    "BSM970519DU8",  # Banco Santander Mexico (variante en estados de cuenta)
    "SIN960904468",  # Scotiabank Inverlat
}

# Marcadores de que un RFC pertenece a la institucion bancaria, no al titular.
_BANK_RFC_CONTEXT = (
    "institucion de banca",
    "banca multiple",
    "grupo financiero",
    "institucion de fondos",
    "banco regional",
)

CLABE_NUMBER_PATTERN = re.compile(r"(?<!\d)(?:\d[\s\-]?){17}\d(?!\d)")
ACCOUNT_PATTERN = re.compile(r"\b(?:\d[\d\s\-]{7,24}\d)\b")
OWNER_NAME_LABELS = [
    "nombre del beneficiario",
    "beneficiario",
    "nombre del titular",
    "titular de la cuenta",
    "titular de cuenta",
    "titular",
    "nombre del cuentahabiente",
    "cuentahabiente",
    "nombre del dueño de la cuenta",
    "dueño de la cuenta",
    "propietario de la cuenta",
    "a nombre de",
]

BENEFICIARY_NAME_LABELS = OWNER_NAME_LABELS + [
    "razon social",
    "nombre razon social",
    "nombre / razon social",
]

RFC_LABELS = [
    "rfc",
    "r f c",
    "r.f.c",
    "r.f.c.",
    "registro federal de contribuyentes",
    "registro federal",
    "clave del rfc",
    "datos fiscales",
    "datos de facturacion",
    "contribuyente",
]


@dataclasses.dataclass
class ExtractionResult:
    file_path: str
    source_type: str
    extracted_text: str
    beneficiary_name: str
    account_number: str
    clabe: str
    bank_name: str
    rfc: str
    clabe_is_valid: bool
    notes: list[str]
    csf_name: str = ""  # nombre tomado de la Constancia de Situacion Fiscal emparejada
    csf_rfc: str = ""   # RFC tomado de la Constancia de Situacion Fiscal emparejada

    @property
    def file_name(self) -> str:
        return os.path.basename(self.file_path)

    @property
    def csf_summary(self) -> str:
        """Resumen del CSF emparejado para validar su lectura en la tabla."""
        if not (self.csf_name or self.csf_rfc):
            return ""
        return f"{self.csf_name or 's/nombre'} - {self.csf_rfc or 's/RFC'}"

    @property
    def status(self) -> str:
        if self.clabe and self.clabe_is_valid:
            return "OK"
        if self.clabe and not self.clabe_is_valid:
            return "CLABE invalida"
        return "Revisar"

    @property
    def report_notes(self) -> str:
        """Notas directas para el reporte: solo los datos que no se encontraron."""
        faltantes: list[str] = []
        if not self.beneficiary_name:
            faltantes.append("Beneficiario no encontrado")
        if not self.account_number:
            faltantes.append("Cuenta no encontrada")
        if not self.clabe:
            faltantes.append("CLABE no encontrada")
        elif not self.clabe_is_valid:
            faltantes.append("CLABE invalida")
        if not self.bank_name:
            faltantes.append("Banco no encontrado")
        if not self.rfc:
            faltantes.append("RFC no encontrado")
        return " | ".join(faltantes)

    def to_export_row(self) -> dict[str, str]:
        return {
            "archivo": self.file_name,
            "tipo": self.source_type,
            "beneficiario": self.beneficiary_name,
            "cuenta": self.account_number,
            "clabe": self.clabe,
            "clabe_valida": "Si" if self.clabe_is_valid else "No",
            "banco": self.bank_name,
            "rfc": self.rfc,
            "csf_nombre": self.csf_name,
            "csf_rfc": self.csf_rfc,
            "estado": self.status,
            "notas": self.report_notes,
        }


def _normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip(" \t\r\n:-|")


def _clean_digits(value: str) -> str:
    return re.sub(r"\D+", "", value or "")


# Confusiones tipicas del OCR letra->digito (para campos que deben ser solo numeros,
# como la CLABE: el OCR a veces lee '5' como 'S', '0' como 'O', etc.).
_DIGIT_OCR_FIX = str.maketrans(
    {"O": "0", "D": "0", "Q": "0", "I": "1", "L": "1", "S": "5", "B": "8", "G": "6", "Z": "2", "T": "7", "A": "4"}
)


def _digits_with_ocr_fix(value: str) -> str:
    """Devuelve solo los digitos del valor, corrigiendo confusiones letra->digito."""
    return re.sub(r"\D+", "", (value or "").upper().translate(_DIGIT_OCR_FIX))


def _normalize_for_match(value: str) -> str:
    ascii_value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", ascii_value.lower()).strip()


def _normalize_rfc_candidate(value: str) -> str:
    return _normalize_for_match(value).replace(" ", "").upper()


# RFC genericos del SAT (placeholders): nacional y de extranjeros. No identifican a
# nadie, asi que se tratan como "sin RFC" para completarlos desde la constancia.
GENERIC_RFCS = {"XAXX010101000", "XEXX010101000"}


def _is_placeholder_rfc(value: str) -> bool:
    return _normalize_rfc_candidate(value) in GENERIC_RFCS


def _looks_like_rfc(value: str) -> bool:
    candidate = _normalize_rfc_candidate(value)
    match = re.fullmatch(r"[A-Z&]{3,4}(\d{2})(\d{2})(\d{2})[A-Z0-9]{3}", candidate)
    if not match:
        return False
    month = int(match.group(2))
    day = int(match.group(3))
    return 1 <= month <= 12 and 1 <= day <= 31


# Confusiones tipicas del OCR, corregidas segun la posicion esperada en el RFC.
_OCR_TO_LETTER = str.maketrans({"0": "O", "1": "I", "5": "S", "8": "B", "6": "G", "2": "Z", "4": "A"})
_OCR_TO_DIGIT = str.maketrans(
    {"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "G": "6", "Z": "2", "Q": "0", "D": "0", "A": "4", "T": "7"}
)


def _repair_rfc_candidate(value: str) -> str:
    """Intenta recuperar un RFC garabateado por el OCR corrigiendo caracteres por posicion.

    Estructura del RFC: 3-4 letras + 6 digitos (AAMMDD) + 3 alfanumericos (homoclave).
    Solo debe usarse cuando hay una etiqueta de RFC cerca, para no fabricar falsos positivos.
    """
    # El OCR suele leer la 'S' inicial como '$'; lo recuperamos antes de limpiar simbolos.
    value = value.replace("$", "S")
    compact = re.sub(r"[^A-Z0-9&]", "", _normalize_rfc_candidate(value))
    if len(compact) < 12:
        return ""
    for length in (13, 12):
        if len(compact) < length:
            continue
        n_letters = 4 if length == 13 else 3
        for start in range(0, len(compact) - length + 1):
            window = compact[start : start + length]
            raw_date = window[n_letters : n_letters + 6]
            # La fecha del RFC son 6 digitos: la reparacion solo corrige 1-2 errores
            # de OCR, NO convierte una palabra entera (p.ej. 'OSILLO' de HERMOSILLO)
            # en digitos. Exigimos que al menos 4 de los 6 ya sean digitos.
            if sum(ch.isdigit() for ch in raw_date) < 4:
                continue
            prefix = window[:n_letters].translate(_OCR_TO_LETTER)
            date = raw_date.translate(_OCR_TO_DIGIT)
            homoclave = window[n_letters + 6 :]
            candidate = prefix + date + homoclave
            if _looks_like_rfc(candidate):
                return candidate
    return ""


def _looks_like_beneficiary_name(value: str) -> bool:
    cleaned = _normalize_spaces(value)
    if not cleaned or any(char.isdigit() for char in cleaned):
        return False

    # Un nombre real trae al menos una mayuscula; descarta basura del OCR en
    # minusculas como 'le marzo' o 'spin' de capturas de apps.
    if not any(char.isupper() for char in cleaned):
        return False

    normalized = _normalize_for_match(cleaned)
    if any(
        marker in normalized
        for marker in (
            "banco",
            "clabe",
            "rfc",
            "cuenta",
            "iban",
            "sucursal",
            "referencia",
            "datos",
            "fiscal",
            "domicilio",
            "direccion",
            "correo",
            "telefono",
            "firma",
            "estado",
            "fecha",
            # Etiquetas y encabezados que el OCR confunde con el nombre.
            "cliente",
            "producto",
            "contrato",
            "tarjeta",
            "nomina",
            "libreton",
            "toque",
            "posterior",
            "solicitante",
            "autorizador",
            "apellido",
            "nacimiento",
            "nacionalidad",
            "celular",
            "empleador",
            "autenticacion",
            "vendedor",
            "saldo",
            "disponible",
            "periodo",
            "pagina",
            # Linea de la institucion financiera (IFPE/banco), no del titular.
            "institucion",
            "fondos de pago",
            "banca multiple",
            "grupo financiero",
            # Nombres de banco (logos/encabezados) que no son el titular. Se listan los
            # que no son palabras comunes (se evitan 'bajio', 'azteca', etc.).
            "santander",
            "banregio",
            "banorte",
            "bancomer",
            "scotiabank",
            "hsbc",
            "citibanamex",
            "bancoppel",
            # Palabras de textos legales/avisos (no aparecen en nombres ni empresas).
            "incumplir",
            "obligaciones",
            "moratorios",
            "comision",
            "intereses",
            "generar",
        )
    ):
        return False

    # Rechaza lineas de domicilio (coincidencia por palabra completa para no afectar
    # nombres que contengan esas letras, p.ej. 'av' dentro de 'GUSTAVO').
    address_words = {
        "av", "ave", "avenida", "calle", "col", "colonia", "blvd", "boulevard",
        "carretera", "priv", "privada", "fracc", "fraccionamiento", "manzana", "mz",
        "lote", "lt", "andador", "cp", "int", "ext", "esq", "circuito", "cda", "cerrada",
        "prolongacion", "retorno", "diagonal", "calzada", "depto", "edificio",
    }
    if any(word in address_words for word in normalized.split()):
        return False

    words = cleaned.split()
    if len(words) < 2 or len(words) > 10:
        return False

    letters = sum(char.isalpha() for char in cleaned)
    if letters < 6:
        return False

    # Descarta oraciones/textos legales (p.ej. "Incumplir tus obligaciones te puede
    # generar comisiones..."). Un nombre va en MAYUSCULAS o Tipo Titulo; una oracion
    # trae varias palabras en minuscula que no son particulas ('tus', 'puede', etc.).
    particles_lower = {particle.lower() for particle in _NAME_PARTICLES}
    lower_non_particle = sum(
        1 for word in words if word.isalpha() and word.islower() and word.lower() not in particles_lower
    )
    if lower_non_particle >= 2:
        return False

    normalized_words = [word.strip(".").lower() for word in words]
    business_markers = {
        "sa",
        "s.a",
        "s.a.",
        "cv",
        "c.v",
        "c.v.",
        "de",
        "rl",
        "r.l",
        "r.l.",
        "decv",
        "sociedad",
        "anónima",
        "anonima",
        "empresa",
        "compania",
        "corp",
        "corporativo",
        "servicios",
        "industriales",
    }
    if any(marker in normalized for marker in ("sa de cv", "s de rl", "sociedad anonima", "razon social")):
        return True
    if len(normalized_words) >= 2 and all(word.isalpha() or word in business_markers for word in normalized_words):
        return True
    return len(words) >= 2


def _app_base_dir() -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "executable"):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(filename: str) -> str:
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if isinstance(bundle_dir, str) and bundle_dir:
        candidate = os.path.join(bundle_dir, filename)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(_app_base_dir(), filename)


def _is_tesseract_ready() -> bool:
    if pytesseract is None:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _ocr_available() -> bool:
    return pytesseract is not None and Image is not None and ImageOps is not None and _is_tesseract_ready()


def _configure_tesseract_path() -> None:
    if pytesseract is None:
        return
    if _is_tesseract_ready():
        return
    bundle_dir = getattr(sys, "_MEIPASS", None)
    extra_candidates = [
        os.path.join(bundle_dir, "Tesseract-OCR", "tesseract.exe") if isinstance(bundle_dir, str) and bundle_dir else "",
        os.path.join(bundle_dir, "tesseract.exe") if isinstance(bundle_dir, str) and bundle_dir else "",
        os.path.join(_app_base_dir(), "Tesseract-OCR", "tesseract.exe"),
        os.path.join(_app_base_dir(), "tesseract.exe"),
    ]
    for candidate in extra_candidates + TESSERACT_DEFAULT_PATHS:
        if candidate and os.path.exists(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            tessdata_dir = os.path.join(os.path.dirname(candidate), "tessdata")
            if os.path.isdir(tessdata_dir):
                os.environ["TESSDATA_PREFIX"] = tessdata_dir
            return


_configure_tesseract_path()


def _extract_labeled_value(text: str, labels: list[str]) -> str:
    lines = [line.strip() for line in text.splitlines()]
    for index, line in enumerate(lines):
        compact = _normalize_spaces(line)
        compact_lower = compact.lower()
        for label in labels:
            if label not in compact_lower:
                continue
            tail = compact_lower.find(label)
            if tail != -1:
                candidate = compact[tail + len(label) :].strip(" \t:-|")
                if candidate:
                    return _normalize_spaces(candidate)
            for next_line in lines[index + 1 :]:
                next_line = _normalize_spaces(next_line)
                if next_line:
                    return next_line
    return ""


def _value_after_label(line: str, normalized_label: str) -> str:
    """Devuelve lo que sigue a la etiqueta dentro del mismo renglon.

    Maneja acentos y casos donde un token del OCR se normaliza a varias palabras
    (por ejemplo '21/04/2026' -> '21 04 2026'), respetando los limites originales.
    """
    label_tokens = normalized_label.split()
    if not label_tokens:
        return ""
    orig_tokens = line.split()
    flat: list[tuple[str, int]] = []
    for idx, token in enumerate(orig_tokens):
        for word in _normalize_for_match(token).split():
            flat.append((word, idx))
    words = [word for word, _ in flat]
    n = len(label_tokens)
    for i in range(len(words) - n + 1):
        if words[i : i + n] == label_tokens:
            after_idx = flat[i + n - 1][1] + 1
            tail = " ".join(orig_tokens[after_idx:]).strip(" \t:-|")
            if tail:
                return tail
    return ""


def _extract_labeled_value_normalized(text: str, labels: list[str]) -> str:
    normalized_labels = [_normalize_for_match(label) for label in labels]
    lines = [_normalize_spaces(line) for line in text.splitlines() if _normalize_spaces(line)]
    normalized_lines = [_normalize_for_match(line) for line in lines]
    # Recorremos las etiquetas en orden de prioridad (las mas especificas primero)
    # sobre todo el texto, para que "no de cuenta" gane a un generico "cuenta" que
    # podria aparecer en lineas como "Estado de Cuenta".
    for normalized_label in normalized_labels:
        for index, line in enumerate(lines):
            if normalized_label not in normalized_lines[index]:
                continue
            same_line = _value_after_label(line, normalized_label)
            if same_line:
                return _normalize_spaces(same_line)
            candidate = re.split(r"[:\-|]", line, maxsplit=1)
            if len(candidate) > 1:
                value = _normalize_spaces(candidate[1])
                if value:
                    return value
            for next_line in lines[index + 1 :]:
                if next_line:
                    return next_line
    return ""


# Particulas cortas validas en nombres de personas y negocios (no son ruido).
_NAME_PARTICLES = {
    "DE", "DEL", "LA", "LAS", "LOS", "Y", "E", "SAN", "SANTA", "DA", "DI", "DO",
    "MC", "MAC", "VON", "VAN", "SA", "CV", "RL", "SC", "SAS", "SAB", "SAPI", "SADECV",
}


def _clean_name_candidate(value: str) -> str:
    """Recorta basura del OCR cuando el nombre viene en MAYUSCULAS.

    En los estados de cuenta el titular aparece en mayusculas dentro del bloque de
    direccion; el OCR suele pegarle texto de columnas vecinas ('... DS al?) Ballas').
    Si el nombre arranca en mayusculas, conservamos solo la racha inicial de palabras
    en mayusculas, cortamos en el primer token con minusculas/digitos/simbolos y
    eliminamos ruido corto del final (p.ej. 'DS') que no sea una particula valida.
    Si no parece estar en mayusculas, se devuelve sin cambios para no danar nombres
    legitimos en minusculas o mixtos.
    """
    tokens = value.split()
    if not tokens:
        return value

    def core(token: str) -> str:
        return token.strip(".,:;|/\\()[]{}\"'?!¡¿-")

    leading = [core(token) for token in tokens[:2]]
    if sum(1 for token in leading if token.isalpha() and token.isupper()) < 2:
        return value

    kept: list[str] = []
    for token in tokens:
        clean = core(token)
        if clean.isalpha() and clean.isupper():
            kept.append(clean)
        else:
            break

    # Quita ruido corto del final (1-2 letras sueltas) que no sea particula valida.
    # Conserva sufijos de negocio (SA, CV, RL...) y conectores ('DE', 'LA').
    while kept and len(kept[-1]) <= 2 and kept[-1] not in _NAME_PARTICLES:
        kept.pop()

    return " ".join(kept) if len(kept) >= 2 else ""


def _finalize_beneficiary(raw: str) -> str:
    """Acepta un candidato solo si el renglon ORIGINAL ya parece un nombre.

    Validar el original primero evita 'rescatar' lineas de datos (p.ej.
    'DEPOSITOS ANULADOS: 0.0 MXN') cuya limpieza les quitaria los digitos. Una vez
    validado, se limpia para quitar basura del OCR al final del nombre real.
    """
    candidate = _normalize_spaces(raw)
    if not candidate or not _looks_like_beneficiary_name(candidate):
        return ""
    cleaned = _normalize_spaces(_clean_name_candidate(candidate))
    if cleaned and _looks_like_beneficiary_name(cleaned):
        return cleaned
    return ""


def _extract_beneficiary_name(text: str, bank_hint: str = "") -> str:
    lines = [_normalize_spaces(line) for line in text.splitlines() if _normalize_spaces(line)]

    # Formato BBVA: el titular va en el renglon siguiente al numero de cliente.
    # Al conocer el banco, enfocamos la busqueda en esa estructura.
    if bank_hint == "BBVA":
        for index, line in enumerate(lines):
            if "no de cliente" in _normalize_for_match(line):
                for next_line in lines[index + 1 :]:
                    result = _finalize_beneficiary(next_line)
                    if result:
                        return result
                break

    value = _extract_labeled_value_normalized(text, BENEFICIARY_NAME_LABELS)
    if value:
        result = _finalize_beneficiary(value)
        if result:
            return result

    # Encabezado de columna de nombre (formularios tipo Citibanamex):
    # 'NOMBRE(S), APELLIDO PATERNO / APELLIDO MATERNO' -> el nombre va en una linea
    # siguiente. Estos formularios traen varios nombres (vendedor, autorizador,
    # titular); el del titular va bajo 'SOLICITANTE', asi que arrancamos la busqueda
    # del encabezado a partir de esa palabra cuando existe. Va primero porque el
    # heuristico de etiquetas se confunde con encabezados como 'TARJETA (TITULAR)'.
    start = 0
    for index, line in enumerate(lines):
        if "solicitante" in _normalize_for_match(line):
            start = index
            break
    if start:
        for index in range(start, len(lines)):
            normalized_line = _normalize_for_match(lines[index])
            if "nombre" in normalized_line and "apellido" in normalized_line:
                for next_line in lines[index + 1 :]:
                    result = _finalize_beneficiary(next_line)
                    if result:
                        return result
                break

    for index, line in enumerate(lines):
        normalized_line = _normalize_for_match(line)
        if not any(_normalize_for_match(label) in normalized_line for label in BENEFICIARY_NAME_LABELS):
            continue
        # Evita etiquetas dentro de oraciones/avisos legales (p.ej. "...a nombre del
        # Titular de la cuenta."): una etiqueta de campo real es corta.
        if len(line.split()) > 6:
            continue
        candidate = re.split(r"[:\-|]", line, maxsplit=1)
        if len(candidate) > 1:
            result = _finalize_beneficiary(candidate[1])
            if result:
                return result
        # El valor va en la misma linea o en la inmediata siguiente; no escaneamos
        # todo el documento (eso tomaba basura lejana del OCR).
        for next_line in lines[index + 1 : index + 3]:
            result = _finalize_beneficiary(next_line)
            if result:
                return result

    for line in lines:
        result = _finalize_beneficiary(line)
        if result:
            return result
    return ""


def _clabe_is_valid(clabe: str) -> bool:
    """Valida una CLABE de 18 digitos con su digito verificador (pesos 3,7,1)."""
    if len(clabe) != 18 or not clabe.isdigit():
        return False
    weights = (3, 7, 1)
    total = sum((int(digit) * weights[i % 3]) % 10 for i, digit in enumerate(clabe[:17]))
    check_digit = (10 - (total % 10)) % 10
    return check_digit == int(clabe[17])


def _extract_clabe(text: str) -> str:
    candidates: list[str] = []

    labeled = _extract_labeled_value(text, ["clabe interbancaria", "clabe"])
    if labeled:
        digits = _clean_digits(labeled)
        if len(digits) == 18:
            candidates.append(digits)
        else:
            # Reintento corrigiendo confusiones del OCR (p.ej. ultimo '5' leido como 'S').
            fixed = _digits_with_ocr_fix(labeled)
            if len(fixed) == 18:
                candidates.append(fixed)

    for candidate in CLABE_NUMBER_PATTERN.findall(text):
        digits = _clean_digits(candidate)
        if len(digits) == 18:
            candidates.append(digits)

    # Ventanas de 18 digitos dentro de cada linea: la CLABE a veces va pegada a otros
    # numeros en una fila de tabla ('... 1322166896 072 744 01322166896 8 $0.00').
    # Solo se aceptan ventanas con prefijo de banco conocido y digito verificador valido.
    for line in text.splitlines():
        digits = _clean_digits(line)
        for i in range(0, len(digits) - 17):
            window = digits[i : i + 18]
            if window[:3] in BANK_CODE_MAP and _clabe_is_valid(window):
                candidates.append(window)

    if not candidates:
        return ""

    # Prefiere la CLABE con prefijo de banco conocido y digito verificador valido.
    def score(clabe: str) -> int:
        return (2 if clabe[:3] in BANK_CODE_MAP else 0) + (1 if _clabe_is_valid(clabe) else 0)

    candidates.sort(key=score, reverse=True)
    return candidates[0]


def _looks_like_date_digits(digits: str) -> bool:
    """Detecta secuencias que en realidad son fechas concatenadas (DDMMAAAA).

    Evita que el periodo del estado de cuenta ('DEL 21/04/2026 AL 20/05/2026' ->
    '2104202620052026') se confunda con un numero de cuenta.
    """

    def valid_ddmmyyyy(value: str) -> bool:
        if len(value) != 8 or not value.isdigit():
            return False
        day, month, year = int(value[:2]), int(value[2:4]), int(value[4:])
        return 1 <= day <= 31 and 1 <= month <= 12 and 1990 <= year <= 2100

    if valid_ddmmyyyy(digits):
        return True
    if len(digits) == 16 and valid_ddmmyyyy(digits[:8]) and valid_ddmmyyyy(digits[8:]):
        return True
    return False


def _extract_account_number(text: str, clabe: str, rfc: str = "") -> str:
    rfc_digits = _clean_digits(rfc)

    def is_valid_account(digits: str) -> bool:
        if not (8 <= len(digits) <= 20):
            return False
        # La CLABE (18 digitos) incluye al numero de cuenta dentro de si misma, asi
        # que solo descartamos la CLABE completa, no que la cuenta sea subcadena suya.
        if digits == clabe or len(digits) == 18:
            return False
        if rfc_digits and digits == rfc_digits:
            return False
        return True

    # 1) Valor junto a una etiqueta de cuenta (las mas especificas tienen prioridad).
    labeled = _extract_labeled_value_normalized(
        text,
        [
            "numero de cuenta",
            "numero cuenta",
            "nro de cuenta",
            "no. de cuenta",
            "no de cuenta",
            "cuenta bancaria",
            "cuenta",
            "cta",
        ],
    )
    if labeled:
        digits = _clean_digits(labeled)
        if is_valid_account(digits):
            return digits

    # 2) Respaldo: alguna secuencia de digitos suelta, descartando fechas.
    for candidate in ACCOUNT_PATTERN.findall(text):
        digits = _clean_digits(candidate)
        if is_valid_account(digits) and not _looks_like_date_digits(digits):
            return digits
    return ""


def _extract_rfc(text: str) -> str:
    lines = [_normalize_spaces(line) for line in text.splitlines() if _normalize_spaces(line)]

    def scan_value(value: str) -> str:
        compact = re.sub(r"[^A-Z0-9&]", "", _normalize_rfc_candidate(value))
        if len(compact) < 12:
            return ""
        for length in (13, 12):
            if len(compact) < length:
                continue
            for start in range(0, len(compact) - length + 1):
                candidate = compact[start : start + length]
                if _looks_like_rfc(candidate):
                    return candidate
        return ""

    normalized_labels = [_normalize_for_match(label).replace(" ", "") for label in RFC_LABELS]

    def line_has_rfc_label(value: str) -> bool:
        compact = _normalize_for_match(value).replace(" ", "")
        return any(label in compact for label in normalized_labels)

    def is_bank_rfc(candidate: str, index: int) -> bool:
        """True si el RFC es de la institucion (por lista conocida o por contexto)."""
        norm = _normalize_rfc_candidate(candidate)
        # Subcadena: el OCR a veces pega una letra de la etiqueta 'R.F.C.' al RFC del
        # banco (p.ej. 'CBSM970519DU8' por 'BSM970519DU8').
        if any(bank_rfc in norm for bank_rfc in BANK_RFCS):
            return True
        context = _normalize_for_match(" ".join(lines[index : index + 2]))
        return any(marker in context for marker in _BANK_RFC_CONTEXT)

    # El RFC del banco suele ir en el encabezado, antes que el del titular. Guardamos
    # el del banco solo como ultimo recurso y devolvemos el primero que NO sea del banco.
    bank_fallback = ""

    def consider(candidate: str, index: int) -> Optional[str]:
        nonlocal bank_fallback
        if not candidate:
            return None
        if is_bank_rfc(candidate, index):
            if not bank_fallback:
                bank_fallback = candidate
            return None
        return candidate

    for index, line in enumerate(lines):
        if line_has_rfc_label(line):
            parts = re.split(r"[:\-|]", line, maxsplit=1)
            candidate = ""
            if len(parts) > 1:
                candidate = _repair_rfc_candidate(parts[1]) or scan_value(parts[1])
            if not candidate:
                candidate = _repair_rfc_candidate(line) or scan_value(line)
            if not candidate:
                for next_line in lines[index + 1 :]:
                    next_normalized = _normalize_for_match(next_line)
                    if any(blocker in next_normalized for blocker in ("banco", "cuenta", "clabe", "nombre", "titular", "beneficiario")):
                        break
                    candidate = _repair_rfc_candidate(next_line) or scan_value(next_line)
                    if candidate:
                        break
            chosen = consider(candidate, index)
            if chosen:
                return chosen

    for index, line in enumerate(lines):
        candidate = scan_value(line)
        if candidate and line_has_rfc_label(line):
            chosen = consider(candidate, index)
            if chosen:
                return chosen
        if candidate and index + 1 < len(lines) and line_has_rfc_label(lines[index + 1]):
            chosen = consider(candidate, index)
            if chosen:
                return chosen

    # Respaldo final: un token aislado con forma de RFC, aunque no haya etiqueta cerca.
    # Se limita a tokens de largo razonable (12-16) para no confundirlo con referencias
    # u otros codigos largos; la validacion estricta de fecha evita falsos positivos.
    for index, line in enumerate(lines):
        for token in line.split():
            compact = re.sub(r"[^A-Z0-9&]", "", _normalize_rfc_candidate(token))
            if not (12 <= len(compact) <= 16):
                continue
            chosen = consider(scan_value(token), index)
            if chosen:
                return chosen

    # No se encontro el RFC del titular. Nunca devolvemos el RFC del banco (seria un
    # dato equivocado); preferimos dejarlo vacio para que se revise o se complete con
    # la constancia fiscal.
    return ""

def _extract_bank_name(text: str, clabe: str) -> str:
    # 1) Codigo CLABE: identifica al banco EMISOR de la cuenta, es lo mas confiable.
    if len(clabe) >= 3:
        code = clabe[:3]
        if code in BANK_CODE_MAP:
            return BANK_CODE_MAP[code]

    # 2) Alias del banco, pero SOLO en el encabezado (primeras lineas). Asi evitamos
    # tomar bancos mencionados en los movimientos ('BANCO ORIGEN: BANORTE/IXE').
    header_lines = [line for line in text.splitlines() if line.strip()][:8]
    header_text = _normalize_for_match("\n".join(header_lines))
    for alias, display in BANK_NAME_ALIASES:
        if _normalize_for_match(alias) in header_text:
            return display

    # 3) Etiqueta 'banco' explicita. Si el valor contiene un alias conocido (p.ej.
    # 'Banco Mercantil del Norte ... Banorte'), se devuelve el banco normalizado;
    # nunca una cadena larga de la institucion.
    labeled = _extract_labeled_value_normalized(text, ["banco"])
    if labeled:
        normalized_labeled = _normalize_for_match(labeled)
        for alias, display in BANK_NAME_ALIASES:
            if _normalize_for_match(alias) in normalized_labeled:
                return display
        cleaned = _normalize_spaces(labeled)
        if cleaned and len(cleaned.split()) <= 4:
            return cleaned

    # 4) Codigo CLABE no catalogado.
    if len(clabe) >= 3:
        return f"Codigo CLABE {clabe[:3]}"

    return ""


def _extract_pdf_text(file_path: str) -> tuple[str, list[str]]:
    notes: list[str] = []
    if PdfReader is None:
        return "", ["Libreria para leer PDF no disponible"]

    try:
        reader = PdfReader(file_path)
    except Exception as exc:
        return "", [f"No fue posible abrir el PDF: {exc}"]

    pages: list[str] = []
    for index, page in enumerate(reader.pages, start=1):
        try:
            page_text = page.extract_text() or ""
            if page_text.strip():
                pages.append(page_text)
        except Exception as exc:
            notes.append(f"No se pudo leer la pagina {index}: {exc}")
    return "\n".join(pages).strip(), notes


def _otsu_threshold(gray_image: "Image.Image") -> int:
    """Calcula el umbral optimo de binarizacion (metodo de Otsu) desde el histograma."""
    hist = gray_image.histogram()[:256]
    total = sum(hist)
    if total == 0:
        return 128
    sum_total = sum(i * h for i, h in enumerate(hist))
    sum_b = 0.0
    weight_b = 0
    max_between = -1.0
    threshold = 128
    for level in range(256):
        weight_b += hist[level]
        if weight_b == 0:
            continue
        weight_f = total - weight_b
        if weight_f == 0:
            break
        sum_b += level * hist[level]
        mean_b = sum_b / weight_b
        mean_f = (sum_total - sum_b) / weight_f
        between = weight_b * weight_f * (mean_b - mean_f) ** 2
        if between > max_between:
            max_between = between
            threshold = level
    return threshold


def _ocr_pil_image(image: "Image.Image") -> tuple[str, list[str]]:
    notes: list[str] = []
    if not _ocr_available():
        return "", ["OCR no disponible. Instala Pillow, pytesseract y Tesseract OCR."]

    base = ImageOps.exif_transpose(image)
    base = ImageOps.grayscale(base)
    base = ImageOps.autocontrast(base)
    # El texto chico de las fotos/tablas necesita mas resolucion para que el OCR no se
    # "salte" filas (como la del RFC). Escalamos hasta ~2400 px en el lado mayor.
    target = 2400
    longest_side = max(base.size)
    if longest_side < target:
        factor = min(4, max(2, round(target / longest_side)))
        base = base.resize((base.width * factor, base.height * factor))

    # Variante 1: enfocada (define bordes de caracteres pequenos).
    sharp = base
    if ImageFilter is not None:
        sharp = base.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
    variants = [("nitida", sharp)]
    # Variante 2: binarizada (Otsu). Ayuda mucho en fotos .jpg/.heic con
    # iluminacion irregular, fondos grises o sombras del telefono.
    try:
        threshold = _otsu_threshold(base)
        variants.append(("binarizada", base.point(lambda px: 255 if px > threshold else 0)))
    except Exception:
        pass

    psm_modes = ("--psm 6", "--psm 4", "--psm 11")
    last_error: Optional[Exception] = None

    def run_language(lang: Optional[str]) -> tuple[str, str]:
        nonlocal last_error
        best_local = ""
        best_label_local = ""
        for variant_name, variant_image in variants:
            for psm in psm_modes:
                try:
                    kwargs = {"config": psm}
                    if lang:
                        kwargs["lang"] = lang
                    text = pytesseract.image_to_string(variant_image, **kwargs).strip()
                except Exception as exc:
                    last_error = exc
                    continue
                # Elegimos la pasada con mas contenido: suele ser la que no omite filas.
                if len(text) > len(best_local):
                    best_local = text
                    best_label_local = f"{lang or 'sin idioma'}, {variant_name}, {psm}"
        return best_local, best_label_local

    # Probamos espanol+ingles; solo si falla recurrimos a ingles o al idioma por defecto.
    best_text, best_label = run_language("spa+eng")
    if not best_text:
        best_text, best_label = run_language("eng")
    if not best_text:
        best_text, best_label = run_language(None)

    if best_text:
        notes.append(f"OCR aplicado con {best_label}.")
        return best_text, notes

    if last_error is not None:
        notes.append(f"Fallo OCR en la ultima pasada: {last_error}")
    return "", notes


def _ocr_image(file_path: str) -> tuple[str, list[str]]:
    if Image is None:
        return "", ["Pillow no esta instalado; no es posible leer JPG/PNG."]
    try:
        with Image.open(file_path) as image:
            return _ocr_pil_image(image)
    except Exception as exc:
        return "", [f"No fue posible leer la imagen: {exc}"]


def _ocr_pdf(file_path: str) -> tuple[str, list[str]]:
    notes: list[str] = []
    if fitz is None or Image is None:
        return "", ["PyMuPDF y Pillow son necesarios para hacer OCR sobre PDFs escaneados."]
    if not _ocr_available():
        return "", ["OCR no disponible. Instala Pillow, pytesseract y Tesseract OCR."]

    try:
        doc = fitz.open(file_path)
    except Exception as exc:
        return "", [f"No fue posible abrir el PDF para OCR: {exc}"]

    pages: list[str] = []
    try:
        for index, page in enumerate(doc, start=1):
            try:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                page_text, page_notes = _ocr_pil_image(image)
                if page_notes:
                    notes.extend(page_notes)
                if page_text.strip():
                    pages.append(page_text)
            except Exception as exc:
                notes.append(f"No se pudo hacer OCR en la pagina {index}: {exc}")
    finally:
        doc.close()
    return "\n".join(pages).strip(), notes


def extract_text(file_path: str) -> tuple[str, str, list[str]]:
    extension = os.path.splitext(file_path)[1].lower()
    notes: list[str] = []

    if extension == ".pdf":
        text, pdf_notes = _extract_pdf_text(file_path)
        notes.extend(pdf_notes)
        if _ocr_available():
            ocr_text, ocr_notes = _ocr_pdf(file_path)
            notes.extend(ocr_notes)
            if ocr_text:
                notes.append("Se aplico OCR al PDF.")
                if text.strip():
                    text = f"{text}\n{ocr_text}".strip()
                else:
                    text = ocr_text
        source_type = "PDF"
    else:
        text, image_notes = _ocr_image(file_path)
        notes.extend(image_notes)
        source_type = "Imagen"

    return text.strip(), source_type, notes


def parse_document(file_path: str, bank_hint: str = "") -> ExtractionResult:
    raw_text, source_type, notes = extract_text(file_path)
    return _result_from_text(file_path, raw_text, source_type, notes, bank_hint)


def _result_from_text(
    file_path: str,
    raw_text: str,
    source_type: str,
    notes: list[str],
    bank_hint: str = "",
) -> ExtractionResult:
    normalized_text = raw_text or ""
    clabe = _extract_clabe(normalized_text)
    beneficiary_name = _extract_beneficiary_name(normalized_text, bank_hint)
    rfc = _extract_rfc(normalized_text)
    account_number = _extract_account_number(normalized_text, clabe, rfc)
    clabe_is_valid = _clabe_is_valid(clabe)

    # Si el usuario indico el banco, se respeta (delimita el formato); si no, se detecta.
    if bank_hint:
        bank_name = bank_hint
        expected_prefix = BANK_EXPECTED_CLABE_PREFIX.get(bank_hint)
        if expected_prefix and clabe[:3] and clabe[:3] != expected_prefix:
            notes.append(
                f"Aviso: indicaste {bank_hint} pero la CLABE empieza en {clabe[:3]} "
                f"(se esperaba {expected_prefix}). Verifica el banco del archivo."
            )
    else:
        bank_name = _extract_bank_name(normalized_text, clabe)

    if not raw_text:
        notes.append("No se encontro texto util para analizar.")
    if clabe and not clabe_is_valid:
        if len(clabe) != 18 or not clabe.isdigit():
            notes.append("La CLABE encontrada no tiene 18 digitos.")
        else:
            notes.append("La CLABE tiene 18 digitos pero el digito verificador no coincide.")
    elif not clabe:
        notes.append("No se encontro CLABE interbancaria.")
    if not beneficiary_name:
        notes.append("No se pudo identificar el nombre del beneficiario.")
    if not rfc:
        notes.append("No se pudo identificar el RFC.")

    return ExtractionResult(
        file_path=file_path,
        source_type=source_type,
        extracted_text=raw_text,
        beneficiary_name=beneficiary_name,
        account_number=account_number,
        clabe=clabe,
        bank_name=bank_name,
        rfc=rfc,
        clabe_is_valid=clabe_is_valid,
        notes=notes,
    )


@dataclasses.dataclass
class ConstanciaFiscal:
    """Datos tomados de una Constancia de Situacion Fiscal (CSF) del SAT."""

    rfc: str
    name: str
    file_path: str


def _is_constancia(text: str) -> bool:
    """Reconoce si el texto corresponde a una Constancia de Situacion Fiscal.

    Usa marcadores EXCLUSIVOS del documento del SAT. No basta con que aparezca la
    frase 'constancia de situacion fiscal', porque los estados de cuenta la mencionan
    en oraciones (p.ej. 'entregue copia de su constancia de situacion fiscal...').
    """
    normalized = _normalize_for_match(text)
    return (
        "cedula de identificacion fiscal" in normalized
        or "idcif" in normalized
        or "datos de identificacion del contribuyente" in normalized
    )


def _extract_constancia_name(text: str) -> str:
    """Obtiene el nombre/razon social desde la CSF (campos del SAT, texto limpio).

    Formato real SAT:
      - Seccion de detalle (persona fisica): 'Nombre (s):', 'Primer Apellido:',
        'Segundo Apellido:'.
      - Encabezado: el nombre/razon social completo va en el renglon JUSTO ANTES de
        la etiqueta 'Nombre, denominacion o razon social' (sirve para morales).
    """
    # 1) Persona fisica: combinar nombre(s) + apellidos de la seccion de detalle.
    nombre = _extract_labeled_value_normalized(text, ["nombre (s)", "nombre(s)"])
    ap1 = _extract_labeled_value_normalized(text, ["primer apellido", "apellido paterno"])
    ap2 = _extract_labeled_value_normalized(text, ["segundo apellido", "apellido materno"])
    parts = [_normalize_spaces(p) for p in (nombre, ap1, ap2) if _normalize_spaces(p)]
    if len(parts) >= 2:  # al menos nombre + un apellido
        return _normalize_spaces(" ".join(parts))

    # 2) Encabezado: el nombre va en el renglon anterior a la etiqueta de nombre/razon.
    lines = [_normalize_spaces(line) for line in text.splitlines() if _normalize_spaces(line)]
    for index, line in enumerate(lines):
        normalized = _normalize_for_match(line)
        if "denominacion o razon" in normalized or "nombre denominacion" in normalized:
            if index > 0:
                candidate = _normalize_spaces(lines[index - 1])
                if candidate and any(ch.isalpha() for ch in candidate) and "contribuyentes" not in _normalize_for_match(candidate):
                    return candidate
            break

    # 3) Respaldo: una sola pieza de nombre o razon social con valor explicito.
    if parts:
        return _normalize_spaces(" ".join(parts))
    razon = _extract_labeled_value_normalized(text, ["denominacion o razon social:", "razon social:"])
    return _normalize_spaces(razon)


def _parse_constancia(text: str, file_path: str) -> ConstanciaFiscal:
    return ConstanciaFiscal(rfc=_extract_rfc(text), name=_extract_constancia_name(text), file_path=file_path)


def _names_match(a: str, b: str) -> bool:
    """Compara dos nombres/razones sociales con tolerancia a OCR y truncamientos."""
    na = _normalize_for_match(a)
    nb = _normalize_for_match(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    tokens_a = set(na.split())
    tokens_b = set(nb.split())
    if tokens_a and tokens_b:
        overlap = len(tokens_a & tokens_b) / min(len(tokens_a), len(tokens_b))
        if overlap >= 0.7:
            return True
    return difflib.SequenceMatcher(None, na, nb).ratio() >= 0.85


def _enrich_with_constancias(result: ExtractionResult, constancias: list[ConstanciaFiscal]) -> None:
    """Corrige el nombre y completa el RFC de una caratula con los datos de su CSF.

    Empareja por RFC y, si no, por similitud de nombre. El nombre de la constancia
    es el oficial del SAT: si hay emparejamiento, se usa como nombre del beneficiario
    para evitar registrar nombres incorrectos o truncados por el OCR. El RFC solo se
    completa si falta o es generico (no se sobrescribe un RFC real ya leido).
    """
    if not constancias:
        return

    # Empareja por RFC y, si no, por similitud de nombre (siempre, para poder mostrar
    # y validar el CSF aunque la caratula ya tuviera los datos).
    match: Optional[ConstanciaFiscal] = None
    matched_by = ""
    if result.rfc:
        rfc_norm = _normalize_rfc_candidate(result.rfc)
        for csf in constancias:
            if csf.rfc and _normalize_rfc_candidate(csf.rfc) == rfc_norm:
                match, matched_by = csf, "rfc"
                break
    if match is None and result.beneficiary_name:
        for csf in constancias:
            if csf.name and _names_match(result.beneficiary_name, csf.name):
                match, matched_by = csf, "name"
                break
    if match is None:
        return

    # Registra el CSF emparejado (para la columna de validacion).
    result.csf_name = match.name
    result.csf_rfc = match.rfc

    # El nombre de la constancia es el oficial del SAT: se usa como nombre del
    # beneficiario para registrar el correcto y evitar nombres incorrectos del OCR.
    if match.name and match.name != result.beneficiary_name:
        if result.beneficiary_name:
            result.notes.append(
                f"Nombre corregido con la constancia fiscal (OCR leyo '{result.beneficiary_name}')."
            )
        else:
            result.notes.append(f"Nombre tomado de la constancia fiscal (RFC {match.rfc or 's/RFC'}).")
        result.beneficiary_name = match.name

    # RFC: se toma el de la constancia cuando falta, es generico (XAXX/XEXX) o cuando
    # el emparejamiento fue por NOMBRE y el RFC leido NO coincide con el de la CSF
    # (senal de que el OCR lo fabrico o tomo el de otra entidad, p.ej. el del banco).
    if match.rfc:
        rfc_mismatch_by_name = (
            matched_by == "name"
            and result.rfc
            and _normalize_rfc_candidate(result.rfc) != _normalize_rfc_candidate(match.rfc)
        )
        if not result.rfc:
            result.notes.append(f"RFC tomado de la constancia fiscal ({match.name or 's/nombre'}).")
            result.rfc = match.rfc
        elif _is_placeholder_rfc(result.rfc):
            result.notes.append(
                f"La caratula traia RFC generico ({result.rfc}); se reemplazo por el de la constancia."
            )
            result.rfc = match.rfc
        elif rfc_mismatch_by_name:
            result.notes.append(
                f"RFC corregido con la constancia fiscal (OCR leyo '{result.rfc}', no coincide con el titular)."
            )
            result.rfc = match.rfc


class BankExtractorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        _configure_tesseract_path()
        self.title("Extractor de informacion bancaria")
        self.geometry("1280x780")
        self.minsize(1100, 680)

        self._selected_files: list[str] = []
        self._file_banks: dict[str, str] = {}  # ruta -> banco indicado ("" = auto)
        self._results: list[ExtractionResult] = []
        self._worker: Optional[threading.Thread] = None
        self._loader: Optional[tk.Toplevel] = None

        self.status_var = tk.StringVar(value=self._initial_status())
        self.summary_var = tk.StringVar(value="Sin archivos cargados")

        self._build_ui()
        self._refresh_buttons()

    def _initial_status(self) -> str:
        if _ocr_available():
            return "OCR listo. Puedes analizar PDF e imagenes."
        if Image is None:
            return "Falta Pillow. Instala dependencias para leer imagenes."
        return "OCR no disponible. Se podran leer PDFs con texto embebido."

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x")

        title = ttk.Label(header, text="Extractor de informacion bancaria", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w")

        subtitle = ttk.Label(
            header,
            text="Lee PDF o imagenes, extrae beneficiario, cuenta, CLABE, banco y RFC. Valida CLABE de 18 digitos.",
        )
        subtitle.pack(anchor="w", pady=(2, 10))

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=(0, 10))

        self.add_button = ttk.Button(buttons, text="Agregar caratulas...", command=self.open_loader)
        self.add_button.pack(side="left")

        self.remove_button = ttk.Button(buttons, text="Quitar seleccion", command=self.remove_selected)
        self.remove_button.pack(side="left", padx=(8, 0))

        self.clear_button = ttk.Button(buttons, text="Limpiar", command=self.clear_all)
        self.clear_button.pack(side="left", padx=(8, 0))

        self.extract_button = ttk.Button(buttons, text="Analizar archivos", command=self.run_extraction)
        self.extract_button.pack(side="left", padx=(8, 0))

        self.copy_button = ttk.Button(buttons, text="Copiar resultado", command=self.copy_selected_details)
        self.copy_button.pack(side="left", padx=(8, 0))

        self.export_csv_button = ttk.Button(buttons, text="Exportar CSV", command=self.export_csv)
        self.export_csv_button.pack(side="left", padx=(8, 0))

        self.export_excel_button = ttk.Button(buttons, text="Exportar Excel", command=self.export_excel)
        self.export_excel_button.pack(side="left", padx=(8, 0))

        summary = ttk.Label(root, textvariable=self.summary_var)
        summary.pack(anchor="w", pady=(0, 8))

        # La barra de estado y la de progreso se anclan al FONDO antes de la zona
        # expandible (paned). Asi quedan siempre visibles aunque la ventana sea chica.
        status = ttk.Label(root, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(side="bottom", fill="x", pady=(8, 0))

        try:
            style.configure("Analisis.Horizontal.TProgressbar", thickness=22)
        except Exception:
            pass

        progress_frame = ttk.Frame(root)
        progress_frame.pack(side="bottom", fill="x", pady=(8, 0))

        self.progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            style="Analisis.Horizontal.TProgressbar",
        )
        self.progress.pack(side="left", fill="x", expand=True)

        self.progress_var = tk.StringVar(value="0%")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var, width=12, anchor="e")
        self.progress_label.pack(side="right", padx=(8, 0))

        paned = ttk.PanedWindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, padding=(0, 0, 8, 0))
        paned.add(left, weight=1)

        right = ttk.Frame(paned)
        paned.add(right, weight=3)

        files_header = ttk.Frame(left)
        files_header.pack(fill="x")

        files_label = ttk.Label(files_header, text="Archivos cargados")
        files_label.pack(side="left", anchor="w")

        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_check = ttk.Checkbutton(
            files_header,
            text="Seleccionar todos",
            variable=self.select_all_var,
            command=self._toggle_select_all,
        )
        self.select_all_check.pack(side="right")

        file_frame = ttk.Frame(left)
        file_frame.pack(fill="both", expand=True, pady=(4, 0))

        self.file_list = tk.Listbox(file_frame, selectmode="extended", height=12)
        file_scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=file_scroll.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        file_scroll.pack(side="right", fill="y")

        results_label = ttk.Label(right, text="Resultados")
        results_label.pack(anchor="w")

        table_frame = ttk.Frame(right)
        table_frame.pack(fill="both", expand=True, pady=(4, 8))

        columns = ("archivo", "beneficiario", "cuenta", "clabe", "clabe_ok", "banco", "rfc", "csf", "estado")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)

        headings = {
            "archivo": "Archivo",
            "beneficiario": "Beneficiario",
            "cuenta": "Cuenta",
            "clabe": "CLABE",
            "clabe_ok": "CLABE 18",
            "banco": "Banco",
            "rfc": "RFC",
            "csf": "Constancia (CSF)",
            "estado": "Estado",
        }
        widths = {
            "archivo": 170,
            "beneficiario": 170,
            "cuenta": 110,
            "clabe": 160,
            "clabe_ok": 65,
            "banco": 120,
            "rfc": 105,
            "csf": 230,
            "estado": 90,
        }
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w", stretch=True)

        tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_selected_details)

        details_label = ttk.Label(right, text="Detalle y texto detectado")
        details_label.pack(anchor="w")

        detail_paned = ttk.PanedWindow(right, orient="vertical")
        detail_paned.pack(fill="both", expand=True, pady=(4, 0))

        detail_frame = ttk.Frame(detail_paned)
        detail_paned.add(detail_frame, weight=1)

        self.detail_text = tk.Text(detail_frame, wrap="word", height=8)
        detail_scroll = ttk.Scrollbar(detail_frame, orient="vertical", command=self.detail_text.yview)
        self.detail_text.configure(yscrollcommand=detail_scroll.set)
        self.detail_text.pack(side="left", fill="both", expand=True)
        detail_scroll.pack(side="right", fill="y")

        raw_label = ttk.Label(right, text="Vista previa OCR crudo")
        raw_label.pack(anchor="w", pady=(8, 0))

        raw_frame = ttk.Frame(detail_paned)
        detail_paned.add(raw_frame, weight=2)

        self.raw_text = tk.Text(raw_frame, wrap="word", height=10)
        raw_scroll = ttk.Scrollbar(raw_frame, orient="vertical", command=self.raw_text.yview)
        self.raw_text.configure(yscrollcommand=raw_scroll.set)
        self.raw_text.tag_configure("highlight_label", background="#fff2a8")
        self.raw_text.tag_configure("highlight_value", background="#cce8ff")
        self.raw_text.tag_configure("highlight_name", background="#dff7d8")
        self.raw_text.pack(side="left", fill="both", expand=True)
        raw_scroll.pack(side="right", fill="y")

    def _refresh_buttons(self) -> None:
        has_files = bool(self._selected_files)
        busy = self._worker is not None and self._worker.is_alive()
        self.add_button.configure(state="disabled" if busy else "normal")
        self.remove_button.configure(state="disabled" if busy or not has_files else "normal")
        self.clear_button.configure(state="disabled" if busy or not has_files else "normal")
        self.extract_button.configure(state="disabled" if busy or not has_files else "normal")
        self.copy_button.configure(state="disabled" if busy or not self._results else "normal")
        self.export_csv_button.configure(state="disabled" if busy or not self._results else "normal")
        self.export_excel_button.configure(state="disabled" if busy or not self._results else "normal")
        self.select_all_check.configure(state="disabled" if busy or not has_files else "normal")

    def _add_supported_paths(self, paths: list[str], bank: str = "") -> int:
        added = 0
        for path in paths:
            if os.path.splitext(path)[1].lower() not in SUPPORTED_EXTENSIONS:
                continue
            if path not in self._selected_files:
                self._selected_files.append(path)
                added += 1
            self._file_banks[path] = bank  # etiqueta (o reetiqueta) el banco indicado
        if added:
            self._refresh_file_list()
            etiqueta = bank if bank else "auto"
            self.status_var.set(f"Se agregaron {added} archivo(s) [{etiqueta}].")
        self._refresh_buttons()
        return added

    def open_loader(self) -> None:
        """Ventana para cargar archivos indicando el banco al que pertenecen."""
        if self._loader is not None and self._loader.winfo_exists():
            self._loader.lift()
            self._loader.focus_force()
            return

        win = tk.Toplevel(self)
        self._loader = win
        win.title("Agregar caratulas")
        win.geometry("480x220")
        win.transient(self)
        win.resizable(False, False)

        frame = ttk.Frame(win, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Banco de las caratulas a cargar:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self._loader_bank_var = tk.StringVar(value=AUTO_BANK_LABEL)
        combo = ttk.Combobox(
            frame,
            textvariable=self._loader_bank_var,
            values=SELECTABLE_BANKS,
            state="readonly",
        )
        combo.pack(fill="x", pady=(4, 10))

        ttk.Label(
            frame,
            text=(
                "Elige archivos o una carpeta; las caratulas quedaran etiquetadas con el "
                "banco seleccionado. Tambien puedes cargar Constancias de Situacion Fiscal: "
                "se detectan solas y se usan para completar el nombre y RFC de las cuentas. "
                "Puedes cambiar el banco y agregar mas; la ventana permanece abierta."
            ),
            wraplength=440,
            foreground="#555555",
            justify="left",
        ).pack(anchor="w", pady=(0, 12))

        btns = ttk.Frame(frame)
        btns.pack(fill="x")
        ttk.Button(btns, text="Elegir archivos...", command=self._loader_pick_files).pack(side="left")
        ttk.Button(btns, text="Elegir carpeta...", command=self._loader_pick_folder).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Cerrar", command=self._close_loader).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", self._close_loader)
        combo.focus_set()

    def _close_loader(self) -> None:
        if self._loader is not None:
            self._loader.destroy()
        self._loader = None

    def _loader_bank_hint(self) -> str:
        label = self._loader_bank_var.get()
        if label in (AUTO_BANK_LABEL, "Otro"):
            return ""
        return label

    def _loader_pick_files(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self._loader,
            title="Selecciona archivos PDF o imagenes",
            filetypes=[
                ("Documentos y imagenes", "*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.heic *.heif"),
                ("PDF", "*.pdf"),
                ("Imagenes", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.heic *.heif"),
                ("Todos", "*.*"),
            ],
        )
        if not paths:
            return
        self._add_supported_paths(list(paths), self._loader_bank_hint())

    def _loader_pick_folder(self) -> None:
        folder = filedialog.askdirectory(parent=self._loader, title="Selecciona una carpeta con caratulas")
        if not folder:
            return
        discovered: list[str] = []
        for root_dir, _dirs, files in os.walk(folder):
            for file_name in files:
                full_path = os.path.join(root_dir, file_name)
                if os.path.splitext(full_path)[1].lower() in SUPPORTED_EXTENSIONS:
                    discovered.append(full_path)
        added = self._add_supported_paths(discovered, self._loader_bank_hint())
        if added == 0:
            self.status_var.set("No se encontraron archivos compatibles en la carpeta.")

    def remove_selected(self) -> None:
        selection = list(self.file_list.curselection())
        if not selection:
            return
        for index in reversed(selection):
            if 0 <= index < len(self._selected_files):
                path = self._selected_files.pop(index)
                self._file_banks.pop(path, None)
        self._refresh_file_list()
        self.status_var.set("Archivos eliminados de la lista.")
        self._refresh_buttons()

    def clear_all(self) -> None:
        self._selected_files.clear()
        self._file_banks.clear()
        self._results.clear()
        self._refresh_file_list()
        self.tree.delete(*self.tree.get_children())
        self.detail_text.delete("1.0", "end")
        self.raw_text.delete("1.0", "end")
        self.summary_var.set("Sin archivos cargados")
        self.status_var.set("Lista limpiada.")
        self.progress.configure(value=0)
        self.progress_var.set("0%")
        self._refresh_buttons()

    def copy_selected_details(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("Copiar", "Selecciona un resultado para copiar su detalle.")
            return
        item_id = selection[0]
        index = self.tree.index(item_id)
        if not (0 <= index < len(self._results)):
            return
        result = self._results[index]
        payload = self._format_result_text(result)
        self.clipboard_clear()
        self.clipboard_append(payload)
        self.status_var.set("Resultado copiado al portapapeles.")

    def _refresh_file_list(self) -> None:
        self.file_list.delete(0, "end")
        for path in self._selected_files:
            bank = self._file_banks.get(path, "")
            etiqueta = bank if bank else "Auto"
            self.file_list.insert("end", f"{os.path.basename(path)}   [{etiqueta}]")
        # Al cambiar la lista se pierde la seleccion; reflejalo en el checkbox.
        if hasattr(self, "select_all_var"):
            self.select_all_var.set(False)

    def _toggle_select_all(self) -> None:
        if self.select_all_var.get():
            self.file_list.select_set(0, "end")
        else:
            self.file_list.select_clear(0, "end")
        self.file_list.focus_set()

    def run_extraction(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        if not self._selected_files:
            messagebox.showinfo("Analizar", "Primero agrega uno o mas archivos.")
            return

        self._set_busy(True)
        self.status_var.set("Analizando archivos...")
        self.detail_text.delete("1.0", "end")
        self.tree.delete(*self.tree.get_children())

        jobs = [(path, self._file_banks.get(path, "")) for path in self._selected_files]
        self.progress.configure(maximum=len(jobs), value=0)
        self.progress_var.set(f"0/{len(jobs)} (0%)")
        self._worker = threading.Thread(target=self._run_extraction_worker, args=(jobs,), daemon=True)
        self._worker.start()

    def _run_extraction_worker(self, jobs: list[tuple[str, str]]) -> None:
        total = len(jobs)
        results: list[ExtractionResult] = []
        constancias: list[ConstanciaFiscal] = []
        # Pasada unica: lee cada archivo, clasifica y va MOSTRANDO cada caratula en la
        # tabla conforme se procesa (las CSF no son filas, se guardan para emparejar).
        for index, (file_path, bank_hint) in enumerate(jobs, start=1):
            # Avisa ANTES de procesar (con el nombre del archivo) para que la barra
            # muestre actividad aunque un archivo lento (OCR) tarde varios segundos.
            self.after(0, lambda done=index - 1, t=total, fn=os.path.basename(file_path): self._update_progress(done, t, fn))
            try:
                raw_text, source_type, notes = extract_text(file_path)
            except Exception as exc:
                raw_text, source_type, notes = "", "Desconocido", [f"Error al leer el archivo: {exc}"]

            if _is_constancia(raw_text):
                constancias.append(_parse_constancia(raw_text, file_path))
            else:
                try:
                    result = _result_from_text(file_path, raw_text, source_type, notes, bank_hint)
                except Exception as exc:
                    result = ExtractionResult(
                        file_path=file_path,
                        source_type="Desconocido",
                        extracted_text=raw_text,
                        beneficiary_name="",
                        account_number="",
                        clabe="",
                        bank_name="",
                        rfc="",
                        clabe_is_valid=False,
                        notes=[f"Error general al procesar el archivo: {exc}"],
                    )
                results.append(result)
                self.after(0, lambda r=result: self._add_result_row(r))

            self.after(0, lambda i=index, t=total: self._update_progress(i, t))

        # Al final, con todas las constancias leidas, emparejamos y refrescamos la tabla
        # para corregir nombres/RFC y llenar la columna de constancia.
        for result in results:
            _enrich_with_constancias(result, constancias)

        self.after(0, lambda r=results, n=len(constancias): self._finish_extraction(r, n))

    def _result_values(self, result: ExtractionResult) -> tuple:
        return (
            result.file_name,
            result.beneficiary_name or "-",
            result.account_number or "-",
            result.clabe or "-",
            "Si" if result.clabe_is_valid else "No",
            result.bank_name or "-",
            result.rfc or "-",
            result.csf_summary or "-",
            result.status,
        )

    def _add_result_row(self, result: ExtractionResult) -> None:
        """Inserta una caratula en la tabla apenas se procesa (avance visible)."""
        self.tree.insert("", "end", values=self._result_values(result))

    def _update_progress(self, done: int, total: int, current: str = "") -> None:
        self.progress.configure(value=done)
        percent = int(done / total * 100) if total else 0
        self.progress_var.set(f"{done}/{total} ({percent}%)")
        if current:
            self.status_var.set(f"Analizando {min(done + 1, total)}/{total}: {current}")
        else:
            self.status_var.set(f"Procesando {done}/{total}...")

    def _finish_extraction(self, results: list[ExtractionResult], constancias_count: int = 0) -> None:
        self._results = results
        self.tree.delete(*self.tree.get_children())

        valid_clabes = 0
        for result in results:
            if result.clabe_is_valid:
                valid_clabes += 1
            self.tree.insert("", "end", values=self._result_values(result))

        resumen = f"Caratulas: {len(results)} | CLABE validas: {valid_clabes}"
        if constancias_count:
            resumen += f" | Constancias fiscales: {constancias_count}"
        resumen += f" | OCR: {'activo' if _ocr_available() else 'no disponible'}"
        self.summary_var.set(resumen)
        total = int(self.progress["maximum"])
        self.progress.configure(value=total)
        self.progress_var.set(f"{total}/{total} (100%)")
        if constancias_count:
            self.status_var.set(
                f"Analisis completado. Se usaron {constancias_count} constancia(s) fiscal(es) para completar datos."
            )
        else:
            self.status_var.set("Analisis completado.")
        if results:
            self._show_result_details(results[0])
        self._set_busy(False)
        self._refresh_buttons()

    def _export_rows(self) -> list[dict[str, str]]:
        return [result.to_export_row() for result in self._results]

    def export_csv(self) -> None:
        if not self._results:
            messagebox.showinfo("Exportar CSV", "Primero analiza uno o mas archivos.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Guardar CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="extractor_bancario.csv",
        )
        if not file_path:
            return

        rows = self._export_rows()
        fieldnames = list(rows[0].keys())
        with open(file_path, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        self.status_var.set(f"CSV exportado en {file_path}.")

    def export_excel(self) -> None:
        if not self._results:
            messagebox.showinfo("Exportar Excel", "Primero analiza uno o mas archivos.")
            return
        if Workbook is None:
            messagebox.showerror("Exportar Excel", "Falta openpyxl para generar archivos .xlsx.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Guardar Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="extractor_bancario.xlsx",
        )
        if not file_path:
            return

        rows = self._export_rows()
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Resultados"

        headers = list(rows[0].keys())
        worksheet.append(headers)
        for row in rows:
            worksheet.append([row.get(header, "") for header in headers])

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > max_length:
                    max_length = len(value)
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 60)

        workbook.save(file_path)
        self.status_var.set(f"Excel exportado en {file_path}.")

    def show_selected_details(self, _event: object) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        index = self.tree.index(selection[0])
        if not (0 <= index < len(self._results)):
            return
        result = self._results[index]
        self._show_result_details(result)

    def _show_result_details(self, result: ExtractionResult) -> None:
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("end", self._format_result_text(result))
        self.raw_text.delete("1.0", "end")
        raw_value = result.extracted_text or "(sin texto OCR)"
        self.raw_text.insert("end", raw_value)
        self._highlight_raw_text(result)

    def _highlight_raw_text(self, result: ExtractionResult) -> None:
        self.raw_text.tag_remove("highlight_label", "1.0", "end")
        self.raw_text.tag_remove("highlight_value", "1.0", "end")
        self.raw_text.tag_remove("highlight_name", "1.0", "end")

        label_terms = [
            "clabe",
            "rfc",
            "banco",
            "cuenta",
            "beneficiario",
            "titular",
            "cuentahabiente",
        ]
        for term in label_terms:
            start = "1.0"
            while True:
                match = self.raw_text.search(term, start, stopindex="end", nocase=True)
                if not match:
                    break
                end = f"{match}+{len(term)}c"
                self.raw_text.tag_add("highlight_label", match, end)
                start = end

        value_terms = [
            result.clabe,
            _clean_digits(result.account_number),
            result.rfc,
            result.bank_name,
            result.beneficiary_name,
        ]
        for term in value_terms:
            normalized_term = _normalize_spaces(term or "")
            if not normalized_term or normalized_term == "sin texto ocr":
                continue
            start = "1.0"
            while True:
                match = self.raw_text.search(normalized_term, start, stopindex="end", nocase=True)
                if not match:
                    break
                end = f"{match}+{len(normalized_term)}c"
                if normalized_term == _normalize_spaces(result.beneficiary_name or ""):
                    self.raw_text.tag_add("highlight_name", match, end)
                else:
                    self.raw_text.tag_add("highlight_value", match, end)
                start = end

    def _format_result_text(self, result: ExtractionResult) -> str:
        lines = [
            f"Archivo: {result.file_path}",
            f"Tipo: {result.source_type}",
            f"Beneficiario: {result.beneficiary_name or 'No identificado'}",
            f"Cuenta: {result.account_number or 'No identificada'}",
            f"CLABE: {result.clabe or 'No identificada'}",
            f"CLABE valida: {'Si' if result.clabe_is_valid else 'No'}",
            f"Banco: {result.bank_name or 'No identificado'}",
            f"RFC: {result.rfc or 'No identificado'}",
            f"Constancia fiscal (CSF): {result.csf_summary or 'Sin emparejar'}",
            "",
            "Notas:",
        ]
        if result.notes:
            lines.extend(f"- {note}" for note in result.notes)
        else:
            lines.append("- Sin observaciones")
        lines.extend(["", "Texto detectado:", result.extracted_text or "(vacío)"])
        return "\n".join(lines)

    def _set_busy(self, busy: bool) -> None:
        self._worker = self._worker if busy else None
        self._refresh_buttons()


def main() -> None:
    app = BankExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
