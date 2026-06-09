# RPA – Solicitudes de Pago (SIPP Petroil)

Robot que registra **solicitudes de pago a ex-colaboradores** (tipo *Pago
Extraordinario*, beneficiario *Acreedor*) en SIPP, leyendo el archivo
**BASE DE DATOS LISTADO 1.csv**.

Por cada renglón con monto, el robot: busca-o-crea al acreedor, llena sus datos
y cuenta bancaria (CLABE/SPEI + carátula), captura el concepto y el importe,
**guarda**, adjunta el **Vo.Bo.** y da clic en **Solicitar Autorización**.

---

## 1. Requisitos (una sola vez)

```powershell
pip install -r requirements.txt
playwright install chromium
```

## 2. Datos necesarios (junto a este archivo)

- **`BASE DE DATOS LISTADO 1.csv`** — con la columna **MONTO** llena en los que
  se van a pagar (los que la dejes vacía, el robot los **omite**).
- Carpeta **`CARATULAS`** — un PDF/imagen por colaborador, nombrado con su nombre
  (ej. `ANGEL MORENO AYALA.pdf`).
- Carpeta **`VOBO`** — el Vo.Bo. de Compras, igual nombrado por colaborador.
  Si alguien no tiene Vo.Bo., el robot lo omite (no truena).

Para revisar que el CSV está bien antes de correr:
```powershell
python validar_csv.py
```

## 3. Configuración (`config.py`)

Banderas de seguridad:
- `AMBIENTE` → `"PRUEBAS"` (stage) o `"PRODUCCION"`.
- `MAX_REGISTROS` → cuántos procesar (`1` para probar, `None` para todos).
- `PAUSAR_ANTES_DE_GUARDAR` → `True`: llena todo y se detiene para que **tú**
  guardes a mano (revisas antes). `False`: el robot guarda, adjunta Vo.Bo. y
  solicita autorización solo.
- `NAVEGADOR_VISIBLE` → `True` para ver el navegador trabajar.
- `CONCEPTO_PAGO` → `"PAGO PTU"` (en producción). *Ojo: en stage no existe ese
  concepto; por eso allá se usa otro.*

Credenciales: en `config.py` (o variables de entorno `SIPP_USUARIO` /
`SIPP_CONTRASENA`).

## 4. Ejecutar

```powershell
python sipp_rpa.py
```

Al final, el robot reporta cuatro listas: **Éxitos**, **Errores**,
**Revisar** (acreedores que ya existían), **Omitidos** (sin monto). Todo queda
en la carpeta `logs/` (con capturas si hubo errores).

---

## 5. Cómo hacer la corrida REAL en producción (recomendado)

1. En `config.py`: pon `AMBIENTE = "PRODUCCION"` y confirma
   `CONCEPTO_PAGO = "PAGO PTU"`.
2. **Primero 1 o 2 con pausa:** deja `PAUSAR_ANTES_DE_GUARDAR = True` y
   `MAX_REGISTROS = 1`. Corre, revisa el formulario lleno, guarda tú el clic
   final y verifica el resultado en SIPP.
3. **Luego automático:** si todo salió bien, pon
   `PAUSAR_ANTES_DE_GUARDAR = False` y `MAX_REGISTROS = None` (o un número) y
   vuelve a correr para procesar el resto.

> ⏱️ Cada registro tarda ~2 min (la subida de la carátula es lo más lento), así
> que el lote completo puede tomar ~1.5 horas. Es normal; deja correr.

---

## 6. Notas / pendientes conocidos

- **Acreedor ya existente:** el robot NO lo duplica; lo marca en *Revisar* para
  que lo atiendas a mano. (El auto-seleccionar un acreedor existente aún no se
  programa porque el lote actual son todos nuevos.)
- **Transferencias:** todas se tratan como **SPEI** (por indicación del área).
- **Diferencias de stage:** el ambiente de pruebas no tiene el concepto
  "PAGO PTU" ni siempre logra subir archivos (error intermitente de Google);
  en producción esto no aplica.

## 7. Archivos del proyecto

| Archivo | Qué es |
|---|---|
| `sipp_rpa.py` | El robot (flujo completo) |
| `config.py` | Configuración y banderas |
| `mapeos.py` | Traducción CSV → catálogos de SIPP |
| `validar_csv.py` | Valida el CSV sin abrir el navegador |
| `requirements.txt` | Dependencias |
| `prueba_*.py` / `explorar_*.py` | Scripts de prueba/diagnóstico usados en el desarrollo |
