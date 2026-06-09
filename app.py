# -*- coding: utf-8 -*-
"""
RPA Solicitudes de Pago (SIPP) — Interfaz gráfica
=================================================
App de escritorio para que un operador: cargue el CSV, vea una vista previa
validada, escriba sus credenciales y ejecute el robot automáticamente.

Apunta SIEMPRE a PRODUCCIÓN. Para correr:
    python app.py
"""

import os
import sys
import queue
import logging
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def ruta_recurso(rel):
    """Ruta a un recurso (logo, navegador) que funciona corriendo como script y
    como .exe empaquetado con PyInstaller (que usa sys._MEIPASS)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# Navegador Chromium: en el .exe va EMPAQUETADO dentro (carpeta 'ms-playwright');
# corriendo como script usa el instalado en el equipo. Debe fijarse ANTES de
# importar sipp_rpa (que importa playwright).
if getattr(sys, "frozen", False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ruta_recurso("ms-playwright")

import config
import sipp_rpa

# ===================================================================== #
#  AMBIENTE DEL EJECUTABLE
#  "PRODUCCION" = versión final | "PRUEBAS" = preprod (para probar)
# ===================================================================== #
AMBIENTE_APP = "PRODUCCION"
# ===================================================================== #

config.AMBIENTE = AMBIENTE_APP
config.URL_LOGIN = config.URLS[AMBIENTE_APP]
# En pruebas (preprod) NO existe "PAGO PTU"; se usa uno que sí existe allá.
config.CONCEPTO_PAGO = ("PAGO PTU" if AMBIENTE_APP == "PRODUCCION"
                        else "PAGO DE NOMINA")
config.PAUSAR_ANTES_DE_GUARDAR = False

COLOR_OK = "#1a7f37"
COLOR_MAL = "#bb0000"


class HandlerCola(logging.Handler):
    """Manda los mensajes del log a una cola para mostrarlos en la ventana."""
    def __init__(self, cola):
        super().__init__()
        self.cola = cola

    def emit(self, record):
        self.cola.put(("log", self.format(record)))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"RPA Solicitudes de Pago — SIPP ({config.AMBIENTE})")
        self.geometry("960x720")
        self.minsize(860, 640)

        self.filas = None
        self.ruta_csv = None
        self.cola = queue.Queue()
        self.detener_flag = threading.Event()
        self.worker = None
        # Control de avance / reanudar.
        self.siguiente = 0                 # índice del próximo registro a procesar
        self.resultados_global = []        # detalle acumulado entre corridas
        self.usuario = ""
        self.contrasena = ""
        self.estado = "idle"               # idle|corriendo|detenido|fin

        self._construir()
        self._enganchar_log()
        self.after(150, self._drenar_cola)

    # ------------------------------------------------------------------ #
    #  Construcción de la interfaz
    # ------------------------------------------------------------------ #
    def _construir(self):
        pad = {"padx": 8, "pady": 4}

        # --- Pie de página (parte inferior) ---
        tk.Label(self, text="Hecho por Quetaltic Solutions", fg="#888",
                 font=("Segoe UI", 8)).pack(side="bottom", pady=4)

        # --- Encabezado: logo + título en la MISMA fila ---
        cab = tk.Frame(self, bg="white")
        cab.pack(fill="x")
        try:
            img = tk.PhotoImage(file=ruta_recurso(os.path.join("Imagenes", "Logo.png")))
            factor = max(1, img.height() // 48)  # ~48 px de alto
            self.logo_img = img.subsample(factor, factor)
            tk.Label(cab, image=self.logo_img, bg="white"
                     ).pack(side="left", padx=(12, 8), pady=6)
        except Exception:
            pass
        tk.Label(cab, text="RPA Solicitudes de Pago — SIPP", bg="white",
                 fg="#00437f", font=("Segoe UI", 14, "bold")
                 ).pack(side="left", pady=8)
        amb_color = "#bb0000" if config.AMBIENTE == "PRODUCCION" else "#b06000"
        tk.Label(cab, text=f"Ambiente: {config.AMBIENTE}   ", bg="white",
                 fg=amb_color, font=("Segoe UI", 9, "bold")).pack(side="right", pady=8)
        # Línea de acento.
        tk.Frame(self, bg="#00437f", height=3).pack(fill="x")

        # --- Paso 1: Archivos (CSV + carpetas de carátulas y Vo.Bo.) ---
        f1 = tk.LabelFrame(self, text="1) Archivos", **pad)
        f1.pack(fill="x", **pad)
        tk.Button(f1, text="Cargar CSV...", width=22, command=self.cargar_csv
                  ).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.lbl_csv = tk.Label(f1, text="(ningún archivo cargado)", fg="#555")
        self.lbl_csv.grid(row=0, column=1, padx=6, sticky="w")
        tk.Button(f1, text="Carpeta de Carátulas...", width=22,
                  command=self.seleccionar_caratulas
                  ).grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.lbl_car = tk.Label(f1, text="(no seleccionada — requerida)", fg="#555")
        self.lbl_car.grid(row=1, column=1, padx=6, sticky="w")
        tk.Button(f1, text="Carpeta de Vo.Bo....", width=22,
                  command=self.seleccionar_vobo
                  ).grid(row=2, column=0, padx=6, pady=4, sticky="w")
        self.lbl_vobo = tk.Label(f1, text="(opcional — no todos lo tienen)", fg="#555")
        self.lbl_vobo.grid(row=2, column=1, padx=6, sticky="w")

        # --- Vista previa ---
        f2 = tk.LabelFrame(self, text="Vista previa", **pad)
        f2.pack(fill="both", expand=True, **pad)
        self.lbl_resumen = tk.Label(f2, text="Carga un CSV para ver el resumen.",
                                    anchor="w", justify="left")
        self.lbl_resumen.pack(fill="x")
        cols = ("col", "empresa", "banco", "monto", "obs")
        self.tabla = ttk.Treeview(f2, columns=cols, show="headings", height=10)
        for c, t, w in [("col", "Colaborador", 240), ("empresa", "Empresa", 110),
                        ("banco", "Banco", 110), ("monto", "Monto", 100),
                        ("obs", "Observaciones", 320)]:
            self.tabla.heading(c, text=t)
            self.tabla.column(c, width=w, anchor="w")
        self.tabla.tag_configure("mal", foreground=COLOR_MAL)
        self.tabla.tag_configure("omit", foreground="#888")
        sb = ttk.Scrollbar(f2, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=sb.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # --- Paso 2: credenciales y opciones ---
        f3 = tk.LabelFrame(self, text="2) Credenciales SIPP", **pad)
        f3.pack(fill="x", **pad)
        tk.Label(f3, text="Usuario:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.ent_usuario = tk.Entry(f3, width=24)
        self.ent_usuario.grid(row=0, column=1, padx=4, pady=4)
        tk.Label(f3, text="Contraseña:").grid(row=0, column=2, sticky="e", padx=4, pady=4)
        self.ent_pass = tk.Entry(f3, width=24, show="*")
        self.ent_pass.grid(row=0, column=3, padx=4, pady=4)
        self.var_visible = tk.BooleanVar(value=True)
        tk.Checkbutton(f3, text="Mostrar navegador", variable=self.var_visible
                       ).grid(row=0, column=4, padx=12)

        # --- Paso 3: ejecutar ---
        f4 = tk.Frame(self)
        f4.pack(fill="x", **pad)
        self.btn_iniciar = tk.Button(f4, text="▶  Iniciar", width=12,
                                     bg="#00437f", fg="white",
                                     font=("Segoe UI", 10, "bold"),
                                     command=self.iniciar, state="disabled")
        self.btn_iniciar.pack(side="left", padx=4)
        self.btn_detener = tk.Button(f4, text="■  Detener", width=11,
                                     command=self.detener, state="disabled")
        self.btn_detener.pack(side="left", padx=4)
        self.btn_reanudar = tk.Button(f4, text="⏵ Reanudar", width=12,
                                      command=self.reanudar, state="disabled")
        self.btn_reanudar.pack(side="left", padx=4)
        self.btn_reiniciar = tk.Button(f4, text="↺ Reiniciar carga", width=15,
                                       command=self.reiniciar, state="disabled")
        self.btn_reiniciar.pack(side="left", padx=4)
        self.btn_reporte = tk.Button(f4, text="📄 Generar reporte", width=16,
                                     command=self.generar_reporte, state="disabled")
        self.btn_reporte.pack(side="left", padx=4)
        self.barra = ttk.Progressbar(f4, mode="determinate")
        self.barra.pack(side="left", fill="x", expand=True, padx=10)
        self.lbl_estado = tk.Label(f4, text="", width=20, anchor="w")
        self.lbl_estado.pack(side="left")

        # --- Log ---
        f5 = tk.LabelFrame(self, text="Avance", **pad)
        f5.pack(fill="both", expand=True, **pad)
        self.txt = tk.Text(f5, height=10, state="disabled", wrap="word",
                           font=("Consolas", 9))
        sb2 = ttk.Scrollbar(f5, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb2.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")

    def _enganchar_log(self):
        h = HandlerCola(self.cola)
        h.setFormatter(logging.Formatter("%(asctime)s | %(message)s", "%H:%M:%S"))
        logging.getLogger("rpa").addHandler(h)

    # ------------------------------------------------------------------ #
    #  Cargar y validar el CSV
    # ------------------------------------------------------------------ #
    def cargar_csv(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el CSV", filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if not ruta:
            return
        carpeta = os.path.dirname(ruta)
        config.ARCHIVO_CSV = ruta
        config.CARPETA_LOGS = os.path.join(carpeta, "logs")
        try:
            self.filas = sipp_rpa.leer_csv(ruta)
        except Exception as e:
            messagebox.showerror("Error al leer el CSV", str(e))
            return
        self.ruta_csv = ruta
        self.lbl_csv.config(text=os.path.basename(ruta), fg="#000")

        # Autodetecta carpetas junto al CSV (el operador puede cambiarlas).
        car = os.path.join(carpeta, "CARATULAS")
        vob = os.path.join(carpeta, "VOBO")
        if os.path.isdir(car):
            config.CARPETA_CARATULAS = car
            self._lbl_carpeta(self.lbl_car, car)
        if os.path.isdir(vob):
            config.CARPETA_VOBO = vob
            self._lbl_carpeta(self.lbl_vobo, vob)

        # CSV nuevo: estado limpio (desde cero).
        self.estado = "idle"
        self.siguiente = 0
        self.resultados_global = []
        self._mostrar_preview()
        self._actualizar_botones()

    def _lbl_carpeta(self, lbl, carpeta):
        n = 0
        if os.path.isdir(carpeta):
            n = sum(1 for f in os.listdir(carpeta)
                    if os.path.splitext(f)[1].lower() in config.EXT_CARATULA)
        lbl.config(text=f"{carpeta}  ({n} archivos)", fg="#000")

    def seleccionar_caratulas(self):
        d = filedialog.askdirectory(title="Carpeta de Carátulas")
        if not d:
            return
        config.CARPETA_CARATULAS = d
        self._lbl_carpeta(self.lbl_car, d)
        if self.filas:
            self._mostrar_preview()
        self._actualizar_botones()

    def seleccionar_vobo(self):
        d = filedialog.askdirectory(title="Carpeta de Vo.Bo.")
        if not d:
            return
        config.CARPETA_VOBO = d
        self._lbl_carpeta(self.lbl_vobo, d)
        if self.filas:
            self._mostrar_preview()

    def _mostrar_preview(self):
        for it in self.tabla.get_children():
            self.tabla.delete(it)
        v = sipp_rpa.validar_datos(self.filas)
        probs = {n: e for n, e in v["problemas"]}
        for fila in self.filas:
            nombre = sipp_rpa.campo(fila, "EX-COLABORADOR (DESCRIPCIÓN)", "NOMBRE DE CUENTA")
            empresa = sipp_rpa.campo(fila, "EMPRESA")
            banco = sipp_rpa.campo(fila, "BANCOS")
            monto = sipp_rpa.campo(fila, "MONTO").strip()
            tags = ()
            obs = ""
            if not sipp_rpa.limpiar_monto(monto):
                obs = "Sin monto (se omite)"
                tags = ("omit",)
            elif nombre in probs:
                obs = "; ".join(probs[nombre])
                tags = ("mal",)
            self.tabla.insert("", "end",
                              values=(nombre, empresa, banco, monto or "—", obs),
                              tags=tags)
        car = "OK" if v["hay_caratulas"] else "carpeta CARATULAS no encontrada"
        vob = "OK" if v["hay_vobo"] else "carpeta VOBO no encontrada"
        nprob = len(v["problemas"])
        self.lbl_resumen.config(
            text=(f"Total: {v['total']}   |   Con monto: {v['con_monto']}   |   "
                  f"Sin monto (se omiten): {v['sin_monto']}   |   "
                  f"Con observaciones: {nprob}\n"
                  f"Carátulas: {car}   |   Vo.Bo.: {vob}"),
            fg=(COLOR_MAL if nprob else COLOR_OK))

    # ------------------------------------------------------------------ #
    #  Ejecutar el robot
    # ------------------------------------------------------------------ #
    def _actualizar_botones(self, *_):
        listo = (self.filas is not None
                 and os.path.isdir(config.CARPETA_CARATULAS))
        e = self.estado
        self.btn_iniciar.config(
            state="normal" if (listo and e == "idle") else "disabled")
        self.btn_detener.config(
            state="normal" if e == "corriendo" else "disabled")
        self.btn_reanudar.config(
            state="normal" if e == "detenido" else "disabled")
        self.btn_reiniciar.config(
            state="normal" if (listo and e in ("detenido", "fin")) else "disabled")
        # 'Generar reporte' disponible mientras NO esté corriendo (al detener,
        # al terminar, o tras cargar). Si no hay datos, el botón lo avisa.
        self.btn_reporte.config(
            state="normal" if (self.filas is not None and e != "corriendo")
            else "disabled")

    def generar_reporte(self):
        if not self.resultados_global:
            messagebox.showinfo(
                "Sin datos",
                "Todavía no hay registros procesados, así que no se generó "
                "ningún reporte.")
            return
        rep = sipp_rpa.escribir_reporte(self.resultados_global)
        if not rep:
            messagebox.showerror("Error", "No se pudo generar el reporte.")
            return
        messagebox.showinfo(
            "Reporte generado",
            f"Se generó el reporte con {len(self.resultados_global)} "
            f"registro(s):\n\n{rep}")
        if os.path.exists(rep) and messagebox.askyesno("Reporte", "¿Deseas abrirlo?"):
            try:
                os.startfile(rep)
            except Exception:
                pass

    def _credenciales(self):
        u = self.ent_usuario.get().strip()
        p = self.ent_pass.get()
        if not u or not p:
            messagebox.showwarning("Faltan credenciales",
                                   "Escribe tu usuario y contraseña de SIPP.")
            return None
        return u, p

    def _confirmar_inicio(self):
        con_monto = sum(1 for f in self.filas
                        if sipp_rpa.limpiar_monto(sipp_rpa.campo(f, "MONTO")))
        faltan = [sipp_rpa.campo(f, "EX-COLABORADOR (DESCRIPCIÓN)", "NOMBRE DE CUENTA")
                  for f in self.filas
                  if sipp_rpa.limpiar_monto(sipp_rpa.campo(f, "MONTO"))
                  and not sipp_rpa.buscar_caratula(
                      sipp_rpa.campo(f, "EX-COLABORADOR (DESCRIPCIÓN)", "NOMBRE DE CUENTA"))]
        if faltan and not messagebox.askyesno(
                "Faltan carátulas",
                f"{len(faltan)} registro(s) con monto NO tienen carátula y "
                f"fallarán al guardar.\nEjemplo: {faltan[0]}\n\n"
                f"¿Continuar de todas formas?"):
            return False
        return messagebox.askyesno(
            f"Confirmar ejecución en {config.AMBIENTE}",
            f"Se registrarán {con_monto} solicitudes de pago en "
            f"{config.AMBIENTE}.\n\n¿Deseas continuar?")

    def iniciar(self):
        cred = self._credenciales()
        if not cred or not self._confirmar_inicio():
            return
        self.usuario, self.contrasena = cred
        self.siguiente = 0
        self.resultados_global = []
        self._log_ui("=== Iniciando desde el primer registro ===")
        self._arrancar()

    def reanudar(self):
        cred = self._credenciales()
        if not cred:
            return
        self.usuario, self.contrasena = cred
        self._log_ui(f"=== Reanudando desde el registro {self.siguiente + 1} ===")
        self._arrancar()

    def reiniciar(self):
        cred = self._credenciales()
        if not cred:
            return
        if not messagebox.askyesno(
                "Reiniciar carga",
                "Se volverá a empezar desde el PRIMER registro.\n\n"
                "OJO: los registros ya guardados en una corrida previa NO se "
                "borran; si los vuelves a procesar podrían duplicarse.\n\n"
                "¿Continuar?"):
            return
        self.usuario, self.contrasena = cred
        self.siguiente = 0
        self.resultados_global = []
        self._log_ui("=== Reiniciando carga desde el primer registro ===")
        self._arrancar()

    def _arrancar(self):
        config.NAVEGADOR_VISIBLE = self.var_visible.get()
        self.detener_flag.clear()
        self.estado = "corriendo"
        self._actualizar_botones()
        self.barra.config(maximum=max(1, len(self.filas)), value=self.siguiente)
        inicio = self.siguiente
        self.worker = threading.Thread(
            target=self._correr, args=(inicio,), daemon=True)
        self.worker.start()

    def _correr(self, inicio):
        try:
            resumen = sipp_rpa.procesar(
                self.usuario, self.contrasena, self.filas[inicio:],
                on_progreso=lambda d: self.cola.put(("prog", {**d, "offset": inicio})),
                detener=self.detener_flag.is_set,
                escribir_rep=False)
            self.cola.put(("fin", resumen))
        except Exception as e:
            self.cola.put(("error", str(e)))

    def detener(self):
        self.detener_flag.set()
        self.lbl_estado.config(text="Deteniendo...")

    # ------------------------------------------------------------------ #
    #  Comunicación hilo -> ventana (cola)
    # ------------------------------------------------------------------ #
    def _log_ui(self, msg):
        self.txt.config(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _drenar_cola(self):
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "log":
                    self._log_ui(dato)
                elif tipo == "prog":
                    # Acumula el resultado EN VIVO (para poder generar el reporte
                    # en cualquier momento, incluso si se detiene).
                    res = dato.get("resultado")
                    if res:
                        self.resultados_global.append(res)
                        self.siguiente = len(self.resultados_global)
                    absoluto = dato.get("offset", 0) + dato.get("i", 0)
                    self.barra.config(value=absoluto)
                    self.lbl_estado.config(
                        text=f"{absoluto}/{len(self.filas)} {dato.get('estado','')}")
                elif tipo == "fin":
                    self._terminar(dato)
                elif tipo == "error":
                    self._terminar(None, error=dato)
        except queue.Empty:
            pass
        self.after(150, self._drenar_cola)

    def _terminar(self, resumen, error=None):
        if error:
            self.estado = "detenido" if self.siguiente > 0 else "idle"
            self._actualizar_botones()
            self.lbl_estado.config(text="Error")
            messagebox.showerror("Error", f"El proceso falló:\n{error}")
            return
        r = resumen
        # Los resultados ya se fueron acumulando EN VIVO (self.resultados_global);
        # self.siguiente = cuántos se han procesado en total.
        detenido = r.get("detenido", False) and self.siguiente < len(self.filas)
        self.estado = "detenido" if detenido else "fin"
        self._actualizar_botones()

        cont = {}
        for x in self.resultados_global:
            cont[x["estado"]] = cont.get(x["estado"], 0) + 1
        estado_txt = "Detenido" if detenido else "Terminado"
        self.lbl_estado.config(text=estado_txt)
        rep = sipp_rpa.escribir_reporte(self.resultados_global)
        msg = (f"{estado_txt}. Procesados {self.siguiente} de {len(self.filas)}.\n\n"
               f"Éxitos: {cont.get('OK', 0)}\n"
               f"Cancelados (sin concepto): {cont.get('CANCELADO', 0)}\n"
               f"Revisar: {cont.get('REVISAR', 0)}\n"
               f"Errores: {cont.get('ERROR', 0)}\n"
               f"Omitidos (sin monto): {cont.get('OMITIDO', 0)}")
        if detenido:
            msg += ("\n\nProceso detenido. Usa 'Reanudar' para continuar desde "
                    "donde quedó, o 'Reiniciar carga' para empezar de nuevo.")
        if rep:
            msg += f"\n\nReporte:\n{rep}"
        messagebox.showinfo("Resultado", msg)
        if rep and os.path.exists(rep) and not detenido and messagebox.askyesno(
                "Reporte", "¿Deseas abrir el reporte final?"):
            try:
                os.startfile(rep)
            except Exception:
                pass


if __name__ == "__main__":
    App().mainloop()
