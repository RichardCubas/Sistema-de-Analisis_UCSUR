# vistas/vista_registro.py
# ============================================================
# FRAME — REPORTE DE REGISTRO DE NOTAS (UCSUR)
# Migración COMPLETA desde registro_notas_GUI.py
# ============================================================

import os
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image
import pandas as pd

# 🔹 Script ORIGINAL (NO SE TOCA)
import Reporte_registro_notas as rn


# ==========================
# ESTILO / COLORES
# ==========================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COL_UCSUR_AZUL = "#003B70"
COL_BG = "#F5F7FA"
COL_CARD = "white"
COL_OK = "#1C7C54"
COL_ERR = "#AA0000"


class FrameRegistroNotas(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COL_BG)

        self.controller = controller

        # ==========================
        # ESTADO
        # ==========================
        self.df = None
        self.cursos = []
        self.evals_globales = []

        self.modo_var = ctk.StringVar(value="CURSO")   # "CURSO" | "GLOBAL"
        self.curso_var = ctk.StringVar(value="")
        self.eval_var = ctk.StringVar(value="")

        # porcentaje como número (slider manda float)
        self.porcentaje_num = ctk.DoubleVar(value=80.0)

        self._logo_img = None

        # ==========================
        # LAYOUT BASE
        # ==========================
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_body()
        self._cargar_datos_iniciales()

    # ==================================================
    # HEADER
    # ==================================================
    def _build_header(self):
        header = ctk.CTkFrame(self, height=170, fg_color=COL_CARD, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(0, weight=1)

        inner = ctk.CTkFrame(header, fg_color=COL_CARD, corner_radius=0)
        inner.grid(row=0, column=0, sticky="nsew", pady=(6, 0))
        inner.columnconfigure(0, weight=1)

        logo_path = getattr(rn, "LOGO_FILE", "logo_ucsur.png")
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((320, 150), Image.LANCZOS)
                self._logo_img = ctk.CTkImage(light_image=img, size=img.size)
                ctk.CTkLabel(inner, image=self._logo_img, text="").grid(row=0, column=0)
            except Exception:
                pass

        ctk.CTkLabel(
            inner,
            text="Reporte de Registro de Notas",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=1, column=0, pady=(2, 0))

        ctk.CTkLabel(
            inner,
            text="Departamento de Cursos Básicos - Equipo de Matemática",
            font=ctk.CTkFont(size=14),
            text_color="#555"
        ).grid(row=2, column=0)

        ctk.CTkFrame(header, fg_color=COL_UCSUR_AZUL, height=3).grid(
            row=1, column=0, sticky="ew", pady=(8, 0)
        )

    # ==================================================
    # BODY
    # ==================================================
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color=COL_BG)
        body.grid(row=1, column=0, sticky="nsew", padx=18, pady=14)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # --------------------------
        # PANEL IZQUIERDO (config)
        # --------------------------
        left = ctk.CTkFrame(body, fg_color=COL_CARD, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(1, weight=1)

        # Tipo
        ctk.CTkLabel(
            left, text="Tipo de reporte:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=0, column=0, padx=15, pady=(18, 8), sticky="w")

        ctk.CTkRadioButton(
            left, text="Por curso",
            variable=self.modo_var, value="CURSO",
            command=self._on_modo_change
        ).grid(row=1, column=0, padx=20, sticky="w")

        ctk.CTkRadioButton(
            left, text="Global (todos los cursos)",
            variable=self.modo_var, value="GLOBAL",
            command=self._on_modo_change
        ).grid(row=1, column=1, padx=20, sticky="w")

        # Curso
        ctk.CTkLabel(
            left, text="Curso:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=2, column=0, padx=15, pady=(18, 6), sticky="w")

        self.combo_curso = ctk.CTkComboBox(
            left,
            values=[],
            state="disabled",
            variable=self.curso_var,
            command=self._on_curso_change
        )
        self.combo_curso.grid(row=2, column=1, padx=15, pady=(18, 6), sticky="ew")

        # Evaluación
        ctk.CTkLabel(
            left, text="Evaluación:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=3, column=0, padx=15, pady=6, sticky="w")

        self.combo_eval = ctk.CTkComboBox(
            left, values=[], variable=self.eval_var
        )
        self.combo_eval.grid(row=3, column=1, padx=15, pady=6, sticky="ew")

        # Porcentaje (Slider)
        ctk.CTkLabel(
            left, text="Porcentaje mínimo:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=4, column=0, padx=15, pady=(10, 6), sticky="w")

        self.lbl_porc = ctk.CTkLabel(
            left, text="80%",
            font=ctk.CTkFont(size=13),
            text_color="#333"
        )
        self.lbl_porc.grid(row=4, column=1, padx=15, pady=(10, 6), sticky="w")

        self.slider_p = ctk.CTkSlider(
            left,
            from_=0, to=100,
            number_of_steps=100,
            variable=self.porcentaje_num,
            command=self._on_slider_change,
            progress_color=COL_UCSUR_AZUL
        )
        self.slider_p.grid(row=5, column=0, columnspan=2, padx=15, pady=(0, 16), sticky="ew")

        # --------------------------
        # PANEL DERECHO (resumen + export)
        # --------------------------
        right = ctk.CTkFrame(body, fg_color=COL_CARD, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            right, text="Resumen",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=0, column=0, padx=15, pady=(15, 6), sticky="w")

        self.txt_resumen = ctk.CTkTextbox(right)
        self.txt_resumen.grid(row=1, column=0, padx=15, pady=6, sticky="nsew")

        btns = ctk.CTkFrame(right, fg_color=COL_CARD)
        btns.grid(row=2, column=0, pady=10)
        btns.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btns,
            text="Exportar PDF",
            fg_color=COL_UCSUR_AZUL,
            hover_color="#00294C",
            command=self._exportar_pdf
        ).grid(row=0, column=0, padx=10)

        ctk.CTkButton(
            btns,
            text="Exportar Excel",
            fg_color=COL_OK,
            hover_color="#239966",
            command=self._exportar_excel
        ).grid(row=0, column=1, padx=10)

        self.lbl_status = ctk.CTkLabel(
            right,
            text="Listo.",
            font=ctk.CTkFont(size=12),
            text_color="#555"
        )
        self.lbl_status.grid(row=3, column=0, padx=15, pady=(0, 12), sticky="w")

    # ==================================================
    # CARGA DE DATOS (igual al antiguo)
    # ==================================================
    def _cargar_datos_iniciales(self):
        try:
            self.df = rn.cargar_df()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar los datos:\n{e}")
            self.lbl_status.configure(text="Error cargando datos.", text_color=COL_ERR)
            return

        self.cursos = sorted(self.df["Curso"].dropna().unique().tolist())
        self.combo_curso.configure(values=self.cursos)

        # evals globales con orden si existe
        evals = sorted(set(self.df["Evaluacion"].dropna().unique()))
        base_order = getattr(rn, "EVAL_ORDER", evals)
        orden = {ev: i for i, ev in enumerate(base_order)}
        self.evals_globales = sorted(evals, key=lambda x: orden.get(x, 999))

        # estado inicial
        self._on_modo_change()
        self._update_resumen()

        self.lbl_status.configure(text="Datos listos.", text_color=COL_OK)
    # ==================================================
    # RECARGAR QUERY (CLAVE PARA CAMBIO CPE / PREGRADO)
    # ==================================================
    def recargar_query(self):
        try:
            self.df = rn.cargar_df()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo recargar el Query:\n{e}")
            self.lbl_status.configure(text="Error recargando query.", text_color=COL_ERR)
            return

        # actualizar cursos
        self.cursos = sorted(self.df["Curso"].dropna().unique().tolist())
        self.combo_curso.configure(values=self.cursos)

        # evaluaciones globales
        evals = sorted(set(self.df["Evaluacion"].dropna().unique()))
        base_order = getattr(rn, "EVAL_ORDER", evals)
        orden = {ev: i for i, ev in enumerate(base_order)}
        self.evals_globales = sorted(evals, key=lambda x: orden.get(x, 999))

        # reset UI
        self.modo_var.set("CURSO")
        self.curso_var.set(self.cursos[0] if self.cursos else "")
        self.eval_var.set("")

        self._actualizar_evals()
        self._update_resumen()

        self.lbl_status.configure(text="Query actualizado.", text_color=COL_OK)
        # ==================================================
    # CUANDO LA VISTA SE MUESTRA
    # ==================================================
    def on_show(self):
        self.recargar_query()

    # ==================================================
    # EVENTOS UI
    # ==================================================
    def _on_modo_change(self):
        if self.modo_var.get() == "CURSO":
            self.combo_curso.configure(state="normal")
            if self.cursos and not self.curso_var.get():
                self.curso_var.set(self.cursos[0])
        else:
            self.combo_curso.configure(state="disabled")
            self.curso_var.set("")

        self._actualizar_evals()
        self._update_resumen()

    def _on_curso_change(self, value=None):
        self._actualizar_evals()
        self._update_resumen()

    def _on_slider_change(self, value):
        v = int(round(float(value)))
        self.porcentaje_num.set(float(v))
        self.lbl_porc.configure(text=f"{v}%")
        self._update_resumen()

    # ==================================================
    # ACTUALIZAR EVALUACIONES (igual al antiguo)
    # ==================================================
    def _actualizar_evals(self):
        if self.df is None:
            return

        modo = self.modo_var.get()
        base_order = getattr(rn, "EVAL_ORDER", [])
        eval_names = getattr(rn, "EVAL_NAMES", {})

        if modo == "CURSO":
            curso = self.curso_var.get()
            df = self.df[self.df["Curso"] == curso] if curso else self.df.iloc[0:0]
            evals = sorted(set(df["Evaluacion"].dropna().unique()))
        else:
            evals = list(self.evals_globales)

        # aplicar orden institucional si existe
        if base_order:
            orden = {ev: i for i, ev in enumerate(base_order)}
            evals = [ev for ev in evals if ev in orden]
            evals = sorted(evals, key=lambda x: orden[x])

        items = [f"{ev} — {eval_names.get(ev, ev)}" for ev in evals]

        self.combo_eval.configure(values=items)
        if items:
            if self.eval_var.get() not in items:
                self.eval_var.set(items[0])
        else:
            self.eval_var.set("")

    # ==================================================
    # PARÁMETROS (FALTABA Y POR ESO CRASHEABA)
    # ==================================================
    def _obtener_parametros(self):
        modo = self.modo_var.get()
        curso = self.curso_var.get().strip()
        eval_display = self.eval_var.get().strip()
        P = float(self.porcentaje_num.get())

        if not eval_display:
            messagebox.showwarning("Falta información", "Debe seleccionar una evaluación.")
            return None

        # código antes del "—"
        eval_code = eval_display.split("—")[0].strip()

        if modo == "CURSO" and not curso:
            messagebox.showwarning("Falta información", "Debe seleccionar un curso.")
            return None

        if not (0 <= P <= 100):
            messagebox.showwarning("Porcentaje inválido", "El porcentaje debe estar entre 0 y 100.")
            return None

        return modo, curso, eval_code, P

    # ==================================================
    # CÁLCULO GLOBAL (igual al antiguo)
    # ==================================================
    def _calcular_registro_global(self, eval_code, P):
        df_eval = self.df[self.df["Evaluacion"] == eval_code]
        if df_eval.empty:
            return None, 0, 0

        rows = []
        for (curso, sec), g in df_eval.groupby(["Curso", "Seccion"]):
            total = len(g)
            con = (g["Nota"] > 0).sum()
            porc = (con / total * 100) if total else 0
            docente = g["Docente"].iloc[0] if "Docente" in g.columns else ""
            cargo = "Sí" if porc >= P else "No"
            rows.append([curso, sec, total, docente, cargo, round(porc, 1)])

        df_res = pd.DataFrame(
            rows,
            columns=["Curso", "Sección", "Total", "Docente", "Cargó Notas", "% con nota"]
        )

        total_si = (df_res["Cargó Notas"] == "Sí").sum()
        total_no = (df_res["Cargó Notas"] == "No").sum()

        return df_res, int(total_si), int(total_no)

    # ==================================================
    # RESUMEN
    # ==================================================
    def _update_resumen(self, df_res=None, total_si=None, total_no=None):
        self.txt_resumen.configure(state="normal")
        self.txt_resumen.delete("1.0", "end")

        modo_txt = "Por curso" if self.modo_var.get() == "CURSO" else "Global"
        curso_txt = self.curso_var.get() or "(no aplica)"
        eval_txt = self.eval_var.get() or "(ninguna)"
        P = int(round(float(self.porcentaje_num.get())))

        self.txt_resumen.insert("end", "⚙ Parámetros del reporte\n")
        self.txt_resumen.insert("end", "─────────────────────────\n")
        self.txt_resumen.insert("end", f"• Tipo: {modo_txt}\n")
        self.txt_resumen.insert("end", f"• Curso: {curso_txt}\n")
        self.txt_resumen.insert("end", f"• Evaluación: {eval_txt}\n")
        self.txt_resumen.insert("end", f"• % mínimo: {P}%\n")

        if df_res is not None and not df_res.empty:
            self.txt_resumen.insert("end", "\n\n📊 Resultados\n")
            self.txt_resumen.insert("end", "─────────────────────────\n")
            self.txt_resumen.insert("end", f"• Total secciones: {len(df_res)}\n")
            self.txt_resumen.insert("end", f"• Cargaron notas: {total_si}\n")
            self.txt_resumen.insert("end", f"• No cargaron notas: {total_no}\n\n")

            preview = df_res.head(6).to_string(index=False)
            self.txt_resumen.insert("end", preview)

        self.txt_resumen.configure(state="disabled")

    # ==================================================
    # EXPORTACIÓN (pide carpeta al momento de exportar)
    # ==================================================
    def _exportar_pdf(self):
        params = self._obtener_parametros()
        if params is None:
            return

        modo, curso, ev, P = params
        eval_nombre = getattr(rn, "EVAL_NAMES", {}).get(ev, ev)

        # calcular
        if modo == "CURSO":
            df_res, si, no = rn.calcular_registro(self.df, curso, ev, P)
        else:
            df_res, si, no = self._calcular_registro_global(ev, P)

        if df_res is None or df_res.empty:
            messagebox.showinfo("Sin datos", "No hay datos para estos filtros.")
            return

        carpeta = filedialog.askdirectory(title="Seleccione carpeta de destino (PDF)")
        if not carpeta:
            return

        try:
            ruta_pdf = rn.exportar_pdf(
                df_res,
                curso if modo == "CURSO" else "TODOS",
                eval_nombre, P, si, no,
                carpeta
            )

            # 👉 ABRIR AUTOMÁTICAMENTE
            if ruta_pdf and os.path.exists(ruta_pdf):
                os.startfile(ruta_pdf)

            self.lbl_status.configure(text="PDF generado correctamente.", text_color=COL_OK)
            self._update_resumen(df_res, si, no)
            messagebox.showinfo("PDF", "PDF generado correctamente.")
        except Exception as e:
            self.lbl_status.configure(text="Error generando PDF.", text_color=COL_ERR)
            messagebox.showerror("Error", f"No se pudo generar PDF:\n{e}")

    def _exportar_excel(self):
        params = self._obtener_parametros()
        if params is None:
            return

        modo, curso, ev, P = params
        eval_nombre = getattr(rn, "EVAL_NAMES", {}).get(ev, ev)

        # calcular
        if modo == "CURSO":
            df_res, si, no = rn.calcular_registro(self.df, curso, ev, P)
        else:
            df_res, si, no = self._calcular_registro_global(ev, P)

        if df_res is None or df_res.empty:
            messagebox.showinfo("Sin datos", "No hay datos para estos filtros.")
            return

        carpeta = filedialog.askdirectory(title="Seleccione carpeta de destino (Excel)")
        if not carpeta:
            return

        try:
            rn.exportar_excel(
                df_res,
                curso if modo == "CURSO" else "TODOS",
                eval_nombre, P, si, no,
                carpeta
            )
            self.lbl_status.configure(text="Excel generado correctamente.", text_color=COL_OK)
            self._update_resumen(df_res, si, no)
            messagebox.showinfo("Excel", "Excel generado correctamente.")
        except Exception as e:
            self.lbl_status.configure(text="Error generando Excel.", text_color=COL_ERR)
            messagebox.showerror("Error", f"No se pudo generar Excel:\n{e}")
