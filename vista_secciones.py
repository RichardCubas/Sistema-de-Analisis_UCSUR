# vistas/vista_secciones.py
# ============================================================
# FRAME — ANÁLISIS POR CURSO, SECCIÓN Y DOCENTE (UCSUR)
# Versión optimizada, funcional y amigable
# ============================================================

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

import Analisis_Secciones_Final as asf

COL_AZUL = "#003B70"
COL_AZUL_CLARO = "#245E9C"


class FrameSecciones(ctk.CTkFrame):

    def __init__(self, parent, controller=None):
        super().__init__(parent, fg_color="white")
        self.controller = controller

        # ======================
        # ESTADO
        # ======================
        self.df_full = asf.cargar_df(asf.PARQUET_IN)
        self.cursos = sorted(
            [c for c in self.df_full["Curso"].astype(str).unique() if c.strip()]
        )

        self.step = 0
        self.selected_curso = None
        self.df_curso = None

        self.modalidad_var = tk.StringVar(value="secciones")
        self.selected_docente = None
        self.docentes = []

        self.evals_vars = {}
        self.selected_evals = []

        # ======================
        # LAYOUT (CAMBIO CLAVE AQUÍ)
        # ======================
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)  # ⬅️ el contenido ahora crece aquí

        self._build_header()
        self._build_steps()
        self._build_nav()        # ⬅️ navegación más arriba
        self._build_content()    # ⬅️ contenido más abajo

        self._update_view()
        # ==================================================
    # RECARGAR QUERY (CPE / PREGRADO)
    # ==================================================
    def recargar_query(self):
        if not os.path.exists(asf.PARQUET_IN):
            messagebox.showerror(
                "Error",
                "No se encontró el archivo de datos.\nPrimero debe procesar un Query."
            )
            return

        # Recargar dataframe
        self.df_full = asf.cargar_df(asf.PARQUET_IN)

        # Actualizar cursos
        self.cursos = sorted(
            [c for c in self.df_full["Curso"].astype(str).unique() if c.strip()]
        )

        # Reset estado
        self.step = 0
        self.selected_curso = None
        self.df_curso = None
        self.selected_docente = None
        self.docentes = []
        self.selected_evals = []
        self.evals_vars.clear()

        # Actualizar combo de cursos
        if hasattr(self, "combo_curso"):
            self.combo_curso.configure(values=self.cursos)
            if self.cursos:
                self.combo_curso.set(self.cursos[0])

        # Volver al primer paso
        self._update_view()

    # ==================================================
    # HEADER
    # ==================================================
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="white", height=140)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)

        ctk.CTkLabel(
            header,
            text="Reporte por Curso, Sección y Docente",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=COL_AZUL
        ).pack(pady=(25, 4))

        ctk.CTkLabel(
            header,
            text="Análisis académico institucional",
            font=ctk.CTkFont(size=14),
            text_color="#555"
        ).pack()

        ctk.CTkFrame(header, fg_color=COL_AZUL, height=3).pack(fill="x", pady=10)

    # ==================================================
    # PASOS
    # ==================================================
    def _build_steps(self):
        self.steps_frame = ctk.CTkFrame(self, fg_color="white")
        self.steps_frame.grid(row=1, column=0, sticky="ew")

        self.step_labels = []
        texts = ["1. Curso", "2. Modalidad", "3. Evaluaciones", "4. Exportar"]

        for t in texts:
            lbl = ctk.CTkLabel(
                self.steps_frame,
                text=t,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color="#777"
            )
            lbl.pack(side="left", expand=True, pady=8)
            self.step_labels.append(lbl)

    # ==================================================
    # NAVEGACIÓN (⬅️ AHORA MÁS ARRIBA)
    # ==================================================
    def _build_nav(self):
        nav = ctk.CTkFrame(self, fg_color="white")
        nav.grid(row=2, column=0, sticky="ew", pady=(5, 10))

        # línea separadora sutil
        ctk.CTkFrame(nav, fg_color="#DDDDDD", height=1).pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            nav, text="⟵ Anterior",
            fg_color="#888",
            command=self._prev,
            width=130
        ).pack(side="left", padx=20)

        ctk.CTkButton(
            nav, text="Siguiente ⟶",
            fg_color=COL_AZUL,
            command=self._next,
            width=130
        ).pack(side="right", padx=20)

    # ==================================================
    # CONTENIDO (⬇️ MÁS ABAJO)
    # ==================================================
    def _build_content(self):
        self.content = ctk.CTkFrame(self, fg_color="white")
        self.content.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.frames = []
        for _ in range(4):
            f = ctk.CTkFrame(self.content, fg_color="white")
            f.grid(row=0, column=0, sticky="nsew")
            f.grid_remove()
            self.frames.append(f)

        self._step_curso()
        self._step_modalidad()
        self._step_evaluaciones()
        self._step_exportar()

    # ---------------- STEP 1 ----------------
    def _step_curso(self):
        f = self.frames[0]

        ctk.CTkLabel(
            f, text="Seleccione el curso",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", pady=10)

        self.combo_curso = ctk.CTkComboBox(
            f, values=self.cursos, width=400, state="readonly"
        )
        self.combo_curso.pack(anchor="w", pady=10)

    # ---------------- STEP 2 ----------------
    def _step_modalidad(self):
        f = self.frames[1]

        ctk.CTkLabel(
            f, text="Seleccione modalidad",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", pady=10)

        for txt, val in [
            ("Por secciones", "secciones"),
            ("Todos los docentes", "todos_docentes"),
            ("Docente específico", "docente")
        ]:
            ctk.CTkRadioButton(
                f, text=txt,
                variable=self.modalidad_var,
                value=val,
                command=self._on_modalidad_change,
                fg_color=COL_AZUL
            ).pack(anchor="w", pady=4)

        self.combo_docente = ctk.CTkComboBox(
            f, values=[], state="disabled", width=400
        )
        self.combo_docente.pack(anchor="w", pady=10)

    # ---------------- STEP 3 ----------------
    def _step_evaluaciones(self):
        f = self.frames[2]

        ctk.CTkLabel(
            f, text="Seleccione evaluaciones",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", pady=10)

        self.frame_checks = ctk.CTkFrame(f, fg_color="white")
        self.frame_checks.pack(anchor="w", pady=10)

    # ---------------- STEP 4 ----------------
    def _step_exportar(self):
        f = self.frames[3]

        self.lbl_resumen = ctk.CTkLabel(
            f, text="", justify="left",
            font=ctk.CTkFont(size=13)
        )
        self.lbl_resumen.pack(anchor="w", pady=15)

        ctk.CTkButton(
            f, text="Exportar PDF",
            fg_color=COL_AZUL,
            command=self._exportar_pdf
        ).pack(anchor="w", pady=6)

        ctk.CTkButton(
            f, text="Exportar Excel",
            fg_color="#1C7C54",
            command=self._exportar_excel
        ).pack(anchor="w", pady=6)

    # ==================================================
    # LÓGICA (NO TOCADA)
    # ==================================================
    def _update_view(self):
        for i, lbl in enumerate(self.step_labels):
            lbl.configure(text_color=COL_AZUL if i == self.step else "#777")

        for i, f in enumerate(self.frames):
            f.grid() if i == self.step else f.grid_remove()

        if self.step == 1:
            self._cargar_docentes()
        elif self.step == 2:
            self._cargar_evaluaciones()
        elif self.step == 3:
            self._actualizar_resumen()

    def _prev(self):
        if self.step > 0:
            self.step -= 1
            self._update_view()

    def _next(self):
        if self.step < 3:
            self.step += 1
            self._update_view()

    # ==================================================
    # DATOS / EXPORTACIÓN (SIN CAMBIOS)
    # ==================================================
    def _on_modalidad_change(self):
        self.combo_docente.configure(
            state="readonly" if self.modalidad_var.get() == "docente" else "disabled"
        )

    def _cargar_docentes(self):
        curso = self.combo_curso.get()
        if not curso:
            return

        self.selected_curso = curso
        self.df_curso = self.df_full[self.df_full["Curso"] == curso]

        self.docentes = sorted(
            self.df_curso["Docente"].dropna().unique().tolist()
        )

        self.combo_docente.configure(values=self.docentes)
        if self.docentes:
            self.combo_docente.set(self.docentes[0])

    def _cargar_evaluaciones(self):
        for w in self.frame_checks.winfo_children():
            w.destroy()
        self.evals_vars.clear()

        if self.modalidad_var.get() == "docente":
            ctk.CTkLabel(
                self.frame_checks,
                text="Se consideran todas las evaluaciones automáticamente."
            ).pack(anchor="w")
            return

        cols = 4
        r = c = 0

        for ev in asf.EVALS:
            var = tk.BooleanVar()
            chk = ctk.CTkCheckBox(
                self.frame_checks,
                text=ev,
                variable=var,
                fg_color=COL_AZUL
            )
            chk.grid(row=r, column=c, padx=12, pady=6, sticky="w")
            self.evals_vars[ev] = var

            c += 1
            if c >= cols:
                c = 0
                r += 1

    def _actualizar_resumen(self):
        modalidad = self.modalidad_var.get()
        docente = self.combo_docente.get() if modalidad == "docente" else "—"

        self.selected_evals = [e for e, v in self.evals_vars.items() if v.get()]

        evals = "Todas" if modalidad == "docente" else ", ".join(self.selected_evals)

        self.lbl_resumen.configure(
            text=(
                f"Curso: {self.selected_curso}\n"
                f"Modalidad: {modalidad}\n"
                f"Docente: {docente}\n"
                f"Evaluaciones: {evals}"
            )
        )

    def _exportar_pdf(self):
        carpeta = filedialog.askdirectory(title="Seleccione carpeta de destino")
        if not carpeta:
            return

        if self.modalidad_var.get() == "secciones":
            asf.exportar_pdf_curso_evals(
                self.df_curso, self.selected_curso, self.selected_evals, carpeta
            )
        elif self.modalidad_var.get() == "todos_docentes":
            asf.exportar_pdf_todos_docentes(
                self.df_curso, self.selected_curso, self.selected_evals, carpeta
            )
        else:
            asf.exportar_pdf_docente(
                self.df_curso, self.selected_curso,
                self.combo_docente.get(), carpeta
            )

        messagebox.showinfo("PDF", "Reporte PDF generado correctamente.")

    def _exportar_excel(self):
        carpeta = filedialog.askdirectory(title="Seleccione carpeta de destino")
        if not carpeta:
            return

        if self.modalidad_var.get() == "secciones":
            asf.exportar_excel_curso_evals(
                self.df_curso, self.selected_curso, self.selected_evals, carpeta
            )
        elif self.modalidad_var.get() == "todos_docentes":
            asf.exportar_excel_todos_docentes(
                self.df_curso, self.selected_curso, self.selected_evals, carpeta
            )
        else:
            asf.exportar_excel_docente(
                self.df_curso, self.selected_curso,
                self.combo_docente.get(), carpeta
            )

        messagebox.showinfo("Excel", "Archivo Excel generado correctamente.")
    # ==================================================
    # CUANDO LA VISTA SE MUESTRA
    # ==================================================
    def on_show(self):
        self.recargar_query()


