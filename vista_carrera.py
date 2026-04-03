# vistas/vista_carrera.py
# ============================================================
# FRAME — ANÁLISIS POR CARRERA (UCSUR)
# Migrado desde Analisis_carrera_GUI.py
# ============================================================

import os
import customtkinter as ctk
from tkinter import messagebox
from PIL import Image, ImageTk

# 🔹 Script de análisis ORIGINAL (NO SE TOCA)
import Analisis_carrera as ac

COL_UCSUR_AZUL = "#003B70"


class FrameCarrera(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#F5F5F5")

        self.controller = controller

        # Estado
        self.df = None
        self.carreras = []
        self.carrera_actual = None
        self.cursos_actuales = []
        self.curso_actual = None
        self.logo_img = None

        # Layout base
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_main()
        self._cargar_datos()

    # ==================================================
    # HEADER
    # ==================================================
    def _build_header(self):
        header = ctk.CTkFrame(self, height=180, fg_color="white", corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.columnconfigure(0, weight=1)

        inner = ctk.CTkFrame(header, fg_color="white", corner_radius=0)
        inner.grid(row=0, column=0, sticky="nsew", pady=(8, 0))
        inner.columnconfigure(0, weight=1)

        logo_path = getattr(ac, "LOGO_PATH", "logo_ucsur.png")
        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path)
                img.thumbnail((320, 150), Image.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                ctk.CTkLabel(inner, image=self.logo_img, text="").grid(row=0, column=0)
            except Exception as e:
                print("Error logo:", e)

        ctk.CTkLabel(
            inner,
            text="Reporte de Notas por Carrera",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=1, column=0, pady=(4, 2))

        ctk.CTkLabel(
            inner,
            text="Departamento de Cursos Básicos - Equipo de Matemática",
            font=ctk.CTkFont(size=14),
            text_color="#555"
        ).grid(row=2, column=0)

        ctk.CTkFrame(header, fg_color=COL_UCSUR_AZUL, height=3).grid(
            row=1, column=0, sticky="ew", pady=(6, 0)
        )

    # ==================================================
    # CUERPO PRINCIPAL
    # ==================================================
    def _build_main(self):
        body = ctk.CTkFrame(self, fg_color="#F5F5F5")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        body.grid_columnconfigure(0, weight=1)

        card = ctk.CTkFrame(body, fg_color="white", corner_radius=12)
        card.grid(row=0, column=0, sticky="n", padx=10, pady=10)
        card.grid_columnconfigure(1, weight=1)

        # ----- Carrera -----
        ctk.CTkLabel(
            card,
            text="Carrera:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(20, 6))

        self.combo_carrera = ctk.CTkComboBox(
            card,
            values=[],
            state="readonly",
            width=420,
            command=self._on_carrera_change
        )
        self.combo_carrera.grid(row=0, column=1, padx=20, pady=(20, 6), sticky="ew")

        # ----- Curso -----
        ctk.CTkLabel(
            card,
            text="Curso:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_UCSUR_AZUL
        ).grid(row=1, column=0, sticky="w", padx=20, pady=6)

        self.combo_curso = ctk.CTkComboBox(
            card,
            values=[],
            state="disabled",
            width=420,
            command=self._on_curso_change
        )
        self.combo_curso.grid(row=1, column=1, padx=20, pady=6, sticky="ew")

        # ----- Botones -----
        btns = ctk.CTkFrame(card, fg_color="white")
        btns.grid(row=2, column=0, columnspan=2, pady=(20, 25))
        btns.grid_columnconfigure((0, 1), weight=1)

        self.btn_pdf = ctk.CTkButton(
            btns,
            text="Exportar PDF",
            fg_color=COL_UCSUR_AZUL,
            hover_color="#004B8D",
            state="disabled",
            command=self._exportar_pdf
        )
        self.btn_pdf.grid(row=0, column=0, padx=10, ipadx=20)

        self.btn_excel = ctk.CTkButton(
            btns,
            text="Exportar Excel",
            fg_color="#1C7C54",
            hover_color="#239966",
            state="disabled",
            command=self._exportar_excel
        )
        self.btn_excel.grid(row=0, column=1, padx=10, ipadx=20)

    # ==================================================
    # LÓGICA
    # ==================================================
    def _cargar_datos(self):
        try:
            self.df = ac.cargar_df(ac.PARQUET_IN)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el parquet:\n{e}")
            return

        carreras = sorted(
            [c for c in self.df["Carrera"].astype(str).unique() if str(c).strip()]
        )
        self.carreras = carreras

        if carreras:
            self.combo_carrera.configure(values=carreras)
        else:
            messagebox.showwarning("Aviso", "No hay carreras disponibles.")
        # ==================================================
    # RECARGAR QUERY (CPE / PREGRADO)
    # ==================================================
    def recargar_query(self):
        try:
            self.df = ac.cargar_df(ac.PARQUET_IN)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo cargar el parquet:\n{e}"
            )
            return

        # Carreras
        self.carreras = sorted(
            [c for c in self.df["Carrera"].astype(str).unique() if str(c).strip()]
        )

        self.combo_carrera.configure(values=self.carreras)
        self.combo_carrera.set("")

        # Reset estado
        self.carrera_actual = None
        self.curso_actual = None
        self.cursos_actuales = []

        self.combo_curso.configure(values=[], state="disabled")

        self.btn_pdf.configure(state="disabled")
        self.btn_excel.configure(state="disabled")


    def _on_carrera_change(self, value):
        self.carrera_actual = value
        self.curso_actual = None

        df_car = self.df[self.df["Carrera"].str.upper() == value.upper()]
        cursos = sorted(
            [c for c in df_car["Curso"].astype(str).unique() if str(c).strip()]
        )
        self.cursos_actuales = cursos

        if cursos:
            self.combo_curso.configure(values=cursos, state="readonly")
            self.combo_curso.set("")
        else:
            self.combo_curso.configure(values=[], state="disabled")

        self.btn_pdf.configure(state="disabled")
        self.btn_excel.configure(state="disabled")

    def _on_curso_change(self, value):
        self.curso_actual = value
        if self.carrera_actual and self.curso_actual:
            self.btn_pdf.configure(state="normal")
            self.btn_excel.configure(state="normal")

    def _get_df_filtrado(self):
        df = self.df[
            (self.df["Carrera"].str.upper() == self.carrera_actual.upper()) &
            (self.df["Curso"] == self.curso_actual)
        ].copy()
        return df

    def _exportar_pdf(self):
        df = self._get_df_filtrado()
        if df.empty:
            messagebox.showwarning("Aviso", "No hay datos para exportar.")
            return

        carpeta = ac.elegir_carpeta()
        if carpeta:
            ac.exportar_pdf(df, self.carrera_actual, self.curso_actual, carpeta)

    def _exportar_excel(self):
        df = self._get_df_filtrado()
        if df.empty:
            messagebox.showwarning("Aviso", "No hay datos para exportar.")
            return

        carpeta = ac.elegir_carpeta()
        if carpeta:
            ac.exportar_excel(df, self.carrera_actual, self.curso_actual, carpeta)
    # ==================================================
    # CUANDO LA VISTA SE MUESTRA
    # ==================================================
    def on_show(self):
        self.recargar_query()
