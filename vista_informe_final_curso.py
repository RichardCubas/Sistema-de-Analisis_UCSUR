# vista_informe_final_curso.py
# -*- coding: utf-8 -*-

import os
import customtkinter as ctk
from tkinter import filedialog, messagebox

import pandas as pd

from informe_final_curso_reportes import (
    cargar_df_informe,
    generar_excel_informe_final_curso,
    generar_pdf_informe_final_curso,
)

COL_UCSUR_AZUL = "#003B70"
PARQUET_IN = "notas_filtradas_ucsur.parquet"


class FrameInformeFinalCurso(ctk.CTkFrame):

    def __init__(self, parent, app):
        super().__init__(parent, fg_color="#F5F5F5")
        self.app = app
        self.df = None

        self.grid_rowconfigure(99, weight=1)
        self.grid_columnconfigure(0, weight=1)

        titulo = ctk.CTkLabel(
            self,
            text="INFORME FINAL DE CURSO",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=COL_UCSUR_AZUL
        )
        titulo.grid(row=0, column=0, padx=18, pady=(18, 8), sticky="w")

        self.lbl_estado = ctk.CTkLabel(self, text="Cargando datos...", text_color="#444")
        self.lbl_estado.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="w")

        box = ctk.CTkFrame(self, fg_color="white", corner_radius=14)
        box.grid(row=2, column=0, padx=18, pady=12, sticky="ew")
        box.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(box, text="Curso:", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, padx=(16, 10), pady=14, sticky="w"
        )

        self.var_curso = ctk.StringVar(value="")
        self.combo_curso = ctk.CTkOptionMenu(
            box,
            variable=self.var_curso,
            values=["(cargando...)"],
            fg_color="white",
            text_color="#222",
            button_color=COL_UCSUR_AZUL,
            button_hover_color="#004B8D",
            dropdown_fg_color="white",
            dropdown_text_color="#222",
            width=420
        )
        self.combo_curso.grid(row=0, column=1, padx=(0, 16), pady=14, sticky="w")

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.grid(row=3, column=0, padx=18, pady=(8, 0), sticky="w")

        self.btn_excel = ctk.CTkButton(
            btns,
            text="📗 GENERAR REPORTE EN EXCEL",
            fg_color=COL_UCSUR_AZUL,
            hover_color="#004B8D",
            height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._generar_excel
        )
        self.btn_excel.grid(row=0, column=0, padx=(0, 12), pady=8)

        self.btn_pdf = ctk.CTkButton(
            btns,
            text="📄 GENERAR REPORTE EN PDF",
            fg_color=COL_UCSUR_AZUL,
            hover_color="#004B8D",
            height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._generar_pdf
        )
        self.btn_pdf.grid(row=0, column=1, padx=(0, 12), pady=8)

        self.txt_log = ctk.CTkTextbox(self, height=260, corner_radius=14)
        self.txt_log.grid(row=4, column=0, padx=18, pady=14, sticky="nsew")

        self._log("✅ Vista lista. Se cargará el parquet al mostrar la vista.")

    def on_show(self):
        self._cargar()

    def _log(self, msg):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")

    def _cargar(self):
        try:
            self.lbl_estado.configure(text="Cargando parquet...")
            self.df = cargar_df_informe(PARQUET_IN)
            cursos = sorted([c for c in self.df["Curso"].astype(str).str.upper().unique().tolist() if c.strip()])
            if not cursos:
                self.combo_curso.configure(values=["(sin cursos)"])
                self.var_curso.set("")
                self.lbl_estado.configure(text="⚠️ No se encontraron cursos.")
                return

            self.combo_curso.configure(values=cursos)
            self.var_curso.set(cursos[0])
            self.lbl_estado.configure(text=f"✅ Parquet cargado. Cursos: {len(cursos)}")
            self._log(f"✅ Parquet cargado: {len(self.df):,} filas.")
        except Exception as e:
            self.lbl_estado.configure(text="❌ Error al cargar datos.")
            self._log(f"❌ Error: {e}")

    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccione carpeta destino")
        return carpeta if carpeta else None

    def _generar_excel(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados.")
            return
        curso = (self.var_curso.get() or "").strip().upper()
        if not curso:
            messagebox.showwarning("Curso", "Seleccione un curso.")
            return

        carpeta = self._elegir_carpeta()
        if not carpeta:
            return

        try:
            self._log(f"➡ Generando EXCEL para curso: {curso}")
            out = generar_excel_informe_final_curso(self.df, curso, carpeta)
            self._log(f"✅ Excel generado: {out}")
            messagebox.showinfo("Listo", f"Excel generado:\n{out}")
            if os.name == "nt":
                os.startfile(out)
        except Exception as e:
            self._log(f"❌ Error Excel: {e}")
            messagebox.showerror("Error", f"No se pudo generar Excel:\n{e}")

    def _generar_pdf(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados.")
            return
        curso = (self.var_curso.get() or "").strip().upper()
        if not curso:
            messagebox.showwarning("Curso", "Seleccione un curso.")
            return

        carpeta = self._elegir_carpeta()
        if not carpeta:
            return

        try:
            self._log(f"➡ Generando PDF para curso: {curso}")
            out = generar_pdf_informe_final_curso(self.df, curso, carpeta)
            self._log(f"✅ PDF generado: {out}")
            messagebox.showinfo("Listo", f"PDF generado:\n{out}")
            if os.name == "nt":
                os.startfile(out)
        except Exception as e:
            self._log(f"❌ Error PDF: {e}")
            messagebox.showerror("Error", f"No se pudo generar PDF:\n{e}")
