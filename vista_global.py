# vistas/vista_global.py
# ============================================================
# FRAME — ANÁLISIS GLOBAL ACADÉMICO (UCSUR)
# SELECCIÓN DE CURSOS EN DOS COLUMNAS (CLICK)
# ============================================================

import os
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image

import Analisis_Global_UCSUR_final as ag

COL_AZUL = "#003B70"
COL_BG = "#F5F7FB"
COL_CARD = "white"
COL_MUTED = "#6b7280"


class FrameGlobal(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color=COL_BG)
        self.controller = controller

        # Estado
        self.df = None
        self.cursos = []
        self.cursos_seleccionados = []

        self.modo_final = tk.BooleanVar(value=True)
        self.tipo_reporte = tk.StringVar(value="PDF")
        self.eval_vars = {ev: tk.BooleanVar(value=True) for ev in ag.EVALS}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._logo_img = None
        self._build_header()
        self._build_body()

        self._cargar_datos()

    # ==================================================
    # HEADER
    # ==================================================
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=COL_CARD, height=110)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        if os.path.exists("logo_ucsur.png"):
            img = Image.open("logo_ucsur.png")
            img.thumbnail((220, 90))
            self._logo_img = ctk.CTkImage(light_image=img, size=img.size)
            ctk.CTkLabel(header, image=self._logo_img, text="").grid(
                row=0, column=0, rowspan=2, padx=16, pady=10
            )

        ctk.CTkLabel(
            header,
            text="ANÁLISIS GLOBAL DE NOTAS",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=COL_AZUL
        ).grid(row=0, column=1, sticky="w", padx=10, pady=(20, 4))

        ctk.CTkLabel(
            header,
            text="Panorama por curso y diagnóstico por evaluaciones",
            font=ctk.CTkFont(size=12),
            text_color=COL_MUTED
        ).grid(row=1, column=1, sticky="w", padx=10)

    # ==================================================
    # BODY
    # ==================================================
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color=COL_BG)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=14)
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self._build_cursos_card(body)
        self._build_opciones_card(body)

    # ==================================================
    # CARD — CURSOS (DOS COLUMNAS)
    # ==================================================
    def _build_cursos_card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=COL_CARD, corner_radius=14)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        card.grid_columnconfigure((0, 1), weight=1)
        card.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            card,
            text="1) Selección de cursos (click para mover)",
            font=ctk.CTkFont(weight="bold"),
            text_color=COL_AZUL
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(14, 6))

        # Buscador (solo disponibles)
        self.txt_buscar = ctk.CTkEntry(card, placeholder_text="Buscar curso...")
        self.txt_buscar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 2))
        self.txt_buscar.bind("<KeyRelease>", lambda e: self._filtrar_disponibles())

        # Labels
        ctk.CTkLabel(card, text="Disponibles", text_color=COL_MUTED)\
            .grid(row=2, column=0, sticky="w", padx=14)
        ctk.CTkLabel(card, text="Seleccionados", text_color=COL_MUTED)\
            .grid(row=2, column=1, sticky="w", padx=14)

        # Listboxes
        self.lst_disp = tk.Listbox(card, height=14, exportselection=False)
        self.lst_sel = tk.Listbox(card, height=14, exportselection=False)

        self.lst_disp.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 6))
        self.lst_sel.grid(row=3, column=1, sticky="nsew", padx=14, pady=(0, 6))


        self.lst_disp.bind("<ButtonRelease-1>", self._agregar_curso)
        self.lst_sel.bind("<ButtonRelease-1>", self._quitar_curso)

        # Botones
        fr_btn = ctk.CTkFrame(card, fg_color=COL_CARD)
        fr_btn.grid(row=4, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 14))
        fr_btn.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(fr_btn, text="Seleccionar TODOS",
                      command=self._seleccionar_todos,
                      fg_color=COL_AZUL).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(fr_btn, text="Limpiar selección",
                      command=self._limpiar_seleccion,
                      fg_color="#334155").grid(row=0, column=1, sticky="ew", padx=(6, 0))

    # ==================================================
    # CARD — OPCIONES
    # ==================================================
    def _build_opciones_card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=COL_CARD, corner_radius=14)
        card.grid(row=0, column=1, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text="2) Configuración",
                     font=ctk.CTkFont(weight="bold"),
                     text_color=COL_AZUL)\
            .grid(row=0, column=0, sticky="w", padx=14, pady=(14, 6))

        ctk.CTkRadioButton(card, text="Condición final",
                           variable=self.modo_final, value=True)\
            .grid(row=1, column=0, sticky="w", padx=14)

        ctk.CTkRadioButton(card, text="Evaluaciones (diagnóstico)",
                           variable=self.modo_final, value=False)\
            .grid(row=2, column=0, sticky="w", padx=14)

        self.panel_evals = ctk.CTkFrame(card, fg_color="#f8fafc", corner_radius=10)
        self.panel_evals.grid(row=3, column=0, sticky="ew", padx=14, pady=8)

        for i, ev in enumerate(ag.EVALS):
            ctk.CTkCheckBox(self.panel_evals, text=ev,
                            variable=self.eval_vars[ev])\
                .grid(row=i // 3, column=i % 3, padx=8, pady=4, sticky="w")

        ctk.CTkRadioButton(card, text="PDF",
                           variable=self.tipo_reporte, value="PDF")\
            .grid(row=4, column=0, sticky="w", padx=14)
        ctk.CTkRadioButton(card, text="Excel",
                           variable=self.tipo_reporte, value="EXCEL")\
            .grid(row=5, column=0, sticky="w", padx=14)

        ctk.CTkButton(
            card, text="GENERAR REPORTE",
            fg_color=COL_AZUL, height=42,
            command=self._generar_reporte
        ).grid(row=6, column=0, sticky="ew", padx=14, pady=(10, 14))

    # ==================================================
    # DATOS Y LÓGICA
    # ==================================================
    def _cargar_datos(self):
        self.df = ag.cargar_df(ag.PARQUET_IN)
        self.cursos = sorted(self.df["Curso"].unique().tolist())
        self.cursos_seleccionados = []
        self._refrescar_listas()

    def _refrescar_listas(self):
        self.lst_disp.delete(0, "end")
        self.lst_sel.delete(0, "end")

        for c in self.cursos:
            if c not in self.cursos_seleccionados:
                self.lst_disp.insert("end", c)

        for c in self.cursos_seleccionados:
            self.lst_sel.insert("end", c)

    def _filtrar_disponibles(self):
        q = self.txt_buscar.get().lower().strip()
        self.lst_disp.delete(0, "end")
        for c in self.cursos:
            if c in self.cursos_seleccionados:
                continue
            if q in c.lower():
                self.lst_disp.insert("end", c)

    def _agregar_curso(self, _):
        sel = self.lst_disp.curselection()
        if not sel:
            return
        curso = self.lst_disp.get(sel[0])
        self.cursos_seleccionados.append(curso)
        self._refrescar_listas()

    def _quitar_curso(self, _):
        sel = self.lst_sel.curselection()
        if not sel:
            return
        curso = self.lst_sel.get(sel[0])
        self.cursos_seleccionados.remove(curso)
        self._refrescar_listas()

    def _seleccionar_todos(self):
        self.cursos_seleccionados = self.cursos[:]
        self._refrescar_listas()

    def _limpiar_seleccion(self):
        self.cursos_seleccionados = []
        self._refrescar_listas()

    # ==================================================
    # GENERAR REPORTE
    # ==================================================
    def _generar_reporte(self):

        # -------------------------
        # Cursos seleccionados
        # -------------------------
        cursos_sel = None if not self.cursos_seleccionados else self.cursos_seleccionados
        df_sel = self.df if cursos_sel is None else self.df[self.df["Curso"].isin(cursos_sel)]

        # -------------------------
        # Evaluaciones
        # -------------------------
        evals_sel = None
        if not self.modo_final.get():
            evals_sel = [ev for ev, v in self.eval_vars.items() if v.get()]
            if not evals_sel:
                messagebox.showwarning(
                    "Selección incompleta",
                    "Debe seleccionar al menos una evaluación."
                )
                return

            if cursos_sel is None:
                messagebox.showwarning(
                    "Selección incompleta",
                    "Debe seleccionar al menos un curso para el análisis por evaluaciones."
                )
                return

        # -------------------------
        # Carpeta destino
        # -------------------------
        carpeta = filedialog.askdirectory()
        if not carpeta:
            return

        # -------------------------
        # Exportación (API COMPATIBLE)
        # -------------------------
        try:
            if self.tipo_reporte.get() == "PDF":
                if self.modo_final.get():
                    ruta = ag.exportar_pdf_final(
                        df=df_sel,
                        cursos_sel=cursos_sel,
                        carpeta_destino=carpeta
                    )
                else:
                    ruta = ag.exportar_pdf_evaluaciones(
                        df=df_sel,
                        evals_sel=evals_sel,
                        cursos_sel=cursos_sel,
                        carpeta_destino=carpeta
                    )
            else:
                ruta = ag.exportar_excel_global(
                    df=df_sel,
                    evals_sel=evals_sel,
                    modo_final=self.modo_final.get(),
                    carpeta_destino=carpeta
                )

            os.startfile(ruta)

        except Exception as e:
            messagebox.showerror("Error al generar reporte", str(e))
    def on_show(self):
        self._cargar_datos()
