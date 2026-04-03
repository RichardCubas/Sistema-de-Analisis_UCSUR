# vista_query.py
# ============================================================
# FRAME — EXTRACCIÓN DE QUERY (UCSUR)
# ============================================================

import os
import shutil
import customtkinter as ctk
from tkinter import filedialog, messagebox

# BACKEND (NO SE TOCA)
import Extraer_Query as eq

COL_AZUL = "#003B70"


class FrameQuery(ctk.CTkFrame):

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#F5F7FB")

        self.controller = controller

        # ==========================
        # ESTADO
        # ==========================
        self.modo = None  # "pregrado" | "cpe"
        self.modo_var = ctk.StringVar(value="")

        self.parquet_pre = "notas_filtradas_ucsur_pregrado.parquet"
        self.parquet_cpe = "notas_filtradas_ucsur_cpe.parquet"

        # ==========================
        # LAYOUT
        # ==========================
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_body()
        self._actualizar_label_ultimo_query()

    # ==================================================
    # HEADER
    # ==================================================
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=COL_AZUL, height=90)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            header,
            text="Extractor de Query UCSUR",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(10, 0))

        self.lbl_modo = ctk.CTkLabel(
            header,
            text="Modo actual: (sin seleccionar)",
            font=ctk.CTkFont(size=14),
            text_color="white"
        )
        self.lbl_modo.grid(row=0, column=1, sticky="e", padx=16, pady=(10, 0))

        self.lbl_ultimo_query = ctk.CTkLabel(
            header,
            text="Último Query: (ninguno cargado aún)",
            font=ctk.CTkFont(size=13),
            text_color="#E6EEF9"
        )
        self.lbl_ultimo_query.grid(row=1, column=0, columnspan=2, sticky="w", padx=16)

    # ==================================================
    # BODY
    # ==================================================
    def _build_body(self):
        body = ctk.CTkFrame(self, fg_color="#F5F7FB")
        body.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # PANEL IZQUIERDO
        left = ctk.CTkFrame(body, fg_color="white", corner_radius=12)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10), pady=10)

        ctk.CTkLabel(
            left, text="1. Tipo de programa",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", padx=12, pady=(12, 4))

        ctk.CTkRadioButton(
            left, text="PREGRADO",
            variable=self.modo_var, value="pregrado",
            command=self._on_modo_change
        ).pack(anchor="w", padx=24, pady=4)

        ctk.CTkRadioButton(
            left, text="CPE",
            variable=self.modo_var, value="cpe",
            command=self._on_modo_change
        ).pack(anchor="w", padx=24, pady=4)

        ctk.CTkLabel(
            left, text="2. Origen del Query",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", padx=12, pady=(14, 4))

        ctk.CTkButton(
            left, text="Cargar archivo CSV / Excel", width=180,
            command=self._cargar_query
        ).pack(padx=12, pady=4)

        ctk.CTkButton(
            left, text="Usar Query existente", width=180,
            command=self._usar_query_existente
        ).pack(padx=12, pady=4)

        ctk.CTkLabel(
            left, text="3. Descargar datos (EXCEL)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=COL_AZUL
        ).pack(anchor="w", padx=12, pady=(14, 4))

        ctk.CTkButton(
            left, text="Generar Excel", width=180,
            command=lambda: eq.generar_excel_desde_parquet(
                self, self.txt_logs, self.progress
            )
        ).pack(padx=12, pady=4)

        # PANEL DERECHO
        right = ctk.CTkFrame(body, fg_color="white", corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", pady=10)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            right, text="Registro de eventos",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COL_AZUL
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))

        self.txt_logs = ctk.CTkTextbox(right)
        self.txt_logs.grid(row=1, column=0, sticky="nsew", padx=12, pady=4)
        self.txt_logs.configure(state="disabled")

        self.progress = ctk.CTkProgressBar(right)
        self.progress.grid(row=2, column=0, sticky="ew", padx=12, pady=(6, 4))
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(
            right, text="Listo.", font=ctk.CTkFont(size=12)
        )
        self.status_label.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 10))

    # ==================================================
    # LÓGICA
    # ==================================================
    def _on_modo_change(self):
        self.modo = self.modo_var.get()
        self.lbl_modo.configure(text=f"Modo actual: {self.modo.upper()}")
        eq.append_log(self.txt_logs, f"▶ Programa seleccionado: {self.modo.upper()}")

    def _cargar_query(self):
        if not self.modo:
            messagebox.showwarning(
                "Modo no seleccionado",
                "Seleccione PREGRADO o CPE."
            )
            return

        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo CSV o Excel",
            filetypes=[
                ("Archivos de datos", "*.csv *.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("Excel", "*.xlsx *.xls")
            ]
        )
        if not archivo:
            return

        parquet = self.parquet_pre if self.modo == "pregrado" else self.parquet_cpe

        self.progress.set(0)
        self.status_label.configure(text="Procesando query...")
        self.update_idletasks()

        # 🔑 USAR FUNCIÓN REAL DEL BACKEND
        eq.procesar_query_archivo(
            archivo,
            parquet,
            self.modo,
            self,
            self.txt_logs,
            self.progress
        )

        self.status_label.configure(text="Proceso completado.")
        self._actualizar_label_ultimo_query()

    def _usar_query_existente(self):
        if not self.modo:
            messagebox.showwarning("Modo no seleccionado",
                                   "Seleccione PREGRADO o CPE.")
            return

        parquet = self.parquet_pre if self.modo == "pregrado" else self.parquet_cpe
        if not os.path.exists(parquet):
            messagebox.showwarning("Query inexistente",
                                   f"No existe {parquet}")
            return

        shutil.copyfile(parquet, "notas_filtradas_ucsur.parquet")
        eq.guardar_info_query(self.modo)
        eq.append_log(self.txt_logs, f"📦 Usando query existente: {parquet}")
        self._actualizar_label_ultimo_query()

    def _actualizar_label_ultimo_query(self):
        modo, fecha = eq.leer_info_query()
        if modo and fecha:
            self.lbl_ultimo_query.configure(
                text=f"Último Query: QUERY-{modo.upper()}, cargado el {fecha}"
            )
        else:
            self.lbl_ultimo_query.configure(
                text="Último Query: (ninguno cargado aún)"
            )

    # ==================================================
    # CUANDO SE MUESTRA LA VISTA
    # ==================================================
    def on_show(self):
        self._actualizar_label_ultimo_query()
        self.status_label.configure(text="Listo.")
