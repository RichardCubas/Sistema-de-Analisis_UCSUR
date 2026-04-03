# ============================================================
# VISTA: ANÁLISIS DE RIESGO ACADÉMICO (UCSUR)
# ARCHIVO: vista_analisis_riesgo.py
# ============================================================
# Fuente: riesgo_academico_reporte.parquet (generado desde vista_riesgo_academico)
# Filtros:
#   - Condición: TODOS / BICA / TRICA / CUARTA
#   - Evaluación: ED, EC1, EP, EC2, EC3, EF, SITUACIÓN FINAL (Promedio Final)
#   - Curso(s): TODOS o seleccionar uno o varios
# Métricas:
#   - Rindieron: nota > 0
#   - No rindieron: nota == 0
#   - Aprobados/Desaprobados solo entre rindieron (umbral >=12.5)
# Exporta Excel con 4 hojas:
#   RESUMEN, APROBADOS, DESAPROBADOS, NO_RINDIERON
# ============================================================

import os
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox, filedialog


class FrameAnalisisRiesgo(ctk.CTkFrame):

    PARQUET_RIESGO_OUT = "riesgo_academico_reporte.parquet"

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#F5F5F5", corner_radius=0)
        self.controller = controller

        self.df = None
        self.df_filtrado = None
        self.df_aprob = None
        self.df_desap = None
        self.df_no_rind = None

        ctk.CTkLabel(
            self,
            text="📊 Análisis — Riesgo Académico",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#003B70"
        ).pack(anchor="w", padx=18, pady=(18, 4))

        self.lbl_estado = ctk.CTkLabel(
            self,
            text="Estado: (sin datos)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#333"
        )
        self.lbl_estado.pack(anchor="w", padx=18, pady=(0, 10))

        panel = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        panel.pack(fill="x", padx=18, pady=(0, 12))

        grid = ctk.CTkFrame(panel, fg_color="transparent")
        grid.pack(fill="x", padx=14, pady=12)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)
        grid.grid_columnconfigure(2, weight=1)
        grid.grid_columnconfigure(3, weight=0)

        # Condición
        ctk.CTkLabel(grid, text="Condición:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w")
        self.var_cond = ctk.StringVar(value="TODOS")
        ctk.CTkOptionMenu(
            grid, variable=self.var_cond,
            values=["TODOS", "BICA", "TRICA", "CUARTA"],
            width=220
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        # Evaluación
        ctk.CTkLabel(grid, text="Tipo de evaluación:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=1, sticky="w")
        self.var_eval = ctk.StringVar(value="SITUACIÓN FINAL")
        ctk.CTkOptionMenu(
            grid, variable=self.var_eval,
            values=["SITUACIÓN FINAL", "ED", "EC1", "EP", "EC2", "EC3", "EF"],
            width=220
        ).grid(row=1, column=1, sticky="w", pady=(6, 0))

        # Cursos (multi)
        ctk.CTkLabel(grid, text="Cursos (puedes elegir varios):", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=2, sticky="w")

        import tkinter as tk
        self.list_cursos = tk.Listbox(grid, selectmode=tk.MULTIPLE, height=7, exportselection=False)
        self.list_cursos.grid(row=1, column=2, sticky="nsew", pady=(6, 0))

        btns = ctk.CTkFrame(grid, fg_color="transparent")
        btns.grid(row=2, column=2, sticky="w", pady=(6, 0))
        ctk.CTkButton(btns, text="Seleccionar todos", height=30, command=self._sel_todos).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btns, text="Limpiar", height=30, command=self._limpiar).pack(side="left")

        # Acciones
        acciones = ctk.CTkFrame(grid, fg_color="transparent")
        acciones.grid(row=1, column=3, sticky="e", padx=(12, 0))

        ctk.CTkButton(
            acciones,
            text="🔎 Analizar",
            height=38,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.analizar
        ).pack(anchor="e", pady=(0, 8))

        ctk.CTkButton(
            acciones,
            text="📄 Exportar Reporte",
            height=38,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.exportar_excel
        ).pack(anchor="e")

        # Caja resumen
        caja = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        caja.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        caja.grid_rowconfigure(0, weight=1)
        caja.grid_columnconfigure(0, weight=1)

        self.box = ctk.CTkTextbox(caja, corner_radius=12)
        self.box.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        self._log("Listo. Esta vista trabaja con el parquet riesgo_academico_reporte.parquet.")

    def on_show(self):
        self._cargar_parquet()
        self._poblar_cursos()

    def _cargar_parquet(self):
        ruta = os.path.join(os.getcwd(), self.PARQUET_RIESGO_OUT)
        if not os.path.exists(ruta):
            self.df = None
            self.lbl_estado.configure(text="Estado: falta parquet riesgo (genérelo desde Riesgo Académico)")
            self._log(f"⚠️ No existe: {self.PARQUET_RIESGO_OUT}")
            return

        try:
            df = pd.read_parquet(ruta)
            if df is None or df.empty:
                self.df = pd.DataFrame()
                self.lbl_estado.configure(text="Estado: parquet riesgo vacío")
                self._log("⚠️ Parquet riesgo está vacío.")
                return

            self.df = df
            self.lbl_estado.configure(text=f"Estado: datos listos ({len(df):,} filas)")
            self._log(f"✅ Parquet cargado: {self.PARQUET_RIESGO_OUT} ({len(df):,} filas)")
        except Exception as e:
            self.df = None
            self.lbl_estado.configure(text="Estado: error leyendo parquet riesgo")
            self._log(f"❌ Error leyendo parquet riesgo: {e}")

    def _poblar_cursos(self):
        try:
            self.list_cursos.delete(0, "end")
            if self.df is None or self.df.empty:
                return
            cursos = sorted(self.df["Curso"].astype(str).str.strip().str.upper().unique().tolist())
            self.list_cursos.insert("end", "TODOS")
            for c in cursos:
                self.list_cursos.insert("end", c)
            self.list_cursos.selection_set(0)
        except Exception:
            pass

    def analizar(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Sin datos", "No hay datos para analizar.")
            return

        try:
            df = self.df.copy()

            # Condición
            cond = self.var_cond.get().strip().upper()
            if cond != "TODOS":
                df = df[df["Riesgo académico"].astype(str).str.strip().str.upper().eq(cond)].copy()

            # Cursos
            cursos = self._cursos_seleccionados()
            if cursos and "TODOS" not in cursos:
                df = df[df["Curso"].astype(str).str.strip().str.upper().isin(cursos)].copy()

            if df.empty:
                self._log("⚠️ Con esos filtros: 0 filas.")
                self._mostrar_resumen_vacio(cond, cursos)
                return

            # Evaluación
            ev = self.var_eval.get().strip().upper()
            col_nota = "Promedio Final" if ev == "SITUACIÓN FINAL" else ev

            if col_nota not in df.columns:
                raise ValueError(f"No existe la columna '{col_nota}' en el parquet de riesgo.")

            df[col_nota] = pd.to_numeric(df[col_nota], errors="coerce").fillna(0)

            total = len(df)
            no_rind = int((df[col_nota] == 0).sum())
            rind = total - no_rind

            # Aprobación (solo rindieron)
            umbral = 12.5
            df_rind = df[df[col_nota] > 0].copy()
            aprob = int((df_rind[col_nota] >= umbral).sum())
            desap = int((df_rind[col_nota] < umbral).sum())

            pct_aprob = (100.0 * aprob / rind) if rind > 0 else 0.0
            pct_desap = (100.0 * desap / rind) if rind > 0 else 0.0

            self.df_filtrado = df.copy()
            self.df_aprob = df_rind[df_rind[col_nota] >= umbral].copy()
            self.df_desap = df_rind[df_rind[col_nota] < umbral].copy()
            self.df_no_rind = df[df[col_nota] == 0].copy()

            # Resumen en pantalla
            self.box.configure(state="normal")
            self.box.delete("1.0", "end")
            self.box.insert("end", "RESUMEN DEL ANÁLISIS\n")
            self.box.insert("end", "===================\n\n")
            self.box.insert("end", f"Condición: {cond}\n")
            self.box.insert("end", f"Evaluación: {ev}\n")
            self.box.insert("end", f"Cursos: {('TODOS' if (not cursos or 'TODOS' in cursos) else ', '.join(cursos))}\n\n")
            self.box.insert("end", f"Cantidad total: {total:,}\n")
            self.box.insert("end", f"Rindieron (nota > 0): {rind:,}\n")
            self.box.insert("end", f"No rindieron (nota = 0): {no_rind:,}\n\n")
            self.box.insert("end", f"Aprobados (>= {umbral}) entre rindieron: {aprob:,} ({pct_aprob:.2f}%)\n")
            self.box.insert("end", f"Desaprobados (< {umbral}) entre rindieron: {desap:,} ({pct_desap:.2f}%)\n\n")
            self.box.insert("end", "Exporta el Excel (4 hojas) con el botón.\n")
            self.box.configure(state="disabled")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo analizar.\n\nDetalle: {e}")

    def exportar_excel(self):
        if self.df_filtrado is None:
            messagebox.showwarning("Falta análisis", "Primero presiona '🔎 Analizar'.")
            return

        ts = pd.Timestamp.now().strftime("%Y-%m-%d_%H-%M-%S")
        destino = filedialog.asksaveasfilename(
            title="Guardar análisis Riesgo Académico (Excel)",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"analisis_riesgo_{ts}.xlsx"
        )
        if not destino:
            return

        try:
            self._exportar_excel_4_hojas_estetico(destino)
            messagebox.showinfo("Exportado", f"Excel guardado:\n{destino}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar.\n\nDetalle: {e}")

    def _exportar_excel_4_hojas_estetico(self, path):
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
            self._sheet_resumen(writer)
            self._sheet_datos(writer, "APROBADOS", self.df_aprob)
            self._sheet_datos(writer, "DESAPROBADOS", self.df_desap)
            self._sheet_datos(writer, "NO_RINDIERON", self.df_no_rind)

    def _sheet_resumen(self, writer):
        wb = writer.book
        ws = wb.add_worksheet("RESUMEN")
        writer.sheets["RESUMEN"] = ws

        azul = "#003B70"
        gris_fondo = "#F7F9FC"
        gris_borde = "#D9D9D9"

        # ===== Formatos =====
        fmt_titulo = wb.add_format({"bold": True, "font_size": 18, "font_color": azul})
        fmt_sub = wb.add_format({"font_size": 11, "font_color": "#555555"})

        fmt_section = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": azul,
            "align": "center", "valign": "vcenter"
        })

        fmt_box = wb.add_format({"border": 1, "border_color": gris_borde, "bg_color": gris_fondo})
        fmt_label = wb.add_format({"bold": True, "font_color": "#333333"})
        fmt_text = wb.add_format({"font_color": "#333333"})
        fmt_wrap = wb.add_format({"font_color": "#333333", "text_wrap": True})

        fmt_kpi_num = wb.add_format({"bold": True, "font_size": 16, "align": "center", "valign": "vcenter"})
        fmt_kpi_cap = wb.add_format({"font_color": "#555555", "align": "center", "valign": "vcenter"})

        fmt_pct = wb.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter",
                                "num_format": '0.00"%"'})

        fmt_tbl_h = wb.add_format({"bold": True, "font_color": "white", "bg_color": azul,
                                "border": 1, "align": "center", "valign": "vcenter"})
        fmt_tbl_c = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})

        # ===== Layout compacto (NO expandir horizontalmente) =====
        # Usaremos solo columnas A..H (0..7)
        ws.set_column("A:A", 18)
        ws.set_column("B:B", 26)
        ws.set_column("C:C", 18)
        ws.set_column("D:D", 18)
        ws.set_column("E:E", 18)
        ws.set_column("F:F", 18)
        ws.set_column("G:G", 18)
        ws.set_column("H:H", 18)

        ws.hide_gridlines(2)

        # ===== Filtros =====
        cond = self.var_cond.get().strip().upper()
        ev = self.var_eval.get().strip().upper()
        cursos = self._cursos_seleccionados()
        cursos_txt = ("TODOS" if (not cursos or "TODOS" in cursos) else ", ".join(cursos))
        # limitar texto para que NO se desborde horizontalmente
        if len(cursos_txt) > 120:
            cursos_txt = cursos_txt[:117] + "..."

        # ===== Métricas =====
        df = self.df_filtrado.copy()
        col_nota = "Promedio Final" if ev == "SITUACIÓN FINAL" else ev
        df[col_nota] = pd.to_numeric(df[col_nota], errors="coerce").fillna(0)

        total = len(df)
        no_rind = int((df[col_nota] == 0).sum())
        rind = total - no_rind

        umbral = 12.5
        df_rind = df[df[col_nota] > 0].copy()
        aprob = int((df_rind[col_nota] >= umbral).sum())
        desap = int((df_rind[col_nota] < umbral).sum())

        pct_aprob = (100.0 * aprob / rind) if rind > 0 else 0.0
        pct_desap = (100.0 * desap / rind) if rind > 0 else 0.0

        # ===== Encabezado =====
        ws.write(0, 0, "Análisis de Riesgo Académico – UCSUR", fmt_titulo)
        ws.write(1, 0, f"Generado: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}", fmt_sub)

        # =========================================================
        # SECCIÓN 1: FILTROS (A3:H7)
        # =========================================================
        ws.merge_range("A3:H3", "FILTROS APLICADOS", fmt_section)

        # fondo “caja”
        for r in range(4, 8):
            for c in range(0, 8):
                ws.write(r, c, "", fmt_box)

        ws.write(4, 0, "Condición:", fmt_label)
        ws.write(4, 1, cond, fmt_text)

        ws.write(5, 0, "Evaluación:", fmt_label)
        ws.write(5, 1, ev, fmt_text)

        ws.write(6, 0, "Cursos:", fmt_label)
        ws.merge_range("6,1,6,7", cursos_txt, fmt_wrap)  # B7:H7 (merge en una sola línea)

        # =========================================================
        # SECCIÓN 2: KPI (A9:H12)
        # =========================================================
        ws.merge_range("A9:H9", "INDICADORES", fmt_section)

        # 3 tarjetas en una fila: TOTAL | RINDIERON | NO RINDIERON
        # TOTAL (A10:C12)
        ws.merge_range("A10:C10", "TOTAL", fmt_section)
        ws.merge_range("A11:C11", total, fmt_kpi_num)
        ws.merge_range("A12:C12", "Registros filtrados", fmt_kpi_cap)

        # RINDIERON (D10:F12)
        ws.merge_range("D10:F10", "RINDIERON", fmt_section)
        ws.merge_range("D11:F11", rind, fmt_kpi_num)
        ws.merge_range("D12:F12", "Nota > 0", fmt_kpi_cap)

        # NO RINDIERON (G10:H12)
        ws.merge_range("G10:H10", "NO RINDIERON", fmt_section)
        ws.merge_range("G11:H11", no_rind, fmt_kpi_num)
        ws.merge_range("G12:H12", "Nota = 0", fmt_kpi_cap)

        # =========================================================
        # SECCIÓN 3: PORCENTAJES (A14:H16)
        # =========================================================
        ws.merge_range("A14:H14", "APROBACIÓN (solo rindieron)", fmt_section)

        ws.merge_range("A15:D15", "% APROBADOS", fmt_section)
        ws.merge_range("A16:D16", pct_aprob, fmt_pct)

        ws.merge_range("E15:H15", "% DESAPROBADOS", fmt_section)
        ws.merge_range("E16:H16", pct_desap, fmt_pct)

        # =========================================================
        # SECCIÓN 4: GRÁFICO (A18:H34) — ABAJO, sin superposición
        # =========================================================
        ws.merge_range("A18:H18", "GRÁFICO", fmt_section)

        # Tabla auxiliar del gráfico (oculta visualmente, pero en el mismo bloque)
        # la ponemos en A20:B22 (ordenada y compacta)
        ws.write(20, 0, "Estado", fmt_tbl_h)
        ws.write(20, 1, "Cantidad", fmt_tbl_h)
        ws.write(21, 0, "APROBADOS", fmt_tbl_c)
        ws.write_number(21, 1, aprob, fmt_tbl_c)
        ws.write(22, 0, "DESAPROBADOS", fmt_tbl_c)
        ws.write_number(22, 1, desap, fmt_tbl_c)

        # Gráfico (insertado dentro del bloque, más abajo para evitar choque)
        chart = wb.add_chart({"type": "pie"})
        chart.add_series({
            "name": "Aprobados vs Desaprobados",
            "categories": ["RESUMEN", 21, 0, 22, 0],  # A22:A23
            "values":     ["RESUMEN", 21, 1, 22, 1],  # B22:B23
            "data_labels": {"percentage": True},
        })
        chart.set_title({"name": "Aprobación (solo rindieron)"})
        chart.set_legend({"position": "bottom"})

        # Insertarlo bien separado: D20 (col 3) hacia el centro del bloque
        ws.insert_chart("D20", chart, {"x_scale": 1.35, "y_scale": 1.35})

        # Ajuste final de alturas para que se vea limpio
        ws.set_row(3, 22)
        ws.set_row(9, 22)
        ws.set_row(14, 22)
        ws.set_row(18, 22)


    def _sheet_datos(self, writer, name, df):
        df = df.copy()
        start_row = 3
        df.to_excel(writer, sheet_name=name, index=False, startrow=start_row)

        wb = writer.book
        ws = writer.sheets[name]

        azul = "#003B70"
        gris = "#F2F2F2"
        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_h = wb.add_format({"bold": True, "font_color": "white", "bg_color": azul, "border": 1, "align": "center"})
        fmt_t = wb.add_format({"border": 1})
        fmt_tz = wb.add_format({"border": 1, "bg_color": gris})
        fmt_n = wb.add_format({"border": 1, "num_format": "0"})
        fmt_nz = wb.add_format({"border": 1, "num_format": "0", "bg_color": gris})

        ws.write(0, 0, name, fmt_title)

        for j, col in enumerate(df.columns):
            ws.write(start_row, j, col, fmt_h)

        ws.freeze_panes(start_row + 1, 0)
        ws.autofilter(start_row, 0, start_row + len(df), len(df.columns) - 1)

        num_cols = {"ED", "EC1", "EP", "EC2", "EC3", "EF", "Promedio Final"}

        for j, col in enumerate(df.columns):
            try:
                max_len = max(df[col].astype(str).map(len).max(), len(col))
            except Exception:
                max_len = len(col)
            ws.set_column(j, j, min(max(max_len + 2, 10), 45))

            for i in range(len(df)):
                zebra = (i % 2 == 1)
                val = df.iloc[i, j]
                if col in num_cols:
                    ws.write(start_row + 1 + i, j, val, fmt_nz if zebra else fmt_n)
                else:
                    ws.write(start_row + 1 + i, j, val, fmt_tz if zebra else fmt_t)

    # helpers
    def _sel_todos(self):
        self.list_cursos.selection_set(0, "end")

    def _limpiar(self):
        self.list_cursos.selection_clear(0, "end")

    def _cursos_seleccionados(self):
        sel = self.list_cursos.curselection()
        if not sel:
            return []
        vals = [self.list_cursos.get(i) for i in sel]
        return [str(v).strip().upper() for v in vals if str(v).strip()]

    def _mostrar_resumen_vacio(self, cond, cursos):
        self.box.configure(state="normal")
        self.box.delete("1.0", "end")
        self.box.insert("end", "RESUMEN DEL ANÁLISIS\n")
        self.box.insert("end", "===================\n\n")
        self.box.insert("end", f"Condición: {cond}\n")
        self.box.insert("end", f"Cursos: {('TODOS' if (not cursos or 'TODOS' in cursos) else ', '.join(cursos))}\n\n")
        self.box.insert("end", "Resultado: 0 filas.\n")
        self.box.configure(state="disabled")

    def _log(self, msg):
        try:
            self.box.configure(state="normal")
            self.box.insert("end", msg + "\n")
            self.box.see("end")
            self.box.configure(state="disabled")
        except Exception:
            pass
        print(msg)
