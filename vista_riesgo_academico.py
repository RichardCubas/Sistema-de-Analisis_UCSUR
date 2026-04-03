# ============================================================
# VISTA: RIESGO ACADÉMICO (UCSUR)
# ARCHIVO: vista_riesgo_academico.py
# ============================================================
# AJUSTES APLICADOS:
# 1) Botón principal: "⚙️ Generar y descargar"
#    - Genera parquet de riesgo (archivo interno persistente)
#    - Exporta Excel estético (descarga) en el momento
# 2) Se elimina panel inferior de exportación (ya no hay botones abajo)
# 3) El Excel de riesgo NO se vuelve a pedir siempre:
#    - Se guarda una copia local (riesgo_academico_input.xlsx) para uso posterior
#    - La vista ofrece "Cargar Excel" solo si deseas reemplazarlo
# 4) Ya no se muestra preview de 40 filas.
#    - Se muestra solo un resumen
# 5) Se agrega botón "📊 Análisis" que abre otra vista:
#    - Vista_analisis_riesgo.py
# 6) NUEVO: Al descargar el Excel, se crean 2 hojas:
#    - "RIESGO" con el reporte cruzado
#    - "NO_ENCONTRADOS" con los alumnos (ID+Curso) del Excel de riesgo que NO aparecen en Query
#       * Incluye NOMBRE (columna G del Excel)
# ============================================================

import os
import sys
import shutil
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox


class FrameRiesgoAcademico(ctk.CTkFrame):

    INPUT_RIESGO_LOCAL = "riesgo_academico_input.xlsx"
    PARQUET_RIESGO_OUT = "riesgo_academico_reporte.parquet"

    def __init__(self, parent, controller):
        super().__init__(parent, fg_color="#F5F5F5", corner_radius=0)
        self.controller = controller

        self.df_excel_riesgo = None
        self.df_resultado = None
        self.df_no_encontrados = None
        self.path_excel = None

        # ====================================================
        # ENCABEZADO
        # ====================================================
        ctk.CTkLabel(
            self,
            text="⚠️ Riesgo Académico",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#003B70"
        ).pack(anchor="w", padx=18, pady=(18, 4))

        self.lbl_estado = ctk.CTkLabel(
            self,
            text="Estado: (sin Excel de riesgo cargado)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#333"
        )
        self.lbl_estado.pack(anchor="w", padx=18, pady=(0, 10))

        # ====================================================
        # PANEL CONTROLES
        # ====================================================
        panel = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        panel.pack(fill="x", padx=18, pady=(0, 12))

        fila1 = ctk.CTkFrame(panel, fg_color="transparent")
        fila1.pack(fill="x", padx=14, pady=(12, 6))

        self.lbl_excel = ctk.CTkLabel(
            fila1,
            text="Excel: (no cargado)",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#333"
        )
        self.lbl_excel.pack(side="left")

        ctk.CTkButton(
            fila1,
            text="📎 Cargar/Reemplazar Excel",
            height=38,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.cargar_excel
        ).pack(side="right")

        fila2 = ctk.CTkFrame(panel, fg_color="transparent")
        fila2.pack(fill="x", padx=14, pady=(6, 12))

        ctk.CTkButton(
            fila2,
            text="⚙️ Generar y descargar",
            height=38,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.generar_y_descargar
        ).pack(side="right")

        ctk.CTkButton(
            fila2,
            text="📊 Análisis",
            height=38,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.ir_a_analisis
        ).pack(side="right", padx=(0, 10))

        # ====================================================
        # LOG / MENSAJES
        # ====================================================
        caja = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        caja.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        caja.grid_rowconfigure(0, weight=1)
        caja.grid_columnconfigure(0, weight=1)

        self.box = ctk.CTkTextbox(caja, corner_radius=12)
        self.box.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        self._log("Listo. Cargue el Excel de riesgo o use el guardado (si existe).")

    # =========================================================
    # on_show: cargar Excel persistido si existe
    # =========================================================
    def on_show(self):
        self._cargar_excel_guardado_si_existe()

    # =========================================================
    # Navegar a vista análisis
    # =========================================================
    def ir_a_analisis(self):
        """
        Al entrar a Análisis:
        1) Regenera y REEMPLAZA el parquet riesgo_academico_reporte.parquet
           (misma lógica que Generar y descargar, pero SIN descargar Excel).
        2) Abre la vista analisis_riesgo.
        """
        self._cargar_excel_guardado_si_existe()
        if self.df_excel_riesgo is None or self.df_excel_riesgo.empty:
            messagebox.showwarning(
                "Falta Excel",
                "No hay Excel de Riesgo cargado/guardado.\n\n"
                "Primero use '📎 Cargar/Reemplazar Excel'."
            )
            return

        df_base = self._obtener_df_base()
        if df_base is None or df_base.empty:
            messagebox.showwarning(
                "Falta Parquet",
                "No encontré el parquet base (notas_filtradas_ucsur.parquet).\n\n"
                "Primero actualice desde '📥 Extraer Query'."
            )
            return

        try:
            self._log("📊 Preparando análisis: regenerando parquet de riesgo (reemplazo)...")
            out = self._generar_df_salida(df_base)

            if out is None or out.empty:
                self._log("⚠️ No hubo coincidencias para riesgo. No se abrirá Análisis.")
                messagebox.showinfo(
                    "Sin coincidencias",
                    "No hubo coincidencias (ID + Curso) entre Excel y parquet.\n\n"
                    "No es posible generar análisis."
                )
                return

            parquet_path = self._ruta_local(self.PARQUET_RIESGO_OUT)
            out.to_parquet(parquet_path, index=False)

            self.df_resultado = out
            self._log(f"✅ Parquet actualizado para análisis: {self.PARQUET_RIESGO_OUT} ({len(out):,} filas)")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo regenerar parquet de riesgo.\n\nDetalle: {e}")
            return

        if hasattr(self.controller, "mostrar_vista"):
            self.controller.mostrar_vista("analisis_riesgo")
        else:
            self._log("⚠️ No encuentro mostrar_vista() en el controller.")

    # =========================================================
    # Cargar Excel (y guardarlo localmente)
    # =========================================================
    def cargar_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel de Riesgo Académico",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            shutil.copyfile(path, self._ruta_local(self.INPUT_RIESGO_LOCAL))

            self.path_excel = path
            self.lbl_excel.configure(text=f"Excel: {os.path.basename(path)}")
            self._log(f"✅ Excel cargado y guardado localmente como: {self.INPUT_RIESGO_LOCAL}")

            self.df_excel_riesgo = self._leer_excel_riesgo(path)
            self.df_resultado = None
            self.df_no_encontrados = None

            self.lbl_estado.configure(
                text=f"Estado: Excel listo ({len(self.df_excel_riesgo):,} filas riesgo válidas)"
            )
            messagebox.showinfo("Listo", f"Excel cargado: {len(self.df_excel_riesgo):,} filas válidas de riesgo.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar/guardar el Excel.\n\nDetalle: {e}")

    def _cargar_excel_guardado_si_existe(self):
        ruta = self._ruta_local(self.INPUT_RIESGO_LOCAL)
        if os.path.exists(ruta) and self.df_excel_riesgo is None:
            try:
                self.df_excel_riesgo = self._leer_excel_riesgo(ruta)
                self.path_excel = ruta
                self.lbl_excel.configure(text=f"Excel: {os.path.basename(ruta)} (guardado)")

                self.lbl_estado.configure(
                    text=f"Estado: Excel guardado listo ({len(self.df_excel_riesgo):,} filas riesgo válidas)"
                )
                self._log(f"📌 Se cargó el Excel guardado: {self.INPUT_RIESGO_LOCAL}")

            except Exception as e:
                self._log(f"⚠️ Existe Excel guardado pero falló la lectura: {e}")

    # =========================================================
    # Generar parquet y descargar Excel (2 hojas)
    # =========================================================
    def generar_y_descargar(self):
        if self.df_excel_riesgo is None or self.df_excel_riesgo.empty:
            messagebox.showwarning("Falta Excel", "Primero carga el Excel de Riesgo Académico (o verifica el guardado).")
            return

        df_base = self._obtener_df_base()
        if df_base is None or df_base.empty:
            messagebox.showwarning(
                "Falta Parquet",
                "No encontré el parquet base (notas_filtradas_ucsur.parquet).\n\n"
                "Primero carga/actualiza el parquet desde '📥 Extraer Query'."
            )
            return

        try:
            self._log("⚙️ Generando cruce y reporte...")

            out = self._generar_df_salida(df_base)
            if out is None or out.empty:
                self.df_resultado = pd.DataFrame()
                self.df_no_encontrados = self._obtener_no_encontrados(df_base)
                self._log("⚠️ No hubo coincidencias entre Excel riesgo y parquet base.")
                messagebox.showinfo("Sin coincidencias", "No hubo coincidencias (ID + Curso) entre Excel y parquet.")
                return

            self.df_resultado = out

            # ✅ NO_ENCONTRADOS con NOMBRE
            self.df_no_encontrados = self._obtener_no_encontrados(df_base)
            self._log(f"✅ Encontrados: {len(out):,} | ❗ No encontrados: {len(self.df_no_encontrados):,}")

            parquet_path = self._ruta_local(self.PARQUET_RIESGO_OUT)
            out.to_parquet(parquet_path, index=False)
            self._log(f"✅ Parquet generado: {self.PARQUET_RIESGO_OUT}")

            ts = pd.Timestamp.now().strftime("%Y-%m-%d_%H-%M-%S")
            destino = filedialog.asksaveasfilename(
                title="Guardar reporte Riesgo Académico (Excel)",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"riesgo_academico_reporte_{ts}.xlsx"
            )

            if not destino:
                self._log("❌ Descarga de Excel cancelada por el usuario.")
                messagebox.showinfo("Listo", "Parquet generado.\nExcel: descarga cancelada.")
                return

            self._exportar_excel_estetico(out, destino, self.df_no_encontrados)
            self._log(f"📄 Excel exportado: {destino}")

            self.lbl_estado.configure(text=f"Estado: Reporte generado ({len(out):,} filas)")
            messagebox.showinfo("Completado", "Reporte generado.\n\nParquet interno listo y Excel descargado (2 hojas).")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar/descargar.\n\nDetalle: {e}")

    # =========================================================
    # Leer excel riesgo:
    # - F=ID (5), G=NOMBRE (6), I=CURSO (8), J=RIESGO (9), M=SEDE (12)
    # =========================================================
    def _leer_excel_riesgo(self, path_excel):
        df = pd.read_excel(path_excel, header=0)
        if df is None or df.empty:
            raise ValueError("El Excel está vacío.")

        ncols = df.shape[1]
        requeridas = [
            (5, "F (ID alumno)"),
            (6, "G (Alumno)"),
            (8, "I (Curso)"),
            (9, "J (Riesgo)"),
            (12, "M (Campus/Sede)")
        ]
        for idx, name in requeridas:
            if ncols <= idx:
                raise ValueError(
                    f"El Excel no tiene suficientes columnas. Falta la columna {name}.\n"
                    f"Columnas detectadas: {ncols}"
                )

        df_r = pd.DataFrame({
            "Sede": df.iloc[:, 12],
            "CodigoAlumno": df.iloc[:, 5],
            "AlumnoExcel": df.iloc[:, 6],        # ✅ NUEVO
            "Curso": df.iloc[:, 8],
            "RiesgoAcademico": df.iloc[:, 9],
        })

        df_r["Sede"] = df_r["Sede"].astype(str).str.strip().str.upper()
        df_r["CodigoAlumno"] = df_r["CodigoAlumno"].astype(str).str.strip()
        df_r["AlumnoExcel"] = df_r["AlumnoExcel"].astype(str).str.strip().str.upper()
        df_r["Curso"] = df_r["Curso"].astype(str).str.strip().str.upper()
        df_r["RiesgoAcademico"] = df_r["RiesgoAcademico"].astype(str).str.strip().str.upper()

        df_r = df_r[
            df_r["CodigoAlumno"].ne("") &
            df_r["Curso"].ne("") &
            df_r["RiesgoAcademico"].ne("")
        ].copy()

        df_r["Veces"] = df_r["RiesgoAcademico"].apply(self._riesgo_a_veces)
        df_r = df_r[df_r["Veces"].isin([2, 3, 4])].copy()

        # ✅ Dejamos el nombre del Excel para "NO_ENCONTRADOS"
        df_r = df_r.drop_duplicates(subset=["Sede", "CodigoAlumno", "Curso", "RiesgoAcademico", "Veces", "AlumnoExcel"])
        return df_r

    # =========================================================
    # Generar DF salida (cruce con parquet base)
    # =========================================================
    def _generar_df_salida(self, df_base):
        base_cols = {"CodigoAlumno", "Alumno", "Curso", "Seccion", "Carrera", "Docente", "Evaluacion", "Nota"}
        faltan = [c for c in base_cols if c not in df_base.columns]
        if faltan:
            raise ValueError(
                "El DF base no tiene el formato esperado.\n"
                f"Faltan columnas: {faltan}\n\n"
                f"Columnas actuales: {list(df_base.columns)}"
            )

        dfb = df_base.copy()
        dfb["CodigoAlumno"] = dfb["CodigoAlumno"].astype(str).str.strip()
        dfb["Curso"] = dfb["Curso"].astype(str).str.strip().str.upper()
        dfb["Evaluacion"] = dfb["Evaluacion"].astype(str).str.strip().str.upper()
        dfb["Nota"] = pd.to_numeric(dfb["Nota"], errors="coerce").fillna(0.0)

        llaves = self.df_excel_riesgo[["CodigoAlumno", "Curso"]].drop_duplicates()

        # ✅ GARANTÍA: filtrado por (ID + Curso)
        df_filtrado = dfb.merge(llaves, on=["CodigoAlumno", "Curso"], how="inner")
        if df_filtrado.empty:
            return pd.DataFrame()

        df_filtrado["EvalStd"] = df_filtrado["Evaluacion"].apply(self._normalizar_evaluacion)

        evals_objetivo = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]
        df_filtrado = df_filtrado[df_filtrado["EvalStd"].isin(evals_objetivo)].copy()

        idx_cols = ["CodigoAlumno", "Alumno", "Carrera", "Curso", "Seccion", "Docente"]
        for c in ["Alumno", "Carrera", "Seccion", "Docente"]:
            df_filtrado[c] = df_filtrado[c].astype(str).str.strip().str.upper()

        pivot = (
            df_filtrado
            .pivot_table(index=idx_cols, columns="EvalStd", values="Nota", aggfunc="max", fill_value=0.0)
            .reset_index()
        )

        for ev in evals_objetivo:
            if ev not in pivot.columns:
                pivot[ev] = 0.0

        excel_info = self.df_excel_riesgo[["Sede", "CodigoAlumno", "Curso", "RiesgoAcademico", "Veces"]].copy()
        excel_info["RiesgoAcademico"] = excel_info["RiesgoAcademico"].astype(str).str.strip().str.upper()

        out = pivot.merge(excel_info, on=["CodigoAlumno", "Curso"], how="left")

        out["Sede"] = out["Sede"].fillna("").astype(str).str.strip().str.upper()
        out["RiesgoAcademico"] = out["RiesgoAcademico"].fillna("").astype(str).str.strip().str.upper()

        out["Promedio Final"] = out.apply(self._promedio_final, axis=1)

        out = out.rename(columns={
            "Seccion": "Sección",
            "RiesgoAcademico": "Riesgo académico",
            "CodigoAlumno": "ID de alumno"
        })

        cols_final = [
            "Sede", "ID de alumno", "Alumno", "Carrera", "Curso", "Sección",
            "Riesgo académico", "Docente", "ED", "EC1", "EP", "EC2", "EC3", "EF",
            "Promedio Final"
        ]
        out = out[cols_final].copy()

        if {"Sede", "Curso", "Riesgo académico", "Alumno"}.issubset(set(out.columns)):
            out = out.sort_values(["Sede", "Curso", "Riesgo académico", "Alumno"], ascending=[True, True, True, True])

        return out

    # =========================================================
    # NO_ENCONTRADOS: ahora incluye "Alumno" desde columna G (AlumnoExcel)
    # =========================================================
    def _obtener_no_encontrados(self, df_base: pd.DataFrame) -> pd.DataFrame:
        if self.df_excel_riesgo is None or self.df_excel_riesgo.empty:
            return pd.DataFrame(columns=["Sede", "ID de alumno", "Alumno", "Curso", "Riesgo académico", "Veces"])

        dfb = df_base.copy()
        dfb["CodigoAlumno"] = dfb["CodigoAlumno"].astype(str).str.strip()
        dfb["Curso"] = dfb["Curso"].astype(str).str.strip().str.upper()

        base_keys = dfb[["CodigoAlumno", "Curso"]].drop_duplicates()

        ex = self.df_excel_riesgo.copy()
        ex["CodigoAlumno"] = ex["CodigoAlumno"].astype(str).str.strip()
        ex["Curso"] = ex["Curso"].astype(str).str.strip().str.upper()

        m = ex.merge(base_keys, on=["CodigoAlumno", "Curso"], how="left", indicator=True)
        no_ok = m[m["_merge"].eq("left_only")].copy()
        no_ok.drop(columns=["_merge"], inplace=True)

        no_ok = no_ok.rename(columns={
            "CodigoAlumno": "ID de alumno",
            "AlumnoExcel": "Alumno",
            "RiesgoAcademico": "Riesgo académico"
        })

        cols = ["Sede", "ID de alumno", "Alumno", "Curso", "Riesgo académico", "Veces"]
        for c in cols:
            if c not in no_ok.columns:
                no_ok[c] = ""

        no_ok["Alumno"] = no_ok["Alumno"].fillna("").astype(str).str.strip().str.upper()

        no_ok = no_ok[cols].drop_duplicates()
        no_ok = no_ok.sort_values(["Sede", "Curso", "Riesgo académico", "Alumno", "ID de alumno"], ascending=True)
        return no_ok

    # =========================================================
    # Exportar Excel estético (2 hojas)
    # =========================================================
    def _exportar_excel_estetico(self, df, path, df_no_encontrados=None):
        with pd.ExcelWriter(path, engine="xlsxwriter") as writer:

            # -------------------------------
            # HOJA 1: RIESGO
            # -------------------------------
            sheet = "RIESGO"
            start_row = 4
            df.to_excel(writer, sheet_name=sheet, index=False, startrow=start_row)

            wb = writer.book
            ws = writer.sheets[sheet]

            azul = "#003B70"
            gris_claro = "#F2F2F2"

            fmt_titulo = wb.add_format({"bold": True, "font_size": 16})
            fmt_sub = wb.add_format({"font_size": 11, "font_color": "#444444"})
            fmt_header = wb.add_format({
                "bold": True, "font_color": "white",
                "bg_color": azul,
                "border": 1, "align": "center", "valign": "vcenter"
            })
            fmt_text = wb.add_format({"border": 1, "valign": "vcenter"})
            fmt_num = wb.add_format({"border": 1, "valign": "vcenter", "num_format": "0"})
            fmt_text_z = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": gris_claro})
            fmt_num_z = wb.add_format({"border": 1, "valign": "vcenter", "num_format": "0", "bg_color": gris_claro})

            ws.write(0, 0, "Reporte de Riesgo Académico – UCSUR", fmt_titulo)
            ws.write(1, 0, f"Generado: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}", fmt_sub)
            if self.path_excel:
                ws.write(2, 0, f"Fuente Excel: {os.path.basename(self.path_excel)}", fmt_sub)

            for col_idx, col_name in enumerate(df.columns):
                ws.write(start_row, col_idx, col_name, fmt_header)

            ws.freeze_panes(start_row + 1, 0)
            ws.autofilter(start_row, 0, start_row + len(df), len(df.columns) - 1)

            num_cols = {"ED", "EC1", "EP", "EC2", "EC3", "EF", "Promedio Final"}

            for col_idx, col_name in enumerate(df.columns):
                try:
                    max_len = max(df[col_name].astype(str).map(len).max(), len(col_name))
                except Exception:
                    max_len = len(col_name)
                ws.set_column(col_idx, col_idx, min(max(max_len + 2, 10), 45))

                for r in range(len(df)):
                    val = df.iloc[r, col_idx]
                    zebra = (r % 2 == 1)
                    if col_name in num_cols:
                        ws.write(start_row + 1 + r, col_idx, val, fmt_num_z if zebra else fmt_num)
                    else:
                        ws.write(start_row + 1 + r, col_idx, val, fmt_text_z if zebra else fmt_text)

            # -------------------------------
            # HOJA 2: NO_ENCONTRADOS (con Alumno)
            # -------------------------------
            if df_no_encontrados is None:
                df_no_encontrados = pd.DataFrame(columns=["Sede", "ID de alumno", "Alumno", "Curso", "Riesgo académico", "Veces"])

            sheet2 = "NO_ENCONTRADOS"
            start_row2 = 4
            df_no_encontrados.to_excel(writer, sheet_name=sheet2, index=False, startrow=start_row2)
            ws2 = writer.sheets[sheet2]

            fmt_titulo2 = wb.add_format({"bold": True, "font_size": 16})
            fmt_sub2 = wb.add_format({"font_size": 11, "font_color": "#444444"})
            fmt_header2 = wb.add_format({
                "bold": True, "font_color": "white",
                "bg_color": azul,
                "border": 1, "align": "center", "valign": "vcenter"
            })
            fmt_text2 = wb.add_format({"border": 1, "valign": "vcenter"})
            fmt_text2z = wb.add_format({"border": 1, "valign": "vcenter", "bg_color": gris_claro})

            ws2.write(0, 0, "Alumnos/Curso del Excel que NO se encontraron en el Query", fmt_titulo2)
            ws2.write(1, 0, f"Generado: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}", fmt_sub2)
            ws2.write(2, 0, "Criterio: coincidencia exacta por (ID de alumno + Curso).", fmt_sub2)

            for col_idx, col_name in enumerate(df_no_encontrados.columns):
                ws2.write(start_row2, col_idx, col_name, fmt_header2)

            ws2.freeze_panes(start_row2 + 1, 0)
            ws2.autofilter(start_row2, 0, start_row2 + len(df_no_encontrados), len(df_no_encontrados.columns) - 1)

            for col_idx, col_name in enumerate(df_no_encontrados.columns):
                try:
                    max_len = max(df_no_encontrados[col_name].astype(str).map(len).max(), len(col_name))
                except Exception:
                    max_len = len(col_name)
                ws2.set_column(col_idx, col_idx, min(max(max_len + 2, 10), 45))

                for r in range(len(df_no_encontrados)):
                    val = df_no_encontrados.iloc[r, col_idx]
                    zebra = (r % 2 == 1)
                    ws2.write(start_row2 + 1 + r, col_idx, val, fmt_text2z if zebra else fmt_text2)

    # =========================================================
    # Obtener parquet base (notas_filtradas_ucsur.parquet)
    # =========================================================
    def _obtener_df_base(self):
        for attr in ["df", "df_base", "df_global", "dataframe", "DF"]:
            if hasattr(self.controller, attr):
                df = getattr(self.controller, attr)
                if isinstance(df, pd.DataFrame) and not df.empty:
                    return df

        try:
            fq = self.controller.frames.get("query", None)
            if fq is not None:
                for m in ["cargar_df", "get_df", "obtener_df"]:
                    if hasattr(fq, m) and callable(getattr(fq, m)):
                        df = getattr(fq, m)()
                        if isinstance(df, pd.DataFrame) and not df.empty:
                            return df
        except Exception:
            pass

        parquet_name = "notas_filtradas_ucsur.parquet"
        posibles_rutas = [os.path.join(os.getcwd(), parquet_name)]
        try:
            posibles_rutas.append(os.path.join(os.path.dirname(__file__), parquet_name))
        except Exception:
            pass
        try:
            posibles_rutas.append(os.path.join(os.path.dirname(sys.executable), parquet_name))
        except Exception:
            pass
        try:
            if hasattr(sys, "_MEIPASS"):
                posibles_rutas.append(os.path.join(sys._MEIPASS, parquet_name))
        except Exception:
            pass

        for ruta in posibles_rutas:
            if os.path.exists(ruta):
                try:
                    df = pd.read_parquet(ruta)
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        return df
                except Exception:
                    pass

        return None

    # =========================================================
    # Helpers
    # =========================================================
    def _ruta_local(self, filename: str) -> str:
        return os.path.join(os.getcwd(), filename)

    def _log(self, msg: str):
        try:
            self.box.configure(state="normal")
            self.box.insert("end", msg + "\n")
            self.box.see("end")
            self.box.configure(state="disabled")
        except Exception:
            pass
        print(msg)

    def _riesgo_a_veces(self, s: str) -> int:
        s = str(s).strip().upper()
        if "BIC" in s:
            return 2
        if "TRIC" in s:
            return 3
        if "CUAR" in s:
            return 4
        return 0

    def _normalizar_evaluacion(self, s: str) -> str:
        t = str(s).strip().upper()
        t = t.replace(" ", "").replace("-", "").replace("_", "")

        if t in {"ED", "EC1", "EP", "EC2", "EC3", "EF"}:
            return t
        if t in {"E.D", "EVALUACIONDIAGNOSTICA", "DIAGNOSTICO", "DIAGNOSTICA"}:
            return "ED"
        if t in {"E.C1", "EVALUACIONCONTINUA1", "CONTINUA1", "EC01"}:
            return "EC1"
        if t in {"E.P", "EXAMENPARCIAL", "PARCIAL"}:
            return "EP"
        if t in {"E.C2", "EVALUACIONCONTINUA2", "CONTINUA2"}:
            return "EC2"
        if t in {"E.C3", "EVALUACIONCONTINUA3", "CONTINUA3"}:
            return "EC3"
        if t in {"E.F", "EXAMENFINAL", "FINAL"}:
            return "EF"

        if "EC" in t and "1" in t:
            return "EC1"
        if "EC" in t and "2" in t:
            return "EC2"
        if "EC" in t and "3" in t:
            return "EC3"
        if "PARC" in t:
            return "EP"
        if "FIN" in t:
            return "EF"
        if "DIAG" in t:
            return "ED"

        return t

    def _round_half_up(self, x: float) -> int:
        try:
            x = float(x)
        except Exception:
            return 0
        return int((x + 0.5) // 1)

    def _promedio_final(self, row) -> int:
        ec1 = float(row.get("EC1", 0) or 0)
        ep = float(row.get("EP", 0) or 0)
        ec2 = float(row.get("EC2", 0) or 0)
        ec3 = float(row.get("EC3", 0) or 0)
        ef = float(row.get("EF", 0) or 0)

        prom = 0.18 * ec1 + 0.20 * ep + 0.18 * ec2 + 0.19 * ec3 + 0.25 * ef
        return self._round_half_up(prom)
