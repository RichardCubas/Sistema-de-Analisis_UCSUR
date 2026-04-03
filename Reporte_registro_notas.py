# -*- coding: utf-8 -*-
"""
SCRIPT 4 — REPORTE DE REGISTRO DE NOTAS (CON REPORTE GLOBAL)
Universidad Científica del Sur
Autor: Richard Cubas (2025)

Incluye:
- Reporte por curso
- Reporte GLOBAL (todos los cursos)
- Orden ED, EC1, EP, EC2, EC3, EF
- EF con el mismo tratamiento que todas
- Agrupación por CURSO en el PDF global
- Separadores por CURSO en el Excel global
- Exportación en PDF y Excel
"""

import os
import unicodedata
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from fpdf import FPDF
from datetime import datetime
import xlsxwriter
LOGO_FILE = "logo_ucsur.png"   # Cambiar por tu archivo real

# ======================================================
# CONFIGURACIÓN GENERAL
# ======================================================
PARQUET_FILE = "notas_filtradas_ucsur.parquet"
DEFAULT_DIR = os.path.join(os.getcwd(), "Reportes")
os.makedirs(DEFAULT_DIR, exist_ok=True)

COLOR_SI = (180, 255, 180)  # Verde
COLOR_NO = (255, 180, 180)  # Rojo

EVAL_ORDER = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]

EVAL_NAMES = {
    "ED": "Evaluación Diagnóstica",
    "EC1": "Evaluación Continua 1",
    "EP": "Evaluación Parcial",
    "EC2": "Evaluación Continua 2",
    "EC3": "Evaluación Continua 3",
    "EF": "Evaluación Final"
}

# ======================================================
# UTILIDADES
# ======================================================
def safe(s):
    return str(s).replace("—", "-").replace("\u2013", "-")


def slug(s):
    s2 = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s2 = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in s2.strip())
    return s2.replace("__", "_") or "reporte"


def limpiar_excel(df):
    return df.replace([np.nan, np.inf, -np.inf], "")

# ======================================================
# SELECCIÓN DE CARPETA
# ======================================================
def elegir_carpeta_con_ventana():
    print("\nSeleccione la carpeta donde desea guardar el archivo…")
    try:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        root.lift()
        root.attributes('-topmost', True)
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de destino", parent=root)
        root.destroy()

        if not carpeta:
            print("⚠️ No seleccionó carpeta. Se usará la carpeta por defecto.")
            return DEFAULT_DIR
        return carpeta

    except:
        return DEFAULT_DIR

# ======================================================
# CARGA DE PARQUET
# ======================================================
def cargar_df():
    if not os.path.exists(PARQUET_FILE):
        print(f"❌ No existe el archivo {PARQUET_FILE}")
        exit()

    df = pd.read_parquet(PARQUET_FILE)

    df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce").fillna(0)
    df["Evaluacion"] = df["Evaluacion"].astype(str).str.upper().str.strip()
    df["Curso"] = df["Curso"].astype(str).str.strip()
    df["Seccion"] = df["Seccion"].astype(str).str.strip()
    df["Docente"] = df["Docente"].astype(str).str.strip()

    return df

# ======================================================
# CÁLCULO DEL REPORTE
# ======================================================
def calcular_registro(df, curso, evaluacion, P):
    df_curso = df[(df["Curso"] == curso) & (df["Evaluacion"] == evaluacion)]
    if df_curso.empty:
        return pd.DataFrame(), 0, 0

    resultado = []
    for sec in sorted(df_curso["Seccion"].unique()):
        df_sec = df_curso[df_curso["Seccion"] == sec]
        total = len(df_sec)
        con_nota = (df_sec["Nota"] > 0).sum()
        porcentaje = con_nota * 100 / total if total else 0
        docente = df_sec["Docente"].iloc[0]
        cargo = "Sí" if porcentaje >= P else "No"
        resultado.append([sec, total, docente, cargo, round(porcentaje, 1)])

    df_res = pd.DataFrame(resultado, columns=["Sección", "Total", "Docente", "Cargó Notas", "% con nota"])
    df_res = df_res.sort_values(by=["Docente", "Sección"]).reset_index(drop=True)

    return df_res, (df_res["Cargó Notas"]=="Sí").sum(), (df_res["Cargó Notas"]=="No").sum()
def generar_resumen_observados(df_res, curso):
    """
    Resumen de secciones OBSERVADAS
    Docente | Curso | Sección | % cargado | Estado
    """

    columnas = ["Docente", "Curso", "Sección", "% cargado", "Estado"]

    if df_res.empty:
        return pd.DataFrame(columns=columnas)

    df_obs = df_res[df_res["Cargó Notas"] == "No"].copy()
    if df_obs.empty:
        return pd.DataFrame(columns=columnas)

    # 🔑 LÓGICA CLAVE
    if "Curso" not in df_obs.columns:
        # Reporte por curso único
        df_obs["Curso"] = curso
    # else:
    # Reporte GLOBAL → usar el curso real (NO tocar)

    df_obs["% cargado"] = df_obs["% con nota"].round(1)
    df_obs["Estado"] = "Observado"

    df_obs = df_obs[
        ["Docente", "Curso", "Sección", "% cargado", "Estado"]
    ].sort_values(
        by=["% cargado", "Docente"],
        ascending=[True, True]
    ).reset_index(drop=True)

    return df_obs
# ======================================================
# PDF
# ======================================================
class PDF(FPDF):
    def header(self):
        # Logo (ajusta width si el logo es más grande)
        if os.path.exists(LOGO_FILE):
            self.image(LOGO_FILE, x=10, y=8, w=22)

        # Encabezado
        self.set_xy(35, 10)
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 6, safe("UNIVERSIDAD CIENTÍFICA DEL SUR"), ln=1)

        self.set_x(35)
        self.set_font("Helvetica", "", 11)
        self.cell(0, 5, safe("Departamento de Cursos Básicos"), ln=1)

        self.ln(5)

def exportar_pdf(df_res, curso, eval_nombre, P, total_si, total_no, carpeta):

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"RegistroNotas_{slug(curso)}_{slug(eval_nombre)}_{ts}.pdf"
    path = os.path.join(carpeta, filename)

    # ===============================
    # CREAR PDF
    # ===============================
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "", 11)

    # ===============================
    # CUADRO RESUMEN OBSERVADOS (PRIMERO)
    # ===============================
    df_obs = generar_resumen_observados(df_res, curso)

    if not df_obs.empty:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 7, "Resumen de Secciones Observadas", ln=1)
        pdf.ln(2)

        headers = ["Docente", "Curso", "Sección", "% cargado"]
        widths = [60, 50, 30, 40]

        pdf.set_font("Helvetica", "B", 9)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, h, border=1, align="C")
        pdf.ln()

        pdf.set_font("Helvetica", "", 9)
        for _, r in df_obs.iterrows():
            pdf.set_fill_color(*COLOR_NO)
            pdf.cell(widths[0], 7, safe(r["Docente"]), border=1)
            pdf.cell(widths[1], 7, safe(r["Curso"]), border=1)
            pdf.cell(widths[2], 7, safe(r["Sección"]), border=1)
            pdf.cell(widths[3], 7, f'{r["% cargado"]}%', border=1, fill=True)
            pdf.ln()

        # Página nueva para el detalle
        pdf.add_page()

    # ===============================
    # ENCABEZADO DEL DETALLE
    # ===============================
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, safe(f"Curso: {curso}"), ln=1)
    pdf.cell(0, 6, safe(f"{eval_nombre}"), ln=1)
    pdf.cell(0, 6, f"Porcentaje mínimo requerido: {P}%", ln=1)
    pdf.ln(2)

    headers = ["Sección", "Total", "Docente", "Cargó Notas"]
    widths = [30, 25, 80, 35]

    pdf.set_font("Helvetica", "B", 10)
    for h, w in zip(headers, widths):
        pdf.cell(w, 7, safe(h), border=1, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", "", 10)

    curso_actual = None

    for _, row in df_res.iterrows():

        if "Curso" in df_res.columns:
            if curso_actual != row["Curso"]:
                curso_actual = row["Curso"]
                pdf.ln(2)
                pdf.set_font("Helvetica", "B", 11)
                pdf.cell(0, 7, safe(f"Curso: {curso_actual}"), ln=1)
                pdf.set_font("Helvetica", "", 10)

        pdf.set_fill_color(*(COLOR_SI if row["Cargó Notas"] == "Sí" else COLOR_NO))

        pdf.cell(widths[0], 7, safe(row["Sección"]), border=1)
        pdf.cell(widths[1], 7, str(row["Total"]), border=1)
        pdf.cell(widths[2], 7, safe(row["Docente"]), border=1)
        pdf.cell(widths[3], 7, safe(row["Cargó Notas"]), border=1, fill=True)
        pdf.ln()

    # ===============================
    # GRÁFICO
    # ===============================
    plt.figure(figsize=(5, 3))
    plt.bar(["Sí", "No"], [total_si, total_no], color=["green", "red"])
    plt.title("Registro de Notas por Sección")
    plt.ylabel("Cantidad")

    tmp = os.path.join(carpeta, "_graf_temp.png")
    plt.savefig(tmp, dpi=140)
    plt.close()

    pdf.ln(5)
    pdf.image(tmp, w=140)
    os.remove(tmp)

    pdf.output(path)
    print(f"✅ PDF generado: {path}")
    return path


# ======================================================
# EXCEL
# ======================================================
def exportar_excel(df_res, curso, eval_nombre, P, total_si, total_no, carpeta):

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"RegistroNotas_{slug(curso)}_{slug(eval_nombre)}_{ts}.xlsx"
    path = os.path.join(carpeta, filename)

    # ===============================
    # RESUMEN OBSERVADOS
    # ===============================
    df_obs = generar_resumen_observados(df_res, curso)

    # ===============================
    # DETALLE (LO QUE YA TENÍAS)
    # ===============================
    filas = []
    curso_actual = None

    if "Curso" in df_res.columns:
        for _, row in df_res.iterrows():
            if curso_actual != row["Curso"]:
                curso_actual = row["Curso"]
                filas.append([f"----- {curso_actual} -----", "", "", "", ""])
            filas.append([
                row["Sección"], row["Total"],
                row["Docente"], row["Cargó Notas"],
                row["% con nota"]
            ])
    else:
        filas = df_res.values.tolist()

    df_detalle = pd.DataFrame(
        filas,
        columns=["Sección", "Total", "Docente", "Cargó Notas", "% con nota"]
    )
    df_detalle = limpiar_excel(df_detalle)

    # ===============================
    # ESCRITURA EXCEL
    # ===============================
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        wb = writer.book

        # -------- HOJA 1: RESUMEN OBSERVADOS --------
        if not df_obs.empty:
            df_obs.to_excel(writer, sheet_name="Resumen_Observados", index=False)
            ws_obs = writer.sheets["Resumen_Observados"]

            fmt_header = wb.add_format({
                "bold": True, "bg_color": "#7F0000",
                "color": "white", "border": 1
            })
            fmt_no = wb.add_format({"bg_color": "#FFB4B4", "border": 1})

            for col, name in enumerate(df_obs.columns):
                ws_obs.write(0, col, name, fmt_header)
                ws_obs.set_column(col, col, 22)

            for r in range(1, len(df_obs) + 1):
                ws_obs.write(r, 3, f'{df_obs.iloc[r-1]["% cargado"]}%', fmt_no)

        # -------- HOJA 2: DETALLE --------
        df_detalle.to_excel(writer, sheet_name="Detalle", index=False)
        ws = writer.sheets["Detalle"]

        fmt_header = wb.add_format({
            "bold": True, "bg_color": "#003B70",
            "color": "white", "border": 1
        })
        fmt_si = wb.add_format({"bg_color": "#B4FFC4"})
        fmt_no = wb.add_format({"bg_color": "#FFB4B4"})

        for col in range(len(df_detalle.columns)):
            ws.write(0, col, df_detalle.columns[col], fmt_header)
            ws.set_column(col, col, 22)

        for r in range(1, len(df_detalle) + 1):
            val = df_detalle.iloc[r - 1]["Cargó Notas"]
            if val in ["Sí", "No"]:
                ws.write(r, 3, val, fmt_si if val == "Sí" else fmt_no)

        # -------- HOJA 3: GRÁFICO --------
        ws2 = wb.add_worksheet("Gráfico")
        ws2.write_row("A1", ["Estado", "Cantidad"])
        ws2.write_row("A2", ["Sí", total_si])
        ws2.write_row("A3", ["No", total_no])

        chart = wb.add_chart({"type": "column"})
        chart.add_series({
            "categories": "=Gráfico!A2:A3",
            "values": "=Gráfico!B2:B3",
            "data_labels": {"value": True}
        })
        chart.set_title({"name": "Registro de Notas por Sección"})
        ws2.insert_chart("D2", chart)

    print(f"✅ Excel generado: {path}")

# ======================================================
# REPORTE GLOBAL
# ======================================================
def reporte_global(df):

    evals_disponibles = [ev for ev in EVAL_ORDER if ev in df["Evaluacion"].unique()]

    print("\nEVALUACIONES (GLOBAL):")
    for i, ev in enumerate(evals_disponibles, 1):
        print(f"{i}. {EVAL_NAMES[ev]} ({ev})")
    print("0. Volver")

    op = input("Seleccione evaluación: ").strip()
    if op == "0":
        return

    if not op.isdigit() or not (1 <= int(op) <= len(evals_disponibles)):
        print("⚠️ Opción inválida.")
        return

    ev = evals_disponibles[int(op) - 1]
    eval_nombre = EVAL_NAMES[ev]

    while True:
        try:
            P = float(input("Mínimo requerido (0-100): "))
            if 0 <= P <= 100:
                break
        except:
            pass
        print("⚠️ Ingrese número válido.")

    df_eval = df[df["Evaluacion"] == ev]
    if df_eval.empty:
        print("⚠️ No hay datos.")
        return

    resultados = []
    for (curso, sec), g in df_eval.groupby(["Curso", "Seccion"]):
        total = len(g)
        con = (g["Nota"] > 0).sum()
        porc = con * 100 / total if total else 0
        docente = g["Docente"].iloc[0]
        cargo = "Sí" if porc >= P else "No"
        resultados.append([curso, sec, total, docente, cargo, round(porc, 1)])

    df_res = pd.DataFrame(resultados, columns=["Curso", "Sección", "Total", "Docente", "Cargó Notas", "% con nota"])
    df_res = df_res.sort_values(by=["Curso", "Sección"]).reset_index(drop=True)

    print("\n───────────────")
    print("GENERAR REPORTE GLOBAL")
    print("───────────────")
    print("1. PDF")
    print("2. Excel")
    print("3. Volver")

    op2 = input("Seleccione opción: ").strip()
    if op2 == "3":
        return

    carpeta = elegir_carpeta_con_ventana()
    nombre_curso = "TODOS LOS CURSOS"

    total_si = (df_res["Cargó Notas"] == "Sí").sum()
    total_no = len(df_res) - total_si

    if op2 == "1":
        exportar_pdf(df_res, nombre_curso, eval_nombre, P, total_si, total_no, carpeta)
    elif op2 == "2":
        exportar_excel(df_res, nombre_curso, eval_nombre, P, total_si, total_no, carpeta)

# ======================================================
# MENÚ PRINCIPAL
# ======================================================
def menu_principal():
    df = cargar_df()
    cursos = sorted(df["Curso"].unique().tolist())

    while True:
        print("\n==============================")
        print(" REPORTE DE REGISTRO DE NOTAS")
        print("==============================")
        print("0. REPORTE GLOBAL (TODOS LOS CURSOS)")

        for i, c in enumerate(cursos, 1):
            print(f"{i}. {c}")

        print("X. Salir")
        op = input("Seleccione una opción: ").strip()

        if op.upper() == "X":
            break

        if op == "0":
            reporte_global(df)
            continue

        if not op.isdigit() or not (1 <= int(op) <= len(cursos)):
            print("⚠️ Opción inválida.")
            continue

        curso = cursos[int(op) - 1]
        df_curso = df[df["Curso"] == curso]

        evals_validas = [ev for ev in EVAL_ORDER if ev in df_curso["Evaluacion"].unique()]

        while True:
            print("\nEVALUACIONES DISPONIBLES:")
            for i, ev in enumerate(evals_validas, 1):
                print(f"{i}. {EVAL_NAMES[ev]} ({ev})")
            print("0. Volver")

            op2 = input("Seleccione evaluación: ").strip()
            if op2 == "0":
                break

            if not op2.isdigit() or not (1 <= int(op2) <= len(evals_validas)):
                print("⚠️ Opción inválida.")
                continue

            ev = evals_validas[int(op2) - 1]
            eval_nombre = EVAL_NAMES[ev]

            while True:
                try:
                    P = float(input("Mínimo requerido (0-100): "))
                    if 0 <= P <= 100:
                        break
                except:
                    pass
                print("⚠️ Ingrese número válido.")

            df_res, total_si, total_no = calcular_registro(df, curso, ev, P)

            print("\n───────────────")
            print("GENERAR REPORTE")
            print("───────────────")
            print("1. PDF")
            print("2. Excel")
            print("3. Volver")

            op3 = input("Seleccione opción: ").strip()

            if op3 == "1":
                carpeta = elegir_carpeta_con_ventana()
                exportar_pdf(df_res, curso, eval_nombre, P, total_si, total_no, carpeta)
            elif op3 == "2":
                carpeta = elegir_carpeta_con_ventana()
                exportar_excel(df_res, curso, eval_nombre, P, total_si, total_no, carpeta)
            elif op3 == "3":
                break

if __name__ == "__main__":
    menu_principal()
