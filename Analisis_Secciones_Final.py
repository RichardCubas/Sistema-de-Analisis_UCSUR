# =========================
# PARTE 1/2
# =========================
# Analisis_Secciones_Final.py
# -*- coding: utf-8 -*-
"""
SCRIPT 9.B — ANÁLISIS POR CURSO, SECCIÓN Y DOCENTE (Compacto)

AJUSTES (enero 2026):
- % Aprobados y % Desaprobados: se calculan SOLO sobre "Rindieron" (Nota > 0).
- Se reporta siempre: Total | No rindieron | Rindieron | % Aprob. | % Desaprob. | % No rind. | Prom.
  * % No rind. se calcula sobre Total.
- Situación Final:
  * EstadoFinal: si todas las evaluaciones están en 0 => "No rindió (todas 0)" (igual que antes).
  * Para % finales, se descarta del denominador a quienes no rindieron ningún examen (promedio 0).
- Nuevo cálculo de FINAL:
  FINAL = 0.18*R(EC1)+0.20*R(EP)+0.18*R(EC2)+0.19*R(EC3)+0.2*0.18*R(EF)
  con redondeo half-up: x.5 -> x+1.
- Aprobado: nota >= 12.5 (por evaluación y en FINAL).
- BLINDAJE: Para evitar duplicados por alumno, todo conteo/porcentaje se calcula a nivel de alumno (max Nota por alumno).
"""

import os, sys, unicodedata, tempfile, warnings
from datetime import datetime

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import xlsxwriter

# Para elegir carpeta de destino (como en el Script 4)
try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:
    tk = None
    filedialog = None

warnings.filterwarnings("ignore", category=FutureWarning)

# ==========================
# CONFIGURACIÓN GENERAL
# ==========================
PARQUET_IN = "notas_filtradas_ucsur.parquet"

EVALS = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]
WEIGHTS = {"ED": 0.0, "EC1": 0.18, "EP": 0.20, "EC2": 0.18, "EC3": 0.19, "EF": 0.25}  # (se mantiene, por si lo usas en otro lado)
FINAL_NAME = "FINAL"

COL_UCSUR_AZUL = "#003B70"
LOGO_PATH = "logo_ucsur.png"

APROBADO_MIN = 12.5

# Pesos FINAL solicitados (EF usa 0.2*0.18)
FINAL_WEIGHTS = {
    "EC1": 0.18,
    "EP":  0.20,
    "EC2": 0.18,
    "EC3": 0.19,
    "EF":  0.2 * 0.18,
}

# Nombres descriptivos para informes
EVAL_NAMES = {
    "ED":   "Evaluación Diagnóstica",
    "EC1":  "Evaluación Continua 1",
    "EP":   "Evaluación Parcial",
    "EC2":  "Evaluación Continua 2",
    "EC3":  "Evaluación Continua 3",
    "EF":   "Evaluación Final",
    FINAL_NAME: "Situación Final",
}

def nombre_eval(ev: str) -> str:
    return EVAL_NAMES.get(ev, ev)

# ==========================
# ORDEN FIJO DE EVALUACIONES
# ==========================
def ordenar_evals(lista):
    """Orden UCSUR fijo: ED, EC1, EP, EC2, EC3, EF, FINAL."""
    orden_base = EVALS + [FINAL_NAME]
    return sorted(list(dict.fromkeys(lista)), key=lambda x: orden_base.index(x))

# ==========================
# UTILIDADES BÁSICAS
# ==========================
def safe_txt(s: str) -> str:
    if s is None:
        return ""
    s = str(s).replace("—", "-").replace("\u2013", "-")
    return s

def slug(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s2 = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in s2.strip())
    s2 = s2.replace("__", "_")
    return s2 or "reporte"

def ensure_columns(df: pd.DataFrame):
    for col in ["Curso", "Seccion", "Evaluacion", "Nota", "CodigoAlumno",
                "Alumno", "Carrera", "Docente"]:
        if col not in df.columns:
            df[col] = 0.0 if col == "Nota" else ""
    return df

def round_half_up(x):
    """Redondeo al entero más cercano y .5 hacia arriba (x.5 -> x+1)."""
    try:
        x = float(x)
    except Exception:
        return 0
    return int(np.floor(x + 0.5)) if x >= 0 else int(np.ceil(x - 0.5))

def consolidar_por_alumno(df_sub: pd.DataFrame) -> pd.DataFrame:
    """
    BLINDAJE:
    Deja 1 fila por alumno (y por Seccion/Docente/Curso/Carrera), usando Nota = max.
    Se usa para evitar inflar 'Total' si el parquet trae duplicados por alumno.
    """
    if df_sub.empty:
        return df_sub.copy()
    keys = ["CodigoAlumno", "Alumno", "Curso", "Seccion", "Carrera", "Docente"]
    df2 = df_sub.copy()
    for k in keys:
        if k not in df2.columns:
            df2[k] = ""
    if "Nota" not in df2.columns:
        df2["Nota"] = 0.0
    df2["Nota"] = pd.to_numeric(df2["Nota"], errors="coerce").fillna(0.0)
    df2 = df2.groupby(keys, dropna=False, as_index=False)["Nota"].max()
    return df2

# ==========================
# SELECCIÓN DE CARPETA
# ==========================
def elegir_carpeta() -> str:
    if tk is None or filedialog is None:
        print("⚠️ No se pudo cargar tkinter; se usará la carpeta actual.")
        return os.getcwd()
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        carpeta = filedialog.askdirectory(title="Seleccione la carpeta de destino")
        root.destroy()
        if not carpeta:
            print("⚠️ No se seleccionó carpeta.")
            return None
        return carpeta
    except Exception as e:
        print(f"⚠️ Error al abrir el selector de carpeta: {e}")
        return None

# ==========================
# FUENTE Y ENCABEZADO PDF
# ==========================
def cargar_fuente(pdf: FPDF):
    try:
        arial_path = r"C:\Windows\Fonts\arial.ttf"
        arial_bold = r"C:\Windows\Fonts\arialbd.ttf"
        if os.path.exists(arial_path):
            pdf.add_font("ArialUnicode", "", arial_path)
        if os.path.exists(arial_bold):
            pdf.add_font("ArialUnicode", "B", arial_bold)
        return "ArialUnicode"
    except Exception:
        return "Helvetica"

def pdf_encabezado(pdf: FPDF, titulo: str):
    try:
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=10, y=8, w=28)
    except Exception as e:
        print(f"⚠️ No se pudo insertar el logo: {e}")

    pdf.set_xy(10, 12)
    pdf.set_font(pdf.font_family, "B", 14)
    pdf.cell(0, 8, "UNIVERSIDAD CIENTÍFICA DEL SUR", align="C", ln=1)

    pdf.set_x(10)
    pdf.set_font(pdf.font_family, "", 11)
    pdf.cell(0, 6, "Departamento de Cursos Básicos", align="C", ln=1)

    pdf.set_x(10)
    pdf.cell(0, 6, safe_txt(titulo), align="C", ln=1)

    pdf.set_y(42)

# ==========================
# CARGA DE DATOS
# ==========================
def cargar_df(parquet_path: str) -> pd.DataFrame:
    if not os.path.exists(parquet_path):
        print(f"❌ No se encuentra el parquet: {parquet_path}")
        sys.exit(1)
    df = pd.read_parquet(parquet_path)
    df = ensure_columns(df)
    for c in ["Curso", "Seccion", "Carrera", "Docente", "Evaluacion"]:
        df[c] = df[c].astype(str).str.strip()
    df["Evaluacion"] = df["Evaluacion"].str.upper()
    df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce").fillna(0.0)
    return df

# ==========================
# CÁLCULOS BASE
# ==========================
def pivot_final_por_estudiante(df_curso: pd.DataFrame) -> pd.DataFrame:
    piv = df_curso.pivot_table(
        index=["CodigoAlumno", "Alumno", "Curso", "Seccion", "Carrera", "Docente"],
        columns="Evaluacion", values="Nota", aggfunc="max", fill_value=0
    ).reset_index()

    for ev in EVALS:
        if ev not in piv.columns:
            piv[ev] = 0.0

    # FINAL con redondeo half-up solicitado
    for ev in ["EC1", "EP", "EC2", "EC3", "EF"]:
        piv[f"__R_{ev}"] = piv[ev].map(round_half_up)

    piv[FINAL_NAME] = (
        FINAL_WEIGHTS["EC1"] * piv["__R_EC1"] +
        FINAL_WEIGHTS["EP"]  * piv["__R_EP"]  +
        FINAL_WEIGHTS["EC2"] * piv["__R_EC2"] +
        FINAL_WEIGHTS["EC3"] * piv["__R_EC3"] +
        FINAL_WEIGHTS["EF"]  * piv["__R_EF"]
    )

    piv = piv.drop(columns=[c for c in piv.columns if c.startswith("__R_")])

    def estado_row(r):
        if all(float(r[ev]) == 0.0 for ev in EVALS):
            return "No rindió (todas 0)"
        return "Aprobado" if float(r[FINAL_NAME]) >= APROBADO_MIN else "Desaprobado"

    piv["EstadoFinal"] = piv.apply(estado_row, axis=1)
    return piv

def _agg_por_grupo(base: pd.DataFrame, group_cols: list, nota_col: str):
    """
    Retorna:
      - cant: conteos (Total, NoRindieron, Rindieron, Aprobados, Desaprobados, Promedio, Desv.Std)
      - pct : mismo orden pero con % calculados (Aprob/Desap sobre Rindieron; No rind sobre Total)
    """
    base = base.copy()

    # Variables auxiliares
    base["Rindio"] = (base[nota_col] > 0).astype(int)
    base["Ausente"] = (base[nota_col] == 0).astype(int)

    umbral = APROBADO_MIN
    base["Aprob"] = ((base[nota_col] >= umbral) & (base["Rindio"] == 1)).astype(int)
    base["Desap"] = ((base[nota_col] > 0) & (base[nota_col] < umbral)).astype(int)

    g = base.groupby(group_cols, dropna=False)

    cant = g.agg(
        Total=("CodigoAlumno", "count"),
        Aprobados=("Aprob", "sum"),
        Desaprobados=("Desap", "sum"),
        NoRindieron=("Ausente", "sum"),
    ).reset_index()

    cant["Rindieron"] = cant["Total"] - cant["NoRindieron"]

    prom = g.apply(
        lambda d: d.loc[d["Rindio"] == 1, nota_col].mean()
        if (d["Rindio"] == 1).any()
        else np.nan
    )
    stdv = g.apply(
        lambda d: d.loc[d["Rindio"] == 1, nota_col].std(ddof=1)
        if (d["Rindio"] == 1).sum() > 1
        else np.nan
    )

    cant["Promedio"] = np.round(prom.values, 2)
    cant["Desv.Std"] = np.round(stdv.values, 2)

    pct = cant.copy()

    # % Aprobados / % Desaprobados sobre Rindieron
    pct["% Aprobados"] = np.round(
        np.where(pct["Rindieron"] > 0, pct["Aprobados"] / pct["Rindieron"] * 100, 0), 1
    )
    pct["% Desaprobados"] = np.round(
        np.where(pct["Rindieron"] > 0, pct["Desaprobados"] / pct["Rindieron"] * 100, 0), 1
    )
    # % No rindieron sobre Total
    pct["% No rindieron"] = np.round(
        np.where(pct["Total"] > 0, pct["NoRindieron"] / pct["Total"] * 100, 0), 1
    )

    pct = pct.sort_values(
        by=["% Aprobados", "Promedio"], ascending=[False, False]
    ).reset_index(drop=True)
    cant = pct[cant.columns]
    return cant, pct

def resumen_por_grupo(df: pd.DataFrame, group_cols: list, usar_final=False):
    base = df.copy()
    nota_col = FINAL_NAME if usar_final else "Nota"
    return _agg_por_grupo(base, group_cols, nota_col)

def tabla_unica_desde_cant_pct(cant: pd.DataFrame, pct: pd.DataFrame, group_cols: list) -> pd.DataFrame:
    merge_cols = group_cols.copy()
    m = cant.merge(
        pct[merge_cols + ["% Aprobados", "% Desaprobados", "% No rindieron"]],
        on=merge_cols,
        how="left",
    )
    m = m.sort_values(
        by=["% Aprobados", "Promedio"], ascending=[False, False]
    ).reset_index(drop=True)
    return m

def resumen_por_seccion_desde_cant(cant: pd.DataFrame) -> pd.DataFrame:
    g = cant.groupby("Seccion", dropna=False).agg(
        Total=("Total", "sum"),
        Aprobados=("Aprobados", "sum"),
        Desaprobados=("Desaprobados", "sum"),
        NoRindieron=("NoRindieron", "sum"),
        Promedio=("Promedio", "mean"),
    ).reset_index()

    g["Rindieron"] = g["Total"] - g["NoRindieron"]

    g["% Aprobados"] = np.round(
        np.where(g["Rindieron"] > 0, g["Aprobados"] / g["Rindieron"] * 100, 0), 1
    )
    g["% Desaprobados"] = np.round(
        np.where(g["Rindieron"] > 0, g["Desaprobados"] / g["Rindieron"] * 100, 0), 1
    )
    g["% No rindieron"] = np.round(
        np.where(g["Total"] > 0, g["NoRindieron"] / g["Total"] * 100, 0), 1
    )

    g = g.sort_values(
        by=["% Aprobados", "Promedio"], ascending=[False, False]
    ).reset_index(drop=True)
    return g

def resumen_por_docente_desde_cant(cant: pd.DataFrame, base: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla por docente:
    Docente | N° secciones | Total | No rindieron | Rindieron | % Aprob. | % Desaprob. | % No rind. | Prom.
    """
    tabla = cant.copy()
    secciones = base.groupby("Docente")["Seccion"].nunique().reset_index(name="Número de secciones")
    tabla = tabla.merge(secciones, on="Docente", how="left")
    tabla.rename(columns={"Número de secciones": "Número.de.secciones"}, inplace=True)

    tabla["Rindieron"] = tabla["Total"] - tabla["NoRindieron"]

    tabla["% Aprobados"] = np.round(
        np.where(tabla["Rindieron"] > 0, tabla["Aprobados"] / tabla["Rindieron"] * 100, 0), 1
    )
    tabla["% Desaprobados"] = np.round(
        np.where(tabla["Rindieron"] > 0, tabla["Desaprobados"] / tabla["Rindieron"] * 100, 0), 1
    )
    tabla["% No rindieron"] = np.round(
        np.where(tabla["Total"] > 0, tabla["NoRindieron"] / tabla["Total"] * 100, 0), 1
    )

    tabla = tabla.sort_values(
        by=["% Aprobados", "Promedio"], ascending=[False, False]
    ).reset_index(drop=True)
    return tabla

def hay_notas_validas(cant: pd.DataFrame) -> bool:
    """False si no hay registros o si todos son 'No rindieron'."""
    if cant.empty:
        return False
    total_sum = cant["Total"].sum()
    no_r_sum = cant["NoRindieron"].sum()
    if total_sum == 0 or total_sum == no_r_sum:
        return False
    return True

def ef_valida_en_df(df_scope: pd.DataFrame) -> bool:
    sub = df_scope[df_scope["Evaluacion"] == "EF"]
    if sub.empty:
        return False
    return (sub["Nota"] > 0).any()

# ==========================
# GRÁFICOS
# ==========================
def grafica_barras(df: pd.DataFrame, label_col: str, val_col: str,
                   title: str, ylabel: str, outfile: str):
    if df.empty:
        return
    n = len(df)
    width = 0.45 if n > 14 else 0.6
    xtick_fs = 7 if n > 18 else (8 if n > 12 else 9)
    rot = 60 if n > 12 else 30
    height = 2.4 + min(0.12 * n, 3.0)
    plt.figure(figsize=(7.2, height))
    plt.bar(df[label_col].astype(str), df[val_col], width=width)
    plt.xticks(rotation=rot, ha="right", fontsize=xtick_fs)
    plt.ylabel(ylabel)
    plt.title(title, fontsize=10)
    plt.tight_layout()
    plt.savefig(outfile, dpi=150)
    plt.close()

# ==========================
# PDF: TABLAS
# ==========================
def pdf_tabla(pdf: FPDF, familia: str, headers, data_rows, widths):
    def _header():
        pdf.set_fill_color(0, 59, 112)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(familia, "B", 9)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, safe_txt(h), border=1, align="C", fill=True)
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.set_font(familia, "", 9)

    _header()
    for row in data_rows:
        if pdf.get_y() > 260:
            pdf.add_page()
            _header()
        for val, w in zip(row, widths):
            pdf.cell(w, 6, safe_txt(str(val)), border=1, align="C")
        pdf.ln()

# ==========================
# EXPORTAR PDF — CASO 1 (Curso + evaluaciones, por sección)
# ==========================
def exportar_pdf_curso_evals(df_curso: pd.DataFrame, curso: str, evals_sel, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf_evals = slug("_".join(evals_sel))[:20]
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    pdf_path = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Secciones_{suf_evals}_{ts}.pdf")

    pdf = FPDF()
    familia = None
    tmpdir = tempfile.mkdtemp()

    for ev in evals_sel:
        if ev == FINAL_NAME:
            piv = pivot_final_por_estudiante(df_curso)
            cant, pct = resumen_por_grupo(piv, ["Seccion", "Docente"], usar_final=True)
        else:
            sub = df_curso[df_curso["Evaluacion"] == ev].copy()
            if sub.empty:
                print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                continue

            # BLINDAJE POR ALUMNO
            sub = consolidar_por_alumno(sub)

            cant, pct = resumen_por_grupo(sub, ["Seccion", "Docente"], usar_final=False)

        if not hay_notas_validas(cant):
            print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
            continue

        tabla = tabla_unica_desde_cant_pct(cant, pct, ["Seccion", "Docente"])
        resumen_sec = resumen_por_seccion_desde_cant(cant)

        g_pct = os.path.join(tmpdir, f"bar_pct_{ev}.png")
        g_prom = os.path.join(tmpdir, f"bar_prom_{ev}.png")
        grafica_barras(resumen_sec, "Seccion", "% Aprobados",
                       f"{nombre_eval(ev)} — % Aprobados por sección (sobre rindieron)",
                       "% Aprobados", g_pct)
        grafica_barras(resumen_sec, "Seccion", "Promedio",
                       f"{nombre_eval(ev)} — Promedio por sección",
                       "Promedio", g_prom)

        pdf.add_page()
        if familia is None:
            familia = cargar_fuente(pdf)
            pdf.set_font(familia, "", 10)
        pdf_encabezado(pdf, f"Análisis por Secciones — {safe_txt(curso)} — {nombre_eval(ev)}")

        pdf.set_font(familia, "", 10)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Curso: {safe_txt(curso)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, nombre_eval(ev), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y')}",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)

        pdf.set_font(familia, "B", 11)
        pdf.cell(0, 7, "Resumen por Sección y Docente", ln=1)

        headers = ["Sección", "Docente", "Total", "No rind.", "Rind.",
                   "% Aprob.", "% Desaprob.", "% No rind.", "Prom."]
        widths = [18, 56, 12, 14, 12, 16, 18, 16, 14]

        rows = []
        for _, r in tabla.iterrows():
            rows.append([
                safe_txt(r["Seccion"]),
                safe_txt(r["Docente"]),
                int(r["Total"]),
                int(r["NoRindieron"]),
                int(r["Rindieron"]),
                f"{r['% Aprobados']}%",
                f"{r['% Desaprobados']}%",
                f"{r['% No rindieron']}%",
                r["Promedio"],
            ])
        pdf_tabla(pdf, familia, headers, rows, widths)

        pdf.ln(3)
        pdf.set_font(familia, "B", 10)
        pdf.cell(0, 6, "Gráfico: % Aprobados por sección (sobre rindieron)", ln=1)
        pdf.image(g_pct, w=170)
        pdf.ln(3)
        pdf.cell(0, 6, "Gráfico: Promedio por sección", ln=1)
        pdf.image(g_prom, w=170)

    try:
        pdf.output(pdf_path)
        print(f"✅ PDF generado: {os.path.abspath(pdf_path)}")
        if os.name == "nt":
            os.startfile(pdf_path)
    except Exception as e:
        print(f"⚠️ Error al guardar el PDF: {e}")

# ==========================
# EXPORTAR PDF — CASO 2 (Todos los docentes del curso)
# ==========================
def exportar_pdf_todos_docentes(df_curso: pd.DataFrame, curso: str, evals_sel, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf_evals = slug("_".join(evals_sel))[:20]
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    pdf_path = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Docentes_{suf_evals}_{ts}.pdf")

    pdf = FPDF()
    familia = None
    tmpdir = tempfile.mkdtemp()

    for ev in evals_sel:
        if ev == FINAL_NAME:
            piv = pivot_final_por_estudiante(df_curso)
            cant, pct = resumen_por_grupo(piv, ["Docente"], usar_final=True)
            base = piv
        else:
            sub = df_curso[df_curso["Evaluacion"] == ev].copy()
            if sub.empty:
                print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                continue

            # BLINDAJE POR ALUMNO
            sub = consolidar_por_alumno(sub)

            cant, pct = resumen_por_grupo(sub, ["Docente"], usar_final=False)
            base = sub

        if not hay_notas_validas(cant):
            print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
            continue

        tabla_doc = resumen_por_docente_desde_cant(cant, base)

        g_pct = os.path.join(tmpdir, f"bar_pct_doc_{ev}.png")
        g_prom = os.path.join(tmpdir, f"bar_prom_doc_{ev}.png")
        grafica_barras(tabla_doc, "Docente", "% Aprobados",
                       f"{nombre_eval(ev)} — % Aprobados por docente (sobre rindieron)",
                       "% Aprobados", g_pct)
        grafica_barras(tabla_doc, "Docente", "Promedio",
                       f"{nombre_eval(ev)} — Promedio por docente",
                       "Promedio", g_prom)

        pdf.add_page()
        if familia is None:
            familia = cargar_fuente(pdf)
            pdf.set_font(familia, "", 10)
        pdf_encabezado(pdf, f"Análisis por Docentes — {safe_txt(curso)} — {nombre_eval(ev)}")

        pdf.set_font(familia, "", 10)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Curso: {safe_txt(curso)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, nombre_eval(ev), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y')}",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)

        pdf.set_font(familia, "B", 11)
        pdf.cell(0, 7, "Resumen por Docente", ln=1)

        headers = ["Docente", "N° Sec.", "Total", "No rind.", "Rind.",
                   "% Aprob.", "% Desaprob.", "% No rind.", "Prom."]
        widths = [54, 14, 12, 14, 12, 16, 18, 16, 14]

        rows = []
        for _, r in tabla_doc.iterrows():
            rows.append([
                safe_txt(r["Docente"]),
                int(r["Número.de.secciones"]) if not pd.isna(r["Número.de.secciones"]) else "",
                int(r["Total"]),
                int(r["NoRindieron"]),
                int(r["Rindieron"]),
                f"{r['% Aprobados']}%",
                f"{r['% Desaprobados']}%",
                f"{r['% No rindieron']}%",
                r["Promedio"],
            ])
        pdf_tabla(pdf, familia, headers, rows, widths)

        pdf.ln(3)
        pdf.set_font(familia, "B", 10)
        pdf.cell(0, 6, "Gráfico: % Aprobados por docente (sobre rindieron)", ln=1)
        pdf.image(g_pct, w=170)
        pdf.ln(3)
        pdf.cell(0, 6, "Gráfico: Promedio por docente", ln=1)
        pdf.image(g_prom, w=170)

    try:
        pdf.output(pdf_path)
        print(f"✅ PDF generado: {os.path.abspath(pdf_path)}")
        if os.name == "nt":
            os.startfile(pdf_path)
    except Exception as e:
        print(f"⚠️ Error al guardar el PDF: {e}")

# ==========================
# EXPORTAR PDF — CASO 3 (Docente específico)
# ==========================
def exportar_pdf_docente(df_curso: pd.DataFrame, curso: str, docente: str, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    pdf_path = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Docente_{slug(docente)}_{ts}.pdf")

    pdf = FPDF()
    familia = None
    tmpdir = tempfile.mkdtemp()

    df_doc = df_curso[df_curso["Docente"] == docente].copy()
    if df_doc.empty:
        print("⚠️ No hay registros para ese docente en el curso seleccionado.")
        return

    ef_ok = ef_valida_en_df(df_doc)
    eval_loop = EVALS + ([FINAL_NAME] if ef_ok else [])

    # 1) Global por evaluación (con Total / No rind / Rind y % sobre rind)
    filas_global = []
    for ev in eval_loop:
        if ev == FINAL_NAME:
            piv = pivot_final_por_estudiante(df_doc)
            if piv.empty:
                continue

            total = int(len(piv))
            no_r = int((piv["EstadoFinal"] == "No rindió (todas 0)").sum())
            rindieron = int(total - no_r)
            if total == 0 or rindieron == 0:
                continue

            rind = piv[piv["EstadoFinal"] != "No rindió (todas 0)"][FINAL_NAME]
            aprob = int((rind >= APROBADO_MIN).sum())
            desap = int((rind < APROBADO_MIN).sum())
            prom = float(rind.mean()) if rindieron else np.nan

            pct_ap = round(aprob / rindieron * 100, 1)
            pct_de = round(desap / rindieron * 100, 1)
            pct_nr = round(no_r / total * 100, 1)

        else:
            sub = df_doc[df_doc["Evaluacion"] == ev].copy()
            if sub.empty:
                continue

            # BLINDAJE POR ALUMNO
            sub = consolidar_por_alumno(sub)

            total = int(len(sub))
            no_r = int((sub["Nota"] == 0).sum())
            rind = sub[sub["Nota"] > 0]["Nota"]
            rindieron = int(len(rind))
            if total == 0 or rindieron == 0:
                continue

            aprob = int((rind >= APROBADO_MIN).sum())
            desap = int((rind < APROBADO_MIN).sum())
            prom = float(rind.mean()) if rindieron else np.nan

            pct_ap = round(aprob / rindieron * 100, 1)
            pct_de = round(desap / rindieron * 100, 1)
            pct_nr = round(no_r / total * 100, 1)

        filas_global.append({
            "CodigoEval": ev,
            "Evaluación": nombre_eval(ev),
            "Total": total,
            "No rindieron": no_r,
            "Rindieron": rindieron,
            "% Aprobados": pct_ap,
            "% Desaprobados": pct_de,
            "% No rindieron": pct_nr,
            "Promedio": round(prom, 2) if not np.isnan(prom) else np.nan,
        })

    if not filas_global:
        print("⚠️ Aún no hay registros válidos para este docente.")
        return

    df_global = pd.DataFrame(filas_global)
    df_global["__ord"] = df_global["CodigoEval"].apply(lambda x: (EVALS + [FINAL_NAME]).index(x))
    df_global = df_global.sort_values("__ord").drop(columns=["__ord"]).reset_index(drop=True)

    g_pct_eval = os.path.join(tmpdir, "bar_pct_eval_global.png")
    grafica_barras(df_global, "Evaluación", "% Aprobados",
                   "Docente — % Aprobados por evaluación (sobre rindieron)",
                   "% Aprobados", g_pct_eval)

    pdf.add_page()
    if familia is None:
        familia = cargar_fuente(pdf)
        pdf.set_font(familia, "", 10)
    pdf_encabezado(pdf, f"Análisis Global por Evaluación — {safe_txt(docente)} — {safe_txt(curso)}")

    pdf.set_font(familia, "", 10)
    pdf.set_x(10)
    pdf.cell(0, 6, f"Curso: {safe_txt(curso)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(10)
    pdf.cell(0, 6, f"Docente: {safe_txt(docente)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(10)
    pdf.cell(0, 6, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y')}",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(3)

    pdf.set_font(familia, "B", 11)
    pdf.cell(0, 7, "Resumen Global por Evaluación", ln=1)

    headers = ["Evaluación", "Total", "No rind.", "Rind.",
               "% Aprob.", "% Desaprob.", "% No rind.", "Prom."]
    widths = [40, 12, 14, 12, 16, 18, 16, 14]

    rows = []
    for _, r in df_global.iterrows():
        rows.append([
            r["Evaluación"],
            int(r["Total"]),
            int(r["No rindieron"]),
            int(r["Rindieron"]),
            f"{r['% Aprobados']}%",
            f"{r['% Desaprobados']}%",
            f"{r['% No rindieron']}%",
            r["Promedio"],
        ])
    pdf_tabla(pdf, familia, headers, rows, widths)

    pdf.ln(3)
    pdf.set_font(familia, "B", 10)
    pdf.cell(0, 6, "Gráfico comparativo: % Aprobados por evaluación (sobre rindieron)", ln=1)
    pdf.image(g_pct_eval, w=170)

    # 2) Por evaluación y sección
    for ev in eval_loop:
        if ev == FINAL_NAME:
            piv = pivot_final_por_estudiante(df_doc)
            if piv.empty:
                continue
            cant, pct = resumen_por_grupo(piv, ["Seccion"], usar_final=True)
        else:
            sub = df_doc[df_doc["Evaluacion"] == ev].copy()
            if sub.empty:
                continue

            # BLINDAJE POR ALUMNO
            sub = consolidar_por_alumno(sub)

            cant, pct = resumen_por_grupo(sub, ["Seccion"], usar_final=False)

        if not hay_notas_validas(cant):
            continue

        tabla_sec = tabla_unica_desde_cant_pct(cant, pct, ["Seccion"])

        g_pct_sec = os.path.join(tmpdir, f"bar_pct_sec_{ev}.png")
        grafica_barras(tabla_sec, "Seccion", "% Aprobados",
                       f"{nombre_eval(ev)} — % Aprobados por sección (sobre rindieron)",
                       "% Aprobados", g_pct_sec)

        pdf.add_page()
        pdf_encabezado(pdf, f"Docente {safe_txt(docente)} — {nombre_eval(ev)} — por Sección")

        pdf.set_font(familia, "", 10)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Curso: {safe_txt(curso)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, f"Docente: {safe_txt(docente)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(10)
        pdf.cell(0, 6, nombre_eval(ev), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(3)

        pdf.set_font(familia, "B", 11)
        pdf.cell(0, 7, "Resumen por Sección", ln=1)

        headers2 = ["Sección", "Total", "No rind.", "Rind.",
                    "% Aprob.", "% Desaprob.", "% No rind.", "Prom."]
        widths2 = [22, 12, 14, 12, 16, 18, 16, 14]

        rows2 = []
        for _, r in tabla_sec.iterrows():
            rows2.append([
                safe_txt(r["Seccion"]),
                int(r["Total"]),
                int(r["NoRindieron"]),
                int(r["Rindieron"]),
                f"{r['% Aprobados']}%",
                f"{r['% Desaprobados']}%",
                f"{r['% No rindieron']}%",
                r["Promedio"],
            ])
        pdf_tabla(pdf, familia, headers2, rows2, widths2)

        pdf.ln(3)
        pdf.set_font(familia, "B", 10)
        pdf.cell(0, 6, "Gráfico: % Aprobados por sección (sobre rindieron)", ln=1)
        pdf.image(g_pct_sec, w=170)

    try:
        pdf.output(pdf_path)
        print(f"✅ PDF generado: {os.path.abspath(pdf_path)}")
        if os.name == "nt":
            os.startfile(pdf_path)
    except Exception as e:
        print(f"⚠️ Error al guardar el PDF: {e}")
# =========================
# PARTE 2/2
# =========================
# (Continuación) Analisis_Secciones_Final.py
# AJUSTES EXCEL: aplicar BLINDAJE POR ALUMNO en CASO 1, CASO 2 y CASO 3
# sin cambiar estructura, formatos ni análisis.

# ==========================
# EXCEL — UTILIDADES
# ==========================
def autosize_columns(ws, df, start_row=0, start_col=0, min_w=10, max_w=38):
    for j, col in enumerate(df.columns):
        col_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).values.tolist()])
        ws.set_column(start_col + j, start_col + j, max(min_w, min(col_len + 2, max_w)))

def write_df_table(ws, df, start_row, start_col, fmt_header, fmt_cell):
    for c, name in enumerate(df.columns):
        ws.write(start_row, start_col + c, name, fmt_header)
    for r in range(len(df)):
        for c in range(len(df.columns)):
            ws.write(start_row + 1 + r, start_col + c, df.iat[r, c], fmt_cell)
    return len(df) + 1, len(df.columns)

def limpiar_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.replace([np.nan, np.inf, -np.inf], "")

# ==========================
# EXCEL — CASO 1 (Secciones)
# ==========================
def exportar_excel_curso_evals(df_curso: pd.DataFrame, curso: str, evals_sel, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf_evals = slug("_".join(evals_sel))[:20]
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    xlsx = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Secciones_{suf_evals}_{ts}.xlsx")

    with pd.ExcelWriter(xlsx, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True, "font_color": "white",
            "bg_color": COL_UCSUR_AZUL, "border": 1, "align": "center"
        })
        fmt_cell = wb.add_format({"border": 1})
        fmt_title = wb.add_format({"bold": True, "font_size": 14})

        for ev in evals_sel:
            if ev == FINAL_NAME:
                piv = pivot_final_por_estudiante(df_curso)
                cant, pct = resumen_por_grupo(piv, ["Seccion", "Docente"], usar_final=True)
            else:
                sub = df_curso[df_curso["Evaluacion"] == ev].copy()
                if sub.empty:
                    print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                    continue

                # BLINDAJE POR ALUMNO
                sub = consolidar_por_alumno(sub)

                cant, pct = resumen_por_grupo(sub, ["Seccion", "Docente"], usar_final=False)

            if not hay_notas_validas(cant):
                print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                continue

            tabla = tabla_unica_desde_cant_pct(cant, pct, ["Seccion", "Docente"])

            sheet_name = f"EV_{ev}"[:31]
            ws = wb.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws

            ws.write(0, 0, f"{curso} — {nombre_eval(ev)} — Secciones", fmt_title)

            df_excel = tabla[[
                "Seccion", "Docente", "Total", "NoRindieron", "Rindieron",
                "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"
            ]].copy()

            df_excel.columns = [
                "Sección", "Docente", "Total", "No rind.", "Rind.",
                "% Aprob.", "% Desaprob.", "% No rind.", "Prom."
            ]
            df_excel = limpiar_excel(df_excel)

            nr, nc = write_df_table(ws, df_excel, 2, 0, fmt_header, fmt_cell)
            autosize_columns(ws, df_excel, 2, 0)

            # Gráfico: % Aprobados (sobre rindieron)
            ch = wb.add_chart({"type": "column"})
            ch.add_series({
                "name": "% Aprobados (sobre rindieron)",
                "categories": [sheet_name, 3, 0, 2 + len(df_excel), 0],
                "values":     [sheet_name, 3, 5, 2 + len(df_excel), 5],
            })
            ch.set_title({"name": f"{nombre_eval(ev)} — % Aprobados (sobre rindieron)"})
            ws.insert_chart(2, nc + 2, ch)

    print(f"✅ Excel generado: {os.path.abspath(xlsx)}")

# ==========================
# EXCEL — CASO 2 (Docentes)
# ==========================
def exportar_excel_todos_docentes(df_curso: pd.DataFrame, curso: str, evals_sel, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf_evals = slug("_".join(evals_sel))[:20]
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    xlsx = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Docentes_{suf_evals}_{ts}.xlsx")

    with pd.ExcelWriter(xlsx, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True, "font_color": "white",
            "bg_color": COL_UCSUR_AZUL, "border": 1, "align": "center"
        })
        fmt_cell = wb.add_format({"border": 1})
        fmt_title = wb.add_format({"bold": True, "font_size": 14})

        for ev in evals_sel:
            if ev == FINAL_NAME:
                piv = pivot_final_por_estudiante(df_curso)
                cant, pct = resumen_por_grupo(piv, ["Docente"], usar_final=True)
                base = piv
            else:
                sub = df_curso[df_curso["Evaluacion"] == ev].copy()
                if sub.empty:
                    print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                    continue

                # BLINDAJE POR ALUMNO
                sub = consolidar_por_alumno(sub)

                cant, pct = resumen_por_grupo(sub, ["Docente"], usar_final=False)
                base = sub

            if not hay_notas_validas(cant):
                print(f"⚠️ Aún no hay registros de estas notas para la evaluación {nombre_eval(ev)}.")
                continue

            tabla_doc = resumen_por_docente_desde_cant(cant, base)

            df_excel = tabla_doc[[
                "Docente", "Número.de.secciones", "Total", "NoRindieron", "Rindieron",
                "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"
            ]].copy()

            df_excel.columns = [
                "Docente", "N° Sec.", "Total", "No rind.", "Rind.",
                "% Aprob.", "% Desaprob.", "% No rind.", "Prom."
            ]
            df_excel = limpiar_excel(df_excel)

            sheet_name = f"EV_{ev}"[:31]
            ws = wb.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws

            ws.write(0, 0, f"{curso} — {nombre_eval(ev)} — Docentes", fmt_title)
            nr, nc = write_df_table(ws, df_excel, 2, 0, fmt_header, fmt_cell)
            autosize_columns(ws, df_excel, 2, 0)

            ch = wb.add_chart({"type": "column"})
            ch.add_series({
                "name": "% Aprobados (sobre rindieron)",
                "categories": [sheet_name, 3, 0, 2 + len(df_excel), 0],
                "values":     [sheet_name, 3, 5, 2 + len(df_excel), 5],
            })
            ch.set_title({"name": f"{nombre_eval(ev)} — % Aprobados (sobre rindieron)"})
            ws.insert_chart(2, nc + 2, ch)

    print(f"✅ Excel generado: {os.path.abspath(xlsx)}")

# ==========================
# EXCEL — CASO 3 (Docente específico)
# ==========================
def exportar_excel_docente(df_curso: pd.DataFrame, curso: str, docente: str, carpeta_salida: str):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not carpeta_salida:
        carpeta_salida = os.getcwd()
    xlsx = os.path.join(carpeta_salida, f"Curso_{slug(curso)}_Docente_{slug(docente)}_{ts}.xlsx")

    df_doc = df_curso[df_curso["Docente"] == docente].copy()
    if df_doc.empty:
        print("⚠️ No hay datos para ese docente.")
        return

    ef_ok = ef_valida_en_df(df_doc)
    eval_loop = EVALS + ([FINAL_NAME] if ef_ok else [])

    with pd.ExcelWriter(xlsx, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True, "font_color": "white",
            "bg_color": COL_UCSUR_AZUL, "border": 1, "align": "center"
        })
        fmt_cell = wb.add_format({"border": 1})
        fmt_title = wb.add_format({"bold": True, "font_size": 14})

        # Hoja global por evaluación
        filas_global = []
        for ev in eval_loop:
            if ev == FINAL_NAME:
                piv = pivot_final_por_estudiante(df_doc)
                if piv.empty:
                    continue

                total = int(len(piv))
                no_r = int((piv["EstadoFinal"] == "No rindió (todas 0)").sum())
                rindieron = int(total - no_r)
                if total == 0 or rindieron == 0:
                    continue

                rind = piv[piv["EstadoFinal"] != "No rindió (todas 0)"][FINAL_NAME]
                aprob = int((rind >= APROBADO_MIN).sum())
                desap = int((rind < APROBADO_MIN).sum())
                prom = float(rind.mean()) if rindieron else np.nan

                pct_ap = round(aprob / rindieron * 100, 1)
                pct_de = round(desap / rindieron * 100, 1)
                pct_nr = round(no_r / total * 100, 1)

            else:
                sub = df_doc[df_doc["Evaluacion"] == ev].copy()
                if sub.empty:
                    continue

                # BLINDAJE POR ALUMNO
                sub = consolidar_por_alumno(sub)

                total = int(len(sub))
                no_r = int((sub["Nota"] == 0).sum())
                rind = sub[sub["Nota"] > 0]["Nota"]
                rindieron = int(len(rind))
                if total == 0 or rindieron == 0:
                    continue

                aprob = int((rind >= APROBADO_MIN).sum())
                desap = int((rind < APROBADO_MIN).sum())
                prom = float(rind.mean()) if rindieron else np.nan

                pct_ap = round(aprob / rindieron * 100, 1)
                pct_de = round(desap / rindieron * 100, 1)
                pct_nr = round(no_r / total * 100, 1)

            filas_global.append({
                "CodigoEval": ev,
                "Evaluación": nombre_eval(ev),
                "Total": total,
                "No rindieron": no_r,
                "Rindieron": rindieron,
                "% Aprobados": pct_ap,
                "% Desaprobados": pct_de,
                "% No rindieron": pct_nr,
                "Promedio": round(prom, 2) if not np.isnan(prom) else np.nan,
            })

        if filas_global:
            df_global = pd.DataFrame(filas_global)
            df_global["__ord"] = df_global["CodigoEval"].apply(lambda x: (EVALS + [FINAL_NAME]).index(x))
            df_global = df_global.sort_values("__ord").drop(columns=["__ord"]).reset_index(drop=True)

            sh = "Global"
            ws = wb.add_worksheet(sh)
            writer.sheets[sh] = ws
            ws.write(0, 0, f"{curso} — {docente} — Global por evaluación", fmt_title)

            df_excel = df_global[[
                "Evaluación", "Total", "No rindieron", "Rindieron",
                "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"
            ]].copy()
            df_excel.columns = [
                "Evaluación", "Total", "No rind.", "Rind.",
                "% Aprob.", "% Desaprob.", "% No rind.", "Prom."
            ]
            df_excel = limpiar_excel(df_excel)

            nr, nc = write_df_table(ws, df_excel, 2, 0, fmt_header, fmt_cell)
            autosize_columns(ws, df_excel, 2, 0)

        # Hojas por evaluación y sección
        for ev in eval_loop:
            if ev == FINAL_NAME:
                piv = pivot_final_por_estudiante(df_doc)
                if piv.empty:
                    continue
                cant, pct = resumen_por_grupo(piv, ["Seccion"], usar_final=True)
            else:
                sub = df_doc[df_doc["Evaluacion"] == ev].copy()
                if sub.empty:
                    continue

                # BLINDAJE POR ALUMNO
                sub = consolidar_por_alumno(sub)

                cant, pct = resumen_por_grupo(sub, ["Seccion"], usar_final=False)

            if not hay_notas_validas(cant):
                continue

            tabla_sec = tabla_unica_desde_cant_pct(cant, pct, ["Seccion"])

            df_excel = tabla_sec[[
                "Seccion", "Total", "NoRindieron", "Rindieron",
                "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"
            ]].copy()
            df_excel.columns = [
                "Sección", "Total", "No rind.", "Rind.",
                "% Aprob.", "% Desaprob.", "% No rind.", "Prom."
            ]
            df_excel = limpiar_excel(df_excel)

            sh = f"EV_{ev}"[:31]
            ws = wb.add_worksheet(sh)
            writer.sheets[sh] = ws
            ws.write(0, 0, f"{curso} — {docente} — {nombre_eval(ev)} por sección", fmt_title)

            nr, nc = write_df_table(ws, df_excel, 2, 0, fmt_header, fmt_cell)
            autosize_columns(ws, df_excel, 2, 0)

            ch = wb.add_chart({"type": "column"})
            ch.add_series({
                "name": "% Aprobados (sobre rindieron)",
                "categories": [sh, 3, 0, 2 + len(df_excel), 0],
                "values":     [sh, 3, 4, 2 + len(df_excel), 4],
            })
            ch.set_title({"name": f"{nombre_eval(ev)} — % Aprobados (sobre rindieron)"})
            ws.insert_chart(2, nc + 2, ch)

    print(f"✅ Excel generado: {os.path.abspath(xlsx)}")

# ==========================
# MENÚS
# ==========================
def menu_evaluaciones(df_scope: pd.DataFrame):
    """
    Devuelve lista de códigos de evaluación seleccionados (entre EVALS + FINAL_NAME),
    o None si el usuario decide volver.
    “Situación Final” solo aparece si hay EF válida.
    """
    ef_ok = ef_valida_en_df(df_scope)

    while True:
        print("\n==============================")
        print(" ELEGIR EVALUACIÓN(ES)")
        print("==============================")
        print("1. Evaluación Diagnóstica (ED)")
        print("2. Evaluación Continua 1 (EC1)")
        print("3. Evaluación Parcial (EP)")
        print("4. Evaluación Continua 2 (EC2)")
        print("5. Evaluación Continua 3 (EC3)")
        print("6. Evaluación Final (EF)")
        if ef_ok:
            print("7. Situación Final")
            print("8. TODAS")
            print("9. Volver")
        else:
            print("7. TODAS")
            print("8. Volver")

        sel = input("Seleccione una o varias opciones (ej. 1,3,4): ").strip()

        if not sel:
            print("⚠️ Debe seleccionar al menos una opción.")
            continue

        partes = [p.strip() for p in sel.split(",") if p.strip()]
        codigos = set(partes)

        if ef_ok:
            mapa = {
                "1": "ED",
                "2": "EC1",
                "3": "EP",
                "4": "EC2",
                "5": "EC3",
                "6": "EF",
                "7": FINAL_NAME,
                "8": "TODAS",
                "9": "VOLVER",
            }
            opt_todas = "8"
            opt_volver = "9"
        else:
            mapa = {
                "1": "ED",
                "2": "EC1",
                "3": "EP",
                "4": "EC2",
                "5": "EC3",
                "6": "EF",
                "7": "TODAS",
                "8": "VOLVER",
            }
            opt_todas = "7"
            opt_volver = "8"

        if opt_volver in codigos:
            return None

        if opt_todas in codigos:
            evals_sel = EVALS.copy()
            if ef_ok:
                evals_sel = EVALS + [FINAL_NAME]
            evals_sel = ordenar_evals(evals_sel)
            return evals_sel

        evals_sel = []
        for c in codigos:
            if c not in mapa:
                continue
            ev_code = mapa[c]
            if ev_code in EVALS or ev_code == FINAL_NAME:
                if ev_code == FINAL_NAME and not ef_ok:
                    continue
                evals_sel.append(ev_code)

        if not evals_sel:
            print("⚠️ Selección no válida. Intente nuevamente.")
            continue

        evals_sel = ordenar_evals(evals_sel)
        return evals_sel

def menu_exportar(func_pdf, func_excel):
    """
    Muestra menú:
    1. PDF → abre selector de carpeta → genera PDF
    2. Excel → abre selector de carpeta → genera Excel
    3. Volver
    """
    while True:
        print("\n───────────────")
        print("1. Exportar PDF")
        print("2. Exportar Excel")
        print("3. Volver")
        op = input("Seleccione una opción: ").strip()
        if op == "1":
            carpeta = elegir_carpeta()
            if not carpeta:
                continue
            func_pdf(carpeta)
        elif op == "2":
            carpeta = elegir_carpeta()
            if not carpeta:
                continue
            func_excel(carpeta)
        elif op == "3":
            break
        else:
            print("⚠️ Opción no válida.")

def submenu_curso(df: pd.DataFrame, curso: str):
    df_curso = df[df["Curso"].str.upper() == curso.upper()].copy()
    if df_curso.empty:
        print("⚠️ No hay datos para ese curso.")
        return

    while True:
        print(f"\n📘 Curso activo: {curso}")
        print("───────────────────────────────")
        print("1. Evaluaciones por sección")
        print("2. Todos los docentes")
        print("3. Elegir un docente")
        print("4. Volver")
        op = input("Seleccione una opción: ").strip()

        if op == "1":
            evals_sel = menu_evaluaciones(df_curso)
            if evals_sel is None:
                continue

            def _pdf(carpeta):
                exportar_pdf_curso_evals(df_curso, curso, evals_sel, carpeta)

            def _xlsx(carpeta):
                exportar_excel_curso_evals(df_curso, curso, evals_sel, carpeta)

            menu_exportar(_pdf, _xlsx)

        elif op == "2":
            evals_sel = menu_evaluaciones(df_curso)
            if evals_sel is None:
                continue

            def _pdf(carpeta):
                exportar_pdf_todos_docentes(df_curso, curso, evals_sel, carpeta)

            def _xlsx(carpeta):
                exportar_excel_todos_docentes(df_curso, curso, evals_sel, carpeta)

            menu_exportar(_pdf, _xlsx)

        elif op == "3":
            docentes = sorted(
                [d for d in df_curso["Docente"].unique().tolist() if str(d).strip()]
            )
            if not docentes:
                print("⚠️ No hay docentes registrados para este curso.")
                continue
            print("\nDocentes disponibles:")
            for i, d in enumerate(docentes, start=1):
                print(f"{i}. {d}")
            ix = input("Seleccione un docente por número: ").strip()
            if not ix.isdigit() or not (1 <= int(ix) <= len(docentes)):
                print("⚠️ Opción inválida.")
                continue
            docente = docentes[int(ix) - 1]

            def _pdf(carpeta):
                exportar_pdf_docente(df_curso, curso, docente, carpeta)

            def _xlsx(carpeta):
                exportar_excel_docente(df_curso, curso, docente, carpeta)

            menu_exportar(_pdf, _xlsx)

        elif op == "4":
            break
        else:
            print("⚠️ Opción no válida.")

def menu_principal():
    df = cargar_df(PARQUET_IN)
    cursos = sorted(
        [c for c in df["Curso"].str.upper().unique().tolist() if str(c).strip()]
    )
    while True:
        print("\n==============================")
        print(" MENÚ PRINCIPAL — ANÁLISIS UCSUR (Script 2)")
        print("==============================")
        print("1. Elegir curso")
        print("2. Salir")
        op = input("Seleccione una opción: ").strip()

        if op == "1":
            print("\nListado de cursos:")
            for i, c in enumerate(cursos, start=1):
                print(f"{i}. {c}")
            sel = input("Seleccione un curso por número: ").strip()
            if not sel.isdigit() or not (1 <= int(sel) <= len(cursos)):
                print("⚠️ Entrada inválida.")
                continue
            curso = cursos[int(sel) - 1]
            print(f"\n✅ Curso seleccionado: {curso}")
            submenu_curso(df, curso)

        elif op == "2":
            print("👋 Saliendo del sistema.")
            break
        else:
            print("⚠️ Opción no válida.")

# ==========================
# MAIN
# ==========================
if __name__ == "__main__":
    menu_principal()
