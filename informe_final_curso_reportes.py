# informe_final_curso_reportes.py
# -*- coding: utf-8 -*-

import os
import unicodedata
import tempfile
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from fpdf import FPDF
import xlsxwriter


# =========================
# CONFIG
# =========================
EVALS = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]
WEIGHTS = {"ED": 0.0, "EC1": 0.18, "EP": 0.20, "EC2": 0.18, "EC3": 0.19, "EF": 0.25}
FINAL_NAME = "SITUACIÓN FINAL"

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

EVAL_NAMES = {
    "ED": "Evaluación Diagnóstica",
    "EC1": "Evaluación Continua 1",
    "EP": "Evaluación Parcial",
    "EC2": "Evaluación Continua 2",
    "EC3": "Evaluación Continua 3",
    "EF": "Evaluación Final",
    FINAL_NAME: "Situación Final",
}

ORDER_EVAL = EVALS + [FINAL_NAME]


# =========================
# UTILIDADES
# =========================
def nombre_eval(ev):
    return EVAL_NAMES.get(ev, ev)


def slug(s):
    s2 = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s2 = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in s2.strip())
    while "__" in s2:
        s2 = s2.replace("__", "_")
    return s2 or "reporte"


def safe_txt(s):
    if s is None:
        return ""
    return str(s).replace("—", "-").replace("\u2013", "-")


def ensure_columns(df):
    cols = ["CodigoAlumno", "Alumno", "Curso", "Seccion", "Carrera", "Docente", "Evaluacion", "Nota"]
    for c in cols:
        if c not in df.columns:
            df[c] = 0.0 if c == "Nota" else ""
    return df


def limpiar_excel(df):
    return df.replace([np.nan, np.inf, -np.inf], "")


def cargar_df_informe(parquet_path):
    if not os.path.exists(parquet_path):
        raise FileNotFoundError(f"No se encuentra el parquet: {parquet_path}")
    df = pd.read_parquet(parquet_path)
    df = ensure_columns(df)
    df["Curso"] = df["Curso"].astype(str).str.strip().str.upper()
    df["Carrera"] = df["Carrera"].astype(str).str.strip()
    df["Docente"] = df["Docente"].astype(str).str.strip()
    df["Seccion"] = df["Seccion"].astype(str).str.strip()
    df["Evaluacion"] = df["Evaluacion"].astype(str).str.strip().str.upper()
    df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce").fillna(0.0)
    return df


def round_half_up(x):
    """Redondeo al entero más cercano y .5 hacia arriba (x.5 -> x+1)."""
    try:
        x = float(x)
    except Exception:
        return 0
    return int(np.floor(x + 0.5)) if x >= 0 else int(np.ceil(x - 0.5))


# =========================
# CÁLCULO FINAL POR ESTUDIANTE (Curso)
# Regla oficial:
# - "No rindió" si NO rindió ninguna evaluación (todas 0).
# - FINAL = 0.18*R(EC1)+0.20*R(EP)+0.18*R(EC2)+0.19*R(EC3)+0.2*0.18*R(EF)
# - R = redondeo half-up al entero
# =========================
def pivot_final_curso(df_curso):
    piv = df_curso.pivot_table(
        index=["CodigoAlumno", "Alumno", "Curso", "Seccion", "Carrera", "Docente"],
        columns="Evaluacion",
        values="Nota",
        aggfunc="max",
        fill_value=0
    ).reset_index()

    for ev in EVALS:
        if ev not in piv.columns:
            piv[ev] = 0.0

    # Redondeos (half-up) solicitados
    for ev in ["EC1", "EP", "EC2", "EC3", "EF"]:
        piv[f"__R_{ev}"] = piv[ev].map(round_half_up)

    piv["Final"] = (
        FINAL_WEIGHTS["EC1"] * piv["__R_EC1"] +
        FINAL_WEIGHTS["EP"]  * piv["__R_EP"]  +
        FINAL_WEIGHTS["EC2"] * piv["__R_EC2"] +
        FINAL_WEIGHTS["EC3"] * piv["__R_EC3"] +
        FINAL_WEIGHTS["EF"]  * piv["__R_EF"]
    )

    piv.drop(columns=[c for c in piv.columns if c.startswith("__R_")], inplace=True)

    def estado(r):
        # “No rindió” SOLO si todas las evaluaciones están en 0
        if all(float(r[ev]) == 0.0 for ev in EVALS):
            return "No rindió"
        return "Aprobado" if float(r["Final"]) >= APROBADO_MIN else "Desaprobado"

    piv["EstadoFinal"] = piv.apply(estado, axis=1)
    return piv


# =========================
# HOJA 1: RESUMEN POR EVALUACIÓN
# Reglas:
# - Total: todos
# - No rindieron: Nota == 0
# - Rindieron: Nota > 0
# - % Aprobados y % Desaprobados: SOBRE RINDIERON
# - % No rindieron: SOBRE TOTAL
# - Promedio: SOLO Nota > 0
# - Aprobado: Nota >= 12.5
# =========================
def resumen_por_evaluacion(df_curso):
    rows = []
    for ev in EVALS:
        sub = df_curso[df_curso["Evaluacion"] == ev].copy()
        total = int(len(sub))

        if total == 0:
            rows.append({
                "Evaluación": ev,
                "Total": 0,
                "No rindieron": 0,
                "Rindieron": 0,
                "% Aprobados": 0.0,
                "% Desaprobados": 0.0,
                "% No rindieron": 0.0,
                "Promedio": 0.0
            })
            continue

        no_r = int((sub["Nota"] == 0).sum())
        rind = sub[sub["Nota"] > 0]["Nota"]
        rind_n = int(len(rind))

        aprob = int((rind >= APROBADO_MIN).sum())
        desap = int((rind < APROBADO_MIN).sum())
        prom = float(rind.mean()) if rind_n else 0.0

        rows.append({
            "Evaluación": ev,
            "Total": total,
            "No rindieron": no_r,
            "Rindieron": rind_n,
            "% Aprobados": round(aprob / rind_n * 100, 1) if rind_n else 0.0,
            "% Desaprobados": round(desap / rind_n * 100, 1) if rind_n else 0.0,
            "% No rindieron": round(no_r / total * 100, 1) if total else 0.0,
            "Promedio": round(prom, 2),
        })

    df1 = pd.DataFrame(rows)
    df1["__ord"] = df1["Evaluación"].map(lambda x: ORDER_EVAL.index(x) if x in ORDER_EVAL else 999)
    df1 = df1.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)
    return df1


# =========================
# HOJA 2: POR CARRERA (SOLO SITUACIÓN FINAL en tu Excel/PDF)
# Reglas:
# - FINAL:
#   * No rindieron: EstadoFinal == "No rindió" (todas 0)
#   * Rindieron: resto
#   * % Aprob / % Desap: SOBRE RINDIERON (descartando los que no rindieron nada)
#   * % No rind: SOBRE TOTAL
#   * Promedio: solo rindieron (Final > 0 por construcción si rindieron algo)
# - Aprobado: Final >= 12.5
# =========================
def resumen_carrera_eval(df_curso, ev_code):
    if ev_code == FINAL_NAME:
        piv = pivot_final_curso(df_curso)
        g = piv.groupby("Carrera", dropna=False)
        rows = []
        for carrera, d in g:
            total = int(len(d))
            no_r = int((d["EstadoFinal"] == "No rindió").sum())
            rind_n = int(total - no_r)

            dr = d[d["EstadoFinal"] != "No rindió"]
            rind = dr["Final"]

            aprob = int((rind >= APROBADO_MIN).sum())
            desap = int((rind < APROBADO_MIN).sum())
            prom = float(rind.mean()) if rind_n else 0.0

            rows.append({
                "Carreras": carrera,
                "Total": total,
                "No rindieron": no_r,
                "Rindieron": rind_n,
                "% Aprobados": round(aprob / rind_n * 100, 1) if rind_n else 0.0,
                "% Desaprobados": round(desap / rind_n * 100, 1) if rind_n else 0.0,
                "% No rindieron": round(no_r / total * 100, 1) if total else 0.0,
                "Promedio": round(prom, 2),
            })
        out = pd.DataFrame(rows)

    else:
        # (Si alguna vez lo usas para evaluaciones individuales)
        sub = df_curso[df_curso["Evaluacion"] == ev_code].copy()
        g = sub.groupby("Carrera", dropna=False)
        rows = []
        for carrera, d in g:
            total = int(len(d))
            no_r = int((d["Nota"] == 0).sum())

            rind = d[d["Nota"] > 0]["Nota"]
            rind_n = int(len(rind))

            aprob = int((rind >= APROBADO_MIN).sum())
            desap = int((rind < APROBADO_MIN).sum())
            prom = float(rind.mean()) if rind_n else 0.0

            rows.append({
                "Carreras": carrera,
                "Total": total,
                "No rindieron": no_r,
                "Rindieron": rind_n,
                "% Aprobados": round(aprob / rind_n * 100, 1) if rind_n else 0.0,
                "% Desaprobados": round(desap / rind_n * 100, 1) if rind_n else 0.0,
                "% No rindieron": round(no_r / total * 100, 1) if total else 0.0,
                "Promedio": round(prom, 2),
            })
        out = pd.DataFrame(rows)

    if out.empty:
        return out

    out = out.sort_values(by=["% Aprobados", "Promedio"], ascending=[False, False]).reset_index(drop=True)
    return out


# =========================
# HOJA 3: POR DOCENTE (FINAL)
# Reglas:
# - No rindieron: EstadoFinal == "No rindió" (todas 0)
# - % Aprob / % Desap: SOBRE RINDIERON
# - % No rind: SOBRE TOTAL
# - Promedio: solo rindieron
# =========================
def resumen_docente_final(df_curso):
    piv = pivot_final_curso(df_curso)

    # N° secciones por docente
    secc = piv.groupby("Docente")["Seccion"].nunique().reset_index(name="N° Sec.")
    g = piv.groupby("Docente", dropna=False)

    rows = []
    for docente, d in g:
        total = int(len(d))
        no_r = int((d["EstadoFinal"] == "No rindió").sum())
        rind_n = int(total - no_r)

        dr = d[d["EstadoFinal"] != "No rindió"]
        rind = dr["Final"]

        aprob = int((rind >= APROBADO_MIN).sum())
        desap = int((rind < APROBADO_MIN).sum())
        prom = float(rind.mean()) if rind_n else 0.0

        rows.append({
            "Docente": docente,
            "Total": total,
            "No rind.": no_r,
            "Rind.": rind_n,
            "% Aprob.": round(aprob / rind_n * 100, 1) if rind_n else 0.0,
            "% Desaprob.": round(desap / rind_n * 100, 1) if rind_n else 0.0,
            "% No rind.": round(no_r / total * 100, 1) if total else 0.0,
            "Prom.": round(prom, 2),
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out = out.merge(secc, on="Docente", how="left")
    out = out[["Docente", "N° Sec.", "Total", "No rind.", "Rind.", "% Aprob.", "% Desaprob.", "% No rind.", "Prom."]]
    out = out.sort_values(by=["% Aprob.", "Prom."], ascending=[False, False]).reset_index(drop=True)
    return out


# =========================
# HOJA 4: GLOBAL FINAL
# Reglas:
# - No rindieron: EstadoFinal == "No rindió" (todas 0)
# - % Aprob / % Desap: SOBRE RINDIERON (descartando “No rindió”)
# - % No rind: SOBRE TOTAL
# - Promedio: solo rindieron
# =========================
def resumen_global_final(df_curso):
    piv = pivot_final_curso(df_curso)
    total = int(len(piv))

    no_r = int((piv["EstadoFinal"] == "No rindió").sum())
    rind_n = int(total - no_r)

    dr = piv[piv["EstadoFinal"] != "No rindió"]
    ap = int((dr["Final"] >= APROBADO_MIN).sum())
    de = int((dr["Final"] < APROBADO_MIN).sum())

    prom = float(dr["Final"].mean()) if rind_n else 0.0

    return pd.DataFrame([{
        "Total de alumnos": total,
        "No rindieron": no_r,
        "Rindieron": rind_n,
        "% Aprob.": round(ap / rind_n * 100, 1) if rind_n else 0.0,
        "% Desaprob.": round(de / rind_n * 100, 1) if rind_n else 0.0,
        "% No rindieron": round(no_r / total * 100, 1) if total else 0.0,
        "Prom.": round(prom, 2),
    }])


# =========================
# EXCEL: Generación completa (4 hojas)
# (Se mantiene estructura: 4 hojas + charts; solo se ajustan columnas/índices)
# =========================
def generar_excel_informe_final_curso(df, curso, carpeta_destino):
    curso = str(curso).strip().upper()
    df_curso = df[df["Curso"].str.upper() == curso].copy()
    if df_curso.empty:
        raise ValueError("No hay datos para el curso seleccionado.")

    os.makedirs(carpeta_destino, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(carpeta_destino, f"INFORME_FINAL_CURSO_{slug(curso)}_{ts}.xlsx")

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_sub = wb.add_format({"bold": True, "font_size": 12})
        fmt_header = wb.add_format({
            "bold": True, "font_color": "white",
            "bg_color": COL_UCSUR_AZUL, "border": 1, "align": "center"
        })
        fmt_cell = wb.add_format({"border": 1})
        fmt_pct = wb.add_format({"border": 1, "num_format": "0.0"})
        fmt_num = wb.add_format({"border": 1, "num_format": "0.00"})

        # -------------------------
        # HOJA 1: EVALUACIONES
        # -------------------------
        sh1 = "1_EVALUACIONES"
        ws1 = wb.add_worksheet(sh1)
        writer.sheets[sh1] = ws1

        ws1.write(0, 0, f"CURSO: {curso}", fmt_title)
        ws1.write(1, 0, "RESUMEN POR TIPO DE EVALUACIÓN", fmt_sub)

        df1 = resumen_por_evaluacion(df_curso)
        df1_show = df1.copy()
        df1_show.columns = ["Evaluación", "Total", "No rindieron", "Rindieron",
                            "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"]
        df1_show = limpiar_excel(df1_show)

        start_row = 3
        for c, col in enumerate(df1_show.columns):
            ws1.write(start_row, c, col, fmt_header)

        for r in range(len(df1_show)):
            for c in range(len(df1_show.columns)):
                val = df1_show.iat[r, c]
                if c in (1, 2, 3):  # conteos
                    ws1.write(start_row + 1 + r, c, int(val) if str(val) != "" else "", fmt_cell)
                elif c in (4, 5, 6):  # porcentajes
                    ws1.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_pct)
                elif c == 7:  # promedio
                    ws1.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_num)
                else:
                    ws1.write(start_row + 1 + r, c, val, fmt_cell)

        ws1.set_column(0, 0, 14)
        ws1.set_column(1, 3, 12)
        ws1.set_column(4, 6, 16)
        ws1.set_column(7, 7, 10)

        # Gráfico 1: Evaluación vs % Aprobados (col 4)
        chart1 = wb.add_chart({"type": "column"})
        chart1.add_series({
            "name": "% Aprobados (sobre rindieron)",
            "categories": [sh1, start_row + 1, 0, start_row + len(df1_show), 0],
            "values": [sh1, start_row + 1, 4, start_row + len(df1_show), 4],
        })
        chart1.set_title({"name": "Evaluación vs % Aprobados (sobre rindieron)"})
        chart1.set_y_axis({"name": "%"})
        ws1.insert_chart(3, 9, chart1, {"x_scale": 1.2, "y_scale": 1.1})

        # Gráfico 2: Evaluación vs Promedio (col 7)
        chart2 = wb.add_chart({"type": "column"})
        chart2.add_series({
            "name": "Promedio (solo rindieron)",
            "categories": [sh1, start_row + 1, 0, start_row + len(df1_show), 0],
            "values": [sh1, start_row + 1, 7, start_row + len(df1_show), 7],
        })
        chart2.set_title({"name": "Evaluación vs Promedio (solo rindieron)"})
        chart2.set_y_axis({"name": "Prom"})
        ws1.insert_chart(18, 9, chart2, {"x_scale": 1.2, "y_scale": 1.1})

        # -------------------------
        # HOJA 2: CARRERAS (SOLO SITUACIÓN FINAL)
        # -------------------------
        sh2 = "2_CARRERAS"
        ws2 = wb.add_worksheet(sh2)
        writer.sheets[sh2] = ws2

        ws2.write(0, 0, f"CURSO: {curso}", fmt_title)
        ws2.write(1, 0, "ANÁLISIS POR CARRERA (SITUACIÓN FINAL)", fmt_sub)

        ev = FINAL_NAME
        df2 = resumen_carrera_eval(df_curso, ev)

        if df2.empty:
            ws2.write(3, 0, "Sin datos para situación final por carrera.", fmt_cell)
        else:
            df2_show = df2.copy()
            df2_show.columns = ["Carreras", "Total", "No rindieron", "Rindieron",
                                "% Aprobados", "% Desaprobados", "% No rindieron", "Promedio"]
            df2_show = limpiar_excel(df2_show)

            start_row = 3
            for c, col in enumerate(df2_show.columns):
                ws2.write(start_row, c, col, fmt_header)

            for r in range(len(df2_show)):
                for c in range(len(df2_show.columns)):
                    val = df2_show.iat[r, c]
                    if c in (1, 2, 3):
                        ws2.write(start_row + 1 + r, c, int(val) if str(val) != "" else "", fmt_cell)
                    elif c in (4, 5, 6):
                        ws2.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_pct)
                    elif c == 7:
                        ws2.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_num)
                    else:
                        ws2.write(start_row + 1 + r, c, val, fmt_cell)

            ws2.set_column(0, 0, 38)
            ws2.set_column(1, 3, 12)
            ws2.set_column(4, 6, 16)
            ws2.set_column(7, 7, 10)

            r0 = start_row + 1
            r1 = start_row + len(df2_show)

            # Gráfico 1: Carrera vs % Aprobados (col 4)
            ch2a = wb.add_chart({"type": "column"})
            ch2a.add_series({
                "name": "% Aprobados (sobre rindieron)",
                "categories": [sh2, r0, 0, r1, 0],
                "values":     [sh2, r0, 4, r1, 4],
            })
            ch2a.set_title({"name": "Situación Final — % Aprobados por carrera (sobre rindieron)"})
            ch2a.set_y_axis({"name": "%"})
            ws2.insert_chart(3, 9, ch2a, {"x_scale": 1.25, "y_scale": 1.1})

            # Gráfico 2: Carrera vs Promedio (col 7)
            ch2p = wb.add_chart({"type": "column"})
            ch2p.add_series({
                "name": "Promedio (solo rindieron)",
                "categories": [sh2, r0, 0, r1, 0],
                "values":     [sh2, r0, 7, r1, 7],
            })
            ch2p.set_title({"name": "Situación Final — Promedio por carrera (solo rindieron)"})
            ch2p.set_y_axis({"name": "Prom"})
            ws2.insert_chart(18, 9, ch2p, {"x_scale": 1.25, "y_scale": 1.1})

        # -------------------------
        # HOJA 3: DOCENTES (FINAL)
        # -------------------------
        sh3 = "3_DOCENTES"
        ws3 = wb.add_worksheet(sh3)
        writer.sheets[sh3] = ws3

        ws3.write(0, 0, f"CURSO: {curso}", fmt_title)
        ws3.write(1, 0, "ANÁLISIS POR DOCENTE (SITUACIÓN FINAL)", fmt_sub)

        df3 = resumen_docente_final(df_curso)
        if df3.empty:
            ws3.write(3, 0, "Sin datos.", fmt_cell)
        else:
            start_row = 3
            for c, col in enumerate(df3.columns):
                ws3.write(start_row, c, col, fmt_header)

            for r in range(len(df3)):
                for c in range(len(df3.columns)):
                    val = df3.iat[r, c]
                    if c in (1, 2, 3, 4):  # N° Sec, Total, No rind, Rind
                        ws3.write(start_row + 1 + r, c, int(val) if str(val) != "" else "", fmt_cell)
                    elif c in (5, 6, 7):  # % cols
                        ws3.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_pct)
                    elif c == 8:  # Prom
                        ws3.write(start_row + 1 + r, c, float(val) if str(val) != "" else "", fmt_num)
                    else:
                        ws3.write(start_row + 1 + r, c, val, fmt_cell)

            ws3.set_column(0, 0, 45)
            ws3.set_column(1, 4, 12)
            ws3.set_column(5, 7, 14)
            ws3.set_column(8, 8, 10)

            # chart % aprob (col 5)
            chd1 = wb.add_chart({"type": "column"})
            chd1.add_series({
                "name": "% Aprob. (sobre rindieron)",
                "categories": [sh3, start_row + 1, 0, start_row + len(df3), 0],
                "values": [sh3, start_row + 1, 5, start_row + len(df3), 5],
            })
            chd1.set_title({"name": "Docente vs % Aprob. (sobre rindieron)"})
            ws3.insert_chart(3, 11, chd1, {"x_scale": 1.25, "y_scale": 1.1})

            # chart promedio (col 8)
            chd2 = wb.add_chart({"type": "column"})
            chd2.add_series({
                "name": "Prom. (solo rindieron)",
                "categories": [sh3, start_row + 1, 0, start_row + len(df3), 0],
                "values": [sh3, start_row + 1, 8, start_row + len(df3), 8],
            })
            chd2.set_title({"name": "Docente vs Prom. (solo rindieron)"})
            ws3.insert_chart(22, 11, chd2, {"x_scale": 1.25, "y_scale": 1.1})

        # -------------------------
        # HOJA 4: GLOBAL (FINAL)
        # -------------------------
        sh4 = "4_GLOBAL"
        ws4 = wb.add_worksheet(sh4)
        writer.sheets[sh4] = ws4

        ws4.write(0, 0, f"CURSO: {curso}", fmt_title)
        ws4.write(1, 0, "GLOBAL (SITUACIÓN FINAL)", fmt_sub)

        df4 = resumen_global_final(df_curso)
        start_row = 3

        for c, col in enumerate(df4.columns):
            ws4.write(start_row, c, col, fmt_header)

        for r in range(len(df4)):
            for c in range(len(df4.columns)):
                val = df4.iat[r, c]
                if c in (0, 1, 2):  # conteos
                    ws4.write(start_row + 1 + r, c, int(val), fmt_cell)
                elif c in (3, 4, 5):  # %
                    ws4.write(start_row + 1 + r, c, float(val), fmt_pct)
                else:  # Prom
                    ws4.write(start_row + 1 + r, c, float(val), fmt_num)

        ws4.set_column(0, 0, 18)
        ws4.set_column(1, 2, 12)
        ws4.set_column(3, 5, 16)
        ws4.set_column(6, 6, 10)

        # Pie chart (% Aprob / % Desap / % No rindieron)
        pie = wb.add_chart({"type": "pie"})
        ws4.write(7, 0, "Aprobados", fmt_cell)
        ws4.write(7, 1, float(df4.loc[0, "% Aprob."]), fmt_pct)
        ws4.write(8, 0, "Desaprobados", fmt_cell)
        ws4.write(8, 1, float(df4.loc[0, "% Desaprob."]), fmt_pct)
        ws4.write(9, 0, "No rindieron", fmt_cell)
        ws4.write(9, 1, float(df4.loc[0, "% No rindieron"]), fmt_pct)

        pie.add_series({
            "name": "% Aprob. / % Desaprob. / % No rindieron",
            "categories": [sh4, 7, 0, 9, 0],
            "values": [sh4, 7, 1, 9, 1],
        })
        pie.set_title({"name": "Distribución porcentual (sobre rindieron / total)"} )
        ws4.insert_chart(3, 9, pie, {"x_scale": 1.25, "y_scale": 1.25})

    return os.path.abspath(out)


# =========================
# PDF: helpers
# =========================
def cargar_fuente(pdf):
    try:
        arial = r"C:\Windows\Fonts\arial.ttf"
        arial_bold = r"C:\Windows\Fonts\arialbd.ttf"
        if os.path.exists(arial):
            pdf.add_font("ArialUnicode", "", arial)
        if os.path.exists(arial_bold):
            pdf.add_font("ArialUnicode", "B", arial_bold)
        return "ArialUnicode"
    except Exception:
        return "Helvetica"


def pdf_encabezado(pdf, titulo):
    try:
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=10, y=8, w=26)
    except Exception:
        pass

    pdf.set_xy(10, 12)
    pdf.set_font(pdf.font_family, "B", 14)
    pdf.cell(0, 7, "UNIVERSIDAD CIENTÍFICA DEL SUR", align="C", ln=1)

    pdf.set_font(pdf.font_family, "", 11)
    pdf.cell(0, 6, "Departamento de Cursos Básicos", align="C", ln=1)
    pdf.cell(0, 6, safe_txt(titulo), align="C", ln=1)
    pdf.set_y(38)


def pdf_tabla(pdf, fam, headers, rows, widths):
    def head():
        pdf.set_fill_color(0, 59, 112)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font(fam, "B", 9)
        for h, w in zip(headers, widths):
            pdf.cell(w, 7, safe_txt(h), border=1, align="C", fill=True)
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.set_font(fam, "", 9)

    head()
    for r in rows:
        if pdf.get_y() > 260:
            pdf.add_page()
            head()
        for v, w in zip(r, widths):
            pdf.cell(w, 6, safe_txt(v), border=1, align="C")
        pdf.ln()


def _plot_bar(labels, values, title, ylabel, out_png):
    plt.figure(figsize=(7.2, 3.2))
    plt.bar(labels, values)
    plt.xticks(rotation=30, ha="right", fontsize=8)
    plt.title(title)
    plt.ylabel(ylabel)
    plt.tight_layout()
    plt.savefig(out_png, dpi=150)
    plt.close()


def _plot_pie(labels, values, title, out_png):
    plt.figure(figsize=(5.2, 4.0))
    plt.pie(values, labels=labels, autopct="%1.1f%%")
    plt.title(title)
    plt.tight_layout()
    plt.savefig(out_png, dpi=150)
    plt.close()


# =========================
# PDF: Generación completa
# (Se mantiene estructura: 4 secciones; solo se ajustan tablas/plots)
# =========================
def generar_pdf_informe_final_curso(df, curso, carpeta_destino):
    curso = str(curso).strip().upper()
    df_curso = df[df["Curso"].str.upper() == curso].copy()
    if df_curso.empty:
        raise ValueError("No hay datos para el curso seleccionado.")

    os.makedirs(carpeta_destino, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(carpeta_destino, f"INFORME_FINAL_CURSO_{slug(curso)}_{ts}.pdf")
    tmp = tempfile.mkdtemp()

    pdf = FPDF()
    pdf.add_page()
    fam = cargar_fuente(pdf)
    pdf.set_font(fam, "", 10)

    # 1) Evaluaciones
    pdf_encabezado(pdf, f"INFORME FINAL DE CURSO — {curso}")
    pdf.set_font(fam, "B", 12)
    pdf.cell(0, 7, "1. Resumen por Evaluación", ln=1)
    pdf.set_font(fam, "", 10)
    pdf.cell(0, 6, f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1)
    pdf.ln(2)

    df1 = resumen_por_evaluacion(df_curso)
    rows = []
    for _, r in df1.iterrows():
        rows.append([
            r["Evaluación"],
            str(int(r["Total"])),
            str(int(r["No rindieron"])),
            str(int(r["Rindieron"])),
            f"{r['% Aprobados']:.1f}%",
            f"{r['% Desaprobados']:.1f}%",
            f"{r['% No rindieron']:.1f}%",
            f"{r['Promedio']:.2f}",
        ])

    pdf_tabla(
        pdf, fam,
        headers=["Evaluación", "Total", "No rind.", "Rind.", "% Aprob.", "% Desaprob.", "% No rind.", "Prom."],
        rows=rows,
        widths=[22, 14, 16, 14, 20, 24, 22, 18]
    )

    g1 = os.path.join(tmp, "eval_pct.png")
    g2 = os.path.join(tmp, "eval_prom.png")
    _plot_bar(df1["Evaluación"], df1["% Aprobados"], "Evaluación vs % Aprobados (sobre rindieron)", "%", g1)
    _plot_bar(df1["Evaluación"], df1["Promedio"], "Evaluación vs Promedio (solo rindieron)", "Prom", g2)

    pdf.ln(2)
    pdf.image(g1, w=185)
    pdf.ln(2)
    pdf.image(g2, w=185)

    # 2) Carreras (SOLO SITUACIÓN FINAL)
    pdf.add_page()
    pdf_encabezado(pdf, f"ANÁLISIS POR CARRERA — {curso}")
    pdf.set_font(fam, "B", 12)
    pdf.cell(0, 7, "2. Análisis por Carrera (Situación Final)", ln=1)

    ev = FINAL_NAME
    df_ev = resumen_carrera_eval(df_curso, ev)

    if df_ev.empty:
        pdf.set_font(fam, "", 10)
        pdf.cell(0, 6, "Sin datos para carreras en situación final.", ln=1)
    else:
        pdf.set_font(fam, "B", 11)
        pdf.cell(0, 7, "SITUACIÓN FINAL — Tabla por Carrera", ln=1)

        rows = []
        for _, r in df_ev.iterrows():
            rows.append([
                safe_txt(r["Carreras"])[:28],
                str(int(r["Total"])),
                str(int(r["No rindieron"])),
                str(int(r["Rindieron"])),
                f"{r['% Aprobados']:.1f}%",
                f"{r['% Desaprobados']:.1f}%",
                f"{r['% No rindieron']:.1f}%",
                f"{r['Promedio']:.2f}",
            ])

        pdf_tabla(
            pdf, fam,
            headers=["Carrera", "Total", "No rind.", "Rind.", "% Aprob.", "% Desaprob.", "% No rind.", "Prom."],
            rows=rows,
            widths=[48, 14, 16, 14, 20, 24, 22, 18]
        )

        g_pct = os.path.join(tmp, "car_pct_final.png")
        g_pr  = os.path.join(tmp, "car_pr_final.png")
        _plot_bar(df_ev["Carreras"], df_ev["% Aprobados"],
                  "Situación Final — % Aprobados por carrera (sobre rindieron)", "%", g_pct)
        _plot_bar(df_ev["Carreras"], df_ev["Promedio"],
                  "Situación Final — Promedio por carrera (solo rindieron)", "Prom", g_pr)

        pdf.ln(2)
        pdf.image(g_pct, w=185)
        pdf.ln(2)
        pdf.image(g_pr, w=185)

    # 3) Docentes
    pdf.add_page()
    pdf_encabezado(pdf, f"ANÁLISIS POR DOCENTE — {curso}")
    pdf.set_font(fam, "B", 12)
    pdf.cell(0, 7, "3. Análisis por Docente (Situación Final)", ln=1)

    df3 = resumen_docente_final(df_curso)
    if not df3.empty:
        rows = []
        for _, r in df3.iterrows():
            rows.append([
                safe_txt(r["Docente"])[:28],
                str(int(r["N° Sec."])) if str(r["N° Sec."]) != "" else "",
                str(int(r["Total"])),
                str(int(r["No rind."])),
                str(int(r["Rind."])),
                f"{float(r['% Aprob.']):.1f}%",
                f"{float(r['% Desaprob.']):.1f}%",
                f"{float(r['% No rind.']):.1f}%",
                f"{float(r['Prom.']):.2f}",
            ])

        pdf_tabla(
            pdf, fam,
            headers=["Docente", "N° Sec.", "Total", "No rind.", "Rind.",
                     "% Aprob.", "% Desaprob.", "% No rind.", "Prom."],
            rows=rows,
            widths=[48, 14, 12, 16, 14, 18, 22, 20, 18]
        )

        gdp = os.path.join(tmp, "doc_pct.png")
        gdm = os.path.join(tmp, "doc_prom.png")
        _plot_bar(df3["Docente"], df3["% Aprob."], "Docente vs % Aprob. (sobre rindieron)", "%", gdp)
        _plot_bar(df3["Docente"], df3["Prom."], "Docente vs Prom. (solo rindieron)", "Prom", gdm)

        pdf.ln(2)
        pdf.image(gdp, w=185)
        pdf.ln(2)
        pdf.image(gdm, w=185)

    # 4) Global + pie
    pdf.add_page()
    pdf_encabezado(pdf, f"GLOBAL — {curso}")
    pdf.set_font(fam, "B", 12)
    pdf.cell(0, 7, "4. Global (Situación Final)", ln=1)

    df4 = resumen_global_final(df_curso)
    r = df4.loc[0].to_dict()

    pdf.set_font(fam, "", 10)
    pdf.ln(2)
    pdf.cell(0, 6, f"Total de alumnos: {int(r['Total de alumnos'])}", ln=1)
    pdf.cell(0, 6, f"No rindieron ninguna evaluación: {int(r['No rindieron'])}", ln=1)
    pdf.cell(0, 6, f"Rindieron al menos una evaluación: {int(r['Rindieron'])}", ln=1)
    pdf.cell(0, 6, f"% Aprobados (sobre rindieron): {r['% Aprob.']}%", ln=1)
    pdf.cell(0, 6, f"% Desaprobados (sobre rindieron): {r['% Desaprob.']}%", ln=1)
    pdf.cell(0, 6, f"% No rindieron (sobre total): {r['% No rindieron']}%", ln=1)
    pdf.cell(0, 6, f"Promedio (solo rindieron): {r['Prom.']}", ln=1)

    gp = os.path.join(tmp, "pie.png")
    _plot_pie(
        ["Aprob. (sobre rind.)", "Desaprob. (sobre rind.)", "No rind. (sobre total)"],
        [float(r["% Aprob."]), float(r["% Desaprob."]), float(r["% No rindieron"])],
        "Distribución porcentual",
        gp
    )
    pdf.ln(3)
    pdf.image(gp, w=140)

    pdf.output(out)
    return os.path.abspath(out)
