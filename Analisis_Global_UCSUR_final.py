# -*- coding: utf-8 -*-
"""
Sistema Analizador Académico UCSUR — Versión 6.5 (estable)

AJUSTES (enero 2026):
- En análisis de aprobados (final y por evaluaciones), el % de aprobados/desaprobados se calcula
  SOLO sobre quienes SÍ rindieron (nota > 0). Los que no rindieron se reportan aparte.
- La situación final (Estado) se mantiene: Final==0 => "Retirado". Además, para porcentajes finales
  se descarta del denominador a los "Retirado" (no rindieron nada).
- Nuevo cálculo de Final:
  Final = 0.18*Redondeo(EC1) + 0.20*Redondeo(EP) + 0.18*Redondeo(EC2) + 0.19*Redondeo(EC3)
          + 0.2*0.18*Redondeo(EF)
  (Se respeta exactamente lo indicado: 0.2*0.18 para EF)
- Redondeo: al entero más cercano y x.5 -> x+1 (round half up).
- Aprobado: nota >= 12.5 (en final y en evaluaciones).
"""

import os
import unicodedata
import tempfile
import warnings
from datetime import datetime

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=FutureWarning)

# ======================================================
# CONFIGURACIÓN
# ======================================================
OUTPUT_DIR = os.path.join(os.getcwd(), "Reportes")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# (La GUI usa esto)
LAST_SAVE_DIR = OUTPUT_DIR

PARQUET_IN = "notas_filtradas_ucsur.parquet"

EVALS = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]
ORDER_EVAL = ["ED", "EC1", "EP", "EC2", "EC3", "EF", "FINAL"]

WEIGHTS = {
    "ED": 0.0,
    "EC1": 0.18,
    "EP": 0.20,
    "EC2": 0.18,
    "EC3": 0.19,
    "EF": 0.25,
}

# Umbrales / reglas
APROBADO_MIN = 12.5

# Pesos del FINAL (según tu indicación exacta)
FINAL_WEIGHTS = {
    "EC1": 0.18,
    "EP": 0.20,
    "EC2": 0.18,
    "EC3": 0.19,
    "EF": 0.2 * 0.18,  # <- se respeta 0.2*0.18
}

COL_UCSUR_AZUL = "#003B70"
LOGO_PATH = "logo_ucsur.png"


# ======================================================
# UTILIDADES
# ======================================================
def safe_txt(s):
    """Evita errores en FPDF (latin-1)."""
    if s is None:
        return ""
    s = str(s).replace("—", "-")
    return s.encode("latin-1", "replace").decode("latin-1")


def slug(s):
    """Para nombres de archivo."""
    s2 = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s2 = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in s2.strip())
    while "__" in s2:
        s2 = s2.replace("__", "_")
    return s2 or "reporte"


def ordenar_eval(df_tabla, col="Evaluación"):
    """Orden institucional ED, EC1, EP, EC2, EC3, EF, FINAL."""
    if df_tabla is None or df_tabla.empty or col not in df_tabla.columns:
        return df_tabla
    orden_map = {ev: i for i, ev in enumerate(ORDER_EVAL)}
    out = df_tabla.copy()
    out["__ord"] = out[col].map(lambda x: orden_map.get(str(x), 999))
    out = out.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)
    return out


def eval_tiene_registros_validos(df, ev):
    """True si la evaluación tiene al menos una nota > 0."""
    sub = df[df["Evaluacion"] == ev]
    if sub.empty:
        return False
    return (sub["Nota"] > 0).any()


def _validar_columnas_minimas(df):
    req = ["CodigoAlumno", "Alumno", "Curso", "Evaluacion", "Nota"]
    faltan = [c for c in req if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas requeridas en parquet: {faltan}")


def _filtrar_evals_validas(df, evals_sel):
    """Mantiene evals en EVALS y con al menos una nota > 0."""
    base = [ev for ev in (evals_sel or []) if ev in EVALS]
    return [ev for ev in base if eval_tiene_registros_validos(df, ev)]


def round_half_up(x):
    """
    Redondeo al entero más cercano y .5 hacia arriba (round half up).
    Ej: 12.4->12, 12.5->13, 12.6->13
    """
    try:
        x = float(x)
    except Exception:
        return 0
    # Notas son >=0 usualmente, pero esto funciona también para negativos
    return int(np.floor(x + 0.5)) if x >= 0 else int(np.ceil(x - 0.5))


# ======================================================
# CARGA DE DATOS (parquet actualizado)
# ======================================================
def cargar_df(path=PARQUET_IN):
    if not os.path.exists(path):
        raise FileNotFoundError(f"No existe el parquet: {path}")

    df = pd.read_parquet(path)
    _validar_columnas_minimas(df)

    df["Curso"] = df["Curso"].astype(str).str.strip()
    df["Evaluacion"] = df["Evaluacion"].astype(str).str.upper().str.strip()
    df["Nota"] = pd.to_numeric(df["Nota"], errors="coerce").fillna(0.0)

    return df


# ======================================================
# CONDICIÓN FINAL
# ======================================================
def compute_final_table(df):
    piv = df.pivot_table(
        index=["CodigoAlumno", "Alumno", "Curso"],
        columns="Evaluacion",
        values="Nota",
        aggfunc="max",
        fill_value=0
    ).reset_index()

    for ev in EVALS:
        if ev not in piv.columns:
            piv[ev] = 0.0

    # --- Nuevo cálculo del Final con redondeo half-up por evaluación ---
    # ED no entra al promedio final
    for ev in ["EC1", "EP", "EC2", "EC3", "EF"]:
        piv[f"__R_{ev}"] = piv[ev].map(round_half_up)

    piv["Final"] = (
        FINAL_WEIGHTS["EC1"] * piv["__R_EC1"] +
        FINAL_WEIGHTS["EP"]  * piv["__R_EP"]  +
        FINAL_WEIGHTS["EC2"] * piv["__R_EC2"] +
        FINAL_WEIGHTS["EC3"] * piv["__R_EC3"] +
        FINAL_WEIGHTS["EF"]  * piv["__R_EF"]
    )

    piv = piv.drop(columns=[c for c in piv.columns if c.startswith("__R_")])

    def estado(row):
        if float(row["Final"]) == 0.0:
            return "Retirado"
        return "Aprobado" if float(row["Final"]) >= APROBADO_MIN else "Desaprobado"

    piv["Estado"] = piv.apply(estado, axis=1)
    return piv


def compute_aprobados_por_curso(df, cursos_sel=None):
    finales = compute_final_table(df)

    if cursos_sel:
        finales = finales[finales["Curso"].isin(cursos_sel)].copy()

    rows = []
    for curso, g in finales.groupby("Curso"):
        total = int(len(g))

        no_rindieron = int((g["Estado"] == "Retirado").sum())
        rindieron = int(total - no_rindieron)

        ap = int((g["Estado"] == "Aprobado").sum())
        de = int((g["Estado"] == "Desaprobado").sum())
        re = int(no_rindieron)

        # % Aprob/Desap: sobre quienes rindieron; % Retirados: sobre total
        pct_ap = round(ap / rindieron * 100, 1) if rindieron else 0.0
        pct_de = round(de / rindieron * 100, 1) if rindieron else 0.0
        pct_re = round(re / total * 100, 1) if total else 0.0

        rows.append({
            "Curso": curso,
            "Total": total,
            "No rindieron": no_rindieron,
            "Rindieron": rindieron,
            "% Aprobados": pct_ap,
            "% Desaprobados": pct_de,
            "% Retirados": pct_re,
        })

    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values("% Aprobados", ascending=False).reset_index(drop=True)
    return out


# ======================================================
# EVALUACIONES
# ======================================================
def compute_aprobados_por_evaluacion(df, curso, evals_sel):
    """Diagnóstico 1 curso: Totales y % aprobados (sobre quienes rindieron) + promedio."""
    df_c = df[df["Curso"] == curso].copy()
    rows = []
    for ev in evals_sel:
        sub = df_c[df_c["Evaluacion"] == ev]
        total = int(len(sub))

        rind_series = sub[sub["Nota"] > 0]["Nota"]
        rindieron = int(len(rind_series))
        no_rindieron = int(total - rindieron)

        ap = int((rind_series >= APROBADO_MIN).sum())
        prom = float(rind_series.mean()) if rindieron else 0.0

        rows.append({
            "Evaluación": ev,
            "Total": total,
            "No rindieron": no_rindieron,
            "Rindieron": rindieron,
            "% Aprobados": round(ap / rindieron * 100, 1) if rindieron else 0.0,
            "Promedio": round(prom, 2) if rindieron else 0.0
        })

    out = pd.DataFrame(rows)
    out = ordenar_eval(out, col="Evaluación")
    return out


def compute_aprobados_eval_por_curso(df, cursos_sel, evals_sel):
    """
    Comparación >=2 cursos:
    devuelve tabla larga:
    Curso, Evaluación, Total, No rindieron, Rindieron, % Aprobados, Promedio
    """
    rows = []
    for curso in cursos_sel:
        df_c = df[df["Curso"] == curso].copy()
        for ev in evals_sel:
            sub = df_c[df_c["Evaluacion"] == ev]
            total = int(len(sub))

            rind_series = sub[sub["Nota"] > 0]["Nota"]
            rindieron = int(len(rind_series))
            no_rindieron = int(total - rindieron)

            ap = int((rind_series >= APROBADO_MIN).sum())
            prom = float(rind_series.mean()) if rindieron else 0.0

            rows.append({
                "Curso": curso,
                "Evaluación": ev,
                "Total": total,
                "No rindieron": no_rindieron,
                "Rindieron": rindieron,
                "% Aprobados": round(ap / rindieron * 100, 1) if rindieron else 0.0,
                "Promedio": round(prom, 2) if rindieron else 0.0
            })

    out = pd.DataFrame(rows)
    if not out.empty:
        out = ordenar_eval(out, col="Evaluación")
    return out


# ======================================================
# PDF helpers
# ======================================================
def _pdf_encabezado(pdf, titulo):
    try:
        if os.path.exists(LOGO_PATH):
            pdf.image(LOGO_PATH, x=10, y=8, w=26)
    except Exception:
        pass

    pdf.set_xy(10, 12)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 7, safe_txt("UNIVERSIDAD CIENTÍFICA DEL SUR"), ln=1, align="C")

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, safe_txt("Departamento de Cursos Básicos"), ln=1, align="C")
    pdf.cell(0, 6, safe_txt(titulo), ln=1, align="C")

    pdf.set_y(38)


def _fmt_cell_value(colname, val):
    """Formatea valores para tabla PDF."""
    if isinstance(val, (int, float, np.integer, np.floating)):
        if "%" in str(colname):
            return f"{float(val):.1f} %"
        if str(colname).strip().lower() in ("n", "total", "no rindieron", "rindieron"):
            return f"{int(val)}"
        return f"{float(val):.2f}"
    return str(val)


def _pdf_tabla_simple(pdf, df, col_widths, header_fill=(0, 59, 112)):
    if df is None or df.empty:
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, safe_txt("Sin datos para mostrar."), ln=1)
        return

    # Header
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(*header_fill)
    pdf.set_text_color(255, 255, 255)

    cols = list(df.columns)
    for c, w in zip(cols, col_widths):
        pdf.cell(w, 7, safe_txt(c), border=1, align="C", fill=True)
    pdf.ln()

    # Body
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(0, 0, 0)

    for _, row in df.iterrows():
        for (c, w) in zip(cols, col_widths):
            txt = _fmt_cell_value(c, row[c])
            pdf.cell(w, 6, safe_txt(txt), border=1)
        pdf.ln()


# ======================================================
# EXPORTACIÓN PDF (CORE)
# ======================================================
def exportar_pdf(df, cursos_sel=None, evals_sel=None, carpeta_destino=None):
    import matplotlib.pyplot as plt
    from fpdf import FPDF

    carpeta = carpeta_destino or OUTPUT_DIR
    os.makedirs(carpeta, exist_ok=True)

    tmpdir = tempfile.mkdtemp()

    # ==========================
    # 1) PANORAMA FINAL (siempre)
    # ==========================
    tabla_cursos = compute_aprobados_por_curso(df, cursos_sel=cursos_sel)

    graf_final = None
    if not tabla_cursos.empty:
        graf_final = os.path.join(tmpdir, "final_por_curso.png")
        plt.figure(figsize=(7.2, 3.2))
        plt.bar(tabla_cursos["Curso"], tabla_cursos["% Aprobados"])
        plt.xticks(rotation=30, ha="right", fontsize=8)
        plt.ylim(0, 100)
        plt.ylabel("% Aprobados (sobre Rindieron)")
        plt.title("% Aprobados por Curso (Condición Final)")
        plt.tight_layout()
        plt.savefig(graf_final, dpi=150)
        plt.close()

    pdf = FPDF()
    pdf.add_page()
    _pdf_encabezado(pdf, "Reporte Académico Global — Condición Final")

    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 6, safe_txt(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), ln=1)
    if cursos_sel:
        pdf.multi_cell(0, 5, safe_txt("Cursos seleccionados: " + ", ".join(cursos_sel)))
    else:
        pdf.cell(0, 6, safe_txt("Cursos: Todos"), ln=1)

    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, safe_txt("Panorama Global — % Aprobados por Curso"), ln=1)

    if not tabla_cursos.empty:
        show = tabla_cursos[
            ["Curso", "Total", "No rindieron", "Rindieron", "% Aprobados", "% Desaprobados", "% Retirados"]
        ].copy()
        _pdf_tabla_simple(pdf, show, col_widths=[55, 16, 22, 18, 22, 27, 25])
    else:
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, safe_txt("No hay datos para panorama (revisa el parquet)."), ln=1)

    if graf_final and os.path.exists(graf_final):
        pdf.ln(3)
        pdf.image(graf_final, w=185)

    # ==========================
    # 2) MODO EVALUACIONES (opcional)
    # ==========================
    if evals_sel:
        if not cursos_sel:
            cursos_sel = sorted(df["Curso"].unique().tolist())

        evals_validas = _filtrar_evals_validas(df, evals_sel)

        if not evals_validas:
            pdf.add_page()
            _pdf_encabezado(pdf, "Modo Evaluaciones")
            pdf.set_font("Helvetica", "", 10)
            pdf.multi_cell(0, 6, safe_txt("No hay evaluaciones válidas con registros (>0)."))
        else:
            # 2.a) Diagnóstico 1 curso
            if cursos_sel and len(cursos_sel) == 1:
                curso = cursos_sel[0]
                df_curso = df[df["Curso"] == curso].copy()
                evals_validas_curso = _filtrar_evals_validas(df_curso, evals_validas)

                pdf.add_page()
                _pdf_encabezado(pdf, f"Diagnóstico por Evaluación — {curso}")

                if not evals_validas_curso:
                    pdf.set_font("Helvetica", "", 10)
                    pdf.multi_cell(0, 6, safe_txt("Este curso no tiene evaluaciones con registros (>0)."))
                else:
                    df_ev = compute_aprobados_por_evaluacion(df, curso, evals_validas_curso)

                    pdf.set_font("Helvetica", "B", 11)
                    pdf.cell(0, 7, safe_txt("Resumen — % Aprobados (sobre Rindieron) y Promedio"), ln=1)
                    _pdf_tabla_simple(
                        pdf,
                        df_ev[["Evaluación", "Total", "No rindieron", "Rindieron", "% Aprobados", "Promedio"]],
                        col_widths=[28, 18, 24, 20, 26, 22]
                    )

                    graf_diag = os.path.join(tmpdir, "diag_curso.png")
                    plt.figure(figsize=(6.5, 3.0))
                    plt.bar(df_ev["Evaluación"], df_ev["% Aprobados"])
                    plt.ylim(0, 100)
                    plt.ylabel("% Aprobados (sobre Rindieron)")
                    plt.title(f"{curso} — % Aprobados por Evaluación")
                    plt.tight_layout()
                    plt.savefig(graf_diag, dpi=150)
                    plt.close()

                    pdf.ln(3)
                    pdf.image(graf_diag, w=175)

            # 2.b) Comparación >=2 cursos
            elif cursos_sel and len(cursos_sel) > 1:
                df_cmp = compute_aprobados_eval_por_curso(df, cursos_sel, evals_validas)

                pdf.add_page()
                _pdf_encabezado(pdf, "Comparación Curso × Evaluación")

                pivot_ap = df_cmp.pivot_table(
                    index="Curso",
                    columns="Evaluación",
                    values="% Aprobados",
                    aggfunc="mean",
                    fill_value=0.0
                )

                pivot_pr = df_cmp.pivot_table(
                    index="Curso",
                    columns="Evaluación",
                    values="Promedio",
                    aggfunc="mean",
                    fill_value=0.0
                )

                pivot_pdf = pivot_ap.copy()
                for ev in pivot_pdf.columns:
                    if ev in pivot_pr.columns:
                        pivot_pdf[ev] = [
                            f"{float(p):.1f} % | {float(m):.2f}"
                            for p, m in zip(pivot_ap[ev].values, pivot_pr[ev].values)
                        ]
                    else:
                        pivot_pdf[ev] = [f"{float(p):.1f} % | 0.00" for p in pivot_ap[ev].values]

                pivot_pdf = pivot_pdf.reset_index()

                cols_ord = ["Curso"] + [ev for ev in ORDER_EVAL if ev in pivot_pdf.columns]
                pivot_pdf = pivot_pdf[cols_ord]

                col_widths = [55] + [22] * (len(pivot_pdf.columns) - 1)
                _pdf_tabla_simple(pdf, pivot_pdf, col_widths=col_widths)

                cursos = pivot_ap.index.tolist()
                x = np.arange(len(cursos))
                m = max(1, len(evals_validas))
                width = min(0.8 / m, 0.18)

                graf_cmp = os.path.join(tmpdir, "cmp.png")
                plt.figure(figsize=(7.6, 3.2))

                for i, ev in enumerate(evals_validas):
                    y = pivot_ap[ev].values if ev in pivot_ap.columns else np.zeros(len(cursos))
                    plt.bar(x + (i - (m - 1) / 2) * width, y, width, label=ev)

                plt.xticks(x, cursos, rotation=30, ha="right", fontsize=8)
                plt.ylim(0, 100)
                plt.ylabel("% Aprobados (sobre Rindieron)")
                plt.title("% Aprobados — Cursos vs Evaluaciones")
                plt.legend(fontsize=8, ncol=min(6, m))
                plt.tight_layout()
                plt.savefig(graf_cmp, dpi=150)
                plt.close()

                pdf.ln(3)
                pdf.image(graf_cmp, w=185)

    # ==========================
    # Guardar
    # ==========================
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf = "TODOS" if not cursos_sel else (slug(cursos_sel[0])[:20] if len(cursos_sel) == 1 else "MULTI")
    out = os.path.join(carpeta, f"Reporte_Global_{suf}_{ts}.pdf")
    pdf.output(out)
    return out


# ======================================================
# EXPORTACIÓN EXCEL (CORE)
# ======================================================
def exportar_excel(df, cursos_sel=None, evals_sel=None, carpeta_destino=None):
    carpeta = carpeta_destino or OUTPUT_DIR
    os.makedirs(carpeta, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suf = "TODOS" if not cursos_sel else (slug(cursos_sel[0])[:20] if len(cursos_sel) == 1 else "MULTI")
    xlsx = os.path.join(carpeta, f"Analisis_Global_{suf}_{ts}.xlsx")

    tabla_final = compute_aprobados_por_curso(df, cursos_sel=cursos_sel)

    with pd.ExcelWriter(xlsx, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_header = wb.add_format({
            "bold": True, "bg_color": COL_UCSUR_AZUL,
            "font_color": "white", "border": 1, "align": "center"
        })
        fmt_cell = wb.add_format({"border": 1})

        # 1) Condición Final
        ws1 = wb.add_worksheet("Condición Final")
        writer.sheets["Condición Final"] = ws1

        if tabla_final.empty:
            ws1.write(0, 0, "Sin datos. Revisa el parquet.", fmt_cell)
        else:
            for c, col in enumerate(tabla_final.columns):
                ws1.write(0, c, col, fmt_header)
            for r in range(len(tabla_final)):
                for c in range(len(tabla_final.columns)):
                    ws1.write(r + 1, c, tabla_final.iat[r, c], fmt_cell)

        # 2) Modo evaluaciones (si aplica)
        if evals_sel:
            if not cursos_sel:
                cursos_sel = sorted(df["Curso"].unique().tolist())

            evals_validas = _filtrar_evals_validas(df, evals_sel)

            # 2.a Diagnóstico (1 curso)
            if cursos_sel and len(cursos_sel) == 1:
                curso = cursos_sel[0]
                df_curso = df[df["Curso"] == curso].copy()
                evals_validas_curso = _filtrar_evals_validas(df_curso, evals_validas)

                ws2 = wb.add_worksheet("Diagnóstico")
                writer.sheets["Diagnóstico"] = ws2
                ws2.write(0, 0, f"Curso: {curso}", wb.add_format({"bold": True}))

                if not evals_validas_curso:
                    ws2.write(2, 0, "No hay evaluaciones válidas (>0).", fmt_cell)
                else:
                    df_ev = compute_aprobados_por_evaluacion(df, curso, evals_validas_curso)
                    for c, col in enumerate(df_ev.columns):
                        ws2.write(2, c, col, fmt_header)
                    for r in range(len(df_ev)):
                        for c in range(len(df_ev.columns)):
                            ws2.write(3 + r, c, df_ev.iat[r, c], fmt_cell)

            # 2.b Comparación (>=2 cursos)
            elif cursos_sel and len(cursos_sel) > 1:
                df_cmp = compute_aprobados_eval_por_curso(df, cursos_sel, evals_validas)

                ws3 = wb.add_worksheet("Comparación")
                writer.sheets["Comparación"] = ws3

                for c, col in enumerate(df_cmp.columns):
                    ws3.write(0, c, col, fmt_header)
                for r in range(len(df_cmp)):
                    for c in range(len(df_cmp.columns)):
                        ws3.write(r + 1, c, df_cmp.iat[r, c], fmt_cell)

                pivot_ap = df_cmp.pivot_table(
                    index="Curso", columns="Evaluación", values="% Aprobados",
                    aggfunc="mean", fill_value=0.0
                ).reset_index()

                cols_ord = ["Curso"] + [ev for ev in ORDER_EVAL if ev in pivot_ap.columns]
                pivot_ap = pivot_ap[cols_ord]

                ws4 = wb.add_worksheet("Pivot %Aprob")
                writer.sheets["Pivot %Aprob"] = ws4
                for c, col in enumerate(pivot_ap.columns):
                    ws4.write(0, c, col, fmt_header)
                for r in range(len(pivot_ap)):
                    for c in range(len(pivot_ap.columns)):
                        val = pivot_ap.iat[r, c]
                        ws4.write(r + 1, c, float(val) if c > 0 else val, fmt_cell)

                pivot_pr = df_cmp.pivot_table(
                    index="Curso", columns="Evaluación", values="Promedio",
                    aggfunc="mean", fill_value=0.0
                ).reset_index()

                cols_ord2 = ["Curso"] + [ev for ev in ORDER_EVAL if ev in pivot_pr.columns]
                pivot_pr = pivot_pr[cols_ord2]

                ws5 = wb.add_worksheet("Pivot Promedio")
                writer.sheets["Pivot Promedio"] = ws5
                for c, col in enumerate(pivot_pr.columns):
                    ws5.write(0, c, col, fmt_header)
                for r in range(len(pivot_pr)):
                    for c in range(len(pivot_pr.columns)):
                        val = pivot_pr.iat[r, c]
                        ws5.write(r + 1, c, float(val) if c > 0 else val, fmt_cell)

    return xlsx


# ======================================================
# FUNCIONES COMPATIBLES CON TU GUI (mantener nombres)
# ======================================================
def exportar_pdf_final(df, cursos_sel=None, carpeta_destino=None):
    return exportar_pdf(df=df, cursos_sel=cursos_sel, evals_sel=None, carpeta_destino=carpeta_destino)


def exportar_pdf_evaluaciones(df, evals_sel, cursos_sel=None, carpeta_destino=None):
    return exportar_pdf(df=df, cursos_sel=cursos_sel, evals_sel=evals_sel, carpeta_destino=carpeta_destino)


def exportar_excel_global(df, evals_sel=None, modo_final=True, carpeta_destino=None):
    """
    Nota: la GUI usualmente ya filtra df por cursos_sel antes de llamar,
    pero aquí igual soportamos si df viene completo (cursos_sel no se pasa).
    """
    if modo_final:
        return exportar_excel(df=df, cursos_sel=None, evals_sel=None, carpeta_destino=carpeta_destino)
    else:
        return exportar_excel(df=df, cursos_sel=None, evals_sel=evals_sel, carpeta_destino=carpeta_destino)
