# 3.Analisis_carrera.py
# -*- coding: utf-8 -*-
"""
SCRIPT 3 — ANÁLISIS POR CARRERA (UCSUR) — Versión B Compacta con Selector de Carpeta

AJUSTES (enero 2026):
- % Aprobados y % Desaprobados se calculan SOLO sobre quienes SÍ rindieron (nota > 0).
  Los "No rindieron" se reportan aparte.
- Situación Final:
  * EstadoFinal se mantiene: si todas las evaluaciones están en 0 => "No rindió (todas 0)".
  * Para porcentajes finales, se descarta del denominador a los que no rindieron nada.
- Nuevo cálculo del promedio final:
  FINAL = 0.18*R(EC1) + 0.20*R(EP) + 0.18*R(EC2) + 0.19*R(EC3) + 0.2*0.18*R(EF)
  con redondeo half-up (x.5 -> x+1).
- Aprobado: nota >= 12.5 (en final y en evaluaciones).
- En cada informe se separa: Total, No rindieron, Rindieron y luego porcentajes.
"""

import os, sys, unicodedata, tempfile, warnings
from datetime import datetime

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import xlsxwriter

# Selector de carpeta
try:
    import tkinter as tk
    from tkinter import filedialog
except:
    tk = None
    filedialog = None

warnings.filterwarnings("ignore", category=FutureWarning)

# ==========================
# CONFIGURACIÓN GENERAL
# ==========================
PARQUET_IN = "notas_filtradas_ucsur.parquet"

EVALS = ["ED", "EC1", "EP", "EC2", "EC3", "EF"]
WEIGHTS = {"ED": 0.0, "EC1": 0.18, "EP": 0.20, "EC2": 0.18, "EC3": 0.19, "EF": 0.25}
FINAL_NAME = "FINAL"

# Orden fijo deseado en todos los reportes:
EVAL_ORDER = EVALS + [FINAL_NAME]

EVAL_NAMES = {
    "ED":   "Evaluación Diagnóstica",
    "EC1":  "Evaluación Continua 1",
    "EP":   "Evaluación Parcial",
    "EC2":  "Evaluación Continua 2",
    "EC3":  "Evaluación Continua 3",
    "EF":   "Evaluación Final",
    FINAL_NAME: "Situación Final",
}

COL_UCSUR_AZUL = "#003B70"
LOGO_PATH = "logo_ucsur.png"

APROBADO_MIN = 12.5  # umbral único solicitado

# Pesos finales según tu fórmula EXACTA (nota: EF usa 0.2*0.18)
FINAL_WEIGHTS = {
    "EC1": 0.18,
    "EP":  0.20,
    "EC2": 0.18,
    "EC3": 0.19,
    "EF":  0.2 * 0.18,
}

def nombre_eval(ev):
    return EVAL_NAMES.get(ev, ev)

# ==========================
# UTILIDADES
# ==========================
def safe_txt(s):
    return str(s).replace("—","-").replace("\u2013", "-") if s is not None else ""

def slug(s):
    s2 = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s2 = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in s2.strip())
    return s2 or "reporte"

def elegir_carpeta():
    """Selector universal de carpeta."""
    if tk is None or filedialog is None:
        return os.getcwd()
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        carpeta = filedialog.askdirectory(title="Seleccione carpeta destino")
        root.destroy()
        if not carpeta:
            print("⚠️ No seleccionó carpeta.")
            return None
        return carpeta
    except:
        return os.getcwd()

def ensure_columns(df):
    cols = ["Curso","Seccion","Carrera","Docente","Evaluacion","Nota","CodigoAlumno","Alumno"]
    for c in cols:
        if c not in df.columns:
            df[c] = "" if c!="Nota" else 0.0
    return df

def round_half_up(x):
    """Redondeo al entero más cercano y .5 hacia arriba (x.5 -> x+1)."""
    try:
        x = float(x)
    except Exception:
        return 0
    return int(np.floor(x + 0.5)) if x >= 0 else int(np.ceil(x - 0.5))

# ==========================
# FUENTE Y ENCABEZADO PDF
# ==========================
def cargar_fuente(pdf):
    try:
        arial = r"C:\Windows\Fonts\arial.ttf"
        arial_bold = r"C:\Windows\Fonts\arialbd.ttf"
        if os.path.exists(arial):
            pdf.add_font("ArialUnicode","",arial)
        if os.path.exists(arial_bold):
            pdf.add_font("ArialUnicode","B",arial_bold)
        return "ArialUnicode"
    except:
        return "Helvetica"

def pdf_encabezado(pdf, titulo):
    if os.path.exists(LOGO_PATH):
        try:
            pdf.image(LOGO_PATH, x=10, y=8, w=28)
        except:
            pass
    pdf.set_xy(10,12)
    pdf.set_font(pdf.font_family,"B",14)
    pdf.cell(0,8,"UNIVERSIDAD CIENTÍFICA DEL SUR",align="C",ln=1)
    pdf.set_font(pdf.font_family,"",11)
    pdf.cell(0,6,"Departamento de Cursos Básicos",align="C",ln=1)
    pdf.cell(0,6,safe_txt(titulo),align="C",ln=1)
    pdf.set_y(42)

# ==========================
# CARGA
# ==========================
def cargar_df(path):
    if not os.path.exists(path):
        print("❌ Archivo parquet no encontrado.")
        sys.exit()
    df = pd.read_parquet(path)
    df = ensure_columns(df)
    for c in ["Curso","Seccion","Carrera","Docente","Evaluacion"]:
        df[c] = df[c].astype(str).str.strip()
    df["Evaluacion"] = df["Evaluacion"].str.upper()
    df["Nota"] = pd.to_numeric(df["Nota"],errors="coerce").fillna(0.0)
    return df

# ==========================
# FINAL
# ==========================
def pivot_final(df_sc):
    piv = df_sc.pivot_table(
        index=["CodigoAlumno","Alumno","Curso","Seccion","Carrera","Docente"],
        columns="Evaluacion",
        values="Nota",
        aggfunc="max",
        fill_value=0
    ).reset_index()

    for e in EVALS:
        if e not in piv.columns:
            piv[e] = 0.0

    # Nuevo FINAL con redondeo half-up por evaluación
    for ev in ["EC1","EP","EC2","EC3","EF"]:
        piv[f"__R_{ev}"] = piv[ev].map(round_half_up)

    piv[FINAL_NAME] = (
        FINAL_WEIGHTS["EC1"] * piv["__R_EC1"] +
        FINAL_WEIGHTS["EP"]  * piv["__R_EP"]  +
        FINAL_WEIGHTS["EC2"] * piv["__R_EC2"] +
        FINAL_WEIGHTS["EC3"] * piv["__R_EC3"] +
        FINAL_WEIGHTS["EF"]  * piv["__R_EF"]
    )

    piv = piv.drop(columns=[c for c in piv.columns if c.startswith("__R_")])

    def est(r):
        if all(float(r[e]) == 0.0 for e in EVALS):
            return "No rindió (todas 0)"
        return "Aprobado" if float(r[FINAL_NAME]) >= APROBADO_MIN else "Desaprobado"

    piv["EstadoFinal"] = piv.apply(est,axis=1)
    return piv

def ef_valida(df_sc):
    sub = df_sc[df_sc["Evaluacion"]=="EF"]
    return not sub.empty and (sub["Nota"]>0).any()

# ==========================
# RESUMEN GLOBAL
# ==========================
def resumen_global(df_sc):
    """
    Devuelve DataFrame con columnas:
    CodigoEval | Evaluación | Total | No rindieron | Rindieron |
    % Aprobados | % Desaprobados | % No rindieron | Promedio

    IMPORTANTES:
    - % Aprobados/% Desaprobados se calculan sobre Rindieron.
    - % No rindieron se calcula sobre Total.
    - Situación Final: se descarta del denominador a los que no rindieron nada.
    - Aprobado: >= 12.5.
    """
    filas = []
    ef_ok = ef_valida(df_sc)

    eval_loop = EVALS + ([FINAL_NAME] if ef_ok else [])

    piv_final = pivot_final(df_sc) if ef_ok else None

    for ev in eval_loop:
        if ev == FINAL_NAME:
            dfv = piv_final
            if dfv is None or dfv.empty:
                continue

            total = int(len(dfv))
            no_r = int((dfv["EstadoFinal"]=="No rindió (todas 0)").sum())
            rindieron = int(total - no_r)

            if total == 0 or rindieron == 0:
                continue

            rind = dfv[dfv["EstadoFinal"]!="No rindió (todas 0)"][FINAL_NAME]
            aprob = int((rind >= APROBADO_MIN).sum())
            desap = int((rind < APROBADO_MIN).sum())
            prom = float(rind.mean()) if rindieron else 0.0

            pct_ap = round(aprob / rindieron * 100, 1) if rindieron else 0.0
            pct_de = round(desap / rindieron * 100, 1) if rindieron else 0.0
            pct_nr = round(no_r / total * 100, 1) if total else 0.0

        else:
            sub = df_sc[df_sc["Evaluacion"]==ev]
            if sub.empty:
                continue

            total = int(len(sub))
            no_r = int((sub["Nota"]==0).sum())
            rind_series = sub[sub["Nota"]>0]["Nota"]
            rindieron = int(len(rind_series))

            if total == 0 or rindieron == 0:
                continue

            aprob = int((rind_series >= APROBADO_MIN).sum())
            desap = int((rind_series < APROBADO_MIN).sum())
            prom = float(rind_series.mean()) if rindieron else 0.0

            pct_ap = round(aprob / rindieron * 100, 1) if rindieron else 0.0
            pct_de = round(desap / rindieron * 100, 1) if rindieron else 0.0
            pct_nr = round(no_r / total * 100, 1) if total else 0.0

        filas.append({
            "CodigoEval": ev,
            "Evaluación": nombre_eval(ev),
            "Total": total,
            "No rindieron": no_r,
            "Rindieron": rindieron,
            "% Aprobados": pct_ap,
            "% Desaprobados": pct_de,
            "% No rindieron": pct_nr,
            "Promedio": round(prom,2)
        })

    if not filas:
        return pd.DataFrame()

    df_gl = pd.DataFrame(filas)

    def _orden(ev_code):
        return EVAL_ORDER.index(ev_code) if ev_code in EVAL_ORDER else 999

    df_gl["__ord"] = df_gl["CodigoEval"].apply(_orden)
    df_gl = df_gl.sort_values("__ord").drop(columns="__ord").reset_index(drop=True)
    return df_gl

# ==========================
# GRÁFICOS
# ==========================
def graf_bar(df,label,val,tit,y,out):
    if df.empty:
        return
    plt.figure(figsize=(6,3))
    plt.bar(df[label],df[val])
    plt.title(tit)
    plt.ylabel(y)
    plt.xticks(rotation=30,ha="right")
    plt.tight_layout()
    plt.savefig(out,dpi=150)
    plt.close()

def graf_3(labels,values,title,out):
    plt.figure(figsize=(4,3))
    plt.bar(labels,values)
    plt.title(title)
    plt.ylabel("Porcentaje (%)")
    plt.tight_layout()
    plt.savefig(out,dpi=150)
    plt.close()

# ==========================
# PDF
# ==========================
def pdf_tabla(pdf,fam,headers,rows,widths):
    def head():
        pdf.set_fill_color(0,59,112)
        pdf.set_text_color(255,255,255)
        pdf.set_font(fam,"B",9)
        for h,w in zip(headers,widths):
            pdf.cell(w,7,str(h),1,0,"C",1)
        pdf.ln()
        pdf.set_text_color(0,0,0)
        pdf.set_font(fam,"",9)

    head()
    for r in rows:
        if pdf.get_y()>260:
            pdf.add_page()
            head()
        for v,w in zip(r,widths):
            pdf.cell(w,6,str(v),1,0,"C")
        pdf.ln()

def exportar_pdf(df_sc,carrera,curso,carpeta):
    df_gl = resumen_global(df_sc)
    if df_gl.empty:
        print("⚠️ No hay registros.")
        return

    if not carpeta:
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(carpeta,f"Carrera_{slug(carrera)}_Curso_{slug(curso)}_{ts}.pdf")
    tmp = tempfile.mkdtemp()

    g1 = os.path.join(tmp,"bar_pct.png")
    graf_bar(df_gl,"Evaluación","% Aprobados","% Aprobados por evaluación","% Aprobados (sobre rindieron)",g1)

    g2 = os.path.join(tmp,"bar_prom.png")
    graf_bar(df_gl,"Evaluación","Promedio","Promedio por evaluación","Promedio",g2)

    pdf = FPDF()
    pdf.add_page()
    fam = cargar_fuente(pdf)
    pdf.set_font(fam,"",10)

    pdf_encabezado(pdf,f"Análisis por Carrera — {carrera} — {curso}")
    pdf.cell(0,6,f"Carrera: {carrera}",new_x=XPos.LMARGIN,new_y=YPos.NEXT)
    pdf.cell(0,6,f"Curso: {curso}",new_x=XPos.LMARGIN,new_y=YPos.NEXT)
    pdf.cell(0,6,f"Fecha: {datetime.now().strftime('%d/%m/%Y')}",new_x=XPos.LMARGIN,new_y=YPos.NEXT)
    pdf.ln(4)

    pdf.set_font(fam,"B",11)
    pdf.cell(0,7,"Resumen global por evaluación",ln=1)

    headers=["Evaluación","Total","No rind.","Rind.","% Aprob.","% Desaprob.","% No rind.","Prom."]
    widths=[36,14,16,14,18,20,18,14]

    rows=[]
    for _,r in df_gl.iterrows():
        rows.append([
            r["Evaluación"],
            int(r["Total"]),
            int(r["No rindieron"]),
            int(r["Rindieron"]),
            f"{r['% Aprobados']}%",
            f"{r['% Desaprobados']}%",
            f"{r['% No rindieron']}%",
            r["Promedio"]
        ])
    pdf_tabla(pdf,fam,headers,rows,widths)

    pdf.ln(4)
    pdf.set_font(fam,"B",10)
    pdf.cell(0,6,"Gráfico: % Aprobados por evaluación (sobre rindieron)",ln=1)
    pdf.image(g1,w=170)
    pdf.ln(4)
    pdf.cell(0,6,"Gráfico: Promedio por evaluación",ln=1)
    pdf.image(g2,w=170)

    for _,r in df_gl.iterrows():
        ev  = r["CodigoEval"]
        evn = r["Evaluación"]

        labels = ["% Aprob.","% Desaprob.","% No rind."]
        values = [r["% Aprobados"],r["% Desaprobados"],r["% No rindieron"]]

        gdet = os.path.join(tmp,f"det_{ev}.png")
        graf_3(labels,values,f"{evn} — Distribución porcentual",gdet)

        pdf.add_page()
        pdf_encabezado(pdf,f"Detalle — {evn}")

        pdf.set_font(fam,"",10)
        pdf.cell(0,6,f"Carrera: {carrera}",new_x=XPos.LMARGIN,new_y=YPos.NEXT)
        pdf.cell(0,6,f"Curso: {curso}",new_x=XPos.LMARGIN,new_y=YPos.NEXT)
        pdf.cell(0,6,evn,new_x=XPos.LMARGIN,new_y=YPos.NEXT)
        pdf.ln(4)

        headers2=["Total","No rind.","Rind.","% Aprob.","% Desaprob.","% No rind.","Prom."]
        widths2=[18,18,16,20,22,20,16]
        row2=[[
            int(r["Total"]),
            int(r["No rindieron"]),
            int(r["Rindieron"]),
            f"{r['% Aprobados']}%",
            f"{r['% Desaprobados']}%",
            f"{r['% No rindieron']}%",
            r["Promedio"]
        ]]
        pdf_tabla(pdf,fam,headers2,row2,widths2)

        pdf.ln(4)
        pdf.cell(0,6,"Gráfico comparativo",ln=1)
        pdf.image(gdet,w=120)

    try:
        pdf.output(path)
        print(f"✅ PDF generado: {path}")
        if os.name=="nt":
            os.startfile(path)
    except Exception as e:
        print(f"⚠️ Error PDF: {e}")

# ==========================
# EXCEL
# ==========================
def autosize(ws,df):
    for j,c in enumerate(df.columns):
        mx = max(len(str(c)), max(df[c].astype(str).apply(len)))
        ws.set_column(j,j,max(10,min(mx+2,38)))

def write(ws,df,fmt_h,fmt_c,row=0,col=0):
    for j,c in enumerate(df.columns):
        ws.write(row,col+j,c,fmt_h)
    for i in range(len(df)):
        for j in range(len(df.columns)):
            ws.write(row+1+i,col+j,df.iat[i,j],fmt_c)

def exportar_excel(df_sc,carrera,curso,carpeta):
    df_gl = resumen_global(df_sc)
    if df_gl.empty:
        print("⚠️ No hay datos.")
        return
    if not carpeta:
        return

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(carpeta,f"Carrera_{slug(carrera)}_Curso_{slug(curso)}_{ts}.xlsx")

    with pd.ExcelWriter(path,engine="xlsxwriter") as writer:
        wb  = writer.book
        fmt_h = wb.add_format({
            "bold":True,
            "font_color":"white",
            "bg_color":COL_UCSUR_AZUL,
            "border":1,
            "align":"center"
        })
        fmt_c = wb.add_format({"border":1})
        fmt_t = wb.add_format({"bold":True,"font_size":14})

        sh = "Global"
        ws = wb.add_worksheet(sh)
        ws.write(0,0,f"Carrera {carrera} — Curso {curso}",fmt_t)

        df_excel = df_gl[["Evaluación","Total","No rindieron","Rindieron","% Aprobados","% Desaprobados","% No rindieron","Promedio"]].copy()
        df_excel.columns = ["Evaluación","Total","No rind.","Rind.","% Aprob.","% Desaprob.","% No rind.","Prom."]
        df_excel = df_excel.replace([np.nan,np.inf,-np.inf],"")

        write(ws,df_excel,fmt_h,fmt_c,2,0)
        autosize(ws,df_excel)

        ch = wb.add_chart({"type":"column"})
        ch.add_series({
            "name":"% Aprobados (sobre rindieron)",
            "categories":[sh,3,0,2+len(df_excel),0],
            "values":[sh,3,4,2+len(df_excel),4],
        })
        ch.set_title({"name":"% Aprobados por evaluación (sobre rindieron)"})
        ws.insert_chart(2,len(df_excel.columns)+2,ch)

        for _,r in df_gl.iterrows():
            ev  = r["CodigoEval"]
            evn = r["Evaluación"]
            df_ev = pd.DataFrame([{
                "Total":int(r["Total"]),
                "No rind.":int(r["No rindieron"]),
                "Rind.":int(r["Rindieron"]),
                "% Aprob.":r["% Aprobados"],
                "% Desaprob.":r["% Desaprobados"],
                "% No rind.":r["% No rindieron"],
                "Prom.":r["Promedio"],
            }]).replace([np.nan,np.inf,-np.inf],"")

            sh_ev = f"EV_{ev}"[:31]
            ws_ev = wb.add_worksheet(sh_ev)
            ws_ev.write(0,0,f"{carrera} — {curso} — {evn}",fmt_t)
            write(ws_ev,df_ev,fmt_h,fmt_c,2,0)
            autosize(ws_ev,df_ev)

    print(f"✅ Excel generado: {path}")

# ==========================
# MENÚS
# ==========================
def menu_exportar(df_sc,carrera,curso):
    while True:
        print("\n───────────────")
        print("1. Exportar PDF")
        print("2. Exportar Excel")
        print("3. Volver")
        op = input("Seleccione opción: ").strip()
        if op=="1":
            carpeta = elegir_carpeta()
            if carpeta:
                exportar_pdf(df_sc,carrera,curso,carpeta)
        elif op=="2":
            carpeta = elegir_carpeta()
            if carpeta:
                exportar_excel(df_sc,carrera,curso,carpeta)
        elif op=="3":
            break
        else:
            print("⚠️ Opción inválida.")

def submenu_carrera(df,carrera):
    df_car = df[df["Carrera"].str.upper()==carrera.upper()]
    if df_car.empty:
        print("⚠️ No hay datos.")
        return

    cursos = sorted(df_car["Curso"].unique())

    while True:
        print(f"\n📘 Carrera: {carrera}")
        for i,c in enumerate(cursos,1):
            print(f"{i}. {c}")
        print(f"{len(cursos)+1}. Volver")

        op = input("Seleccione curso: ")
        if not op.isdigit():
            continue
        op = int(op)

        if op == len(cursos)+1:
            break
        if not (1<=op<=len(cursos)):
            continue

        curso = cursos[op-1]
        df_sc = df_car[df_car["Curso"]==curso]
        print(f"\n➡ Carrera {carrera} — Curso {curso}")
        menu_exportar(df_sc,carrera,curso)

def menu_principal():
    df = cargar_df(PARQUET_IN)
    carreras = sorted(df["Carrera"].unique())

    while True:
        print("\n==============================")
        print(" ANÁLISIS POR CARRERA — UCSUR")
        print("==============================")
        print("1. Elegir carrera")
        print("2. Salir")

        op = input("Seleccione opción: ")
        if op=="1":
            for i,c in enumerate(carreras,1):
                print(f"{i}. {c}")
            sel = input("Seleccione carrera: ")
            if not sel.isdigit():
                continue
            sel = int(sel)
            if not (1<=sel<=len(carreras)):
                continue
            carrera = carreras[sel-1]
            submenu_carrera(df,carrera)
        elif op=="2":
            break
        else:
            print("⚠️ Opción inválida.")

# ==========================
# MAIN
# ==========================
if __name__ == "__main__":
    menu_principal()
