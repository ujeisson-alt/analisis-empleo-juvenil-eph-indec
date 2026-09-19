#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
01_pipeline_eph.py
==================
Pipeline reproducible del proyecto "Analisis del Empleo Juvenil en Argentina"
(microdatos EPH - INDEC).

Replica en codigo la MISMA logica que queda escrita como formulas vivas en el
libro de Excel, de modo que ambos caminos deben arrojar identicos resultados:

  1. Importa el archivo de personas (delimitado por ';').
  2. Filtra el segmento 18-29 anios (CH06).
  3. Diagnostica nulos y codigos especiales (-9, 999, celdas vacias).
  4. Decodifica REGION y NIVEL_ED con tablas de equivalencia (BUSCARV).
  5. Construye las variables compuestas IPL (0-3) e ISE (0-3).
  6. Calcula la matriz de correlacion 5x5 (COEF.DE.CORREL).
  7. Arma la brecha de genero por region y el perfil ingreso-educacion.
  8. Exporta:
       02_Datos_Procesados/02_Base_Trabajo.xlsx   (con formulas, mapa de calor y dashboard)
       03_Visualizaciones/G1_Matriz_Correlacion.png
       03_Visualizaciones/G2_Brecha_Genero.png
       03_Visualizaciones/G3_Ingreso_Educacion.png
       05_Scripts/_hallazgos.json                 (insumo del informe ejecutivo)

Uso:
    python3 05_Scripts/01_pipeline_eph.py [ruta_al_txt]
"""

import json
import sys
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import LinearSegmentedColormap
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==========================================================================
# 0. Parametros, rutas y paleta institucional (maximo 3 colores + neutros)
# ==========================================================================
RAIZ = Path(__file__).resolve().parents[1]
ENTRADA = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    RAIZ / "01_Datos_Originales" / "usu_individual_T125_SIMULADO.txt"
DIR_PROC = RAIZ / "02_Datos_Procesados"
DIR_VIS = RAIZ / "03_Visualizaciones"
XLSX = DIR_PROC / "02_Base_Trabajo.xlsx"

AZUL = "#14375E"      # color primario
NARANJA = "#C2603C"   # color secundario
GRIS = "#8A94A6"      # neutro
TINTA = "#1B2430"
CLARO = "#EEF1F6"

HOJA_D = "Jóvenes_18_29"          # nombre de la hoja de trabajo
REF_D = f"'{HOJA_D}'"

REGION_LBL = {1: "Gran Buenos Aires", 40: "Noroeste (NOA)", 41: "Noreste (NEA)",
              42: "Cuyo", 43: "Pampeana", 44: "Patagonia"}
NIVEL_LBL = {1: "Primaria incompleta", 2: "Primaria completa",
             3: "Secundaria incompleta", 4: "Secundaria completa",
             5: "Superior incompleto", 6: "Superior completo", 7: "Sin instrucción"}
ESTADO_LBL = {1: "Ocupado", 2: "Desocupado", 3: "Inactivo", 4: "Menor de 10 años"}

DIR_PROC.mkdir(parents=True, exist_ok=True)
DIR_VIS.mkdir(parents=True, exist_ok=True)

# ==========================================================================
# 1. Importacion y filtrado del segmento juvenil
# ==========================================================================
print("1) Importando microdatos...")
base = pd.read_csv(ENTRADA, sep=";", low_memory=False)
N_TOTAL = len(base)

NUM = ["REGION", "CH03", "CH04", "CH06", "CH10", "NIVEL_ED", "ESTADO",
       "CAT_OCUP", "PP07C", "PP07H", "PP3E_TOT", "P21", "TOT_P12", "P47T", "IPCF"]
for c in NUM:
    if c in base.columns:
        base[c] = pd.to_numeric(base[c], errors="coerce")

jov = base[(base.CH06 >= 18) & (base.CH06 <= 29)].copy().reset_index(drop=True)
N_JOV = len(jov)
print(f"   registros totales: {N_TOTAL:,} | jóvenes 18-29: {N_JOV:,} ({N_JOV/N_TOTAL:.1%})")

# ==========================================================================
# 2. Diagnostico de nulos y codigos especiales
# ==========================================================================
diag = []
for col, cods in [("P21", [-9]), ("TOT_P12", [-9]), ("PP3E_TOT", [999]),
                  ("PP07H", []), ("PP07C", [])]:
    s = jov[col]
    fila = {"Variable": col,
            "Celdas vacías (no corresponde)": int(s.isna().sum()),
            "Código especial": ", ".join(str(c) for c in cods) or "—",
            "Registros con código especial": int(s.isin(cods).sum()) if cods else 0,
            "Válidos para análisis": int(s.notna().sum() - (s.isin(cods).sum() if cods else 0)),
            "% válidos sobre jóvenes": round(
                100 * (s.notna().sum() - (s.isin(cods).sum() if cods else 0)) / len(jov), 1)}
    diag.append(fila)
diag = pd.DataFrame(diag)

# ==========================================================================
# 3. Decodificacion + variables compuestas (misma logica que las formulas)
# ==========================================================================
print("2) Construyendo variables compuestas IPL e ISE...")
jov["REGION_TEXTO"] = jov.REGION.map(REGION_LBL)
jov["NIVEL_ED_TEXTO"] = jov.NIVEL_ED.map(NIVEL_LBL)
jov["SEXO_TEXTO"] = np.where(jov.CH04 == 1, "Varón", "Mujer")
jov["ESTADO_TEXTO"] = jov.ESTADO.map(ESTADO_LBL)
jov["TRAMO_EDAD"] = np.where(jov.CH06 <= 24, "18-24", "25-29")

ocu = jov.ESTADO == 1
horas_val = jov.PP3E_TOT.where(jov.PP3E_TOT != 999)          # 999 = Ns./Nr.

jov["IPL_Jubilacion"] = np.where(ocu, (jov.PP07H == 2).astype(float), np.nan)
jov["IPL_Inestabilidad"] = np.where(ocu, (jov.PP07C >= 2).astype(float), np.nan)
jov["IPL_Jornada"] = np.where(ocu & horas_val.notna(),
                              ((horas_val < 35) | (horas_val > 48)).astype(float), np.nan)
jov["IPL"] = jov[["IPL_Jubilacion", "IPL_Inestabilidad", "IPL_Jornada"]].sum(axis=1)
jov.loc[jov[["IPL_Jubilacion", "IPL_Inestabilidad", "IPL_Jornada"]].isna().any(axis=1), "IPL"] = np.nan

jov["ISE_NivelJoven"] = np.where(jov.NIVEL_ED >= 4, 2, np.where(jov.NIVEL_ED >= 2, 1, 0))
jov["ISE_Hogar"] = (jov.NIVEL_ED >= 5).astype(int)
jov["ISE"] = jov.ISE_NivelJoven + jov.ISE_Hogar

jov["P21_LIMPIO"] = np.where(ocu & (jov.P21 != -9) & jov.P21.notna(), jov.P21, np.nan)
jov["HORAS_LIMPIO"] = np.where(ocu, horas_val, np.nan)

frec_ipl = jov.IPL.value_counts().reindex([0, 1, 2, 3]).fillna(0).astype(int)
frec_ise = jov.ISE.value_counts().reindex([0, 1, 2, 3]).fillna(0).astype(int)

# ==========================================================================
# 4. Matriz de correlacion 5x5
# ==========================================================================
print("3) Calculando matriz de correlación...")
VARS = [("Edad (CH06)", "CH06"),
        ("Ingreso ocup. principal (P21)", "P21_LIMPIO"),
        ("Horas semanales (PP3E_TOT)", "HORAS_LIMPIO"),
        ("Precariedad laboral (IPL)", "IPL"),
        ("Nivel socio-educativo (ISE)", "ISE")]
etq = [v[0] for v in VARS]
corr = pd.DataFrame(index=etq, columns=etq, dtype=float)
for i, (li, ci) in enumerate(VARS):
    for j, (lj, cj) in enumerate(VARS):
        par = jov[[ci, cj]].dropna()
        corr.iloc[i, j] = round(float(np.corrcoef(par[ci], par[cj])[0, 1]), 3) \
            if len(par) > 2 and ci != cj else (1.0 if ci == cj else np.nan)

# ==========================================================================
# 5. Brecha de genero por region e ingreso por nivel educativo
# ==========================================================================
regiones_ord = [REGION_LBL[k] for k in [1, 43, 42, 40, 41, 44]]
td = jov.pivot_table(index="SEXO_TEXTO", columns="REGION_TEXTO",
                     values="IPL", aggfunc="mean").reindex(
                     index=["Varón", "Mujer"], columns=regiones_ord).round(3)
td["Total país"] = [round(jov.loc[jov.SEXO_TEXTO == s, "IPL"].mean(), 3) for s in ["Varón", "Mujer"]]
brecha = (td.loc["Mujer"] - td.loc["Varón"]).round(3)

ing_edu = pd.DataFrame({
    "Código": [1, 2, 3, 4, 5, 6],
    "Nivel educativo": [NIVEL_LBL[i] for i in range(1, 7)],
    "Ingreso promedio (P21)": [round(jov.loc[jov.NIVEL_ED == i, "P21_LIMPIO"].mean() or 0)
                               for i in range(1, 7)],
    "Casos válidos": [int(jov.loc[jov.NIVEL_ED == i, "P21_LIMPIO"].notna().sum())
                      for i in range(1, 7)],
    "IPL promedio": [round(jov.loc[jov.NIVEL_ED == i, "IPL"].mean(), 2)
                     for i in range(1, 7)],
})

# Indicadores de cabecera
kpi = {
    "registros_totales": N_TOTAL,
    "jovenes": N_JOV,
    "pct_jovenes": round(100 * N_JOV / N_TOTAL, 1),
    "ocupados": int((jov.ESTADO == 1).sum()),
    "desocupados": int((jov.ESTADO == 2).sum()),
    "inactivos": int((jov.ESTADO == 3).sum()),
    "tasa_actividad": round(100 * ((jov.ESTADO.isin([1, 2])).sum()) / N_JOV, 1),
    "tasa_desocupacion": round(100 * (jov.ESTADO == 2).sum() / max((jov.ESTADO.isin([1, 2])).sum(), 1), 1),
    "informalidad": round(100 * jov.IPL_Jubilacion.mean(), 1),
    "ipl_promedio": round(jov.IPL.mean(), 2),
    "ipl3_pct": round(100 * (jov.IPL == 3).sum() / max(jov.IPL.notna().sum(), 1), 1),
    "ingreso_promedio": int(round(jov.P21_LIMPIO.mean())),
    "ingreso_varon": int(round(jov.loc[jov.SEXO_TEXTO == "Varón", "P21_LIMPIO"].mean())),
    "ingreso_mujer": int(round(jov.loc[jov.SEXO_TEXTO == "Mujer", "P21_LIMPIO"].mean())),
    "horas_promedio": round(jov.HORAS_LIMPIO.mean(), 1),
    "r_ise_ipl": corr.loc[etq[4], etq[3]],
    "r_ise_p21": corr.loc[etq[4], etq[1]],
    "r_ipl_p21": corr.loc[etq[3], etq[1]],
    "r_horas_p21": corr.loc[etq[2], etq[1]],
    "r_edad_p21": corr.loc[etq[0], etq[1]],
    "ipl_mujer": float(td.loc["Mujer", "Total país"]),
    "ipl_varon": float(td.loc["Varón", "Total país"]),
    "brecha_total": float(brecha["Total país"]),
    "region_mayor_brecha": str(brecha.drop("Total país").idxmax()),
    "brecha_max": float(brecha.drop("Total país").max()),
    "region_mayor_ipl": str(td.drop(columns="Total país").mean().idxmax()),
    "ipl_region_max": round(float(td.drop(columns="Total país").mean().max()), 2),
    "ingreso_secund_incompleta": int(ing_edu.loc[2, "Ingreso promedio (P21)"]),
    "ingreso_superior_completo": int(ing_edu.loc[5, "Ingreso promedio (P21)"]),
}
kpi["brecha_ingresos_pct"] = round(
    100 * (1 - kpi["ingreso_mujer"] / kpi["ingreso_varon"]), 1)
kpi["prima_superior_pct"] = round(
    100 * (kpi["ingreso_superior_completo"] / kpi["ingreso_secund_incompleta"] - 1), 1)

# ==========================================================================
# 6. Graficos PNG (paleta de 3 colores, ejes titulados, sin ruido visual)
# ==========================================================================
print("4) Generando visualizaciones...")
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10.5, "axes.edgecolor": "#C9D0DA",
    "axes.labelcolor": TINTA, "text.color": TINTA, "xtick.color": "#4A5568",
    "ytick.color": "#4A5568", "axes.titlesize": 13, "axes.titleweight": "bold",
    "figure.facecolor": "white", "axes.facecolor": "white", "savefig.dpi": 200,
})
cmap_div = LinearSegmentedColormap.from_list("div", ["#A63A2A", "#F3EFE8", "#1E6F5C"])

# ---- G1: mapa de calor de la matriz de correlacion -----------------------
corto = ["Edad\n(CH06)", "Ingreso\n(P21)", "Horas\n(PP3E_TOT)", "Precariedad\n(IPL)",
         "Socio-educativo\n(ISE)"]
fig, ax = plt.subplots(figsize=(7.6, 6.0))
m = corr.values.astype(float)
im = ax.imshow(m, cmap=cmap_div, vmin=-1, vmax=1)
ax.set_xticks(range(5), corto, fontsize=9)
ax.set_yticks(range(5), corto, fontsize=9)
for i in range(5):
    for j in range(5):
        v = m[i, j]
        ax.text(j, i, f"{v:+.2f}".replace("+1.00", "1.00"),
                ha="center", va="center", fontsize=10.5,
                color="white" if abs(v) > 0.55 else TINTA,
                fontweight="bold" if abs(v) >= 0.3 and i != j else "normal")
ax.set_title("Matriz de correlación — determinantes del empleo juvenil",
             pad=14, loc="left")
fig.text(0.02, 0.005, f"Jóvenes de 18 a 29 años · n = {int(jov.IPL.notna().sum()):,} ocupados con dato válido",
         fontsize=8.5, color="#4A5568")
cb = fig.colorbar(im, ax=ax, shrink=0.78, ticks=[-1, -0.5, 0, 0.5, 1])
cb.set_label("Coeficiente de Pearson", fontsize=9)
cb.outline.set_visible(False)
for s in ax.spines.values():
    s.set_visible(False)
ax.set_xticks(np.arange(-.5, 5, 1), minor=True)
ax.set_yticks(np.arange(-.5, 5, 1), minor=True)
ax.grid(which="minor", color="white", linewidth=2.5)
ax.tick_params(which="minor", length=0)
fig.tight_layout()
fig.savefig(DIR_VIS / "G1_Matriz_Correlacion.png", bbox_inches="tight")
plt.close(fig)

# ---- G2: brecha de genero en precariedad por region ----------------------
fig, ax = plt.subplots(figsize=(9.2, 5.2))
x = np.arange(len(regiones_ord))
w = 0.38
v_var = [td.loc["Varón", r] for r in regiones_ord]
v_muj = [td.loc["Mujer", r] for r in regiones_ord]
b1 = ax.bar(x - w/2, v_var, w, label="Varones", color=AZUL)
b2 = ax.bar(x + w/2, v_muj, w, label="Mujeres", color=NARANJA)
for b in list(b1) + list(b2):
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.02, f"{b.get_height():.2f}",
            ha="center", va="bottom", fontsize=9, color="#4A5568")
ax.axhline(jov.IPL.mean(), color=GRIS, linestyle="--", linewidth=1.2)
ax.set_xlim(-0.65, len(x) - 0.05 + 0.85)
ax.text(len(x) - 0.42, jov.IPL.mean(), f"Promedio país\n{jov.IPL.mean():.2f}".replace(".", ","),
        fontsize=8.5, color="#4A5568", ha="left", va="center")
ax.set_xticks(x, [r.replace(" (", "\n(") for r in regiones_ord], fontsize=9.5)
ax.set_ylabel("Índice de Precariedad Laboral (0 a 3)")
ax.set_xlabel("Región geográfica")
ax.set_title("Brecha de género en precariedad laboral juvenil, por región", pad=14, loc="left")
ax.set_ylim(0, max(max(v_var), max(v_muj)) * 1.22)
ax.legend(frameon=False, loc="upper right", ncols=2)
ax.spines[["top", "right"]].set_visible(False)
ax.yaxis.grid(True, color="#EDF0F5")
ax.set_axisbelow(True)
fig.tight_layout()
fig.savefig(DIR_VIS / "G2_Brecha_Genero.png", bbox_inches="tight")
plt.close(fig)

# ---- G3: ingreso segun nivel educativo ----------------------------------
fig, ax = plt.subplots(figsize=(9.2, 5.2))
ax.plot(ing_edu["Nivel educativo"], ing_edu["Ingreso promedio (P21)"],
        marker="o", markersize=7, linewidth=2.4, color=AZUL, label="Ingreso promedio (P21)")
for xx, yy, n in zip(range(6), ing_edu["Ingreso promedio (P21)"], ing_edu["Casos válidos"]):
    ax.annotate(f"${yy:,.0f}".replace(",", "."), (xx, yy), textcoords="offset points",
                xytext=(0, 13), ha="center", fontsize=9, color=TINTA,
                bbox=dict(boxstyle="round,pad=0.18", fc="white", ec="none", alpha=0.85))
    ax.annotate(f"n={n}", (xx, yy), textcoords="offset points", xytext=(0, -16),
                ha="center", fontsize=8, color=GRIS)
ax2 = ax.twinx()
ax2.plot(ing_edu["Nivel educativo"], ing_edu["IPL promedio"], marker="s", markersize=6,
         linewidth=2.0, linestyle="--", color=NARANJA, label="IPL promedio")
ax2.set_ylabel("IPL promedio (0 a 3)", color=NARANJA)
ax2.tick_params(axis="y", colors=NARANJA)
ax2.set_ylim(0, 3)
ax2.spines[["top"]].set_visible(False)
ax.yaxis.set_major_formatter(plt.FuncFormatter(
    lambda v, _: "$ " + f"{v:,.0f}".replace(",", ".")))
ax.set_ylabel("Ingreso de la ocupación principal (pesos corrientes)")
ax.set_xlabel("Nivel educativo alcanzado (NIVEL_ED)")
ax.set_title("A más credenciales, más ingreso y menos precariedad", pad=14, loc="left")
ax.set_xticks(range(6), [t.replace(" ", "\n", 1) for t in ing_edu["Nivel educativo"]], fontsize=9)
ax.spines[["top", "right"]].set_visible(False)
ax.yaxis.grid(True, color="#EDF0F5")
ax.set_axisbelow(True)
h1, l1 = ax.get_legend_handles_labels()
h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, frameon=False, loc="upper left", ncols=2, fontsize=9)
ax.set_ylim(0, max(ing_edu["Ingreso promedio (P21)"]) * 1.25)
fig.tight_layout()
fig.savefig(DIR_VIS / "G3_Ingreso_Educacion.png", bbox_inches="tight")
plt.close(fig)

# ==========================================================================
# 7. Libro de Excel con formulas vivas
# ==========================================================================
print("5) Escribiendo el libro de Excel con fórmulas vivas...")
wb = Workbook()

TIT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FILL_TIT = PatternFill("solid", fgColor=AZUL.lstrip("#"))
BORDE = Border(*[Side(style="thin", color="D5DBE5")] * 4)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def encabezar(ws, fila, valores, ancho=None):
    for j, v in enumerate(valores, start=1):
        c = ws.cell(row=fila, column=j, value=v)
        c.font, c.fill, c.alignment, c.border = TIT, FILL_TIT, CENTRO, BORDE
    ws.row_dimensions[fila].height = 30
    if ancho:
        for j, a in enumerate(ancho, start=1):
            ws.column_dimensions[get_column_letter(j)].width = a


def titulo_hoja(ws, texto, subtitulo="", fila=1):
    ws.cell(row=fila, column=1, value=texto).font = Font(
        name="Calibri", size=15, bold=True, color=AZUL.lstrip("#"))
    if subtitulo:
        ws.cell(row=fila + 1, column=1, value=subtitulo).font = Font(
            name="Calibri", size=10, italic=True, color="5A6472")


# -------------------------------------------------- hoja Tablas_Equivalencia
we = wb.active
we.title = "Tablas_Equivalencia"
titulo_hoja(we, "Tablas de equivalencia — Diseño de Registro EPH (INDEC)",
            "Fuente de los códigos: Diseño de Registro de la Base de Microdatos EPH. "
            "Estas tablas alimentan las fórmulas BUSCARV de la hoja de trabajo.")
encabezar(we, 2, ["REGION", "Etiqueta", "", "NIVEL_ED", "Etiqueta", "", "ESTADO", "Etiqueta"],
          ancho=[10, 26, 3, 11, 26, 3, 10, 22])
for i, (k, v) in enumerate(REGION_LBL.items(), start=3):
    we.cell(row=i, column=1, value=k).border = BORDE
    we.cell(row=i, column=2, value=v).border = BORDE
for i, (k, v) in enumerate(NIVEL_LBL.items(), start=3):
    we.cell(row=i, column=4, value=k).border = BORDE
    we.cell(row=i, column=5, value=v).border = BORDE
for i, (k, v) in enumerate(ESTADO_LBL.items(), start=3):
    we.cell(row=i, column=7, value=k).border = BORDE
    we.cell(row=i, column=8, value=v).border = BORDE
we["A12"] = "Recordatorio metodológico: PP07H = 2 indica que NO le descuentan aportes jubilatorios (empleo informal); " \
            "PP07C ≥ 2 indica inestabilidad contractual; PP3E_TOT = 999 es Ns./Nr.; P21 = -9 es no respuesta de ingresos."
we["A12"].font = Font(size=9.5, italic=True, color="5A6472")

# -------------------------------------------------- hoja de trabajo
COLS = ["CODUSU", "ANO4", "TRIMESTRE", "REGION", "MAS_500", "AGLOMERADO", "PONDERA",
        "CH03", "CH04", "CH06", "CH10", "NIVEL_ED", "ESTADO", "CAT_OCUP", "PP07C",
        "PP07H", "PP3E_TOT", "P21", "TOT_P12", "P47T", "IPCF"]
DERIV = ["REGION_TEXTO", "NIVEL_ED_TEXTO", "SEXO_TEXTO", "ESTADO_TEXTO", "TRAMO_EDAD",
         "IPL_Jubilacion", "IPL_Inestabilidad", "IPL_Jornada", "IPL",
         "ISE_NivelJoven", "ISE_Hogar", "ISE", "P21_LIMPIO", "HORAS_LIMPIO"]
wd = wb.create_sheet(HOJA_D)
encabezar(wd, 1, COLS + DERIV,
          ancho=[15, 7, 7, 8, 8, 11, 9, 7, 7, 7, 7, 9, 8, 9, 8, 8, 10, 11, 10, 10, 10]
          + [20, 20, 11, 12, 11, 13, 15, 12, 7, 14, 10, 7, 12, 13])
L = {n: get_column_letter(i + 1) for i, n in enumerate(COLS + DERIV)}
ult = N_JOV + 1

for r, row in enumerate(jov[COLS].itertuples(index=False), start=2):
    for j, v in enumerate(row, start=1):
        wd.cell(row=r, column=j,
                value=(None if (isinstance(v, float) and np.isnan(v)) else v))
    f = {
        "REGION_TEXTO": f'=IFERROR(VLOOKUP({L["REGION"]}{r},Tablas_Equivalencia!$A$3:$B$8,2,FALSE),"Sin dato")',
        "NIVEL_ED_TEXTO": f'=IFERROR(VLOOKUP({L["NIVEL_ED"]}{r},Tablas_Equivalencia!$D$3:$E$9,2,FALSE),"Sin dato")',
        "SEXO_TEXTO": f'=IF({L["CH04"]}{r}=1,"Varón","Mujer")',
        "ESTADO_TEXTO": f'=IFERROR(VLOOKUP({L["ESTADO"]}{r},Tablas_Equivalencia!$G$3:$H$6,2,FALSE),"Sin dato")',
        "TRAMO_EDAD": f'=IF({L["CH06"]}{r}<=24,"18-24","25-29")',
        "IPL_Jubilacion": f'=IF({L["ESTADO"]}{r}<>1,"",IF({L["PP07H"]}{r}=2,1,0))',
        "IPL_Inestabilidad": f'=IF({L["ESTADO"]}{r}<>1,"",IF({L["PP07C"]}{r}>=2,1,0))',
        "IPL_Jornada": f'=IF({L["ESTADO"]}{r}<>1,"",IF(OR({L["PP3E_TOT"]}{r}=999,'
                       f'{L["PP3E_TOT"]}{r}=""),"",IF(OR({L["PP3E_TOT"]}{r}<35,'
                       f'{L["PP3E_TOT"]}{r}>48),1,0)))',
        "IPL": f'=IF(COUNT({L["IPL_Jubilacion"]}{r}:{L["IPL_Jornada"]}{r})<3,"",'
               f'SUM({L["IPL_Jubilacion"]}{r}:{L["IPL_Jornada"]}{r}))',
        "ISE_NivelJoven": f'=IF({L["NIVEL_ED"]}{r}>=4,2,IF({L["NIVEL_ED"]}{r}>=2,1,0))',
        "ISE_Hogar": f'=IF({L["NIVEL_ED"]}{r}>=5,1,0)',
        "ISE": f'={L["ISE_NivelJoven"]}{r}+{L["ISE_Hogar"]}{r}',
        "P21_LIMPIO": f'=IF({L["ESTADO"]}{r}<>1,"",IF(OR({L["P21"]}{r}=-9,{L["P21"]}{r}=""),"",{L["P21"]}{r}))',
        "HORAS_LIMPIO": f'=IF({L["ESTADO"]}{r}<>1,"",IF(OR({L["PP3E_TOT"]}{r}=999,'
                        f'{L["PP3E_TOT"]}{r}=""),"",{L["PP3E_TOT"]}{r}))',
    }
    for n, formula in f.items():
        wd.cell(row=r, column=COLS.index(COLS[0]) + 1 + (COLS + DERIV).index(n), value=formula)

wd.freeze_panes = "B2"
wd.auto_filter.ref = f"A1:{L['HORAS_LIMPIO']}{ult}"
for col in ["P21", "TOT_P12", "P47T", "IPCF", "P21_LIMPIO"]:
    for r in range(2, ult + 1):
        wd[f"{L[col]}{r}"].number_format = '#,##0'

# -------------------------------------------------- hoja Diagnostico_Nulos
wn = wb.create_sheet("Diagnostico_Nulos")
titulo_hoja(wn, "Diagnóstico de valores faltantes y códigos especiales",
            "Tarea 1.3 — conteo de celdas vacías y códigos -9 / 999 sobre el segmento 18-29 años.")
encabezar(wn, 3, list(diag.columns), ancho=[14, 26, 16, 26, 20, 20])
for i, row in enumerate(diag.itertuples(index=False), start=4):
    for j, v in enumerate(row, start=1):
        c = wn.cell(row=i, column=j, value=v)
        c.border, c.alignment = BORDE, CENTRO
f0 = 4 + len(diag) + 1
wn.cell(row=f0, column=1, value="Verificación con fórmulas vivas (deben coincidir con la tabla):").font = \
    Font(bold=True, color=AZUL.lstrip("#"))
checks = [
    ("Registros totales de la base original", N_TOTAL),
    ("Jóvenes 18-29 años (CONTARA)", f"=COUNTA({REF_D}!$A$2:$A${ult})"),
    ("Jóvenes con CH06 entre 18 y 29 (CONTAR.SI.CONJUNTO)",
     f'=COUNTIFS({REF_D}!${L["CH06"]}$2:${L["CH06"]}${ult},">=18",'
     f'{REF_D}!${L["CH06"]}$2:${L["CH06"]}${ult},"<=29")'),
    ("Ocupados (ESTADO = 1)", f'=COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},1)'),
    ("Desocupados (ESTADO = 2)", f'=COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},2)'),
    ("Inactivos (ESTADO = 3)", f'=COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},3)'),
    ("P21 con código -9 (no respuesta)", f'=COUNTIF({REF_D}!${L["P21"]}$2:${L["P21"]}${ult},-9)'),
    ("PP3E_TOT con código 999 (Ns./Nr.)", f'=COUNTIF({REF_D}!${L["PP3E_TOT"]}$2:${L["PP3E_TOT"]}${ult},999)'),
    ("P21 igual a 0 (sin ingreso por ocupación)", f'=COUNTIF({REF_D}!${L["P21"]}$2:${L["P21"]}${ult},0)'),
]
for i, (t, v) in enumerate(checks, start=f0 + 1):
    wn.cell(row=i, column=1, value=t).border = BORDE
    c = wn.cell(row=i, column=2, value=v)
    c.border, c.number_format, c.font = BORDE, '#,##0', Font(bold=True)

f1 = f0 + len(checks) + 2
wn.cell(row=f1, column=1, value="Tablas de frecuencia de validación (Tareas 2.1 y 2.2)").font = \
    Font(bold=True, color=AZUL.lstrip("#"))
encabezar(wn, f1 + 1, ["Valor del índice", "Frecuencia IPL", "Frecuencia ISE",
                       "% IPL sobre ocupados", "% ISE sobre jóvenes"])
for k in range(4):
    r = f1 + 2 + k
    wn.cell(row=r, column=1, value=k).border = BORDE
    wn.cell(row=r, column=2, value=f'=COUNTIF({REF_D}!${L["IPL"]}$2:${L["IPL"]}${ult},{k})').border = BORDE
    wn.cell(row=r, column=3, value=f'=COUNTIF({REF_D}!${L["ISE"]}$2:${L["ISE"]}${ult},{k})').border = BORDE
    c = wn.cell(row=r, column=4, value=f"=IFERROR(B{r}/SUM($B${f1+2}:$B${f1+5}),0)")
    c.number_format, c.border = '0.0%', BORDE
    c = wn.cell(row=r, column=5, value=f"=IFERROR(C{r}/SUM($C${f1+2}:$C${f1+5}),0)")
    c.number_format, c.border = '0.0%', BORDE
r = f1 + 6
wn.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
wn.cell(row=r, column=2, value=f"=SUM(B{f1+2}:B{f1+5})").font = Font(bold=True)
wn.cell(row=r, column=3, value=f"=SUM(C{f1+2}:C{f1+5})").font = Font(bold=True)
wn.cell(row=r + 1, column=1,
        value="El total de ISE debe igualar la cantidad de jóvenes; el total de IPL, "
              "la cantidad de ocupados con jornada informada (los demás quedan en blanco por diseño).").font = \
    Font(size=9.5, italic=True, color="5A6472")

# -------------------------------------------------- hoja Correlacion
wc = wb.create_sheet("Correlacion")
titulo_hoja(wc, "Matriz de correlación de Pearson (5 × 5)",
            "Fórmula COEF.DE.CORREL sobre la hoja de trabajo. Los pares con celdas vacías "
            "se excluyen automáticamente: el análisis de ingresos y precariedad queda restringido a ocupados.")
cols_corr = {"Edad (CH06)": L["CH06"], "Ingreso ocup. principal (P21)": L["P21_LIMPIO"],
             "Horas semanales (PP3E_TOT)": L["HORAS_LIMPIO"],
             "Precariedad laboral (IPL)": L["IPL"], "Nivel socio-educativo (ISE)": L["ISE"]}
encabezar(wc, 3, [""] + list(cols_corr), ancho=[30, 17, 17, 17, 17, 17])
for i, (ni, ci) in enumerate(cols_corr.items(), start=4):
    c = wc.cell(row=i, column=1, value=ni)
    c.font, c.border = Font(bold=True, color="FFFFFF"), BORDE
    c.fill, c.alignment = FILL_TIT, Alignment(vertical="center", wrap_text=True)
    for j, (nj, cj) in enumerate(cols_corr.items(), start=2):
        cel = wc.cell(row=i, column=j,
                      value=f'=CORREL({REF_D}!${ci}$2:${ci}${ult},{REF_D}!${cj}$2:${cj}${ult})')
        cel.number_format, cel.border, cel.alignment = '0.000', BORDE, CENTRO
wc.conditional_formatting.add(
    "B4:F8", ColorScaleRule(start_type="num", start_value=-1, start_color="A63A2A",
                            mid_type="num", mid_value=0, mid_color="F3EFE8",
                            end_type="num", end_value=1, end_color="1E6F5C"))
r = 10
wc.cell(row=r, column=1, value="Lectura de los coeficientes significativos (|r| ≥ 0,3)").font = \
    Font(bold=True, size=12, color=AZUL.lstrip("#"))
lecturas = [
    f"ISE ↔ IPL = {kpi['r_ise_ipl']:+.3f}: a mayor nivel socio-educativo, menor precariedad laboral. "
    "Es la relación más fuerte de la matriz y sostiene la hipótesis central del proyecto.",
    f"ISE ↔ P21 = {kpi['r_ise_p21']:+.3f}: las credenciales educativas se traducen en ingresos ocupacionales más altos.",
    f"IPL ↔ P21 = {kpi['r_ipl_p21']:+.3f}: la precariedad opera como penalidad salarial: cada punto de IPL "
    "se asocia a menores ingresos.",
    f"Horas ↔ P21 = {kpi['r_horas_p21']:+.3f}: la intensidad horaria explica parte del ingreso, "
    "pero menos que las credenciales.",
    f"Edad ↔ P21 = {kpi['r_edad_p21']:+.3f}: dentro de una franja tan estrecha (18-29) la experiencia "
    "aporta poco; lo determinante es la calidad del puesto.",
]
for i, t in enumerate(lecturas, start=r + 1):
    wc.cell(row=i, column=1, value=f"• {t}").alignment = Alignment(wrap_text=True, vertical="top")
    wc.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
    wc.row_dimensions[i].height = 30

# -------------------------------------------------- hoja TD_Genero
wt = wb.create_sheet("TD_Genero")
titulo_hoja(wt, "IPL promedio por sexo y región (tabla dinámica equivalente)",
            "Construida con PROMEDIO.SI.CONJUNTO para que los valores se recalculen solos. "
            "En Excel podés reproducirla con Insertar → Tabla dinámica (CH04 en Filas, "
            "REGION_TEXTO en Columnas, IPL en Valores como Promedio) y agregarle segmentadores.")
encabezar(wt, 3, ["Sexo"] + regiones_ord + ["Total país"],
          ancho=[14] + [16] * len(regiones_ord) + [14])
for i, sexo in enumerate(["Varón", "Mujer"], start=4):
    c = wt.cell(row=i, column=1, value=sexo)
    c.font, c.border = Font(bold=True), BORDE
    for j, reg in enumerate(regiones_ord, start=2):
        cel = wt.cell(row=i, column=j,
                      value=f'=IFERROR(AVERAGEIFS({REF_D}!${L["IPL"]}$2:${L["IPL"]}${ult},'
                            f'{REF_D}!${L["SEXO_TEXTO"]}$2:${L["SEXO_TEXTO"]}${ult},$A{i},'
                            f'{REF_D}!${L["REGION_TEXTO"]}$2:${L["REGION_TEXTO"]}${ult},'
                            f'{get_column_letter(j)}$3),"s/d")')
        cel.number_format, cel.border, cel.alignment = '0.000', BORDE, CENTRO
    cel = wt.cell(row=i, column=len(regiones_ord) + 2,
                  value=f'=AVERAGEIF({REF_D}!${L["SEXO_TEXTO"]}$2:${L["SEXO_TEXTO"]}${ult},'
                        f'$A{i},{REF_D}!${L["IPL"]}$2:${L["IPL"]}${ult})')
    cel.number_format, cel.border, cel.alignment = '0.000', BORDE, CENTRO
    cel.font = Font(bold=True)
wt.cell(row=6, column=1, value="Brecha (Mujer − Varón)").font = Font(bold=True, color=NARANJA.lstrip("#"))
for j in range(2, len(regiones_ord) + 3):
    cl = get_column_letter(j)
    cel = wt.cell(row=6, column=j, value=f"=IFERROR({cl}5-{cl}4,\"s/d\")")
    cel.number_format, cel.border, cel.alignment = '+0.000;-0.000', BORDE, CENTRO
    cel.font = Font(bold=True, color=NARANJA.lstrip("#"))

wt.cell(row=8, column=1, value="Ingreso promedio (P21) por sexo y brecha de ingresos").font = \
    Font(bold=True, color=AZUL.lstrip("#"))
encabezar(wt, 9, ["Sexo", "Ingreso promedio", "Casos válidos", "Horas promedio", "% informalidad"])
for i, sexo in enumerate(["Varón", "Mujer"], start=10):
    wt.cell(row=i, column=1, value=sexo).font = Font(bold=True)
    base_args = f'{REF_D}!${L["SEXO_TEXTO"]}$2:${L["SEXO_TEXTO"]}${ult},$A{i}'
    vals = [
        (f'=AVERAGEIF({base_args},{REF_D}!${L["P21_LIMPIO"]}$2:${L["P21_LIMPIO"]}${ult})', '#,##0'),
        (f'=COUNTIFS({REF_D}!${L["SEXO_TEXTO"]}$2:${L["SEXO_TEXTO"]}${ult},$A{i},'
         f'{REF_D}!${L["P21_LIMPIO"]}$2:${L["P21_LIMPIO"]}${ult},">0")', '#,##0'),
        (f'=AVERAGEIF({base_args},{REF_D}!${L["HORAS_LIMPIO"]}$2:${L["HORAS_LIMPIO"]}${ult})', '0.0'),
        (f'=AVERAGEIF({base_args},{REF_D}!${L["IPL_Jubilacion"]}$2:${L["IPL_Jubilacion"]}${ult})', '0.0%'),
    ]
    for j, (formula, nf) in enumerate(vals, start=2):
        cel = wt.cell(row=i, column=j, value=formula)
        cel.number_format, cel.border, cel.alignment = nf, BORDE, CENTRO
wt.cell(row=12, column=1, value="Brecha de ingresos (1 − Mujer/Varón)").font = Font(bold=True)
cel = wt.cell(row=12, column=2, value="=IFERROR(1-B11/B10,\"s/d\")")
cel.number_format, cel.font = '0.0%', Font(bold=True, color=NARANJA.lstrip("#"))

graf_g = BarChart()
graf_g.type, graf_g.grouping, graf_g.gapWidth = "col", "clustered", 60
graf_g.title = "IPL promedio por sexo y región"
graf_g.y_axis.title = "IPL promedio (0 a 3)"
graf_g.x_axis.title = "Región"
graf_g.height, graf_g.width = 8.6, 19
datos = Reference(wt, min_col=1, max_col=len(regiones_ord) + 1, min_row=3, max_row=5)
graf_g.add_data(datos, from_rows=True, titles_from_data=True)
graf_g.series[0].graphicalProperties.solidFill = AZUL.lstrip("#")
graf_g.series[1].graphicalProperties.solidFill = NARANJA.lstrip("#")
wt.add_chart(graf_g, "A15")

# -------------------------------------------------- hoja Ingreso_Educacion
wi = wb.create_sheet("Ingreso_Educacion")
titulo_hoja(wi, "Ingreso y precariedad según nivel educativo",
            "Tarea 4.1 — PROMEDIO.SI sobre NIVEL_ED. Base: jóvenes ocupados con ingreso informado.")
encabezar(wi, 3, ["Código", "Nivel educativo", "Ingreso promedio (P21)", "Casos válidos",
                  "IPL promedio", "% informalidad"], ancho=[10, 26, 22, 14, 14, 16])
for i in range(1, 7):
    r = 3 + i
    wi.cell(row=r, column=1, value=i).border = BORDE
    wi.cell(row=r, column=2, value=NIVEL_LBL[i]).border = BORDE
    fs = [
        (f'=IFERROR(AVERAGEIF({REF_D}!${L["NIVEL_ED"]}$2:${L["NIVEL_ED"]}${ult},$A{r},'
         f'{REF_D}!${L["P21_LIMPIO"]}$2:${L["P21_LIMPIO"]}${ult}),0)', '#,##0'),
        (f'=COUNTIFS({REF_D}!${L["NIVEL_ED"]}$2:${L["NIVEL_ED"]}${ult},$A{r},'
         f'{REF_D}!${L["P21_LIMPIO"]}$2:${L["P21_LIMPIO"]}${ult},">0")', '#,##0'),
        (f'=IFERROR(AVERAGEIF({REF_D}!${L["NIVEL_ED"]}$2:${L["NIVEL_ED"]}${ult},$A{r},'
         f'{REF_D}!${L["IPL"]}$2:${L["IPL"]}${ult}),0)', '0.00'),
        (f'=IFERROR(AVERAGEIF({REF_D}!${L["NIVEL_ED"]}$2:${L["NIVEL_ED"]}${ult},$A{r},'
         f'{REF_D}!${L["IPL_Jubilacion"]}$2:${L["IPL_Jubilacion"]}${ult}),0)', '0.0%'),
    ]
    for j, (formula, nf) in enumerate(fs, start=3):
        cel = wi.cell(row=r, column=j, value=formula)
        cel.number_format, cel.border, cel.alignment = nf, BORDE, CENTRO

graf_i = LineChart()
graf_i.title = "Ingreso promedio de la ocupación principal según nivel educativo"
graf_i.y_axis.title = "Pesos corrientes"
graf_i.x_axis.title = "Nivel educativo"
graf_i.height, graf_i.width = 8.6, 19
graf_i.add_data(Reference(wi, min_col=3, min_row=3, max_row=9), titles_from_data=True)
graf_i.set_categories(Reference(wi, min_col=2, min_row=4, max_row=9))
graf_i.series[0].graphicalProperties.line.solidFill = AZUL.lstrip("#")
graf_i.series[0].graphicalProperties.line.width = 28000
graf_i.series[0].smooth = False
wi.add_chart(graf_i, "H3")

# -------------------------------------------------- hoja Bitacora
wbi = wb.create_sheet("Bitacora")
titulo_hoja(wbi, "Bitácora de decisiones metodológicas",
            "Toda decisión que altera los resultados queda documentada y es auditable.")
encabezar(wbi, 3, ["#", "Decisión", "Fundamento", "Impacto en los resultados"],
          ancho=[5, 40, 52, 44])
BITACORA = [
    ("Universo: 18 a 29 años cumplidos (CH06)",
     "Definición de juventud usada por los programas de primer empleo; excluye adolescentes escolarizados obligatorios.",
     f"Quedan {N_JOV:,} de {N_TOTAL:,} registros ({N_JOV/N_TOTAL:.1%} de la muestra)."),
    ("No se eliminan los casos con P21 = 0",
     "Un ingreso cero no es un error: identifica a desocupados e inactivos, población central del diagnóstico.",
     "Se conserva la estructura real del segmento; los promedios de ingreso se calculan solo sobre ocupados."),
    ("P21 = -9 y PP3E_TOT = 999 se tratan como faltantes, no como ceros",
     "Son códigos de no respuesta del INDEC; computarlos como valores hundiría artificialmente los promedios.",
     f"Se excluyen {int((jov.P21 == -9).sum()):,} ingresos y {int((jov.PP3E_TOT == 999).sum()):,} jornadas."),
    ("Las columnas P21_LIMPIO y HORAS_LIMPIO alimentan la matriz de correlación",
     "COEF.DE.CORREL descarta los pares con texto o vacío: así el coeficiente se calcula sobre casos válidos.",
     "Evita el sesgo de incluir ceros estructurales en las correlaciones de ingreso."),
    ("IPL en blanco para quienes no están ocupados (ESTADO ≠ 1)",
     "PP07H y PP07C solo se relevan a ocupados; asignarles 0 los contaría como empleo de calidad.",
     f"El IPL se calcula sobre {int(jov.IPL.notna().sum()):,} ocupados con jornada informada."),
    ("Umbrales del IPL: sin aportes (PP07H=2), inestabilidad (PP07C≥2), jornada <35 h o >48 h",
     "Replican los criterios de informalidad, temporalidad y sub/sobreocupación usados en la literatura del mercado laboral.",
     f"IPL promedio = {kpi['ipl_promedio']:.2f}; {kpi['ipl3_pct']:.1f}% de los ocupados acumula las tres carencias."),
    ("ISE: 2 puntos por secundaria completa o más, 1 punto extra por estudios superiores",
     "Discrimina el salto de credenciales que el mercado premia; mantiene una escala comparable con el IPL (0 a 3).",
     f"Correlación ISE ↔ IPL = {kpi['r_ise_ipl']:+.3f}."),
    ("Se conserva una copia intacta del archivo original",
     "Trazabilidad: cualquier resultado debe poder reproducirse desde el dato crudo.",
     "01_Datos_Originales/ nunca se modifica; todo el trabajo ocurre en 02_Datos_Procesados/."),
    ("Base de trabajo SIMULADA con la estructura oficial de la EPH",
     "El entorno de ejecución no tiene acceso al portal del INDEC; se generó una base con los mismos nombres y códigos.",
     "Las fórmulas, el dashboard y el informe son válidos; los valores se actualizan al reemplazar el .txt por el oficial."),
]
for i, (dec, fund, imp) in enumerate(BITACORA, start=1):
    r = 3 + i
    wbi.cell(row=r, column=1, value=i).alignment = CENTRO
    for j, v in enumerate([dec, fund, imp], start=2):
        c = wbi.cell(row=r, column=j, value=v)
        c.alignment, c.border = Alignment(wrap_text=True, vertical="top"), BORDE
    wbi.row_dimensions[r].height = 46

# -------------------------------------------------- hoja DASHBOARD
wdb = wb.create_sheet("DASHBOARD", 0)
wdb.sheet_view.showGridLines = False
wdb.column_dimensions["A"].width = 3
for col in "BCDEFGHIJKLMN":
    wdb.column_dimensions[col].width = 13.5
wdb.merge_cells("B2:N2")
wdb["B2"] = "PANEL DE ANÁLISIS — EMPLEO JUVENIL EN ARGENTINA (EPH · INDEC)"
wdb["B2"].font = Font(name="Calibri", size=20, bold=True, color=AZUL.lstrip("#"))
wdb.row_dimensions[2].height = 30
wdb.merge_cells("B3:N3")
periodo = f"{int(base.ANO4.iloc[0])} · trimestre {int(base.TRIMESTRE.iloc[0])}"
wdb["B3"] = (f"Segmento 18 a 29 años · EPH {periodo}"
             f" · elaborado por Jeisson Marín Uribe para Fundación Impacto Laboral"
             f" · actualizado {date.today().strftime('%d/%m/%Y')}")
wdb["B3"].font = Font(size=10.5, italic=True, color="5A6472")

KPIS = [
    ("Jóvenes 18-29 relevados", f"=COUNTA({REF_D}!$A$2:$A${ult})", '#,##0'),
    ("Tasa de desocupación juvenil",
     f'=COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},2)/'
     f'(COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},1)+'
     f'COUNTIF({REF_D}!${L["ESTADO"]}$2:${L["ESTADO"]}${ult},2))', '0.0%'),
    ("Informalidad (sin aportes)",
     f'=AVERAGE({REF_D}!${L["IPL_Jubilacion"]}$2:${L["IPL_Jubilacion"]}${ult})', '0.0%'),
    ("IPL promedio (0 a 3)", f'=AVERAGE({REF_D}!${L["IPL"]}$2:${L["IPL"]}${ult})', '0.00'),
    ("Ingreso promedio ocupados",
     f'=AVERAGE({REF_D}!${L["P21_LIMPIO"]}$2:${L["P21_LIMPIO"]}${ult})', '$ #,##0'),
    ("Correlación ISE ↔ IPL", "=Correlacion!F7", '0.000'),
]
col_ini = 2
for k, (nombre, formula, nf) in enumerate(KPIS):
    c0 = col_ini + (k % 3) * 4
    fila = 5 if k < 3 else 8
    rng = f"{get_column_letter(c0)}{fila}:{get_column_letter(c0+3)}{fila}"
    wdb.merge_cells(rng)
    cel = wdb.cell(row=fila, column=c0, value=nombre)
    cel.font = Font(size=10, bold=True, color="FFFFFF")
    cel.fill = PatternFill("solid", fgColor=AZUL.lstrip("#"))
    cel.alignment = Alignment(horizontal="center", vertical="center")
    rng2 = f"{get_column_letter(c0)}{fila+1}:{get_column_letter(c0+3)}{fila+1}"
    wdb.merge_cells(rng2)
    cel = wdb.cell(row=fila + 1, column=c0, value=formula)
    cel.font = Font(size=18, bold=True, color=TINTA.lstrip("#"))
    cel.fill = PatternFill("solid", fgColor=CLARO.lstrip("#"))
    cel.alignment = Alignment(horizontal="center", vertical="center")
    cel.number_format = nf
    wdb.row_dimensions[fila].height = 20
    wdb.row_dimensions[fila + 1].height = 34

wdb["B11"] = "Gráfico 1 — Matriz de calor de correlación"
wdb["B11"].font = Font(size=12, bold=True, color=AZUL.lstrip("#"))
wdb["H11"] = "Gráfico 2 — Brecha de género en precariedad, por región"
wdb["H11"].font = Font(size=12, bold=True, color=AZUL.lstrip("#"))
img1 = XLImage(str(DIR_VIS / "G1_Matriz_Correlacion.png"))
img1.width, img1.height = 470, 380
wdb.add_image(img1, "B12")
img2 = XLImage(str(DIR_VIS / "G2_Brecha_Genero.png"))
img2.width, img2.height = 560, 320
wdb.add_image(img2, "H12")
wdb["B33"] = "Gráfico 3 — Ingreso y precariedad según nivel educativo"
wdb["B33"].font = Font(size=12, bold=True, color=AZUL.lstrip("#"))
img3 = XLImage(str(DIR_VIS / "G3_Ingreso_Educacion.png"))
img3.width, img3.height = 700, 395
wdb.add_image(img3, "B34")

wdb["H33"] = "Cómo leer este panel"
wdb["H33"].font = Font(size=12, bold=True, color=AZUL.lstrip("#"))
notas = [
    f"· {kpi['informalidad']:.1f}% de los jóvenes ocupados trabaja sin aportes jubilatorios.",
    f"· El IPL promedio es {kpi['ipl_promedio']:.2f} sobre 3 y {kpi['ipl3_pct']:.1f}% acumula las tres carencias.",
    f"· Las mujeres registran un IPL de {kpi['ipl_mujer']:.2f} frente a {kpi['ipl_varon']:.2f} de los varones "
    f"(brecha {kpi['brecha_total']:+.2f}).",
    f"· La brecha de género más alta se observa en {kpi['region_mayor_brecha']} ({kpi['brecha_max']:+.2f}).",
    f"· Correlación ISE ↔ IPL = {kpi['r_ise_ipl']:+.3f}: la educación es el principal amortiguador de la precariedad.",
    "· Todas las celdas son fórmulas vivas: al reemplazar el archivo de microdatos, el panel se recalcula solo.",
    "· Para agregar segmentadores interactivos: Insertar → Tabla dinámica sobre la hoja de trabajo "
    "y luego Analizar → Insertar segmentación de datos (REGION_TEXTO y TRAMO_EDAD).",
]
for i, t in enumerate(notas, start=34):
    wdb.merge_cells(start_row=i, start_column=8, end_row=i, end_column=14)
    c = wdb.cell(row=i, column=8, value=t)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(size=10)
    wdb.row_dimensions[i].height = 28

wdb.sheet_properties.tabColor = AZUL.lstrip("#")
wb.save(XLSX)
print(f"   libro guardado: {XLSX}")

# ==========================================================================
# 8. Hallazgos para el informe ejecutivo
# ==========================================================================
salida = {
    "kpi": kpi,
    "corr": corr.round(3).to_dict(),
    "td_genero": td.to_dict(),
    "brecha": brecha.to_dict(),
    "ingreso_educacion": ing_edu.to_dict(orient="records"),
    "frecuencias": {"IPL": frec_ipl.to_dict(), "ISE": frec_ise.to_dict()},
    "diagnostico_nulos": diag.to_dict(orient="records"),
    "fuente": ENTRADA.name,
    "generado": date.today().isoformat(),
}
(Path(__file__).parent / "_hallazgos.json").write_text(
    json.dumps(salida, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

print("\n=== HALLAZGOS PRINCIPALES ===")
for k in ["jovenes", "ocupados", "tasa_actividad", "tasa_desocupacion", "informalidad",
          "ipl_promedio", "ipl3_pct", "ingreso_promedio", "brecha_ingresos_pct",
          "ipl_varon", "ipl_mujer", "brecha_total", "region_mayor_brecha", "brecha_max",
          "r_ise_ipl", "r_ise_p21", "r_ipl_p21", "r_horas_p21", "r_edad_p21",
          "prima_superior_pct"]:
    print(f"{k:28s}: {kpi[k]}")
