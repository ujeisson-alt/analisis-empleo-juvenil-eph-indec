#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
00_generar_base_simulada.py
===========================
Genera una base de microdatos SIMULADA con la estructura, los nombres de
columna y los codigos oficiales del Diseno de Registro de la EPH (INDEC),
para poder ejecutar y auditar el pipeline completo del proyecto sin
depender de la descarga del .zip oficial.

IMPORTANTE: los datos NO son oficiales. Las distribuciones estan calibradas
con valores publicos de referencia del mercado laboral argentino (tasas de
actividad, desocupacion e informalidad juvenil, brecha de genero, estructura
educativa y regional), de modo que los ordenes de magnitud y los signos de
las correlaciones sean plausibles, pero ningun valor puntual debe citarse
como dato del INDEC.

Para trabajar con datos reales: ver 01_Datos_Originales/LEEME_datos.md

Salida: 01_Datos_Originales/usu_individual_T125_SIMULADO.txt  (delimitado por ';')
"""

import numpy as np
import pandas as pd
from pathlib import Path

SEMILLA = 2026
N_TOTAL = 32000          # personas encuestadas (todas las edades)
ANO4, TRIMESTRE = 2025, 1

rng = np.random.default_rng(SEMILLA)
RAIZ = Path(__file__).resolve().parents[1]
SALIDA = RAIZ / "01_Datos_Originales" / "usu_individual_T125_SIMULADO.txt"

# --------------------------------------------------------------------------
# 1. Identificacion, region y aglomerado
# --------------------------------------------------------------------------
REGIONES = [1, 40, 41, 42, 43, 44]              # GBA, NOA, NEA, Cuyo, Pampeana, Patagonia
PESO_REG = [0.31, 0.15, 0.11, 0.12, 0.24, 0.07]
AGLO_POR_REG = {1: [32, 33], 40: [7, 9, 22, 23], 41: [8, 12, 15],
                42: [10, 27, 25], 43: [17, 19, 20, 26, 29, 30, 34, 36, 38],
                44: [2, 4, 5, 6, 13, 14, 93]}

region = rng.choice(REGIONES, size=N_TOTAL, p=PESO_REG)
aglomerado = np.array([rng.choice(AGLO_POR_REG[r]) for r in region])

codusu = np.array([f"TQRMNOP{str(100000 + i)}" for i in range(N_TOTAL)])

# --------------------------------------------------------------------------
# 2. Demografia: edad, sexo, parentesco
# --------------------------------------------------------------------------
# Piramide poblacional argentina simplificada por tramos
tramos = [(0, 9), (10, 17), (18, 24), (25, 29), (30, 44), (45, 59), (60, 74), (75, 95)]
p_tramos = [0.145, 0.125, 0.105, 0.080, 0.215, 0.180, 0.110, 0.040]
idx = rng.choice(len(tramos), size=N_TOTAL, p=p_tramos)
ch06 = np.array([rng.integers(tramos[i][0], tramos[i][1] + 1) for i in idx])

ch04 = rng.choice([1, 2], size=N_TOTAL, p=[0.485, 0.515])      # 1 varon / 2 mujer

# CH03 relacion de parentesco: 1 jefe, 2 conyuge, 3 hijo, 4 yerno/nuera,
# 5 nieto, 6 madre/padre, 7 suegro, 8 hermano, 9 otros, 10 no familiares
ch03 = np.where(
    ch06 < 18, rng.choice([3, 5, 9], size=N_TOTAL, p=[0.86, 0.10, 0.04]),
    np.where(ch06 < 30,
             rng.choice([1, 2, 3, 10], size=N_TOTAL, p=[0.30, 0.18, 0.47, 0.05]),
             rng.choice([1, 2, 3, 6, 9], size=N_TOTAL, p=[0.55, 0.30, 0.07, 0.04, 0.04])))

# --------------------------------------------------------------------------
# 3. Nivel educativo (NIVEL_ED) y asistencia (CH10)
# --------------------------------------------------------------------------
# 1 primaria incompleta / 2 primaria completa / 3 secundaria incompleta
# 4 secundaria completa / 5 superior incompleto / 6 superior completo / 7 sin instruccion
# La estructura educativa depende de la edad y (levemente) de la region.
VENTAJA_REG = {1: 0.06, 43: 0.04, 44: 0.03, 42: 0.00, 40: -0.05, 41: -0.07}

def perfil_educativo(edad, reg, sexo):
    v = VENTAJA_REG[reg] + (0.03 if sexo == 2 else 0.0)   # mujeres con mas credenciales
    if edad < 15:
        base = [0.55, 0.20, 0.20, 0.00, 0.00, 0.00, 0.05]
    elif edad < 18:
        base = [0.05, 0.22, 0.65, 0.06, 0.00, 0.00, 0.02]
    elif edad < 22:
        base = [0.03, 0.06, 0.28, 0.35, 0.26, 0.01, 0.01]
    elif edad < 30:
        base = [0.02, 0.05, 0.19, 0.31, 0.24, 0.18, 0.01]
    elif edad < 60:
        base = [0.04, 0.11, 0.20, 0.27, 0.15, 0.22, 0.01]
    else:
        base = [0.13, 0.22, 0.17, 0.20, 0.09, 0.17, 0.02]
    p = np.array(base, dtype=float)
    # el "empuje" hacia credenciales altas segun ventaja regional / de genero
    p[3:6] *= (1 + 2.5 * v)
    p[0:3] *= (1 - 2.0 * v)
    p = np.clip(p, 1e-6, None)
    return p / p.sum()

nivel_ed = np.empty(N_TOTAL, dtype=int)
for i in range(N_TOTAL):
    nivel_ed[i] = rng.choice([1, 2, 3, 4, 5, 6, 7], p=perfil_educativo(ch06[i], region[i], ch04[i]))

# CH10 asiste a establecimiento educativo: 1 si, 2 no asiste pero asistio, 3 nunca
ch10 = np.where(ch06 < 18, 1,
       np.where((ch06 < 30) & np.isin(nivel_ed, [5]), rng.choice([1, 2], N_TOTAL, p=[0.72, 0.28]),
       np.where(ch06 < 30, rng.choice([1, 2], N_TOTAL, p=[0.22, 0.78]),
                rng.choice([1, 2, 3], N_TOTAL, p=[0.06, 0.92, 0.02]))))

# --------------------------------------------------------------------------
# 4. Indice socio-educativo latente (no se publica: solo ordena el modelo)
# --------------------------------------------------------------------------
cred = {7: 0.0, 1: 0.1, 2: 0.25, 3: 0.40, 4: 0.60, 5: 0.72, 6: 0.90}
capital = np.array([cred[n] for n in nivel_ed]) + np.array([VENTAJA_REG[r] for r in region])
capital = np.clip(capital + rng.normal(0, 0.05, N_TOTAL), 0, 1.1)

# --------------------------------------------------------------------------
# 5. Condicion de actividad (ESTADO)
# --------------------------------------------------------------------------
# 0 entrevista no realizada / 1 ocupado / 2 desocupado / 3 inactivo / 4 menor de 10
estado = np.full(N_TOTAL, 3)
estado[ch06 < 10] = 4

for i in range(N_TOTAL):
    e, edad = estado[i], ch06[i]
    if e == 4:
        continue
    if edad < 18:
        p_act = 0.10
    elif edad < 25:
        p_act = 0.56 + 0.22 * capital[i] - (0.10 if ch04[i] == 2 else 0.0)
    elif edad < 30:
        p_act = 0.78 + 0.14 * capital[i] - (0.13 if ch04[i] == 2 else 0.0)
    elif edad < 60:
        p_act = 0.80 + 0.10 * capital[i] - (0.16 if ch04[i] == 2 else 0.0)
    elif edad < 70:
        p_act = 0.33
    else:
        p_act = 0.07
    if rng.random() > np.clip(p_act, 0.02, 0.97):
        estado[i] = 3
        continue
    # dentro de los activos: probabilidad de desocupacion
    if edad < 25:
        p_des = 0.24 - 0.09 * capital[i] + (0.05 if ch04[i] == 2 else 0.0)
    elif edad < 30:
        p_des = 0.15 - 0.07 * capital[i] + (0.04 if ch04[i] == 2 else 0.0)
    else:
        p_des = 0.07 - 0.03 * capital[i] + (0.02 if ch04[i] == 2 else 0.0)
    estado[i] = 2 if rng.random() < np.clip(p_des, 0.01, 0.6) else 1

ocupado = estado == 1

# --------------------------------------------------------------------------
# 6. Caracteristicas del puesto (solo ocupados)
# --------------------------------------------------------------------------
NA = ""   # celda vacia: la EPH deja en blanco lo que "no corresponde"

cat_ocup = np.full(N_TOTAL, np.nan)          # 1 patron 2 cuenta propia 3 obrero/empleado 4 trab. familiar
pp07h = np.full(N_TOTAL, np.nan)             # 1 si le descuentan jubilacion / 2 no
pp07c = np.full(N_TOTAL, np.nan)             # 1 estable / 2 temporario / 3 sin relacion estable
pp3e_tot = np.full(N_TOTAL, np.nan)          # horas semanales trabajadas (999 = Ns/Nr)
p21 = np.full(N_TOTAL, np.nan)               # ingreso de la ocupacion principal

for i in np.flatnonzero(ocupado):
    joven = ch06[i] < 30
    k = capital[i]
    # categoria ocupacional
    p_cp = np.clip(0.26 - 0.10 * k + (0.03 if joven else 0.0), 0.05, 0.45)   # cuenta propia
    p_pat = np.clip(0.05 + 0.05 * k - (0.03 if joven else 0.0), 0.005, 0.12)
    p_fam = 0.03 if joven else 0.012
    cat_ocup[i] = rng.choice([1, 2, 3, 4],
                             p=[p_pat, p_cp, 1 - p_pat - p_cp - p_fam, p_fam])
    # informalidad: sin descuento jubilatorio (PP07H = 2)
    base_inf = 1.26 if joven else 0.92
    p_inf = base_inf - 1.06 * k + (0.05 if ch04[i] == 2 else 0.0) + (0.10 if cat_ocup[i] == 2 else 0.0) \
            + (0.22 if cat_ocup[i] == 4 else 0.0) - (0.06 if region[i] in (1, 43, 44) else 0.04)
    informal = rng.random() < np.clip(p_inf, 0.05, 0.95)
    pp07h[i] = 2 if informal else 1
    # estabilidad contractual: fuertemente ligada a la formalidad
    p_inest = float(np.clip((0.92 if informal else 0.26) - 0.26 * k, 0.03, 0.95))
    pp07c[i] = rng.choice([1, 2, 3], p=[1 - p_inest, p_inest * 0.72, p_inest * 0.28])
    # horas trabajadas
    mu = 45 - 7 * informal - 9 * (1 - k) - (4.5 if ch04[i] == 2 else 0) - (1.5 if joven else 0)
    h = rng.normal(mu, 11)
    if rng.random() < 0.04 + 0.07 * (1 - k):      # jornadas extremas (mas frecuentes con baja calificacion)
        h = rng.choice([rng.uniform(6, 20), rng.uniform(55, 80)])
    pp3e_tot[i] = int(np.clip(round(h), 1, 90))
    if rng.random() < 0.015:                      # Ns./Nr. codificado 999
        pp3e_tot[i] = 999

# Ingreso de la ocupacion principal (modelo log-normal, pesos corrientes 2025)
mask = np.flatnonzero(ocupado)
horas_ef = np.where(pp3e_tot[mask] == 999, 40, pp3e_tot[mask])
log_ing = (np.log(255000)
           + 0.95 * capital[mask]                                   # credenciales
           + 0.30 * np.log(np.clip(horas_ef, 6, 80) / 40)           # intensidad horaria
           + 0.34 * (pp07h[mask] == 1)                              # prima de formalidad
           + 0.012 * np.clip(ch06[mask] - 18, 0, 30)                # experiencia
           - 0.19 * (ch04[mask] == 2)                               # brecha de genero
           + np.array([{1: 0.10, 43: 0.05, 44: 0.16, 42: -0.02, 40: -0.10, 41: -0.13}[r]
                       for r in region[mask]])
           + rng.normal(0, 0.42, mask.size))
ing = np.round(np.exp(log_ing) / 1000) * 1000
p21[mask] = ing
# no respuesta de ingresos: el INDEC codifica -9
no_resp = rng.random(mask.size) < 0.075
p21[mask[no_resp]] = -9

# Desocupados e inactivos: P21 = 0 (sin ingreso por ocupacion)
p21[(estado == 2) | (estado == 3)] = 0

# Otros ingresos y agregados del hogar
tot_p12 = np.where(rng.random(N_TOTAL) < 0.17,
                   np.round(rng.lognormal(11.2, 0.6, N_TOTAL) / 1000) * 1000, 0)
tot_p12 = np.where(estado == 4, np.nan, tot_p12)
p47t = np.where(np.isnan(p21), np.nan, np.where(p21 < 0, -9, p21)) \
       + np.nan_to_num(tot_p12, nan=0.0)
p47t = np.where(np.isnan(p21), np.nan, np.maximum(p47t, 0))
ipcf = np.round(np.clip(rng.lognormal(11.9, 0.55, N_TOTAL), 20000, 4e6) / 100) * 100

pondera = np.round(np.clip(rng.normal(1150, 420, N_TOTAL), 90, 4200)).astype(int)

# --------------------------------------------------------------------------
# 7. Ensamblado y exportacion con el formato del INDEC
# --------------------------------------------------------------------------
def fmt(v, entero=True):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return NA
    return str(int(v)) if entero else str(v)

df = pd.DataFrame({
    "CODUSU": codusu,
    "ANO4": ANO4,
    "TRIMESTRE": TRIMESTRE,
    "REGION": region,
    "MAS_500": np.where(np.isin(aglomerado, [32, 33, 17, 19, 20, 7, 8, 10, 13, 30, 36]), "S", "N"),
    "AGLOMERADO": aglomerado,
    "PONDERA": pondera,
    "CH03": ch03,
    "CH04": ch04,
    "CH06": ch06,
    "CH10": ch10,
    "NIVEL_ED": nivel_ed,
    "ESTADO": estado,
    "CAT_OCUP": [fmt(v) for v in cat_ocup],
    "PP07C": [fmt(v) for v in pp07c],
    "PP07H": [fmt(v) for v in pp07h],
    "PP3E_TOT": [fmt(v) for v in pp3e_tot],
    "P21": [fmt(v) for v in p21],
    "TOT_P12": [fmt(v) for v in tot_p12],
    "P47T": [fmt(v) for v in p47t],
    "IPCF": [fmt(v) for v in ipcf],
})

SALIDA.parent.mkdir(parents=True, exist_ok=True)
df.to_csv(SALIDA, sep=";", index=False, encoding="utf-8")

# --------------------------------------------------------------------------
# 8. Diagnostico en consola (queda registrado en la bitacora)
# --------------------------------------------------------------------------
jov = df[(df.CH06 >= 18) & (df.CH06 <= 29)]
ocu = jov[jov.ESTADO == 1]
print(f"Archivo generado : {SALIDA}")
print(f"Registros totales: {len(df):,}")
print(f"Jovenes 18-29    : {len(jov):,}  ({len(jov)/len(df):.1%})")
print(f"  ocupados       : {(jov.ESTADO == 1).sum():,}")
print(f"  desocupados    : {(jov.ESTADO == 2).sum():,}")
print(f"  inactivos      : {(jov.ESTADO == 3).sum():,}")
print(f"Informalidad ocupados jovenes (PP07H=2): "
      f"{(ocu.PP07H.astype(str) == '2').mean():.1%}")
print(f"P21 = -9 entre jovenes ocupados        : {(ocu.P21.astype(str) == '-9').sum():,}")
print(f"PP3E_TOT = 999 entre jovenes ocupados  : {(ocu.PP3E_TOT.astype(str) == '999').sum():,}")
