#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera el logo de la fundacion FICTICIA usada como cliente del proyecto."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from pathlib import Path

AZUL, NARANJA = "#14375E", "#C2603C"
SAL = Path(__file__).resolve().parents[1] / "03_Visualizaciones" / "logo_fundacion.png"

fig, ax = plt.subplots(figsize=(4.6, 1.5))
ax.set_xlim(0, 4.6); ax.set_ylim(0, 1.5); ax.axis("off")

ax.add_patch(FancyBboxPatch((0.08, 0.28), 0.95, 0.95, boxstyle="round,pad=0.02,rounding_size=0.14",
                            linewidth=0, facecolor=AZUL))
for i, (h, col) in enumerate([(0.30, "#FFFFFF"), (0.48, "#FFFFFF"), (0.66, NARANJA)]):
    ax.add_patch(Rectangle((0.26 + i * 0.23, 0.42), 0.14, h, facecolor=col, linewidth=0))

ax.text(1.22, 0.86, "FUNDACIÓN", fontsize=15, color=AZUL, fontweight="bold",
        family="DejaVu Sans", va="center")
ax.text(1.22, 0.56, "IMPACTO LABORAL", fontsize=15, color=NARANJA, fontweight="bold",
        family="DejaVu Sans", va="center")
ax.text(1.24, 0.30, "entidad ficticia · caso de estudio", fontsize=7.5, color="#6B7280",
        family="DejaVu Sans", va="center")

fig.savefig(SAL, dpi=300, transparent=True, bbox_inches="tight")
print("logo:", SAL)
