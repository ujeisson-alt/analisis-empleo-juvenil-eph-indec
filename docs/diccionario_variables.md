# Diccionario de variables utilizadas (Diseño de Registro EPH — INDEC)

Variables de la base de **personas** (`usu_individual`) empleadas en este proyecto.

| Variable | Descripción | Códigos / unidad |
|---|---|---|
| `CODUSU` | Identificador del hogar | alfanumérico |
| `ANO4` / `TRIMESTRE` | Año y trimestre del relevamiento | 2025 / 1 a 4 |
| `REGION` | Región geográfica | 1 Gran Buenos Aires · 40 Noroeste (NOA) · 41 Noreste (NEA) · 42 Cuyo · 43 Pampeana · 44 Patagonia |
| `AGLOMERADO` | Aglomerado urbano relevado | códigos INDEC |
| `MAS_500` | Aglomerado de 500.000 habitantes o más | S / N |
| `PONDERA` | Ponderador muestral | entero (no aplicado en este MVP) |
| `CH03` | Relación de parentesco con el jefe de hogar | 1 jefe/a · 2 cónyuge · 3 hijo/a · 5 nieto/a · 6 madre/padre · 9 otros · 10 no familiares |
| `CH04` | Sexo | **1 varón · 2 mujer** |
| `CH06` | Edad | años cumplidos |
| `CH10` | Asiste o asistió a un establecimiento educativo | 1 asiste · 2 no asiste, asistió · 3 nunca asistió |
| `NIVEL_ED` | Nivel educativo alcanzado | 1 primaria incompleta · 2 primaria completa · 3 secundaria incompleta · 4 secundaria completa · 5 superior incompleto · 6 superior completo · 7 sin instrucción |
| `ESTADO` | Condición de actividad | **1 ocupado · 2 desocupado · 3 inactivo · 4 menor de 10 años** |
| `CAT_OCUP` | Categoría ocupacional | 1 patrón · 2 cuenta propia · 3 obrero/empleado · 4 trabajador familiar |
| `PP07C` | Estabilidad del puesto | 1 estable · ≥ 2 temporario o sin relación estable |
| `PP07H` | ¿Le descuentan aportes jubilatorios? | **1 sí (formal) · 2 no (informal)** |
| `PP3E_TOT` | Horas semanales trabajadas en la ocupación principal | horas · **999 = Ns./Nr.** |
| `P21` | Ingreso de la ocupación principal | pesos · **−9 = no respuesta** · 0 = sin ingreso por ocupación |
| `TOT_P12` | Ingreso por otras ocupaciones | pesos |
| `P47T` | Ingreso total individual | pesos |
| `IPCF` | Ingreso per cápita familiar | pesos |

## Variables derivadas creadas en este proyecto

| Variable | Fórmula (notación Excel en español) | Rango |
|---|---|---|
| `REGION_TEXTO` | `=BUSCARV(REGION;Tablas_Equivalencia!$A$3:$B$8;2;FALSO)` | texto |
| `NIVEL_ED_TEXTO` | `=BUSCARV(NIVEL_ED;Tablas_Equivalencia!$D$3:$E$9;2;FALSO)` | texto |
| `SEXO_TEXTO` | `=SI(CH04=1;"Varón";"Mujer")` | texto |
| `TRAMO_EDAD` | `=SI(CH06<=24;"18-24";"25-29")` | texto |
| `IPL_Jubilacion` | `=SI(ESTADO<>1;"";SI(PP07H=2;1;0))` | 0-1 |
| `IPL_Inestabilidad` | `=SI(ESTADO<>1;"";SI(PP07C>=2;1;0))` | 0-1 |
| `IPL_Jornada` | `=SI(ESTADO<>1;"";SI(O(PP3E_TOT=999;PP3E_TOT="");"";SI(O(PP3E_TOT<35;PP3E_TOT>48);1;0)))` | 0-1 |
| `IPL` | `=SI(CONTAR(IPL_Jubilacion:IPL_Jornada)<3;"";SUMA(IPL_Jubilacion:IPL_Jornada))` | **0-3** |
| `ISE_NivelJoven` | `=SI(NIVEL_ED>=4;2;SI(NIVEL_ED>=2;1;0))` | 0-2 |
| `ISE_Hogar` | `=SI(NIVEL_ED>=5;1;0)` | 0-1 |
| `ISE` | `=ISE_NivelJoven+ISE_Hogar` | **0-3** |
| `P21_LIMPIO` | `=SI(ESTADO<>1;"";SI(O(P21=-9;P21="");"";P21))` | pesos |
| `HORAS_LIMPIO` | `=SI(ESTADO<>1;"";SI(O(PP3E_TOT=999;PP3E_TOT="");"";PP3E_TOT))` | horas |

> En el archivo `.xlsx` las fórmulas están guardadas con los nombres internos en inglés (`IF`, `OR`, `VLOOKUP`,
> `CORREL`…). Excel en español las muestra automáticamente como `SI`, `O`, `BUSCARV` y `COEF.DE.CORREL`.
