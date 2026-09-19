# Bitácora de decisiones metodológicas

Proyecto: **Análisis del Empleo Juvenil en Argentina — microdatos EPH (INDEC)**
Analista: Jeisson Marín Uribe · Última actualización: septiembre de 2026

Toda decisión que altera un resultado queda documentada acá, con su fundamento y su impacto. El criterio es
simple: cualquier persona debe poder reconstruir el número exacto del informe siguiendo esta bitácora.

---

## Semana 1 — Extracción, filtrado y normalización

**Conteos de control**

| Control | Valor |
|---|---|
| Registros de la base original | 32.000 |
| Jóvenes de 18 a 29 años (`CH06`) | **5.933** (18,5 %) |
| Ocupados (`ESTADO = 1`) | 3.521 |
| Desocupados (`ESTADO = 2`) | 704 |
| Inactivos (`ESTADO = 3`) | 1.708 |

**Diagnóstico de faltantes en el segmento juvenil**

| Variable | Celdas vacías | Código especial | Casos con el código | Válidos |
|---|---|---|---|---|
| `P21` (ingreso de la ocupación principal) | 0 | −9 | 304 | 5.629 |
| `PP3E_TOT` (horas semanales) | 2.412 | 999 | 55 | 3.466 |
| `PP07H` (aportes jubilatorios) | 2.412 | — | — | 3.521 |
| `PP07C` (estabilidad contractual) | 2.412 | — | — | 3.521 |

> Las 2.412 celdas vacías no son un error de datos: son las personas **no ocupadas**, a quienes la encuesta no
> les hace las preguntas del puesto de trabajo.

**Decisiones**

1. **Universo 18 a 29 años cumplidos.** Es la definición usada por los programas de primer empleo. Excluye a
   adolescentes en escolaridad obligatoria, cuya situación laboral responde a otra lógica.
2. **No se eliminan los casos con `P21 = 0`.** Un ingreso cero identifica a desocupados e inactivos, población
   central del diagnóstico. Eliminarlos maquillaría el promedio de ingresos del segmento.
3. **`P21 = −9` y `PP3E_TOT = 999` se tratan como faltantes, no como ceros.** Son códigos de no respuesta del
   INDEC; computarlos como valores hundiría artificialmente los promedios.
4. **Decodificación con `BUSCARV` contra una hoja de equivalencias**, en lugar de reemplazar los códigos a mano:
   así el dato crudo queda intacto y la traducción es auditable y reutilizable.

## Semana 2 — Variables compuestas

5. **IPL — Índice de Precariedad Laboral (0 a 3).** Un punto por cada carencia:
   `PP07H = 2` (sin aportes jubilatorios), `PP07C ≥ 2` (inestabilidad contractual) y jornada atípica
   (`PP3E_TOT < 35` o `> 48` horas). Los tres umbrales replican criterios usuales en el análisis del mercado
   laboral: informalidad, temporalidad y sub/sobreocupación.
6. **El IPL queda en blanco para quien no está ocupado** (`=SI(ESTADO<>1;"";…)`). `PP07H` y `PP07C` solo se
   relevan a ocupados; asignarles 0 los contaría como empleo de calidad. Mismo criterio cuando la jornada es
   `999`: sin jornada informada no hay índice.
7. **ISE — Índice Socio-Educativo (0 a 3).** 2 puntos por secundaria completa o más y 1 punto adicional por
   estudios superiores. La escala se mantiene comparable con la del IPL para poder correlacionarlos.
   *Limitación asumida:* el MVP usa el nivel educativo de la propia persona joven como aproximación del clima
   educativo del hogar. Con la base de hogares correspondería usar el nivel del jefe o jefa de hogar.

**Frecuencias de validación**

| Valor del índice | IPL | ISE |
|---|---|---|
| 0 | 602 | 132 |
| 1 | 909 | 1.501 |
| 2 | 815 | 1.928 |
| 3 | 1.140 | 2.372 |
| **Total** | **3.466** (ocupados con jornada informada) | **5.933** (todos los jóvenes) |

## Semana 3 — Correlación y dashboard

8. **Columnas `P21_LIMPIO` y `HORAS_LIMPIO` para la matriz de correlación.** `COEF.DE.CORREL` descarta los pares
   donde alguna celda es texto o está vacía, de modo que el coeficiente se calcula únicamente sobre casos válidos
   y sin arrastrar los ceros estructurales de quienes no trabajan.
9. **Correlación de Pearson sobre cinco variables**: edad, ingreso, horas, IPL e ISE. Se documentan solo las
   relaciones con |r| ≥ 0,3.

| Par | Coeficiente | Lectura |
|---|---|---|
| IPL ↔ ingreso | **−0,446** | Cada punto de precariedad opera como penalidad salarial. |
| ISE ↔ ingreso | **+0,386** | Las credenciales se traducen en ingresos. |
| Horas ↔ ingreso | **+0,368** | La intensidad horaria pesa, pero menos que la educación. |
| ISE ↔ IPL | **−0,365** | A mayor nivel socio-educativo, menor precariedad. |
| Edad ↔ ingreso | +0,112 | Dentro de la franja 18-29 la experiencia casi no discrimina. |

10. **La tabla de brecha de género se construyó con `PROMEDIO.SI.CONJUNTO`** en lugar de una tabla dinámica
    estática: el resultado es idéntico y se recalcula solo al cambiar la base. En Excel puede reproducirse como
    tabla dinámica (CH04 en Filas, `REGION_TEXTO` en Columnas, IPL en Valores como Promedio) para sumarle
    segmentadores interactivos.
11. **Cautela con Patagonia.** Registra la mayor brecha de género (+0,424) pero es la región con menos casos
    (279 ocupados jóvenes): el informe la menciona con la advertencia correspondiente y apoya la recomendación
    en Gran Buenos Aires (+0,236), donde el respaldo muestral es mayor.

## Semana 4 — Informe y entrega

12. **Los ponderadores muestrales (`PONDERA`) no se aplicaron.** Las cifras describen la muestra, no la
    proyección poblacional. Queda registrado como próximo paso porque cambia la magnitud de los porcentajes,
    aunque no el signo de las correlaciones.
13. **Paleta de tres colores** (azul institucional, naranja de contraste y gris neutro), tipografía uniforme y
    ejes titulados en los tres gráficos, para que el informe se lea como un documento corporativo y no como una
    salida de software.
14. **Ningún número del informe se escribe a mano**: el PDF se genera desde el archivo de hallazgos que produce
    el pipeline, de modo que Excel, gráficos e informe no puedan desincronizarse.

## Transversal

15. **Se conserva una copia intacta del archivo original** en `01_Datos_Originales/`. Trazabilidad: cualquier
    resultado debe poder reproducirse desde el dato crudo.
16. **Base de trabajo simulada con la estructura oficial de la EPH.** El entorno no tenía acceso al portal del
    INDEC. Las fórmulas, el dashboard y el informe son definitivos; los valores se actualizan al reemplazar el
    `.txt` por el oficial (ver `01_Datos_Originales/LEEME_datos.md`).
