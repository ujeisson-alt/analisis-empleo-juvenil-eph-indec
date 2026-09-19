# Datos de origen

## Qué hay en esta carpeta

| Archivo | Descripción |
|---|---|
| `usu_individual_T125_SIMULADO.txt` | Base de **personas SIMULADA**, delimitada por `;`, con la estructura, los nombres de columna y los códigos del Diseño de Registro oficial de la EPH (INDEC). 32.000 registros. |

> Esta carpeta representa el dato crudo: **nunca se modifica**. Todo el trabajo ocurre en `02_Datos_Procesados/`.

## Por qué la base es simulada

El entorno en el que se construyó el proyecto no tenía acceso de red al portal del INDEC, de modo que no fue
posible descargar el `.zip` oficial. Para no detener el desarrollo, se generó una base con:

- los mismos **nombres de columna** y **códigos** del Diseño de Registro (`CH04`, `CH06`, `NIVEL_ED`, `ESTADO`, `PP07H`, `PP07C`, `PP3E_TOT`, `P21`, `REGION`…);
- los mismos **códigos de no respuesta** (`-9` en ingresos, `999` en horas) y celdas vacías donde la encuesta no releva la pregunta;
- distribuciones calibradas con valores públicos de referencia del mercado laboral argentino (tasas de actividad,
  desocupación e informalidad juvenil, brecha de género, estructura educativa y peso relativo de cada región).

El generador es auditable: [`../05_Scripts/00_generar_base_simulada.py`](../05_Scripts/00_generar_base_simulada.py).

**Ningún valor puede citarse como dato oficial del INDEC.**

## Cómo trabajar con los datos reales

1. Entrar a [www.indec.gob.ar](https://www.indec.gob.ar) → **Estadísticas Sociales** → **Hogares** →
   **Encuesta Permanente de Hogares (EPH)** → sección **Bases de microdatos**.
2. Descargar el `.zip` del último trimestre publicado. Adentro hay dos archivos: **personas**
   (`usu_individual_TXXX.txt`, el que se usa acá) y **hogares**.
3. Descargar también el **Diseño de Registro** (diccionario de variables). Es obligatorio para interpretar los códigos.
4. Copiar `usu_individual_TXXX.txt` en esta carpeta.
5. Recalcular todo:

   ```bash
   python3 05_Scripts/01_pipeline_eph.py 01_Datos_Originales/usu_individual_TXXX.txt
   node 05_Scripts/03_informe_ejecutivo.js
   ```

   El libro de Excel, los tres gráficos y el informe se regeneran con los datos oficiales.

6. Si preferís hacerlo a mano en Excel: **Datos → Desde texto/CSV**, delimitador `;`, y pegar los registros
   filtrados por `CH06` entre 18 y 29 en la hoja `Jóvenes_18_29` del libro de trabajo. Todas las fórmulas de
   las columnas derivadas, la matriz de correlación y el dashboard se recalculan solos.

## Variables utilizadas

El detalle completo está en [`../docs/diccionario_variables.md`](../docs/diccionario_variables.md).
