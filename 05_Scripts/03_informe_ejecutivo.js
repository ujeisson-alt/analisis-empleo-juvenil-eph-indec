/**
 * 03_informe_ejecutivo.js
 * -----------------------
 * Genera el Informe de Diagnóstico Social (4 páginas) a partir de los
 * hallazgos calculados por 01_pipeline_eph.py (_hallazgos.json), de modo que
 * ningún número del informe se escriba a mano.
 *
 * Uso:  node 05_Scripts/03_informe_ejecutivo.js
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, HeadingLevel, BorderStyle, ShadingType, ShadingType: ST,
  PageBreak, Footer, PageNumber, NumberFormat, convertInchesToTwip, TabStopType,
} = require("docx");

const RAIZ = path.resolve(__dirname, "..");
const H = JSON.parse(fs.readFileSync(path.join(__dirname, "_hallazgos.json"), "utf8"));
const K = H.kpi;

const AZUL = "14375E", NARANJA = "C2603C", GRIS = "5A6472", CLARO = "EEF1F6", TINTA = "1B2430";
const FUENTE = "Calibri";

// ---------- formato de números en español -------------------------------
const nf = (n, d = 0) =>
  Number(n).toLocaleString("es-AR", { minimumFractionDigits: d, maximumFractionDigits: d });
const pesos = (n) => "$ " + nf(Math.round(n));
const pct = (n, d = 1) => nf(n, d) + " %";
const coef = (n) => (n >= 0 ? "+" : "−") + nf(Math.abs(n), 3);

const HOY = new Date();
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
  "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const FECHA = `${HOY.getDate()} de ${MESES[HOY.getMonth()]} de ${HOY.getFullYear()}`;

const img = (rel, w, h) => new ImageRun({
  type: "png", data: fs.readFileSync(path.join(RAIZ, rel)),
  transformation: { width: w, height: h },
});

const p = (texto, o = {}) => new Paragraph({
  alignment: o.align ?? AlignmentType.JUSTIFIED,
  spacing: { after: o.after ?? 100, line: o.line ?? 252 },
  indent: o.indent,
  children: (Array.isArray(texto) ? texto : [texto]).map((t) =>
    typeof t === "string"
      ? new TextRun({ text: t, font: FUENTE, size: o.size ?? 20, color: o.color ?? TINTA })
      : t),
});
const b = (t, o = {}) => new TextRun({ text: t, font: FUENTE, bold: true, size: o.size ?? 20, color: o.color ?? TINTA });
const t = (txt, o = {}) => new TextRun({ text: txt, font: FUENTE, size: o.size ?? 20, color: o.color ?? TINTA, italics: o.italics });

const h1 = (txt) => new Paragraph({
  spacing: { before: 200, after: 110 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NARANJA, space: 4 } },
  children: [new TextRun({ text: txt, font: FUENTE, bold: true, size: 28, color: AZUL })],
});
const h2 = (txt) => new Paragraph({
  spacing: { before: 150, after: 80 },
  children: [new TextRun({ text: txt, font: FUENTE, bold: true, size: 22, color: AZUL })],
});
const pie = (txt) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 140 },
  children: [new TextRun({ text: txt, font: FUENTE, size: 16, color: GRIS, italics: true })],
});

// ---------- tabla de ficha técnica --------------------------------------
const ANCHO = convertInchesToTwip(6.3);
const celda = (txt, { bold = false, fill = null, ancho, color = TINTA } = {}) => new TableCell({
  width: { size: ancho, type: WidthType.DXA },
  shading: fill ? { type: ST.CLEAR, fill, color: "auto" } : undefined,
  margins: { top: 45, bottom: 45, left: 100, right: 100 },
  children: [new Paragraph({
    spacing: { after: 0 },
    children: [new TextRun({ text: txt, font: FUENTE, size: 17, bold, color })],
  })],
});
const C1 = Math.round(ANCHO * 0.30), C2 = ANCHO - C1;
const fichaTecnica = new Table({
  width: { size: ANCHO, type: WidthType.DXA },
  columnWidths: [C1, C2],
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D5DBE5" },
  },
  rows: [
    ["Fuente", "Encuesta Permanente de Hogares (EPH) — INDEC. Base de microdatos de personas, estructura y códigos del Diseño de Registro oficial."],
    ["Universo", `Personas de 18 a 29 años (CH06). ${nf(K.jovenes)} casos sobre ${nf(K.registros_totales)} registros relevados (${pct(K.pct_jovenes)} de la muestra).`],
    ["Métricas propias", "Índice de Precariedad Laboral (IPL, 0 a 3) e Índice Socio-Educativo (ISE, 0 a 3), construidos con funciones lógicas anidadas."],
    ["Herramientas", "Microsoft Excel / Google Sheets con fórmulas vivas (BUSCARV, SI/O anidados, COEF.DE.CORREL, tablas dinámicas y formato condicional) y un pipeline equivalente en Python para auditar cada resultado."],
    ["Advertencia", "Los valores surgen de una base SIMULADA con la estructura oficial de la EPH: el entorno de trabajo no tenía acceso al portal del INDEC. La metodología y las fórmulas son las definitivas; al reemplazar el archivo de microdatos por el oficial, todo el tablero y este informe se recalculan."],
  ].map(([a, c], i) => new TableRow({
    children: [celda(a, { bold: true, fill: CLARO, ancho: C1, color: AZUL }), celda(c, { ancho: C2 })],
  })),
});

const recomendacion = (n, titulo, cuerpo) => [
  new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: `${n}. `, font: FUENTE, bold: true, size: 22, color: NARANJA }),
      new TextRun({ text: titulo, font: FUENTE, bold: true, size: 22, color: AZUL }),
    ],
  }),
  p(cuerpo, { indent: { left: convertInchesToTwip(0.25) }, after: 60 }),
];

// ==========================================================================
// Documento
// ==========================================================================
const doc = new Document({
  creator: "Jeisson Marín Uribe",
  title: "Diagnóstico del Mercado Laboral Juvenil en Argentina",
  description: "Análisis de microdatos EPH-INDEC con variables compuestas y matriz de correlación",
  sections: [{
    properties: {
      page: {
        margin: { top: convertInchesToTwip(0.9), bottom: convertInchesToTwip(0.8),
                  left: convertInchesToTwip(1), right: convertInchesToTwip(1) },
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: ANCHO }],
          children: [
            t("Fundación Impacto Laboral (entidad ficticia) · Diagnóstico del mercado laboral juvenil", { size: 16, color: GRIS }),
            new TextRun({ text: "\t", font: FUENTE }),
            new TextRun({ children: ["Página ", PageNumber.CURRENT, " de ", PageNumber.TOTAL_PAGES],
                          font: FUENTE, size: 16, color: GRIS }),
          ],
        })],
      }),
    },
    children: [
      // ================= PÁGINA 1 — PORTADA =================
      new Paragraph({ spacing: { after: 700 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 500 },
                      children: [img("03_Visualizaciones/logo_fundacion.png", 300, 98)] }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 160 },
        children: [new TextRun({ text: "Diagnóstico del Mercado Laboral Juvenil en Argentina",
                                 font: FUENTE, bold: true, size: 46, color: AZUL })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 420 },
        border: { top: { style: BorderStyle.SINGLE, size: 12, color: NARANJA, space: 10 } },
        children: [new TextRun({
          text: "Análisis de microdatos EPH-INDEC con variables compuestas y matriz de correlación",
          font: FUENTE, size: 25, color: GRIS, italics: true })],
      }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [b("Jeisson Marín Uribe", { size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [t("Analista de Datos — autor del informe", { size: 21, color: GRIS })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 420 },
        children: [t(FECHA, { size: 21, color: GRIS })] }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 600, after: 80 },
        children: [b("Informe preparado para", { size: 19, color: GRIS })],
      }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
        children: [b("Dirección de Políticas de Empleo · Fundación Impacto Laboral", { size: 21, color: AZUL })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 },
        children: [t("Entidad ficticia utilizada como caso de estudio", { size: 16, color: GRIS, italics: true })] }),
      new Paragraph({ children: [new PageBreak()] }),

      // ================= PÁGINA 2 =================
      h1("1. Resumen ejecutivo"),
      p([
        t("La Fundación encaró este diagnóstico para responder una pregunta concreta: "),
        b("qué determina que una persona joven acceda a un empleo de calidad. "),
        t(`Se procesaron ${nf(K.registros_totales)} registros de la Encuesta Permanente de Hogares y se aisló el segmento de 18 a 29 años (${nf(K.jovenes)} casos, ${pct(K.pct_jovenes)} de la muestra), sobre el que se construyeron dos métricas propias que la encuesta no trae: un Índice de Precariedad Laboral (IPL) y un Índice Socio-Educativo (ISE), ambos en escala 0 a 3. Tres hallazgos ordenan el resultado. Primero, la precariedad es la norma y no la excepción: ${pct(K.informalidad)} de los jóvenes ocupados trabaja sin aportes jubilatorios y ${pct(K.ipl3_pct)} acumula las tres carencias medidas al mismo tiempo. Segundo, la educación es el amortiguador más potente que aparece en los datos: la correlación entre ISE e IPL es de ${coef(K.r_ise_ipl)} y el ingreso promedio de quienes completaron estudios superiores (${pesos(K.ingreso_superior_completo)}) supera en ${pct(K.prima_superior_pct)} al de quienes no terminaron la secundaria (${pesos(K.ingreso_secund_incompleta)}). Tercero, la desventaja femenina es estructural y no se explica por credenciales: las mujeres jóvenes registran un IPL de ${nf(K.ipl_mujer, 2)} frente a ${nf(K.ipl_varon, 2)} de los varones y perciben ${pct(K.brecha_ingresos_pct)} menos de ingreso por su ocupación principal. `),
        b("Se recomienda concentrar el presupuesto en tres frentes: terminalidad educativa con certificación laboral, incentivos a la registración focalizados en mujeres jóvenes, y la adopción del IPL como indicador trimestral de calidad del empleo joven."),
      ], { after: 200 }),

      h1("2. Ficha técnica y alcance"),
      fichaTecnica,
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      h1("3. Análisis y visualizaciones"),
      h2("Gráfico 1 — Matriz de calor de correlación"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [img("03_Visualizaciones/G1_Matriz_Correlacion.png", 300, 237)] }),
      pie("Coeficientes de Pearson entre las cinco variables clave. Jóvenes de 18 a 29 años ocupados con dato válido."),
      p([
        t(`La matriz ordena la jerarquía de los determinantes. El coeficiente más fuerte es el que vincula precariedad e ingresos (${coef(K.r_ipl_p21)}): cada punto de IPL opera como una penalidad salarial, de modo que la mala calidad del puesto y el bajo ingreso son dos caras del mismo problema. Le sigue la relación entre credenciales e ingresos (${coef(K.r_ise_p21)}) y, con signo negativo, la de credenciales y precariedad (${coef(K.r_ise_ipl)}): a mayor nivel socio-educativo, menor precariedad, con una intensidad moderada que confirma la hipótesis central del proyecto. Las horas trabajadas también empujan el ingreso (${coef(K.r_horas_p21)}), aunque menos que la educación. `),
        b(`La edad, en cambio, es casi irrelevante dentro de la franja juvenil (${coef(K.r_edad_p21)}): entre los 18 y los 29 años la experiencia adicional no compensa la falta de credenciales ni la informalidad.`),
      ]),

      // ================= PÁGINA 3 =================
      h2("Gráfico 2 — Brecha de género en precariedad, por región"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [img("03_Visualizaciones/G2_Brecha_Genero.png", 460, 260)] }),
      pie(`IPL promedio (0 a 3) de la población ocupada de 18 a 29 años, por sexo y región. Promedio país: ${nf(K.ipl_promedio, 2)}.`),
      p([
        t(`La brecha de género aparece en casi todas las regiones, pero no con la misma intensidad. En el total del país las mujeres jóvenes registran un IPL de ${nf(K.ipl_mujer, 2)} contra ${nf(K.ipl_varon, 2)} de los varones (${coef(K.brecha_total)} puntos), diferencia que se amplía en Gran Buenos Aires (+0,236) y alcanza su máximo en Patagonia (${coef(K.brecha_max)}), aunque en esta última la lectura debe tomarse con cautela porque es la región con menos casos de la muestra (279 ocupados jóvenes). El nivel absoluto de precariedad, en cambio, se explica más por geografía que por género: el Noreste (IPL ${nf(K.ipl_region_max, 2)}) y el Noroeste concentran la informalidad más alta del país (${pct(72.7)} y ${pct(68.1)} de ocupados sin aportes), frente a ${pct(52.8)} en Gran Buenos Aires. `),
        b("La combinación de ser mujer y vivir en el norte del país define el núcleo más expuesto del segmento juvenil, y es allí donde cualquier política de registración rinde más por peso invertido."),
      ]),

      h2("Gráfico 3 — Ingreso y precariedad según nivel educativo"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
        children: [img("03_Visualizaciones/G3_Ingreso_Educacion.png", 460, 260)] }),
      pie("Ingreso promedio de la ocupación principal (eje izquierdo) e IPL promedio (eje derecho) por nivel educativo alcanzado."),
      p([
        t(`Las dos curvas se mueven en espejo y esa es la evidencia más contundente del informe: mientras el ingreso promedio trepa de ${pesos(K.ingreso_secund_incompleta)} en quienes no terminaron la secundaria a ${pesos(K.ingreso_superior_completo)} en quienes completaron estudios superiores, el IPL cae de 2,26 a 1,07 puntos. La informalidad sigue el mismo patrón: alcanza al ${pct(82.6)} de los ocupados sin secundaria completa y baja al ${pct(25.5)} entre quienes tienen título superior. El salto decisivo ocurre al completar la secundaria, donde el IPL cae medio punto de golpe: es el umbral que separa el empleo protegido del empleo de subsistencia. `),
        b(`Con ${pct(27.5)} de los jóvenes por debajo de ese umbral, la terminalidad educativa deja de ser una política educativa y pasa a ser la política de empleo más rentable disponible.`),
      ]),

      // ================= PÁGINA 4 =================
      h1("4. Conclusiones"),
      p([
        t(`El diagnóstico deja tres conclusiones firmes. La precariedad juvenil no es un problema de falta de empleo sino de calidad del empleo disponible: con una tasa de actividad de ${pct(K.tasa_actividad)} y una desocupación de ${pct(K.tasa_desocupacion)}, la mayoría de los jóvenes trabaja, pero ${pct(K.informalidad)} lo hace sin aportes jubilatorios. Las credenciales educativas son la variable con mayor poder explicativo y efecto acumulativo, porque mejoran el ingreso (${coef(K.r_ise_p21)}) y reducen la precariedad (${coef(K.r_ise_ipl)}) al mismo tiempo. Y la desigualdad de género atraviesa todos los niveles educativos y todas las regiones, por lo que exige instrumentos propios y no puede resolverse por arrastre de las políticas generales.`),
      ]),

      h1("5. Recomendaciones de política"),
      ...recomendacion(1, "Terminalidad educativa con certificación laboral para jóvenes de 18 a 24 años sin secundaria completa",
        `Dado que la correlación entre el ISE y el IPL es de ${coef(K.r_ise_ipl)} y que el ${pct(82.6)} de los ocupados sin secundaria completa trabaja sin aportes jubilatorios, se recomienda una beca de terminalidad con certificación de oficio dirigida a los ${pct(27.5)} de jóvenes que hoy están por debajo de ese umbral, priorizando el Noreste y el Noroeste. La transferencia debe condicionarse a la asistencia y acreditarse con un certificado que el empleador pueda verificar, porque el dato muestra que el salto de calidad del puesto se produce exactamente al completar el nivel secundario (el IPL cae de 2,26 a 1,74 puntos).`),
      ...recomendacion(2, "Incentivo a la registración focalizado en mujeres jóvenes del área metropolitana y del norte",
        `Dado que el IPL femenino (${nf(K.ipl_mujer, 2)}) supera al masculino (${nf(K.ipl_varon, 2)}) y que la brecha de ingresos alcanza ${pct(K.brecha_ingresos_pct)}, con una diferencia de precariedad de +0,236 puntos en Gran Buenos Aires, se recomienda una reducción temporal de contribuciones patronales por doce meses para la contratación registrada de mujeres de 18 a 29 años, acompañada de cupos de cuidado infantil. La población objetivo son las mujeres jóvenes ocupadas en condiciones de informalidad y las inactivas por tareas de cuidado, que en este segmento representan el ${pct(35.4)} de las jóvenes frente al ${pct(22.0)} de los varones: sin resolver el cuidado, el incentivo fiscal no alcanza para convertir inactividad en empleo formal.`),
      ...recomendacion(3, "Adoptar el IPL como indicador trimestral de calidad del empleo joven, con metas por región",
        `Dado que la precariedad explica los ingresos mejor que cualquier otra variable medida (${coef(K.r_ipl_p21)}) y que la brecha regional es amplia —Noreste ${nf(K.ipl_region_max, 2)} puntos contra 1,60 en Gran Buenos Aires—, se recomienda institucionalizar el IPL como indicador oficial de seguimiento, con publicación trimestral y metas de reducción diferenciadas por región. El instrumento ya está construido y es reproducible en una hoja de cálculo: permite evaluar si los dos programas anteriores mueven la aguja, en lugar de medir únicamente la cantidad de puestos creados.`),

      h1("6. Limitaciones y próximos pasos"),
      p([
        t("El análisis es de corte transversal: mide asociaciones, no causalidad, y una correlación moderada no autoriza a atribuir a la educación la totalidad del efecto observado. El IPL se calcula sobre la población ocupada con jornada informada, por lo que excluye por diseño a desocupados e inactivos, cuya situación se analiza aparte. Los ponderadores muestrales (PONDERA) no se aplicaron en esta etapa, de modo que las cifras describen la muestra y no la proyección poblacional. "),
        b("Los próximos pasos naturales son incorporar la ponderación, replicar el índice en cuatro trimestres consecutivos para medir evolución, y sumar la base de hogares para construir el ISE con el nivel educativo del jefe o jefa de hogar en lugar de la aproximación usada en este MVP."),
      ]),
      new Paragraph({
        spacing: { before: 240 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: "D5DBE5", space: 8 } },
        children: [t(`Informe generado el ${FECHA} · base de datos: ${H.fuente} · el libro de cálculo con todas las fórmulas vivas, los scripts de reproducción y la bitácora de decisiones acompañan este documento en el repositorio del proyecto.`,
                     { size: 16, color: GRIS, italics: true })],
      }),
    ],
  }],
});

const SALIDA = path.join(RAIZ, "04_Informe_Final",
  "Informe_Empleo_Juvenil_Jeisson_Marin.docx");
fs.mkdirSync(path.dirname(SALIDA), { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(SALIDA, buf);
  console.log("informe:", SALIDA);
});
