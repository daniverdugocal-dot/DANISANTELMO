// Presentación de defensa — ALCOPALET, El precio del crecimiento
// Daniel Verdugo Calvo · Programa LYDES · San Telmo Business School
// 12 diapositivas para una exposición de 10 minutos.

const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.3 x 7.5 pulgadas
pres.author = "Daniel Verdugo Calvo";
pres.title = "Alcopalet — El precio del crecimiento";

// ---------------------------------------------------------------- paleta
const OSCURO = "2B2118";   // madera oscura
const AMBAR = "C87F2A";    // acento
const AMBAR_C = "EFD9B8";  // acento claro
const GRIS = "6B6259";
const CLARO = "F7F5F2";
const BLANCO = "FFFFFF";
const ROJO = "B3261E";
const VERDE = "2E6F40";

const TIT = "Cambria";
const CPO = "Calibri";

const M = 0.7;             // margen lateral
const ANCHO = 13.33 - 2 * M;

// ---------------------------------------------------------------- helpers
function titulo(s, texto, sub) {
  s.addText(texto, {
    x: M, y: 0.45, w: ANCHO, h: 0.62, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 32, bold: true, color: OSCURO,
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.08, w: ANCHO, h: 0.34, isTextBox: true, margin: 0,
      fontFace: CPO, fontSize: 14, color: GRIS,
    });
  }
}

function circulo(s, n, x, y, d, relleno, colorTexto) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: relleno || AMBAR },
  });
  s.addText(String(n), {
    x, y, w: d, h: d, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 17, bold: true,
    color: colorTexto || BLANCO, align: "center", valign: "middle",
  });
}

function tarjeta(s, x, y, w, h, relleno) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: relleno || CLARO },
    line: { color: relleno === BLANCO ? "E3DED6" : relleno || CLARO, width: 1 },
  });
}

function pie(s, n) {
  s.addText("Alcopalet · El precio del crecimiento", {
    x: M, y: 6.92, w: 6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 9, color: "A79E93",
  });
  s.addText(String(n), {
    x: 13.33 - M - 1, y: 6.92, w: 1, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 9, color: "A79E93", align: "right",
  });
}

// ================================================================ 1 PORTADA
let s = pres.addSlide();
s.background = { color: OSCURO };
s.addText("CASO PRÁCTICO", {
  x: M, y: 2.0, w: ANCHO, h: 0.35, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 14, color: AMBAR, charSpacing: 5, bold: true,
});
s.addText("ALCOPALET", {
  x: M, y: 2.42, w: ANCHO, h: 1.15, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 60, bold: true, color: BLANCO,
});
s.addText("El precio del crecimiento", {
  x: M, y: 3.6, w: ANCHO, h: 0.55, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 26, italic: true, color: AMBAR_C,
});
s.addShape(pres.ShapeType.ellipse, { x: M, y: 4.62, w: 0.13, h: 0.13, fill: { color: AMBAR } });
s.addText("Daniel Verdugo Calvo", {
  x: M + 0.32, y: 4.5, w: 6, h: 0.35, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 15, color: BLANCO,
});
s.addText("Programa LYDES  ·  San Telmo Business School  ·  31 de agosto de 2026", {
  x: M, y: 5.0, w: ANCHO, h: 0.32, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, color: "A79E93",
});
s.addNotes("Buenos días. Voy a defender el caso Alcopalet. En diez minutos: cuál es el problema, qué decisión propongo y cómo se ejecuta.");

// ================================================================ 2 LA MESA
s = pres.addSlide();
titulo(s, "Tres asuntos sobre la mesa", "15 de diciembre de 2025 · Comité de dirección");

const asuntos = [
  ["Contrato Persán", "90.000 palés/año a 10 € fijos\ndurante 5 años · cobro a 180 días", "BORRADOR sin firmar"],
  ["Nueva máquina", "40.000 € · más flexibilidad\ny menos dependencia de mano de obra", "RESPUESTA HOY"],
  ["Reparto a los fundadores", "Los padres quieren recoger\nlos frutos de dieciocho años", "SIN RESOLVER"],
];
asuntos.forEach((a, i) => {
  const x = M + i * (ANCHO / 3);
  const w = ANCHO / 3 - 0.35;
  tarjeta(s, x, 1.85, w, 3.28, CLARO);
  circulo(s, i + 1, x + 0.35, 2.15, 0.52);
  s.addText(a[0], {
    x: x + 0.35, y: 2.82, w: w - 0.7, h: 0.78, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 18, bold: true, color: OSCURO,
  });
  s.addText(a[1], {
    x: x + 0.35, y: 3.62, w: w - 0.7, h: 1.0, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 13, color: GRIS, lineSpacing: 18,
  });
  s.addText(a[2], {
    x: x + 0.35, y: 4.68, w: w - 0.7, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 11, bold: true, color: AMBAR, charSpacing: 1,
  });
});
s.addText("Las dos primeras decisiones se han tratado como una sola. Tienen plazos distintos.", {
  x: M, y: 5.55, w: ANCHO, h: 0.45, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 18, italic: true, color: OSCURO,
});
pie(s, 2);
s.addNotes("No voy a contar el caso, que ya conocen. Solo subrayo una cosa: la máquina hay que decidirla hoy, el contrato es todavía un borrador. Son dos decisiones, no una.");

// ================================================================ 3 EL PROBLEMA
s = pres.addSlide();
titulo(s, "El problema no es de capacidad. Es de precio.", "Escandallo de costes unitarios · Anexo 3");

const filas = [
  ["Ingreso por palé", "10,00 €", false],
  ["Materia prima (madera y clavos)", "8,27 €", false],
  ["Mano de obra directa", "0,42 €", false],
  ["Transporte", "0,53 €", false],
  ["Otros costes variables", "1,30 €", false],
  ["COSTE UNITARIO", "10,52 €", true],
];
filas.forEach((f, i) => {
  const y = 1.9 + i * 0.46;
  if (f[2]) tarjeta(s, M, y - 0.05, 6.6, 0.46, AMBAR_C);
  s.addText(f[0], {
    x: M + 0.2, y, w: 4.6, h: 0.36, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 14, bold: f[2], color: f[2] ? OSCURO : GRIS,
  });
  s.addText(f[1], {
    x: M + 4.8, y, w: 1.6, h: 0.36, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 14, bold: f[2], color: OSCURO, align: "right",
  });
});

tarjeta(s, 8.1, 1.85, ANCHO - 7.4, 2.9, OSCURO);
s.addText("MARGEN POR PALÉ", {
  x: 8.45, y: 2.15, w: 4.3, h: 0.3, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, bold: true, color: AMBAR, charSpacing: 3,
});
s.addText("−0,52 €", {
  x: 8.45, y: 2.5, w: 4.3, h: 1.1, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 62, bold: true, color: "FF7B6B",
});
s.addText("Cada palé vendido a Persán resta margen.\nNi el volumen ni la máquina lo corrigen.", {
  x: 8.45, y: 3.72, w: 4.3, h: 0.8, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 13, color: AMBAR_C, lineSpacing: 19,
});

s.addText("La madera representa el 82,7 % del precio pactado, y el contrato lo fija durante cinco años sin revisión.", {
  x: M, y: 5.35, w: ANCHO, h: 0.4, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 14, color: GRIS,
});
pie(s, 3);
s.addNotes("El precio es de 10 € y el coste unitario 10,52 €. El margen es negativo antes de cualquier coste fijo. Y la madera es el 83 % del precio, así que ninguna mejora de productividad lo arregla.");

// ================================================================ 4 IMPACTO
s = pres.addSlide();
titulo(s, "Lo que costaría firmar tal como está", "Cuenta de resultados incremental del contrato · Anexo 6");

const comp = [
  ["Margen de contribución", "−46.800 €", "90.000 ud × (−0,52 €)"],
  ["Alquiler de la campa", "−60.000 €", "5.000 €/mes"],
  ["Leasing máquina + camión", "−48.000 €", "4.000 €/mes"],
  ["Coste del circulante", "−34.077 €", "454.356 € al 7,5 %"],
];
comp.forEach((c, i) => {
  const y = 1.95 + i * 0.72;
  tarjeta(s, M, y, 7.1, 0.62, CLARO);
  s.addText(c[0], {
    x: M + 0.25, y: y + 0.06, w: 3.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 14, bold: true, color: OSCURO,
  });
  s.addText(c[2], {
    x: M + 0.25, y: y + 0.33, w: 3.5, h: 0.24, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 11, color: GRIS,
  });
  s.addText(c[1], {
    x: M + 4.5, y: y + 0.15, w: 2.35, h: 0.35, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 17, bold: true, color: ROJO, align: "right",
  });
});

tarjeta(s, 8.35, 1.95, ANCHO - 7.65, 3.32, OSCURO);
s.addText("IMPACTO ANUAL", {
  x: 8.7, y: 2.25, w: 4.1, h: 0.3, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, bold: true, color: AMBAR, charSpacing: 3,
});
s.addText("−188.877 €", {
  x: 8.7, y: 2.6, w: 4.1, h: 1.0, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 44, bold: true, color: "FF7B6B",
});
s.addText("67,6 %", {
  x: 8.7, y: 3.72, w: 1.6, h: 0.55, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 30, bold: true, color: BLANCO,
});
s.addText("del beneficio de\ntoda la empresa", {
  x: 10.35, y: 3.78, w: 2.4, h: 0.6, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, color: AMBAR_C, lineSpacing: 15,
});
s.addText("944.383 € acumulados en los cinco años del contrato", {
  x: 8.7, y: 4.55, w: 4.1, h: 0.4, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, color: AMBAR_C,
});
pie(s, 4);
s.addNotes("Sumando campa, leasing y el coste de cobrar a 180 días, el contrato resta 188.877 € al año: dos tercios del beneficio de toda la empresa, no solo de la línea.");

// ================================================================ 5 OPCIONES
s = pres.addSlide();
titulo(s, "Cuatro alternativas", "Presentación y análisis de opciones");

const ops = [
  ["Aceptar el contrato", "tal como está", "−188.877 € al año", ROJO],
  ["Rechazarlo y no invertir", "mantener la situación actual", "Sin crecimiento", GRIS],
  ["Comprar solo el robot", "sin vincularlo a Persán", "171.262 € de ahorro", VERDE],
  ["Renegociar y después firmar", "12,50 € y cobro a 90 días", "+48.606 € al año", VERDE],
];
ops.forEach((o, i) => {
  const y = 1.85 + i * 1.18;
  const destacada = i === 3;
  tarjeta(s, M, y, ANCHO, 1.02, destacada ? AMBAR_C : CLARO);
  circulo(s, i + 1, M + 0.3, y + 0.25, 0.52, destacada ? OSCURO : AMBAR);
  s.addText(o[0], {
    x: M + 1.05, y: y + 0.16, w: 5.2, h: 0.36, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 19, bold: true, color: OSCURO,
  });
  s.addText(o[1], {
    x: M + 1.05, y: y + 0.56, w: 5.2, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 13, color: GRIS,
  });
  s.addText(o[2], {
    x: M + 6.6, y: y + 0.3, w: 5.2, h: 0.42, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 18, bold: true, color: o[3], align: "right",
  });
});
pie(s, 5);
s.addNotes("Cuatro alternativas viables. Las dos primeras son las que estaban sobre la mesa; las dos últimas salen del análisis. Descarté una quinta, esperar y buscar otra máquina, porque el fabricante exige respuesta hoy.");

// ================================================================ 6 CRITERIOS
s = pres.addSlide();
titulo(s, "Tres perspectivas para decidir", "Fijación de criterios");

const crits = [
  ["Económico-financiera", ["Rentabilidad de cada alternativa", "Recursos necesarios y efecto en la liquidez", "Que el crecimiento no deteriore el margen"]],
  ["Estratégica y de mercado", ["Contribución al crecimiento futuro", "Capacidad disponible para acompañarlo", "Riesgo de que Persán pese en exceso"]],
  ["Organizativa y humana", ["Nivel de servicio a los clientes actuales", "Aprovechar la plantilla sin despidos", "Compatibilidad con los intereses de la familia"]],
];
crits.forEach((c, i) => {
  const x = M + i * (ANCHO / 3);
  const w = ANCHO / 3 - 0.35;
  tarjeta(s, x, 1.85, w, 3.55, i === 0 ? OSCURO : CLARO);
  circulo(s, i + 1, x + 0.35, 2.15, 0.5, i === 0 ? AMBAR : AMBAR);
  s.addText(c[0], {
    x: x + 0.35, y: 2.82, w: w - 0.7, h: 0.7, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 18, bold: true, color: i === 0 ? BLANCO : OSCURO,
  });
  s.addText(c[1].map((t, k) => ({
    text: t, options: { bullet: true, breakLine: k < c[1].length - 1 },
  })), {
    x: x + 0.35, y: 3.55, w: w - 0.7, h: 1.65, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 12.5, color: i === 0 ? AMBAR_C : GRIS,
    paraSpaceAfter: 8,
  });
});
pie(s, 6);
s.addNotes("La decisión se valora desde tres perspectivas. La económico-financiera es la que más pesa, porque el problema es de margen y de caja.");

// ================================================================ 7 RECHAZOS
s = pres.addSlide();
titulo(s, "Por qué se descartan tres", "Elección justificada · rechazo de alternativas");

const rech = [
  ["1", "Aceptar tal como está", "Sería crecer destruyendo valor: el precio de 10 € es inferior al coste de 10,52 €.", "−188.877 €"],
  ["2", "Rechazar y no invertir", "No resuelve las limitaciones de capacidad y espacio, y renuncia a 99.670 palés de recorrido en palé usado.", "Sin futuro"],
  ["3", "Solo el robot", "Favorable en lo productivo, pero desaprovecha la oportunidad comercial si Persán puede ser rentable.", "Plan B"],
];
rech.forEach((r, i) => {
  const y = 1.9 + i * 1.42;
  tarjeta(s, M, y, ANCHO, 1.22, CLARO);
  circulo(s, r[0], M + 0.32, y + 0.35, 0.52, GRIS);
  s.addText(r[1], {
    x: M + 1.1, y: y + 0.2, w: 4.3, h: 0.36, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 18, bold: true, color: OSCURO,
  });
  s.addText(r[2], {
    x: M + 1.1, y: y + 0.6, w: 7.6, h: 0.55, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 13, color: GRIS, lineSpacing: 17,
  });
  s.addText(r[3], {
    x: M + 9.0, y: y + 0.38, w: 2.8, h: 0.45, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 17, bold: true, color: i === 0 ? ROJO : GRIS, align: "right",
  });
});
pie(s, 7);
s.addNotes("La 1 destruye valor. La 2 no prepara a la empresa para crecer. La 3 es buena, pero desaprovecha a Persán si se puede hacer rentable; la mantengo como plan B.");

// ================================================================ 8 DECISIÓN
s = pres.addSlide();
s.background = { color: OSCURO };
s.addText("LA DECISIÓN", {
  x: M, y: 0.85, w: ANCHO, h: 0.35, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 13, bold: true, color: AMBAR, charSpacing: 5,
});
s.addText("Invertir en el robot y renegociar antes de firmar", {
  x: M, y: 1.25, w: ANCHO - 0.5, h: 1.4, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 36, bold: true, color: BLANCO,
});

const dec = [
  ["Precio", "10,00 €", "12,50 €"],
  ["Plazo de cobro", "180 días", "90 días"],
  ["Margen por palé", "−0,52 €", "+1,98 €"],
];
dec.forEach((d, i) => {
  const x = M + i * (ANCHO / 3);
  const w = ANCHO / 3 - 0.4;
  s.addText(d[0].toUpperCase(), {
    x, y: 2.85, w, h: 0.3, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 11, bold: true, color: AMBAR, charSpacing: 2,
  });
  s.addText(d[1], {
    x, y: 3.2, w: w * 0.45, h: 0.62, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 22, color: "8A8079", strike: true,
  });
  s.addText("→", {
    x: x + w * 0.45, y: 3.26, w: 0.42, h: 0.5, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 20, color: AMBAR, align: "center",
  });
  s.addText(d[2], {
    x: x + w * 0.45 + 0.42, y: 3.2, w: w * 0.52, h: 0.62, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 22, bold: true, color: BLANCO,
  });
});

s.addText("El robot se compra hoy: 40.000 € que se justifican por flexibilidad y se recuperan en 84 días.\nLa firma queda condicionada a que Persán acepte unas condiciones rentables.", {
  x: M, y: 4.35, w: ANCHO - 0.5, h: 0.9, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 15, color: AMBAR_C, lineSpacing: 24,
});
s.addNotes("La decisión: comprar el robot hoy y condicionar la firma a renegociar precio y plazo de cobro. Son dos decisiones separadas, con plazos distintos.");

// ================================================================ 9 EFECTO
s = pres.addSlide();
titulo(s, "El efecto de renegociar", "Impacto anual del contrato · Anexo 8");

s.addChart(pres.ChartType.bar, [{
  name: "Impacto anual (€)",
  labels: ["Contrato actual", "Contrato renegociado"],
  values: [-188877, 48606],
}], {
  x: M, y: 1.75, w: 7.5, h: 3.6,
  barDir: "col", chartColors: [ROJO, VERDE], varyColors: true,
  showValue: true, dataLabelPosition: "outEnd",
  dataLabelFontFace: CPO, dataLabelFontSize: 13, dataLabelColor: OSCURO,
  dataLabelFormatCode: '#,##0 "€";-#,##0 "€"',
  showLegend: false, showTitle: false,
  catAxisLabelFontFace: CPO, catAxisLabelFontSize: 13, catAxisLabelColor: OSCURO,
  valAxisLabelFontFace: CPO, valAxisLabelFontSize: 10, valAxisLabelColor: GRIS,
  valGridLine: { color: "E8E3DB", size: 1 },
  catGridLine: { style: "none" },
  valAxisMaxVal: 120000, valAxisMinVal: -240000,
  barGapWidthPct: 110,
});

tarjeta(s, 8.7, 1.9, ANCHO - 8.0, 1.55, AMBAR_C);
s.addText("MEJORA ANUAL", {
  x: 9.0, y: 2.12, w: 3.7, h: 0.28, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 11, bold: true, color: OSCURO, charSpacing: 2,
});
s.addText("+237.483 €", {
  x: 9.0, y: 2.42, w: 3.7, h: 0.8, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 34, bold: true, color: OSCURO,
});

const efectos = [
  ["Margen por palé", "de −0,52 € a +1,98 €"],
  ["Circulante inmovilizado", "se reduce a la mitad"],
  ["Resultado de la empresa", "vuelve a crecer"],
];
efectos.forEach((e, i) => {
  const y = 3.7 + i * 0.62;
  s.addShape(pres.ShapeType.ellipse, { x: 8.75, y: y + 0.13, w: 0.11, h: 0.11, fill: { color: AMBAR } });
  s.addText(e[0], {
    x: 9.0, y, w: 3.7, h: 0.28, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 13, bold: true, color: OSCURO,
  });
  s.addText(e[1], {
    x: 9.0, y: y + 0.26, w: 3.7, h: 0.28, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 12, color: GRIS,
  });
});
pie(s, 9);
s.addNotes("Con 12,50 € y cobro a 90 días el contrato pasa de restar 188.877 € a aportar 48.606 €. La mejora es de 237.483 € anuales.");

// ================================================================ 10 ROBOT
s = pres.addSlide();
titulo(s, "El robot se justifica por sí solo", "Inversión independiente del contrato · Anexo 7");

const cifras = [
  ["40.000 €", "de inversión", "Frente a un beneficio\nde 279.364 €"],
  ["84 días", "de recuperación", "Sobre el ahorro neto\nde 171.262 € anuales"],
  ["5 operarios", "liberados", "Sin despidos: se recolocan\nen palé usado"],
];
cifras.forEach((c, i) => {
  const x = M + i * (ANCHO / 3);
  const w = ANCHO / 3 - 0.35;
  tarjeta(s, x, 1.85, w, 2.35, CLARO);
  s.addText(c[0], {
    x: x + 0.35, y: 2.12, w: w - 0.7, h: 0.82, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 32, bold: true, color: AMBAR,
  });
  s.addText(c[1], {
    x: x + 0.35, y: 2.99, w: w - 0.7, h: 0.32, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 14, bold: true, color: OSCURO,
  });
  s.addText(c[2], {
    x: x + 0.35, y: 3.35, w: w - 0.7, h: 0.7, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 12, color: GRIS, lineSpacing: 16,
  });
});

tarjeta(s, M, 4.55, ANCHO, 1.25, OSCURO);
s.addText("Lo que compra la máquina es flexibilidad", {
  x: M + 0.4, y: 4.76, w: 6.3, h: 0.44, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 19, bold: true, color: BLANCO,
});
s.addText("El cambio de formato pasa de 300 a 25 minutos, lo que hace viables las series cortas.", {
  x: M + 0.4, y: 5.18, w: 6.2, h: 0.4, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 13, color: AMBAR_C,
});
s.addText("300 min", {
  x: M + 7.0, y: 4.95, w: 1.7, h: 0.5, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 26, color: "8A8079", strike: true, align: "right",
});
s.addText("→", {
  x: M + 8.75, y: 5.02, w: 0.5, h: 0.4, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 20, color: AMBAR, align: "center",
});
s.addText("25 min", {
  x: M + 9.3, y: 4.95, w: 2.5, h: 0.5, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 26, bold: true, color: BLANCO,
});
pie(s, 10);
s.addNotes("La máquina no se compra por capacidad. Se compra por flexibilidad y por reducir la dependencia de una mano de obra difícil de contratar. Se paga en 84 días.");

// ================================================================ 11 PLAN
s = pres.addSlide();
titulo(s, "Plan de acción", "Secuenciación temporal de actuaciones");

const fases = [
  ["Negociación y decisión", "Confirmar el robot y negociar precio y plazo. La firma queda condicionada a la rentabilidad."],
  ["Implantación productiva", "Instalación, pruebas por formato y operario formado. Campa y transporte, solo si hay acuerdo."],
  ["Reorganización de recursos", "Reasignar sin despidos a palé usado y revisar la distribución del espacio."],
  ["Seguimiento y control", "Cuadro de mando mensual y revisión a los tres y seis meses."],
  ["Consolidación", "Usar la capacidad para ampliar cartera y evitar la dependencia de un solo cliente."],
];
fases.forEach((f, i) => {
  const y = 1.9 + i * 0.98;
  circulo(s, i + 1, M, y + 0.09, 0.5);
  if (i < fases.length - 1) {
    s.addShape(pres.ShapeType.line, {
      x: M + 0.25, y: y + 0.59, w: 0, h: 0.48,
      line: { color: AMBAR_C, width: 2 },
    });
  }
  s.addText(`Fase ${i + 1}. ${f[0]}`, {
    x: M + 0.85, y: y + 0.02, w: 11.2, h: 0.34, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 17, bold: true, color: OSCURO,
  });
  s.addText(f[1], {
    x: M + 0.85, y: y + 0.38, w: 11.2, h: 0.34, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 12.5, color: GRIS,
  });
});
pie(s, 11);
s.addNotes("Cinco fases. La clave está en la primera: el robot se confirma hoy, pero la campa y el camión no se contratan hasta que haya firma.");

// ================================================================ 12 CIERRE
s = pres.addSlide();
s.background = { color: OSCURO };
s.addText("¿Y SI PERSÁN NO ACEPTA?", {
  x: M, y: 1.15, w: ANCHO, h: 0.35, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 13, bold: true, color: AMBAR, charSpacing: 5,
});
s.addText("Se rechaza el contrato y se mantiene la inversión", {
  x: M, y: 1.55, w: ANCHO - 0.5, h: 1.25, isTextBox: true, margin: 0,
  fontFace: TIT, fontSize: 34, bold: true, color: BLANCO,
});
s.addText("La máquina deja de depender de Persán y pasa a formar parte de la modernización de la estructura productiva. La capacidad liberada se orienta a captar clientes rentables.", {
  x: M, y: 2.88, w: 11.3, h: 0.85, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 15, color: AMBAR_C, lineSpacing: 23,
});

const cierre = [
  ["Crecer", "sí, pero con margen"],
  ["Invertir", "sí, pero sin atarlo al contrato"],
  ["Repartir", "sí, cuando el negocio genere caja"],
];
cierre.forEach((c, i) => {
  const x = M + i * (ANCHO / 3);
  const w = ANCHO / 3 - 0.4;
  s.addShape(pres.ShapeType.ellipse, { x, y: 4.25, w: 0.13, h: 0.13, fill: { color: AMBAR } });
  s.addText(c[0], {
    x, y: 4.5, w, h: 0.45, isTextBox: true, margin: 0,
    fontFace: TIT, fontSize: 24, bold: true, color: BLANCO,
  });
  s.addText(c[1], {
    x, y: 4.98, w, h: 0.4, isTextBox: true, margin: 0,
    fontFace: CPO, fontSize: 14, color: AMBAR_C,
  });
});

s.addText("Daniel Verdugo Calvo  ·  Programa LYDES  ·  San Telmo Business School", {
  x: M, y: 6.35, w: ANCHO, h: 0.32, isTextBox: true, margin: 0,
  fontFace: CPO, fontSize: 12, color: "8A8079",
});
s.addNotes("Si Persán no acepta, se rechaza y se conserva la máquina. Crecer sí, pero con margen. Invertir sí, pero sin atar la inversión al contrato. Y repartir sí, cuando el negocio genere caja. Muchas gracias.");

pres.writeFile({ fileName: "/home/user/DANISANTELMO/Alcopalet_Defensa.pptx" })
  .then(f => console.log("Generado:", f));
