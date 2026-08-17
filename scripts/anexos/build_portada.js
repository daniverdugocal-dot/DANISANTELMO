const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
  Footer, PageNumber, PositionalTab, PositionalTabAlignment, PositionalTabLeader,
  TabStopType, HeadingLevel
} = d;

const FONT = 'Calibri';
const GREY = '767171';
const INK = '000000';

const gap = (n = 1) => Array.from({ length: n }, () =>
  new Paragraph({ children: [], spacing: { after: 0, line: 300 } }));

// ---------- PORTADA ----------

const portada = [
  ...gap(2),

  // hueco para el logo
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 200, line: 240 },
    border: {
      top:    { style: BorderStyle.DASHED, size: 4, color: 'BFBFBF', space: 18 },
      bottom: { style: BorderStyle.DASHED, size: 4, color: 'BFBFBF', space: 18 },
      left:   { style: BorderStyle.DASHED, size: 4, color: 'BFBFBF', space: 18 },
      right:  { style: BorderStyle.DASHED, size: 4, color: 'BFBFBF', space: 18 },
    },
    children: [new TextRun({
      text: 'Insertar aquí el logotipo de LYDES',
      font: FONT, size: 18, color: 'A6A6A6', italics: true,
    })],
  }),

  ...gap(4),

  // referencia del caso
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 260, line: 240 },
    children: [new TextRun({
      text: 'CASO DGI-445  ·  EL VALOR DE LA SOSTENIBILIDAD. PIÑAS CULTIVADAS DE COSTA RICA',
      font: FONT, size: 17, color: GREY, bold: true,
    })],
  }),

  // título propio
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120, line: 300 },
    children: [new TextRun({
      text: 'Más natural, mismo precio',
      font: FONT, size: 56, bold: true, color: INK,
    })],
  }),

  // filete
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000', space: 6 } },
    children: [new TextRun({ text: '', font: FONT, size: 2 })],
  }),

  // subtítulo
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 0, line: 280 },
    children: [new TextRun({
      text: 'Estrategia de marca para la piña libre de pesticidas de PCC Fresh',
      font: FONT, size: 26, color: INK,
    })],
  }),

  ...gap(9),

  // bloque de autoría
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 40, line: 260 },
    children: [new TextRun({ text: 'Trabajo realizado por:', font: FONT, size: 22, color: GREY })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 160, line: 260 },
    children: [new TextRun({ text: 'Daniel Verdugo Calvo', font: FONT, size: 26, bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 20, line: 260 },
    children: [new TextRun({ text: 'Programa LYDES 2026  ·  Fundación San Telmo', font: FONT, size: 20, color: GREY })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 0, line: 260 },
    children: [new TextRun({ text: 'Sevilla, [fecha de entrega]', font: FONT, size: 20, color: GREY })],
  }),

  new Paragraph({ children: [new d.PageBreak()] }),
];

// ---------- ÍNDICE ----------

function entrada(texto, pagina, negrita = false) {
  return new Paragraph({
    spacing: { after: 90, line: 260 },
    children: [
      new TextRun({ text: texto, font: FONT, size: 22, bold: negrita }),
      new TextRun({
        children: [new PositionalTab({
          alignment: PositionalTabAlignment.RIGHT,
          relativeTo: 'margin',
          leader: PositionalTabLeader.DOT,
        })],
        font: FONT, size: 22,
      }),
      new TextRun({ text: String(pagina), font: FONT, size: 22, bold: negrita }),
    ],
  });
}

const indice = [
  new Paragraph({
    spacing: { after: 60, line: 260 },
    children: [new TextRun({ text: 'ÍNDICE', font: FONT, size: 28, bold: true })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000', space: 6 } },
  }),
  ...gap(1),
  entrada('1.  Introducción', 3),
  entrada('2.  Resumen ejecutivo', 3),
  entrada('3.  Identificación del problema planteado', 4),
  entrada('4.  Presentación y análisis de opciones', 4),
  entrada('5.  Fijación de criterios', 5),
  entrada('6.  Elección justificada de una opción', 6),
  entrada('     a.  Elección justificada', 6),
  entrada('     b.  Rechazo justificado de las otras', 6),
  entrada('7.  Conclusión: decisión y plan de acción', 7),
  entrada('8.  Anexos', 9),
  entrada('     Anexo 1.  Análisis DAFO de PCC Fresh', 9),
  entrada('     Anexo 2.  Matriz de evaluación de alternativas', 9),
  entrada('     Anexo 3.  Cronograma del plan de acción', 10),
  entrada('     Anexo 4.  Concepto de comunicación en el punto de venta', 10),
  ...gap(2),
  new Paragraph({
    spacing: { line: 260 },
    children: [new TextRun({
      text: 'Nota: los números de página son orientativos; deben actualizarse una vez montado el documento definitivo.',
      font: FONT, size: 17, italics: true, color: GREY,
    })],
  }),
];

// ---------- DOCUMENTO ----------

const doc = new Document({
  creator: 'Daniel Verdugo Calvo',
  title: 'Más natural, mismo precio — Caso DGI-445',
  styles: { default: { document: { run: { font: FONT, size: 22 }, paragraph: { spacing: { line: 240 } } } } },
  sections: [{
    properties: {
      titlePage: true,                                   // primera página diferente: sin numerar
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 },
      },
    },
    footers: {
      first: new Footer({ children: [new Paragraph({ children: [] })] }),
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY })],
        })],
      }),
    },
    children: [...portada, ...indice],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(process.argv[2], b);
  console.log('escrito', b.length, 'bytes');
});
