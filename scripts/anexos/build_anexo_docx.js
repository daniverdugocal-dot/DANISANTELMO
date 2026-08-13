const fs = require('fs');
const d = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, ImageRun, Footer, PageNumber
} = d;

const FONT = 'Calibri';
const ACCENT = '1F4E3D';
const GREY = '595959';
const W = 9020;

function leg(parts) {
  return new Paragraph({
    children: parts.map(p => new TextRun({ text: p.t, bold: !!p.b, font: FONT, size: 18, color: p.b ? '000000' : GREY })),
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 120, after: 100 },
  });
}

function src(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 16, italics: true, color: GREY })],
    spacing: { line: 240, before: 60, after: 260 },
  });
}

function cell(text, width, o = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill } : undefined,
    margins: { top: 55, bottom: 55, left: 100, right: 100 },
    borders: o.topRule ? { top: { style: BorderStyle.SINGLE, size: 8, color: ACCENT } } : undefined,
    children: [new Paragraph({
      children: [new TextRun({ text, font: FONT, size: 19, bold: !!o.bold, color: o.head ? 'FFFFFF' : '000000' })],
      alignment: o.align || AlignmentType.LEFT,
      spacing: { line: 240, after: 0 },
    })],
  });
}

const R = AlignmentType.RIGHT;
const widths = [4240, 2700, 2080];

const rows = [
  ['Superficie en producción', '1.500 ha', 'Caso DGI-445'],
  ['Productividad exportadora del sector', '4.167 cajas/ha', 'CANAPEP'],
  ['Volumen anual estimado', '6,25 M cajas (5,6–6,9)', 'Cálculo propio'],
  ['Precio por caja', '8,0 USD (7,5–9,0)', 'PROCOMER / mercado'],
  ['Ingresos anuales estimados', '50 M USD (42–62)', 'Cálculo propio'],
  ['Margen operativo de referencia', '12 % · 6,0 M USD', 'Supuesto propio'],
  ['Coste unitario de la banda individual', '0,50 USD/caja', 'Caso DGI-445'],
  ['Coste del co-branding en el 100 % del volumen', '3,1 M USD/año', 'Cálculo propio'],
  ['Coste del co-branding en el 25 % del volumen', '0,8 M USD/año', 'Cálculo propio'],
];

const trs = [new TableRow({
  tableHeader: true,
  children: [
    cell('Concepto', widths[0], { fill: ACCENT, head: true, bold: true }),
    cell('Valor', widths[1], { fill: ACCENT, head: true, bold: true, align: R }),
    cell('Origen', widths[2], { fill: ACCENT, head: true, bold: true }),
  ],
})];

rows.forEach((r, i) => {
  const last = i === rows.length - 1;
  trs.push(new TableRow({
    children: [
      cell(r[0], widths[0], { fill: i % 2 === 0 ? 'F2F5F3' : undefined, bold: last, topRule: last }),
      cell(r[1], widths[1], { fill: i % 2 === 0 ? 'F2F5F3' : undefined, align: R, bold: last, topRule: last }),
      cell(r[2], widths[2], { fill: i % 2 === 0 ? 'F2F5F3' : undefined, bold: last, topRule: last }),
    ],
  }));
});

const table = new Table({
  columnWidths: widths,
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'D9E2DD' },
    insideVertical: { style: BorderStyle.NONE },
  },
  rows: trs,
});

const imgW = 560, imgH = Math.round(560 * 0.335);

const children = [
  new Paragraph({
    children: [new TextRun({ text: 'ANEXO 1. Dimensionado económico de la división de piña de PCC Fresh', font: FONT, size: 24, bold: true, color: ACCENT })],
    spacing: { after: 40, line: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, space: 6, color: ACCENT } },
  }),
  leg([
    { t: 'Tabla A1.1. Volumen, ingresos y margen anuales estimados de PCC Fresh, y coste del esquema de co-branding según su alcance. ', b: true },
    { t: 'El caso no aporta información económica de la empresa, por lo que la estimación parte de las 1.500 hectáreas en producción y de la productividad exportadora del sector costarricense (175 millones de cajas de 12 kg sobre 42.000 hectáreas). La columna de origen distingue el dato público del supuesto propio. Escenario base; entre paréntesis, el rango de trabajo.' },
  ]),
  table,
  src('Fuentes: caso DGI-445; CANAPEP (2025); PROCOMER (2025); FAOSTAT (2024); cotización del mercado mayorista (2026).'),
  leg([
    { t: 'Figura A1.1. Margen operativo anual resultante según el alcance del esquema de co-branding. ', b: true },
    { t: 'Millones de dólares que restan del margen operativo estimado una vez descontado el coste de 0,50 dólares por caja, aplicado a la totalidad del volumen exportado o solo al canal seleccionado.' },
  ]),
  new Paragraph({
    children: [new ImageRun({
      type: 'png',
      data: fs.readFileSync(__dirname + '/chart.png'),
      transformation: { width: imgW, height: imgH },
    })],
    spacing: { before: 60, after: 60 },
  }),
  src('Elaboración propia a partir de la Tabla A1.1.'),
];

const doc = new Document({
  creator: 'Programa LYDES 2026',
  title: 'Anexo 1 — Dimensionado económico de PCC Fresh',
  styles: { default: { document: { run: { font: FONT, size: 22 }, paragraph: { spacing: { line: 240 } } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1250, right: 1418, bottom: 1020, left: 1418 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('escrito', b.length); });
