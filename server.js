require('dotenv').config();
const express = require('express');
const multer = require('multer');
const PDFParser = require('pdf2json');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, HeadingLevel, AlignmentType, WidthType, BorderStyle } = require('docx');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

app.use(express.static('public'));

function extraerTextoPDF(buffer) {
  return new Promise((resolve, reject) => {
    const parser = new PDFParser();
    parser.on('pdfParser_dataReady', (data) => {
      const texto = data.Pages.map(page =>
        page.Texts.map(t => decodeURIComponent(t.R.map(r => r.T).join(''))).join(' ')
      ).join('\n');
      resolve(texto);
    });
    parser.on('pdfParser_dataError', reject);
    parser.parseBuffer(buffer);
  });
}

app.post('/analizar', upload.single('pdf'), async (req, res) => {
  try {
    console.log('Archivo recibido:', req.file?.originalname);
    const texto = await extraerTextoPDF(req.file.buffer);

    const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

    const prompt = `Eres un arquitecto técnico experto en mediciones y presupuestos de obra en España.

Analiza esta memoria descriptiva y devuelve ÚNICAMENTE un JSON puro sin markdown, sin explicaciones, solo el JSON:

{
  "proyecto": {
    "titulo": "título del proyecto",
    "promotor": "nombre",
    "emplazamiento": "dirección",
    "pem": 0,
    "plazo": "x días"
  },
  "mediciones": [
    {
      "ref": "V01",
      "planta": "Planta baja",
      "tipo": "Ventana",
      "ancho": 1.00,
      "alto": 1.10,
      "uds": 1,
      "material": "descripción material"
    }
  ],
  "presupuesto": {
    "capitulos": [
      {
        "codigo": "C01",
        "nombre": "Nombre capítulo",
        "partidas": [
          {
            "codigo": "C01.01",
            "descripcion": "descripción partida",
            "uds": 1.0,
            "unidad": "ud",
            "precio_unitario": 100.00,
            "observaciones": "nota breve"
          }
        ]
      }
    ]
  }
}

Precios realistas para zona Valencia 2026. El total de partidas debe aproximarse al PEM indicado en la memoria. Incluye capítulos de demolición, carpintería, acristalamiento, ayudas albañilería, gestión residuos y seguridad y salud.

MEMORIA:
${texto}`;

    const result = await model.generateContent(prompt);
    const responseText = result.response.text().trim();
    const clean = responseText.replace(/```json|```/g, '').trim();
    const jsonStart = clean.indexOf('{');
    const jsonEnd = clean.lastIndexOf('}');
    const parsed = JSON.parse(clean.substring(jsonStart, jsonEnd + 1));

    res.json({ ok: true, data: parsed });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.post('/exportar-word', express.json({ limit: '10mb' }), async (req, res) => {
  try {
    const d = req.body;
    const p = d.proyecto;

    const borderNone = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const borderThin = { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' };
    const borderHeader = { style: BorderStyle.SINGLE, size: 8, color: '0A0A0A' };

    const cellHeader = (text, shade) => new TableCell({
      shading: { fill: shade || '0A0A0A' },
      borders: { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone },
      children: [new Paragraph({
        children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 18, font: 'Calibri' })],
      })],
    });

    const cellData = (text, opts = {}) => new TableCell({
      shading: { fill: opts.shade || 'FFFFFF' },
      borders: { top: borderThin, bottom: borderThin, left: borderNone, right: borderNone },
      children: [new Paragraph({
        alignment: opts.right ? AlignmentType.RIGHT : AlignmentType.LEFT,
        children: [new TextRun({
          text: String(text),
          size: 18,
          bold: opts.bold || false,
          color: opts.color || '1a1a1a',
          font: 'Calibri'
        })],
      })],
    });

    // TABLA MEDICIONES
    const rowsMed = [
      new TableRow({
        children: ['Ref.','Planta','Tipo','Ancho m','Alto m','Uds.','Sup. m²','Material']
          .map(h => cellHeader(h))
      })
    ];
    let totalSup = 0;
    d.mediciones.forEach(m => {
      const sup = +(m.ancho * m.alto * m.uds).toFixed(2);
      totalSup += sup;
      rowsMed.push(new TableRow({
        children: [
          cellData(m.ref, { color: '6B7280' }),
          cellData(m.planta),
          cellData(m.tipo, { bold: true }),
          cellData(m.ancho.toFixed(2), { right: true }),
          cellData(m.alto.toFixed(2), { right: true }),
          cellData(m.uds, { right: true }),
          cellData(sup.toFixed(2), { right: true, bold: true }),
          cellData(m.material, { color: '6B7280' }),
        ]
      }));
    });
    rowsMed.push(new TableRow({
      children: [
        cellData('Total superficie carpintería', { shade: '0A0A0A', color: 'FFFFFF', bold: true }),
        cellData('', { shade: '0A0A0A' }),
        cellData('', { shade: '0A0A0A' }),
        cellData('', { shade: '0A0A0A' }),
        cellData('', { shade: '0A0A0A' }),
        cellData('', { shade: '0A0A0A' }),
        cellData(totalSup.toFixed(2) + ' m²', { shade: '0A0A0A', color: 'FFFFFF', bold: true, right: true }),
        cellData('', { shade: '0A0A0A' }),
      ]
    }));

    // TABLA PRESUPUESTO
    const rowsPres = [
      new TableRow({
        children: ['Cód.','Descripción','Uds.','Ud.','P. unit.','Importe','Observaciones']
          .map(h => cellHeader(h))
      })
    ];
    let totalGlobal = 0;
    d.presupuesto.capitulos.forEach(cap => {
      let totalCap = 0;
      rowsPres.push(new TableRow({
        children: [
          cellData(cap.codigo, { shade: 'F3F4F6', color: '6B7280', bold: true }),
          cellData(cap.nombre, { shade: 'F3F4F6', bold: true }),
          cellData('', { shade: 'F3F4F6' }),
          cellData('', { shade: 'F3F4F6' }),
          cellData('', { shade: 'F3F4F6' }),
          cellData('', { shade: 'F3F4F6' }),
          cellData('', { shade: 'F3F4F6' }),
        ]
      }));
      cap.partidas.forEach(pp => {
        const imp = +(pp.uds * pp.precio_unitario).toFixed(2);
        totalCap += imp;
        rowsPres.push(new TableRow({
          children: [
            cellData(pp.codigo, { color: '9CA3AF' }),
            cellData(pp.descripcion),
            cellData(Number.isInteger(pp.uds) ? pp.uds : pp.uds.toFixed(2), { right: true }),
            cellData(pp.unidad, { color: '6B7280' }),
            cellData(pp.precio_unitario.toFixed(2) + ' €', { right: true }),
            cellData(imp.toFixed(2) + ' €', { right: true, bold: true }),
            cellData(pp.observaciones, { color: '6B7280' }),
          ]
        }));
      });
      totalGlobal += totalCap;
      rowsPres.push(new TableRow({
        children: [
          cellData('', { shade: 'F9FAFB' }),
          cellData('', { shade: 'F9FAFB' }),
          cellData('', { shade: 'F9FAFB' }),
          cellData('', { shade: 'F9FAFB' }),
          cellData('Total ' + cap.codigo, { shade: 'F9FAFB', color: '6B7280', right: true }),
          cellData(totalCap.toFixed(2) + ' €', { shade: 'F9FAFB', bold: true, right: true }),
          cellData('', { shade: 'F9FAFB' }),
        ]
      }));
    });

    const pec = +(totalGlobal * 1.19).toFixed(2);
    const pca = +(pec * 1.21).toFixed(2);
    [
      ['PEM', totalGlobal.toFixed(2) + ' €'],
      ['GG 13% + BI 6%', (totalGlobal * 0.19).toFixed(2) + ' €'],
      ['PEC', pec.toFixed(2) + ' €'],
      ['IVA 21%', (pec * 0.21).toFixed(2) + ' €'],
      ['PCA (contrata)', pca.toFixed(2) + ' €'],
    ].forEach(([label, val], i) => {
      const isTotal = i === 4;
      rowsPres.push(new TableRow({
        children: [
          cellData('', { shade: isTotal ? '0A0A0A' : 'F5F5F5' }),
          cellData('', { shade: isTotal ? '0A0A0A' : 'F5F5F5' }),
          cellData('', { shade: isTotal ? '0A0A0A' : 'F5F5F5' }),
          cellData('', { shade: isTotal ? '0A0A0A' : 'F5F5F5' }),
          cellData(label, { shade: isTotal ? '0A0A0A' : 'F5F5F5', color: isTotal ? 'FFFFFF' : '6B7280', bold: isTotal, right: true }),
          cellData(val, { shade: isTotal ? '0A0A0A' : 'F5F5F5', color: isTotal ? 'FFFFFF' : '1a1a1a', bold: true, right: true }),
          cellData('', { shade: isTotal ? '0A0A0A' : 'F5F5F5' }),
        ]
      }));
    });

    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } },
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'ArquIA — Informe técnico', bold: true, size: 32, font: 'Calibri', color: '0A0A0A' })],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [new TextRun({ text: p.promotor + ' — ' + p.emplazamiento, size: 20, font: 'Calibri', color: '6B7280' })],
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [new TextRun({ text: 'PEM: ' + p.pem.toLocaleString('es-ES') + ' €   |   Plazo: ' + p.plazo, size: 18, font: 'Calibri', color: '6B7280' })],
            spacing: { after: 400 },
          }),
          new Paragraph({
            children: [new TextRun({ text: 'Mediciones de carpintería', bold: true, size: 24, font: 'Calibri' })],
            spacing: { after: 200 },
          }),
          new Table({ rows: rowsMed, width: { size: 100, type: WidthType.PERCENTAGE } }),
          new Paragraph({ children: [], spacing: { after: 400 } }),
          new Paragraph({
            children: [new TextRun({ text: 'Presupuesto por capítulos', bold: true, size: 24, font: 'Calibri' })],
            spacing: { after: 200 },
          }),
          new Table({ rows: rowsPres, width: { size: 100, type: WidthType.PERCENTAGE } }),
          new Paragraph({ children: [], spacing: { after: 400 } }),
          new Paragraph({
            children: [new TextRun({ text: 'ArquIA es una herramienta de asistencia técnica. Los resultados son orientativos y no sustituyen el criterio del técnico competente.', size: 16, font: 'Calibri', color: '9CA3AF', italics: true })],
          }),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="arquia-informe.docx"');
    res.send(buffer);

  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});
app.listen(3000, () => console.log('Servidor en http://localhost:3000'));