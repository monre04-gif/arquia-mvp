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

    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

    const prompt = `Eres un arquitecto técnico experto en mediciones y presupuestos de obra en España.

Analiza esta memoria descriptiva, detecta automáticamente el tipo de obra y genera mediciones y presupuesto adaptados a ese tipo de intervención.

INSTRUCCIONES:
1. Lee la memoria y determina el tipo de obra (carpintería, reforma interior, cubierta, fachada, instalaciones, obra nueva, etc.)
2. Extrae TODAS las unidades de obra mencionadas en la memoria con sus dimensiones cuando estén disponibles
3. Genera el presupuesto con los capítulos y partidas específicos para ese tipo de obra
4. Adapta los precios a la zona geográfica del proyecto

Devuelve ÚNICAMENTE un JSON puro sin markdown, sin explicaciones, solo el JSON:

{
  "proyecto": {
    "titulo": "título del proyecto",
    "tipo_obra": "tipo detectado ej: Sustitución carpintería exterior",
    "promotor": "nombre",
    "emplazamiento": "dirección completa",
    "municipio": "municipio",
    "provincia": "provincia",
    "zona_climatica": "zona climática según municipio ej: B3",
    "pem": 0,
    "plazo": "x días/semanas"
  },
  "mediciones": [
    {
      "ref": "P01",
      "planta": "Planta baja o zona",
      "tipo": "tipo de elemento",
      "ancho": 1.00,
      "alto": 1.00,
      "uds": 1,
      "material": "descripción del material o elemento"
    }
  ],
  "presupuesto": {
    "capitulos": [
      {
        "codigo": "C01",
        "nombre": "Nombre del capítulo según tipo de obra",
        "partidas": [
          {
            "codigo": "C01.01",
            "descripcion": "descripción de la partida",
            "uds": 1.0,
            "unidad": "ud/m²/ml/pa",
            "precio_unitario": 100.00,
            "observaciones": "nota breve"
          }
        ]
      }
    ]
  }
}

IMPORTANTE:
- Detecta el tipo de obra y adapta los capítulos a ese tipo específico
- Si hay dimensiones en la memoria úsalas, si no estímalas según lo descrito
- Precios realistas para la zona geográfica y tipo de obra en 2026
- El total debe aproximarse al PEM indicado en la memoria
- Si no hay PEM indicado, estímalo según el alcance de la obra

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
app.post('/verificar-cte', express.json({ limit: '10mb' }), async (req, res) => {
  try {
    const d = req.body;
    const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

    const prompt = `Eres un arquitecto técnico experto en normativa española, especialmente en el Código Técnico de la Edificación (CTE).

Analiza los siguientes datos de un proyecto y verifica el cumplimiento del DB-HE 1 y DB-HS 3.

DATOS DEL PROYECTO:
- Promotor: ${d.proyecto.promotor}
- Emplazamiento: ${d.proyecto.emplazamiento}
- Municipio: ${d.proyecto.municipio || 'ver emplazamiento'}
- Tipo de obra: Sustitución de carpintería exterior en vivienda unifamiliar
- Zona climática: detectar automáticamente según el municipio indicado en el emplazamiento del proyecto
- Mediciones de carpintería: ${JSON.stringify(d.mediciones)}
- Material especificado: Carpintería PVC imitación madera, vidrio doble bajo emisivo con control solar

Devuelve ÚNICAMENTE un JSON puro sin markdown:
{
  "resumen": "conclusión general en 2 frases",
  "zona_climatica": "zona detectada ej: B3",
  "db_he1": {
    "cumple": true,
    "transmitancia_maxima": "2.70 W/m²K",
    "transmitancia_estimada": "1.8-2.2 W/m²K",
    "factor_solar_requerido": "≤ 0.60",
    "superficie_total_huecos": 0,
    "porcentaje_fachada": "estimado",
    "observaciones": "texto",
    "advertencias": ["advertencia 1"]
  },
  "db_hs3": {
    "cumple": null,
    "caudal_minimo_requerido": "según estancias",
    "observaciones": "texto",
    "advertencias": ["advertencia 1"]
  },
  "recomendaciones": ["recomendación 1", "recomendación 2"],
  "texto_justificacion": "texto completo listo para copiar en la memoria, redactado en español técnico formal, mínimo 150 palabras"
}`;

    const result = await model.generateContent(prompt);
    const responseText = result.response.text().trim();
    const clean = responseText.replace(/```json|```/g, '').trim();
    const jsonStart = clean.indexOf('{');
    const jsonEnd = clean.lastIndexOf('}');
    const parsed = JSON.parse(clean.substring(jsonStart, jsonEnd + 1));

    res.json({ ok: true, data: parsed });
  } catch(e) {
    res.json({ ok: false, error: e.message });
  }
});
app.listen(3000, () => console.log('Servidor en http://localhost:3000'));