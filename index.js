const fs = require('fs');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { PDFDocument, rgb } = require('pdf-lib');

// Ruta del archivo Excel de entrada y PDF de salida
const inputExcelPath = 'ruta/del/archivo.xlsx';
const outputPdfPath = 'ruta/del/resultado.pdf';

// Función para leer el archivo Excel
async function readExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  return data;
}

// Función para generar el PDF
async function generatePdf(data) {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([600, 400]);
  const { height } = page.getSize();

  const font = await pdfDoc.embedFont(PDFDocument.Font.Helvetica);
  const fontSize = 12;
  const lineHeight = 20;
  const textX = 50;
  const textY = height - 50;

  data.forEach((row, rowIndex) => {
    const text = row.join('\t');
    const textWidth = font.widthOfTextAtSize(text, fontSize);
    const textHeight = font.heightAtSize(fontSize);

    page.drawText(text, {
      x: textX,
      y: textY - (rowIndex * lineHeight) - textHeight,
      size: fontSize,
      font: font,
      color: rgb(0, 0, 0),
    });
  });

  const pdfBytes = await pdfDoc.save();
  return pdfBytes;
}

// Función para guardar el PDF
async function savePdf(outputPath, pdfBytes) {
  fs.writeFileSync(outputPath, pdfBytes);
  console.log('PDF generado con éxito:', outputPath);
}

// Ejecutar el generador
(async () => {
  try {
    const excelData = await readExcel(inputExcelPath);
    const pdfBytes = await generatePdf(excelData);
    await savePdf(outputPdfPath, pdfBytes);
  } catch (error) {
    console.error('Error al generar el PDF:', error.message);
  }
})();
