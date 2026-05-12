const fs = require('node:fs')
const path = require('node:path')

const mainPath = path.join(__dirname, '..', 'electron', 'main.cjs')
const rendererPath = path.join(__dirname, '..', 'electron', 'report-pdf.cjs')
const packagePath = path.join(__dirname, '..', 'package.json')
const oldTemplatePath = path.join(__dirname, '..', 'src', 'ReportPdfDocument.tsx')

const mainSource = fs.readFileSync(mainPath, 'utf8')
const rendererSource = fs.readFileSync(rendererPath, 'utf8')
const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf8'))

if (fs.existsSync(oldTemplatePath)) {
  throw new Error('The old React PDF-only template must not be used for final export.')
}

if (!packageJson.dependencies?.pdfkit) {
  throw new Error('Professional PDF export must depend on PDFKit.')
}

if (!/const \{ renderReportPdfBuffer \} = require\('\.\/report-pdf\.cjs'\)/.test(mainSource)) {
  throw new Error('Electron export must import the measured PDFKit renderer.')
}

if (!/const pdf = await renderReportPdfBuffer\(printReport\)/.test(mainSource)) {
  throw new Error('pdf:export must use the measured PDFKit renderer instead of Chromium printToPDF.')
}

if (!/embedEditableReportData\(pdf,\s*printReport\)/.test(mainSource)) {
  throw new Error('PDF export must keep embedding editable report data.')
}

const exportHandler = /ipcMain\.handle\('pdf:export'[\s\S]*?\n\}\)/.exec(mainSource)?.[0] || ''
if (/printToPDF|renderReportPdf\(/.test(exportHandler)) {
  throw new Error('pdf:export must not call browser print rendering.')
}

const requiredRendererParts = [
  ['PDFDocument', 'PDFKit renderer must create real PDF documents.'],
  ['heightOfString', 'PDFKit renderer must measure text with real PDF font metrics.'],
  ['drawTextCard', 'PDFKit renderer must use reusable measured text cards.'],
  ['splitBlockToFit', 'PDFKit renderer must split long text by measured height.'],
  ['drawPhotoPair', 'PDFKit renderer must render controlled exhibit photo pairs.'],
  ['drawEquipmentCard', 'PDFKit renderer must render protected equipment cards.'],
  ['drawSignatureBlock', 'PDFKit renderer must render a professional signature block.'],
  ['finalizePageNumbers', 'PDFKit renderer must add final page numbers after layout.'],
]

for (const [token, message] of requiredRendererParts) {
  if (!rendererSource.includes(token)) throw new Error(message)
}

if (!/link:\s*url/.test(rendererSource) || !/Purchase Link/.test(rendererSource)) {
  throw new Error('Purchase links must be clickable PDF annotations.')
}

if (!/sectionIV\(\)[\s\S]*this\.addPage\('IV', 'Evidentiary Documentation'\)/.test(rendererSource)) {
  throw new Error('Each exhibit must start from a controlled PDF page.')
}

if (!/this\.addPage\('V', 'Clinical Justification Display'\)/.test(rendererSource)) {
  throw new Error('Clinical justification must start on its own controlled PDF page.')
}

if (!/this\.addPage\('VI', 'Procurement Notes & Justification Summary'\)/.test(rendererSource)) {
  throw new Error('Procurement summary must start on its own controlled PDF page.')
}

console.log('Report layout check passed: final export uses measured PDFKit layout, clickable links, protected cards, page numbers, and editable import metadata.')
