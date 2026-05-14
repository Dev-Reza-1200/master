const { app, BrowserWindow, dialog, ipcMain, Menu, nativeImage } = require('electron')
const path = require('node:path')
const fs = require('node:fs')
const crypto = require('node:crypto')
const os = require('node:os')
const zlib = require('node:zlib')
const { fileURLToPath, pathToFileURL } = require('node:url')
const { fork } = require('node:child_process')
const AdmZip = require('adm-zip')
const pdfParse = require('pdf-parse')
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js')
const mammoth = require('mammoth')
const XLSX = require('xlsx')
const { PNG } = require('pngjs')
const { renderReportPdfBuffer } = require('./report-pdf.cjs')

let mainWindow
let db

const devServerUrl = process.env.ERGONOMIC_DEV_SERVER_URL
const gotSingleInstanceLock = app.requestSingleInstanceLock()
const editablePdfSchema = 'com.ergonomicworks.report.v1'
const editablePdfStartMarker = '%ERGONOMIC_REPORT_DATA_BEGIN'
const editablePdfEndMarker = '%ERGONOMIC_REPORT_DATA_END'

function getStartupLogPath() {
  const dir = app.isReady() ? app.getPath('userData') : (process.env.TEMP || process.cwd())
  return path.join(dir, 'ergonomic-report-startup.log')
}

function log(message, error) {
  const detail = error ? `\n${error.stack || error.message || error}` : ''
  try {
    fs.appendFileSync(getStartupLogPath(), `[${new Date().toISOString()}] ${message}${detail}\n`)
  } catch {
    // Logging should never prevent the desktop app from opening.
  }
}

process.on('uncaughtException', (error) => log('Uncaught exception', error))
process.on('unhandledRejection', (error) => log('Unhandled rejection', error))

if (!gotSingleInstanceLock) {
  log('Another app instance already has the lock; quitting.')
  app.quit()
} else {
  app.on('second-instance', () => {
    if (!mainWindow) return
    if (mainWindow.isMinimized()) mainWindow.restore()
    mainWindow.show()
    mainWindow.focus()
  })
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true })
}

function getAppIcon() {
  const icons = [
    path.join(__dirname, '..', 'dist', 'assets', 'app-icon.ico'),
    path.join(__dirname, '..', 'public', 'assets', 'app-icon.ico'),
  ]
  return icons.find((iconPath) => fs.existsSync(iconPath))
}

function walkFiles(dir) {
  if (!fs.existsSync(dir)) return []
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(dir, entry.name)
    if (entry.isDirectory()) return walkFiles(fullPath)
    const stat = fs.statSync(fullPath)
    return [{
      name: entry.name,
      path: fullPath,
      relativePath: path.relative(dir, fullPath),
      size: stat.size,
      modifiedAt: stat.mtime.toISOString(),
    }]
  })
}

function addFileIfExists(zip, zipPath, filePath) {
  if (!fs.existsSync(filePath)) return false
  zip.addFile(zipPath, fs.readFileSync(filePath))
  return true
}

function mimeForFile(filePath) {
  const ext = path.extname(filePath || '').toLowerCase()
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg'
  if (ext === '.png') return 'image/png'
  if (ext === '.webp') return 'image/webp'
  if (ext === '.gif') return 'image/gif'
  if (ext === '.svg') return 'image/svg+xml'
  if (ext === '.pdf') return 'application/pdf'
  if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  if (ext === '.doc') return 'application/msword'
  if (ext === '.xlsx') return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  if (ext === '.xls') return 'application/vnd.ms-excel'
  if (ext === '.csv') return 'text/csv'
  if (ext === '.txt') return 'text/plain'
  if (ext === '.rtf') return 'application/rtf'
  return 'application/octet-stream'
}

function filePathFromPhoto(photo) {
  if (!photo) return ''
  if (photo.path && fs.existsSync(photo.path)) return photo.path
  if (typeof photo.url === 'string' && photo.url.startsWith('file://')) {
    try {
      const decoded = fileURLToPath(photo.url)
      return fs.existsSync(decoded) ? decoded : ''
    } catch (error) {
      log('Unable to decode photo file URL', error)
    }
  }
  return ''
}

function makePortablePhoto(photo) {
  if (!photo) return photo
  if (typeof photo.url === 'string' && photo.url.startsWith('data:')) {
    return { ...photo, path: '' }
  }

  const filePath = filePathFromPhoto(photo)
  if (!filePath) return photo

  try {
    const data = fs.readFileSync(filePath)
    return {
      name: photo.name || path.basename(filePath),
      path: '',
      url: `data:${mimeForFile(filePath)};base64,${data.toString('base64')}`,
    }
  } catch (error) {
    log('Unable to embed photo into editable PDF payload', error)
    return photo
  }
}

function imageFromPhoto(photo) {
  if (!photo) return null
  if (typeof photo.url === 'string' && photo.url.startsWith('data:')) {
    try {
      const image = nativeImage.createFromDataURL(photo.url)
      return image.isEmpty() ? null : image
    } catch (error) {
      log('Unable to read photo data URL', error)
      return null
    }
  }

  const filePath = filePathFromPhoto(photo)
  if (!filePath) return null
  const image = nativeImage.createFromPath(filePath)
  return image.isEmpty() ? null : image
}

function optimizedPhoto(photo, maxSide = 1200) {
  if (!photo) return photo
  const image = imageFromPhoto(photo)
  if (!image) return photo

  const size = image.getSize()
  if (!size.width || !size.height) return photo
  const scale = Math.min(1, maxSide / Math.max(size.width, size.height))
  const width = Math.max(1, Math.round(size.width * scale))
  const height = Math.max(1, Math.round(size.height * scale))
  const resized = scale < 1 ? image.resize({ width, height, quality: 'best' }) : image
  const data = resized.toJPEG(82)

  const cacheDir = path.join(app.getPath('userData'), 'print-cache')
  ensureDir(cacheDir)
  const fileName = `${Date.now()}-${crypto.randomUUID()}-${(photo.name || 'photo').replace(/[^a-z0-9-_]+/gi, '-').slice(0, 40)}.jpg`
  const optimizedPath = path.join(cacheDir, fileName)
  fs.writeFileSync(optimizedPath, data)
  return {
    name: photo.name || fileName,
    path: optimizedPath,
    url: pathToFileURL(optimizedPath).href,
  }
}

function optimizeReportForPdf(report) {
  const clone = JSON.parse(JSON.stringify(report || {}))
  if (Array.isArray(clone.exhibits)) {
    clone.exhibits = clone.exhibits.map((exhibit) => ({
      ...exhibit,
      pre: optimizedPhoto(exhibit?.pre),
      post: optimizedPhoto(exhibit?.post),
    }))
  }
  clone.signature = optimizedPhoto(clone.signature, 900)
  return clone
}

function makePortableReport(report) {
  const clone = JSON.parse(JSON.stringify(report || {}))
  if (Array.isArray(clone.exhibits)) {
    clone.exhibits = clone.exhibits.map((exhibit) => ({
      ...exhibit,
      pre: makePortablePhoto(exhibit?.pre),
      post: makePortablePhoto(exhibit?.post),
    }))
  }
  clone.signature = makePortablePhoto(clone.signature)
  return clone
}

function buildEditablePdfPayload(report) {
  const payload = {
    schema: editablePdfSchema,
    exportedAt: new Date().toISOString(),
    appVersion: app.getVersion(),
    report: makePortableReport(report),
  }
  const compressed = zlib.gzipSync(Buffer.from(JSON.stringify(payload), 'utf8')).toString('base64')
  const lines = compressed.match(/.{1,96}/g) || []
  return Buffer.from(`${editablePdfStartMarker}\n${lines.map((line) => `%${line}`).join('\n')}\n${editablePdfEndMarker}\n`, 'latin1')
}

function embedEditableReportData(pdf, report) {
  if (!report) return pdf
  const payload = buildEditablePdfPayload(report)
  const eofMarker = Buffer.from('%%EOF', 'latin1')
  const eofIndex = pdf.lastIndexOf(eofMarker)
  if (eofIndex === -1) return Buffer.concat([pdf, Buffer.from('\n', 'latin1'), payload])
  return Buffer.concat([pdf.subarray(0, eofIndex), payload, pdf.subarray(eofIndex)])
}

function extractEditableReportData(pdf) {
  const startMarker = Buffer.from(editablePdfStartMarker, 'latin1')
  const endMarker = Buffer.from(editablePdfEndMarker, 'latin1')
  const start = pdf.lastIndexOf(startMarker)
  if (start === -1) return null
  const end = pdf.indexOf(endMarker, start + startMarker.length)
  if (end === -1) throw new Error('The PDF contains incomplete editable report data.')

  const encoded = pdf
    .subarray(start + startMarker.length, end)
    .toString('latin1')
    .split(/\r?\n/)
    .map((line) => line.trim().replace(/^%/, ''))
    .join('')

  const payload = JSON.parse(zlib.gunzipSync(Buffer.from(encoded, 'base64')).toString('utf8'))
  if (payload.schema !== editablePdfSchema || !payload.report) {
    throw new Error('The PDF editable report data is not compatible with this app.')
  }
  return payload.report
}

function safeImportedFileName(filePath) {
  const parsed = path.parse(filePath)
  const safeStem = parsed.name.replace(/[^a-z0-9-_]+/gi, '-').slice(0, 60) || 'imported-file'
  return `${Date.now()}-${crypto.randomUUID()}-${safeStem}${parsed.ext.toLowerCase()}`
}

function copyImportedFile(filePath) {
  const importDir = path.join(app.getPath('userData'), 'imports')
  ensureDir(importDir)
  const fileName = safeImportedFileName(filePath)
  const destination = path.join(importDir, fileName)
  fs.copyFileSync(filePath, destination)
  return destination
}

function multiplyMatrix(m1, m2) {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ]
}

function joinVisualLineItems(items) {
  const sorted = [...items].sort((a, b) => a.x - b.x)
  let line = ''
  let previousEnd = null

  sorted.forEach((item) => {
    const value = String(item.str || '').trim()
    if (!value) return
    const gap = previousEnd === null ? 0 : item.x - previousEnd
    const needsSpace =
      line &&
      gap > 1.8 &&
      value !== '-' &&
      !line.endsWith('-') &&
      !/^[,.;:)]/.test(value)
    line += `${needsSpace ? ' ' : ''}${value}`
    previousEnd = item.x + (item.width || value.length * 5)
  })

  return line.replace(/\s*-\s*/g, '-').trim()
}

async function getPdfVisualLines(page) {
  const textContent = await page.getTextContent()
  const rows = []

  textContent.items.forEach((item) => {
    const value = String(item.str || '').trim()
    if (!value) return
    const y = Math.round(item.transform[5] * 2) / 2
    let row = rows.find((candidate) => Math.abs(candidate.y - y) <= 2)
    if (!row) {
      row = { y, items: [] }
      rows.push(row)
    }
    row.items.push({
      str: value,
      x: item.transform[4],
      width: item.width || 0,
    })
  })

  return rows
    .map((row) => ({ y: row.y, text: joinVisualLineItems(row.items) }))
    .filter((row) => row.text)
    .sort((a, b) => b.y - a.y)
}

function pngBufferFromPdfImage(image) {
  const width = image.width
  const height = image.height
  const data = image.data
  if (!width || !height || !data) return null

  const pixelCount = width * height
  const channels = data.length >= pixelCount * 4 ? 4 : data.length >= pixelCount * 3 ? 3 : data.length >= pixelCount ? 1 : 0
  if (!channels) return null

  const maxSide = 1200
  const scale = Math.min(1, maxSide / Math.max(width, height))
  const outputWidth = Math.max(1, Math.round(width * scale))
  const outputHeight = Math.max(1, Math.round(height * scale))
  const png = new PNG({ width: outputWidth, height: outputHeight })

  for (let y = 0; y < outputHeight; y += 1) {
    const sourceY = Math.min(height - 1, Math.floor(y / scale))
    for (let x = 0; x < outputWidth; x += 1) {
      const sourceX = Math.min(width - 1, Math.floor(x / scale))
      const source = (sourceY * width + sourceX) * channels
      const target = (y * outputWidth + x) * 4
      if (channels === 4) {
        png.data[target] = data[source]
        png.data[target + 1] = data[source + 1]
        png.data[target + 2] = data[source + 2]
        png.data[target + 3] = data[source + 3]
      } else if (channels === 3) {
        png.data[target] = data[source]
        png.data[target + 1] = data[source + 1]
        png.data[target + 2] = data[source + 2]
        png.data[target + 3] = 255
      } else {
        png.data[target] = data[source]
        png.data[target + 1] = data[source]
        png.data[target + 2] = data[source]
        png.data[target + 3] = 255
      }
    }
  }

  return PNG.sync.write(png)
}

async function extractPdfImagesFromPage(page) {
  const opList = await page.getOperatorList()
  const candidates = []
  const { OPS } = pdfjsLib
  const stack = []
  let ctm = [1, 0, 0, 1, 0, 0]

  for (let index = 0; index < opList.fnArray.length; index += 1) {
    const fn = opList.fnArray[index]
    const args = opList.argsArray[index]

    if (fn === OPS.save) {
      stack.push([...ctm])
    } else if (fn === OPS.restore) {
      ctm = stack.pop() || [1, 0, 0, 1, 0, 0]
    } else if (fn === OPS.transform) {
      ctm = multiplyMatrix(ctm, args)
    } else if (fn === OPS.paintImageXObject || fn === OPS.paintJpegXObject || fn === OPS.paintInlineImageXObject) {
      const name = args[0]
      const displayWidth = Math.abs(ctm[0])
      const displayHeight = Math.abs(ctm[3])
      if (displayWidth < 90 || displayHeight < 70) continue

      const image = await new Promise((resolve) => page.objs.get(name, resolve))
      if (!image || image.width < 120 || image.height < 100) continue

      candidates.push({
        name,
        x: ctm[4],
        y: ctm[5],
        displayWidth,
        displayHeight,
        image,
      })
    }
  }

  return candidates
}

async function extractPdfVisualData(filePath, buffer) {
  const doc = await pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    disableWorker: true,
    verbosity: 0,
  }).promise
  const photoDir = path.join(app.getPath('userData'), 'photos')
  ensureDir(photoDir)

  const visualPages = []
  const exhibitPhotos = []
  const assigned = new Set()
  let carryExhibitNumber = null
  const safeStem = path.parse(filePath).name.replace(/[^a-z0-9-_]+/gi, '-').slice(0, 50) || 'imported-pdf'

  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
    const page = await doc.getPage(pageNumber)
    const lines = await getPdfVisualLines(page)
    const pageText = lines.map((line) => line.text).join('\n')
    visualPages.push(pageText)

    const headers = lines
      .map((line) => {
        const match = line.text.match(/^Exhibit\s+([1-3])\b/i)
        return match ? { number: Number(match[1]), y: line.y } : null
      })
      .filter(Boolean)
      .sort((a, b) => b.y - a.y)

    const isPhotoEvidencePage = headers.length > 0 && /Photographic Evidence|EXHIBIT A:|Click to upload/i.test(pageText)
    if (assigned.size < 6 && isPhotoEvidencePage) {
      const images = await extractPdfImagesFromPage(page)
      images
        .sort((a, b) => b.y - a.y || a.x - b.x)
        .forEach((candidate) => {
          const headerAbove = headers
            .filter((header) => header.y > candidate.y)
            .sort((a, b) => (a.y - candidate.y) - (b.y - candidate.y))[0]
          const exhibitNumber = headerAbove?.number || carryExhibitNumber
          if (!exhibitNumber || exhibitNumber < 1 || exhibitNumber > 3) return

          const side = candidate.x < page.view[2] / 2 ? 'pre' : 'post'
          const key = `${exhibitNumber}-${side}`
          if (assigned.has(key)) return

          const png = pngBufferFromPdfImage(candidate.image)
          if (!png) return

          const fileName = `${Date.now()}-${crypto.randomUUID()}-${safeStem}-exhibit-${exhibitNumber}-${side}.png`
          const photoPath = path.join(photoDir, fileName)
          fs.writeFileSync(photoPath, png)
          assigned.add(key)
          exhibitPhotos.push({
            exhibitIndex: exhibitNumber - 1,
            side,
            photo: {
              name: `Exhibit ${exhibitNumber} ${side === 'pre' ? 'pre' : 'corrective'} photo`,
              path: photoPath,
              url: pathToFileURL(photoPath).href,
            },
          })
        })
    }

    if (headers.length) {
      carryExhibitNumber = headers[headers.length - 1].number
    }

    await new Promise((resolve) => setImmediate(resolve))
  }

  return {
    text: visualPages.join('\n\n').trim(),
    exhibitPhotos,
  }
}

function decodeXmlEntities(text) {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
}

function plainTextFromRtf(text) {
  return text
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\'[0-9a-fA-F]{2}/g, ' ')
    .replace(/[{}]/g, '')
    .replace(/\\[a-z]+\d* ?/gi, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function textLooksReadable(text) {
  if (!text.trim()) return false
  const sample = text.slice(0, 2000)
  const printable = sample.replace(/[\x09\x0a\x0d\x20-\x7e]/g, '')
  return printable.length / sample.length < 0.15
}

async function extractImportedDocumentData(filePath, buffer) {
  const ext = path.extname(filePath).toLowerCase()

  if (ext === '.pdf') {
    try {
      return await extractPdfVisualData(filePath, buffer)
    } catch (error) {
      log('Visual PDF import failed; falling back to basic text extraction', error)
      const parsed = await pdfParse(buffer)
      return { text: parsed.text.trim(), exhibitPhotos: [] }
    }
  }

  if (ext === '.docx') {
    const result = await mammoth.extractRawText({ buffer })
    return { text: result.value.trim(), exhibitPhotos: [] }
  }

  if (ext === '.xlsx' || ext === '.xls' || ext === '.xlsm') {
    const workbook = XLSX.read(buffer, { type: 'buffer' })
    const text = workbook.SheetNames.map((sheetName) => {
      const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], { blankrows: false })
      return [`Sheet: ${sheetName}`, csv].filter(Boolean).join('\n')
    }).join('\n\n').trim()
    return { text, exhibitPhotos: [] }
  }

  if (ext === '.csv' || ext === '.txt' || ext === '.md' || ext === '.json') {
    return { text: buffer.toString('utf8').trim(), exhibitPhotos: [] }
  }

  if (ext === '.rtf') {
    return { text: plainTextFromRtf(buffer.toString('utf8')), exhibitPhotos: [] }
  }

  if (ext === '.doc') {
    throw new Error('Legacy .doc files can be attached, but text extraction is not supported. Please save the document as .docx for text extraction.')
  }

  const text = buffer.toString('utf8')
  return { text: textLooksReadable(text) ? text.trim() : '', exhibitPhotos: [] }
}

async function buildImportedSourceDocument(filePath) {
  const buffer = fs.readFileSync(filePath)
  const storedPath = copyImportedFile(filePath)
  let extractedText = ''
  let exhibitPhotos = []
  let extractionStatus = 'No readable text was extracted from this file.'

  try {
    const importedData = await extractImportedDocumentData(filePath, buffer)
    extractedText = importedData.text
    extractionStatus = extractedText
      ? `Extracted ${extractedText.length.toLocaleString()} characters of text.`
      : extractionStatus
    if (importedData.exhibitPhotos?.length) {
      extractionStatus += ` Imported ${importedData.exhibitPhotos.length} exhibit photo${importedData.exhibitPhotos.length === 1 ? '' : 's'}.`
    }
    exhibitPhotos = importedData.exhibitPhotos || []
  } catch (error) {
    extractionStatus = error.message || String(error)
    log(`Imported file text extraction failed for ${filePath}`, error)
  }

  return {
    name: path.basename(filePath),
    path: storedPath,
    url: pathToFileURL(storedPath).href,
    type: mimeForFile(filePath),
    extension: path.extname(filePath).toLowerCase(),
    importedAt: new Date().toISOString(),
    text: extractedText,
    extractionStatus,
    exhibitPhotos: exhibitPhotos || [],
  }
}

function buildImportedSourceDocumentInWorker(filePath) {
  return new Promise((resolve, reject) => {
    const workerPath = path.join(__dirname, 'import-worker.cjs')
    const worker = fork(workerPath, [], {
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      stdio: ['ignore', 'ignore', 'ignore', 'ipc'],
    })

    let settled = false
    const settle = (fn, value) => {
      if (settled) return
      settled = true
      clearTimeout(timeout)
      fn(value)
    }

    const timeout = setTimeout(() => {
      worker.kill()
      settle(reject, new Error('Timed out while importing the selected file.'))
    }, 120000)

    worker.once('message', (message) => {
      if (message?.ok) {
        settle(resolve, message.sourceDocument)
      } else {
        settle(reject, new Error(message?.error || 'File import worker failed.'))
      }
      worker.disconnect()
    })

    worker.once('error', (error) => {
      settle(reject, error)
    })

    worker.once('exit', (code) => {
      settle(reject, new Error(code ? `File import worker exited with code ${code}.` : 'File import worker exited unexpectedly.'))
    })

    worker.send({ filePath, userData: app.getPath('userData') })
  })
}

async function writeDiagnosticsZip(filePath) {
  log('Exporting diagnostics')
  const userData = app.getPath('userData')
  const photosDir = path.join(userData, 'photos')
  const zip = new AdmZip()
  const reports = db
    ? db.prepare('SELECT id, title, updated_at AS updatedAt FROM reports ORDER BY updated_at DESC').all()
    : []
  const photoFiles = walkFiles(photosDir).map((file) => ({
    name: file.name,
    relativePath: path.relative(photosDir, file.path),
    size: file.size,
    modifiedAt: file.modifiedAt,
  }))

  zip.addFile('diagnostics/system-info.json', Buffer.from(JSON.stringify({
    generatedAt: new Date().toISOString(),
    app: {
      name: app.getName(),
      version: app.getVersion(),
      packaged: app.isPackaged,
      userData,
    },
    runtime: {
      electron: process.versions.electron,
      chrome: process.versions.chrome,
      node: process.versions.node,
      v8: process.versions.v8,
    },
    system: {
      platform: process.platform,
      arch: process.arch,
      osType: os.type(),
      osRelease: os.release(),
      osVersion: os.version?.(),
      totalMemory: os.totalmem(),
      freeMemory: os.freemem(),
      cpus: os.cpus().map((cpu) => cpu.model),
    },
    reports,
    photos: photoFiles,
  }, null, 2)))

  addFileIfExists(zip, 'logs/ergonomic-report-startup.log', getStartupLogPath())

  if (db) {
    const backupPath = path.join(app.getPath('temp'), `ergonomic-reports-${Date.now()}.sqlite`)
    try {
      await db.backup(backupPath)
      addFileIfExists(zip, 'data/ergonomic-reports.sqlite', backupPath)
    } finally {
      fs.rmSync(backupPath, { force: true })
    }
  }

  zip.addFile('README.txt', Buffer.from([
    'Ergonomic Evaluation Report diagnostics',
    '',
    'Send this ZIP to the developer when reporting app issues.',
    'It contains app/system metadata, startup logs, and a copy of the local SQLite reports database.',
    'Uploaded photo image files are not included; only photo filenames and sizes are listed in diagnostics/system-info.json.',
  ].join('\n')))

  zip.writeZip(filePath)
  log(`Diagnostics exported to ${filePath}`)
  return filePath
}

async function exportDiagnostics(parentWindow) {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
  const { canceled, filePath } = await dialog.showSaveDialog(parentWindow || mainWindow, {
    title: 'Export Diagnostics',
    defaultPath: `Ergonomic-Report-Diagnostics-${stamp}.zip`,
    filters: [{ name: 'ZIP archive', extensions: ['zip'] }],
  })
  if (canceled || !filePath) return null
  const savedPath = await writeDiagnosticsZip(filePath)
  if (parentWindow) {
    await dialog.showMessageBox(parentWindow, {
      type: 'info',
      title: 'Diagnostics Exported',
      message: 'Diagnostics ZIP created successfully.',
      detail: savedPath,
    })
  }
  return savedPath
}

function createApplicationMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        {
          label: 'Export Diagnostics',
          click: () => exportDiagnostics(BrowserWindow.getFocusedWindow() || mainWindow)
            .catch((error) => {
              log('Diagnostics export failed', error)
              dialog.showErrorBox('Diagnostics Export Failed', error.message || String(error))
            }),
        },
        { type: 'separator' },
        { role: 'quit' },
      ],
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
    {
      role: 'help',
      submenu: [
        {
          label: 'Export Diagnostics',
          click: () => exportDiagnostics(BrowserWindow.getFocusedWindow() || mainWindow)
            .catch((error) => {
              log('Diagnostics export failed', error)
              dialog.showErrorBox('Diagnostics Export Failed', error.message || String(error))
            }),
        },
      ],
    },
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

function initDatabase() {
  log('Initializing SQLite')
  const Database = require('better-sqlite3')
  const dataDir = app.getPath('userData')
  ensureDir(dataDir)
  db = new Database(path.join(dataDir, 'ergonomic-reports.sqlite'))
  db.exec(`
    CREATE TABLE IF NOT EXISTS reports (
      id TEXT PRIMARY KEY,
      title TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      data_json TEXT NOT NULL
    );
  `)
  log('SQLite initialized')
}

function createWindow() {
  log('Creating BrowserWindow')
  mainWindow = new BrowserWindow({
    title: 'Ergonomic Evaluation Report',
    width: 1420,
    height: 960,
    minWidth: 1180,
    minHeight: 760,
    backgroundColor: '#f4f6f8',
    icon: getAppIcon(),
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  mainWindow.once('ready-to-show', () => {
    log('Window ready-to-show')
    mainWindow.center()
    mainWindow.show()
    mainWindow.focus()
  })

  if (devServerUrl) {
    log(`Loading dev server ${devServerUrl}`)
    mainWindow.loadURL(devServerUrl)
  } else {
    const filePath = path.join(__dirname, '..', 'dist', 'index.html')
    log(`Loading file ${filePath}`)
    mainWindow.loadFile(filePath).catch((error) => log('Failed to load file', error))
  }

  setTimeout(() => {
    if (mainWindow && !mainWindow.isVisible()) {
      log('Forcing hidden window visible')
      mainWindow.center()
      mainWindow.show()
      mainWindow.focus()
    }
  }, 2000)
}

app.whenReady().then(() => {
  log('App ready')
  app.setAppUserModelId('com.ergonomicworks.report')
  initDatabase()
  createApplicationMenu()
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('reports:list', () => {
  return db.prepare('SELECT id, title, updated_at AS updatedAt FROM reports ORDER BY updated_at DESC').all()
})

ipcMain.handle('reports:load', (_event, id) => {
  const row = db.prepare('SELECT data_json FROM reports WHERE id = ?').get(id)
  return row ? JSON.parse(row.data_json) : null
})

ipcMain.handle('reports:save', (_event, report) => {
  const now = new Date().toISOString()
  const title = report.employeeName || report.caseType || 'Untitled Ergonomic Report'
  const payload = { ...report, updatedAt: now }
  db.prepare(`
    INSERT INTO reports (id, title, updated_at, data_json)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
      title = excluded.title,
      updated_at = excluded.updated_at,
      data_json = excluded.data_json
  `).run(payload.id, title, now, JSON.stringify(payload))
  return payload
})

ipcMain.handle('reports:delete', (_event, id) => {
  db.prepare('DELETE FROM reports WHERE id = ?').run(id)
  return true
})

ipcMain.handle('photos:save', (_event, photo) => {
  const matches = /^data:(.+);base64,(.+)$/.exec(photo.dataUrl)
  if (!matches) throw new Error('Photo must be sent as a data URL.')
  const mime = matches[1]
  const ext = mime.split('/')[1]?.replace('jpeg', 'jpg') || 'jpg'
  const photoDir = path.join(app.getPath('userData'), 'photos')
  ensureDir(photoDir)
  const safeStem = path.parse(photo.name || 'photo').name.replace(/[^a-z0-9-_]+/gi, '-').slice(0, 50)
  const fileName = `${Date.now()}-${crypto.randomUUID()}-${safeStem}.${ext}`
  const filePath = path.join(photoDir, fileName)
  fs.writeFileSync(filePath, Buffer.from(matches[2], 'base64'))
  return {
    name: photo.name,
    path: filePath,
    url: pathToFileURL(filePath).href,
  }
})

ipcMain.handle('pdf:import', async () => {
  const focused = BrowserWindow.getFocusedWindow() || mainWindow
  const { canceled, filePaths } = await dialog.showOpenDialog(focused, {
    title: 'Import Report File',
    properties: ['openFile'],
    filters: [
      { name: 'Supported report files', extensions: ['pdf', 'docx', 'doc', 'xlsx', 'xls', 'xlsm', 'csv', 'txt', 'rtf', 'md', 'json', 'jpg', 'jpeg', 'png', 'webp'] },
      { name: 'PDF', extensions: ['pdf'] },
      { name: 'Word documents', extensions: ['docx', 'doc'] },
      { name: 'Excel workbooks', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
      { name: 'Text files', extensions: ['txt', 'rtf', 'md', 'json'] },
      { name: 'Images', extensions: ['jpg', 'jpeg', 'png', 'webp'] },
      { name: 'All files', extensions: ['*'] },
    ],
  })
  if (canceled || !filePaths?.[0]) {
    log('PDF import canceled')
    return null
  }

  const filePath = filePaths[0]
  log(`File import selected: ${filePath}`)
  try {
    const buffer = await fs.promises.readFile(filePath)
    const isPdf = path.extname(filePath).toLowerCase() === '.pdf'
    const report = isPdf ? extractEditableReportData(buffer) : null
    if (report) {
      log(`PDF import succeeded with editable report data: ${filePath}`)
      await dialog.showMessageBox(focused, {
        type: 'info',
        title: 'Editable Report Imported',
        message: 'Editable report imported successfully.',
        detail: filePath,
      })
      return { filePath, report }
    }

    const sourceDocument = await buildImportedSourceDocumentInWorker(filePath)
    log(`Imported source file: ${filePath} (${sourceDocument.extractionStatus})`)
    if (!sourceDocument.text) {
      await dialog.showMessageBox(focused, {
        type: 'warning',
        title: 'File Imported',
        message: 'The file was attached to the report, but no readable text was extracted.',
        detail: `${sourceDocument.extractionStatus}\n\nFor full editable restoration, import a PDF that was exported from this app. For Word and Excel text extraction, use .docx, .xlsx, .csv, or .txt when possible.`,
      })
      return {
        filePath,
        sourceDocument,
        warning: sourceDocument.extractionStatus,
      }
    }

    await dialog.showMessageBox(focused, {
      type: 'info',
      title: 'File Imported',
      message: 'The file was imported and readable text was extracted.',
      detail: `${sourceDocument.name}\n${sourceDocument.extractionStatus}`,
    })
    return { filePath, sourceDocument }
  } catch (error) {
    log('File import failed', error)
    await dialog.showMessageBox(focused, {
      type: 'error',
      title: 'File Import Failed',
      message: 'The selected file could not be imported.',
      detail: error.message || String(error),
    })
    return {
      filePath,
      error: error.message || String(error),
    }
  }
})

ipcMain.handle('pdf:export', async (_event, title, report) => {
  const focused = BrowserWindow.getFocusedWindow() || mainWindow

  try {
    const safeTitle = (title || report?.employeeName || 'Ergonomic Report')
      .replace(/[\\/:*?"<>|]+/g, '-')
      .trim()

    const { canceled, filePath } = await dialog.showSaveDialog(focused, {
      title: 'Export Ergonomic Report PDF',
      defaultPath: `${safeTitle}.pdf`,
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    })

    if (canceled || !filePath) {
      log('PDF export canceled')
      return null
    }

    const finalPath = filePath.toLowerCase().endsWith('.pdf')
      ? filePath
      : `${filePath}.pdf`

    log(`Starting PDF export to: ${finalPath}`)

    const printReport = optimizeReportForPdf(report)

    const pdf = await renderReportPdfBuffer(printReport)

    const editablePdf = embedEditableReportData(pdf, printReport)

    fs.writeFileSync(finalPath, editablePdf)

    if (!fs.existsSync(finalPath)) {
      throw new Error(`PDF was not created at: ${finalPath}`)
    }

    const stats = fs.statSync(finalPath)

    if (!stats.size || stats.size < 1000) {
      throw new Error(`PDF was created but appears invalid or empty: ${finalPath}`)
    }

    log(`PDF exported successfully: ${finalPath} (${stats.size} bytes)`)

    await dialog.showMessageBox(focused, {
      type: 'info',
      title: 'PDF Exported',
      message: 'PDF created successfully.',
      detail: finalPath,
    })

    return finalPath
  } catch (error) {
    log('PDF export failed', error)

    await dialog.showMessageBox(focused, {
      type: 'error',
      title: 'PDF Export Failed',
      message: 'The PDF could not be created.',
      detail: error.message || String(error),
    })

    return {
      error: error.message || String(error),
    }
  }
})

ipcMain.handle('diagnostics:export', async () => {
  return exportDiagnostics(BrowserWindow.getFocusedWindow() || mainWindow)
})
