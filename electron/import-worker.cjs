const fs = require('node:fs')
const path = require('node:path')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const pdfParse = require('pdf-parse')
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js')
const mammoth = require('mammoth')
const XLSX = require('xlsx')
const { PNG } = require('pngjs')

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true })
}

function mimeForFile(filePath) {
  const ext = path.extname(filePath || '').toLowerCase()
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg'
  if (ext === '.png') return 'image/png'
  if (ext === '.webp') return 'image/webp'
  if (ext === '.gif') return 'image/gif'
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

function safeImportedFileName(filePath) {
  const parsed = path.parse(filePath)
  const safeStem = parsed.name.replace(/[^a-z0-9-_]+/gi, '-').slice(0, 60) || 'imported-file'
  return `${Date.now()}-${crypto.randomUUID()}-${safeStem}${parsed.ext.toLowerCase()}`
}

function copyImportedFile(filePath, userData) {
  const importDir = path.join(userData, 'imports')
  ensureDir(importDir)
  const destination = path.join(importDir, safeImportedFileName(filePath))
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
    const needsSpace = line && gap > 1.8 && value !== '-' && !line.endsWith('-') && !/^[,.;:)]/.test(value)
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
    row.items.push({ str: value, x: item.transform[4], width: item.width || 0 })
  })

  return rows
    .map((row) => ({ y: row.y, text: joinVisualLineItems(row.items) }))
    .filter((row) => row.text)
    .sort((a, b) => b.y - a.y)
}

function rotatePng90CW(src, srcW, srcH) {
  const dst = new PNG({ width: srcH, height: srcW })
  for (let outRow = 0; outRow < srcW; outRow += 1) {
    for (let outCol = 0; outCol < srcH; outCol += 1) {
      const srcIdx = ((srcH - 1 - outCol) * srcW + outRow) * 4
      const dstIdx = (outRow * srcH + outCol) * 4
      dst.data[dstIdx] = src.data[srcIdx]
      dst.data[dstIdx + 1] = src.data[srcIdx + 1]
      dst.data[dstIdx + 2] = src.data[srcIdx + 2]
      dst.data[dstIdx + 3] = src.data[srcIdx + 3]
    }
  }
  return PNG.sync.write(dst)
}

function rotatePng90CCW(src, srcW, srcH) {
  const dst = new PNG({ width: srcH, height: srcW })
  for (let outRow = 0; outRow < srcW; outRow += 1) {
    for (let outCol = 0; outCol < srcH; outCol += 1) {
      const srcIdx = (outCol * srcW + (srcW - 1 - outRow)) * 4
      const dstIdx = (outRow * srcH + outCol) * 4
      dst.data[dstIdx] = src.data[srcIdx]
      dst.data[dstIdx + 1] = src.data[srcIdx + 1]
      dst.data[dstIdx + 2] = src.data[srcIdx + 2]
      dst.data[dstIdx + 3] = src.data[srcIdx + 3]
    }
  }
  return PNG.sync.write(dst)
}

function pngBufferFromPdfImage(image, rotationDeg = 0) {
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

  const snap = ((Math.round(rotationDeg / 90) * 90) % 360 + 360) % 360
  if (snap === 270) return rotatePng90CW(png, outputWidth, outputHeight)
  if (snap === 90) return rotatePng90CCW(png, outputWidth, outputHeight)
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
      // Use Euclidean magnitude so rotated images (ctm[0]=0, ctm[3]=0) are not missed.
      const displayWidth = Math.sqrt(ctm[0] * ctm[0] + ctm[1] * ctm[1])
      const displayHeight = Math.sqrt(ctm[2] * ctm[2] + ctm[3] * ctm[3])
      if (displayWidth < 90 || displayHeight < 70) continue

      const image = await new Promise((resolve) => {
        const timer = setTimeout(() => resolve(null), 6000)
        page.objs.get(name, (obj) => { clearTimeout(timer); resolve(obj) })
      })
      if (!image || image.width < 120 || image.height < 100) continue
      const rotationDeg = Math.round(Math.atan2(ctm[1], ctm[0]) * 180 / Math.PI / 90) * 90
      candidates.push({ name, x: ctm[4], y: ctm[5], displayWidth, displayHeight, image, rotationDeg })
    }
  }

  return candidates
}

async function extractPdfVisualData(filePath, buffer, userData) {
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer), disableWorker: true, verbosity: 0 }).promise
  const photoDir = path.join(userData, 'photos')
  ensureDir(photoDir)

  const visualPages = []
  const exhibitPhotos = []
  const assigned = new Set()
  let carryExhibitNumber = null
  let signaturePhoto = null
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

          const png = pngBufferFromPdfImage(candidate.image, candidate.rotationDeg)
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

    const isCertificationPage = !signaturePhoto && /Secure Electronic Signature Block/i.test(pageText)
    if (isCertificationPage) {
      const images = await extractPdfImagesFromPage(page)
      const sigCandidate = images
        .filter((c) => c.image.width > 200 && c.image.height > 50)
        .sort((a, b) => b.image.width * b.image.height - a.image.width * a.image.height)[0]
      if (sigCandidate) {
        const png = pngBufferFromPdfImage(sigCandidate.image, sigCandidate.rotationDeg)
        if (png) {
          const fileName = `${Date.now()}-${crypto.randomUUID()}-${safeStem}-signature.png`
          const photoPath = path.join(photoDir, fileName)
          fs.writeFileSync(photoPath, png)
          signaturePhoto = {
            name: 'Electronic signature',
            path: photoPath,
            url: `file://${photoPath.replace(/\\/g, '/')}`,
          }
        }
      }
    }

    if (headers.length) carryExhibitNumber = headers[headers.length - 1].number
    await new Promise((resolve) => setImmediate(resolve))
  }

  let pdfKeywords = ''
  try {
    const parsed = await pdfParse(buffer)
    pdfKeywords = parsed.info?.Keywords || ''
  } catch {}
  const visualText = visualPages.join('\n\n').trim()
  return { text: pdfKeywords ? `${pdfKeywords}\n\n${visualText}` : visualText, exhibitPhotos, signaturePhoto }
}

const WIN1252 = {
  0x80: '€', 0x82: '‚', 0x83: 'ƒ', 0x84: '„', 0x85: '…',
  0x86: '†', 0x87: '‡', 0x88: 'ˆ', 0x89: '‰', 0x8A: 'Š',
  0x8B: '‹', 0x8C: 'Œ', 0x8E: 'Ž', 0x91: '‘', 0x92: '’',
  0x93: '“', 0x94: '”', 0x95: '•', 0x96: '–', 0x97: '—',
  0x98: '˜', 0x99: '™', 0x9A: 'š', 0x9B: '›', 0x9C: 'œ',
  0x9E: 'ž', 0x9F: 'Ÿ',
}

function plainTextFromRtf(text) {
  return text
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\'([0-9a-fA-F]{2})/g, (_, hex) => {
      const code = parseInt(hex, 16)
      return WIN1252[code] ?? Buffer.from([code]).toString('latin1')
    })
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

async function extractImportedDocumentData(filePath, buffer, userData) {
  const ext = path.extname(filePath).toLowerCase()

  if (ext === '.pdf') {
    try {
      return await extractPdfVisualData(filePath, buffer, userData)
    } catch {
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

async function buildImportedSourceDocument(filePath, userData) {
  const buffer = fs.readFileSync(filePath)
  const storedPath = copyImportedFile(filePath, userData)
  let extractedText = ''
  let exhibitPhotos = []
  let signaturePhoto = null
  let extractionStatus = 'No readable text was extracted from this file.'

  const importedData = await extractImportedDocumentData(filePath, buffer, userData)
  extractedText = importedData.text
  extractionStatus = extractedText
    ? `Extracted ${extractedText.length.toLocaleString()} characters of text.`
    : extractionStatus
  if (importedData.exhibitPhotos?.length) {
    extractionStatus += ` Imported ${importedData.exhibitPhotos.length} exhibit photo${importedData.exhibitPhotos.length === 1 ? '' : 's'}.`
  }
  exhibitPhotos = importedData.exhibitPhotos || []
  if (importedData.signaturePhoto) {
    signaturePhoto = importedData.signaturePhoto
    extractionStatus += ' Imported electronic signature.'
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
    exhibitPhotos,
    signaturePhoto,
  }
}

process.once('message', async (message) => {
  try {
    const sourceDocument = await buildImportedSourceDocument(message.filePath, message.userData)
    process.send?.({ ok: true, sourceDocument })
  } catch (error) {
    process.send?.({ ok: false, error: error.message || String(error) })
  }
})
