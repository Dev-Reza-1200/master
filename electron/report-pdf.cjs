const fs = require('node:fs')
const path = require('node:path')
const { fileURLToPath } = require('node:url')
const PDFDocument = require('pdfkit')

const navy = '#193c69'
const gold = '#d9a129'
const ink = '#17345b'
const muted = '#607087'
const line = '#a8b6c6'
const ice = '#e9f7fc'
const paper = '#fbfdff'
const danger = '#ffd3da'
const medium = '#fff5bc'
const low = '#c5f5d2'

const margin = { top: 30, right: 40, bottom: 38, left: 40 }
const contentWidth = 612 - margin.left - margin.right
const pageHeaderY = 30
const sectionBarY = 78
const pageTopY = 124
const footerSafeZone = 70

const dataSources = [
  'On-Site Observation',
  'Job Function Review',
  'Medical Records',
  'Employee Interview',
]

const assessmentTools = ['RULA', 'REBA', 'ROSA', 'NIOSH']
const benchmarks = ['ANSI/HFES 100-2007', 'Cal/OSHA Title 8 §5110']
const controls = ['Micro-Breaks', 'Software', 'Self-Adjustment']
const softwareOptions = ['RSI Guard', 'Auto-Click', 'Keyboard Shortcuts']

const justifications = [
  'Neutral Wrist Alignment',
  'Pronation Mitigation',
  'Ulnar Deviation Correction',
  'Neutral Shoulder Adduction',
  'Elbow Flexion Management',
  'Cervical Neutrality and Gaze Dynamics',
  'Lumbar Integrity and Lordotic Support',
  'Thoracic/Scapular Support',
  'Ocular Accommodation & Glare',
  'Midline Alignment (The Neutral Pathway)',
  'Plantar Surface Stability',
  'Pregnancy Postural Support',
  'Pregnancy Circulatory Enhancement',
  'Metabolic Flux (Sit-Stand Transitions)',
  'Contact Stress Mitigation',
]

const clinicalHeadingNames = [
  ...justifications,
  'Metabolic Flux & Height Accommodation',
  'Cervical Neutrality & Gaze Dynamics',
  'Neutral Reach Zone Optimization',
  'Dynamic Neutral Reach Zone Optimization',
  'Display Neutral Reach Zone Optimization',
  'Engineering Control & Regulatory Compliance',
  'Biomechanical Requirement for Negative Tilt & Neutral Reach Zone Engineering Controls & Contact Stress',
  'Biochemical Requirement for Negative Tilt & Neutral Reach Zone Engineering Controls & Contact Stress',
  'Mitigation of Wrist Extension',
]

const postureRows = [
  { key: 'neck', label: 'Neck / Head', factors: ['Extension', 'Rotation', 'Forward Lean'] },
  { key: 'shoulders', label: 'Shoulders', factors: ['Elevation', 'Protraction', 'Abduction'] },
  { key: 'wrists', label: 'Wrists / Arms', factors: ['Extension', 'Ulnar Deviation', 'Contact Stress'] },
  { key: 'back', label: 'Back / Trunk', factors: ['Slumping', 'Lack of Lumbar Support'] },
]

function sanitizePdfText(text) {
  return String(text || '')
    .replace(/\r/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[\uFFFE\uFFFF]/g, ' ')
    .replace(/[\u2018\u2019\u201b]/g, "'")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/[\u2010\u2011\u2012\u2013\u2014]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/[^\x20-\x7E\n§°]/g, ' ')
}

function cleanText(text) {
  return sanitizePdfText(text)
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\b(\d+(?:st|nd|rd|th))[-\s]+(percentile)\b/gi, '$1 $2')
    .replace(/\bleg[-\s]+crossing\b/gi, 'leg crossing')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function reflowText(text) {
  return cleanText(text)
    .split(/\n{2,}/)
    .map((block) => {
      if (/^-\s+/m.test(block)) return block
      return block.replace(/\n+/g, ' ').replace(/\s{2,}/g, ' ').trim()
    })
    .filter(Boolean)
    .join('\n\n')
}

function cleanInlineText(text) {
  return cleanText(text)
    .replace(/\s*\n+\s*/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim()
}

function cleanEquipmentModel(text) {
  return cleanInlineText(text)
    .replace(/([A-Za-z0-9")])\s*-\s*(?=\$)/g, '$1 - ')
    .replace(/([A-Za-z0-9")])\s*-\s+(?=[A-Za-z0-9])/g, '$1 - ')
    .replace(/\s{2,}/g, ' ')
    .trim()
}

function escapeRegExp(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function formatJustificationDisplay(text) {
  let formatted = cleanText(text)
    .replace(/\bEditable Clinical Justification Display\b\s*/gi, '')
    .replace(/\b(Metabolic Flux\s*&\s*Height Accommodation):\s*/gi, '\n\n$1: ')
    .replace(/\b(Cervical Neutrality\s*&\s*Gaze Dynamics):\s*/gi, '\n\n$1: ')
    .replace(/\b(Biomechanical Requirement for Negative Tilt & Neutral Reach Zone)\s+(Engineering Controls & Contact Stress):\s*/gi, '\n\n$1 $2: ')
    .replace(/([.!?])\s+(?=[A-Z][A-Za-z][A-Za-z0-9/&()'" -]{8,150}:)/g, '$1\n\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
  clinicalHeadingNames
    .slice()
    .sort((a, b) => b.length - a.length)
    .forEach((heading) => {
      const pattern = new RegExp(`(^|\\s)(${escapeRegExp(heading)}):\\s*`, 'gi')
      formatted = formatted.replace(pattern, (match, prefix, title, offset) => {
        const needsBreak = offset > 0 && !/\n\n$/.test(prefix)
        return `${needsBreak ? '\n\n' : ''}${title}: `
      })
    })
  return formatted.replace(/\n{3,}/g, '\n\n').trim()
}

function formatEvaluatorSummary(text) {
  return cleanText(text)
    .replace(/(?:^|\s)-\s*(?=[A-Z])/g, '\n- ')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function formatProcurementSummary(text) {
  return cleanText(text)
    .replace(/\s+-\s+(?=[A-Z0-9])/g, '\n- ')
    .trim()
}

function formatObservationText(text) {
  return segmentLongParagraphs(cleanText(text))
}

function formatAnalysisText(text) {
  return segmentLongParagraphs(reflowText(text))
}

function segmentLongParagraphs(text) {
  const blocks = cleanText(text)
    .split(/\n{2,}/)
    .map((block) => block.replace(/\n+/g, ' ').trim())
    .filter(Boolean)
  if (!blocks.length) return ''
  return blocks.flatMap((block) => {
    if (block.length <= 720) return [block]
    const sentences = block.match(/[^.!?]+[.!?]+(?:\s+|$)|[^.!?]+$/g)?.map((sentence) => sentence.trim()).filter(Boolean) || [block]
    const grouped = []
    let current = ''
    sentences.forEach((sentence) => {
      const next = current ? `${current} ${sentence}` : sentence
      if (current && next.length > 900) {
        grouped.push(current)
        current = sentence
      } else {
        current = next
      }
    })
    if (current) grouped.push(current)
    return grouped
  }).join('\n\n')
}

function visibleRecommendationRows(rows) {
  return (rows || []).filter((row) => [row.current, row.action, row.notes].some((value) => String(value || '').trim()))
}

function extractUrl(text) {
  return cleanText(text)
    .split('\n')
    .map((line) => /h\s*t\s*t\s*p/i.test(line)
      ? line
        .replace(/\s+/g, '')
        .replace(/h\s*t\s*t\s*p\s*s\s*:\s*\/\s*\//i, 'https://')
        .replace(/h\s*t\s*t\s*p\s*:\s*\/\s*\//i, 'http://')
      : line)
    .join('\n')
    .match(/https?:\/\/[^\s]+/i)?.[0] || ''
}

function logoPath() {
  const candidates = [
    path.join(__dirname, '..', 'dist', 'assets', 'extracted-1.jpeg'),
    path.join(__dirname, '..', 'public', 'assets', 'extracted-1.jpeg'),
  ]
  return candidates.find((candidate) => fs.existsSync(candidate)) || ''
}

function bufferFromPhoto(photo) {
  if (!photo) return null
  if (typeof photo.url === 'string' && photo.url.startsWith('data:')) {
    const match = /^data:[^;]+;base64,(.+)$/.exec(photo.url)
    if (!match) return null
    return Buffer.from(match[1], 'base64')
  }
  const candidates = []
  if (photo.path) candidates.push(photo.path)
  if (typeof photo.url === 'string' && photo.url.startsWith('file://')) {
    try {
      candidates.push(fileURLToPath(photo.url))
    } catch {
      // Ignore invalid file URLs and fall through to the placeholder.
    }
  }
  const filePath = candidates.find((candidate) => candidate && fs.existsSync(candidate))
  return filePath ? fs.readFileSync(filePath) : null
}

class ReportPdfRenderer {
  constructor(report, options = {}) {
    this.report = report || {}
    this.tocPageNumbers = options.tocPageNumbers || null
    this.doc = new PDFDocument({
      size: 'LETTER',
      autoFirstPage: false,
      bufferPages: true,
      margins: margin,
      info: {
        Title: `${this.report.employeeName || 'Ergonomic'} Evaluation Report`,
        Author: 'Ergonomic Works',
        Subject: 'Ergonomic Evaluation Report',
      },
    })
    const drawPdfText = this.doc.text.bind(this.doc)
    this.doc.text = (text, ...args) => drawPdfText(sanitizePdfText(text), ...args)
    this.logo = logoPath()
    this.y = margin.top
    this.pageNumber = 0
    this.sectionCode = ''
    this.sectionTitle = ''
    this.sectionStarts = {}
    this.pageCount = 0
  }

  render() {
    this.doc.info.Keywords = `ERGO_IMPORT_V1;mode=${this.report.assessmentMode || ''};src=${(this.report.dataSources || []).join('|')};tools=${(this.report.assessmentTools || []).join('|')};bench=${(this.report.benchmarks || []).join('|')}`
    this.sectionI()
    this.sectionII()
    this.sectionIII()
    this.sectionIV()
    this.sectionV()
    this.sectionVI()
    this.sectionVII()
    this.finalizePageNumbers()
  }

  addPage(sectionCode, sectionTitle, options = {}) {
    this.doc.addPage({ size: 'LETTER', margins: margin })
    this.pageNumber += 1
    this.sectionCode = sectionCode
    this.sectionTitle = sectionTitle
    if (!this.sectionStarts[sectionCode]) this.sectionStarts[sectionCode] = this.pageNumber
    this.drawHeader()
    this.drawSectionBar(sectionCode, sectionTitle, options.continued)
    this.y = pageTopY
  }

  bottom() {
    return this.doc.page.height - margin.bottom - footerSafeZone
  }

  ensureSpace(height) {
    if (this.y + height <= this.bottom()) return
    this.addPage(this.sectionCode, this.sectionTitle, { continued: true })
  }

  drawHeader() {
    const x = margin.left
    const width = contentWidth
    const y = pageHeaderY
    if (this.logo) {
      this.doc.image(this.logo, x, y, { fit: [34, 30] })
    }
    this.doc
      .font('Helvetica-Bold')
      .fontSize(15)
      .fillColor(navy)
      .text('ERGONOMIC EVALUATION REPORT', x + 46, y + 4, { width: width - 46, lineBreak: false })
    this.doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(muted)
      .text('Professional Ergonomic Assessment', x + 46, y + 21, { width: width - 46, lineBreak: false })
    this.doc
      .strokeColor(navy)
      .lineWidth(1.5)
      .moveTo(x, y + 37)
      .lineTo(x + width, y + 37)
      .stroke()
  }

  drawSectionBar(code, title, continued = false) {
    const h = 24
    const label = continued ? `${title} - Continued` : title
    const y = sectionBarY
    this.doc.roundedRect(margin.left, y, contentWidth, h, 2).fill(navy)
    this.doc.rect(margin.left, y, 4, h).fill(gold)
    this.doc.font('Helvetica-Bold').fontSize(11).fillColor('#ffffff')
    this.doc.text(code, margin.left + 12, y + 6, { width: 24, align: 'center', height: 12, lineBreak: false })
    this.doc.text(label.toUpperCase(), margin.left + 44, y + 6, { width: contentWidth - 52, height: 12, lineBreak: false })
  }

  drawSubhead(text) {
    this.ensureSpace(20)
    this.doc.font('Helvetica-Bold').fontSize(11).fillColor(navy)
    this.doc.text(text, margin.left, this.y, { width: contentWidth })
    this.doc.strokeColor(line).lineWidth(0.7).moveTo(margin.left, this.y + 15).lineTo(margin.left + contentWidth, this.y + 15).stroke()
    this.y += 22
  }

  continuationSubhead(title, continuation) {
    return continuation <= 1 ? `${title} (Continued)` : `${title} (Continued ${continuation})`
  }

  textHeight(text, width, options = {}) {
    const font = options.bold ? 'Helvetica-Bold' : options.italic ? 'Helvetica-Oblique' : 'Helvetica'
    this.doc.font(font).fontSize(options.size || 9)
    return this.doc.heightOfString(reflowText(text) || '-', {
      width,
      lineGap: options.lineGap || 1.4,
      align: options.align || 'justify',
    })
  }

  paragraphLineHeight(options = {}) {
    const size = options.size || 9
    return size + 2.2 + (options.lineGap || 1.4)
  }

  wrapTextLines(text, width, options = {}) {
    const font = options.bold ? 'Helvetica-Bold' : options.italic ? 'Helvetica-Oblique' : 'Helvetica'
    this.doc.font(font).fontSize(options.size || 9)

    const words = cleanInlineText(text).split(/\s+/).filter(Boolean)
    if (!words.length) return ['-']

    const lines = []
    let current = ''

    words.forEach((word) => {
      const candidate = current ? `${current} ${word}` : word
      if (current && this.doc.widthOfString(candidate) > width) {
        lines.push(current)
        current = word
      } else {
        current = candidate
      }
    })

    if (current) lines.push(current)
    return lines
  }

  justifiedTextHeight(text, width, options = {}) {
    const lines = this.wrapTextLines(text, width, options)
    return lines.length * this.paragraphLineHeight(options)
  }

  drawJustifiedText(text, x, y, width, options = {}) {
    const font = options.bold ? 'Helvetica-Bold' : options.italic ? 'Helvetica-Oblique' : 'Helvetica'
    const size = options.size || 9
    const lineHeight = this.paragraphLineHeight(options)
    const lines = this.wrapTextLines(text, width, options)

    this.doc.font(font).fontSize(size).fillColor(options.color || ink)

    lines.forEach((lineText, lineIndex) => {
      const words = lineText.split(/\s+/).filter(Boolean)
      const isLastLine = lineIndex === lines.length - 1

      if (isLastLine || words.length <= 1 || options.align === 'left') {
        this.doc.text(lineText, x, y, {
          width,
          height: lineHeight,
          lineBreak: false,
          align: 'left',
        })
        y += lineHeight
        return
      }

      const wordsWidth = words.reduce((sum, word) => sum + this.doc.widthOfString(word), 0)
      const gapCount = words.length - 1
      const naturalGap = this.doc.widthOfString(' ')
      const extraGap = Math.max(0, (width - wordsWidth) / gapCount)

      // Safety cap avoids ugly rivers on very short lines.
      // Lines that would need extreme spacing are left-aligned instead.
      if (extraGap > naturalGap * 4.5) {
        this.doc.text(lineText, x, y, {
          width,
          height: lineHeight,
          lineBreak: false,
          align: 'left',
        })
        y += lineHeight
        return
      }

      let cursorX = x
      words.forEach((word, wordIndex) => {
        this.doc.text(word, cursorX, y, {
          lineBreak: false,
          continued: false,
        })
        cursorX += this.doc.widthOfString(word)
        if (wordIndex < words.length - 1) cursorX += extraGap
      })

      y += lineHeight
    })

    return y
  }

  paragraphBlocks(text) {
    const clean = reflowText(text)
    if (!clean) return ['-']
    return clean.split(/\n{2,}/).map((block) => block.trim()).filter(Boolean)
  }

  blockHeight(block, width, options = {}) {
    const blockGap = options.blockGap ?? 6
    const lines = block.split(/\n+/).map((line) => line.trim()).filter(Boolean)
    if (lines.some((line) => /^-\s+/.test(line))) {
      return lines.reduce((sum, line) => {
        const body = line.replace(/^-\s+/, '')
        return sum + this.justifiedTextHeight(body, width - 15, options) + 4
      }, 0)
    }
    const match = block.match(/^([^:\n]{4,120}):\s*([\s\S]*)$/)
    if (match) {
      return this.textHeight(match[1], width, { ...options, bold: true, align: 'left' }) + this.justifiedTextHeight(match[2], width, options) + blockGap
    }
    return this.justifiedTextHeight(block, width, options) + blockGap
  }

  splitBlockToFit(block, width, maxHeight, options = {}) {
    const words = cleanText(block).split(/\s+/).filter(Boolean)
    if (words.length <= 1) return { fit: block, rest: '' }
    let low = 1
    let high = words.length
    let fitCount = 1
    while (low <= high) {
      const mid = Math.floor((low + high) / 2)
      const candidate = words.slice(0, mid).join(' ')
      if (this.blockHeight(candidate, width, options) <= maxHeight) {
        fitCount = mid
        low = mid + 1
      } else {
        high = mid - 1
      }
    }
    return {
      fit: words.slice(0, fitCount).join(' '),
      rest: words.slice(fitCount).join(' '),
    }
  }

  collectTextSegment(blocks, width, maxHeight, options = {}) {
    const segment = []
    let used = 0
    while (blocks.length && !cleanText(blocks[0])) blocks.shift()
    while (blocks.length) {
      const block = blocks[0]
      const h = this.blockHeight(block, width, options)
      if (used + h <= maxHeight || !segment.length && h <= maxHeight + 8) {
        segment.push(block)
        used += h
        blocks.shift()
        continue
      }
      const remaining = maxHeight - used
      if (options.keepBlocksWhole && segment.length) break
      if (remaining > 38) {
        const split = this.splitBlockToFit(block, width, remaining, options)
        if (split.fit) {
          segment.push(split.fit)
          if (cleanText(split.rest)) {
            blocks[0] = split.rest
          } else {
            blocks.shift()
          }
          used += this.blockHeight(split.fit, width, options)
        }
      }
      break
    }
    return { segment, used }
  }

  drawTextCard(title, text, options = {}) {
    if (title) this.drawSubhead(title)
    const fontSize = options.fontSize || 9
    const blocks = this.paragraphBlocks(text)
    const pad = options.padding || 10
    const maxInnerWidth = contentWidth - pad * 2
    const innerWidth = Math.min(maxInnerWidth, options.innerWidth || maxInnerWidth)
    const innerX = margin.left + pad + Math.max(0, (maxInnerWidth - innerWidth) / 2)
    const minHeight = options.minHeight || 0
    const minimumReadableSpace = options.minimumReadableSpace || 72
    let continuation = 0

    while (blocks.length) {
      if (this.bottom() - this.y < minimumReadableSpace) {
        this.addPage(this.sectionCode, this.sectionTitle, { continued: true })
        continuation += 1
        if (title) this.drawSubhead(this.continuationSubhead(title, continuation))
      }

      const cardY = this.y
      const currentMinHeight = continuation ? (options.continuationMinHeight ?? minHeight) : minHeight
      const pageSpace = this.bottom() - cardY

      // Default behavior: use the whole remaining page before creating a continued card.
      // Only apply maxHeight when a caller explicitly opts into limiting the first card.
      const configuredMaxHeight = continuation
        ? options.continuationMaxHeight
        : options.maxHeight

      const maxCardHeight = configuredMaxHeight
        ? Math.min(configuredMaxHeight, pageSpace)
        : pageSpace

      const maxContentHeight = Math.max(32, maxCardHeight - pad * 2)
      const textOptions = {
        size: fontSize,
        lineGap: options.lineGap || 1.4,
        blockGap: options.blockGap,
        keepBlocksWhole: options.keepBlocksWhole,
        align: options.align || 'justify',
      }

      const collected = this.collectTextSegment(blocks, innerWidth, maxContentHeight, textOptions)
      const { segment } = collected
      let used = collected.used

      if (!segment.length && blocks.length) {
        const split = this.splitBlockToFit(blocks[0], innerWidth, Math.max(48, maxContentHeight), textOptions)
        segment.push(split.fit || blocks[0])
        used = this.blockHeight(segment[0], innerWidth, textOptions)
        if (cleanText(split.rest)) {
          blocks[0] = split.rest
        } else {
          blocks.shift()
        }
      }

      // Card height follows the actual text height, capped only by the available page space.
      const naturalCardHeight = used + pad * 2 + 2
      const cardHeight = Math.min(maxCardHeight, Math.max(currentMinHeight, naturalCardHeight))

      this.doc.roundedRect(margin.left, cardY, contentWidth, cardHeight, 5).fillAndStroke(options.fill || '#ffffff', options.border || '#8ea5c1')

      let cursor = cardY + pad
      segment.forEach((block) => {
        cursor = this.drawBlock(block, innerX, cursor, innerWidth, textOptions)
      })
      this.y = cardY + cardHeight + (options.after || 12)

      if (blocks.length) {
        this.addPage(this.sectionCode, this.sectionTitle, { continued: true })
        continuation += 1
        if (title) this.drawSubhead(this.continuationSubhead(title, continuation))
      }
    }
  }

  drawBlock(block, x, y, width, options = {}) {
    const hasBullets = /^-\s+/m.test(block)
    const normalizedBlock = options.preserveLineBreaks || hasBullets
      ? cleanText(block)
      : cleanInlineText(block)
    const lines = normalizedBlock.split(/\n+/).map((line) => line.trim()).filter(Boolean)
    const blockGap = options.blockGap ?? 6

    if (lines.some((line) => /^-\s+/.test(line))) {
      lines.forEach((line) => {
        const text = line.replace(/^-\s+/, '')
        this.doc.circle(x + 4, y + 5.6, 2.7).fill(gold)
        y = this.drawJustifiedText(text, x + 15, y, width - 15, {
          ...options,
          color: ink,
        })
        y += 4
      })
      return y + 2
    }

    const match = normalizedBlock.match(/^([^:\n]{4,120}):\s*([\s\S]*)$/)
    if (match) {
      const heading = cleanInlineText(match[1])
      const body = cleanInlineText(match[2]) || '-'

      this.doc
        .font('Helvetica-Bold')
        .fontSize(options.size || 9)
        .fillColor(navy)
        .text(`${heading}:`, x, y, {
          width,
          lineGap: options.lineGap || 1.4,
          continued: false,
          lineBreak: true,
          align: 'left',
        })

      y += this.textHeight(heading, width, { ...options, bold: true, align: 'left' }) + 1
      y = this.drawJustifiedText(body, x, y, width, {
        ...options,
        color: ink,
      })

      return y + blockGap
    }

    y = this.drawJustifiedText(normalizedBlock, x, y, width, {
      ...options,
      color: ink,
    })

    return y + blockGap
  }

  drawCheck(on, x, y) {
    this.doc.rect(x, y, 8.5, 8.5).strokeColor('#526d8f').lineWidth(0.7).stroke()
    if (!on) return
    this.doc.strokeColor(navy).lineWidth(1.2).moveTo(x + 1.8, y + 4.8).lineTo(x + 3.8, y + 7.5).lineTo(x + 7.4, y + 1.5).stroke()
  }

  drawOptionLine(value, selected, description, x, y, width) {
    this.drawCheck(selected.includes(value), x, y + 1)
    this.doc.font('Helvetica-Bold').fontSize(9).fillColor(navy).text(`${value}:`, x + 15, y, { width: 130, continued: false })
    this.doc.font('Helvetica').fontSize(9).fillColor(ink).text(description, x + 145, y, { width: width - 145, lineGap: 1 })
    return y + Math.max(15, this.textHeight(description, width - 145, { size: 9, lineGap: 1 })) + 2
  }

  drawOptionBlock(title, values, selected, descriptions) {
    const heights = descriptions.map((description) => Math.max(15, this.textHeight(description, contentWidth - 170, { size: 9, lineGap: 1 })) + 2)
    const blockHeight = 30 + heights.reduce((a, b) => a + b, 0)
    this.ensureSpace(blockHeight)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, blockHeight, 5).fillAndStroke(paper, line)
    this.doc.font('Helvetica-Bold').fontSize(10).fillColor(navy).text(title, margin.left + 9, top + 8, { width: contentWidth - 18 })
    let y = top + 25
    values.forEach((value, index) => {
      y = this.drawOptionLine(value, selected || [], descriptions[index], margin.left + 12, y, contentWidth - 24)
    })
    this.y = top + blockHeight + 10
  }

  drawGrid(rows, columns = 2) {
    const colW = contentWidth / columns
    const rowH = 28
    rows.forEach((row, index) => {
      const col = index % columns
      if (col === 0) this.ensureSpace(rowH)
      const rowIndexY = this.y
      const x = margin.left + col * colW
      this.doc.rect(x, rowIndexY, colW, rowH).fillAndStroke(index % 2 === 0 ? ice : '#ffffff', line)
      this.doc.font('Helvetica-Bold').fontSize(8.5).fillColor(navy).text(row[0], x + 6, rowIndexY + 5, { width: 92 })
      this.doc.font('Helvetica').fontSize(8.5).fillColor(ink).text(row[1] || '-', x + 100, rowIndexY + 5, { width: colW - 106, lineGap: 1 })
      if (col === columns - 1 || index === rows.length - 1) this.y += rowH
    })
    this.y += 12
  }

  sectionI() {
    this.addPage('I', 'Administrative & Jurisdictional Data')
    this.doc.font('Helvetica-Bold').fontSize(16).fillColor(navy).text('ERGONOMIC EVALUATION REPORT', margin.left, this.y, { width: contentWidth })
    this.y += 26
    this.drawToc()
    this.drawGrid([
      ['Case Type', this.report.caseType],
      ['Case / Claim #', this.report.claimNumber],
      ['Employee Name', this.report.employeeName],
      ['Employee ID', this.report.employeeId],
      ['Department', this.report.department],
      ['Location / Floor', this.report.location],
      ['Date of Eval', this.report.dateOfEval],
      ['Report Date', this.report.reportDate],
      ['Manager Name', this.report.managerName],
      ['Job Title', this.report.jobTitle],
    ], 2)
  }

  drawToc() {
    const rows = [
      ['I', 'Administrative & Jurisdictional Data'],
      ['II', 'Assessment Methodology & Compliance Standards'],
      ['III', 'Comprehensive Workstation Analysis'],
      ['IV', 'Evidentiary Documentation'],
      ['V', 'Findings & Corrective Action Plan'],
      ['VI', 'Ergonomic Equipment Recommendation'],
      ['VII', 'Opinion & Professional Certification'],
    ]
    const x = margin.left
    const col = [48, contentWidth - 96, 48]
    const rowH = 20
    this.doc.rect(x, this.y, contentWidth, rowH).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(9).fillColor('#ffffff')
      .text('Section', x + 6, this.y + 6, { width: col[0] })
      .text('Description', x + col[0] + 6, this.y + 6, { width: col[1] })
      .text('Pg', x + col[0] + col[1], this.y + 6, { width: col[2], align: 'center' })
    this.y += rowH
    rows.forEach((row, index) => {
      const fill = index % 2 ? '#ffffff' : ice
      const rowY = this.y
      const pgX = x + col[0] + col[1]
      this.doc.rect(x, rowY, contentWidth, rowH).fillAndStroke(fill, '#ffffff')
      const pageNumber = this.tocPageNumbers?.[row[0]] || '-'
      this.doc.font('Helvetica').fontSize(9).fillColor(ink)
        .text(row[0], x + 6, rowY + 6, { width: col[0], height: 10, lineBreak: false })
        .text(row[1], x + col[0] + 6, rowY + 6, { width: col[1], height: 10, lineBreak: false })
        .text(String(pageNumber), pgX, rowY + 6, { width: col[2], height: 10, align: 'center', lineBreak: false })
      this.y += rowH
    })
    this.y += 18
  }

  sectionII() {
    this.addPage('II', 'Assessment Methodology & Compliance Standards')
    this.doc.font('Helvetica-Oblique').fontSize(9).fillColor(ink).text('This evaluation follows recognized industry protocols to ensure objective findings and minimize bias.', margin.left, this.y, { width: contentWidth })
    this.y += 20
    this.drawModeRow()
    this.drawOptionBlock('A. Scope of Evaluation', dataSources, this.report.dataSources, [
      'Direct assessment of the employee at their workstation.',
      'Analysis of Essential Job Functions (EJF).',
      'Review of physician-mandated work restrictions.',
      'Subjective report of discomfort and work habits.',
    ])
    this.drawOptionBlock('B. Evidence-Based Assessment Tools', assessmentTools, this.report.assessmentTools, [
      'Rapid Upper Limb Assessment (Neck, trunk, and upper limbs).',
      'Rapid Entire Body Assessment (Whole-body posture and load).',
      'Rapid Office Strain Assessment (Computer/desk environments).',
      'Lifting Equation (Manual material handling tasks).',
    ])
    this.drawOptionBlock('C. Regulatory & Industry Benchmarks', benchmarks, this.report.benchmarks, [
      'Human Factors Engineering of Computer Workstations.',
      'California Repetitive Motion Injury Standard.',
    ])
  }

  drawModeRow() {
    const h = 26
    this.doc.roundedRect(margin.left, this.y, contentWidth, h, 4).fillAndStroke(ice, '#c3cfdd')
    this.doc.font('Helvetica-Bold').fontSize(9).fillColor(navy).text('Assessment Mode', margin.left + 8, this.y + 8, { width: 110 })
    this.drawCheck(this.report.assessmentMode === 'In-Person', margin.left + 130, this.y + 8)
    this.doc.font('Helvetica').fontSize(9).fillColor(ink).text('In-Person', margin.left + 145, this.y + 7)
    this.drawCheck(this.report.assessmentMode === 'Remote/Virtual', margin.left + 220, this.y + 8)
    this.doc.text('Remote/Virtual', margin.left + 235, this.y + 7)
    this.y += h + 10
  }

  sectionIII() {
    this.addPage('III', 'Comprehensive Workstation Analysis')
    this.doc.font('Helvetica-Oblique').fontSize(9).fillColor(ink).text('Measurements are recorded to identify deviations from neutral Home Row positioning.', margin.left, this.y, { width: contentWidth })
    this.y += 20
    this.drawWorkstationData()
    this.drawPostureMatrix()
    this.drawTextCard('Postural Observations', formatObservationText(this.report.posturalObservations), {
      fontSize: 7.75,
      minHeight: 0,
      lineGap: 3.1,
      blockGap: 10,
      padding: 10,
      innerWidth: contentWidth - 20,
    })
  }

  drawWorkstationData() {
    this.drawSubhead('A. Anthropometric & Equipment Data')
    this.drawGrid([
      ['Stature', this.report.stature],
      ['Dominant Hand', this.report.dominantHand],
      ['Vision', this.report.vision],
      ['Desk Type', this.report.deskType],
      ['Current Desk Height', this.report.currentDeskHeight],
      ['Ideal Keying Height', this.report.idealKeyingHeight],
      ['Delta', this.report.delta],
    ], 4)
  }

  drawPostureMatrix() {
    this.drawSubhead('B. Postural Deviation Matrix')
    const x = margin.left
    const widths = [110, contentWidth - 170, 60]
    const rowH = 38
    this.ensureSpace(28 + rowH * postureRows.length)
    this.doc.rect(x, this.y, contentWidth, 22).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(9).fillColor('#ffffff')
      .text('Body Segment', x + 6, this.y + 7, { width: widths[0] })
      .text('Observed Posture / Risk Factor', x + widths[0] + 6, this.y + 7, { width: widths[1] })
      .text('Risk', x + widths[0] + widths[1], this.y + 7, { width: widths[2], align: 'center' })
    this.y += 22
    postureRows.forEach((row, index) => {
      const posture = this.report.postures?.[row.key] || { factors: [], risk: 'L' }
      this.doc.rect(x, this.y, contentWidth, rowH).fillAndStroke(index % 2 ? '#ffffff' : ice, line)
      this.doc.font('Helvetica-Bold').fontSize(9).fillColor(ink).text(row.label, x + 6, this.y + 12, { width: widths[0] - 12 })
      row.factors.forEach((factor, factorIndex) => {
        const fx = x + widths[0] + 8 + factorIndex * ((widths[1] - 16) / 3)
        this.drawCheck((posture.factors || []).includes(factor), fx, this.y + 14)
        this.doc.font('Helvetica').fontSize(7.8).fillColor(ink).text(factor, fx + 13, this.y + 13, { width: (widths[1] - 16) / 3 - 15 })
      })
      const riskColor = posture.risk === 'H' ? danger : posture.risk === 'M' ? medium : low
      this.doc.roundedRect(x + widths[0] + widths[1] + 15, this.y + 10, 30, 18, 2).fill(riskColor)
      this.doc.font('Helvetica-Bold').fontSize(10).fillColor(ink).text(posture.risk || 'L', x + widths[0] + widths[1] + 15, this.y + 14, { width: 30, align: 'center' })
      this.y += rowH
    })
    this.y += 12
  }

  sectionIV() {
    ;(this.report.exhibits || []).forEach((exhibit, index) => {
      this.addPage('IV', 'Evidentiary Documentation')
      this.drawSubhead(`Exhibit ${index + 1} - Photographic Evidence`)
      this.drawPhotoPair(exhibit)
      this.drawTextCard('Analysis', formatAnalysisText(exhibit.analysis), {
        fontSize: 7.75,
        minHeight: 0,
        lineGap: 3.1,
        blockGap: 10,
        padding: 10,
        innerWidth: contentWidth - 20,
        align: 'justify',
        after: 12,
      })
    })
  }

  drawPhotoPair(exhibit) {
    const hasPre = Boolean(exhibit?.pre)
    const hasPost = Boolean(exhibit?.post)
    if (!hasPre && !hasPost) return

    const gap = 8
    const frameH = 210

    // Two-photo layout
    if (hasPre && hasPost) {
      const frameW = (contentWidth - gap) / 2
      this.ensureSpace(24 + frameH + 12)
      const top = this.y
      this.drawPhotoLabel('Initial Findings (Pre)', margin.left, top, frameW)
      this.drawPhotoLabel('Corrective Action', margin.left + frameW + gap, top, frameW)
      this.drawPhotoFrame(exhibit.pre, margin.left, top + 25, frameW, frameH)
      this.drawPhotoFrame(exhibit.post, margin.left + frameW + gap, top + 25, frameW, frameH)
      this.y = top + 25 + frameH + 14
      return
    }

    // Single-photo layout occupies full width
    const frameW = contentWidth
    this.ensureSpace(24 + frameH + 12)
    const top = this.y
    if (hasPre) {
      this.drawPhotoLabel('Initial Findings (Pre)', margin.left, top, frameW)
      this.drawPhotoFrame(exhibit.pre, margin.left, top + 25, frameW, frameH)
    } else {
      this.drawPhotoLabel('Corrective Action', margin.left, top, frameW)
      this.drawPhotoFrame(exhibit.post, margin.left, top + 25, frameW, frameH)
    }
    this.y = top + 25 + frameH + 14
  }

  drawPhotoLabel(text, x, y, width) {
    this.doc.roundedRect(x, y, width, 20, 3).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(9).fillColor('#ffffff').text(text, x + 6, y + 6, { width: width - 12, align: 'center' })
  }

  drawPhotoFrame(photo, x, y, width, height) {
    this.doc.roundedRect(x, y, width, height, 5).fillAndStroke(paper, navy)
    const image = bufferFromPhoto(photo)
    if (image) {
      try {
        this.doc.image(image, x + 4, y + 4, { fit: [width - 8, height - 8], align: 'center', valign: 'center' })
        return
      } catch {
        // Fall through to placeholder if PDFKit cannot read the image.
      }
    }
    this.doc.font('Helvetica-Oblique').fontSize(10).fillColor(muted).text('Photo Placeholder', x, y + height / 2 - 6, { width, align: 'center' })
  }

  sectionV() {
    this.addPage('V', 'Findings & Corrective Action Plan')
    this.drawOptionBlock('Administrative Controls & Training', controls, this.report.controls, [
      'Educated on 20/20/20 rule and movement pacing.',
      `Reviewed use of: ${softwareOptions.join(', ')}.`,
      'Trained on chair mechanics and monitor height optimization.',
    ])
    this.drawJustificationChecklist()
    this.addPage('V', 'Clinical Justification Display')
    this.drawTextCard('', formatJustificationDisplay(this.report.justificationDisplay), {
      fontSize: 8.05,
      lineGap: 1.65,
      blockGap: 10,
      padding: 12,
      maxHeight: 410,
      continuationMinHeight: 190,
      continuationMaxHeight: 500,
      keepBlocksWhole: true,
      fill: '#ffffff',
      border: '#b4c6d9',
    })
  }

  drawClinicalJustificationCards(text) {
    const blocks = this.paragraphBlocks(text)
    const bodyOptions = { size: 8.05, lineGap: 1.85 }
    blocks.forEach((block, index) => {
      if (index > 0 && index % 3 === 0) {
        this.addPage('V', 'Clinical Justification Display', { continued: true })
      }
      const match = block.match(/^([^:\n]{4,160}):\s*([\s\S]*)$/)
      const title = cleanText(match ? match[1] : (index === 0 ? 'Clinical Rationale' : 'Clinical Rationale Continued'))
      let body = segmentLongParagraphs(cleanText(match ? match[2] : block))
      if (!body) body = '-'
      let continued = 0
      while (body) {
        const pageCapacity = this.bottom() - pageTopY
        const fullBodyH = this.textHeight(body, contentWidth - 28, bodyOptions)
        const fullCardH = fullBodyH + 52
        if (this.y + fullCardH > this.bottom() && fullCardH <= pageCapacity) {
          this.addPage('V', 'Clinical Justification Display', { continued: true })
        }
        const available = this.bottom() - this.y
        const maxBodyH = Math.max(70, available - 62)
        const fits = this.textHeight(body, contentWidth - 28, bodyOptions) <= maxBodyH
        const segment = fits ? { fit: body, rest: '' } : this.splitBlockToFit(body, contentWidth - 28, maxBodyH, bodyOptions)
        this.drawJustificationCardSegment(continued ? `${title} Continued` : title, segment.fit || body, bodyOptions)
        body = cleanText(segment.rest)
        continued += 1
        if (body) this.addPage('V', 'Clinical Justification Display', { continued: true })
      }
    })
  }

  drawJustificationCardSegment(title, body, bodyOptions) {
    const bodyH = this.textHeight(body, contentWidth - 28, bodyOptions)
    const cardH = Math.max(70, bodyH + 52)
    this.ensureSpace(cardH + 10)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, cardH, 5).fillAndStroke('#ffffff', '#b4c6d9')
    this.doc.rect(margin.left, top, 4, cardH).fill(gold)
    this.doc.roundedRect(margin.left + 4, top, contentWidth - 4, 24, 4).fill(ice)
    this.doc.font('Helvetica-Bold').fontSize(8.4).fillColor(navy).text(title, margin.left + 12, top + 8, {
      width: contentWidth - 24,
      height: 10,
      lineBreak: false,
    })
    this.doc.font('Helvetica').fontSize(bodyOptions.size).fillColor(ink).text(reflowText(body), margin.left + 14, top + 34, {
      width: contentWidth - 28,
      lineGap: bodyOptions.lineGap,
      continued: false,
      lineBreak: true,
      align: 'justify',
    })
    this.y = top + cardH + 10
  }

  drawJustificationChecklist() {
    this.drawSubhead('Expert Justification Repository')
    this.doc.font('Helvetica-Oblique').fontSize(8.5).fillColor(muted).text('Selected items populate the editable clinical justification display.', margin.left, this.y, { width: contentWidth })
    this.y += 16
    const colW = (contentWidth - 8) / 2
    const rowH = 20
    justifications.forEach((item, index) => {
      const col = index % 2
      if (col === 0) this.ensureSpace(rowH)
      const x = margin.left + col * (colW + 8)
      this.doc.roundedRect(x, this.y, colW, rowH - 3, 3).fillAndStroke(paper, '#c4d0de')
      this.drawCheck((this.report.justifications || []).includes(item), x + 5, this.y + 4)
      this.doc.font('Helvetica').fontSize(8.4).fillColor(ink).text(item, x + 20, this.y + 4, { width: colW - 25 })
      if (col === 1 || index === justifications.length - 1) this.y += rowH
    })
  }

  sectionVI() {
    this.addPage('VI', 'Ergonomic Equipment Recommendation')
    const rows = visibleRecommendationRows(this.report.recommendations)
    if (!rows.length) {
      this.drawEmptyState('No equipment recommendations have been entered for this report.')
    } else {
      rows.forEach((row) => this.drawEquipmentCard(row))
    }
    this.addPage('VI', 'Procurement Notes & Justification Summary')
    this.drawProcurementSummary(this.normalizedProcurementText())
  }

  drawEquipmentCard(row) {
    const action = cleanInlineText(row.action) || '-'
    const model = cleanEquipmentModel(row.current) || '-'
    const category = cleanInlineText(row.category) || 'Equipment Recommendation'
    const url = extractUrl(row.notes)
    const headerH = 29
    const bodyX = margin.left + 12
    const bodyW = contentWidth - 24
    const panelGap = 5
    const modelTextW = url ? bodyW - 118 : bodyW - 18
    const modelH = Math.max(42, this.textHeight(model, modelTextW, { size: 8.35, lineGap: 0.9 }) + 25)
    const actionH = Math.max(58, this.textHeight(action, bodyW - 20, { size: 8.35, lineGap: 1.05 }) + 29)
    const cardH = headerH + 6 + modelH + panelGap + actionH + 8
    this.ensureSpace(cardH + 14)
    const top = this.y
    this.doc.lineWidth(1.15)
    this.doc.roundedRect(margin.left, top, contentWidth, cardH, 7).fillAndStroke('#f7f9fc', '#b8c8d9')
    this.doc.rect(margin.left, top, 6, cardH).fill(gold)
    this.doc.roundedRect(margin.left + 6, top, contentWidth - 6, headerH, 6).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(10.4).fillColor('#ffffff').text(category, margin.left + 15, top + 8.6, {
      width: contentWidth - 128,
      height: 12,
      lineBreak: false,
    })
    this.drawPriority(row.priority, margin.left + contentWidth - 62, top + 5.5)
    const modelY = top + headerH + 6
    this.doc.roundedRect(bodyX, modelY, bodyW, modelH, 4).fillAndStroke('#ffffff', '#d0dce9')
    this.doc.rect(bodyX, modelY, 3, modelH).fill('#8fb3d1')
    this.doc.font('Helvetica-Bold').fontSize(7).fillColor('#7a8798').text('Product / model', bodyX + 10, modelY + 7, { width: 105, height: 9, lineBreak: false })
    this.doc.font('Helvetica').fontSize(8.15).fillColor('#4d627a').text(model, bodyX + 10, modelY + 19, { width: modelTextW, lineGap: 0.9 })
    if (url) {
      const linkX = bodyX + bodyW - 100
      this.doc.roundedRect(linkX, modelY + 17, 86, 17, 4).fillAndStroke('#f5fbff', '#8eb5dd')
      this.doc.font('Helvetica-Bold').fontSize(6.7).fillColor('#005fa8').text('Open Purchase Link', linkX, modelY + 22.2, {
        width: 86,
        height: 8,
        align: 'center',
        link: url,
        underline: true,
        lineBreak: false,
      })
    }
    const actionY = modelY + modelH + panelGap
    this.doc.roundedRect(bodyX, actionY, bodyW, actionH, 4).fillAndStroke('#fbfdff', '#c6d5e5')
    this.doc.rect(bodyX, actionY, 3, actionH).fill(gold)
    this.doc.font('Helvetica-Bold').fontSize(7.4).fillColor(navy).text('Recommended action', bodyX + 10, actionY + 8, { width: bodyW - 20, height: 9, lineBreak: false })
    this.doc.font('Helvetica').fontSize(8.45).fillColor(ink).text(action, bodyX + 10, actionY + 21, { width: bodyW - 20, lineGap: 1.05 })
    this.doc.strokeColor('#244f80').lineWidth(0.8).roundedRect(margin.left, top, contentWidth, cardH, 7).stroke()
    this.y = top + cardH + 14
  }

  drawPriority(priority, x, y) {
    const normalized = priority || ''
    const fill = normalized === 'High' ? danger : normalized === 'Medium' ? medium : normalized === 'Low' ? low : '#edf2f7'
    this.doc.roundedRect(x, y, 50, 18, 9).fillAndStroke(fill, '#ffffff')
    this.doc.font('Helvetica-Bold').fontSize(8).fillColor(ink).text(normalized || '-', x, y + 5.2, { width: 50, align: 'center' })
  }

  normalizedProcurementText() {
    const text = formatProcurementSummary(this.report.procurementSummary)
    if (/-\s+/.test(text)) return text
    const known = [
      'Two 48"-55" sit-stand desks',
      'Full back, head and neck support ergonomic chair',
      'Adjustable height, telescopic (max 21") laptop riser',
      'Dual monitor arms system',
      'Document holder',
      'Under-Desk Keyboard tray',
    ]
    const found = known.filter((item) => text.toLowerCase().includes(item.toLowerCase()))
    if (found.length) {
      const intro = text.match(/^([\s\S]*?(?:configuration|recommended items)\s*:)/i)?.[1] || 'The recommended items for this ergonomic configuration:'
      const closing = text.match(/These items can\b[\s\S]*$/i)?.[0] || ''
      return [intro, ...found.map((item) => `- ${item}`), closing].filter(Boolean).join('\n')
    }
    // Fall back to auto-generating bullets from the recommendations list
    const recBullets = (this.report.recommendations || [])
      .filter((r) => String(r.category || '').trim())
      .map((r) => {
        const cat = String(r.category || '').trim()
        const current = String(r.current || '').trim()
        return `- ${cat}${current ? ` — ${current}` : ''}`
      })
    if (recBullets.length) {
      const introLine = text.match(/^[^\n-]+/)?.[0]?.trim() || 'Recommended Equipment'
      return [introLine, ...recBullets].join('\n')
    }
    return text
  }

  drawProcurementSummary(text) {
    const lines = cleanText(text).split(/\n+/).map((line) => line.trim()).filter(Boolean)
    const bulletLines = lines.filter((line) => /^-\s+/.test(line)).map((line) => line.replace(/^-\s+/, ''))
    const nonBullets = lines.filter((line) => !/^-\s+/.test(line))
    const intro = nonBullets[0] || 'Recommended procurement items:'
    const closing = nonBullets.slice(1).join(' ')
    const introH = this.textHeight(intro, contentWidth - 28, { size: 8.7, lineGap: 1 })
    const bulletHeights = bulletLines.map((line) => Math.max(16, this.textHeight(line, contentWidth - 48, { size: 8.8, lineGap: 1.05 }) + 5))
    const footerText = 'Delivery ETA and vendor availability should be confirmed during procurement.'
    const closingH = closing ? this.textHeight(closing, contentWidth - 28, { size: 8.2, italic: true, lineGap: 0.9 }) + 8 : 0
    const cardH = Math.max(165, 72 + introH + bulletHeights.reduce((sum, h) => sum + h, 0) + closingH)
    this.ensureSpace(cardH + 10)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, cardH, 6).fillAndStroke('#ffffff', '#244f80')
    this.doc.rect(margin.left, top, 6, cardH).fill(gold)
    this.doc.roundedRect(margin.left + 6, top, contentWidth - 6, 28, 5).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(9.2).fillColor('#ffffff').text('Procurement Summary', margin.left + 16, top + 10, {
      width: contentWidth - 28,
      height: 10,
      lineBreak: false,
    })
    let y = top + 40
    this.doc.font('Helvetica').fontSize(8.7).fillColor(ink).text(intro, margin.left + 14, y, {
      width: contentWidth - 28,
      lineGap: 1,
    })
    y += introH + 10
    ;(bulletLines.length ? bulletLines : ['No procurement items entered.']).forEach((line, index) => {
      const rowH = bulletLines.length ? bulletHeights[index] : 18
      this.doc.circle(margin.left + 21, y + 5.6, 3.1).fill(gold)
      this.doc.font('Helvetica-Bold').fontSize(8.8).fillColor(ink).text(line, margin.left + 33, y, {
        width: contentWidth - 48,
        lineGap: 1.05,
      })
      y += rowH
    })
    if (closing) {
      y += 3
      this.doc.font('Helvetica-Oblique').fontSize(8.2).fillColor(muted).text(closing, margin.left + 14, y, {
        width: contentWidth - 28,
        lineGap: 0.9,
      })
    }
    const footerY = top + cardH - 28
    this.doc.roundedRect(margin.left + 12, footerY, contentWidth - 24, 18, 3).fillAndStroke('#f5fbff', '#c8d8e8')
    this.doc.font('Helvetica-Bold').fontSize(7.6).fillColor(navy).text(footerText, margin.left + 20, footerY + 5.5, {
      width: contentWidth - 40,
      height: 8,
      lineBreak: false,
    })
    this.y = top + cardH + 12
  }

  drawEmptyState(text) {
    this.ensureSpace(70)
    this.doc.roundedRect(margin.left, this.y, contentWidth, 56, 5).fillAndStroke(paper, line)
    this.doc.font('Helvetica-Oblique').fontSize(10).fillColor(muted).text(text, margin.left + 12, this.y + 20, { width: contentWidth - 24, align: 'center' })
    this.y += 68
  }

  sectionVII() {
  this.addPage('VII', 'Opinion & Professional Certification')

  this.drawOpinionBlock()
  this.y += 6

  const summary = formatEvaluatorSummary(this.report.evaluatorSummary)

  if (cleanText(summary)) {
    this.drawTextCard('Evaluator Summary', summary, {
      fontSize: 8.8,
      minHeight: 0,
      lineGap: 1.6,
      blockGap: 8,
      padding: 10,
      innerWidth: contentWidth - 20,
      align: 'justify',
      after: 16,
    })
  }

  if (this.bottom() - this.y < 240) {
    this.addPage('VII', 'Opinion & Professional Certification', { continued: true })
  }

  this.y += 8

  this.drawSignatureBlock()

  this.y += 6

  this.drawCorrespondence()
}

  drawOpinionBlock() {
    const text = 'It is my professional opinion, held to a reasonable degree of certainty within the field of ergonomics, that the recommended modifications significantly reduce the identified risk factors for cumulative trauma.'
    const h = 58
    this.ensureSpace(h + 10)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, h, 5).fillAndStroke('#f7fbff', '#9fb4cc')
    this.doc.rect(margin.left, top, 5, h).fill(gold)
    this.doc.font('Helvetica-Bold').fontSize(10.2).fillColor(navy).text('Professional Opinion', margin.left + 14, top + 10, { width: 180, height: 12, lineBreak: false })
    this.doc.font('Helvetica-Oblique').fontSize(8.7).fillColor(ink).text(text, margin.left + 14, top + 28, {
      width: contentWidth - 28,
      lineGap: 1.3,
    })
    this.y = top + h + 12
  }

  drawSignatureBlock() {
    const h = 154
    this.ensureSpace(h + 10)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, h, 6).fillAndStroke('#ffffff', '#244f80')
    this.doc.rect(margin.left, top, 5, h).fill(gold)
    this.doc.roundedRect(margin.left + 5, top, contentWidth - 5, 28, 5).fill(navy)
    this.doc.font('Helvetica-Bold').fontSize(9.6).fillColor('#ffffff').text('Professional Certification', margin.left + 16, top + 10, {
      width: contentWidth - 32,
      height: 10,
      lineBreak: false,
    })
    const labelY = top + 40
    const valueY = top + 53
    this.doc.font('Helvetica-Bold').fontSize(7.9).fillColor(navy)
      .text('Evaluator Name (Printed)', margin.left + 14, labelY, { width: 120 })
      .text('Date', margin.left + 395, labelY, { width: 36 })
    this.doc.font('Helvetica').fontSize(8.2).fillColor(ink)
      .text(this.report.evaluatorName || '-', margin.left + 135, labelY, { width: 235 })
      .text(this.report.certificationDate || '-', margin.left + 430, labelY, { width: 78, align: 'center' })
    this.doc.strokeColor('#9aacc1').lineWidth(0.7)
      .moveTo(margin.left + 135, valueY).lineTo(margin.left + 372, valueY)
      .moveTo(margin.left + 430, valueY).lineTo(margin.left + 508, valueY)
      .stroke()

    const sigTop = top + 64
    if (this.logo) this.doc.image(this.logo, margin.left + 16, sigTop + 7, { fit: [42, 36] })
    this.doc.roundedRect(margin.left + 66, sigTop, contentWidth - 82, 48, 3).fillAndStroke('#fbfdff', '#526d8f')
    const sig = bufferFromPhoto(this.report.signature)
    if (sig) {
      try {
        this.doc.image(sig, margin.left + 76, sigTop + 6, { fit: [contentWidth - 102, 36], align: 'center', valign: 'center' })
      } catch {
        this.doc.font('Helvetica-Oblique').fontSize(8.5).fillColor(muted).text('Electronic Signature Pending', margin.left + 66, sigTop + 18, { width: contentWidth - 82, align: 'center' })
      }
    } else {
      this.doc.font('Helvetica-Oblique').fontSize(8.5).fillColor(muted).text('Electronic Signature Pending', margin.left + 66, sigTop + 18, { width: contentWidth - 82, align: 'center' })
    }

    const id = `${String(this.report.id || 'REPORT00').slice(0, 8).toUpperCase()}-${String(this.report.signatureSignedAt || this.report.certificationDate || '').replace(/[^0-9]/g, '').slice(0, 8)}`
    const secureTop = top + 120
    this.doc.roundedRect(margin.left + 12, secureTop, contentWidth - 24, 26, 4).fillAndStroke('#f6fbff', '#b7c6d8')
    this.doc.font('Helvetica-Bold').fontSize(8.4).fillColor(navy).text('Secure Electronic Signature Block', margin.left + 20, secureTop + 6, { width: 190 })
    this.doc.font('Helvetica').fontSize(7.3).fillColor(ink)
      .text(`Status: ${this.report.signatureConfirmed ? 'Certified by evaluator' : 'Not certified'}`, margin.left + 20, secureTop + 17, { width: 155 })
      .text(`Signature ID: ${id || 'Pending'}`, margin.left + 230, secureTop + 17, { width: 155, align: 'center' })
      .text(`Timestamp: ${this.report.signatureSignedAt ? new Date(this.report.signatureSignedAt).toLocaleString() : 'Pending'}`, margin.left + 382, secureTop + 17, { width: 138, align: 'right' })
    this.y = top + h + 12
  }

  drawCorrespondence() {
    const text = 'Should you have any questions or clinical concerns regarding the findings of this report, please contact our administrative team at Info@ergonomicworks.com. If you are corresponding with the evaluator directly, please ensure that Info@ergonomicworks.com is copied on all emails to maintain a formal record and ensure a timely response.'
    const footer = 'This report was prepared by a certified ergonomic specialist. All findings are based on direct observation, interview, and validated assessment protocols.'
    const mainH = this.textHeight(text, contentWidth - 20, { size: 7.7, italic: true, lineGap: 0.7 })
    const footerH = this.textHeight(footer, contentWidth - 20, { size: 7, italic: true, lineGap: 0.6 })
    const boxH = mainH + footerH + 44
    this.ensureSpace(boxH + 10)
    const top = this.y
    this.doc.roundedRect(margin.left, top, contentWidth, boxH, 4).fillAndStroke('#f7fbff', '#c6d5e5')
    this.doc.roundedRect(margin.left, top, contentWidth, 22, 4).fill(ice)
    this.doc.rect(margin.left, top, 4, boxH).fill(gold)
    this.doc.font('Helvetica-Bold').fontSize(8.8).fillColor(navy).text('Inquiries & Correspondence', margin.left + 12, top + 7, {
      width: contentWidth - 24,
      height: 10,
      lineBreak: false,
    })
    this.doc.font('Helvetica-BoldOblique').fontSize(7.7).fillColor('#0d2e57').text(text, margin.left + 10, top + 30, { width: contentWidth - 20, lineGap: 0.7 })
    this.doc.font('Helvetica-Oblique').fontSize(7).fillColor(muted).text(footer, margin.left + 10, top + 34 + mainH, { width: contentWidth - 20, lineGap: 0.6 })
    this.y = top + boxH + 8
  }

  finalizePageNumbers() {
    const range = this.doc.bufferedPageRange()
    this.pageCount = range.count
    for (let index = range.start; index < range.start + range.count; index += 1) {
      this.doc.switchToPage(index)
      const label = `Page ${index + 1} of ${range.count}`
      this.doc.strokeColor('#d9e1ea').lineWidth(0.6).moveTo(margin.left, 714).lineTo(margin.left + contentWidth, 714).stroke()
      this.doc.font('Helvetica').fontSize(7.5).fillColor('#898f98').text(label, margin.left, 722, {
        width: contentWidth,
        height: 10,
        align: 'right',
        lineBreak: false,
      })
    }
  }
}

function renderReportPdfPass(report, options = {}) {
  return new Promise((resolve, reject) => {
    const renderer = new ReportPdfRenderer(report, options)
    const chunks = []
    renderer.doc.on('data', (chunk) => chunks.push(chunk))
    renderer.doc.on('end', () => resolve({
      buffer: Buffer.concat(chunks),
      sectionStarts: { ...renderer.sectionStarts },
      pageCount: renderer.pageCount,
    }))
    renderer.doc.on('error', reject)
    try {
      renderer.render()
      renderer.doc.end()
    } catch (error) {
      reject(error)
    }
  })
}

async function renderReportPdfBuffer(report) {
  const firstPass = await renderReportPdfPass(report)
  const secondPass = await renderReportPdfPass(report, { tocPageNumbers: firstPass.sectionStarts })
  return secondPass.buffer
}

module.exports = {
  renderReportPdfBuffer,
}
