import React, { useEffect, useMemo, useState } from 'react'
import type { ChangeEvent } from 'react'
import { createRoot } from 'react-dom/client'
import {
  Bug,
  Camera,
  Download,
  FilePlus2,
  FileUp,
  FolderOpen,
  PenLine,
  Save,
  Trash2,
} from 'lucide-react'
import { v4 as uuid } from 'uuid'
import './styles.css'

type Risk = 'L' | 'M' | 'H'
type Priority = '' | 'High' | 'Medium' | 'Low'
type PhotoSide = 'pre' | 'post'

const caseTypes = [
  'Preventative Evaluation',
  'ADA Interactive Process',
  'Workers’ Compensation Claim',
] as const

type CaseType = (typeof caseTypes)[number]

type PhotoSlot = {
  name: string
  path: string
  url: string
}

type ImportedSourceDocument = {
  name: string
  path: string
  url: string
  type: string
  extension: string
  importedAt: string
  text: string
  extractionStatus: string
  exhibitPhotos?: Array<{
    exhibitIndex: number
    side: PhotoSide
    photo: PhotoSlot
  }>
  signaturePhoto?: PhotoSlot
}

type Exhibit = {
  pre?: PhotoSlot
  post?: PhotoSlot
  analysis: string
}

type EquipmentRow = {
  category: string
  current: string
  action: string
  priority: Priority
  notes: string
}

type Report = {
  id: string
  caseType: CaseType
  claimNumber: string
  employeeName: string
  employeeId: string
  department: string
  location: string
  dateOfEval: string
  reportDate: string
  managerName: string
  jobTitle: string
  assessmentMode: string
  dataSources: string[]
  assessmentTools: string[]
  benchmarks: string[]
  stature: string
  dominantHand: 'R' | 'L'
  vision: string
  deskType: string
  currentDeskHeight: string
  idealKeyingHeight: string
  delta: string
  postures: Record<string, { factors: string[]; risk: Risk }>
  posturalObservations: string
  exhibits: Exhibit[]
  recommendations: EquipmentRow[]
  controls: string[]
  software: string[]
  justifications: string[]
  justificationDisplay: string
  procurementSummary: string
  evaluatorSummary: string
  evaluatorName: string
  certificationDate: string
  signature?: PhotoSlot
  signatureConfirmed: boolean
  signatureSignedAt: string
  sourceDocument?: ImportedSourceDocument
  importedSourceText: string
  updatedAt?: string
}

type ReportSummary = {
  id: string
  title: string
  updatedAt: string
}

type PdfImportResult = {
  filePath: string
  report?: Partial<Report>
  sourceDocument?: ImportedSourceDocument
  warning?: string
  error?: string
}

declare global {
  interface Window {
    reportApi?: {
      listReports: () => Promise<ReportSummary[]>
      loadReport: (id: string) => Promise<Report | null>
      saveReport: (report: Report) => Promise<Report>
      deleteReport: (id: string) => Promise<boolean>
      savePhoto: (photo: { name: string; dataUrl: string }) => Promise<PhotoSlot>
      importPdf: () => Promise<PdfImportResult | null>
      exportPdf: (title: string, report: Report) => Promise<string | null>
      exportDiagnostics: () => Promise<string | null>
    }
  }
}

const logo = `${import.meta.env.BASE_URL}assets/extracted-1.jpeg`
const today = new Date().toISOString().slice(0, 10)
const exhibitCount = 3
const observationTextLimit = 900
const observationHardLimit = 1080
const observationContinuationLimit = 2850
const observationContinuationHardLimit = 3350
const singleExhibitAnalysisLimit = 1150
const singleExhibitAnalysisHardLimit = 1320
const exhibitAnalysisContinuationLimit = 3300
const exhibitAnalysisContinuationHardLimit = 3750
const minimumContinuationText = 420
const equipmentRowsPerPage = 3
const evaluatorCertificationTextLimit = 650
const justificationTextLimit = 2050
const justificationHardLimit = 2450
const justificationContinuationLimit = 2450
const justificationContinuationHardLimit = 2900
const posturalObservationsRecommendedLimit = 1500
const exhibitAnalysisRecommendedLimit = 1200
const procurementSummaryRecommendedLimit = 900
const evaluatorSummaryRecommendedLimit = 650
const justificationRecommendedLimit = 2200
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

const clinicalJustifications = [
  {
    title: 'Neutral Wrist Alignment',
    text: 'The carpal tunnel is a narrow, rigid passageway of ligament and bone at the base of the hand housing the median nerve and tendons. Maintaining a 0° neutral plane is the single most effective way to prevent nerve hypoxia, as pressures above 30 mmHg can impede blood flow to the nerve.',
  },
  {
    title: 'Pronation Mitigation',
    text: 'Standard computer input devices typically require full pronation (palm flat), requiring sustained isometric contraction of the pronator teres and brachioradialis which can lead to tennis elbow. A vertical or slanted interface is recommended to transition the forearm into a handshake position (45° to 90°), offloading extensor tendons and reducing muscular oxygen demand.',
  },
  {
    title: 'Ulnar Deviation Correction',
    text: 'Standard keyboards force the wrists to fan outward in ulnar deviation, compressing the ulnar artery and nerve against the carpal tunnel walls. Split keyboard geometry is justified to allow linear wrist to forearm alignment, reducing deviation from approximately 15 degrees to less than 5 degrees, preserving circulation and preventing tenosynovitis.',
  },
  {
    title: 'Neutral Shoulder Adduction',
    text: 'Shoulder stress is frequently caused by "reaching" (abduction) for a mouse or numeric keypad, forcing the humerus away from the torso and compressing rotator cuff tendons. To achieve shoulder neutrality, the elbows must remain tucked close to the ribs; a compact or "tenkeyless" keyboard allows the mouse to be positioned closer to the midline, eliminating outward reach and shoulder impingement risks.',
  },
  {
    title: 'Elbow Flexion Management',
    text: 'Sustained elbow flexion greater than 90° can stretch and irritate the ulnar nerve at the cubital tunnel ("funny bone"). Ensuring workstation height allows for an open elbow angle (90° to 110°) reduces mechanical tension on the nerve and prevents the onset of Cubital Tunnel Syndrome.',
  },
  {
    title: 'Cervical Neutrality and Gaze Dynamics',
    text: 'The human head weighs approximately 12 lbs, but for every inch the head tilts forward, the effective weight on the neck doubles. To achieve neutrality, the top third of the monitor must be at eye level, allowing for a slight downward gaze of 15° to 20°, which releases the levator scapulae from constant tension.',
  },
  {
    title: 'Lumbar Integrity and Lordotic Support',
    text: 'Sitting flattens the natural lordotic curve and forces intervertebral discs to bulge. Adjustable lumbar support acts as a mechanical brace that preserves this curve and shifts body weight from soft tissues to the structural bones of the spine, which is a critical necessity for managing intradiscal pressure.',
  },
  {
    title: 'Thoracic/Scapular Support',
    text: 'High-back or contoured seating configurations are required to support the scapular (mid-back) region. This reduces the isometric workload on the erector spinae and trapezius muscles, preventing mid-back fatigue and slumped postures that lead to secondary cervical strain.',
  },
  {
    title: 'Ocular Accommodation & Glare',
    text: 'Digital eye strain occurs when ciliary muscles are forced to maintain a constant focal length for extended periods. Proper focal distance (20–30 inches) and the "20-20-20 Rule" are required to relax ciliary muscles and reduce visual fatigue.',
  },
  {
    title: 'Midline Alignment (The Neutral Pathway)',
    text: 'Repetitive neck rotation (twisting to look at papers) causes asymmetrical loading. Implementing midline alignment through an inline document holder ensures the head remains in a neutral, forward-facing orientation and prevents repetitive twisting motions.',
  },
  {
    title: 'Plantar Surface Stability',
    text: 'If feet dangle or are tucked back, the weight of the legs pulls on the lower back. A height-adjustable footrest is required to ensure the feet are flat and the ankles are at a neutral angle, which stabilizes the pelvis and ensures the user maintains contact with the chair\'s lumbar support.',
  },
  {
    title: 'Pregnancy Postural Support',
    text: 'Hormonal laxity (relaxin) and a shifting center of gravity increase the risk of SI joint instability. A chair with a tilting seat pan is recommended to open the hip angle, reducing upward pressure on the diaphragm and supporting the increased lumbar load associated with abdominal weight gain.',
  },
  {
    title: 'Pregnancy Circulatory Enhancement',
    text: 'Increased blood volume and uterine pressure on pelvic veins increase the risk of edema. A sit-stand workstation allows for postural transitions that activate the "calf pump" to return blood to the heart; when seated, a footrest encourages drainage by keeping knees slightly higher than hips.',
  },
  {
    title: 'Metabolic Flux (Sit-Stand Transitions)',
    text: 'Human physiology is optimized for movement. Postural flux (changing positions every 30–60 minutes) improves blood oxygenation and manages pain sensitivity by preventing the build-up of inflammatory markers in the muscles. A height-adjustable workstation is the primary tool for achieving this metabolic flux.',
  },
  {
    title: 'Contact Stress Mitigation',
    text: 'Contact stress occurs when body tissues are compressed against hard workstation surfaces. Padded supports and rounded desk edges reduce tissue compression and prevent localized irritation.',
  },
]

const justifications = clinicalJustifications.map((item) => item.title)

const oldJustificationMap: Record<string, string> = {
  'Neutral Wrist Alignment (The 0° Plane)': 'Neutral Wrist Alignment',
  'Pronation Mitigation (Handshake Position)': 'Pronation Mitigation',
  'Neutral Shoulder Adduction (Comfort Zone)': 'Neutral Shoulder Adduction',
}

const postureRows = [
  { key: 'neck', label: 'Neck / Head', factors: ['Extension', 'Rotation', 'Forward Lean'] },
  { key: 'shoulders', label: 'Shoulders', factors: ['Elevation', 'Protraction', 'Abduction'] },
  { key: 'wrists', label: 'Wrists / Arms', factors: ['Extension', 'Ulnar Deviation', 'Contact Stress'] },
  { key: 'back', label: 'Back / Trunk', factors: ['Slumping', 'Lack of Lumbar Support'] },
]

function newReport(): Report {
  return {
    id: uuid(),
    caseType: 'ADA Interactive Process',
    claimNumber: '',
    employeeName: '',
    employeeId: '',
    department: '',
    location: '',
    dateOfEval: today,
    reportDate: today,
    managerName: '',
    jobTitle: '',
    assessmentMode: 'In-Person',
    dataSources: [],
    assessmentTools: [],
    benchmarks: [],
    stature: '',
    dominantHand: 'R',
    vision: 'Single',
    deskType: 'Fixed',
    currentDeskHeight: '',
    idealKeyingHeight: '',
    delta: '',
    postures: {
      neck: { factors: [], risk: 'M' },
      shoulders: { factors: [], risk: 'L' },
      wrists: { factors: [], risk: 'L' },
      back: { factors: [], risk: 'L' },
    },
    posturalObservations: 'Enter detailed postural observations here...',
    exhibits: Array.from({ length: exhibitCount }, () => ({ analysis: '' })),
    recommendations: [
      { category: 'Mouse', current: '', action: '', priority: 'High', notes: '' },
      { category: 'Footrest', current: '', action: '', priority: 'Medium', notes: '' },
      { category: 'Other', current: '', action: '', priority: 'Low', notes: '' },
      { category: 'Mouse', current: '', action: '', priority: '', notes: '' },
      { category: 'Chair', current: '', action: '', priority: '', notes: '' },
      { category: 'Chair', current: '', action: '', priority: '', notes: '' },
      { category: 'Chair', current: '', action: '', priority: '', notes: '' },
      { category: 'Chair', current: '', action: '', priority: '', notes: '' },
    ],
    controls: [],
    software: [],
    justifications: [],
    justificationDisplay: '',
    procurementSummary: '',
    evaluatorSummary: '',
    evaluatorName: '',
    certificationDate: today,
    signatureConfirmed: false,
    signatureSignedAt: '',
    importedSourceText: '',
  }
}

export function normalizeReport(report: Partial<Report> | null): Report {
  const base = newReport()
  if (!report) return base
  const normalized: Report = {
    ...base,
    ...report,
    caseType: caseTypes.includes(report.caseType as CaseType)
      ? report.caseType as CaseType
      : 'ADA Interactive Process',
    postures: { ...base.postures, ...report.postures },
    exhibits: Array.from({ length: exhibitCount }, (_, index) => ({ ...base.exhibits[index], ...report.exhibits?.[index] })),
    recommendations: base.recommendations.map((baseRow, index) => {
      const row = report.recommendations?.[index]
      const priority = row?.priority && ['High', 'Medium', 'Low'].includes(row.priority) ? row.priority as Priority : ''
      return repairEquipmentRow({ ...baseRow, ...row, priority })
    }),
    justifications: (report.justifications || [])
      .map((item) => oldJustificationMap[item] || item)
      .filter((item, index, items) => justifications.includes(item) && items.indexOf(item) === index),
    signatureConfirmed: Boolean(report.signatureConfirmed),
    signatureSignedAt: report.signatureSignedAt || '',
  }
  if (!normalized.justificationDisplay.trim() && normalized.justifications.length > 0) {
    normalized.justificationDisplay = buildJustificationDisplay(normalized.justifications)
  }
  normalized.justificationDisplay = formatJustificationDisplay(normalized.justificationDisplay)
  normalized.procurementSummary = formatProcurementSummary(normalized.procurementSummary)
  normalized.evaluatorSummary = formatEvaluatorSummary(normalized.evaluatorSummary)
  ;(['stature', 'currentDeskHeight', 'idealKeyingHeight', 'delta'] as const).forEach((key) => {
    normalized[key] = trimInlineImportedValue(normalized[key])
  })
  return normalized
}

function formatJustificationBlock(title: string) {
  const item = clinicalJustifications.find((justification) => justification.title === title)
  if (!item) return ''
  return `${item.title}:\n${item.text}`
}

function buildJustificationDisplay(selected: string[]) {
  return selected.map(formatJustificationBlock).filter(Boolean).join('\n\n')
}

function formatJustificationDisplay(text: string) {
  return cleanImportedText(text)
    .replace(/\bEditable Clinical Justification Display\b\s*/gi, '')
    .replace(/([.!?])\s+(?=[A-Z][A-Za-z][A-Za-z0-9/&()'" -]{8,150}:)/g, '$1\n\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function formatEvaluatorSummary(text: string) {
  return cleanImportedText(text)
    .replace(/(?:^|\s)-\s*(?=[A-Z])/g, '\n- ')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function addJustificationBlock(display: string, title: string) {
  const block = formatJustificationBlock(title)
  if (!block || display.includes(`${title}:`)) return display
  return [display.trim(), block].filter(Boolean).join('\n\n')
}

function removeJustificationBlock(display: string, title: string) {
  return display
    .split(/\n{2,}/)
    .filter((block) => !block.trimStart().startsWith(`${title}:`))
    .join('\n\n')
}

type ChunkTextOptions = {
  hardMaxChars?: number
  nextMaxChars?: number
  nextHardMaxChars?: number
  minTailChars?: number
}

function tokenizeText(text: string) {
  return text
    .split(/\n{2,}/)
    .map((paragraph) => paragraph.trim())
    .filter(Boolean)
    .flatMap((paragraph, index) => {
      const words = paragraph.split(/\s+/).filter(Boolean)
      return index === 0 ? words : ['\n\n', ...words]
    })
}

function joinTextTokens(tokens: string[]) {
  return tokens.reduce((text, token) => {
    if (token === '\n\n') return `${text.trimEnd()}\n\n`
    return text && !text.endsWith('\n\n') ? `${text} ${token}` : `${text}${token}`
  }, '').trim()
}

function endsSentence(token: string) {
  return /[.!?;)]["']?$/.test(token)
}

function splitTokens(tokens: string[], maxChars: number): [string[], string[]] {
  let fitCount = 0
  for (let index = 0; index < tokens.length; index += 1) {
    const candidate = joinTextTokens(tokens.slice(0, index + 1))
    if (candidate.length > maxChars) break
    fitCount = index + 1
  }

  if (fitCount === 0) return [[tokens[0]], tokens.slice(1)]

  let breakAt = fitCount
  for (let index = 0; index < fitCount - 1; index += 1) {
    const candidate = joinTextTokens(tokens.slice(0, index + 1))
    if (candidate.length >= maxChars * 0.76 && endsSentence(tokens[index])) {
      breakAt = index + 1
    }
  }

  return [tokens.slice(0, breakAt), tokens.slice(breakAt)]
}

function rebalanceTinyTail(chunkTokens: string[], restTokens: string[], minTailChars: number): [string[], string[]] {
  if (!restTokens.length || joinTextTokens(restTokens).length >= minTailChars) return [chunkTokens, restTokens]

  let currentChunk = [...chunkTokens]
  let currentRest = [...restTokens]

  while (currentRest.length && joinTextTokens(currentRest).length < minTailChars && joinTextTokens(currentChunk).length > minTailChars) {
    let transferStart = Math.max(1, currentChunk.length - 14)
    for (let index = currentChunk.length - 2; index > 0; index -= 1) {
      if (endsSentence(currentChunk[index - 1]) && joinTextTokens(currentChunk.slice(index)).length >= 90) {
        transferStart = index
        break
      }
    }

    currentRest = [...currentChunk.slice(transferStart), ...currentRest]
    currentChunk = currentChunk.slice(0, transferStart)
  }

  return [currentChunk, currentRest]
}

function chunkText(text: string, maxChars: number, options: ChunkTextOptions = {}) {
  const clean = (text || '').replace(/\r\n/g, '\n').trim()
  if (!clean) return ['']
  let tokens = tokenizeText(clean)
  const chunks: string[] = []

  while (tokens.length) {
    const isFirstChunk = chunks.length === 0
    const softLimit = isFirstChunk ? maxChars : options.nextMaxChars || maxChars
    const hardLimit = isFirstChunk
      ? options.hardMaxChars || softLimit
      : options.nextHardMaxChars || options.nextMaxChars || options.hardMaxChars || softLimit
    const remaining = joinTextTokens(tokens)

    if (remaining.length <= hardLimit) {
      chunks.push(remaining)
      break
    }

    let [chunkTokens, restTokens] = splitTokens(tokens, softLimit)
    ;[chunkTokens, restTokens] = rebalanceTinyTail(chunkTokens, restTokens, options.minTailChars || 0)
    chunks.push(joinTextTokens(chunkTokens))
    tokens = restTokens
  }

  return chunks.length ? chunks : ['']
}

function normalizedPriority(value: string): Priority {
  const match = value.match(/\b(High|Medium|Low)\b/i)?.[1]
  if (!match) return ''
  return match.replace(/^./, (letter) => letter.toUpperCase()) as Priority
}

function splitEquipmentCategoryFromModel(lines: string[]) {
  const cleaned = lines.map((line) => line.trim()).filter(Boolean)
  if (cleaned.length < 2) return null
  const first = cleaned[0]
  const knownPrefixes = [
    'Under-Desk Keyboard Tray',
    'Document Holder',
    'Dual monitor arms system',
    'Adjustable height',
    'Chair',
    'Mouse',
    'Footrest',
    'Keyboard',
    'Monitor',
  ]

  for (const prefix of knownPrefixes) {
    const match = first.match(new RegExp(`^(${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})(?:\\s+(.+))?$`, 'i'))
    if (match) {
      return {
        category: match[1],
        model: [match[2] || '', ...cleaned.slice(1)].filter(Boolean),
      }
    }
  }

  if (/electric\s+sit-stand\s+desk/i.test(first) || /laptop\s+riser/i.test(first)) {
    return { category: first, model: cleaned.slice(1) }
  }

  return { category: first, model: cleaned.slice(1) }
}

function removeActionTextFromModel(current: string, action: string) {
  const actionKey = normalizeLookup(action)
  const currentLines = cleanImportedText(current)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)

  if (!actionKey || !currentLines.length) return cleanImportedText(current)

  const cleanedLines = currentLines
    .map((line) => {
      const priceMatch = /\$\s*\d[\d,.]*(?:\.\d{2})?/.exec(line)
      if (!priceMatch || priceMatch.index === undefined) return line

      const priceEnd = priceMatch.index + priceMatch[0].length
      const trailingText = line.slice(priceEnd).trim()
      const trailingKey = normalizeLookup(trailingText)
      if (trailingKey.length >= 18 && actionKey.includes(trailingKey)) {
        return line.slice(0, priceEnd).trim()
      }

      return line
    })
    .filter((line) => {
      const lineKey = normalizeLookup(line)
      if (!lineKey) return false
      if (lineKey.length < 24) return true
      if (/(?:\$|h\s*t\s*t\s*p)/i.test(line)) return true
      return !actionKey.includes(lineKey)
    })

  return cleanImportedText(cleanedLines.join('\n'))
}

function repairEquipmentRow(row: EquipmentRow): EquipmentRow {
  const repaired = {
    ...row,
    category: row.category === '-' ? '' : cleanImportedText(row.category || ''),
    current: cleanImportedText(row.current || ''),
    action: cleanImportedText(row.action || ''),
    notes: cleanImportedPurchaseLink(row.notes || ''),
  }

  ;(['Product / model', 'Model / SKU', 'Clinical Priority', 'Recommended action', 'Recommended Action / Risk Mitigation', 'Purchase Link', 'Open Purchase Link'] as const).forEach((label) => {
    const labelPattern = new RegExp(`(?:^|\\n)\\s*${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*:?\\s*`, 'gi')
    repaired.category = cleanImportedText(repaired.category.replace(labelPattern, '\n'))
    repaired.current = cleanImportedText(repaired.current.replace(labelPattern, '\n'))
    repaired.action = cleanImportedText(repaired.action.replace(labelPattern, '\n'))
  })

  const categoryPriority = repaired.category.match(/\b(High|Medium|Low)\s*$/i)?.[1]
  if (!repaired.priority && categoryPriority) repaired.priority = normalizedPriority(categoryPriority)
  if (categoryPriority) repaired.category = cleanImportedText(repaired.category.replace(/\b(High|Medium|Low)\s*$/i, ''))

  repaired.priority = normalizedPriority(repaired.priority)

  const currentLines = cleanImportedText(repaired.current)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
  const priorityLineIndex = currentLines.findIndex((line) => /^(High|Medium|Low)$/i.test(line))
  if (priorityLineIndex !== -1) {
    const priority = normalizedPriority(currentLines[priorityLineIndex])
    if (!repaired.priority) repaired.priority = priority

    const beforePriority = currentLines.slice(0, priorityLineIndex)
    const afterPriority = currentLines.slice(priorityLineIndex + 1)
    if (!repaired.category && beforePriority.length > 1) {
      repaired.category = beforePriority.shift() || ''
    }

    const modelLines = [...beforePriority]
    const possiblePriceLine = afterPriority[0] || ''
    if (possiblePriceLine && /(?:\$|\d+\s*(?:["”]|x)|-\s*\$)/i.test(possiblePriceLine) && !/h\s*t\s*t\s*p/i.test(possiblePriceLine)) {
      modelLines.push(possiblePriceLine)
    }
    if (modelLines.length) {
      repaired.current = cleanImportedText(modelLines.join('\n'))
    }

    const embeddedLink = afterPriority.find((line) => /h\s*t\s*t\s*p/i.test(line))
    if (!repaired.notes.trim() && embeddedLink) {
      repaired.notes = cleanImportedPurchaseLink(embeddedLink)
    }
  }

  repaired.current = cleanImportedText(repaired.current)
    .split('\n')
    .filter((line) => !/h\s*t\s*t\s*p/i.test(line))
    .join('\n')
    .trim()

  if (!repaired.priority) {
    repaired.priority = normalizedPriority(`${repaired.current}\n${repaired.action}`)
  }

  if (!repaired.category.trim()) {
    const split = splitEquipmentCategoryFromModel(
      repaired.current
        .split('\n')
        .map((line) => line.trim())
        .filter(Boolean),
    )
    if (split) {
      repaired.category = split.category
      repaired.current = cleanImportedText(split.model.join('\n'))
    }
  }

  if (repaired.current && repaired.action) {
    repaired.current = removeActionTextFromModel(repaired.current, repaired.action)
  }

  repaired.category = cleanImportedText(repaired.category.replace(/\s+-\s*$/g, ''))
  repaired.current = cleanImportedText(repaired.current
    .replace(/\s+Open\s+Purchase\s+Link\s*$/i, '')
    .replace(/\s+-\s*(?=\$)/g, ' - ')
    .replace(/\s{2,}/g, ' '))
  repaired.action = cleanImportedText(repaired.action)
  repaired.notes = cleanImportedPurchaseLink(repaired.notes)

  return repaired
}


function visibleRecommendationRows(rows: EquipmentRow[]) {
  return rows.filter((row) => [row.current, row.action, row.notes].some((value) => value.trim()))
}

function recommendationWeight(row: EquipmentRow) {
  const contentLength = [row.category, row.current, row.action, row.notes].join('\n').length
  if (contentLength > 900) return equipmentRowsPerPage
  if (contentLength > 520) return 2
  return 1
}

function chunkEquipmentRows(rows: EquipmentRow[]) {
  const groups: EquipmentRow[][] = []
  let current: EquipmentRow[] = []
  let currentWeight = 0

  rows.forEach((row) => {
    const weight = recommendationWeight(row)
    if (current.length && (current.length >= equipmentRowsPerPage || currentWeight + weight > equipmentRowsPerPage)) {
      groups.push(current)
      current = []
      currentWeight = 0
    }

    current.push(row)
    currentWeight += weight

    if (weight >= equipmentRowsPerPage) {
      groups.push(current)
      current = []
      currentWeight = 0
    }
  })

  if (current.length) groups.push(current)
  return groups
}

function cleanImportedText(text: string) {
  return (text || '')
    .replace(/\r/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/^\s*(?:Page\s*)?\d+\s*(?:of\s*\d+)?\s*$/gim, '')
    .replace(/^\s*Page\s*$/gim, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function importedLines(text: string) {
  return cleanImportedText(text)
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
}

function normalizeLookup(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '')
}

function isImportedLabelOrHeader(value: string) {
  const normalized = normalizeLookup(value)
  if (!normalized) return true
  const labels = [
    'casetype',
    'caseclaim',
    'employeename',
    'employeeid',
    'department',
    'locationfloor',
    'dateofeval',
    'reportdate',
    'managername',
    'jobtitle',
    'stature',
    'dominanthand',
    'vision',
    'desktype',
    'currentdeskhgt',
    'idealkeyinghgt',
    'delta',
    'equipmentrecommendation',
    'modelsku',
    'clinicalpriority',
    'recommendedactionriskmitigation',
    'purchaselink',
  ]
  return labels.some((label) => normalized === label || normalized.startsWith(label)) ||
    /^[ivx]+[.)]/i.test(value) ||
    /^section\b/i.test(value) ||
    /^pg$/i.test(value) ||
    /^(ergonomic evaluation report|report navigation|table of contents)$/i.test(value)
}

function trimInlineImportedValue(value: string) {
  const inlineStopLabels = [
    'Case Type',
    'Case / Claim #',
    'Case Claim #',
    'Employee Name',
    'Employee ID',
    'Department',
    'Location / Floor',
    'Location Floor',
    'Date of Eval',
    'Report Date',
    'Manager Name',
    'Job Title',
    'Dominant Hand',
    'Vision',
    'Desk Type',
    'Current Desk Hgt',
    'Ideal Keying Hgt',
    'Delta',
    'Risk Level',
    'Body Segment',
    'Observed Posture / Risk Factor',
    'Postural Observations',
    'Assessment Mode',
  ]
  const escapedStops = inlineStopLabels.map((label) => label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')
  const stopPattern = new RegExp(`\\s+(?:${escapedStops})(?:\\s*:|\\b)`, 'i')
  const sectionStopPattern = /\s+(?:[IVX]+\.\s+[A-Z][A-Za-z0-9 &/()—-]+|[A-Z]\.\s+[A-Z][A-Za-z0-9 &/()—-]+).*$/i
  const brokenMatrixPattern = /\s+B\.\s*Po\s*stural[\s\S]*$/i
  const match = stopPattern.exec(value)
  const sliced = match ? value.slice(0, match.index) : value
  return cleanImportedText(sliced
    .replace(brokenMatrixPattern, '')
    .replace(sectionStopPattern, '')
    .replace(/\s+(?:R\s+L|Single\s+Bifocal\s+Progressive)\s*$/i, '')
    .replace(/\s+(?:Fixed\s+Electric\s+Sit\/Stand\s+Desk\s+Converter)\s*$/i, ''))
}


function valueAfterImportedLabel(lines: string[], labels: string[]) {
  const normalizedLabels = labels.map(normalizeLookup)

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index]
    const lineLookup = normalizeLookup(line)
    const label = normalizedLabels.find((item) => lineLookup === item || lineLookup.startsWith(item))
    if (!label) continue

    const rawAfterLabel = line
      .replace(new RegExp(`^\\s*${labels.find((item) => normalizeLookup(item) === label)?.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*:?\\s*`, 'i'), '')
      .trim()
    const inlineValue = trimInlineImportedValue(rawAfterLabel)

    if (inlineValue && normalizeLookup(inlineValue) !== label && !isImportedLabelOrHeader(inlineValue)) {
      return inlineValue
    }

    for (let nextIndex = index + 1; nextIndex < Math.min(lines.length, index + 5); nextIndex += 1) {
      const candidate = lines[nextIndex]
      if (!candidate) continue
      const inlineCandidate = trimInlineImportedValue(candidate)
      if (isImportedLabelOrHeader(inlineCandidate)) return ''
      return inlineCandidate
    }
  }

  return ''
}

function stripImportedLabel(line: string, labels: string[]) {
  for (const label of labels) {
    const escapedLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const stripped = line.replace(new RegExp(`^\\s*${escapedLabel}\\s*:?\\s*`, 'i'), '').trim()
    if (stripped !== line.trim()) return stripped
  }
  return ''
}

function findImportedLabelIndex(lines: string[], labels: string[], startIndex = 0) {
  const lookups = labels.map(normalizeLookup)
  return lines.findIndex((line, index) => index >= startIndex && lookups.some((label) => normalizeLookup(line).startsWith(label)))
}

function cleanImportedFieldValue(lines: string[]) {
  return cleanImportedText(lines
    .map((line) => line.trim())
    .filter((line) => line && line !== '-' && !isImportedLabelOrHeader(line))
    .join('\n'))
}

function collectImportedField(lines: string[], startIndex: number, endIndex: number, labels: string[]) {
  if (startIndex === -1) return ''
  const values: string[] = []
  const firstLineValue = stripImportedLabel(lines[startIndex] || '', labels)
  if (firstLineValue && firstLineValue !== '-' && !isImportedLabelOrHeader(firstLineValue)) {
    values.push(firstLineValue)
  }

  for (let index = startIndex + 1; index < endIndex; index += 1) {
    const line = lines[index]
    if (!line || line === '-' || isImportedLabelOrHeader(line)) continue
    values.push(line)
  }

  return cleanImportedText(values.join('\n'))
}

function cleanImportedPurchaseLink(value: string) {
  return cleanImportedText(value
    .split('\n')
    .map((line) => /h\s*t\s*t\s*p/i.test(line)
      ? line
        .replace(/\s+/g, '')
        .replace(/h\s*t\s*t\s*p\s*s\s*:\s*\/\s*\//i, 'https://')
        .replace(/h\s*t\s*t\s*p\s*:\s*\/\s*\//i, 'http://')
      : line)
    .join('\n'))
}

function textBetweenImportedMarkers(text: string, start: RegExp, end: RegExp) {
  const clean = cleanImportedText(text)
  const startMatch = start.exec(clean)
  if (!startMatch) return ''
  const afterStart = startMatch.index + startMatch[0].length
  const tail = clean.slice(afterStart)
  const endMatch = end.exec(tail)
  return cleanImportedText(tail.slice(0, endMatch ? endMatch.index : undefined))
}

function cleanImportedSection(text: string) {
  return cleanImportedText(text)
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => {
      if (!line) return false
      if (/^EXHIBIT [AB]:/i.test(line)) return false
      if (/^Click to upload/i.test(line)) return false
      if (/^Analysis (?:—|-|Continued)/i.test(line)) return false
      if (/^Exhibit \d+ .*Photographic Evidence/i.test(line)) return false
      if (/^Postural Observations(?: Continued)?:?/i.test(line)) return false
      if (/^Continuation\s*\d*$/i.test(line)) return false
      if (/^[ivx]+\.\s+/i.test(line)) return false
      if (/^\d+\.\s*[A-Z]/.test(line)) return false
      if (/^JUSTIFICATION DISPLAY$/i.test(line)) return false
      if (/^Inquiries & Correspondence$/i.test(line)) return false
      if (/^Should you have any questions/i.test(line)) return false
      if (/^contact our administrative team/i.test(line)) return false
      if (/^copied on all emails/i.test(line)) return false
      if (/^This report was prepared/i.test(line)) return false
      if (/^IV\. EXHIBIT \d+ ANALYSIS/i.test(line)) return false
      if (/^IV\. EVIDENTIARY DOCUMENTATION/i.test(line)) return false
      if (/^V\. FINDINGS/i.test(line)) return false
      if (/^VI\. ERGONOMIC/i.test(line)) return false
      if (/^VII\. OPINION/i.test(line)) return false
      if (/^II\. ASSESSMENT/i.test(line)) return false
      if (/^III\. COMPREHENSIVE/i.test(line)) return false
      if (/^ERGONOMIC EVALUATION REPORT$/i.test(line)) return false
      if (/^Professional Ergonomic Assessment$/i.test(line)) return false
      return true
    })
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function dedupeRepeatedSentences(text: string) {
  const normalized = cleanImportedText(text)
  const parts = normalized.match(/[^.!?]+[.!?]+["']?|[^.!?]+$/g) || []
  const seen = new Set<string>()
  const kept: string[] = []

  parts.forEach((part) => {
    const clean = cleanImportedText(part)
    if (!clean) return
    const key = normalizeLookup(clean)
    if (key.length > 80 && seen.has(key)) return
    seen.add(key)
    kept.push(clean)
  })

  return cleanImportedText(kept.join(' '))
}

function sentenceParts(text: string) {
  return cleanImportedText(text)
    .replace(/\n+/g, ' ')
    .match(/[^.!?]+[.!?]+["”']?|[^.!?]+$/g) || []
}

const importedNarrativeBleedPatterns = [
  /\bThe primary risk factor identified during this assessment\b/i,
  /\bStephanie'?s job duties are highly repetitive\b/i,
  /\bApproximately 15-20% of these tasks\b/i,
  /\bDirect observation and interview reveal\b/i,
  /\bThe workstation is comprised of two free-standing\b/i,
  /\bThis "vulture neck" posture\b/i,
]

function trimImportedNarrativeBleed(text: string, contaminationSource = '') {
  const clean = cleanImportedText(text)
  const source = cleanImportedText(contaminationSource)
  const activePatterns = source
    ? importedNarrativeBleedPatterns.filter((pattern) => pattern.test(source))
    : importedNarrativeBleedPatterns
  const firstBleedIndex = activePatterns.reduce((lowest, pattern) => {
    const match = pattern.exec(clean)
    if (!match || match.index <= 0) return lowest
    return Math.min(lowest, match.index)
  }, clean.length)
  return cleanImportedText(clean.slice(0, firstBleedIndex))
}

function removeImportedSectionOverlap(text: string, contaminationSource: string) {
  const clean = trimImportedNarrativeBleed(text, contaminationSource)
  const contamination = cleanImportedText(contaminationSource)
  if (!clean || !contamination) return clean

  const contaminationKey = normalizeLookup(contamination)
  const contaminationSentenceKeys = new Set(
    sentenceParts(contamination)
      .map((sentence) => normalizeLookup(sentence))
      .filter((key) => key.length >= 50),
  )

  const cleanedLines = clean
    .split('\n')
    .map((line) => {
      const trimmed = line.trim()
      if (!trimmed) return ''
      const lineKey = normalizeLookup(trimmed)
      if (lineKey.length >= 90 && contaminationKey.includes(lineKey)) return ''

      const isBullet = /^-\s*/.test(trimmed)
      const keptSentences = sentenceParts(trimmed)
        .filter((sentence) => {
          const key = normalizeLookup(sentence)
          if (key.length < 50) return true
          return !contaminationSentenceKeys.has(key) && !contaminationKey.includes(key)
        })
      const joined = cleanImportedText(keptSentences.join(' '))
      if (!joined) return ''
      return isBullet && !joined.startsWith('-') ? `- ${joined}` : joined
    })
    .filter(Boolean)

  return cleanImportedText(cleanedLines.join('\n'))
}

function sanitizeReportSectionBleed(report: Report) {
  const posturalObservations = cleanImportedText(report.posturalObservations)
  if (!posturalObservations || posturalObservations === 'Enter detailed postural observations here...') return report

  report.exhibits = report.exhibits.map((exhibit) => ({
    ...exhibit,
    analysis: dedupeRepeatedSentences(removeImportedSectionOverlap(exhibit.analysis, posturalObservations)),
  }))
  report.procurementSummary = formatProcurementSummary(removeImportedSectionOverlap(report.procurementSummary, posturalObservations))
  report.evaluatorSummary = removeImportedSectionOverlap(report.evaluatorSummary, posturalObservations)
  return report
}

function formatProcurementSummary(text: string) {
  const clean = cleanImportedText(text)
    .replace(/\s+-\s+(?=[A-Z0-9])/g, '\n- ')
  if (!clean) return ''
  if (!/recommended items|ergonomic configuration/i.test(clean)) return clean

  const introMatch = clean.match(/^([\s\S]*?(?:configuration|recommended items)\s*:)\s*([\s\S]*)$/i)
  if (!introMatch) return clean

  let rest = introMatch[2]
  const bleedIndex = importedNarrativeBleedPatterns.reduce((lowest, pattern) => {
    const match = pattern.exec(rest)
    if (!match || match.index <= 0) return lowest
    return Math.min(lowest, match.index)
  }, rest.length)
  rest = rest.slice(0, bleedIndex).trim()

  const closingMatch = rest.match(/(?:\n|\s)+(These items can\b[\s\S]*)$/i)
  const closing = closingMatch
    ? cleanImportedText((closingMatch[1].match(/^These items can\b[\s\S]*?(?:[.!?](?=\s|$)|$)/i)?.[0] || closingMatch[1]))
    : ''
  if (closingMatch) rest = rest.slice(0, closingMatch.index).trim()

  const itemPatterns: RegExp[] = [
    /Two\s+48["“”]?\s*[-–]\s*55["“”]?\s+sit-stand desks/i,
    /Full back,\s*head and neck support ergonomic chair/i,
    /Adjustable height(?:\s*\(max\s+21["“”]?\))?\s*telescopic\s*laptop riser(?:\s*\(max\s+21["“”]?\s*height\))?/i,
    /Contoured\s+split\s+keyboard(?:\s+with\s+built-in\s+wrist\s+rest)?/i,
    /Dual monitor arms system/i,
    /Document holder/i,
    /Under-Desk Keyboard tray/i,
  ]
  const items = itemPatterns
    .map((pattern) => {
      const match = pattern.exec(rest)
      return match ? { index: match.index, value: cleanImportedText(match[0]) } : null
    })
    .filter((item): item is { index: number; value: string } => Boolean(item))
    .sort((a, b) => a.index - b.index)
    .filter((item, index, allItems) => allItems.findIndex((candidate) => normalizeLookup(candidate.value) === normalizeLookup(item.value)) === index)
    .map((item) => item.value)

  if (!items.length) return clean
  return [
    introMatch[1].trim(),
    ...items.map((item) => `- ${item}`),
    closing,
  ].filter(Boolean).join('\n')
}

function filenameEmployeeName(name: string) {
  const stem = name.replace(/\.[^.]+$/, '')
  const candidate = stem
    .replace(/[_-]\d{1,2}[._-]\d{1,2}[._-]\d{2,4}.*/, '')
    .replace(/[_-]+/g, ' ')
    .trim()
  return /^[a-z ,.'-]{3,}$/i.test(candidate) ? candidate : ''
}

function parseInlineValue(text: string, label: string, stopLabels: string[]) {
  const escapedLabel = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const stops = stopLabels.map((stopLabel) => stopLabel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')
  const pattern = new RegExp(`${escapedLabel}\\s*:?\\s*([\\s\\S]*?)(?=\\s+(?:${stops})(?:\\s*:|\\b)|\\n|\\s+[IVX]+\\.\\s+[A-Z]|\\s+[A-Z]\\.\\s+[A-Z]|$)`, 'i')
  return trimInlineImportedValue(pattern.exec(text)?.[1] || '')
}


function flexibleLabel(label: string) {
  return label
    .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    .replace(/\s+/g, '\\s+')
}

function valueFromImportedText(text: string, label: string, stopLabels: string[]) {
  const oneLine = cleanImportedText(text).replace(/\n+/g, ' ')
  const stops = stopLabels.map(flexibleLabel).join('|')
  const sectionStops = String.raw`[IVX]+\.\s+[A-Z][A-Za-z0-9 &/()—-]+|[A-Z]\.\s+[A-Z][A-Za-z0-9 &/()—-]+`
  const pattern = new RegExp(`${flexibleLabel(label)}\\s*:?\\s*([\\s\\S]*?)(?=\\s+(?:${stops})(?:\\s*:|\\b)|\\s+(?:${sectionStops})|$)`, 'i')
  return trimInlineImportedValue(pattern.exec(oneLine)?.[1] || '')
}


function parseRiskLevel(text: string, start: string, end: string) {
  const startIndex = text.indexOf(start)
  if (startIndex === -1) return null
  const endIndex = end ? text.indexOf(end, startIndex + start.length) : -1
  const block = text.slice(startIndex, endIndex === -1 ? undefined : endIndex)
  const match = block.match(/\b([LMH])\b(?![\s\S]*\b[LMH]\b)/)
  return match ? match[1] as Risk : null
}

function mergeUnique(values: string[]) {
  return values.filter((value, index, items) => value && items.indexOf(value) === index)
}

function parseImportedExhibit(text: string, exhibitNumber: number, posturalObservations = '') {
  const clean = cleanImportedText(text)
  const start = new RegExp(`(?:^|\\n)\\s*Analysis\\s+[—-]\\s+Exhibit\\s+${exhibitNumber}\\s*:`, 'i')
  const end = exhibitNumber < exhibitCount
    ? new RegExp(`(?:^|\\n)\\s*Analysis\\s+[—-]\\s+Exhibit\\s+${exhibitNumber + 1}\\s*:`, 'i')
    : /V\.\s+FINDINGS\s*&\s*CORRECTIVE\s+ACTION\s+PLAN/i
  return dedupeRepeatedSentences(removeImportedSectionOverlap(
    cleanImportedSection(textBetweenImportedMarkers(clean, start, end)),
    posturalObservations,
  ))
}

function exactHeading(label: string) {
  return new RegExp(`(?:^|\\n)\\s*${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*(?:\\n|$)`, 'i')
}

function cleanEvaluatorSummary(text: string, posturalObservations = '') {
  const clean = removeImportedSectionOverlap(cleanImportedSection(text), posturalObservations)
  const bulletStart = clean.search(/(?:^|\n)\s*-/)
  if (bulletStart !== -1) {
    const bulletText = clean.slice(bulletStart)
    const bullets = bulletText.match(/-\s*[\s\S]*?(?=(?:\s+-\s*[A-Z])|$)/g)
    if (bullets?.length) return cleanImportedText(bullets.slice(0, 6).join('\n'))
    return cleanImportedText(bulletText)
  }
  const inquiryStart = clean.search(/(?:^|\n)\s*Inquiries & Correspondence/i)
  return cleanImportedText(clean.slice(0, inquiryStart === -1 ? undefined : inquiryStart))
}

export function parseImportedRecommendations(text: string) {
  const section = textBetweenImportedMarkers(
    text,
    /VI\.\s+ERGONOMIC\s+EQUIPMENT\s+RECOMMENDATION/i,
    /VI\.\s+ERGONOMIC\s+EQUIPMENT\s+RECOMMENDATION\s+[—-]\s+PROCUREMENT|VI\.\s+PROCUREMENT\s+NOTES|VII\.\s+OPINION/i,
  )
  if (!section) return []

  const lines = importedLines(section)
  const blocks: string[][] = []

  // Primary: each card's first field is "Model / SKU" — the line before it is the category+priority
  const modelIndices = lines.reduce<number[]>((acc, line, i) => {
    if (/^(?:Model\s*\/\s*SKU|Product\s*\/\s*model)\s*$/i.test(line)) acc.push(i)
    return acc
  }, [])

  if (modelIndices.length) {
    modelIndices.forEach((modelIndex, i) => {
      const start = modelIndex > 0 ? modelIndex - 1 : modelIndex
      const end = i + 1 < modelIndices.length ? modelIndices[i + 1] - 1 : lines.length
      blocks.push(lines.slice(start, end))
    })
  } else {
    // Legacy: split on standalone "Equipment Recommendation" label (empty category)
    let currentBlock: string[] = []
    lines.forEach((line) => {
      if (/^Equipment Recommendation\s*$/i.test(line)) {
        if (currentBlock.length) blocks.push(currentBlock)
        currentBlock = []
        return
      }
      currentBlock.push(line)
    })
    if (currentBlock.length) blocks.push(currentBlock)
  }

  const sourceBlocks = blocks.length ? blocks : section.split(/Equipment Recommendation/i).slice(1).map(importedLines)

  return sourceBlocks
    .map((blockLines) => {
      const modelIndex = findImportedLabelIndex(blockLines, ['Model / SKU', 'Product / model'])
      const priorityIndex = findImportedLabelIndex(blockLines, ['Clinical Priority'])
      const actionIndex = findImportedLabelIndex(blockLines, ['Recommended Action / Risk Mitigation', 'Recommended action'])
      const linkIndex = findImportedLabelIndex(blockLines, ['Purchase Link', 'Open Purchase Link'])

      const firstFieldIndex = [modelIndex, priorityIndex, actionIndex, linkIndex]
        .filter((index) => index >= 0)
        .sort((a, b) => a - b)[0] ?? blockLines.length

      const category = cleanImportedFieldValue(blockLines.slice(0, firstFieldIndex))
      const currentEnd = [priorityIndex, actionIndex, linkIndex, blockLines.length]
        .filter((index) => index > modelIndex)
        .sort((a, b) => a - b)[0] ?? blockLines.length
      const priorityEnd = [actionIndex, linkIndex, blockLines.length]
        .filter((index) => index > priorityIndex)
        .sort((a, b) => a - b)[0] ?? blockLines.length
      const actionEnd = [linkIndex, blockLines.length]
        .filter((index) => index > actionIndex)
        .sort((a, b) => a - b)[0] ?? blockLines.length

      const priorityValue = priorityIndex >= 0
        ? collectImportedField(blockLines, priorityIndex, priorityEnd, ['Clinical Priority'])
        : ''
      const detectedPriority = priorityValue.match(/\b(High|Medium|Low)\b/i)?.[1] ||
        blockLines.find((line) => /^(High|Medium|Low)$/i.test(line.trim())) ||
        category.match(/\b(High|Medium|Low)\s*$/i)?.[1] ||
        ''

      const row: EquipmentRow = {
        category,
        current: modelIndex >= 0 ? collectImportedField(blockLines, modelIndex, currentEnd, ['Model / SKU', 'Product / model']) : '',
        priority: detectedPriority.replace(/^./, (letter) => letter.toUpperCase()) as Priority,
        action: actionIndex >= 0 ? collectImportedField(blockLines, actionIndex, actionEnd, ['Recommended Action / Risk Mitigation', 'Recommended action']) : '',
        notes: linkIndex >= 0 ? cleanImportedPurchaseLink(collectImportedField(blockLines, linkIndex, blockLines.length, ['Purchase Link', 'Open Purchase Link'])) : '',
      }

      if (!['High', 'Medium', 'Low'].includes(row.priority)) row.priority = ''
      return repairEquipmentRow(row)
    })
    .filter((row) => [row.category, row.current, row.action, row.notes].some((value) => value.trim()))
}


export function autofillReportFromSource(current: Report, sourceDocument: ImportedSourceDocument, warning?: string) {
  const text = cleanImportedText(sourceDocument.text || '')
  const lines = importedLines(text)
  const next: Report = normalizeReport({
    ...current,
    sourceDocument,
    importedSourceText: sourceDocument.text || warning || '',
  })
  let filledCount = 0

  const assign = <K extends keyof Report>(key: K, value: Report[K] | '') => {
    if (typeof value !== 'string' || !value.trim()) return
    if (next[key] === value) return
    next[key] = value as Report[K]
    filledCount += 1
  }

  const parsedCaseType = caseTypes.find((caseType) => text.toLowerCase().includes(caseType.toLowerCase()))
  if (parsedCaseType) {
    next.caseType = parsedCaseType
    filledCount += 1
  }

  assign('claimNumber', valueFromImportedText(text, 'Case / Claim #', ['Employee Name', 'Employee ID']) || valueAfterImportedLabel(lines, ['Case / Claim #', 'Case Claim #']))
  assign('employeeName', valueFromImportedText(text, 'Employee Name', ['Employee ID', 'Department']) || valueAfterImportedLabel(lines, ['Employee Name']) || filenameEmployeeName(sourceDocument.name))
  assign('employeeId', valueFromImportedText(text, 'Employee ID', ['Department', 'Location / Floor']) || valueAfterImportedLabel(lines, ['Employee ID']))
  assign('department', valueFromImportedText(text, 'Department', ['Location / Floor', 'Date of Eval']) || valueAfterImportedLabel(lines, ['Department']))
  assign('location', valueFromImportedText(text, 'Location / Floor', ['Date of Eval', 'Report Date']) || valueAfterImportedLabel(lines, ['Location / Floor', 'Location Floor']))
  assign('dateOfEval', valueFromImportedText(text, 'Date of Eval', ['Report Date', 'Manager Name']) || valueAfterImportedLabel(lines, ['Date of Eval']))
  assign('reportDate', valueFromImportedText(text, 'Report Date', ['Manager Name', 'Job Title']) || valueAfterImportedLabel(lines, ['Report Date']))
  assign('managerName', valueFromImportedText(text, 'Manager Name', ['Job Title', 'I.', 'II.', 'Assessment Methodology']) || valueAfterImportedLabel(lines, ['Manager Name']))
  assign('jobTitle', valueFromImportedText(text, 'Job Title', ['I.', 'II.', 'Assessment Mode', 'Assessment Methodology']) || valueAfterImportedLabel(lines, ['Job Title']))
  assign('stature', valueFromImportedText(text, 'Stature', ['Dominant Hand', 'Vision', 'Desk Type']) || parseInlineValue(text, 'Stature', ['Dominant Hand', 'Vision', 'Desk Type']))
  assign('currentDeskHeight', valueFromImportedText(text, 'Current Desk Hgt', ['Ideal Keying Hgt', 'Delta']) || parseInlineValue(text, 'Current Desk Hgt', ['Ideal Keying Hgt', 'Delta']))
  assign('idealKeyingHeight', valueFromImportedText(text, 'Ideal Keying Hgt', ['Delta', 'B. Postural Deviation Matrix']) || parseInlineValue(text, 'Ideal Keying Hgt', ['Delta']))
  assign('delta', valueFromImportedText(text, 'Delta', ['B. Postural Deviation Matrix', 'Postural Deviation Matrix', 'Risk Level', 'Body Segment']) || parseInlineValue(text, 'Delta', ['B. Postural Deviation Matrix', 'Postural Deviation Matrix', 'Risk Level', 'Body Segment']))

  // In visually-extracted PDFs the signed measurement (e.g. "+8"") can appear above the
  // label row, leaving only the descriptive caption after "Delta:" in the one-line parse.
  // If the imported value contains no digit, recover the number from the nearby region.
  if (!/\d/.test(next.delta || '')) {
    const anthropRegion = textBetweenImportedMarkers(
      text,
      /Desk\s+Type\b/i,
      /B\.\s+Postural\s+Deviation\s+Matrix/i,
    )
    const numericDelta = /([+-]\d+(?:\.\d+)?[^\s,\n(]?)/.exec(anthropRegion)?.[1]
    if (numericDelta) next.delta = numericDelta
  }

  next.dataSources = mergeUnique(dataSources.filter((value) => text.includes(value)))
  next.assessmentTools = mergeUnique(assessmentTools.filter((value) => text.includes(value)))
  next.benchmarks = mergeUnique(benchmarks.filter((value) => text.includes(value) || normalizeLookup(text).includes(normalizeLookup(value))))
  next.controls = mergeUnique(controls.filter((value) => text.includes(value)))

  const riskMap: Array<[string, string, string]> = [
    ['neck', 'Neck / Head', 'Shoulders'],
    ['shoulders', 'Shoulders', 'Wrists / Arms'],
    ['wrists', 'Wrists / Arms', 'Back / Trunk'],
    ['back', 'Back / Trunk', 'Postural Observations'],
  ]
  riskMap.forEach(([key, start, end]) => {
    const risk = parseRiskLevel(text, start, end)
    if (risk) next.postures[key] = { ...next.postures[key], risk }
  })

  const posturalObservations = textBetweenImportedMarkers(
    text,
    /Postural Observations(?: Continued)?:/i,
    /IV\.\s+EVIDENTIARY\s+DOCUMENTATION/i,
  )
  // cleanImportedSection filters /^\d+\.\s*[A-Z]/ which strips numbered observation lists,
  // so use a lighter clean that only removes section navigation artifacts
  const importedPosturalObservations = dedupeRepeatedSentences(
    cleanImportedText(posturalObservations)
      .split('\n')
      .map((line) => line.trim())
      .filter((line) => {
        if (!line) return false
        if (/^[ivx]+\.\s+/i.test(line)) return false
        if (/^Postural Observations(?: Continued)?:?\s*$/i.test(line)) return false
        if (/^ERGONOMIC EVALUATION REPORT$/i.test(line)) return false
        if (/^Professional Ergonomic Assessment$/i.test(line)) return false
        return true
      })
      .join('\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim(),
  )
  assign('posturalObservations', importedPosturalObservations)

  next.exhibits = next.exhibits.map((exhibit, index) => {
    const analysis = parseImportedExhibit(text, index + 1, importedPosturalObservations || next.posturalObservations)
    const sourcePhotoPatch = Object.fromEntries(
      (sourceDocument.exhibitPhotos || [])
        .filter((item) => item.exhibitIndex === index)
        .map((item) => [item.side, item.photo]),
    ) as Partial<Exhibit>
    if (!analysis && !Object.keys(sourcePhotoPatch).length) return exhibit
    filledCount += 1
    return { ...exhibit, ...sourcePhotoPatch, ...(analysis ? { analysis } : {}) }
  })

  if (sourceDocument.signaturePhoto && !next.signature) {
    next.signature = sourceDocument.signaturePhoto
    filledCount += 1
  }

  const recommendations = parseImportedRecommendations(text)
  if (recommendations.length) {
    next.recommendations = next.recommendations.map((row, index) => recommendations[index] || row)
    filledCount += recommendations.length
  }

  assign('justificationDisplay', cleanImportedSection(textBetweenImportedMarkers(
    text,
    exactHeading('Editable Clinical Justification Display'),
    /VI\.\s+ERGONOMIC\s+EQUIPMENT\s+RECOMMENDATION/i,
  )))
  if (next.justificationDisplay) {
    next.justifications = justifications.filter((title) => next.justificationDisplay.includes(`${title}:`))
    if (next.justifications.length) filledCount += 1
  }
  assign('procurementSummary', formatProcurementSummary(removeImportedSectionOverlap(cleanImportedSection(textBetweenImportedMarkers(
    text,
    /Procurement Notes\s*&\s*Justification Summary/i,
    /VII\.\s+OPINION\s*&\s*PROFESSIONAL\s+CERTIFICATION/i,
  )), importedPosturalObservations || next.posturalObservations)))
  const mainEvalSummary = textBetweenImportedMarkers(
    text,
    exactHeading('Evaluator Summary'),
    /Evaluator Name/i,
  )
  const contEvalSummary = textBetweenImportedMarkers(
    text,
    /(?:^|\n)\s*Evaluator Summary Continued\s*(?:\n|$)/i,
    /Inquiries\s*&\s*Correspondence|This\s+report\s+was\s+prepared\b/i,
  )
  assign('evaluatorSummary', cleanEvaluatorSummary(
    [mainEvalSummary, contEvalSummary].filter(Boolean).join('\n\n'),
    importedPosturalObservations || next.posturalObservations,
  ))
  assign('evaluatorName', valueAfterImportedLabel(lines, ['Evaluator Name', 'Evaluator Name (Printed)']))

  return { report: normalizeReport(sanitizeReportSectionBleed(next)), filledCount }
}

const field = (value: string) => <span className="pdf-field">{value}</span>
const mark = (on: boolean) => <span className={`box ${on ? 'checked' : ''}`} />

function waitForPaint() {
  return new Promise<void>((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  })
}

function App() {
  const [report, setReport] = useState<Report>(() => normalizeReport(newReport()))
  const [savedReports, setSavedReports] = useState<ReportSummary[]>([])
  const [status, setStatus] = useState('Ready')

  const api = window.reportApi

  useEffect(() => {
    void refreshReports()
  }, [])

  async function refreshReports() {
    if (!api) return
    setSavedReports(await api.listReports())
  }

  function update<K extends keyof Report>(key: K, value: Report[K]) {
    setReport((current) => ({ ...current, [key]: value }))
  }

  function toggleList(key: keyof Report, value: string) {
    setReport((current) => {
      const list = current[key] as string[]
      const next = list.includes(value) ? list.filter((item) => item !== value) : [...list, value]
      return { ...current, [key]: next }
    })
  }

  function toggleJustification(value: string) {
    setReport((current) => {
      const selected = current.justifications.includes(value)
      return {
        ...current,
        justifications: selected
          ? current.justifications.filter((item) => item !== value)
          : [...current.justifications, value],
        justificationDisplay: selected
          ? removeJustificationBlock(current.justificationDisplay, value)
          : addJustificationBlock(current.justificationDisplay, value),
      }
    })
  }

  function updateRecommendation(index: number, key: keyof EquipmentRow, value: string) {
    setReport((current) => ({
      ...current,
      recommendations: current.recommendations.map((row, i) => (i === index ? { ...row, [key]: value } : row)),
    }))
  }

  function updateExhibit(index: number, patch: Partial<Exhibit>) {
    setReport((current) => ({
      ...current,
      exhibits: current.exhibits.map((row, i) => (i === index ? { ...row, ...patch } : row)),
    }))
  }

  async function saveFile(file: File) {
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(String(reader.result))
      reader.onerror = () => reject(reader.error)
      reader.readAsDataURL(file)
    })
    return api ? api.savePhoto({ name: file.name, dataUrl }) : { name: file.name, path: '', url: dataUrl }
  }

  async function attachPhoto(event: ChangeEvent<HTMLInputElement>, exhibitIndex: number, side: PhotoSide) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    const photo = await saveFile(file)
    updateExhibit(exhibitIndex, { [side]: photo })
    setStatus(`Attached ${file.name}`)
  }

  function deletePhoto(exhibitIndex: number, side: PhotoSide) {
    updateExhibit(exhibitIndex, side === 'pre' ? { pre: undefined } : { post: undefined })
    setStatus('Photo removed')
  }

  async function attachSignature(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    event.target.value = ''
    if (!file) return
    const signature = await saveFile(file)
    setReport((current) => ({
      ...current,
      signature,
      signatureConfirmed: true,
      signatureSignedAt: new Date().toISOString(),
    }))
    setStatus('Electronic signature attached')
  }

  async function save() {
    if (!api) return null
    const saved = await api.saveReport(report)
    setReport(normalizeReport(saved))
    await refreshReports()
    setStatus('Saved to SQLite')
    return saved
  }

  async function load(id: string) {
    if (!api || !id) return
    const loaded = await api.loadReport(id)
    if (loaded) {
      setReport(normalizeReport(loaded))
      setStatus('Loaded saved report')
    }
  }

  async function remove() {
    if (!api) return
    await api.deleteReport(report.id)
    setReport(newReport())
    await refreshReports()
    setStatus('Deleted report')
  }

  async function importPdf() {
    if (!api) return
    setStatus('Importing file and reading report data...')
    await waitForPaint()
    const result = await api.importPdf()
    if (!result) {
      setStatus('PDF import canceled')
      return
    }
    if (result.error || !result.report) {
      if (result.sourceDocument) {
        const { report: nextReport, filledCount } = autofillReportFromSource(report, result.sourceDocument, result.warning)
        const saved = await api.saveReport(nextReport)
        setReport(normalizeReport(saved))
        await refreshReports()
        setStatus(result.warning
          ? `Imported file with warning: ${result.warning}`
          : `Imported file and filled ${filledCount} report fields: ${result.sourceDocument.name}`)
        return
      }
      setStatus(`File import failed: ${result.error || 'No editable data or readable text found.'}`)
      return
    }
    const imported = normalizeReport(sanitizeReportSectionBleed(normalizeReport(result.report)))
    const saved = await api.saveReport(imported)
    setReport(normalizeReport(saved))
    await refreshReports()
    setStatus(`Imported editable PDF: ${result.filePath}`)
  }

  async function exportPdf() {
    setStatus('Preparing PDF layout...')
    await waitForPaint()
    const saved = await save()
    const reportForExport = normalizeReport(saved || report)
    setReport(reportForExport)
    await waitForPaint()
    const path = await api?.exportPdf(reportForExport.employeeName || 'Ergonomic Report', reportForExport)
    setStatus(path ? `Exported PDF: ${path}` : 'PDF export canceled')
  }

  async function exportDiagnostics() {
    const path = await api?.exportDiagnostics()
    setStatus(path ? `Exported diagnostics: ${path}` : 'Diagnostics export canceled')
  }

  const title = useMemo(() => report.employeeName || 'New ergonomic report', [report.employeeName])

  return (
    <main className="workspace">
      <aside className="sidebar">
        <div className="brand-block">
          <img src={logo} alt="Ergonomic Works" />
          <div>
            <h1>Ergonomic Evaluation</h1>
            <p>{title}</p>
          </div>
        </div>

        <div className="toolbar">
          <button onClick={() => setReport(newReport())}><FilePlus2 /> New</button>
          <button onClick={save}><Save /> Save</button>
          <button onClick={exportPdf}><Download /> Export PDF</button>
          <button onClick={exportDiagnostics}><Bug /> Diagnostics</button>
          <button className="danger" onClick={remove}><Trash2 /> Delete</button>
        </div>

        <label className="stacked">
          <span><FolderOpen size={15} /> Saved reports</span>
          <select onChange={(event) => load(event.target.value)} value="">
            <option value="">Choose a saved report...</option>
            {savedReports.map((saved) => (
              <option key={saved.id} value={saved.id}>{saved.title} - {new Date(saved.updatedAt).toLocaleString()}</option>
            ))}
          </select>
        </label>

        <div className="stacked">
          <span><FileUp size={15} /> Import file</span>
          <button onClick={importPdf}>Choose PDF, Word, or Excel</button>
        </div>

        <Editor
          report={report}
          update={update}
          toggleList={toggleList}
          toggleJustification={toggleJustification}
          updateRecommendation={updateRecommendation}
          updateExhibit={updateExhibit}
          attachPhoto={attachPhoto}
          deletePhoto={deletePhoto}
          attachSignature={attachSignature}
        />
        <p className="status">{status}</p>
      </aside>

      <section className="preview-panel">
        <ReportDocument report={report} onPhotoUpload={attachPhoto} />
      </section>
    </main>
  )
}

type EditorProps = {
  report: Report
  update: <K extends keyof Report>(key: K, value: Report[K]) => void
  toggleList: (key: keyof Report, value: string) => void
  toggleJustification: (value: string) => void
  updateRecommendation: (index: number, key: keyof EquipmentRow, value: string) => void
  updateExhibit: (index: number, patch: Partial<Exhibit>) => void
  attachPhoto: (event: ChangeEvent<HTMLInputElement>, exhibitIndex: number, side: PhotoSide) => void
  deletePhoto: (exhibitIndex: number, side: PhotoSide) => void
  attachSignature: (event: ChangeEvent<HTMLInputElement>) => void
}

export function Editor({ report, update, toggleList, toggleJustification, updateRecommendation, updateExhibit, attachPhoto, deletePhoto, attachSignature }: EditorProps) {
  const adminFields: [keyof Report, string][] = [
    ['claimNumber', 'Case / Claim #'],
    ['employeeName', 'Employee name'],
    ['employeeId', 'Employee ID'],
    ['department', 'Department'],
    ['location', 'Location / Floor'],
    ['dateOfEval', 'Date of Eval'],
    ['reportDate', 'Report date'],
    ['managerName', 'Manager name'],
    ['jobTitle', 'Job title'],
  ]

  return (
    <div className="editor">
      {(report.sourceDocument || report.importedSourceText) && (
        <section>
          <h2>Imported Source File</h2>
          {report.sourceDocument && (
            <div className="import-summary">
              <b>{report.sourceDocument.name}</b>
              <span>{report.sourceDocument.extractionStatus}</span>
              <span>Imported {new Date(report.sourceDocument.importedAt).toLocaleString()}</span>
            </div>
          )}
          <label>
            <span>Extracted source text</span>
            <textarea
              className="tall-editor source-text-editor"
              value={report.importedSourceText}
              onChange={(event) => update('importedSourceText', event.target.value)}
            />
          </label>
        </section>
      )}

      <section>
        <h2>Administrative Data</h2>
        <label className="full-field">
          <span>Case type</span>
          <select value={report.caseType} onChange={(event) => update('caseType', event.target.value as CaseType)}>
            {caseTypes.map((type) => <option key={type}>{type}</option>)}
          </select>
        </label>
        <div className="field-grid">
          {adminFields.map(([key, label]) => (
            <label key={key}>
              <span>{label}</span>
              <input value={String(report[key])} onChange={(event) => update(key, event.target.value as never)} />
            </label>
          ))}
        </div>
      </section>

      <section>
        <h2>Assessment</h2>
        <div className="inline-options">
          {['In-Person', 'Remote/Virtual'].map((option) => (
            <label key={option}><input type="radio" checked={report.assessmentMode === option} onChange={() => update('assessmentMode', option)} /> {option}</label>
          ))}
        </div>
        <Checklist title="Data sources" values={dataSources} selected={report.dataSources} onToggle={(value) => toggleList('dataSources', value)} />
        <Checklist title="Assessment tools" values={assessmentTools} selected={report.assessmentTools} onToggle={(value) => toggleList('assessmentTools', value)} />
        <Checklist title="Benchmarks" values={benchmarks} selected={report.benchmarks} onToggle={(value) => toggleList('benchmarks', value)} />
      </section>

      <section>
        <h2>Workstation Analysis</h2>
        <div className="field-grid compact">
          {(['stature', 'currentDeskHeight', 'idealKeyingHeight', 'delta'] as const).map((key) => (
            <label key={key}><span>{labelFor(key)}</span><input value={report[key]} onChange={(event) => update(key, event.target.value)} /></label>
          ))}
        </div>
        <div className="inline-options">
          {(['R', 'L'] as const).map((hand) => <label key={hand}><input type="radio" checked={report.dominantHand === hand} onChange={() => update('dominantHand', hand)} /> {hand}</label>)}
          {['Single', 'Bifocal', 'Progressive'].map((vision) => <label key={vision}><input type="radio" checked={report.vision === vision} onChange={() => update('vision', vision)} /> {vision}</label>)}
        </div>
        <div className="inline-options">
          {['Fixed', 'Electric Sit/Stand', 'Desk Converter'].map((desk) => <label key={desk}><input type="radio" checked={report.deskType === desk} onChange={() => update('deskType', desk)} /> {desk}</label>)}
        </div>
        {postureRows.map((row) => (
          <div className="posture-edit" key={row.key}>
            <strong>{row.label}</strong>
            {row.factors.map((factor) => (
              <label key={factor}><input type="checkbox" checked={report.postures[row.key].factors.includes(factor)} onChange={() => {
                const factors = report.postures[row.key].factors.includes(factor)
                  ? report.postures[row.key].factors.filter((item) => item !== factor)
                  : [...report.postures[row.key].factors, factor]
                update('postures', { ...report.postures, [row.key]: { ...report.postures[row.key], factors } })
              }} /> {factor}</label>
            ))}
            <select value={report.postures[row.key].risk} onChange={(event) => update('postures', { ...report.postures, [row.key]: { ...report.postures[row.key], risk: event.target.value as Risk } })}>
              <option>L</option><option>M</option><option>H</option>
            </select>
          </div>
        ))}
        <label><span>Postural observations</span><textarea className="tall-editor" value={report.posturalObservations} onChange={(event) => update('posturalObservations', event.target.value)} /></label>
        <LongTextWarning text={report.posturalObservations} limit={posturalObservationsRecommendedLimit} />
      </section>

      <section>
        <h2>Photographs</h2>
        {report.exhibits.map((exhibit, index) => (
          <div className="photo-editor" key={index}>
            <strong>Exhibit {index + 1}</strong>
            <div className="photo-field">
              <label><Camera size={15} /> Pre photo<input type="file" accept="image/*" onChange={(event) => attachPhoto(event, index, 'pre')} /></label>
              {exhibit.pre && (
                <div className="photo-thumb-row">
                  <img className="photo-thumb" src={exhibit.pre.url} alt={exhibit.pre.name} />
                  <button type="button" className="photo-delete" title="Remove photo" onClick={() => deletePhoto(index, 'pre')}>×</button>
                </div>
              )}
            </div>
            <div className="photo-field">
              <label><Camera size={15} /> Corrective photo<input type="file" accept="image/*" onChange={(event) => attachPhoto(event, index, 'post')} /></label>
              {exhibit.post && (
                <div className="photo-thumb-row">
                  <img className="photo-thumb" src={exhibit.post.url} alt={exhibit.post.name} />
                  <button type="button" className="photo-delete" title="Remove photo" onClick={() => deletePhoto(index, 'post')}>×</button>
                </div>
              )}
            </div>
            <textarea placeholder="Analysis" value={exhibit.analysis} onChange={(event) => updateExhibit(index, { analysis: event.target.value })} />
            <LongTextWarning text={exhibit.analysis} limit={exhibitAnalysisRecommendedLimit} />
          </div>
        ))}
      </section>

      <section>
        <h2>Recommendations</h2>
        {report.recommendations.map((row, index) => (
          <div className="recommendation-edit" key={index}>
            <label><span>Equipment recommendation</span><input value={row.category} onChange={(event) => updateRecommendation(index, 'category', event.target.value)} /></label>
            <label><span>Model / SKU</span><input value={row.current} onChange={(event) => updateRecommendation(index, 'current', event.target.value)} /></label>
            <label><span>Clinical priority</span><select value={row.priority} onChange={(event) => updateRecommendation(index, 'priority', event.target.value)}><option value="">Select...</option><option>High</option><option>Medium</option><option>Low</option></select></label>
            <label className="full-field"><span>Recommended action / risk mitigation</span><textarea value={row.action} onChange={(event) => updateRecommendation(index, 'action', event.target.value)} /></label>
            <label className="full-field"><span>Purchase link</span><textarea value={row.notes} onChange={(event) => updateRecommendation(index, 'notes', event.target.value)} /></label>
          </div>
        ))}
      </section>

      <section>
        <h2>Controls, Justifications & Certification</h2>
        <Checklist title="Administrative controls" values={controls} selected={report.controls} onToggle={(value) => toggleList('controls', value)} />
        <Checklist title="Software" values={softwareOptions} selected={report.software} onToggle={(value) => toggleList('software', value)} />
        <Checklist title="Expert justification repository" values={justifications} selected={report.justifications} onToggle={toggleJustification} />
        <label><span>Editable justification display</span><textarea className="tall-editor" value={report.justificationDisplay} onChange={(event) => update('justificationDisplay', event.target.value)} /></label>
        <LongTextWarning text={report.justificationDisplay} limit={justificationRecommendedLimit} />
        <label><span>Procurement summary</span><textarea className="tall-editor" value={report.procurementSummary} onChange={(event) => update('procurementSummary', event.target.value)} /></label>
        <LongTextWarning text={report.procurementSummary} limit={procurementSummaryRecommendedLimit} />
        <label><span>Evaluator summary</span><textarea className="tall-editor" value={report.evaluatorSummary} onChange={(event) => update('evaluatorSummary', event.target.value)} /></label>
        <LongTextWarning text={report.evaluatorSummary} limit={evaluatorSummaryRecommendedLimit} />
        <div className="field-grid compact">
          <label><span>Evaluator name</span><input value={report.evaluatorName} onChange={(event) => update('evaluatorName', event.target.value)} /></label>
          <label><span>Date</span><input value={report.certificationDate} onChange={(event) => update('certificationDate', event.target.value)} /></label>
        </div>
        <div className="signature-editor">
          <label><PenLine size={15} /> Upload signature image<input type="file" accept="image/*" onChange={attachSignature} /></label>
          <label><input type="checkbox" checked={report.signatureConfirmed} onChange={(event) => update('signatureConfirmed', event.target.checked)} /> I certify this electronic signature for this report.</label>
          {report.signature && <span>Attached: {report.signature.name}</span>}
        </div>
      </section>
    </div>
  )
}

function Checklist({ title, values, selected, onToggle }: { title: string; values: string[]; selected: string[]; onToggle: (value: string) => void }) {
  return (
    <fieldset>
      <legend>{title}</legend>
      {values.map((value) => (
        <label key={value}><input type="checkbox" checked={selected.includes(value)} onChange={() => onToggle(value)} /> {value}</label>
      ))}
    </fieldset>
  )
}

function labelFor(value: string) {
  return value.replace(/([A-Z])/g, ' $1').replace(/^./, (letter) => letter.toUpperCase())
}

function LongTextWarning({ text, limit }: { text: string; limit: number }) {
  const length = cleanImportedText(text).length
  if (length <= limit) return null
  return (
    <p className="field-warning">
      This section is long and may create continuation pages. Recommended maximum: {limit.toLocaleString()} characters.
    </p>
  )
}

type SectionPages = Record<'I' | 'II' | 'III' | 'IV' | 'V' | 'VI' | 'VII', number>

type PlannedExhibit = {
  exhibit: Exhibit
  index: number
  chunks: string[]
}

type PhotoPlanItem =
  | { type: 'photo'; exhibits: PlannedExhibit[]; section: string }
  | { type: 'continuation'; exhibitNumber: number; chunks: string[] }

export function buildPhotoPlan(exhibits: Exhibit[]): PhotoPlanItem[] {
  const plan: PhotoPlanItem[] = []
  let isFirstPhotoPage = true

  const addPhotoPage = (index: number) => {
    const plannedExhibits = [{
      exhibit: exhibits[index],
      index,
      chunks: chunkText(exhibits[index].analysis, singleExhibitAnalysisLimit, {
        hardMaxChars: singleExhibitAnalysisHardLimit,
        nextMaxChars: exhibitAnalysisContinuationLimit,
        nextHardMaxChars: exhibitAnalysisContinuationHardLimit,
        minTailChars: minimumContinuationText,
      }),
    }]

    plan.push({
      type: 'photo',
      exhibits: plannedExhibits,
      section: isFirstPhotoPage
        ? 'IV.  EVIDENTIARY DOCUMENTATION  (PHOTOGRAPHS)'
        : 'IV.  EVIDENTIARY DOCUMENTATION  (PHOTOGRAPHS) — CONTINUATION',
    })
    isFirstPhotoPage = false

    plannedExhibits.forEach((item) => {
      const overflow = item.chunks.slice(1)
      if (overflow.length) {
        plan.push({ type: 'continuation', exhibitNumber: item.index + 1, chunks: overflow })
      }
    })
  }

  exhibits.forEach((_exhibit, index) => addPhotoPage(index))

  return plan
}

function ReportDocument({ report, onPhotoUpload }: { report: Report; onPhotoUpload?: (event: ChangeEvent<HTMLInputElement>, exhibitIndex: number, side: PhotoSide) => void }) {
  const normalized = normalizeReport(report)
  const observationChunks = chunkText(normalized.posturalObservations, observationTextLimit, {
    hardMaxChars: observationHardLimit,
    nextMaxChars: observationContinuationLimit,
    nextHardMaxChars: observationContinuationHardLimit,
    minTailChars: minimumContinuationText,
  })
  const justificationChunks = chunkText(normalized.justificationDisplay, justificationTextLimit, {
    hardMaxChars: justificationHardLimit,
    nextMaxChars: justificationContinuationLimit,
    nextHardMaxChars: justificationContinuationHardLimit,
    minTailChars: minimumContinuationText,
  })
  // Keep procurement as one formatted block so list formatting is not destroyed by chunkText().
  const procurementChunks = normalized.procurementSummary.trim()
    ? [normalized.procurementSummary]
    : ['']
  const evaluatorSummaryText = normalized.evaluatorSummary.trim()
  const evaluatorSummaryFitsCertification = evaluatorSummaryText.length <= evaluatorCertificationTextLimit
  const evaluatorSummaryChunks = evaluatorSummaryText && !evaluatorSummaryFitsCertification
    ? chunkText(evaluatorSummaryText, 3400, {
      hardMaxChars: 3900,
      nextMaxChars: 3400,
      nextHardMaxChars: 3900,
      minTailChars: minimumContinuationText,
    })
    : []
  const photoPlan = buildPhotoPlan(normalized.exhibits)
  const photoPageCount = photoPlan.reduce((count, item) => count + (item.type === 'photo' ? 1 : item.chunks.length), 0)
  const equipmentRows = visibleRecommendationRows(normalized.recommendations)
  const equipmentGroups = chunkEquipmentRows(equipmentRows)
  const equipmentPageCount = equipmentGroups.length
  const evaluatorPageCount = evaluatorSummaryChunks.length + 1

  const sectionPages: SectionPages = {
    I: 1,
    II: 2,
    III: 3,
    IV: 3 + observationChunks.length,
    V: 3 + observationChunks.length + photoPageCount,
    VI: 3 + observationChunks.length + photoPageCount + 1 + justificationChunks.length,
    VII: 3 + observationChunks.length + photoPageCount + 1 + justificationChunks.length + equipmentPageCount + procurementChunks.length,
  }
  const totalPages = sectionPages.VII + evaluatorPageCount - 1
  const pages: Array<(page: number) => React.ReactNode> = []

  pages.push((page) => (
    <Page page={page} total={totalPages} title="ERGONOMIC EVALUATION REPORT" section="REPORT NAVIGATION  —  TABLE OF CONTENTS">
      <TableOfContents sectionPages={sectionPages} />
      <SectionTitle text="I.  ADMINISTRATIVE & JURISDICTIONAL DATA" />
      <div className="case-line"><b>CASE TYPE</b>{field(normalized.caseType)}<b>Case / Claim #:</b>{field(normalized.claimNumber)}</div>
      <FormRows rows={[
        ['Employee Name:', normalized.employeeName],
        ['Employee ID:', normalized.employeeId],
        ['Department:', normalized.department],
        ['Location / Floor:', normalized.location],
        ['Date of Eval:', normalized.dateOfEval],
        ['Report Date:', normalized.reportDate],
        ['Manager Name:', normalized.managerName],
        ['Job Title:', normalized.jobTitle],
      ]} />
    </Page>
  ))

  pages.push((page) => (
    <Page page={page} total={totalPages} section="II.  ASSESSMENT METHODOLOGY & COMPLIANCE STANDARDS" className="methodology-page">
      <p className="italic">This evaluation follows recognized industry protocols to ensure objective findings and minimize bias.</p>
      <div className="pdf-options"><b>Assessment Mode:</b>{mark(normalized.assessmentMode === 'In-Person')}<b>In-Person</b>{mark(normalized.assessmentMode === 'Remote/Virtual')}<b>Remote/Virtual</b></div>
      <div className="methodology-section">
        <Subhead text="A. Scope of Evaluation (Data Sources)" />
        <CheckLines values={dataSources} selected={normalized.dataSources} descriptions={[
          'Direct assessment of the employee at their workstation.',
          'Analysis of Essential Job Functions (EJF).',
          'Review of physician-mandated work restrictions.',
          'Subjective report of discomfort and work habits.',
        ]} />
      </div>
      <div className="methodology-section">
        <Subhead text="B. Evidence-Based Assessment Tools (Quantifying Risk)" />
        <CheckLines values={assessmentTools} selected={normalized.assessmentTools} descriptions={[
          'Rapid Upper Limb Assessment (Neck, trunk, and upper limbs).',
          'Rapid Entire Body Assessment (Whole-body posture and load).',
          'Rapid Office Strain Assessment (Computer/desk environments).',
          'Lifting Equation (Manual material handling tasks).',
        ]} />
      </div>
      <div className="methodology-section">
        <Subhead text="C. Regulatory & Industry Benchmarks" />
        <CheckLines values={benchmarks} selected={normalized.benchmarks} descriptions={[
          'Human Factors Engineering of Computer Workstations.',
          'California Repetitive Motion Injury Standard.',
        ]} />
      </div>
    </Page>
  ))

  pages.push((page) => (
    <Page page={page} total={totalPages} section="III.  COMPREHENSIVE WORKSTATION ANALYSIS">
      <p className="italic roomy">Measurements are recorded to identify deviations from neutral "Home Row" positioning.</p>
      <Subhead text="A. Anthropometric & Equipment Data" />
      <div className="equipment-band">
        <b>Stature:</b>{field(normalized.stature)} <b>Dominant Hand:</b>{mark(normalized.dominantHand === 'R')}<b>R</b>{mark(normalized.dominantHand === 'L')}<b>L</b>
        <b>Vision:</b>{['Single', 'Bifocal', 'Progressive'].map((v) => <React.Fragment key={v}>{mark(normalized.vision === v)}<b>{v}</b></React.Fragment>)}
      </div>
      <div className="equipment-band white"><b>Desk Type:</b>{['Fixed', 'Electric Sit/Stand', 'Desk Converter'].map((v) => <React.Fragment key={v}>{mark(normalized.deskType === v)}<b>{v}</b></React.Fragment>)}</div>
      <div className="equipment-band"><b>Current Desk Hgt:</b>{field(normalized.currentDeskHeight)}<b>Ideal Keying Hgt:</b>{field(normalized.idealKeyingHeight)}<b>Delta:</b>{field(normalized.delta)}</div>
      <Subhead text="B. Postural Deviation Matrix" />
      <div className="risk-key"><b>Risk Level:</b><span className="risk l">L = Low</span><span className="risk m">M = Medium</span><span className="risk h">H = High</span></div>
      <table className="matrix"><thead><tr><th>Body Segment</th><th>Observed Posture / Risk Factor</th><th>Risk Level</th></tr></thead><tbody>
        {postureRows.map((row) => <tr key={row.key}><td><b>{row.label}</b></td><td>{row.factors.map((factor) => <span className="factor" key={factor}>{mark(normalized.postures[row.key].factors.includes(factor))}<b>{factor}</b></span>)}</td><td><span className={`risk ${normalized.postures[row.key].risk.toLowerCase()}`}>{normalized.postures[row.key].risk}</span></td></tr>)}
      </tbody></table>
      <h3>Postural Observations:</h3>
      <TextBox className="observations" text={observationChunks[0]} />
    </Page>
  ))

  observationChunks.slice(1).forEach((chunk, index) => {
    pages.push((page) => (
      <Page page={page} total={totalPages} section={`III.  COMPREHENSIVE WORKSTATION ANALYSIS — POSTURAL OBSERVATIONS CONTINUED ${index + 1}`}>
        <TextBox className="continuation-box continuation-clean" text={chunk} />
      </Page>
    ))
  })

  photoPlan.forEach((item) => {
    if (item.type === 'photo') {
      pages.push((page) => (
        <PhotoPage
          page={page}
          total={totalPages}
          section={item.section}
          exhibits={item.exhibits}
          onUpload={onPhotoUpload}
        />
      ))
      return
    }

    item.chunks.forEach((chunk, chunkIndex) => {
      pages.push((page) => (
        <Page page={page} total={totalPages} section={`IV.  EXHIBIT ${item.exhibitNumber} ANALYSIS — CONTINUATION ${chunkIndex + 1}`}>
          <TextBox className={`analysis-continuation continuation-clean ${chunk.length < 1250 ? 'analysis-continuation-compact' : ''}`} text={chunk} />
        </Page>
      ))
    })
  })

  pages.push((page) => (
      <Page page={page} total={totalPages} section="V.  FINDINGS & CORRECTIVE ACTION PLAN" className="controls-page">
      <Subhead text="1. Administrative Controls & Training" />
      <CheckLines values={controls} selected={normalized.controls} descriptions={[
        'Educated on 20/20/20 rule and movement pacing.',
        `Reviewed use of: ${softwareOptions.join('   ')}`,
        'Trained on chair mechanics and monitor height optimization.',
      ]} />
      <Subhead text="2. Expert Justification Repository" />
      <p className="small italic">Select all applicable justifications. Selected items will appear in the editable clinical justification display.</p>
      <ol className="justification-list">
        {justifications.map((item) => <li key={item}>{mark(normalized.justifications.includes(item))}<span>{item}</span></li>)}
      </ol>
    </Page>
  ))

  justificationChunks.forEach((chunk, index) => {
    pages.push((page) => (
      <Page page={page} total={totalPages} section={index === 0 ? 'V.  FINDINGS & CORRECTIVE ACTION PLAN — CLINICAL JUSTIFICATION' : `V.  CLINICAL JUSTIFICATION — CONTINUATION ${index}`}>
        <TextBox className="justification-display continuation-clean" text={chunk} />
      </Page>
    ))
  })

  equipmentGroups.forEach((group, index) => {
    pages.push((page) => (
      <Page page={page} total={totalPages} section={index === 0 ? 'VI.   ERGONOMIC EQUIPMENT RECOMMENDATION' : `VI.   ERGONOMIC EQUIPMENT RECOMMENDATION — CONTINUATION ${index}`} className="equipment-page">
        <EquipmentRecommendationSummary rows={group} />
      </Page>
    ))
  })

  procurementChunks.forEach((chunk, index) => {
    const isCompactProcurement = procurementChunks.length === 1 && chunk.length < 1200
    pages.push((page) => (
      <Page page={page} total={totalPages} section={equipmentGroups.length === 0 && index === 0 ? 'VI.   ERGONOMIC EQUIPMENT RECOMMENDATION' : `VI.   ERGONOMIC EQUIPMENT RECOMMENDATION — PROCUREMENT ${index === 0 ? 'SUMMARY' : `CONTINUATION ${index}`}`}>
        <h3>Procurement Notes & Justification Summary{index > 0 ? ' Continued' : ''}</h3>
        <ProcurementBox className={`procurement ${isCompactProcurement ? 'procurement-compact' : 'procurement-full'}`} text={chunk} />
      </Page>
    ))
  })

  evaluatorSummaryChunks.forEach((chunk, index) => {
    pages.push((page) => (
      <Page page={page} total={totalPages} section={index === 0 ? 'VII. OPINION & PROFESSIONAL CERTIFICATION — EVALUATOR SUMMARY' : `VII. OPINION & PROFESSIONAL CERTIFICATION — SUMMARY CONTINUATION ${index}`}>
        <TextBox className="evaluator evaluator-full continuation-clean" text={chunk} />
      </Page>
    ))
  })

  pages.push((page) => (
    <Page
      page={page}
      total={totalPages}
      section={evaluatorSummaryChunks.length ? 'VII. OPINION & PROFESSIONAL CERTIFICATION — SIGNATURE & CORRESPONDENCE' : 'VII. OPINION & PROFESSIONAL CERTIFICATION'}
      className="certification-page"
    >
      <p className="certification-opinion">It is my professional opinion, held to a reasonable degree of certainty within the field of ergonomics, that the recommended modifications significantly reduce the identified risk factors for cumulative trauma.</p>
      {evaluatorSummaryFitsCertification && (
        <>
          <h3>Evaluator Summary</h3>
          <TextBox className="evaluator evaluator-certification" text={normalized.evaluatorSummary} />
        </>
      )}
      <SignatureBlock report={normalized} />
      <div className="correspondence-block">
        <Subhead text="Inquiries & Correspondence" />
        <p className="contact">Should you have any questions or clinical concerns regarding the findings of this report, please contact our administrative team at <u>Info@ergonomicworks.com</u>. If you are corresponding with the evaluator directly, please ensure that <u>Info@ergonomicworks.com</u> is copied on all emails to maintain a formal record and ensure a timely response.</p>
        <p className="footer-note">This report was prepared by a certified ergonomic specialist. All findings are based on direct observation, interview, and validated assessment protocols.</p>
      </div>
    </Page>
  ))

  return <div className="report-document">{pages.map((renderPage, index) => <React.Fragment key={index}>{renderPage(index + 1)}</React.Fragment>)}</div>
}

function Page({ page, total, title, section, className = '', children }: { page: number; total: number; title?: string; section?: string; className?: string; children: React.ReactNode }) {
  return (
    <section className={`pdf-page ${className}`.trim()}>
      <header className="pdf-header"><img src={logo} alt="" />{title && <h2>{title}</h2>}</header>
      {section && <SectionTitle text={section} />}
      <div className="pdf-content">{children}</div>
      <footer>Page {page} of {total}</footer>
    </section>
  )
}

function SectionTitle({ text }: { text: string }) {
  return <div className="section-title">{text}</div>
}

function Subhead({ text }: { text: string }) {
  return <h3 className="subhead">{text}</h3>
}

function TableOfContents({ sectionPages }: { sectionPages: SectionPages }) {
  const rows = [
    ['I', 'Administrative & Jurisdictional Data', sectionPages.I],
    ['II', 'Assessment Methodology & Compliance Standards', sectionPages.II],
    ['III', 'Comprehensive Workstation Analysis', sectionPages.III],
    ['IV', 'Evidentiary Documentation (Photographs)', sectionPages.IV],
    ['V', 'Findings & Corrective Action Plan', sectionPages.V],
    ['VI', 'Ergonomic Equipment Recommendation', sectionPages.VI],
    ['VII', 'Opinion & Professional Certification', sectionPages.VII],
  ]
  return <table className="toc"><thead><tr><th /><th>Section</th><th>Pg</th></tr></thead><tbody>{rows.map((row) => <tr key={row[0]}><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td></tr>)}</tbody></table>
}

function FormRows({ rows }: { rows: [string, string][] }) {
  return <div className="form-rows">{rows.map(([label, value]) => <div key={label}><b>{label}</b>{field(value)}</div>)}</div>
}

function CheckLines({ values, selected, descriptions }: { values: string[]; selected: string[]; descriptions: string[] }) {
  return <div className="check-lines">{values.map((value, index) => <p key={value}>{mark(selected.includes(value))}<b>{value}:</b> {descriptions[index]}</p>)}</div>
}

export function PhotoPage({ page, total, section, exhibits, onUpload }: { page: number; total: number; section?: string; exhibits: PlannedExhibit[]; onUpload?: (event: ChangeEvent<HTMLInputElement>, exhibitIndex: number, side: PhotoSide) => void }) {
  return (
    <Page page={page} total={total} section={section} className={exhibits.length === 1 ? 'photo-page single-exhibit-page' : 'photo-page multi-exhibit-page'}>
      {exhibits.map((item) => {
        const number = item.index + 1
        const analysis = normalizeBoxText(item.chunks[0] || item.exhibit.analysis).join('\n\n')
        const hasPre = Boolean(item.exhibit.pre)
        const hasPost = Boolean(item.exhibit.post)
        const hasAnyPhoto = hasPre || hasPost

        return <div className="exhibit" key={number}>
          <Subhead text={`Exhibit ${number}  —  Photographic Evidence`} />
          {hasAnyPhoto && (
            <>
              <div className={`photo-labels ${hasPre && hasPost ? '' : 'single-photo'}`}>
                {hasPre && <b>EXHIBIT A: INITIAL FINDINGS (PRE)</b>}
                {hasPost && <b>EXHIBIT B: CORRECTIVE ACTION</b>}
              </div>

              <div className={`photo-grid ${hasPre && hasPost ? '' : 'single-photo'}`}>
                {hasPre && <PhotoBox photo={item.exhibit.pre} exhibitIndex={item.index} side="pre" onUpload={onUpload} />}
                {hasPost && <PhotoBox photo={item.exhibit.post} exhibitIndex={item.index} side="post" onUpload={onUpload} />}
              </div>
            </>
          )}
          <h3>Analysis — Exhibit {number}:</h3>
          <TextBox className={`analysis-box ${analysis.length < 850 ? 'analysis-compact' : 'analysis-roomy'}`} text={analysis} />
        </div>
      })}
    </Page>
  )
}

function PhotoBox({ photo, exhibitIndex, side, onUpload }: { photo?: PhotoSlot; exhibitIndex: number; side: PhotoSide; onUpload?: (event: ChangeEvent<HTMLInputElement>, exhibitIndex: number, side: PhotoSide) => void }) {
  return (
    <label className="photo-box upload-box">
      {photo ? <img src={photo.url} alt={photo.name} /> : <span className="photo-placeholder">Click to Upload</span>}
      {onUpload && <input className="preview-upload-input" type="file" accept="image/*" onChange={(event) => onUpload(event, exhibitIndex, side)} />}
    </label>
  )
}

function normalizeBoxText(text: string) {
  return cleanImportedText(text)
    .split(/\n{2,}/)
    .map((paragraph) =>
      paragraph
        .replace(/\n+/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim()
    )
    .filter(Boolean)
}

function TextBox({
  text,
  className,
  linkify = false,
}: {
  text: string
  className: string
  linkify?: boolean
}) {
  const paragraphs = normalizeBoxText(text)

  return (
    <div className={`large-box justified-box ${className}`}>
      {paragraphs.length ? (
        paragraphs.map((paragraph, index) => (
          <p key={index}>
            {linkify ? <LinkifiedText text={paragraph} /> : paragraph}
          </p>
        ))
      ) : (
        <p>-</p>
      )}
    </div>
  )
}

function LinkifiedText({ text }: { text: string }) {
  const parts = text.split(/(https?:\/\/[^\s]+)/g)
  return (
    <>
      {parts.map((part, index) => part.match(/^https?:\/\//)
        ? <a key={index} href={part} target="_blank" rel="noreferrer">{part}</a>
        : <React.Fragment key={index}>{part}</React.Fragment>)}
    </>
  )
}

function ProcurementBox({ text, className }: { text: string; className: string }) {
  const formatted = formatProcurementSummary(text)
  const normalized = formatted.replace(/\s+-\s+(?=[A-Z0-9])/g, '\n- ')
  const lines = normalized.split('\n').map((line) => line.trim()).filter(Boolean)
  const firstBullet = lines.findIndex((line) => /^-\s+/.test(line))

  if (firstBullet === -1) {
    const knownItems = [
      'Two 48"-55" sit-stand desks',
      'Full back, head and neck support ergonomic chair',
      'Adjustable height, telescopic (max 21") laptop riser',
      'Dual monitor arms system',
      'Document holder',
      'Under-Desk Keyboard tray',
    ]
    const intro = lines[0] || ''
    const closing = lines.find((line) => /^These items can\b/i.test(line)) || ''
    const detected = knownItems.filter((item) => normalized.toLowerCase().includes(item.toLowerCase()))
    if (!detected.length) return <TextBox className={className} text={normalized} linkify />

    return (
      <div className={`large-box ${className} procurement-formatted`}>
      {intro && <p><LinkifiedText text={intro} /></p>}
      <ul>
        {detected.map((item, index) => <li key={`${item}-${index}`}><span className="procurement-bullet" /><span><LinkifiedText text={item} /></span></li>)}
      </ul>
      {closing && <p><LinkifiedText text={closing} /></p>}
    </div>
    )
  }

  const lastBullet = lines.reduce((last, line, index) => /^-\s+/.test(line) ? index : last, firstBullet)
  const intro = lines.slice(0, firstBullet).join(' ')
  const items = lines.slice(firstBullet, lastBullet + 1).map((line) => line.replace(/^-\s+/, '').trim())
  const closing = lines.slice(lastBullet + 1).join(' ')

  return (
    <div className={`large-box ${className} procurement-formatted`}>
      {intro && <p><LinkifiedText text={intro} /></p>}
      <ul>
        {items.map((item, index) => <li key={`${item}-${index}`}><span className="procurement-bullet" /><span><LinkifiedText text={item} /></span></li>)}
      </ul>
      {closing && <p><LinkifiedText text={closing} /></p>}
    </div>
  )
}

function EquipmentRecommendationSummary({ rows }: { rows: EquipmentRow[] }) {
  if (!rows.length) {
    return null
  }

  return (
    <div className="equipment-summary">
      {rows.map((row, index) => (
        <section className="equipment-card" key={index}>
          <header className="equipment-card-head">
            <b>{row.category || 'Equipment Recommendation'}</b>
            <span className={`priority-pill ${row.priority.toLowerCase()}`}>{row.priority || '-'}</span>
          </header>
          <div className="equipment-card-body">
            <div className="equipment-detail">
              <b>Model / SKU</b>
              <span>{row.current || '-'}</span>
            </div>
            <div className="equipment-detail equipment-link">
              <b>Purchase Link</b>
              <span><LinkifiedText text={row.notes || '-'} /></span>
            </div>
            <div className="equipment-detail equipment-action">
              <b>Recommended Action / Risk Mitigation</b>
              <span><LinkifiedText text={row.action || '-'} /></span>
            </div>
          </div>
        </section>
      ))}
    </div>
  )
}

function SignatureBlock({ report }: { report: Report }) {
  const signatureId = `${report.id.slice(0, 8).toUpperCase()}-${(report.signatureSignedAt || report.certificationDate || today).replace(/[^0-9]/g, '').slice(0, 8)}`
  return (
    <div className="secure-signature">
      <div className="signature-line"><b>Evaluator Name (Printed):</b>{field(report.evaluatorName)}</div>
      <div className="signature-grid">
        <img src={logo} alt="" />
        <div className="signature-box">{report.signature ? <img src={report.signature.url} alt="Evaluator signature" /> : <span>Electronic Signature Pending</span>}</div>
        <b>Date</b>{field(report.certificationDate)}
      </div>
      <div className="signature-certification">
        <b>Secure Electronic Signature Block</b>
        <span>Status: {report.signatureConfirmed ? 'Certified by evaluator' : 'Not certified'}</span>
        <span>Signature ID: {signatureId}</span>
        <span>Timestamp: {report.signatureSignedAt ? new Date(report.signatureSignedAt).toLocaleString() : 'Pending'}</span>
      </div>
    </div>
  )
}

if (typeof document !== 'undefined') {
  const rootEl = document.getElementById('app')
  if (rootEl) createRoot(rootEl).render(<App />)
}
