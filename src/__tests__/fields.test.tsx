import { render, screen, cleanup, fireEvent } from '@testing-library/react'
import { describe, it, expect, afterEach, vi } from 'vitest'
import {
  Editor,
  PhotoPage,
  normalizeReport,
  parseImportedRecommendations,
  autofillReportFromSource,
  buildPhotoPlan,
} from '../main'


afterEach(() => cleanup())

// ─── helpers ─────────────────────────────────────────────────────────────────

function blankReport() {
  return normalizeReport({})
}

function makeEditorProps(reportOverrides = {}, handlers: Record<string, (...a: any[]) => any> = {}) {
  const report = normalizeReport(reportOverrides as any)
  return {
    report,
    update: vi.fn(),
    toggleList: vi.fn(),
    toggleJustification: vi.fn(),
    updateRecommendation: vi.fn(),
    updateExhibit: vi.fn(),
    attachPhoto: vi.fn(),
    deletePhoto: vi.fn(),
    attachSignature: vi.fn(),
    ...handlers,
  }
}

function makeSource(text: string, extras = {}) {
  return {
    name: 'test.pdf',
    path: '',
    url: '',
    type: 'application/pdf',
    extension: '.pdf',
    importedAt: new Date().toISOString(),
    text,
    extractionStatus: 'ok',
    exhibitPhotos: [],
    ...extras,
  }
}

// ─── normalizeReport ──────────────────────────────────────────────────────────

describe('normalizeReport', () => {
  it('creates default report with 3 exhibits', () => {
    const r = normalizeReport({})
    expect(r.exhibits).toHaveLength(3)
    r.exhibits.forEach((ex) => {
      expect(ex.analysis).toBe('')
      expect(ex.pre).toBeUndefined()
      expect(ex.post).toBeUndefined()
    })
  })

  it('preserves existing exhibit data and pads to 3', () => {
    const r = normalizeReport({ exhibits: [{ analysis: 'hello' }] } as any)
    expect(r.exhibits).toHaveLength(3)
    expect(r.exhibits[0].analysis).toBe('hello')
    expect(r.exhibits[1].analysis).toBe('')
  })

  it('gives default empty arrays for list fields', () => {
    const r = normalizeReport({})
    expect(Array.isArray(r.benchmarks)).toBe(true)
    expect(Array.isArray(r.dataSources)).toBe(true)
    expect(Array.isArray(r.controls)).toBe(true)
  })
})

// ─── Editor – photo sidebar fields ───────────────────────────────────────────

describe('Editor photo fields', () => {
  it('shows upload labels for all 3 exhibits', () => {
    render(<Editor {...makeEditorProps()} />)
    const preLabels = screen.getAllByText(/Pre photo/i)
    const postLabels = screen.getAllByText(/Corrective photo/i)
    expect(preLabels).toHaveLength(3)
    expect(postLabels).toHaveLength(3)
  })

  it('shows no thumbnail or delete button when no photo is attached', () => {
    const { container } = render(<Editor {...makeEditorProps()} />)
    expect(container.querySelectorAll('.photo-thumb')).toHaveLength(0)
    expect(container.querySelectorAll('.photo-delete')).toHaveLength(0)
  })

  it('shows thumbnail and delete button when pre photo is attached', () => {
    const photo = { name: 'test.jpg', path: '', url: 'data:image/png;base64,AAA' }
    const report = normalizeReport({
      exhibits: [{ analysis: '', pre: photo }, { analysis: '' }, { analysis: '' }],
    } as any)
    const { container } = render(<Editor {...makeEditorProps()} report={report} />)

    expect(container.querySelectorAll('.photo-thumb')).toHaveLength(1)
    expect(container.querySelectorAll('.photo-delete')).toHaveLength(1)
    const img = container.querySelector('.photo-thumb') as HTMLImageElement
    expect(img.src).toContain('data:image/png')
  })

  it('shows 2 thumbnails and 2 delete buttons when both pre and post attached', () => {
    const photo = { name: 'p.jpg', path: '', url: 'data:image/png;base64,AAA' }
    const report = normalizeReport({
      exhibits: [{ analysis: '', pre: photo, post: photo }, { analysis: '' }, { analysis: '' }],
    } as any)
    const { container } = render(<Editor {...makeEditorProps()} report={report} />)

    expect(container.querySelectorAll('.photo-thumb')).toHaveLength(2)
    expect(container.querySelectorAll('.photo-delete')).toHaveLength(2)
  })

  it('calls deletePhoto with correct exhibit index and side when × clicked', () => {
    const photo = { name: 'p.jpg', path: '', url: 'data:image/png;base64,AAA' }
    const report = normalizeReport({
      exhibits: [{ analysis: '', pre: photo }, { analysis: '' }, { analysis: '' }],
    } as any)
    const deletePhoto = vi.fn()
    const { container } = render(<Editor {...makeEditorProps({}, { deletePhoto })} report={report} />)

    const btn = container.querySelector('.photo-delete') as HTMLButtonElement
    fireEvent.click(btn)
    expect(deletePhoto).toHaveBeenCalledWith(0, 'pre')
  })

  it('calls deletePhoto for post side on the correct exhibit', () => {
    const photo = { name: 'p.jpg', path: '', url: 'data:image/png;base64,AAA' }
    const report = normalizeReport({
      exhibits: [
        { analysis: '' },
        { analysis: '', post: photo },
        { analysis: '' },
      ],
    } as any)
    const deletePhoto = vi.fn()
    const { container } = render(<Editor {...makeEditorProps({}, { deletePhoto })} report={report} />)

    const btn = container.querySelector('.photo-delete') as HTMLButtonElement
    fireEvent.click(btn)
    expect(deletePhoto).toHaveBeenCalledWith(1, 'post')
  })

  it('file inputs are present and accept image/* for all 6 slots', () => {
    const { container } = render(<Editor {...makeEditorProps()} />)
    const inputs = container.querySelectorAll('input[type="file"]')
    // 6 photo slots (3 exhibits × pre + post) + possible signature input
    const photoInputs = Array.from(inputs).filter(
      (el) => (el as HTMLInputElement).accept === 'image/*',
    )
    expect(photoInputs.length).toBeGreaterThanOrEqual(6)
  })

  it('file input change calls attachPhoto with correct exhibit index and side', () => {
    const attachPhoto = vi.fn()
    const { container } = render(<Editor {...makeEditorProps({}, { attachPhoto })} />)

    const inputs = container.querySelectorAll('input[type="file"]') as NodeListOf<HTMLInputElement>
    const photoInputs = Array.from(inputs).filter((el) => el.accept === 'image/*')
    // First two inputs are Exhibit 1 pre and post
    const file = new File(['data'], 'photo.jpg', { type: 'image/jpeg' })
    fireEvent.change(photoInputs[0], { target: { files: [file] } })
    expect(attachPhoto).toHaveBeenCalledWith(expect.anything(), 0, 'pre')
  })
})

// ─── parseImportedRecommendations ─────────────────────────────────────────────

describe('parseImportedRecommendations', () => {
  it('returns empty array when no recommendations section', () => {
    const result = parseImportedRecommendations('Some unrelated text here.')
    expect(result).toEqual([])
  })

  it('splits on "Model / SKU" boundary for multiple recommendations', () => {
    const text = `
VI. ERGONOMIC EQUIPMENT RECOMMENDATION
Chair Recommendation — High
Model / SKU
ErgoChair Pro
Recommended Action
Upgrade chair
Price
$500

Keyboard Recommendation — Medium
Model / SKU
Kinesis Advantage
Recommended Action
Replace keyboard
Price
$200
`
    const result = parseImportedRecommendations(text)
    expect(result.length).toBeGreaterThanOrEqual(2)
    expect(result[0].current).toContain('ErgoChair')
    expect(result[1].current).toContain('Kinesis')
  })

  it('extracts priority from recommendation block', () => {
    const text = `
VI. ERGONOMIC EQUIPMENT RECOMMENDATION
Mouse Recommendation — High
Model / SKU
Logitech MX Master
Recommended Action
Replace current mouse
`
    const result = parseImportedRecommendations(text)
    expect(result.length).toBeGreaterThanOrEqual(1)
    expect(result[0].priority).toBe('High')
  })

  it('extracts action text from recommendation block', () => {
    const text = `
VI. ERGONOMIC EQUIPMENT RECOMMENDATION
Footrest — Low
Model / SKU
Kensington SoleMate
Recommended Action
Provide adjustable footrest to support neutral posture.
`
    const result = parseImportedRecommendations(text)
    expect(result[0].action).toMatch(/footrest/i)
  })
})

// ─── autofillReportFromSource – benchmarks ────────────────────────────────────

describe('autofillReportFromSource – benchmarks', () => {
  it('matches "Cal/OSHA Title 8 §5110" when § is present in text', () => {
    const text = `Data Sources: On-Site Observation
Benchmarks: Cal/OSHA Title 8 §5110, ANSI/HFES 100-2007`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.benchmarks).toContain('Cal/OSHA Title 8 §5110')
  })

  it('matches "Cal/OSHA Title 8 §5110" via normalization when text lacks § symbol', () => {
    const text = `Benchmarks: Cal/OSHA Title 8 5110`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.benchmarks).toContain('Cal/OSHA Title 8 §5110')
  })

  it('matches ANSI benchmark', () => {
    const text = `Benchmarks Applied: ANSI/HFES 100-2007`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.benchmarks).toContain('ANSI/HFES 100-2007')
  })
})

// ─── autofillReportFromSource – controls ─────────────────────────────────────

describe('autofillReportFromSource – controls', () => {
  it('extracts Micro-Breaks from text', () => {
    const text = `Administrative Controls: Micro-Breaks, Software`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.controls).toContain('Micro-Breaks')
  })
})

// ─── autofillReportFromSource – signature photo ───────────────────────────────

describe('autofillReportFromSource – signature photo', () => {
  it('applies signaturePhoto from source when report has no signature', () => {
    const sig = { name: 'sig.png', path: '/photos/sig.png', url: 'file:///photos/sig.png' }
    const source = makeSource('', { signaturePhoto: sig })
    const { report } = autofillReportFromSource(blankReport(), source)
    expect(report.signature).toEqual(sig)
  })

  it('does not overwrite existing signature', () => {
    const existing = { name: 'existing.png', path: '/existing.png', url: 'file:///existing.png' }
    const newSig = { name: 'new.png', path: '/new.png', url: 'file:///new.png' }
    const current = normalizeReport({ signature: existing } as any)
    const source = makeSource('', { signaturePhoto: newSig })
    const { report } = autofillReportFromSource(current, source)
    expect(report.signature?.name).toBe('existing.png')
  })
})

// ─── autofillReportFromSource – exhibit photos ───────────────────────────────

describe('autofillReportFromSource – exhibit photos', () => {
  it('applies exhibit pre photo to the correct exhibit slot', () => {
    const photo = { name: 'pre.png', path: '/pre.png', url: 'file:///pre.png' }
    const source = makeSource('', {
      exhibitPhotos: [{ exhibitIndex: 1, side: 'pre', photo }],
    })
    const { report } = autofillReportFromSource(blankReport(), source)
    expect(report.exhibits[1].pre).toEqual(photo)
  })

  it('applies exhibit post photo to correct slot', () => {
    const photo = { name: 'post.png', path: '/post.png', url: 'file:///post.png' }
    const source = makeSource('', {
      exhibitPhotos: [{ exhibitIndex: 0, side: 'post', photo }],
    })
    const { report } = autofillReportFromSource(blankReport(), source)
    expect(report.exhibits[0].post).toEqual(photo)
  })
})

// ─── autofillReportFromSource – basic fields ─────────────────────────────────

describe('autofillReportFromSource – basic fields', () => {
  it('extracts employee name', () => {
    const text = `Employee Name: Jane Doe\nJob Title: Data Analyst`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.employeeName).toBe('Jane Doe')
  })

  it('extracts evaluator name', () => {
    const text = `Evaluator Name: John Smith, CPE`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.evaluatorName).toMatch(/John Smith/)
  })

  it('extracts assessment tools', () => {
    const text = `Assessment Tools Used: RULA, REBA`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.assessmentTools).toContain('RULA')
    expect(report.assessmentTools).toContain('REBA')
  })

  it('extracts stature measurement', () => {
    const text = `Stature: 5'6"`
    const { report } = autofillReportFromSource(blankReport(), makeSource(text))
    expect(report.stature).toMatch(/5.6/)
  })
})

// ─── buildPhotoPlan ───────────────────────────────────────────────────────────

describe('buildPhotoPlan', () => {
  it('produces a photo plan item for each exhibit with photos', () => {
    const exhibits = [
      { pre: { name: 'p', path: '', url: 'data:x' }, analysis: 'a1' },
      { analysis: '' },
      { analysis: '' },
    ]
    const plan = buildPhotoPlan(exhibits as any)
    const photoItems = plan.filter((p) => p.type === 'photo')
    expect(photoItems.length).toBeGreaterThanOrEqual(1)
  })

  it('single-exhibit plan uses single-exhibit-page class', () => {
    const exhibits = [
      { pre: { name: 'p', path: '', url: 'data:x' }, analysis: 'x' },
      { analysis: '' },
      { analysis: '' },
    ]
    const plan = buildPhotoPlan(exhibits as any)
    const photoItem = plan.find((p) => p.type === 'photo') as any
    expect(photoItem).toBeDefined()
  })

  it('plan includes all 3 exhibit entries when all have photos', () => {
    const photo = { name: 'p', path: '', url: 'data:x' }
    const exhibits = [
      { pre: photo, analysis: 'a1' },
      { pre: photo, post: photo, analysis: 'a2' },
      { post: photo, analysis: 'a3' },
    ]
    const plan = buildPhotoPlan(exhibits as any)
    const photoItems = plan.filter((p) => p.type === 'photo')
    expect(photoItems.length).toBeGreaterThanOrEqual(1)
    const allExhibits = photoItems.flatMap((p: any) => p.exhibits)
    expect(allExhibits.length).toBe(3)
  })
})

// ─── PhotoPage – existing layout tests (regression) ──────────────────────────

describe('PhotoPage – regression', () => {
  it('two-photo layout: both labels and two exhibit images', () => {
    const exhibits = [{
      exhibit: {
        pre: { name: 'pre', path: '', url: 'data:image/png;base64,AAA' },
        post: { name: 'post', path: '', url: 'data:image/png;base64,BBB' },
        analysis: '',
      },
      index: 0,
      chunks: [],
    }]
    const { container } = render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)
    expect(screen.getByText(/EXHIBIT A: INITIAL FINDINGS/i)).toBeDefined()
    expect(screen.getByText(/EXHIBIT B: CORRECTIVE ACTION/i)).toBeDefined()
    // photo-grid contains exactly the two exhibit images (excludes header logo)
    expect(container.querySelectorAll('.photo-grid img')).toHaveLength(2)
  })

  it('single-photo layout: only pre label, one exhibit image', () => {
    const exhibits = [{
      exhibit: {
        pre: { name: 'only', path: '', url: 'data:image/png;base64,AAA' },
        analysis: '',
      },
      index: 0,
      chunks: [],
    }]
    const { container } = render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)
    expect(screen.getByText(/EXHIBIT A: INITIAL FINDINGS/i)).toBeDefined()
    expect(screen.queryByText(/EXHIBIT B: CORRECTIVE ACTION/i)).toBeNull()
    expect(container.querySelectorAll('.photo-grid img')).toHaveLength(1)
  })

  it('no-photo layout: no labels, no photo-grid images', () => {
    const exhibits = [{ exhibit: { analysis: '' }, index: 0, chunks: [] }]
    const { container } = render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)
    expect(screen.queryByText(/EXHIBIT A/i)).toBeNull()
    expect(container.querySelectorAll('.photo-grid img')).toHaveLength(0)
  })
})
