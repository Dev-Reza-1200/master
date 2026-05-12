import { render, screen } from '@testing-library/react'
import { describe, it, expect, afterEach } from 'vitest'
import { cleanup } from '@testing-library/react'
import { PhotoPage } from '../main'

describe('PhotoPage', () => {
  afterEach(() => cleanup())
  it('renders both photo labels and images when both pre and post exist', () => {
    const exhibits = [{
      exhibit: {
        pre: { name: 'pre', path: '', url: 'data:image/png;base64,AAA' },
        post: { name: 'post', path: '', url: 'data:image/png;base64,BBB' },
        analysis: '',
      },
      index: 0,
      chunks: [],
    }]

    render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    expect(screen.getByText(/EXHIBIT A: INITIAL FINDINGS/i)).not.toBeNull()
    expect(screen.getByText(/EXHIBIT B: CORRECTIVE ACTION/i)).not.toBeNull()
    expect(screen.getByAltText('pre')).not.toBeNull()
    expect(screen.getByAltText('post')).not.toBeNull()
  })

  it('renders single-photo layout for only pre photo', () => {
    const exhibits = [{
      exhibit: {
        pre: { name: 'only', path: '', url: 'data:image/png;base64,AAA' },
        analysis: '',
      },
      index: 0,
      chunks: [],
    }]

    render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    expect(screen.getByText(/EXHIBIT A: INITIAL FINDINGS/i)).not.toBeNull()
    expect(screen.queryByText(/EXHIBIT B: CORRECTIVE ACTION/i)).toBeNull()
    expect(screen.getByAltText('only')).not.toBeNull()
  })

  it('renders no photo area when neither photo exists', () => {
    const exhibits = [{
      exhibit: { analysis: '' },
      index: 0,
      chunks: [],
    }]

    render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    expect(screen.queryByText(/EXHIBIT A: INITIAL FINDINGS/i)).toBeNull()
    expect(screen.queryByText(/Click to Upload/i)).toBeNull()
  })
})
