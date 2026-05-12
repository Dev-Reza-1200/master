import { render, screen, cleanup } from '@testing-library/react'
import { describe, it, expect, afterEach } from 'vitest'
import { PhotoPage } from '../main'

afterEach(() => cleanup())

describe('PhotoPage - review tests', () => {
  it('renders both pre and post photos with labels and two images', () => {
    const exhibits = [
      {
        exhibit: {
          pre: { name: 'pre.jpg', url: 'data:image/png;base64,AAAA' },
          post: { name: 'post.jpg', url: 'data:image/png;base64,BBBB' },
          analysis: 'Some analysis here',
        },
        index: 0,
        chunks: ['Some analysis here'],
      },
    ] as any

    const { container } = render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    // Labels
    expect(screen.getByText('EXHIBIT A: INITIAL FINDINGS (PRE)')).toBeDefined()
    expect(screen.getByText('EXHIBIT B: CORRECTIVE ACTION')).toBeDefined()

    // Grid contains two images (photo pre + photo post)
    const grid = container.querySelector('.photo-grid')
    expect(grid).not.toBeNull()
    expect(grid!.querySelectorAll('img').length).toBe(2)
  })

  it('renders single pre photo with single-photo classes and one image', () => {
    const exhibits = [
      {
        exhibit: {
          pre: { name: 'pre.jpg', url: 'data:image/png;base64,AAAA' },
          analysis: 'Single pre analysis',
        },
        index: 0,
        chunks: ['Single pre analysis'],
      },
    ] as any

    const { container } = render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    expect(container.querySelector('.photo-labels.single-photo')).not.toBeNull()
    expect(container.querySelector('.photo-grid.single-photo')).not.toBeNull()
    expect(container.querySelectorAll('.photo-grid img').length).toBe(1)
  })

  it('renders no photo labels or grid when no photos present', () => {
    const exhibits = [
      {
        exhibit: { analysis: 'No photos' },
        index: 0,
        chunks: ['No photos'],
      },
    ] as any

    render(<PhotoPage page={1} total={1} section="IV" exhibits={exhibits} />)

    expect(screen.queryByText('EXHIBIT A: INITIAL FINDINGS (PRE)')).toBeNull()
    expect(screen.queryByText('EXHIBIT B: CORRECTIVE ACTION')).toBeNull()
    expect(screen.queryByText('Click to Upload')).toBeNull()
  })
})
