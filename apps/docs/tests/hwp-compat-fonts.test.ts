import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  cssFontFamily,
  hangulSpaceOffsets,
  hwpLayoutOn,
  setHwpCompatFonts,
} from '../src/renderer/line-metrics'
import { HWP_COMPAT_FACES } from '../../../packages/hwpx2ooxml/src/compat-faces'

const FONTS = join(__dirname, '..', 'src', 'renderer', 'fonts')

describe('HWP compat faces', () => {
  afterEach(() => setHwpCompatFonts(false))

  it('leads a Hancom face with its compat face only in an HWP-origin document', () => {
    const word = cssFontFamily('휴먼명조')
    expect(word).not.toContain('HWP Compat')
    setHwpCompatFonts(true)
    expect(hwpLayoutOn()).toBe(true)
    const hwp = cssFontFamily('휴먼명조')
    expect(hwp.startsWith("'HWP Compat HumanMyeongjo',")).toBe(true)
    // the Word chain still follows for what the compat face does not cover
    expect(hwp.endsWith(word)).toBe(true)
    // Hancom's Latin faces chain the face rhwp lays them out with
    expect(cssFontFamily('HCI Poppy').startsWith("'Palatino Linotype','Palatino',")).toBe(true)
    // names outside the table are untouched
    expect(cssFontFamily('Calibri')).not.toContain('HWP Compat')
  })

  it('sets every space half an em wide in an HWP-origin document', () => {
    // Word widens only spaces touching hangul
    expect(hangulSpaceOffsets('  가 b c')).toEqual([1, 3])
    setHwpCompatFonts(true)
    expect(hangulSpaceOffsets('  가 b c')).toEqual([0, 1, 3, 5])
  })

  it('declares a face file for every compat family the table names', () => {
    const css = readFileSync(join(FONTS, 'hwp-compat.css'), 'utf8')
    const families = new Set(
      Object.values(HWP_COMPAT_FACES)
        .flat()
        .filter((f) => f.startsWith('HWP Compat ')),
    )
    expect(families.size).toBeGreaterThan(20)
    for (const family of families) {
      expect(css).toContain(`font-family: '${family}'`)
      const slug = family.slice('HWP Compat '.length)
      expect(readFileSync(join(FONTS, 'hwp-compat', `${slug}.woff2`)).byteLength).toBeGreaterThan(
        1000,
      )
    }
  })
})
