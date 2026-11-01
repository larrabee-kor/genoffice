/**
 * Hangul's condense rule (hh:paraPr condense, "줄 나눔 최소 공백"): to keep a
 * started word on a line, Hangul may shrink the line's spaces by up to that
 * share of their width. Word has no such rule, so a converted paragraph
 * wraps one or two syllables early wherever Hangul condensed.
 *
 * Hangul saves the line layout it chose (hp:linesegarray: where each line
 * starts and how wide it may run). With the Hancom advances the compat faces
 * carry (compat-metrics.ts), the natural width of each saved line is known;
 * a line wider than its room is one Hangul condensed, and its inner spaces
 * get exactly the negative letter spacing that brings it back, so Word and
 * GenOffice break where Hangul did. Leading spaces never shrink (rhwp's
 * Hangul 2024 ladder: 8/8 paragraphs only with that rule). A line the model
 * cannot explain within the condense share is left alone.
 */

import type { CharShape } from './header'
import { HWP_COMPAT_METRICS } from './compat-metrics'
import { attr, elements, firstChild, intAttr, type XNode } from './xml'

export interface PlainChar {
  ch: string
  char: CharShape
}

const tables = new Map<string, Map<number, number>>()

function advancesOf(face: string): Map<number, number> | null {
  const cached = tables.get(face)
  if (cached) return cached
  const metric = HWP_COMPAT_METRICS[face]
  if (!metric) return null
  const map = new Map<number, number>()
  for (const [first, ...advances] of metric.runs) {
    advances.forEach((advance, index) => map.set(first + index, advance))
  }
  tables.set(face, map)
  return map
}

const SYLLABLE = /[가-힣]/
const WIDE = /[ᄀ-ᇿ⺀-鿿豈-﫿＀-｠]/

/** advance of one character in points, as Hangul sets it; null when no table covers its face */
export function advancePt(ch: string, char: CharShape): number | null {
  const size = (char.sizeHalfPoints ?? 20) / 2
  const ratio = (char.scalePct ?? 100) / 100
  const spacing = (char.spacingPct ?? 0) / 100
  const syllable = SYLLABLE.test(ch)
  const face = syllable || WIDE.test(ch) ? char.fontHangul : (char.fontLatin ?? char.fontHangul)
  const metric = face ? HWP_COMPAT_METRICS[face] : undefined
  if (!face || !metric) return null
  let em: number
  if (syllable) em = metric.hangul
  else em = advancesOf(face)?.get(ch.codePointAt(0)!) ?? (WIDE.test(ch) ? 1000 : 500)
  // letter spacing is a share of the glyph's own (scaled) advance
  return ((size * em) / 1000) * ratio * (1 + spacing)
}

/**
 * Extra letter spacing (points, negative) per character index for the spaces
 * of each saved line Hangul condensed; empty when nothing applies.
 */
export function condensedSpaces(
  p: XNode,
  chars: PlainChar[],
  condense: number,
  format: { indentLeft?: number; indentRight?: number; indentFirstLine?: number } = {},
): Map<number, number> {
  const out = new Map<number, number>()
  const array = firstChild(p, 'hp:linesegarray')
  if (!array || condense <= 0) return out
  const lines = elements(array)
    .filter((seg) => seg.name === 'hp:lineseg')
    .map((seg) => ({
      start: intAttr(seg.node, 'textpos', 0),
      room: intAttr(seg.node, 'horzsize', 0) / 100,
    }))
  for (let k = 0; k < lines.length - 1; k++) {
    const { start } = lines[k]
    // the saved width is the column's; the paragraph's indents narrow it,
    // the first line by its own indent, the rest by the hanging one
    const indentTw =
      (format.indentLeft ?? 0) +
      (k === 0 ? (format.indentFirstLine ?? 0) : 0) +
      (format.indentRight ?? 0)
    const room = lines[k].room - indentTw / 20
    const end = lines[k + 1].start
    if (room <= 0 || end <= start || end > chars.length) return new Map()
    let last = end
    while (last > start && chars[last - 1].ch === ' ') last--
    let lead = start
    while (lead < last && chars[lead].ch === ' ') lead++
    let natural = 0
    let spaceWidth = 0
    const spaces: number[] = []
    for (let i = start; i < last; i++) {
      const width = advancePt(chars[i].ch, chars[i].char)
      if (width === null) return new Map()
      natural += width
      if (i >= lead && chars[i].ch === ' ') {
        spaces.push(i)
        spaceWidth += width
      }
    }
    const need = natural - room
    if (need <= 0.05 || spaces.length === 0) continue
    // beyond the paragraph's share the saved line is not a condensed one
    if (need > (spaceWidth * condense) / 100 + 0.5) continue
    // a quarter point of slack so the browser's breaker takes the same cut
    const perSpace = -(need + 0.25) / spaces.length
    for (const i of spaces) out.set(i, perSpace)
  }
  return out
}

/** a paragraph whose saved layout maps one to one onto its text: plain runs only */
export function plainParagraph(p: XNode): boolean {
  for (const run of elements(p)) {
    if (run.name === 'hp:linesegarray') continue
    if (run.name !== 'hp:run') return false
    for (const child of elements(run.node)) {
      if (child.name !== 'hp:t') return false
      if (elements(child.node).length > 0) return false
    }
  }
  return attr(p, 'paraPrIDRef') !== undefined
}
