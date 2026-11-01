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

import type { CharShape, ScriptSlot } from './header'
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
// full-width outside the face tables too: jamo, CJK, compatibility forms and
// the symbols Korean faces set an em wide (※, ①, box drawing, □○, ☞)
const WIDE =
  /[\u1100-\u11ff\u203b\u2460-\u24ff\u2500-\u27bf\u2e80-\u9fff\uf900-\ufaff\uff00-\uff60]/

/**
 * The language slot Hangul shapes a character with: each slot of a character
 * shape carries its own scale and letter spacing. Spaces stay with the
 * Hangul slot.
 */
export function slotOf(ch: string): 'hangul' | ScriptSlot {
  const cp = ch.codePointAt(0)!
  if (cp === 0x20 || (cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0x3130 && cp <= 0x318f))
    return 'hangul'
  if ((cp >= 0xa960 && cp <= 0xa97f) || (cp >= 0xac00 && cp <= 0xd7ff)) return 'hangul'
  if (cp < 0x0250 || (cp >= 0x1e00 && cp <= 0x1eff)) return 'latin'
  if ((cp >= 0x2e80 && cp <= 0x2fdf) || (cp >= 0x3400 && cp <= 0x9fff)) return 'hanja'
  if ((cp >= 0xf900 && cp <= 0xfaff) || cp >= 0x20000) return 'hanja'
  if ((cp >= 0x3040 && cp <= 0x30ff) || (cp >= 0x31f0 && cp <= 0x31ff)) return 'japanese'
  if ((cp >= 0x2000 && cp <= 0x2bff) || (cp >= 0x3000 && cp <= 0x303f)) return 'symbol'
  if ((cp >= 0x3200 && cp <= 0x33ff) || (cp >= 0xff00 && cp <= 0xffef)) return 'symbol'
  return 'other'
}

/** a character's scale and letter spacing percents, from its language slot */
export function slotMetrics(ch: string, char: CharShape): { scalePct: number; spacingPct: number } {
  const slot = char.slots ? slotOf(ch) : 'hangul'
  const own = slot === 'hangul' ? undefined : char.slots?.[slot]
  if (own) return { scalePct: own.scalePct ?? 100, spacingPct: own.spacingPct ?? 0 }
  return { scalePct: char.scalePct ?? 100, spacingPct: char.spacingPct ?? 0 }
}

/**
 * A glyph's Hancom advance in 1/1000 em (before scale and letter spacing);
 * null when no table covers its face. Hangul's fixed-width space
 * (hwpx-docx FIXED_SPACE) is a quarter em.
 */
export function glyphEm(ch: string, char: CharShape): number | null {
  if (ch === '\u2005') return 250
  const syllable = SYLLABLE.test(ch)
  const face = syllable || WIDE.test(ch) ? char.fontHangul : (char.fontLatin ?? char.fontHangul)
  const metric = face ? HWP_COMPAT_METRICS[face] : undefined
  if (!face || !metric) return null
  if (syllable) return metric.hangul
  return advancesOf(face)?.get(ch.codePointAt(0)!) ?? (WIDE.test(ch) ? 1000 : 500)
}

/** advance of one character in points, as Hangul sets it; null when no table covers its face */
export function advancePt(ch: string, char: CharShape): number | null {
  const em = glyphEm(ch, char)
  if (em === null) return null
  const size = (char.sizeHalfPoints ?? 20) / 2
  const metrics = slotMetrics(ch, char)
  // letter spacing is a share of the glyph's own (scaled) advance
  return ((size * em) / 1000) * (metrics.scalePct / 100) * (1 + metrics.spacingPct / 100)
}

/** a glyph Word draws an em wide when the table does not say (the face's default) */
export const isWideGlyph = (ch: string): boolean => SYLLABLE.test(ch) || WIDE.test(ch)

/**
 * Extra letter spacing (points, negative) per character index for the spaces
 * of each saved line Hangul condensed; empty when nothing applies.
 */
export function condensedSpaces(
  p: XNode,
  chars: PlainChar[],
  condense: number,
  format: {
    indentLeft?: number
    indentRight?: number
    indentFirstLine?: number
    /** the paragraph head in front of the first line (twips) */
    headTw?: number
    /**
     * the room the text really has (a cell's width inside its margins):
     * Hangul may save, and draw, a cell line wider than its cell (lawgo-17
     * runs 3.7pt past a borderless edge), which Word wraps instead
     */
    maxRoomPt?: number
  } = {},
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
  for (let k = 0; k < lines.length; k++) {
    const { start } = lines[k]
    // the saved width is the column's; the paragraph's indents narrow it,
    // the first line by its own indent, the rest by the hanging one
    const indentTw =
      (format.indentLeft ?? 0) +
      (k === 0 ? (format.indentFirstLine ?? 0) + (format.headTw ?? 0) : 0) +
      (format.indentRight ?? 0)
    const room = Math.min(lines[k].room, format.maxRoomPt ?? Infinity) - indentTw / 20
    // the last line runs to the end; one wider than its room still fit, so
    // Hangul condensed it too (lawgo-17's space-aligned one-line form rows)
    const end = k + 1 < lines.length ? lines[k + 1].start : chars.length
    if (k === lines.length - 1 && end <= start) break
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

/** inline marks that take one text position each (hwpx-docx textOfT) */
const ONE_CHAR = new Set(['hp:fwSpace', 'hp:nbSpace'])

/**
 * A paragraph whose saved layout maps one to one onto its text: plain runs,
 * fixed and non-breaking spaces. A column definition (hp:colPr) draws nothing,
 * but where it sits in the text positions is not known, so it is accepted
 * only in a one-line paragraph, whose only line starts at 0 anyway.
 */
export function plainParagraph(p: XNode): boolean {
  let controls = false
  for (const run of elements(p)) {
    if (run.name === 'hp:linesegarray') continue
    if (run.name !== 'hp:run') return false
    for (const child of elements(run.node)) {
      if (child.name === 'hp:ctrl') {
        if (!elements(child.node).every((c) => c.name === 'hp:colPr')) return false
        controls = true
        continue
      }
      if (child.name !== 'hp:t') return false
      if (!elements(child.node).every((c) => ONE_CHAR.has(c.name))) return false
    }
  }
  if (controls) {
    const array = firstChild(p, 'hp:linesegarray')
    const lines = array ? elements(array).filter((seg) => seg.name === 'hp:lineseg') : []
    if (lines.length !== 1) return false
  }
  return attr(p, 'paraPrIDRef') !== undefined
}
