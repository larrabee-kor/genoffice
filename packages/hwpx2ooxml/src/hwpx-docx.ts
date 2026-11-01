/**
 * .hwpx → .docx import filter. The OWPML body (Contents/sectionN.xml) is
 * walked into self-contained OOXML fragments and rebuilt as a fresh .docx
 * through docx-engine's buildBlankDocx → parseDocx → saveDocx — the same
 * assembly the .odt filter and pdf2docx use, so the result is an ordinary
 * editable Word file.
 *
 * Scope, by the "prove it or drop it" rule the other import filters follow:
 * paragraphs and character shapes (font slots, size, bold/italic/underline/
 * strike, color, shading, super/subscript, width scale, letter spacing),
 * paragraph shapes (alignment, indents incl. Hangul's negative indent, spacing, line
 * spacing, page breaks), tables (spans, column widths, row heights, per-cell
 * borders/fills/margins/vertical alignment, nested tables), pictures (inline
 * or floating → inline, also inside cells; BMP re-encoded as PNG), text boxes
 * flattened into the flow, and the first section's page geometry.
 * Headers/footers, notes, equations, shapes, OLE objects and auto-numbering
 * markers are counted in `dropped` instead of guessed at.
 */

import {
  applySectionSettings,
  buildBlankDocx,
  generateParagraphXml,
  inlinePictureDrawingXml,
  keepsHangulKanaSingle,
  NEW_PICTURE_TOKEN_PREFIX,
  parseDocx,
  saveDocx,
  type GeneratedBlock,
  type NewImage,
  type Run,
  type SaveBlock,
  type SectionSettings,
} from '../../docx-engine/src/index'
import { bmpToPng } from './bmp'
import { withFontTable } from './fonts'
import {
  advancePt,
  condensedSpaces,
  glyphEm,
  isWideGlyph,
  plainParagraph,
  slotMetrics,
  type PlainChar,
} from './condense'
import { withHeaderFooter, type HeaderFooterParts } from './headerfooter'
import { HeadingCounter, readHeadings, type Head } from './numbering'
import { withHangulLayoutSettings } from './settings'
import {
  COMPOUND,
  readHeader,
  type BorderFill,
  type BorderSide,
  type CharShape,
  type HeaderTables,
} from './header'
import { HwpxConvertError, loadHwpxPackage, type HwpxPackage } from './package'
import { hwpToPx, hwpToTwips, intrinsicSize, sniffImage } from './units'
import {
  attr,
  children,
  elements,
  escapeXml,
  firstChild,
  firstDescendant,
  intAttr,
  nodeName,
  textOf,
  type XNode,
} from './xml'

export { HwpxConvertError }

export interface HwpxToDocxResult {
  docx: Uint8Array
  /** top-level blocks written (paragraphs + tables) */
  blocks: number
  /** constructs the filter could not represent, with how often each occurred */
  dropped: Array<{ what: string; count: number }>
  /** set when the source was an HWP binary read through rhwp */
  via?: 'rhwp'
}

class Drops {
  private counts = new Map<string, number>()
  add(what: string, count = 1): void {
    this.counts.set(what, (this.counts.get(what) ?? 0) + count)
  }
  list(): Array<{ what: string; count: number }> {
    return [...this.counts].map(([what, count]) => ({ what, count }))
  }
}

const GEN_CTX = {
  headingStyleIds: new Map<number, string>(),
  allocateHyperlinkRel: (href: string): string => {
    throw new Error(`hyperlink reached generated XML: ${href}`)
  },
}

/** a paragraph is assembled from text runs and inline pictures in order */
type Segment =
  | { kind: 'run'; run: Run }
  /** overHwp: the picture's height when it is set over a cell too short for it */
  | { kind: 'picture'; xml: string; overHwp?: number; floating?: boolean; heightHwp?: number }

interface Ctx {
  header: HeaderTables
  pack: HwpxPackage
  drops: Drops
  pictures: Record<string, { base64: string; mime: NewImage['mime'] }>
  pictureSeq: number
  /** resolved BinData → token, so a logo repeated on every page lands once */
  pictureCache: Map<string, { token: string; widthPx: number; heightPx: number } | null>
  section: SectionSettings | null
  /** the sheet of the section being written (its hp:secPr) */
  current: SectionSettings | null
  /** nesting depth of tables being written (cells inside cells) */
  tableDepth: number
  headings: HeadingCounter
  /** the cell whose paragraphs are being written: its content height and vertical alignment */
  cell: {
    contentHwp: number
    /** width inside the cell's left and right margins */
    textWidthHwp: number
    marginsHwp: number
    vAlign: CellInfo['vAlign']
    overHwp: number
  } | null
  /** headers, footers and the page number collected from the controls */
  hf: HeaderFooterParts
  /** the page a top-level paragraph is on (its first line's vertpos restarts per page) */
  page: number
  lastVertpos: number
  /** the table default vertical cell margins (hp:inMargin) of the table being written */
  tableMarginHwp: { top: number; bottom: number; left: number; right: number }
  /** numbering the section's outline levels use (hp:secPr outlineShapeIDRef) */
  outlineId: string
}

const MAX_TABLE_DEPTH = 8

export async function convertHwpxToDocx(hwpx: Uint8Array): Promise<HwpxToDocxResult> {
  const pack = await loadHwpxPackage(hwpx)
  const ctx: Ctx = {
    header: readHeader(pack.header),
    headings: new HeadingCounter(readHeadings(pack.header)),
    cell: null,
    tableMarginHwp: { top: 0, bottom: 0, left: 0, right: 0 },
    hf: {
      header: null,
      footer: null,
      pageNumber: null,
      hideFirstPageNumber: false,
      pageStart: null,
      headerDistTw: null,
      footerDistTw: null,
    },
    page: 0,
    lastVertpos: -1,
    outlineId: '1',
    pack,
    drops: new Drops(),
    pictures: {},
    pictureSeq: 0,
    pictureCache: new Map(),
    section: null,
    current: null,
    tableDepth: 0,
  }

  const body: string[] = []
  let ended: SectionSettings | null = null
  for (const [index, root] of pack.sections.entries()) {
    const sec = root.find((node) => nodeName(node) === 'hs:sec')
    if (!sec) continue
    if (index > 0 && body.length > 0) {
      // trailing empty lines of the previous section only pad the page before
      // the break; where a wider fallback font pushes them over, they would
      // leave a blank page behind
      while (body.length > 1 && isEmptyParagraph(body[body.length - 1])) body.pop()
      // each Hangul section keeps its own sheet (a landscape result table
      // after portrait pages): Word closes a section in its last paragraph
      endSection(body, ctx.current)
      ctx.page++
      ctx.lastVertpos = -1
    }
    for (const { name, node } of elements(sec)) {
      if (name !== 'hp:p') continue
      notePage(node, ctx)
      body.push(...(await paragraphBlocks(node, ctx)))
    }
    ended = ctx.current
  }

  const blocks: SaveBlock[] = body.map((xml) => ({ kind: 'xml', xml }))
  // a document that ends on a table needs a paragraph after it (Word merges two
  // adjacent tables and swallows a trailing one into the sectPr otherwise).
  // Hangul has no such line, so it is a 1pt sliver: a form table that fills
  // its page would otherwise push a full-height empty line onto a blank page
  if (blocks.length === 0) {
    blocks.push({ kind: 'generated', block: { type: 'paragraph', runs: [] } })
  } else if (/^<w:tbl[\s>]/.test(body[body.length - 1])) {
    blocks.push({ kind: 'xml', xml: TABLE_END_PARAGRAPH })
  }

  const parsed = await parseDocx(await buildBlankDocx())
  const saved = await saveDocx(parsed, blocks, {
    ...((ended ?? ctx.section) ? { section: (ended ?? ctx.section)! } : {}),
    ...(Object.keys(ctx.pictures).length > 0 ? { newPictures: ctx.pictures } : {}),
  })
  // original font names stay on the runs; the table tells Word/GenOffice how to substitute
  const withParts = await withHeaderFooter(saved, ctx.hf, pageNumberRunPr(ctx))
  const docx = await withHangulLayoutSettings(await withFontTable(withParts))
  return { docx, blocks: blocks.length, dropped: ctx.drops.list() }
}

const TABLE_END_PARAGRAPH =
  '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>' +
  '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'

// -- paragraphs ---------------------------------------------------------------

/**
 * One hp:p becomes one or more blocks: Hangul anchors tables inside a
 * paragraph's run, so the text before a table, the table itself and the text
 * after it are separate OOXML blocks. A paragraph whose only content is a
 * table does not leave an empty paragraph behind.
 */
async function paragraphBlocks(p: XNode, ctx: Ctx, lastInCell = false): Promise<string[]> {
  const out: string[] = []
  const shape = ctx.header.paraShape(attr(p, 'paraPrIDRef'))
  const format = { ...shape.format }
  if (attr(p, 'pageBreak') === '1') format.pageBreakBefore = true
  // the paragraph head (bullet / number) goes in front of the first text
  let head = shape.numbering
    ? ctx.headings.next(
        shape.numbering,
        shape.numbering === 'outline' ? ctx.outlineId : shape.headingIdRef,
        shape.level ?? 0,
      )
    : undefined
  if (head) ctx.drops.add('auto-numbering-static')
  else if (shape.numbering) ctx.drops.add('auto-numbering')

  let segments: Segment[] = []
  let markSize: number | undefined
  let emittedAny = false
  let hadTable = false
  const trailing: string[] = []

  // Hangul's condense rule needs the saved line layout matched char by char
  let fitChars: PlainChar[] | null = shape.condense && plainParagraph(p) ? [] : null
  // the head shares the first saved line with the text
  let headTw = 0
  const flushFitted = (): void => {
    if (!fitChars) return
    const squeeze = condensedSpaces(p, fitChars, shape.condense ?? 0, {
      ...format,
      headTw,
      ...(ctx.cell ? { maxRoomPt: ctx.cell.textWidthHwp / 100 } : {}),
    })
    let i = 0
    while (i < fitChars.length) {
      const { char } = fitChars[i]
      const extra = squeeze.get(i)
      if (extra !== undefined) {
        const run = runOf(fitChars[i].ch, char)
        run.charSpacingTwips = Math.round((run.charSpacingTwips ?? 0) + extra * 20)
        segments.push({ kind: 'run', run })
        i++
        continue
      }
      let j = i
      while (j < fitChars.length && fitChars[j].char === char && !squeeze.has(j)) j++
      appendText(
        segments,
        fitChars
          .slice(i, j)
          .map((c) => c.ch)
          .join(''),
        char,
      )
      i = j
    }
    fitChars = null
  }

  const flush = (force: boolean, final = false): void => {
    if (final) flushFitted()
    const hasContent = segments.some(
      (segment) => segment.kind === 'picture' || segment.run.text.length > 0,
    )
    if (!hasContent && !force) {
      segments = []
      return
    }
    const paraFormat = emittedAny ? { ...format, pageBreakBefore: undefined } : { ...format }
    const textless = segments.every((s) => s.kind === 'picture' || !s.run.text.trim())
    // a lone picture taller than its cell floats over it (overCellPictureXml)
    const over =
      segments.length === 1 && segments[0].kind === 'picture' ? (segments[0].overHwp ?? 0) : 0
    if (shape.linePercent) {
      // Hangul's percent line spacing is a share of the font size itself, so
      // the line is the largest glyph × percent
      const sizes = segments.flatMap((segment) =>
        segment.kind === 'run' && segment.run.sizeHalfPoints ? [segment.run.sizeHalfPoints] : [],
      )
      const largest = Math.max(markSize ?? 20, ...sizes)
      // the font does not enter Hangul's rule, so an exact line keeps a
      // substituted face's taller line box (Malgun pitches at 1.73 em) from
      // growing the page; only lines holding pictures stay "at least", since an
      // exact line would clip them
      // a placed (floating) picture does not sit on the line
      const hasPicture =
        !over && segments.some((segment) => segment.kind === 'picture' && !segment.floating)
      paraFormat.lineRule = hasPicture ? 'atLeast' : 'exact'
      paraFormat.lineRawTwips = Math.round(largest * 10 * (shape.linePercent / 100))
      // a picture set like a character makes the line as tall as itself, and
      // Hangul spaces that line by the percent too (below 100% a share of the
      // picture: lawgo-9 sets 610 Hanja as 15.6pt pictures in lines advancing
      // 11pt and draws them over each other's margins); the saved layout has
      // that advance, so the line keeps it exactly (GenOffice lets such a
      // picture overlap in HWP-origin documents, styles.css). A paragraph of
      // pictures alone is already as tall as its pictures (ftc-57's product
      // photos), so it keeps the picture line GenOffice gives it
      const saved = hasPicture && !textless ? savedLineAdvance(p) : 0
      if (saved > 0) {
        paraFormat.lineRule = 'exact'
        paraFormat.lineRawTwips = saved
      }
      delete paraFormat.lineSpacing
      // the gap under a cell's last line is not laid out (the row ends at the
      // glyphs), so its lines share the height Hangul recorded for them
      const cellEnd =
        final && lastInCell && !hadTable && (!hasPicture || saved > 0) && !over
          ? lastLinesHeight(p)
          : 0
      if (cellEnd > 0) paraFormat.lineRawTwips = cellEnd
    }
    // nor is the space under a cell's last paragraph: Hangul's row ends at its
    // last line (lawgo-9's Hanja rows keep 30.5pt with 10pt below each cell)
    if (final && lastInCell && !hadTable) paraFormat.spaceAfter = 0
    if (over && segments[0].kind === 'picture') {
      segments = [
        {
          kind: 'picture',
          xml: overCellPictureXml(segments[0].xml, paraFormat, over, ctx.cell!.vAlign),
        },
      ]
    } else if (ctx.cell && textless && segments.some((s) => s.kind === 'picture' && s.floating)) {
      // a placed picture alone in a cell paragraph does not push the empty line
      // down in Hangul: it sits in the cell by the cell's vertical alignment
      const cell = ctx.cell
      segments = segments.map((segment) =>
        segment.kind === 'picture' && segment.floating && segment.heightHwp
          ? {
              ...segment,
              xml: cellFloatingPictureXml(segment.xml, paraFormat, segment.heightHwp, cell.vAlign),
            }
          : segment,
      )
    }
    out.push(paragraphXml(segments, paraFormat, markSize, shape.eastAsian))
    emittedAny = true
    segments = []
  }

  for (const run of elements(p)) {
    if (run.name !== 'hp:run') continue
    const char = ctx.header.charShape(attr(run.node, 'charPrIDRef'))
    if (char.sizeHalfPoints) markSize = char.sizeHalfPoints
    for (const { name, node } of elements(run.node)) {
      switch (name) {
        case 'hp:t': {
          const text = textOfT(node, ctx)
          // Hangul leaves half the glyph size between the head and the text (textOffset 50%)
          if (head && text) {
            appendText(segments, `${head.text}\u2002`, char)
            headTw = headWidthTw(head, char)
            // autoIndent: the following lines start under the text, not the head
            // (a paragraph with its own first-line indent keeps that one)
            if (head.autoIndent && !format.indentFirstLine) {
              format.indentLeft = (format.indentLeft ?? 0) + headTw
              format.indentFirstLine = -headTw
            }
          }
          if (text) head = undefined
          // a condensed paragraph's text waits for its line fitting (flush)
          if (fitChars) for (const ch of text) fitChars.push({ ch, char })
          else appendText(segments, text, char)
          break
        }
        case 'hp:secPr':
          ctx.current = sectionSettings(node) ?? ctx.current
          if (!ctx.section) {
            ctx.section = ctx.current
            const margin = firstChild(firstChild(node, 'hp:pagePr'), 'hp:margin')
            ctx.hf.headerDistTw = hwpToTwips(intAttr(margin, 'top', 0))
            ctx.hf.footerDistTw = hwpToTwips(intAttr(margin, 'bottom', 0))
          }
          ctx.outlineId = attr(node, 'outlineShapeIDRef') ?? ctx.outlineId
          break
        case 'hp:ctrl':
          await controls(node, ctx)
          break
        case 'hp:tbl': {
          flush(false)
          // a table caption sits above or below the table as its own paragraphs
          const caption = firstChild(node, 'hp:caption')
          const captionBlocks = caption ? await subListBlocks(caption, ctx) : []
          const below = attr(caption, 'side') !== 'TOP'
          if (!below) out.push(...captionBlocks)
          out.push(await tableXml(node, ctx, format.align))
          if (below) out.push(...captionBlocks)
          hadTable = true
          emittedAny = true
          break
        }
        case 'hp:pic': {
          const xml = await pictureRunXml(node, ctx)
          if (!xml) break
          // Hangul keeps a cell at its own height when a picture set like a
          // character is taller than the content area: the picture rides over
          // the margins and rules (korea-156786997 slogan, 32pt in a 27pt cell)
          const height = intAttr(firstChild(node, 'hp:sz'), 'height', 0)
          const asChar = attr(firstChild(node, 'hp:pos'), 'treatAsChar') === '1'
          const over = ctx.cell && asChar && height > ctx.cell.contentHwp ? height : 0
          if (ctx.cell && (over || !asChar)) {
            // the cell still grows to the picture plus its own vertical margins
            // (korea-156786997: 3199 + 282 = the table's recorded 3481)
            ctx.cell.overHwp = Math.max(ctx.cell.overHwp, height + ctx.cell.marginsHwp)
          }
          segments.push({
            kind: 'picture',
            xml,
            heightHwp: height,
            ...(over ? { overHwp: over } : {}),
            ...(asChar ? {} : { floating: true }),
          })
          break
        }
        case 'hp:container':
        case 'hp:rect':
        case 'hp:ellipse':
        case 'hp:polygon':
        case 'hp:curve':
        case 'hp:arc':
        case 'hp:line':
        case 'hp:connectLine':
          await shapeContent(node, ctx, segments, trailing)
          break
        case 'hp:equation':
          ctx.drops.add('equations')
          break
        case 'hp:ole':
          ctx.drops.add('ole-objects')
          break
        case 'hp:textart':
          ctx.drops.add('shapes')
          break
        default:
          break
      }
    }
  }
  // an empty paragraph is a real line in Hangul (vertical space); keep it,
  // unless the paragraph only existed to anchor a table
  flush(!hadTable || segments.length > 0 ? !emittedAny : false, true)
  out.push(...trailing)
  return out
}

/** width of a paragraph head and the gap after it, in twips, with Hangul's advances */
function headWidthTw(head: Head, char: CharShape): number {
  const size = (char.sizeHalfPoints ?? 20) / 2
  let pt = 0
  for (const ch of head.text) {
    pt += advancePt(ch, char) ?? (/[\u1100-\uffff]/.test(ch) && ch !== ' ' ? size : size / 2)
  }
  pt += head.textOffset.percent ? (size * head.textOffset.value) / 100 : head.textOffset.value / 100
  return Math.round(pt * 20)
}

/** the tallest line advance (height + gap) of a paragraph's saved layout, in twips; 0 without one */
function savedLineAdvance(p: XNode): number {
  const array = firstChild(p, 'hp:linesegarray')
  if (!array) return 0
  let tallest = 0
  for (const { name, node } of elements(array)) {
    if (name !== 'hp:lineseg') continue
    tallest = Math.max(tallest, intAttr(node, 'vertsize', 0) + intAttr(node, 'spacing', 0))
  }
  return tallest > 0 ? hwpToTwips(tallest) : 0
}

/**
 * Per-line height of a paragraph whose last line gap is dropped, from the
 * layout Hangul saved with it (hp:linesegarray): the lines' heights plus the
 * gaps between them, spread evenly; 0 when the paragraph carries no layout.
 */
function lastLinesHeight(p: XNode): number {
  const array = firstChild(p, 'hp:linesegarray')
  if (!array) return 0
  const lines = elements(array).filter((line) => line.name === 'hp:lineseg')
  if (lines.length === 0) return 0
  let total = 0
  for (const [index, { node }] of lines.entries()) {
    total += intAttr(node, 'vertsize', 0)
    if (index < lines.length - 1) total += intAttr(node, 'spacing', 0)
  }
  return total > 0 ? Math.round(hwpToTwips(total) / lines.length) : 0
}

/** the paragraphs of an element's hp:subList (captions, text boxes, cells) */
async function subListBlocks(owner: XNode, ctx: Ctx): Promise<string[]> {
  const subList = firstChild(owner, 'hp:subList')
  const out: string[] = []
  if (!subList) return out
  for (const { name, node } of elements(subList)) {
    if (name === 'hp:p') out.push(...(await paragraphBlocks(node, ctx)))
  }
  return out
}

/** stand-in for Hangul's fixed-width space (hp:fwSpace) until its run is built */
const FIXED_SPACE = '\u2005'

function appendText(segments: Segment[], text: string, char: CharShape): void {
  if (!text) return
  if (text.includes(FIXED_SPACE)) {
    text.split(FIXED_SPACE).forEach((part, index) => {
      if (index > 0) pushRun(segments, fixedSpaceRun(char))
      appendText(segments, part, char)
    })
    return
  }
  if (!char.sizeHalfPoints || (!char.spacingPct && !char.slots)) {
    pushRun(segments, runOf(text, char))
    return
  }
  spacedRuns(segments, text, char)
}

/** the running difference (twips) between Hangul's letter spacing and the written one, per paragraph piece */
const spacingCarry = new WeakMap<Segment[], number>()
/** how far the written spacing may drift from Hangul's before the value changes (0.15pt) */
const SPACING_SLACK_TWIPS = 3

/**
 * Text whose character shape has letter spacing or per-language scales.
 * Hangul's letter spacing is a percent of each glyph's own advance (rhwp
 * text_measurement: 휴먼명조 '*' 0.5 em at -9% spaces 0.455 em); OOXML's
 * w:spacing is whole twips per glyph, and under the balanced-width setting
 * (settings.ts) Word doubles it on double-byte glyphs except hangul and kana
 * in the faces it keeps single (Malgun Gothic, NanumGothic…). GenOffice
 * spreads that doubling over a run by its share of double-byte glyphs, so
 * double-byte and other glyphs go in separate runs. The whole-twip value
 * misses Hangul's (-1% of 14pt: 0.14pt a syllable, written 0.1pt), which over
 * a 30-syllable line pushed the last syllable down (ftc-58, korea-156787108),
 * so the written value alternates between its neighbours to keep the running
 * total within SPACING_SLACK_TWIPS of Hangul's. Each language slot keeps its
 * own scale and spacing (lawgo-14: digits at 0%, hangul at -5%).
 */
function spacedRuns(segments: Segment[], text: string, char: CharShape): void {
  const size = (char.sizeHalfPoints ?? 20) / 2
  const singleFace = keepsHangulKanaSingle(char.fontHangul ?? char.fontLatin)
  let carry = spacingCarry.get(segments) ?? 0
  let written: number | null = null
  let group = ''
  let key = ''
  let groupValue = 0
  let groupScale = 100
  const flush = (): void => {
    if (!group) return
    const run = runOf(group, char)
    if (groupScale !== 100) run.charScalePct = Math.min(600, Math.max(1, groupScale))
    else delete run.charScalePct
    if (groupValue) run.charSpacingTwips = groupValue
    else delete run.charSpacingTwips
    pushRun(segments, run)
    group = ''
  }
  for (const ch of text) {
    const metrics = slotMetrics(ch, char)
    const doubled = doubledSpacing(ch, singleFace)
    let value = 0
    if (metrics.spacingPct) {
      const em = glyphEm(ch, char) ?? (isWideGlyph(ch) ? 1000 : 500)
      const exact = (metrics.spacingPct / 100) * (em / 1000) * size * (metrics.scalePct / 100) * 20
      const mult = doubled ? 2 : 1
      // only the whole values either side of Hangul's: keep the current one
      // while the running total allows, else take the one that pulls it back
      const lo = Math.floor(exact / mult)
      const hi = Math.ceil(exact / mult)
      const error = (v: number): number => Math.abs(carry + exact - v * mult)
      if (
        written !== null &&
        (written === lo || written === hi) &&
        error(written) <= SPACING_SLACK_TWIPS
      ) {
        value = written
      } else {
        value = error(lo) <= error(hi) ? lo : hi
      }
      carry += exact - value * mult
      written = value
    }
    const next = `${doubled}|${value}|${metrics.scalePct}`
    if (next !== key) {
      flush()
      key = next
      groupValue = value
      groupScale = metrics.scalePct
    }
    group += ch
  }
  flush()
  spacingCarry.set(segments, carry)
}

/** Word doubles letter spacing on this glyph under the balanced-width setting (docx-engine doubledGlyphFraction) */
function doubledSpacing(ch: string, singleFace: boolean): boolean {
  const cp = ch.codePointAt(0)!
  if (!isDoubleByte(cp)) return false
  const hangulKana =
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x3040 && cp <= 0x30ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0x31f0 && cp <= 0x31ff) ||
    (cp >= 0xac00 && cp <= 0xd7af)
  return !(hangulKana && singleFace)
}

/**
 * append a run, merged into the previous one when the formatting matches
 * (fewer w:r); spaced double-byte and other glyphs stay apart (spacedRuns)
 */
function pushRun(segments: Segment[], run: Run): void {
  const last = segments[segments.length - 1]
  if (
    last &&
    last.kind === 'run' &&
    sameFormat(last.run, run) &&
    (!run.charSpacingTwips || spacingClass(last.run) === spacingClass(run))
  ) {
    last.run.text += run.text
  } else segments.push({ kind: 'run', run })
}

/** whether a run's glyphs take Word's doubled letter spacing: all, none or a mix */
function spacingClass(run: Run): 'all' | 'none' | 'mixed' {
  const single = keepsHangulKanaSingle(run.eastAsiaFont)
  let doubled = 0
  let n = 0
  for (const ch of run.text) {
    if (doubledSpacing(ch, single)) doubled++
    n++
  }
  return doubled === 0 ? 'none' : doubled === n ? 'all' : 'mixed'
}

/**
 * Hangul's fixed-width space (고정폭 빈칸) is a quarter of the glyph size wide,
 * letter spacing included (korea-156787289 "위원장 : 출입국" at 14pt -8%:
 * 3.2pt; korea-156787338 at 18pt: 4.4pt), and justification leaves it alone.
 * An en space, which no justifier stretches, carries it: half an em in the
 * compat faces, brought down to the quarter by its letter spacing.
 */
function fixedSpaceRun(char: CharShape): Run {
  const run = runOf('\u2002', char)
  const size = (char.sizeHalfPoints ?? 20) / 2
  const ratio = (char.scalePct ?? 100) / 100
  const spacing = (char.spacingPct ?? 0) / 100
  run.charSpacingTwips = Math.round(20 * size * ratio * (0.25 * (1 + spacing) - 0.5))
  return run
}

/** what Word counts as double-byte for balanced spacing (docx-engine doubledGlyphFraction) */
function isDoubleByte(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x9fff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0x31f0 && cp <= 0x31ff) ||
    (cp >= 0xac00 && cp <= 0xd7af) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xff00 && cp <= 0xff60) ||
    cp >= 0x20000
  )
}

function sameFormat(a: Run, b: Run): boolean {
  const { text: _a, ...restA } = a
  const { text: _b, ...restB } = b
  return JSON.stringify(restA) === JSON.stringify(restB)
}

const isHangulText = (text: string): boolean =>
  /^[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]+$/.test(text)

function runOf(text: string, char: CharShape): Run {
  const run: Run = { text }
  if (char.bold) run.bold = true
  if (char.italic) run.italic = true
  if (char.underline) run.underline = true
  if (char.strike) run.strike = true
  if (char.sizeHalfPoints) run.sizeHalfPoints = char.sizeHalfPoints
  if (char.color) run.color = char.color
  if (char.shading) run.shading = char.shading
  if (char.vertAlign) run.vertAlign = char.vertAlign
  const eastAsia = char.fontHangul ?? char.fontLatin
  const latin = char.fontLatin ?? char.fontHangul
  if (eastAsia) {
    run.font = eastAsia
    run.eastAsiaFont = eastAsia
  }
  if (latin) run.fontAscii = latin
  if (char.scalePct) run.charScalePct = Math.min(600, Math.max(1, char.scalePct))
  if (char.spacingPct && char.sizeHalfPoints) {
    // a single value for the whole text (spacedRuns refines it per glyph):
    // full-width glyphs take the whole percent of the size, half-width ones
    // half of it, and Word doubles w:spacing on double-byte glyphs except
    // hangul in the faces it keeps single
    const twips = (char.spacingPct / 100) * char.sizeHalfPoints * 10
    const single = keepsHangulKanaSingle(run.eastAsiaFont) && isHangulText(text)
    run.charSpacingTwips = Math.round(single ? twips : twips / 2)
  }
  return run
}

/** hp:t mixes literal text with inline markers */
function textOfT(t: XNode, ctx: Ctx): string {
  let out = ''
  for (const child of children(t)) {
    const literal = textOf(child)
    if (literal !== null) {
      out += literal
      continue
    }
    switch (nodeName(child)) {
      case 'hp:tab':
        out += '\t'
        break
      case 'hp:lineBreak':
        out += '\n'
        break
      case 'hp:fwSpace':
        out += FIXED_SPACE
        break
      case 'hp:nbSpace':
        out += ' '
        break
      case 'hp:hyphen':
        out += '-'
        break
      case 'hp:titleMark':
      case 'hp:markpenBegin':
      case 'hp:markpenEnd':
        break
      default:
        ctx.drops.add('inline-markers')
        break
    }
  }
  return out
}

function paragraphXml(
  segments: Segment[],
  format: GeneratedBlock['format'],
  markSize: number | undefined,
  eastAsian?: { autoSpaceDE: boolean; autoSpaceDN: boolean; wordWrap: boolean },
): string {
  const block: GeneratedBlock = { type: 'paragraph', runs: [] }
  if (format && Object.keys(format).length > 0) block.format = format
  let xml = generateParagraphXml(block, GEN_CTX)
  if (eastAsian) xml = withEastAsianBreaking(xml, eastAsian)
  // the paragraph mark carries the size, so an empty Hangul line keeps its height
  if (markSize) {
    const mark = `<w:rPr><w:sz w:val="${markSize}"/><w:szCs w:val="${markSize}"/></w:rPr>`
    xml = xml.includes('</w:pPr>')
      ? xml.replace('</w:pPr>', `${mark}</w:pPr>`)
      : xml.replace(/^<w:p\b[^>]*>/, (open) => `${open}<w:pPr>${mark}</w:pPr>`)
  }
  let runsXml = ''
  let pending: Run[] = []
  const flushRuns = (): void => {
    if (pending.length === 0) return
    const generated = generateParagraphXml({ type: 'paragraph', runs: pending }, GEN_CTX)
    runsXml += innerRuns(generated)
    pending = []
  }
  for (const segment of segments) {
    if (segment.kind === 'run') pending.push(segment.run)
    else {
      flushRuns()
      runsXml += segment.xml
    }
  }
  flushRuns()
  if (xml.endsWith('/>')) return xml.replace(/\/>$/, `>${runsXml}</w:p>`)
  return xml.replace(/<\/w:p>$/, `${runsXml}</w:p>`)
}

/**
 * w:wordWrap / w:autoSpaceDE / w:autoSpaceDN in CT_PPr order: after tabs and
 * shading, before snapToGrid/spacing/ind/jc/rPr. Only non-default values are
 * written (Word defaults are all on).
 */
function withEastAsianBreaking(
  xml: string,
  east: { autoSpaceDE: boolean; autoSpaceDN: boolean; wordWrap: boolean },
): string {
  const flags =
    (east.wordWrap ? '' : '<w:wordWrap w:val="0"/>') +
    (east.autoSpaceDE ? '' : '<w:autoSpaceDE w:val="0"/>') +
    (east.autoSpaceDN ? '' : '<w:autoSpaceDN w:val="0"/>')
  if (!flags) return xml
  if (!xml.includes('<w:pPr>')) {
    return xml.replace(/^<w:p\b[^>]*?(\/?)>/, (open, selfClose: string) =>
      selfClose
        ? `${open.slice(0, -2)}><w:pPr>${flags}</w:pPr></w:p>`
        : `${open}<w:pPr>${flags}</w:pPr>`,
    )
  }
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)![0]
  const anchor =
    /<w:(snapToGrid|spacing|ind|contextualSpacing|jc|textAlignment|outlineLvl|rPr)\b/.exec(pPr)
  const updated = anchor
    ? pPr.slice(0, anchor.index) + flags + pPr.slice(anchor.index)
    : pPr.replace('</w:pPr>', `${flags}</w:pPr>`)
  return xml.replace(pPr, updated)
}

/** the run part of a generated paragraph (between its pPr and </w:p>) */
function innerRuns(paragraph: string): string {
  return paragraph
    .replace(/^<w:p\b[^>]*>/, '')
    .replace(/^<w:pPr>[\s\S]*?<\/w:pPr>/, '')
    .replace(/<\/w:p>$/, '')
}

/** a paragraph with no text, picture or break: vertical space only */
function isEmptyParagraph(xml: string): boolean {
  return (
    /^<w:p[\s>/]/.test(xml) &&
    !/<w:t[\s>]|<w:drawing|<w:br\b|<w:tab\/>/.test(xml) &&
    !xml.includes('<w:pageBreakBefore/>')
  )
}

/**
 * Close the section the body ends with: its sheet as a paragraph-level
 * sectPr in the last paragraph (a closing table gets a 1pt paragraph for it).
 */
function endSection(body: string[], settings: SectionSettings | null): void {
  if (!settings) return
  const base =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
  const sectPr = applySectionSettings(base, settings)
  const last = body[body.length - 1]
  if (last && /^<w:p[\s>]/.test(last) && !last.includes('<w:sectPr')) {
    body[body.length - 1] = last.includes('</w:pPr>')
      ? last.replace('</w:pPr>', `${sectPr}</w:pPr>`)
      : last.replace(/^<w:p\b[^>]*?(\/?)>/, (open, selfClose: string) =>
          selfClose ? `<w:p><w:pPr>${sectPr}</w:pPr></w:p>` : `${open}<w:pPr>${sectPr}</w:pPr>`,
        )
    return
  }
  body.push(TABLE_END_PARAGRAPH.replace('</w:pPr>', `${sectPr}</w:pPr>`))
}

/** a new page starts where a top-level paragraph's first line goes back up the page */
function notePage(p: XNode, ctx: Ctx): void {
  const array = firstChild(p, 'hp:linesegarray')
  const first = array ? elements(array).find((seg) => seg.name === 'hp:lineseg') : undefined
  if (!first) return
  const vertpos = intAttr(first.node, 'vertpos', 0)
  if (attr(p, 'pageBreak') === '1' || vertpos < ctx.lastVertpos) ctx.page++
  ctx.lastVertpos = vertpos
}

async function controls(ctrl: XNode, ctx: Ctx): Promise<void> {
  for (const { name, node } of elements(ctrl)) {
    switch (name) {
      case 'hp:header':
      case 'hp:footer': {
        const key = name === 'hp:header' ? 'header' : 'footer'
        // odd/even variants: the first BOTH (or any) definition stands
        if (ctx.hf[key] && attr(node, 'applyPageType') !== 'BOTH') break
        const cell = ctx.cell
        ctx.cell = null
        try {
          ctx.hf[key] = await subListBlocks(node, ctx)
        } finally {
          ctx.cell = cell
        }
        break
      }
      case 'hp:pageNum':
        ctx.hf.pageNumber ??= {
          pos: attr(node, 'pos') ?? 'BOTTOM_CENTER',
          formatType: attr(node, 'formatType') ?? 'DIGIT',
          sideChar: attr(node, 'sideChar') ?? '',
        }
        break
      case 'hp:pageHiding':
        if (attr(node, 'hidePageNum') === '1' && ctx.page === 0) ctx.hf.hideFirstPageNumber = true
        break
      case 'hp:newNum':
        if (attr(node, 'numType') === 'PAGE' && ctx.page === 0) {
          ctx.hf.pageStart = intAttr(node, 'num', 1)
        }
        break
      case 'hp:footNote':
        ctx.drops.add('footnotes')
        break
      case 'hp:endNote':
        ctx.drops.add('endnotes')
        break
      default:
        break // column settings, fields, bookmarks: layout/meta only
    }
  }
}

// -- section geometry ------------------------------------------------------------

function sectionSettings(secPr: XNode): SectionSettings | null {
  const pagePr = firstChild(secPr, 'hp:pagePr')
  if (!pagePr) return null
  let width = hwpToTwips(intAttr(pagePr, 'width', 59528))
  let height = hwpToTwips(intAttr(pagePr, 'height', 84188))
  // landscape="NARROWLY" rotates the sheet; width/height stay portrait-spelled
  if (attr(pagePr, 'landscape') === 'NARROWLY' && width < height) [width, height] = [height, width]
  const margin = firstChild(pagePr, 'hp:margin')
  const m = (name: string): number => hwpToTwips(intAttr(margin, name, 0))
  // Hangul's top/bottom margins exclude the header/footer bands; Word's body
  // margins are measured from the page edge
  return {
    pageWidth: width,
    pageHeight: height,
    orientation: width > height ? 'landscape' : 'portrait',
    marginTop: m('top') + m('header'),
    marginBottom: m('bottom') + m('footer'),
    marginLeft: m('left') + m('gutter'),
    marginRight: m('right'),
    // the header and footer bands begin at Hangul's own top/bottom margins
    headerDist: m('top'),
    footerDist: m('bottom'),
    pageBorder: false,
    columns: 1,
  }
}

// -- shapes: text boxes flatten, pictures inside groups survive ---------------------

async function shapeContent(
  shape: XNode,
  ctx: Ctx,
  segments: Segment[],
  trailing: string[],
): Promise<void> {
  if (nodeName(shape) === 'hp:container') {
    for (const { name, node } of elements(shape)) {
      if (name === 'hp:pic') {
        const xml = await pictureRunXml(node, ctx)
        if (xml) segments.push({ kind: 'picture', xml })
      } else if (name !== 'hp:offset' && name !== 'hp:orgSz' && name !== 'hp:curSz') {
        await shapeContent(node, ctx, segments, trailing)
      }
    }
    return
  }
  const drawText = firstChild(shape, 'hp:drawText')
  const subList = drawText ? firstChild(drawText, 'hp:subList') : null
  if (subList) {
    ctx.drops.add('textboxes-flattened')
    for (const { name, node } of elements(subList)) {
      if (name === 'hp:p') trailing.push(...(await paragraphBlocks(node, ctx)))
    }
    return
  }
  if (!['hp:offset', 'hp:orgSz', 'hp:curSz', 'hp:flip'].includes(nodeName(shape) ?? '')) {
    ctx.drops.add('shapes')
  }
}

// -- pictures ---------------------------------------------------------------------------

const MIME: Record<string, NewImage['mime']> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

/**
 * A lone picture taller than its cell, as Hangul sets it: anchored to the
 * paragraph without wrapping, centred across the cell (or at the paragraph
 * alignment) and placed by the cell's vertical alignment against the
 * paragraph's own line, so the row keeps its height and the picture overlaps
 * the margins the way the original does. Behind the text layer, so the
 * cell rules still draw over it as Hangul draws them.
 */
function overCellPictureXml(
  runXml: string,
  format: { align?: string; lineRawTwips?: number },
  heightHwp: number,
  vAlign: CellInfo['vAlign'],
): string {
  const lineTw = format.lineRawTwips ?? 0
  const picTw = hwpToTwips(heightHwp)
  const offsetTw =
    vAlign === 'top' ? 0 : vAlign === 'bottom' ? lineTw - picTw : (lineTw - picTw) / 2
  const align = format.align === 'right' ? 'right' : format.align === 'left' ? 'left' : 'center'
  return runXml
    .replace(
      /<wp:inline distT="0" distB="0" distL="0" distR="0">/,
      '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264"' +
        ' behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
        `<wp:positionH relativeFrom="column"><wp:align>${align}</wp:align></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(offsetTw * 635)}</wp:posOffset></wp:positionV>`,
    )
    .replace(/(<wp:effectExtent [^>]*\/>)/, '$1<wp:wrapNone/>')
    .replace('</wp:inline>', '</wp:anchor>')
}

/** a placed picture in a text-less cell paragraph: no wrap, vertical place by the cell alignment */
function cellFloatingPictureXml(
  runXml: string,
  format: { lineRawTwips?: number },
  heightHwp: number,
  vAlign: CellInfo['vAlign'],
): string {
  const lineTw = format.lineRawTwips ?? 0
  const picTw = hwpToTwips(heightHwp)
  const offsetTw =
    vAlign === 'top' ? 0 : vAlign === 'bottom' ? lineTw - picTw : (lineTw - picTw) / 2
  return runXml
    .replace(
      /<wp:positionV relativeFrom="[^"]*">.*?<\/wp:positionV>/,
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(offsetTw * 635)}</wp:posOffset></wp:positionV>`,
    )
    .replace(/<wp:wrap(?:TopAndBottom|Square)[^>]*\/>/, '<wp:wrapNone/>')
}

async function pictureRunXml(pic: XNode, ctx: Ctx): Promise<string | null> {
  const img = firstDescendant(pic, 'hc:img')
  const id = attr(img, 'binaryItemIDRef')
  if (!id) {
    ctx.drops.add('images')
    return null
  }
  const pos = firstChild(pic, 'hp:pos')

  let resolved = ctx.pictureCache.get(id)
  if (resolved === undefined) {
    resolved = await resolvePicture(id, ctx)
    ctx.pictureCache.set(id, resolved)
  }
  if (!resolved) {
    ctx.drops.add('images')
    return null
  }
  // displayed size: hp:sz (or curSz) in HWPUNIT; intrinsic pixels as the fallback
  const sz = firstChild(pic, 'hp:sz') ?? firstChild(pic, 'hp:curSz')
  const w = intAttr(sz, 'width', 0)
  const h = intAttr(sz, 'height', 0)
  const widthPx = w > 0 ? hwpToPx(w) : resolved.widthPx
  const heightPx = h > 0 ? hwpToPx(h) : resolved.heightPx
  const drawing = inlinePictureDrawingXml({
    rId: resolved.token,
    docPrId: resolved.token,
    widthPx,
    heightPx,
  })
  const placed = pos && attr(pos, 'treatAsChar') !== '1'
  return `<w:r>${placed ? floatingPictureXml(drawing, pic, pos) : drawing}</w:r>`
}

/** HWP positions are unsigned 32-bit HWPUNIT fields: wrapped values are negative */
const signedHwp = (value: number): number => (value > 0x7fffffff ? value - 0x100000000 : value)
const hwpToEmu = (hwp: number): number => Math.round(hwpToTwips(hwp) * 635)

const H_RELATIVE: Record<string, string> = {
  PAPER: 'page',
  PAGE: 'margin',
  COLUMN: 'column',
  PARA: 'column',
}
const V_RELATIVE: Record<string, string> = { PAPER: 'page', PAGE: 'margin', PARA: 'paragraph' }
const ALIGN_H: Record<string, string> = {
  CENTER: 'center',
  RIGHT: 'right',
  INSIDE: 'inside',
  OUTSIDE: 'outside',
}
const ALIGN_V: Record<string, string> = {
  CENTER: 'center',
  BOTTOM: 'bottom',
  INSIDE: 'inside',
  OUTSIDE: 'outside',
}

/**
 * A picture Hangul places (treatAsChar="0") as a Word anchor: the same
 * reference frame (paper, page, column, paragraph), alignment or offset,
 * text wrap and outside margins, so it no longer takes a line of its own.
 */
function floatingPictureXml(inlineXml: string, pic: XNode, pos: XNode): string {
  const horz = attr(pos, 'horzAlign') ?? 'LEFT'
  const vert = attr(pos, 'vertAlign') ?? 'TOP'
  const positionH = ALIGN_H[horz]
    ? `<wp:align>${ALIGN_H[horz]}</wp:align>`
    : `<wp:posOffset>${hwpToEmu(signedHwp(intAttr(pos, 'horzOffset', 0)))}</wp:posOffset>`
  const positionV = ALIGN_V[vert]
    ? `<wp:align>${ALIGN_V[vert]}</wp:align>`
    : `<wp:posOffset>${hwpToEmu(signedHwp(intAttr(pos, 'vertOffset', 0)))}</wp:posOffset>`
  const margin = firstChild(pic, 'hp:outMargin')
  const dist = (side: string) => hwpToEmu(intAttr(margin, side, 0))
  const wrapMode = attr(pic, 'textWrap') ?? 'TOP_AND_BOTTOM'
  const behind = wrapMode === 'BEHIND_TEXT'
  const wrap =
    wrapMode === 'TOP_AND_BOTTOM'
      ? '<wp:wrapTopAndBottom/>'
      : wrapMode === 'BEHIND_TEXT' || wrapMode === 'IN_FRONT_OF_TEXT'
        ? '<wp:wrapNone/>'
        : '<wp:wrapSquare wrapText="bothSides"/>'
  const z = 251659264 + Math.max(0, intAttr(pic, 'zOrder', 0))
  return inlineXml
    .replace(
      /<wp:inline distT="0" distB="0" distL="0" distR="0">/,
      `<wp:anchor distT="${dist('top')}" distB="${dist('bottom')}" distL="${dist('left')}" distR="${dist('right')}"` +
        ` simplePos="0" relativeHeight="${z}" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1"` +
        ` allowOverlap="${attr(pos, 'allowOverlap') === '1' ? 1 : 0}"><wp:simplePos x="0" y="0"/>` +
        `<wp:positionH relativeFrom="${H_RELATIVE[attr(pos, 'horzRelTo') ?? 'COLUMN'] ?? 'column'}">${positionH}</wp:positionH>` +
        `<wp:positionV relativeFrom="${V_RELATIVE[attr(pos, 'vertRelTo') ?? 'PARA'] ?? 'paragraph'}">${positionV}</wp:positionV>`,
    )
    .replace(/(<wp:effectExtent [^>]*\/>)/, `$1${wrap}`)
    .replace('</wp:inline>', '</wp:anchor>')
}

async function resolvePicture(
  id: string,
  ctx: Ctx,
): Promise<{ token: string; widthPx: number; heightPx: number } | null> {
  const binary = await ctx.pack.readBinary(id)
  if (!binary || binary.bytes.length === 0) return null
  let bytes = binary.bytes
  let kind = sniffImage(bytes)
  if (kind === 'bmp') {
    const png = bmpToPng(bytes)
    if (!png) return null
    bytes = png
    kind = 'png'
  }
  const mime = kind ? MIME[kind] : undefined
  if (!mime) {
    ctx.drops.add(kind === 'wmf' || kind === 'emf' ? 'vector-images' : 'unknown-images')
    return null
  }
  const size = intrinsicSize(bytes) ?? { width: 400, height: 300 }
  const token = `${NEW_PICTURE_TOKEN_PREFIX}${++ctx.pictureSeq}`
  ctx.pictures[token] = { base64: Buffer.from(bytes).toString('base64'), mime }
  return { token, widthPx: size.width, heightPx: size.height }
}

// -- tables -----------------------------------------------------------------------------

interface CellInfo {
  row: number
  col: number
  rowSpan: number
  colSpan: number
  widthHwp: number
  heightHwp: number
  border: BorderFill | null
  vAlign: 'top' | 'center' | 'bottom'
  margin: { left: number; right: number; top: number; bottom: number } | null
  /** height of the tallest picture set over this cell (overCellPictureXml); 0 = none */
  overHwp: number
  contentXml: string
}

const V_ALIGN: Record<string, CellInfo['vAlign']> = {
  TOP: 'top',
  CENTER: 'center',
  BOTTOM: 'bottom',
}

async function tableXml(tbl: XNode, ctx: Ctx, hostAlign?: string): Promise<string> {
  if (ctx.tableDepth >= MAX_TABLE_DEPTH) {
    ctx.drops.add('deep-nested-tables')
    return '<w:p/>'
  }
  ctx.tableDepth++
  try {
    const rowCount = Math.max(1, intAttr(tbl, 'rowCnt', 0))
    const colCount = Math.max(1, intAttr(tbl, 'colCnt', 0))
    const cells: CellInfo[] = []
    let rowIndex = 0
    const inMarginNode = firstChild(tbl, 'hp:inMargin')
    const tableMarginHwp = {
      top: intAttr(inMarginNode, 'top'),
      bottom: intAttr(inMarginNode, 'bottom'),
      left: intAttr(inMarginNode, 'left'),
      right: intAttr(inMarginNode, 'right'),
    }
    for (const { name, node: tr } of elements(tbl)) {
      if (name !== 'hp:tr') continue
      for (const { name: cellName, node: tc } of elements(tr)) {
        if (cellName !== 'hp:tc') continue
        ctx.tableMarginHwp = tableMarginHwp
        cells.push(await cellInfo(tc, rowIndex, ctx))
      }
      rowIndex++
    }
    const rows = Math.max(rowCount, ...cells.map((cell) => cell.row + cell.rowSpan))
    const cols = Math.max(colCount, ...cells.map((cell) => cell.col + cell.colSpan))
    const colHwp = columnWidths(cells, cols, intAttr(firstChild(tbl, 'hp:sz'), 'width', 0))
    const colTw = colHwp.map((width) => Math.max(1, hwpToTwips(width)))

    // occupancy grid → origin cell per slot
    const grid: Array<Array<CellInfo | null>> = Array.from({ length: rows }, () =>
      Array.from({ length: cols }, () => null),
    )
    for (const cell of cells) {
      for (let r = cell.row; r < Math.min(rows, cell.row + cell.rowSpan); r++) {
        for (let c = cell.col; c < Math.min(cols, cell.col + cell.colSpan); c++) {
          if (!grid[r][c]) grid[r][c] = cell
        }
      }
    }
    // row heights: the smallest single-row cell height in each row
    const rowHeights = Array.from({ length: rows }, (_, r) => {
      const heights = cells
        .filter((cell) => cell.row === r && cell.rowSpan === 1 && cell.heightHwp > 0)
        .map((cell) => cell.heightHwp)
      const declared = heights.length > 0 ? Math.min(...heights) : 0
      // a picture set over a cell still stretches its row to the picture height
      const over = cells
        .filter((cell) => cell.row === r && cell.rowSpan === 1)
        .reduce((max, cell) => Math.max(max, cell.overHwp), 0)
      return Math.max(declared, over)
    })

    const inMargin = firstChild(tbl, 'hp:inMargin')
    const tableMargin = {
      top: hwpToTwips(intAttr(inMargin, 'top')),
      bottom: hwpToTwips(intAttr(inMargin, 'bottom')),
    }
    const defaultMar = inMargin
      ? `<w:tblCellMar><w:top w:w="${hwpToTwips(intAttr(inMargin, 'top'))}" w:type="dxa"/>` +
        `<w:left w:w="${hwpToTwips(intAttr(inMargin, 'left'))}" w:type="dxa"/>` +
        `<w:bottom w:w="${hwpToTwips(intAttr(inMargin, 'bottom'))}" w:type="dxa"/>` +
        `<w:right w:w="${hwpToTwips(intAttr(inMargin, 'right'))}" w:type="dxa"/></w:tblCellMar>`
      : ''
    const parts: string[] = [
      `<w:tbl><w:tblPr><w:tblW w:w="${colTw.reduce((a, b) => a + b, 0)}" w:type="dxa"/>` +
        `${tablePlacementXml(tbl, hostAlign)}<w:tblLayout w:type="fixed"/>${defaultMar}<w:tblLook w:val="0000"/></w:tblPr>`,
      `<w:tblGrid>${colTw.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>`,
    ]
    // Hangul never splits a table set like a character across pages: one that
    // fits moves to the next page whole (lawgo-17's second-page form), so its
    // rows are kept together and with the next row
    const whole =
      ctx.tableDepth === 1 &&
      attr(firstChild(tbl, 'hp:pos'), 'treatAsChar') === '1' &&
      hwpToTwips(intAttr(firstChild(tbl, 'hp:sz'), 'height', 0)) <= bodyHeightTwips(ctx)
    for (let r = 0; r < rows; r++) {
      let contentTw = rowContentHeight(r, rowHeights[r], cells, tableMargin)
      // a row Hangul split across pages declares the sum of its pieces; as a
      // Word floor it would pad the last page and spill onto a blank one
      if (contentTw > bodyHeightTwips(ctx)) contentTw = 0
      const rowPr =
        (contentTw > 0 ? `<w:trHeight w:val="${contentTw}" w:hRule="atLeast"/>` : '') +
        (whole ? '<w:cantSplit/>' : '')
      const trPr = rowPr ? `<w:trPr>${rowPr}</w:trPr>` : ''
      // Word chains a row to the next by its first cell's first paragraph
      const keep = whole && r < rows - 1
      parts.push(`<w:tr>${trPr}`)
      for (let c = 0; c < cols; c++) {
        const cell = grid[r][c]
        if (!cell) {
          const empty = keep && c === 0 ? withKeepNext('<w:p/>') : '<w:p/>'
          parts.push(
            `<w:tc><w:tcPr><w:tcW w:w="${colTw[c]}" w:type="dxa"/></w:tcPr>${empty}</w:tc>`,
          )
          continue
        }
        if (c !== cell.col) continue // horizontal span shadow: gridSpan covers it
        const widthTw = colTw.slice(cell.col, cell.col + cell.colSpan).reduce((a, b) => a + b, 0)
        const tcPr: string[] = [`<w:tcW w:w="${widthTw}" w:type="dxa"/>`]
        if (cell.colSpan > 1) tcPr.push(`<w:gridSpan w:val="${cell.colSpan}"/>`)
        const origin = cell.row === r
        if (cell.rowSpan > 1) tcPr.push(origin ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>')
        const borders = bordersXml(cell.border)
        if (borders) tcPr.push(borders)
        if (cell.border?.fill) {
          tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cell.border.fill}"/>`)
        }
        if (cell.margin) {
          tcPr.push(
            `<w:tcMar><w:top w:w="${cell.margin.top}" w:type="dxa"/><w:left w:w="${cell.margin.left}" w:type="dxa"/>` +
              `<w:bottom w:w="${cell.margin.bottom}" w:type="dxa"/><w:right w:w="${cell.margin.right}" w:type="dxa"/></w:tcMar>`,
          )
        }
        tcPr.push(`<w:vAlign w:val="${cell.vAlign}"/>`)
        let content = origin ? cell.contentXml : '<w:p/>'
        if (keep && c === 0) content = withKeepNext(content)
        parts.push(`<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${content}</w:tc>`)
      }
      parts.push('</w:tr>')
    }
    parts.push('</w:tbl>')
    return parts.join('')
  } finally {
    ctx.tableDepth--
  }
}

/** the first paragraph of a cell's content with keep-with-next set (CT_PPr: after pStyle) */
function withKeepNext(xml: string): string {
  const m = /<w:p\/>|<w:p>(<w:pPr>(?:<w:pStyle [^>]*\/>)?)?/.exec(xml)
  if (!m) return xml
  const pPr = /^<w:p><w:pPr>(.*?)<\/w:pPr>/.exec(xml.slice(m.index))
  if (pPr && pPr[1].includes('<w:keepNext')) return xml
  let next: string
  if (m[0] === '<w:p/>') next = '<w:p><w:pPr><w:keepNext/></w:pPr></w:p>'
  else if (m[1]) next = `<w:p>${m[1]}<w:keepNext/>`
  else next = '<w:p><w:pPr><w:keepNext/></w:pPr>'
  return xml.slice(0, m.index) + next + xml.slice(m.index + m[0].length)
}

/**
 * Where the table sits across the column: a table set like a character
 * (treatAsChar) follows its host paragraph's alignment, a placed one its
 * horizontal alignment and offset. Vertical placement is the flow position.
 */
function tablePlacementXml(tbl: XNode, hostAlign?: string): string {
  const pos = firstChild(tbl, 'hp:pos')
  if (attr(pos, 'treatAsChar') === '1') {
    return hostAlign === 'center' || hostAlign === 'right' ? `<w:jc w:val="${hostAlign}"/>` : ''
  }
  const horz = attr(pos, 'horzAlign')
  if (attr(pos, 'horzRelTo') === 'PAGE') return ''
  if (horz === 'CENTER') return '<w:jc w:val="center"/>'
  if (horz === 'RIGHT') return '<w:jc w:val="right"/>'
  const offset = hwpToTwips(intAttr(pos, 'horzOffset', 0))
  // horzOffset is an unsigned 32-bit field: wrapped values are negative offsets
  return offset > 0 && offset < 20000 ? `<w:tblInd w:w="${offset}" w:type="dxa"/>` : ''
}

/** the page number's run properties: the document's base character shape (id 0) */
function pageNumberRunPr(ctx: Ctx): string {
  const char = ctx.header.charShape('0')
  const face = char.fontHangul ?? char.fontLatin
  const size = char.sizeHalfPoints ?? 20
  const fonts = face
    ? `<w:rFonts w:ascii="${escapeXml(char.fontLatin ?? face)}" w:eastAsia="${escapeXml(face)}" w:hAnsi="${escapeXml(char.fontLatin ?? face)}"/>`
    : ''
  return `<w:rPr>${fonts}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>`
}

/** usable page height (A4 with Hangul's default margins until a secPr is read) */
function bodyHeightTwips(ctx: Ctx): number {
  const section = ctx.current ?? ctx.section
  if (!section) return 13606
  return section.pageHeight - section.marginTop - section.marginBottom
}

/** rendered border width in twips (w:sz is eighths of a point; a double line is three of them) */
const borderTwips = (side: BorderSide | null): number =>
  side ? side.sz * 2.5 * (COMPOUND.has(side.val) ? 3 : 1) : 0

/**
 * Hangul's cell height is the whole row: content, the vertical cell margins
 * and the gridline. Word's atLeast w:trHeight floors the content alone and
 * lays the row's largest top and bottom margin plus one gridline on top, so
 * the declared height drops those, or every form row grows by ~4pt and a
 * one-page form spills onto a second page.
 */
function rowContentHeight(
  r: number,
  heightHwp: number,
  cells: CellInfo[],
  tableMargin: { top: number; bottom: number },
): number {
  if (heightHwp <= 0) return 0
  // the row's own cells: a cell merged down from a row above sits in that
  // row's w:tr (one rowspan cell in GenOffice), so its margins and rules do
  // not add to this row (ftc-57's discount table lost 1.5pt a row to the
  // double rule of the cell spanning it)
  const starting = cells.filter((cell) => cell.row === r)
  const inRow =
    starting.length > 0
      ? starting
      : cells.filter((cell) => cell.row <= r && r < cell.row + cell.rowSpan)
  let top = 0
  let bottom = 0
  let gridline = 0
  for (const cell of inRow) {
    top = Math.max(top, cell.margin?.top ?? tableMargin.top)
    bottom = Math.max(bottom, cell.margin?.bottom ?? tableMargin.bottom)
    gridline = Math.max(
      gridline,
      borderTwips(cell.border?.top ?? null),
      borderTwips(cell.border?.bottom ?? null),
    )
  }
  return Math.max(1, Math.round(hwpToTwips(heightHwp) - top - bottom - gridline))
}

async function cellInfo(tc: XNode, rowIndex: number, ctx: Ctx): Promise<CellInfo> {
  const addr = firstChild(tc, 'hp:cellAddr')
  const span = firstChild(tc, 'hp:cellSpan')
  const size = firstChild(tc, 'hp:cellSz')
  const subList = firstChild(tc, 'hp:subList')
  // a cell's own hp:cellMargin only applies with hasMargin="1"; otherwise the
  // table's hp:inMargin (written once as tblCellMar) is in force
  const marginNode = attr(tc, 'hasMargin') === '1' ? firstChild(tc, 'hp:cellMargin') : null
  const vAlign = V_ALIGN[attr(subList, 'vertAlign') ?? ''] ?? 'top'
  const marginsHwp = marginNode
    ? intAttr(marginNode, 'top') + intAttr(marginNode, 'bottom')
    : ctx.tableMarginHwp.top + ctx.tableMarginHwp.bottom
  const blocks: string[] = []
  const outerCell = ctx.cell
  const outerMargin = ctx.tableMarginHwp
  const sidesHwp = marginNode
    ? intAttr(marginNode, 'left') + intAttr(marginNode, 'right')
    : ctx.tableMarginHwp.left + ctx.tableMarginHwp.right
  const cell = {
    contentHwp: intAttr(size, 'height', 0) - marginsHwp,
    textWidthHwp: intAttr(size, 'width', 0) - sidesHwp,
    marginsHwp,
    vAlign,
    overHwp: 0,
  }
  ctx.cell = cell
  try {
    if (subList) {
      const paragraphs = elements(subList).filter((child) => child.name === 'hp:p')
      for (const [index, { node }] of paragraphs.entries()) {
        blocks.push(...(await paragraphBlocks(node, ctx, index === paragraphs.length - 1)))
      }
    }
  } finally {
    ctx.cell = outerCell
    ctx.tableMarginHwp = outerMargin
  }
  // OOXML: a cell's last child must be a paragraph
  if (blocks.length === 0 || /^<w:tbl[\s>]/.test(blocks[blocks.length - 1])) blocks.push('<w:p/>')
  return {
    row: addr ? intAttr(addr, 'rowAddr', rowIndex) : rowIndex,
    col: intAttr(addr, 'colAddr', 0),
    rowSpan: Math.max(1, intAttr(span, 'rowSpan', 1)),
    colSpan: Math.max(1, intAttr(span, 'colSpan', 1)),
    widthHwp: intAttr(size, 'width', 0),
    heightHwp: intAttr(size, 'height', 0),
    border: ctx.header.borderFill(attr(tc, 'borderFillIDRef')),
    vAlign,
    overHwp: cell.overHwp,
    margin: marginNode
      ? {
          left: hwpToTwips(intAttr(marginNode, 'left')),
          right: hwpToTwips(intAttr(marginNode, 'right')),
          top: hwpToTwips(intAttr(marginNode, 'top')),
          bottom: hwpToTwips(intAttr(marginNode, 'bottom')),
        }
      : null,
    contentXml: blocks.join(''),
  }
}

/**
 * Column widths from the cells. Hangul draws a table on one grid whatever its
 * cells claim: the right edge is the table's own width, the boundaries come
 * from the narrowest cells, and a wider cell that disagrees is stretched to
 * the grid (lawgo-7 draws a 44790 heading over 2830 + 18839 + 23791 columns
 * across all 45460; lawgo-17 draws 22 cells of 38933 out to the 50897 edge).
 * So the boundaries are settled from the edges inward, narrow cells first,
 * each on the value most of its cells agree on; any the cells leave open
 * split the remaining width evenly.
 */
function columnWidths(cells: CellInfo[], cols: number, tableWidth: number): number[] {
  const x: Array<number | null> = Array.from({ length: cols + 1 }, () => null)
  x[0] = 0
  if (tableWidth > 0) x[cols] = tableWidth
  const constraints = cells
    .filter((cell) => cell.widthHwp > 0 && cell.col < cols)
    .map((cell) => ({ a: cell.col, b: Math.min(cols, cell.col + cell.colSpan), w: cell.widthHwp }))
    .filter((c) => c.a < c.b)
  const widest = constraints.reduce((max, c) => Math.max(max, c.b - c.a), 0)
  for (let span = 1; span <= widest; span++) {
    for (let pass = 0; pass <= cols; pass++) {
      // the votes for each open boundary from cells of this span or narrower
      const votes = new Map<number, Map<number, number>>()
      const vote = (k: number, value: number): void => {
        const tally = votes.get(k) ?? new Map<number, number>()
        tally.set(value, (tally.get(value) ?? 0) + 1)
        votes.set(k, tally)
      }
      for (const c of constraints) {
        if (c.b - c.a > span) continue
        if (x[c.a] !== null && x[c.b] === null) vote(c.b, x[c.a]! + c.w)
        else if (x[c.b] !== null && x[c.a] === null) vote(c.a, x[c.b]! - c.w)
      }
      if (votes.size === 0) break
      for (const [k, tally] of votes) {
        let best = 0
        let count = -1
        for (const [value, n] of tally) {
          if (n > count) {
            best = value
            count = n
          }
        }
        x[k] = best
      }
    }
  }
  if (x[cols] === null) x[cols] = 48000
  // boundaries no cell pins: spread evenly between the known ones
  for (let c = 1; c < cols; c++) {
    if (x[c] !== null) continue
    let next = c + 1
    while (x[next] === null) next++
    const prev = x[c - 1]!
    x[c] = prev + (x[next]! - prev) / (next - c + 1)
  }
  // a boundary placed past its neighbour (cells that contradict each other
  // outright) leaves that column a sliver instead of a negative width
  return Array.from({ length: cols }, (_, c) => Math.max(1, Math.round(x[c + 1]! - x[c]!)))
}

function bordersXml(border: BorderFill | null): string {
  const side = (tag: string, value: BorderFill['left']): string =>
    value
      ? `<w:${tag} w:val="${value.val}" w:sz="${value.sz}" w:space="0" w:color="${value.color}"/>`
      : `<w:${tag} w:val="nil"/>`
  if (!border) return ''
  return `<w:tcBorders>${side('top', border.top)}${side('left', border.left)}${side('bottom', border.bottom)}${side('right', border.right)}</w:tcBorders>`
}
