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
import { HeadingCounter, readHeadings } from './numbering'
import { withHangulLayoutSettings } from './settings'
import {
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
type Segment = { kind: 'run'; run: Run } | { kind: 'picture'; xml: string }

interface Ctx {
  header: HeaderTables
  pack: HwpxPackage
  drops: Drops
  pictures: Record<string, { base64: string; mime: NewImage['mime'] }>
  pictureSeq: number
  /** resolved BinData → token, so a logo repeated on every page lands once */
  pictureCache: Map<string, { token: string; widthPx: number; heightPx: number } | null>
  section: SectionSettings | null
  /** nesting depth of tables being written (cells inside cells) */
  tableDepth: number
  headings: HeadingCounter
  /** numbering the section's outline levels use (hp:secPr outlineShapeIDRef) */
  outlineId: string
}

const MAX_TABLE_DEPTH = 8

export async function convertHwpxToDocx(hwpx: Uint8Array): Promise<HwpxToDocxResult> {
  const pack = await loadHwpxPackage(hwpx)
  const ctx: Ctx = {
    header: readHeader(pack.header),
    headings: new HeadingCounter(readHeadings(pack.header)),
    outlineId: '1',
    pack,
    drops: new Drops(),
    pictures: {},
    pictureSeq: 0,
    pictureCache: new Map(),
    section: null,
    tableDepth: 0,
  }

  const body: string[] = []
  for (const [index, root] of pack.sections.entries()) {
    const sec = root.find((node) => nodeName(node) === 'hs:sec')
    if (!sec) continue
    if (index > 0) {
      // trailing empty lines of the previous section only pad the page before
      // the break; where a wider fallback font pushes them over, they would
      // leave a blank page behind
      while (body.length > 0 && isEmptyParagraph(body[body.length - 1])) body.pop()
    }
    const before = body.length
    for (const { name, node } of elements(sec)) {
      if (name === 'hp:p') body.push(...(await paragraphBlocks(node, ctx)))
    }
    // a later section starts on its own page in nearly every real document
    // (Hangul's section break); without per-section sectPr this keeps the flow
    if (index > 0 && body.length > before) body[before] = withPageBreakBefore(body[before])
  }
  if (pack.sections.length > 1) ctx.drops.add('section-settings-merged', pack.sections.length - 1)

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
    ...(ctx.section ? { section: ctx.section } : {}),
    ...(Object.keys(ctx.pictures).length > 0 ? { newPictures: ctx.pictures } : {}),
  })
  // original font names stay on the runs; the table tells Word/GenOffice how to substitute
  const docx = await withHangulLayoutSettings(await withFontTable(saved))
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

  const flush = (force: boolean, final = false): void => {
    const hasContent = segments.some(
      (segment) => segment.kind === 'picture' || segment.run.text.length > 0,
    )
    if (!hasContent && !force) {
      segments = []
      return
    }
    const paraFormat = emittedAny ? { ...format, pageBreakBefore: undefined } : { ...format }
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
      const hasPicture = segments.some((segment) => segment.kind === 'picture')
      paraFormat.lineRule = hasPicture ? 'atLeast' : 'exact'
      paraFormat.lineRawTwips = Math.round(largest * 10 * (shape.linePercent / 100))
      delete paraFormat.lineSpacing
      // the gap under a cell's last line is not laid out (the row ends at the
      // glyphs), so its lines share the height Hangul recorded for them
      const cellEnd = final && lastInCell && !hadTable && !hasPicture ? lastLinesHeight(p) : 0
      if (cellEnd > 0) paraFormat.lineRawTwips = cellEnd
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
          if (head && text) appendText(segments, `${head}\u2002`, char)
          if (text) head = undefined
          appendText(segments, text, char)
          break
        }
        case 'hp:secPr':
          if (!ctx.section) ctx.section = sectionSettings(node)
          ctx.outlineId = attr(node, 'outlineShapeIDRef') ?? ctx.outlineId
          break
        case 'hp:ctrl':
          dropControls(node, ctx)
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
          if (xml) segments.push({ kind: 'picture', xml })
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

function appendText(segments: Segment[], text: string, char: CharShape): void {
  if (!text) return
  for (const piece of textPieces(text, char)) {
    const run = runOf(piece.text, char, piece.kind)
    const last = segments[segments.length - 1]
    // merge with the previous run when the formatting matches (fewer w:r)
    if (last && last.kind === 'run' && sameFormat(last.run, run)) last.run.text += piece.text
    else segments.push({ kind: 'run', run })
  }
}

type PieceKind = 'wide' | 'other'

/**
 * Text split where Word applies letter spacing differently from the rest of
 * the run: under the balanced-width setting (settings.ts) Word doubles
 * w:spacing on double-byte glyphs except hangul in a few faces (runOf).
 * Glyph widths themselves are the compat faces' business (tools/hwp-compat),
 * not the converter's.
 */
function textPieces(text: string, char: CharShape): Array<{ text: string; kind: PieceKind }> {
  // only a face Word keeps single gives double-byte stretches a different value
  const single = keepsHangulKanaSingle(char.fontHangul ?? char.fontLatin)
  if (!char.spacingPct || !char.sizeHalfPoints || !single) return [{ text, kind: 'other' }]
  const pieces: Array<{ text: string; kind: PieceKind }> = []
  for (const ch of text) {
    const kind: PieceKind = isDoubleByte(ch.codePointAt(0)!) ? 'wide' : 'other'
    const last = pieces[pieces.length - 1]
    if (last?.kind === kind) last.text += ch
    else pieces.push({ text: ch, kind })
  }
  return pieces
}

/** what Word counts as double-byte for balanced spacing (docx-engine doubledGlyphFraction) */
function isDoubleByte(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x9fff) ||
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

function runOf(text: string, char: CharShape, kind: PieceKind = 'other'): Run {
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
    // Hangul letter spacing is a percent of each glyph's own advance (rhwp
    // text_measurement: 휴먼명조 '*' 0.5 em at -9% spaces 0.455 em), OOXML's a
    // fixed twip amount per glyph. Full-width glyphs take the whole percent of
    // the size, half-width ones (Latin, digits, spaces) about half of it. Word
    // doubles w:spacing on double-byte stretches (textPieces), so both write
    // half the percent, except hangul in the faces Word keeps single (Malgun
    // Gothic, NanumGothic…), which needs it whole
    const twips = (char.spacingPct / 100) * char.sizeHalfPoints * 10
    const single = kind === 'wide' && keepsHangulKanaSingle(run.eastAsiaFont) && isHangulText(text)
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
        out += '　'
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

function withPageBreakBefore(xml: string): string {
  if (!/^<w:p[\s>]/.test(xml) || xml.includes('<w:pageBreakBefore/>')) return xml
  if (xml.includes('<w:pPr>')) return xml.replace('<w:pPr>', '<w:pPr><w:pageBreakBefore/>')
  return xml.replace(/^<w:p\b[^>]*>/, (open) => `${open}<w:pPr><w:pageBreakBefore/></w:pPr>`)
}

function dropControls(ctrl: XNode, ctx: Ctx): void {
  for (const { name } of elements(ctrl)) {
    switch (name) {
      case 'hp:header':
        ctx.drops.add('headers')
        break
      case 'hp:footer':
        ctx.drops.add('footers')
        break
      case 'hp:footNote':
        ctx.drops.add('footnotes')
        break
      case 'hp:endNote':
        ctx.drops.add('endnotes')
        break
      case 'hp:pageNum':
        ctx.drops.add('page-numbers')
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

async function pictureRunXml(pic: XNode, ctx: Ctx): Promise<string | null> {
  const img = firstDescendant(pic, 'hc:img')
  const id = attr(img, 'binaryItemIDRef')
  if (!id) {
    ctx.drops.add('images')
    return null
  }
  const pos = firstChild(pic, 'hp:pos')
  if (pos && attr(pos, 'treatAsChar') !== '1') ctx.drops.add('floating-images-inlined')

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
  return `<w:r>${drawing}</w:r>`
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
    for (const { name, node: tr } of elements(tbl)) {
      if (name !== 'hp:tr') continue
      for (const { name: cellName, node: tc } of elements(tr)) {
        if (cellName !== 'hp:tc') continue
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
      return heights.length > 0 ? Math.min(...heights) : 0
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
    for (let r = 0; r < rows; r++) {
      let contentTw = rowContentHeight(r, rowHeights[r], cells, tableMargin)
      // a row Hangul split across pages declares the sum of its pieces; as a
      // Word floor it would pad the last page and spill onto a blank one
      if (contentTw > bodyHeightTwips(ctx)) contentTw = 0
      const trPr =
        contentTw > 0 ? `<w:trPr><w:trHeight w:val="${contentTw}" w:hRule="atLeast"/></w:trPr>` : ''
      parts.push(`<w:tr>${trPr}`)
      for (let c = 0; c < cols; c++) {
        const cell = grid[r][c]
        if (!cell) {
          parts.push(`<w:tc><w:tcPr><w:tcW w:w="${colTw[c]}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`)
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
        parts.push(
          `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${origin ? cell.contentXml : '<w:p/>'}</w:tc>`,
        )
      }
      parts.push('</w:tr>')
    }
    parts.push('</w:tbl>')
    return parts.join('')
  } finally {
    ctx.tableDepth--
  }
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

/** usable page height (A4 with Hangul's default margins until a secPr is read) */
function bodyHeightTwips(ctx: Ctx): number {
  const section = ctx.section
  if (!section) return 13606
  return section.pageHeight - section.marginTop - section.marginBottom
}

/** border width in twips (w:sz is eighths of a point) */
const borderTwips = (side: BorderSide | null): number => (side ? side.sz * 2.5 : 0)

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
  const inRow = cells.filter((cell) => cell.row <= r && r < cell.row + cell.rowSpan)
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
  const blocks: string[] = []
  if (subList) {
    const paragraphs = elements(subList).filter((child) => child.name === 'hp:p')
    for (const [index, { node }] of paragraphs.entries()) {
      blocks.push(...(await paragraphBlocks(node, ctx, index === paragraphs.length - 1)))
    }
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
    vAlign: V_ALIGN[attr(subList, 'vertAlign') ?? ''] ?? 'top',
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
 * Column widths from the cells: single-column cells pin their column; spanned
 * cells then share their remaining width among the columns still unknown;
 * whatever is left splits the table width evenly.
 */
function columnWidths(cells: CellInfo[], cols: number, tableWidth: number): number[] {
  const widths: Array<number | null> = Array.from({ length: cols }, () => null)
  for (const cell of cells) {
    if (cell.colSpan === 1 && cell.widthHwp > 0 && cell.col < cols && widths[cell.col] === null) {
      widths[cell.col] = cell.widthHwp
    }
  }
  for (let pass = 0; pass < 4; pass++) {
    let changed = false
    for (const cell of [...cells].sort((a, b) => a.colSpan - b.colSpan)) {
      if (cell.colSpan === 1 || cell.widthHwp <= 0) continue
      const range = widths.slice(cell.col, cell.col + cell.colSpan)
      const unknown = range.filter((width) => width === null).length
      if (unknown === 0) continue
      const known = range.reduce<number>((sum, width) => sum + (width ?? 0), 0)
      const share = Math.max(100, Math.round((cell.widthHwp - known) / unknown))
      for (let c = cell.col; c < Math.min(cols, cell.col + cell.colSpan); c++) {
        if (widths[c] === null) widths[c] = share
      }
      changed = true
    }
    if (!changed) break
  }
  const known = widths.reduce<number>((sum, width) => sum + (width ?? 0), 0)
  const unknown = widths.filter((width) => width === null).length
  const fallback =
    unknown > 0 ? Math.max(500, Math.round(((tableWidth || 48000) - known) / unknown)) : 0
  return widths.map((width) => width ?? fallback)
}

function bordersXml(border: BorderFill | null): string {
  const side = (tag: string, value: BorderFill['left']): string =>
    value
      ? `<w:${tag} w:val="${value.val}" w:sz="${value.sz}" w:space="0" w:color="${value.color}"/>`
      : `<w:${tag} w:val="nil"/>`
  if (!border) return ''
  return `<w:tcBorders>${side('top', border.top)}${side('left', border.left)}${side('bottom', border.bottom)}${side('right', border.right)}</w:tcBorders>`
}
