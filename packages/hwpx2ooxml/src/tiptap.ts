/**
 * Tiptap entry: the same HWP/HWPX → .docx filter GenOffice opens with, read
 * back through docx-engine and emitted as standard Tiptap JSON (StarterKit,
 * TextStyle/Color/FontFamily/FontSize, Highlight, TextAlign, Subscript/
 * Superscript, Link, Image, Table). There is no second HWP reader: anything
 * the docx filter learns, a Tiptap editor gets with it, and both stay
 * comparable on the same document.
 *
 * Images are emitted inside paragraphs, so the Image extension needs
 * `inline: true`. Attributes a stock extension does not declare are dropped
 * by the editor schema, never misread.
 */

import {
  parseDocx,
  readSections,
  type Block,
  type CellBorders,
  type ParaAlign,
  type ParaFormat,
  type Run,
  type TableCell,
  type TableModel,
  type TableParagraph,
} from '../../docx-engine/src/index'
import { HWP_COMPAT_CLASS, HWP_COMPAT_FACES } from './compat-faces'
import { classifyFont } from './fonts'
import { convertHwpToDocx } from './hwp'
import type { HwpxToDocxResult } from './hwpx-docx'

export interface TiptapMark {
  type: string
  attrs?: Record<string, unknown>
}

export interface TiptapNode {
  type: string
  attrs?: Record<string, unknown>
  content?: TiptapNode[]
  marks?: TiptapMark[]
  text?: string
}

/** the first section's sheet, in points: a host lays the editor out on it */
export interface TiptapPage {
  widthPt: number
  heightPt: number
  marginTopPt: number
  marginRightPt: number
  marginBottomPt: number
  marginLeftPt: number
}

export interface TiptapDocument {
  /** Tiptap document JSON (`editor.commands.setContent(doc)`) */
  doc: TiptapNode
  /** docx blocks with no Tiptap counterpart (section/body passthrough parts) */
  skipped: number
  page?: TiptapPage
}

export type HwpToTiptapResult = HwpxToDocxResult & TiptapDocument

const TWIPS_PER_PX = 15

const ALIGN: Partial<Record<ParaAlign, string>> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'justify',
  // Hangul's distributed alignment has no CSS counterpart beyond justify
  distribute: 'justify',
}

/** Any Hancom word-processor file → Tiptap JSON, through the docx filter */
export async function convertHwpToTiptap(bytes: Uint8Array): Promise<HwpToTiptapResult> {
  const converted = await convertHwpToDocx(bytes)
  return { ...converted, ...(await docxToTiptap(converted.docx)) }
}

/** A .docx (as written by the filter, or any other) → Tiptap JSON */
export async function docxToTiptap(docx: Uint8Array): Promise<TiptapDocument> {
  const parsed = await parseDocx(docx)
  // an HWP-origin document names Hancom faces: chain their compat faces
  // (apps/docs/src/renderer/fonts/hwp-compat.css) so a host that loads them
  // breaks lines with Hangul's widths, as GenOffice does
  hwpFaces = parsed.hwpSource === true
  const content: TiptapNode[] = []
  let skipped = 0
  for (const block of parsed.blocks) {
    const node = blockNode(block)
    if (node) content.push(node)
    else skipped++
  }
  // a Tiptap document needs at least one block
  if (content.length === 0) content.push({ type: 'paragraph' })
  const sheet = readSections(parsed)[0]?.settings
  const pt = (twips: number) => Math.round(twips / 2) / 10
  const page = sheet
    ? {
        widthPt: pt(sheet.pageWidth),
        heightPt: pt(sheet.pageHeight),
        marginTopPt: pt(sheet.marginTop),
        marginRightPt: pt(sheet.marginRight),
        marginBottomPt: pt(sheet.marginBottom),
        marginLeftPt: pt(sheet.marginLeft),
      }
    : undefined
  return { doc: { type: 'doc', content }, skipped, ...(page ? { page } : {}) }
}

let hwpFaces = false

/** CSS font-family for a run: the face, then its compat face in an HWP-origin document */
function fontFamilyCss(face: string): string {
  const quote = (name: string) => `'${name.replace(/'/g, '')}'`
  const compat = hwpFaces ? HWP_COMPAT_FACES[face.normalize('NFC')] : undefined
  if (!compat) return face
  const generic =
    HWP_COMPAT_CLASS[compat[0]] ??
    ({ serif: 'serif', sans: 'sans-serif', mono: 'monospace' } as const)[classifyFont(face).cls]
  return [quote(face), ...compat.map(quote), generic].join(', ')
}

/** paragraph geometry as inline CSS for a `style` attribute (stock Paragraph drops it) */
function paragraphCss(format: ParaFormat): string {
  const pt = (twips: number) => `${Math.round(twips / 2) / 10}pt`
  const css: string[] = []
  if (format.lineRawTwips && (format.lineRule === 'exact' || format.lineRule === 'atLeast')) {
    css.push(`line-height:${pt(format.lineRawTwips)}`)
  }
  if (format.spaceBefore) css.push(`margin-top:${pt(format.spaceBefore)}`)
  if (format.spaceAfter) css.push(`margin-bottom:${pt(format.spaceAfter)}`)
  if (format.indentLeft) css.push(`padding-left:${pt(format.indentLeft)}`)
  if (format.indentRight) css.push(`padding-right:${pt(format.indentRight)}`)
  if (format.indentFirstLine) css.push(`text-indent:${pt(format.indentFirstLine)}`)
  return css.join(';')
}

function blockNode(block: Block): TiptapNode | null {
  switch (block.type) {
    case 'table':
      return block.table ? tableNode(block.table) : null
    case 'image':
      return block.imageDataUrl
        ? paragraphNode({ align: block.imageAlign }, [
            imageNode(block.imageDataUrl, block.imageWidthPx, block.imageHeightPx),
          ])
        : null
    case 'heading': {
      const node = paragraphNode(block.format ?? {}, inlineNodes(block.runs ?? []))
      return {
        ...node,
        type: 'heading',
        attrs: { ...node.attrs, level: Math.min(6, Math.max(1, block.level ?? 1)) },
      }
    }
    case 'paragraph':
    case 'listItem':
      return paragraphNode(block.format ?? {}, inlineNodes(block.runs ?? []))
    default:
      return null
  }
}

function paragraphNode(format: ParaFormat, content: TiptapNode[]): TiptapNode {
  const align = format.align ? ALIGN[format.align] : undefined
  const style = paragraphCss(format)
  const attrs = {
    ...(align && align !== 'left' ? { textAlign: align } : {}),
    ...(style ? { style } : {}),
  }
  return {
    type: 'paragraph',
    ...(Object.keys(attrs).length > 0 ? { attrs } : {}),
    ...(content.length > 0 ? { content } : {}),
  }
}

function imageNode(src: string, width?: number, height?: number): TiptapNode {
  return {
    type: 'image',
    attrs: {
      src,
      ...(width ? { width: Math.round(width) } : {}),
      ...(height ? { height: Math.round(height) } : {}),
    },
  }
}

/** runs → text / hardBreak / image nodes; a line break inside a run is a hard break */
function inlineNodes(runs: Run[]): TiptapNode[] {
  const out: TiptapNode[] = []
  for (const run of runs) {
    if (run.vanish) continue
    if (run.image?.dataUrl) {
      out.push(imageNode(run.image.dataUrl, run.image.widthPx, run.image.heightPx))
    }
    if (!run.text) continue
    const marks = runMarks(run)
    run.text.split('\n').forEach((piece, index) => {
      if (index > 0) out.push({ type: 'hardBreak' })
      if (!piece) return
      // runs the docx splits for Word-only reasons (a space's Latin face) fold back together
      const last = out[out.length - 1]
      if (last?.type === 'text' && JSON.stringify(last.marks ?? []) === JSON.stringify(marks)) {
        last.text += piece
      } else {
        out.push({ type: 'text', text: piece, ...(marks.length > 0 ? { marks } : {}) })
      }
    })
  }
  return out
}

function runMarks(run: Run): TiptapMark[] {
  const marks: TiptapMark[] = []
  if (run.bold) marks.push({ type: 'bold' })
  if (run.italic) marks.push({ type: 'italic' })
  if (run.underline) marks.push({ type: 'underline' })
  if (run.strike || run.dstrike) marks.push({ type: 'strike' })
  if (run.vertAlign === 'superscript') marks.push({ type: 'superscript' })
  if (run.vertAlign === 'subscript') marks.push({ type: 'subscript' })
  if (run.link?.href) marks.push({ type: 'link', attrs: { href: run.link.href } })
  const style: Record<string, string> = {}
  if (run.color && run.color !== 'auto') style.color = `#${run.color}`
  const font = run.eastAsiaFont ?? run.font
  if (font) style.fontFamily = fontFamilyCss(font)
  if (run.sizeHalfPoints) style.fontSize = `${run.sizeHalfPoints / 2}pt`
  if (Object.keys(style).length > 0) marks.push({ type: 'textStyle', attrs: style })
  const highlight = run.shading ?? run.highlight
  if (highlight && /^[0-9a-f]{6}$/i.test(highlight)) {
    marks.push({ type: 'highlight', attrs: { color: `#${highlight}` } })
  }
  return marks
}

// -- tables -------------------------------------------------------------------

/**
 * docx-engine keeps Word's shape: horizontal spans as gridSpan, vertical ones
 * as a restart cell over continue cells. Tiptap wants rowspan on the origin
 * and no cell in the slots it covers.
 */
function tableNode(table: TableModel): TiptapNode {
  const colPx = (table.colWidthsTwips ?? []).map((tw) => Math.max(1, Math.round(tw / TWIPS_PER_PX)))
  // grid column of every cell, then how far each restart reaches down
  const placed = table.rows.map((row) => {
    let col = 0
    return row.map((cell) => {
      const at = col
      col += Math.max(1, cell.colSpan ?? 1)
      return { cell, col: at }
    })
  })
  const rows: TiptapNode[] = placed.map((row, r) => {
    const cells: TiptapNode[] = []
    for (const { cell, col } of row) {
      if (cell.vMerge === 'continue') continue
      let rowspan = 1
      if (cell.vMerge === 'restart') {
        while (
          r + rowspan < placed.length &&
          placed[r + rowspan].some((next) => next.col === col && next.cell.vMerge === 'continue')
        ) {
          rowspan++
        }
      }
      cells.push(cellNode(cell, col, rowspan, colPx))
    }
    return { type: 'tableRow', content: cells.length > 0 ? cells : [emptyCell()] }
  })
  return { type: 'table', content: rows }
}

function cellNode(cell: TableCell, col: number, rowspan: number, colPx: number[]): TiptapNode {
  const colspan = Math.max(1, cell.colSpan ?? 1)
  const widths = colPx.slice(col, col + colspan)
  const content = cellContent(cell)
  return {
    type: 'tableCell',
    attrs: {
      colspan,
      rowspan,
      colwidth: widths.length === colspan ? widths : null,
      ...(cell.fill && /^[0-9a-f]{6}$/i.test(cell.fill)
        ? { backgroundColor: `#${cell.fill}` }
        : {}),
      ...(cell.borders ? { style: borderCss(cell.borders) } : {}),
    },
    content: content.length > 0 ? content : [{ type: 'paragraph' }],
  }
}

const CSS_BORDER: Record<string, string> = {
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  thinThickSmallGap: 'double',
  thickThinSmallGap: 'double',
}

/**
 * Cell borders as inline CSS for a `style` attribute: stock TableCell has no
 * border attribute, so a host that wants form lines declares `style` as a
 * global attribute; without it the editor drops the value.
 */
function borderCss(borders: CellBorders): string {
  const sides = ['top', 'right', 'bottom', 'left'] as const
  return sides
    .map((side) => {
      const line = borders[side]
      if (!line || line.style === 'nil' || line.style === 'none') return `border-${side}:none`
      const px = Math.max(1, Math.round(((line.szEighths ?? 4) / 8) * (96 / 72) * 10) / 10)
      const color = line.color && line.color !== 'auto' ? `#${line.color}` : '#000'
      return `border-${side}:${px}px ${CSS_BORDER[line.style] ?? 'solid'} ${color}`
    })
    .join(';')
}

/** cell paragraphs with nested tables placed before the paragraph they anchor to */
function cellContent(cell: TableCell): TiptapNode[] {
  const paras: TableParagraph[] =
    cell.richParas ?? cell.paras.map((text) => ({ runs: text ? [{ text }] : [] }))
  const nested = cell.nestedTables ?? []
  const anchors = cell.nestedTableAnchors ?? nested.map(() => paras.length)
  const out: TiptapNode[] = []
  for (let i = 0; i <= paras.length; i++) {
    nested.forEach((table, t) => {
      if (Math.min(anchors[t] ?? paras.length, paras.length) === i) out.push(tableNode(table))
    })
    if (i < paras.length) out.push(paragraphNode(paras[i], inlineNodes(paras[i].runs)))
  }
  return out
}

function emptyCell(): TiptapNode {
  return { type: 'tableCell', attrs: { colspan: 1, rowspan: 1 }, content: [{ type: 'paragraph' }] }
}
