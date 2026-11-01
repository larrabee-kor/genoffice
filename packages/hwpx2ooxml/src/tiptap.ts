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
  type Block,
  type CellBorders,
  type ParaAlign,
  type Run,
  type TableCell,
  type TableModel,
  type TableParagraph,
} from '../../docx-engine/src/index'
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

export interface HwpToTiptapResult extends HwpxToDocxResult {
  /** Tiptap document JSON (`editor.commands.setContent(doc)`) */
  doc: TiptapNode
  /** docx blocks with no Tiptap counterpart (section/body passthrough parts) */
  skipped: number
}

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
  const { doc, skipped } = await docxToTiptap(converted.docx)
  return { ...converted, doc, skipped }
}

/** A .docx (as written by the filter, or any other) → Tiptap JSON */
export async function docxToTiptap(
  docx: Uint8Array,
): Promise<{ doc: TiptapNode; skipped: number }> {
  const parsed = await parseDocx(docx)
  const content: TiptapNode[] = []
  let skipped = 0
  for (const block of parsed.blocks) {
    const node = blockNode(block)
    if (node) content.push(node)
    else skipped++
  }
  // a Tiptap document needs at least one block
  if (content.length === 0) content.push({ type: 'paragraph' })
  return { doc: { type: 'doc', content }, skipped }
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

function paragraphNode(format: { align?: ParaAlign }, content: TiptapNode[]): TiptapNode {
  const align = format.align ? ALIGN[format.align] : undefined
  return {
    type: 'paragraph',
    ...(align && align !== 'left' ? { attrs: { textAlign: align } } : {}),
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
  if (font) style.fontFamily = font
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
