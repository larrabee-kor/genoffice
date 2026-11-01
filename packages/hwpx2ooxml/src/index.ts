/**
 * @genoffice/hwpx2ooxml — local HWPX (Hancom OWPML, KS X 6101) import
 * filter: .hwpx → .docx. Pure-function package: bytes in, bytes out; nothing
 * here touches Electron or the filesystem. Like the ODF filters, a construct
 * that cannot be converted provably-correct is counted in `dropped` instead of
 * approximated into wrong markup.
 */

export { convertHwpxToDocx, HwpxConvertError, type HwpxToDocxResult } from './hwpx-docx'
export { isHwpBinary, MAX_HWPX_PART_BYTES, type HwpxConvertErrorCode } from './package'
export { convertHwpToDocx, hwpToHwpx } from './hwp'
export { bmpToPng } from './bmp'
export { classifyFont, type FontClass, type FontHint } from './fonts'
export {
  convertHwpToTiptap,
  docxToTiptap,
  type HwpToTiptapResult,
  type TiptapMark,
  type TiptapNode,
} from './tiptap'
