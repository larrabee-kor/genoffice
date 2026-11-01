import { convertHwpToDocx, HwpxConvertError } from '@genoffice/hwpx2ooxml'
import { CliError, EXIT } from '../result'

/** .hwp/.hwpx → .docx through the local import filter (same engine as the app's open route; routed by content). */
export async function convertHwpx(input: Uint8Array): Promise<{
  bytes: Uint8Array
  detail: Record<string, unknown>
}> {
  try {
    const result = await convertHwpToDocx(input)
    return {
      bytes: result.docx,
      detail: {
        blocks: result.blocks,
        dropped: result.dropped,
        ...(result.via ? { via: result.via } : {}),
      },
    }
  } catch (err) {
    if (err instanceof HwpxConvertError) {
      const message = err.message
      throw new CliError(EXIT.conversion, message, { reason: err.code })
    }
    throw err
  }
}
