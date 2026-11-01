/**
 * word/settings.xml for the converted document: the two East Asian
 * compatibility flags under which Word sets a space touching Hangul half an
 * em wide (w:balanceSingleByteDoubleByteWidth + w:useFELayout; either alone
 * keeps the face's own 0.25-0.36 em space). That is Hangul's own rule: the
 * original PDFs measure 0.5 em per space in Korean faces, and the narrower
 * spaces pulled one or two extra syllables onto every body line.
 * No compatibilityMode is declared, as before the part existed.
 */

import JSZip from 'jszip'
import { addDocumentPart } from './parts'

const SETTINGS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

/** CT_Compat order: balanceSingleByteDoubleByteWidth precedes useFELayout */
const COMPAT = '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/></w:compat>'

export async function withHangulLayoutSettings(docx: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(docx)
  if (zip.file('word/settings.xml')) return docx
  await addDocumentPart(
    zip,
    'settings.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${COMPAT}</w:settings>`,
    SETTINGS_REL,
    SETTINGS_CT,
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
