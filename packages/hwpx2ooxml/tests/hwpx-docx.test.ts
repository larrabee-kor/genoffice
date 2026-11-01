import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../../docx-engine/src/index'
import { bmpToPng, convertHwpxToDocx, HwpxConvertError } from '../src/index'

const NS =
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"' +
  ' xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"'

const HEADER = `<?xml version="1.0" encoding="UTF-8"?><hh:head ${NS}><hh:refList>
<hh:fontfaces itemCnt="2">
  <hh:fontface lang="HANGUL" fontCnt="1"><hh:font id="0" face="함초롬바탕" type="TTF"/></hh:fontface>
  <hh:fontface lang="LATIN" fontCnt="1"><hh:font id="0" face="Times New Roman" type="TTF"/></hh:fontface>
</hh:fontfaces>
<hh:borderFills itemCnt="2">
  <hh:borderFill id="1"><hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/><hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/><hh:topBorder type="NONE" width="0.1 mm" color="#000000"/><hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/></hh:borderFill>
  <hh:borderFill id="2"><hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/><hh:topBorder type="DOUBLE_SLIM" width="0.5 mm" color="#FF0000"/><hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/><hc:fillBrush><hc:winBrush faceColor="#D9D9D9" hatchColor="#000000" alpha="0"/></hc:fillBrush></hh:borderFill>
</hh:borderFills>
<hh:charProperties itemCnt="3">
  <hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none"><hh:fontRef hangul="0" latin="0"/><hh:ratio hangul="100"/><hh:spacing hangul="0"/><hh:underline type="NONE"/><hh:strikeout shape="3D"/></hh:charPr>
  <hh:charPr id="1" height="1600" textColor="#0000FF" shadeColor="none"><hh:fontRef hangul="0" latin="0"/><hh:ratio hangul="95"/><hh:spacing hangul="-5"/><hh:bold/><hh:underline type="BOTTOM"/><hh:strikeout shape="SOLID"/></hh:charPr>
  <hh:charPr id="2" height="900" textColor="#000000" shadeColor="none"><hh:fontRef hangul="0" latin="0"/><hh:italic/><hh:supscript/><hh:strikeout shape="NONE"/></hh:charPr>
</hh:charProperties>
<hh:paraProperties itemCnt="2">
  <hh:paraPr id="0"><hh:align horizontal="JUSTIFY"/><hh:heading type="NONE" idRef="0" level="0"/><hh:breakSetting pageBreakBefore="0" keepWithNext="0"/>
    <hp:switch><hp:case hp:required-namespace="x"><hh:margin><hc:intent value="999" unit="HWPUNIT"/></hh:margin></hp:case>
    <hp:default><hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/></hp:default></hp:switch></hh:paraPr>
  <hh:paraPr id="1"><hh:align horizontal="CENTER"/><hh:heading type="NONE" idRef="0" level="0"/>
    <hh:margin><hc:intent value="-2000" unit="HWPUNIT"/><hc:left value="1000" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="500" unit="HWPUNIT"/><hc:next value="1000" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="FIXED" value="2000" unit="HWPUNIT"/></hh:paraPr>
  <hh:paraPr id="2"><hh:align horizontal="LEFT"/><hh:heading type="NONE" idRef="0" level="0"/>
    <hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"><hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="500" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="100" unit="HWPUNIT"/></hp:case>
    <hp:default><hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="1000" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="100" unit="HWPUNIT"/></hp:default></hp:switch></hh:paraPr>
</hh:paraProperties>
</hh:refList></hh:head>`

const SEC_PR =
  '<hp:secPr><hp:pagePr landscape="WIDELY" width="59528" height="84188"><hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668" bottom="4252"/></hp:pagePr></hp:secPr>'

function cell(
  col: number,
  row: number,
  text: string,
  span = { col: 1, row: 1 },
  width = 10000,
): string {
  return (
    `<hp:tc borderFillIDRef="2" hasMargin="1"><hp:subList vertAlign="CENTER"><hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p></hp:subList>` +
    `<hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="${span.col}" rowSpan="${span.row}"/>` +
    `<hp:cellSz width="${width * span.col}" height="2000"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>`
  )
}

function section(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><hs:sec ${NS}>${body}</hs:sec>`
}

async function hwpx(
  sectionXml: string,
  extra: Record<string, Uint8Array> = {},
): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' })
  zip.file('Contents/header.xml', HEADER)
  zip.file('Contents/section0.xml', sectionXml)
  const items = Object.keys(extra)
    .map((path, i) => `<opf:item id="image${i + 1}" href="${path}" media-type="image/bmp"/>`)
    .join('')
  zip.file(
    'Contents/content.hpf',
    `<?xml version="1.0"?><opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:manifest>` +
      `<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>` +
      `<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>${items}` +
      `</opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`,
  )
  for (const [path, bytes] of Object.entries(extra)) zip.file(path, bytes)
  return zip.generateAsync({ type: 'uint8array' })
}

async function documentXml(docx: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(docx)).file('word/document.xml')!.async('string')
}

/** 2×2 24-bit BMP: red, green / blue, white (bottom-up rows) */
function tinyBmp(): Uint8Array {
  const rowBytes = 8 // 2 px × 3 bytes, padded to 4
  const pixels = [
    [255, 0, 0, 255, 255, 255], // bottom row: blue, white (BGR)
    [0, 0, 255, 0, 255, 0], // top row: red, green (BGR)
  ]
  const size = 54 + rowBytes * 2
  const out = new Uint8Array(size)
  const view = new DataView(out.buffer)
  out[0] = 0x42
  out[1] = 0x4d
  view.setUint32(2, size, true)
  view.setUint32(10, 54, true)
  view.setUint32(14, 40, true)
  view.setInt32(18, 2, true)
  view.setInt32(22, 2, true)
  view.setUint16(26, 1, true)
  view.setUint16(28, 24, true)
  pixels.forEach((row, y) => out.set(row, 54 + y * rowBytes))
  return out
}

describe('convertHwpxToDocx', () => {
  it('carries character shapes, paragraph shapes and page geometry', async () => {
    const xml = section(
      `<hp:p paraPrIDRef="1"><hp:run charPrIDRef="0">${SEC_PR}</hp:run>` +
        `<hp:run charPrIDRef="1"><hp:t>제목<hp:tab/>굵게</hp:t></hp:run>` +
        `<hp:run charPrIDRef="2"><hp:t>위첨자</hp:t></hp:run></hp:p>` +
        `<hp:p paraPrIDRef="0" pageBreak="1"><hp:run charPrIDRef="0"><hp:t>본문<hp:lineBreak/>둘째 줄</hp:t></hp:run></hp:p>`,
    )
    const result = await convertHwpxToDocx(await hwpx(xml))
    const doc = await documentXml(result.docx)
    expect(doc).toContain('<w:jc w:val="center"/>')
    // 내어쓰기: left 1000 + 2000 = 3000 HWPUNIT = 600 twips, hanging 400
    expect(doc).toMatch(/<w:ind w:left="600" w:hanging="400"\/>/)
    expect(doc).toMatch(/w:line="400" w:lineRule="exact"/)
    expect(doc).toContain('w:eastAsia="함초롬바탕"')
    expect(doc).toContain('w:ascii="Times New Roman"')
    expect(doc).toContain('<w:b/>')
    expect(doc).toContain('<w:strike/>')
    expect(doc).toContain('<w:u w:val="single"/>')
    expect(doc).toContain('<w:color w:val="0000FF"/>')
    expect(doc).toContain('<w:w w:val="95"/>')
    expect(doc).toContain('<w:vertAlign w:val="superscript"/>')
    expect(doc).toContain('<w:sz w:val="32"/>')
    expect(doc).toContain('<w:tab/>')
    expect(doc).toContain('<w:br/>')
    expect(doc).toContain('<w:pageBreakBefore/>')
    // PERCENT 160 of a 10pt line = exactly 16pt (320 twips)
    expect(doc).toMatch(/w:line="320" w:lineRule="exact"/)
    // A4 with Hangul's header band folded into the body margin
    expect(doc).toMatch(/<w:pgSz w:w="11906" w:h="16838"/)
    expect(doc).toMatch(/w:top="1984"/)
    // the strikeout shape="3D" of charPr 0 is Hangul's "off" value
    const plain = doc.slice(doc.indexOf('본문') - 400, doc.indexOf('본문'))
    expect(plain).not.toContain('<w:strike/>')
    const parsed = await parseDocx(result.docx)
    expect(parsed.blocks.length).toBeGreaterThanOrEqual(2)
  })

  it('reads paragraph margins from the HwpUnitChar case, not the doubled default', async () => {
    const xml = section(
      `<hp:p paraPrIDRef="2"><hp:run charPrIDRef="0">${SEC_PR}<hp:t>들여쓴 문단</hp:t></hp:run></hp:p>`,
    )
    const doc = await documentXml((await convertHwpxToDocx(await hwpx(xml))).docx)
    // case left 500 HWPUNIT = 100 twips (the default branch says 1000 = 200)
    expect(doc).toContain('<w:ind w:left="100"/>')
  })

  it('drops trailing empty lines before a section break instead of leaving a blank page', async () => {
    const empty = '<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"/></hp:p>'
    const zip = new JSZip()
    zip.file('mimetype', 'application/hwp+zip')
    zip.file('Contents/header.xml', HEADER)
    zip.file(
      'Contents/section0.xml',
      section(
        `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}<hp:t>첫 구역</hp:t></hp:run></hp:p>${empty}${empty}`,
      ),
    )
    zip.file(
      'Contents/section1.xml',
      section(
        `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>둘째 구역</hp:t></hp:run></hp:p>`,
      ),
    )
    const doc = await documentXml(
      (await convertHwpxToDocx(await zip.generateAsync({ type: 'uint8array' }))).docx,
    )
    const between = doc.slice(doc.indexOf('첫 구역'), doc.indexOf('둘째 구역'))
    expect(between.match(/<w:p[\s>]/g) ?? []).toHaveLength(1) // only the break paragraph itself
    expect(between).toContain('<w:pageBreakBefore/>')
  })

  it('spells out Hangul line breaking and auto spacing (Word defaults are the opposite)', async () => {
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}<hp:t>한글 English 123</hp:t></hp:run></hp:p>`,
    )
    const doc = await documentXml((await convertHwpxToDocx(await hwpx(xml))).docx)
    // paraPr 0 has no hh:autoSpacing → both off; breakSetting without
    // breakNonLatinWord is KEEP_WORD, which Hangul lays out between syllables
    expect(doc).toMatch(
      /<w:wordWrap w:val="0"\/><w:autoSpaceDE w:val="0"\/><w:autoSpaceDN w:val="0"\/>(<w:snapToGrid|<w:spacing)/,
    )
  })

  it('writes tables with spans, borders, fills, margins and a caption', async () => {
    const table =
      `<hp:tbl rowCnt="2" colCnt="3" borderFillIDRef="1"><hp:sz width="30000" height="4000"/>` +
      `<hp:caption side="TOP"><hp:subList><hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>표 1. 현황</hp:t></hp:run></hp:p></hp:subList></hp:caption>` +
      `<hp:inMargin left="141" right="141" top="141" bottom="141"/>` +
      `<hp:tr>${cell(0, 0, '병합', { col: 1, row: 2 })}${cell(1, 0, '가로', { col: 2, row: 1 })}</hp:tr>` +
      `<hp:tr>${cell(1, 1, '셀비')}${cell(2, 1, '셀씨')}</hp:tr></hp:tbl>`
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}<hp:t>앞</hp:t>${table}<hp:t>뒤</hp:t></hp:run></hp:p>`,
    )
    const doc = await documentXml((await convertHwpxToDocx(await hwpx(xml))).docx)
    const order = [
      '앞',
      '표 1. 현황',
      '<w:tbl>',
      '병합',
      '가로',
      '셀비',
      '셀씨',
      '</w:tbl>',
      '뒤',
    ].map((s) => doc.indexOf(s))
    expect(
      order.every((at, i) => at >= 0 && (i === 0 || at > order[i - 1])),
      JSON.stringify(order),
    ).toBe(true)
    expect(doc).toContain('<w:gridCol w:w="2000"/>')
    expect(doc).toContain('<w:gridSpan w:val="2"/>')
    expect(doc).toContain('<w:vMerge w:val="restart"/>')
    expect(doc).toMatch(/<w:vMerge\/>/)
    expect(doc).toContain('<w:top w:val="double" w:sz="11" w:space="0" w:color="FF0000"/>')
    expect(doc).toContain('w:fill="D9D9D9"')
    expect(doc).toContain('<w:vAlign w:val="center"/>')
    expect(doc).toContain('<w:left w:w="102" w:type="dxa"/>')
    // the table default (inMargin 141 HWPUNIT = 28 twips) is written once
    expect(doc).toContain('<w:tblCellMar><w:top w:w="28" w:type="dxa"/>')
  })

  it("declares Word's balanced-width layout, Hangul's half-em spaces", async () => {
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="1">${SEC_PR}<hp:t>자간 좁힘</hp:t></hp:run>` +
        `<hp:run charPrIDRef="0"><hp:t> </hp:t></hp:run><hp:run charPrIDRef="1"><hp:t>ABC</hp:t></hp:run></hp:p>`,
    )
    const docx = (await convertHwpxToDocx(await hwpx(xml))).docx
    const zip = await JSZip.loadAsync(docx)
    expect(await zip.file('word/settings.xml')!.async('string')).toContain(
      '<w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/></w:compat>',
    )
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
      'Target="settings.xml"',
    )
    const parsed = await parseDocx(docx)
    // and marks the document HWP-origin, so GenOffice draws the Hancom faces with their compat faces
    expect(parsed).toMatchObject({ balanceDbcsSpacing: true, useFELayout: true, hwpSource: true })
    // spacing -5% of 16pt = -16 twips on a full-width glyph: Word doubles the
    // -8 written on Hangul under the flag; Hangul's percent of a half-width
    // Latin glyph's own advance is about -8 too
    const doc = await documentXml(docx)
    const hangul = doc.slice(doc.indexOf('자간') - 300, doc.indexOf('자간'))
    expect(hangul).toContain('<w:spacing w:val="-8"/>')
    const latin = doc.slice(doc.indexOf('ABC') - 300, doc.indexOf('ABC'))
    expect(latin).toContain('<w:spacing w:val="-8"/>')
  })

  it("fits table rows to Hangul's row heights", async () => {
    const lineseg = (vert: number, gap: number) =>
      `<hp:lineseg textpos="0" vertpos="0" vertsize="${vert}" textheight="${vert}" baseline="850" spacing="${gap}" horzpos="0" horzsize="9000" flags="0"/>`
    const plainCell = (row: number, height: number, segs: string) =>
      `<hp:tc borderFillIDRef="1" hasMargin="0"><hp:subList vertAlign="TOP">` +
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0"><hp:t>둘째 줄까지 가는 셀 ${row}</hp:t></hp:run><hp:linesegarray>${segs}</hp:linesegarray></hp:p></hp:subList>` +
      `<hp:cellAddr colAddr="0" rowAddr="${row}"/><hp:cellSpan colSpan="1" rowSpan="1"/>` +
      `<hp:cellSz width="10000" height="${height}"/><hp:cellMargin left="0" right="0" top="0" bottom="0"/></hp:tc>`
    const table =
      `<hp:tbl rowCnt="2" colCnt="1" borderFillIDRef="1"><hp:sz width="10000" height="82000"/>` +
      `<hp:inMargin left="141" right="141" top="141" bottom="141"/>` +
      `<hp:tr>${plainCell(0, 2000, lineseg(1000, 600) + lineseg(1000, 600))}</hp:tr>` +
      `<hp:tr>${plainCell(1, 80000, lineseg(1000, 600))}</hp:tr></hp:tbl>`
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}${table}</hp:run></hp:p>`,
    )
    const doc = await documentXml((await convertHwpxToDocx(await hwpx(xml))).docx)
    const heights = [...doc.matchAll(/<w:trHeight w:val="(\d+)"/g)].map((m) => Number(m[1]))
    // 2000 HWPUNIT = 400 twips holds the inMargin 28 + 28 Word lays on top of the floor
    expect(heights).toEqual([344])
    // the second row declares a page and a half (split across pages): no floor
    expect(doc.match(/<w:tr>/g)).toHaveLength(2)
    // a cell's last paragraph drops the gap under its last line:
    // (1000 + 600 + 1000) HWPUNIT over 2 lines = 260 twips each, 1 line = 200
    expect(doc).toMatch(/w:line="260" w:lineRule="exact"/)
    expect(doc).toMatch(/w:line="200" w:lineRule="exact"/)
    // the paragraph Word needs after a closing table is a 1pt sliver
    expect(doc).toMatch(/<\/w:tbl><w:p><w:pPr><w:spacing [^>]*w:line="20" w:lineRule="exact"\/>/)
  })

  it('embeds BinData pictures, re-encoding BMP as PNG', async () => {
    const pic =
      `<hp:pic><hp:sz width="7500" height="7500"/><hp:pos treatAsChar="1"/>` +
      `<hc:img binaryItemIDRef="image1"/></hp:pic>`
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}${pic}</hp:run></hp:p>`,
    )
    const result = await convertHwpxToDocx(await hwpx(xml, { 'BinData/image1.bmp': tinyBmp() }))
    const zip = await JSZip.loadAsync(result.docx)
    const media = Object.keys(zip.files).filter(
      (name) => name.startsWith('word/media/') && !zip.files[name].dir,
    )
    expect(media).toHaveLength(1)
    const png = await zip.file(media[0])!.async('uint8array')
    expect([...png.slice(1, 4)]).toEqual([0x50, 0x4e, 0x47])
    const doc = await documentXml(result.docx)
    // 7500 HWPUNIT = 100 px = 952500 EMU
    expect(doc).toContain('cx="952500"')
    expect(result.dropped).toEqual([])
  })

  it('rejects an HWP 5 binary served under a .hwpx name', async () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0])
    await expect(convertHwpxToDocx(ole)).rejects.toBeInstanceOf(HwpxConvertError)
    await expect(convertHwpxToDocx(ole)).rejects.toMatchObject({ code: 'hwp-binary' })
  })

  it('rejects a zip that is not an HWPX package', async () => {
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text')
    await expect(
      convertHwpxToDocx(await zip.generateAsync({ type: 'uint8array' })),
    ).rejects.toMatchObject({ code: 'not-hwpx' })
  })

  it('counts what it drops instead of guessing', async () => {
    const xml = section(
      `<hp:p paraPrIDRef="0"><hp:run charPrIDRef="0">${SEC_PR}<hp:ctrl><hp:header/><hp:pageNum/></hp:ctrl>` +
        `<hp:equation><hp:script>x^2</hp:script></hp:equation><hp:t>본문</hp:t></hp:run></hp:p>`,
    )
    const result = await convertHwpxToDocx(await hwpx(xml))
    const dropped = Object.fromEntries(result.dropped.map((d) => [d.what, d.count]))
    expect(dropped).toMatchObject({ headers: 1, 'page-numbers': 1, equations: 1 })
  })
})

describe('bmpToPng', () => {
  it('decodes bottom-up 24-bit rows into an RGB PNG', () => {
    const png = bmpToPng(tinyBmp())!
    expect(png).not.toBeNull()
    const view = new DataView(png.buffer, png.byteOffset)
    expect(view.getUint32(16)).toBe(2)
    expect(view.getUint32(20)).toBe(2)
    expect(png[25]).toBe(2) // color type RGB
  })

  it('refuses RLE-compressed bitmaps', () => {
    const bmp = tinyBmp()
    new DataView(bmp.buffer).setUint32(30, 1, true) // BI_RLE8
    expect(bmpToPng(bmp)).toBeNull()
  })
})
