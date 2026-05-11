// Generates fixture .xlsx files for PoC and round-trip tests.
// Run with: npx --package=exceljs@^4 -- node frontend/scripts/generate-fixtures.mjs
// Adds NO production deps — exceljs is invoked transiently via npx.
//
// Notes:
// - We support two resolution paths:
//   1. Local install (`npm install -D exceljs`) — `import 'exceljs/excel.js'`
//      resolves via Node's package resolver. (Bare `import 'exceljs'` would
//      also work; we use the explicit subpath so the fallback below stays
//      consistent with what we file://-import.)
//   2. Transient via `npx --package=exceljs@^4 --` — npm puts exceljs under
//      `~/.npm/_npx/<hash>/node_modules/` and only adds its `.bin/` to PATH.
//      Node's import resolver does NOT walk PATH, so the bare import fails.
//      We bootstrap by scanning PATH for `<x>/node_modules/.bin` entries and
//      hand-resolving `<x>/node_modules/exceljs/excel.js` via a `file://` URL.
//
// Generators are deterministic for cell content (see Math.sin pseudo-random
// in generateSimple), but ExcelJS stamps `wb.created` into core.xml, so the
// archive bytes drift between runs. Treat the .xlsx as a regenerable artifact,
// not a stable hash target.
import path from 'node:path'
import fs from 'node:fs'
import { fileURLToPath, pathToFileURL } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = path.resolve(__dirname, '../public/fixtures')

function locateExcelJsFromNpx() {
  // npx prepends `<cache>/node_modules/.bin` to PATH; walk PATH to find an
  // entry whose sibling `node_modules/exceljs/excel.js` exists.
  const sep = process.platform === 'win32' ? ';' : ':'
  for (const entry of (process.env.PATH || '').split(sep)) {
    if (!entry.endsWith(path.sep + 'node_modules' + path.sep + '.bin') &&
        !entry.endsWith('/node_modules/.bin')) continue
    const nodeModules = path.dirname(entry) // .../node_modules
    const candidate = path.join(nodeModules, 'exceljs', 'excel.js')
    if (fs.existsSync(candidate)) return candidate
  }
  return null
}

async function loadExcelJs() {
  // Prefer the conventional resolver (works if exceljs is installed locally).
  try {
    const m = await import('exceljs/excel.js')
    return m.default ?? m
  } catch (e) {
    if (e?.code !== 'ERR_MODULE_NOT_FOUND' && e?.code !== 'MODULE_NOT_FOUND') throw e
  }
  const npxPath = locateExcelJsFromNpx()
  if (!npxPath) {
    throw new Error(
      'ExcelJS not found. Run via:\n' +
      '  npx --package=exceljs@^4 -- node frontend/scripts/generate-fixtures.mjs',
    )
  }
  const m = await import(pathToFileURL(npxPath).href)
  return m.default ?? m
}

async function generateSimple(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()
  const ws = wb.addWorksheet('Sheet1')

  // A1:J100 — A 列文本, B-J 列数字
  for (let r = 1; r <= 100; r++) {
    ws.getCell(`A${r}`).value = `row-${r}`
    for (let c = 2; c <= 10; c++) {
      // 用稳定的伪随机（基于 r,c）确保可重现
      const v = Math.round((Math.sin(r * 100 + c) + 1) * 5000) / 100
      ws.getRow(r).getCell(c).value = v
    }
  }

  // K 列：合并 K1:K3 + 边框
  ws.mergeCells('K1:K3')
  const k1 = ws.getCell('K1')
  k1.value = 'merged'
  k1.alignment = { vertical: 'middle', horizontal: 'center' }
  const thin = { style: 'thin', color: { argb: 'FF000000' } }
  for (const row of [1, 2, 3]) {
    ws.getCell(`K${row}`).border = { top: thin, left: thin, bottom: thin, right: thin }
  }

  // 列宽（让它看起来更像真实文件）
  ws.getColumn(1).width = 12
  for (let c = 2; c <= 11; c++) ws.getColumn(c).width = 10

  const out = path.join(FIXTURES_DIR, 'simple-100x10.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

async function generateMerged(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()
  const ws = wb.addWorksheet('Sheet1')

  // 5x5 grid with various merge patterns
  ws.getCell('A1').value = '标题区'
  ws.mergeCells('A1:E1')
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' }
  ws.getCell('A1').font = { bold: true, size: 14 }

  // Vertical merge
  ws.getCell('A2').value = '类别'
  ws.mergeCells('A2:A5')
  ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' }

  // Block merge
  ws.getCell('B2').value = '总计'
  ws.mergeCells('B2:C3')
  ws.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' }

  // Filling out the rest
  for (let r = 4; r <= 10; r++) {
    for (let c = 1; c <= 5; c++) {
      const cell = ws.getRow(r).getCell(c)
      if (!cell.value) cell.value = `R${r}C${c}`
    }
  }

  // Borders on all merged regions
  const thin = { style: 'thin', color: { argb: 'FF000000' } }
  for (let r = 1; r <= 10; r++) {
    for (let c = 1; c <= 5; c++) {
      ws.getRow(r).getCell(c).border = { top: thin, left: thin, bottom: thin, right: thin }
    }
  }

  for (let c = 1; c <= 5; c++) ws.getColumn(c).width = 12

  const out = path.join(FIXTURES_DIR, 'merged-cells.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

async function generateFormulas(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()

  // Sheet 1: data table
  const data = wb.addWorksheet('Data')
  data.getRow(1).values = ['ID', 'Name', 'Price', 'Qty', 'Total']
  data.getRow(1).font = { bold: true }
  const products = [
    [1, 'Apple', 1.5, 100],
    [2, 'Banana', 0.5, 200],
    [3, 'Cherry', 3.0, 50],
    [4, 'Date', 5.0, 30],
    [5, 'Elderberry', 4.5, 40],
  ]
  for (let i = 0; i < products.length; i++) {
    const r = i + 2
    const [id, name, price, qty] = products[i]
    data.getCell(`A${r}`).value = id
    data.getCell(`B${r}`).value = name
    data.getCell(`C${r}`).value = price
    data.getCell(`D${r}`).value = qty
    // Formula: Total = Price * Qty
    data.getCell(`E${r}`).value = { formula: `C${r}*D${r}` }
  }

  // Summary row with SUM
  data.getCell('A8').value = 'Sum'
  data.getCell('A8').font = { bold: true }
  data.getCell('E8').value = { formula: 'SUM(E2:E6)' }
  data.getCell('E8').font = { bold: true }

  // Sheet 2: VLOOKUP demo
  const lookup = wb.addWorksheet('Lookup')
  lookup.getRow(1).values = ['Lookup ID', 'Found Name', 'Found Total']
  lookup.getRow(1).font = { bold: true }
  for (let i = 0; i < 3; i++) {
    const r = i + 2
    const lookupId = (i % 5) + 1
    lookup.getCell(`A${r}`).value = lookupId
    lookup.getCell(`B${r}`).value = { formula: `VLOOKUP(A${r},Data!A:E,2,FALSE)` }
    lookup.getCell(`C${r}`).value = { formula: `VLOOKUP(A${r},Data!A:E,5,FALSE)` }
  }

  for (const ws of [data, lookup]) {
    for (let c = 1; c <= 5; c++) ws.getColumn(c).width = 14
  }

  const out = path.join(FIXTURES_DIR, 'formulas-sum-vlookup.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

async function generateConditionalFormat(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()
  const ws = wb.addWorksheet('Sheet1')

  // Header
  ws.getRow(1).values = ['Score', 'Name', 'Status']
  ws.getRow(1).font = { bold: true }

  // 20 rows of data with varying scores
  for (let r = 2; r <= 21; r++) {
    const score = Math.round(Math.sin(r * 17) * 50 + 50)  // 0-100
    ws.getCell(`A${r}`).value = score
    ws.getCell(`B${r}`).value = `User${r - 1}`
    ws.getCell(`C${r}`).value = score >= 60 ? 'Pass' : 'Fail'
  }

  // Conditional formatting on Score column: 3-color scale
  ws.addConditionalFormatting({
    ref: 'A2:A21',
    rules: [
      {
        type: 'colorScale',
        priority: 1,
        cfvo: [
          { type: 'min' },
          { type: 'percentile', value: 50 },
          { type: 'max' },
        ],
        color: [
          { argb: 'FFF8696B' },  // red
          { argb: 'FFFFEB84' },  // yellow
          { argb: 'FF63BE7B' },  // green
        ],
      },
    ],
  })

  // Bold red for Fail status
  ws.addConditionalFormatting({
    ref: 'C2:C21',
    rules: [
      {
        type: 'cellIs',
        priority: 1,
        operator: 'equal',
        formulae: ['"Fail"'],
        style: { font: { bold: true, color: { argb: 'FFFF0000' } } },
      },
    ],
  })

  for (let c = 1; c <= 3; c++) ws.getColumn(c).width = 12

  const out = path.join(FIXTURES_DIR, 'conditional-format.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

async function generateMultiSheet(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()

  // Sheet 1: Summary
  const summary = wb.addWorksheet('Summary')
  summary.getRow(1).values = ['Region', 'Q1', 'Q2', 'Q3', 'Q4']
  summary.getRow(1).font = { bold: true }
  const regions = ['North', 'South', 'East', 'West']
  for (let i = 0; i < regions.length; i++) {
    const r = i + 2
    summary.getCell(`A${r}`).value = regions[i]
    for (let q = 1; q <= 4; q++) {
      summary.getRow(r).getCell(q + 1).value = Math.round(Math.sin(i * 4 + q) * 1000 + 5000)
    }
  }

  // Sheet 2: Details
  const details = wb.addWorksheet('Details')
  details.getRow(1).values = ['Item', 'Region', 'Quantity']
  details.getRow(1).font = { bold: true }
  for (let r = 2; r <= 21; r++) {
    details.getCell(`A${r}`).value = `Item-${r - 1}`
    details.getCell(`B${r}`).value = regions[(r - 2) % 4]
    details.getCell(`C${r}`).value = Math.floor(Math.abs(Math.sin(r) * 100))
  }

  // Sheet 3: Notes
  const notes = wb.addWorksheet('Notes')
  notes.getCell('A1').value = 'Multi-sheet fixture'
  notes.getCell('A1').font = { italic: true, size: 14 }
  notes.getCell('A3').value = 'Use this to verify round-trip preserves sheet order, names, and per-sheet content.'

  for (const ws of [summary, details, notes]) {
    for (let c = 1; c <= 5; c++) ws.getColumn(c).width = 14
  }

  const out = path.join(FIXTURES_DIR, 'multi-sheet.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

async function generateLarge10kRows(ExcelJS) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'fixture-generator'
  wb.created = new Date()
  const ws = wb.addWorksheet('Sheet1')

  ws.getRow(1).values = ['Row', 'Value']
  ws.getRow(1).font = { bold: true }

  // 10,000 rows
  for (let r = 2; r <= 10001; r++) {
    ws.getCell(`A${r}`).value = `row_${r - 1}`
    ws.getCell(`B${r}`).value = Math.round(Math.sin(r) * 10000) / 100
  }

  ws.getColumn(1).width = 14
  ws.getColumn(2).width = 12

  const out = path.join(FIXTURES_DIR, 'large-10k-rows.xlsx')
  await wb.xlsx.writeFile(out)
  console.log(`Wrote ${out}`)
}

const generators = [
  generateSimple,
  generateMerged,
  generateFormulas,
  generateConditionalFormat,
  generateMultiSheet,
  generateLarge10kRows,
]
const ExcelJS = await loadExcelJs()
fs.mkdirSync(FIXTURES_DIR, { recursive: true })
for (const gen of generators) {
  await gen(ExcelJS)
}
console.log('All fixtures generated.')
