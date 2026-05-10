// Generates fixture .xlsx files for PoC and round-trip tests.
// Run with: npx --package=exceljs@^4 -- node frontend/scripts/generate-fixtures.mjs
// Adds NO production deps — exceljs is invoked transiently via npx.
//
// Notes:
// - ExcelJS 4.x is CommonJS-only without an ESM `exports` field, so Node's
//   strict ESM resolver cannot find a bare `import 'exceljs'` from this .mjs
//   file. We import via the package's main file `exceljs/excel.js`, which
//   round-trips through CJS interop and works on Node 20+.
// - When invoked through `npx --package=exceljs@^4 --`, npm puts the package
//   under `~/.npm/_npx/<hash>/node_modules/` and only adds the `.bin` dir to
//   PATH — Node's import resolver does NOT see it. We bootstrap by reading the
//   first PATH entry (npx prepends its own `.bin`) and hand-resolving exceljs's
//   main file via a `file://` URL. This keeps the npx invocation simple.
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
  } catch {}
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

const ExcelJS = await loadExcelJs()
await generateSimple(ExcelJS)
console.log('All fixtures generated.')
