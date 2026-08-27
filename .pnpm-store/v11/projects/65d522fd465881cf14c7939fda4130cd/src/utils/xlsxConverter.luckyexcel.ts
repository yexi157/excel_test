import type { ICellData, IRange, IWorkbookData } from '@univerjs/core'
import LuckyExcel from '@mertdeveci55/univer-import-export'
import type { XlsxConverter } from './xlsxConverter'

const DATA_VALIDATION_PLUGIN_NAME = 'SHEET_DATA_VALIDATION_PLUGIN'

type DataValidationRule = {
  uid: string
  ranges: IRange[]
  type: string
  formula1?: string
  formula2?: string
  allowBlank?: boolean
  operator?: string
  showErrorMessage?: boolean
  showInputMessage?: boolean
  error?: string
  errorStyle?: number
  errorTitle?: string
  prompt?: string
  promptTitle?: string
}

type ExcelDataValidation = {
  type: string
  formulae?: unknown[]
  allowBlank?: boolean
  operator?: string
  showErrorMessage?: boolean
  showInputMessage?: boolean
  error?: string
  errorStyle?: string
  errorTitle?: string
  prompt?: string
  promptTitle?: string
}

type ExcelDataValidationModel = Record<string, ExcelDataValidation>

function parseValidationResource(workbook: IWorkbookData): Record<string, DataValidationRule[]> {
  const resource = workbook.resources?.find(item => item.name === DATA_VALIDATION_PLUGIN_NAME)
  if (!resource?.data) return {}
  try {
    return JSON.parse(resource.data) as Record<string, DataValidationRule[]>
  } catch {
    return {}
  }
}

function hasValidationRules(rulesBySheet: Record<string, DataValidationRule[]>): boolean {
  return Object.values(rulesBySheet).some(rules => rules.length > 0)
}

function isCellOrRangeReference(value: string): boolean {
  return /^(?:(?:'[^']+'|[^!,]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(value)
}

function normalizeListFormulaForUniver(formula: string): string {
  const unquoted = formula.length >= 2 && formula.startsWith('"') && formula.endsWith('"')
    ? formula.slice(1, -1)
    : formula
  return !unquoted.startsWith('=') && isCellOrRangeReference(unquoted) ? `=${unquoted}` : unquoted
}

function normalizeImportedValidationResources(workbook: IWorkbookData): IWorkbookData {
  const resourceIndex = workbook.resources?.findIndex(item => item.name === DATA_VALIDATION_PLUGIN_NAME) ?? -1
  if (resourceIndex < 0) return workbook

  const rulesBySheet = parseValidationResource(workbook)
  let changed = false
  const normalized = Object.fromEntries(Object.entries(rulesBySheet).map(([sheetId, rules]) => [
    sheetId,
    rules.map((rule) => {
      if (rule.type !== 'list' || !rule.formula1) return rule
      const formula1 = normalizeListFormulaForUniver(rule.formula1)
      if (formula1 === rule.formula1) return rule
      changed = true
      return { ...rule, formula1 }
    }),
  ]))
  if (!changed) return workbook

  const resources = [...workbook.resources!]
  resources[resourceIndex] = { ...resources[resourceIndex], data: JSON.stringify(normalized) }
  return { ...workbook, resources }
}

function columnIndex(columnName: string): number {
  return [...columnName.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1
}

function parseCellAddress(address: string): { row: number; column: number } | null {
  const match = address.replaceAll('$', '').match(/^([A-Z]+)(\d+)$/i)
  if (!match) return null
  return { row: Number(match[2]) - 1, column: columnIndex(match[1]) }
}

function parseValidationRange(address: string): IRange | null {
  const [startAddress, endAddress = startAddress] = address.split(':')
  const start = parseCellAddress(startAddress)
  const end = parseCellAddress(endAddress)
  if (!start || !end) return null
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  }
}

function excelErrorStyle(style?: string): number | undefined {
  if (!style) return undefined
  return ({ information: 0, stop: 1, warning: 2 } as Record<string, number>)[style]
}

function toFormulaString(value: unknown, type: string): string | undefined {
  if (value === undefined || value === null) return undefined
  const formula = value instanceof Date ? value.toISOString() : String(value)
  return type === 'list' ? normalizeListFormulaForUniver(formula) : formula
}

async function recoverValidationResources(blob: Blob, workbookData: IWorkbookData): Promise<IWorkbookData> {
  const normalized = normalizeImportedValidationResources(workbookData)
  if (hasValidationRules(parseValidationResource(normalized))) return normalized

  const { default: ExcelJS } = await import('@zwight/exceljs')
  const excelWorkbook = new ExcelJS.Workbook()
  await excelWorkbook.xlsx.load(await blob.arrayBuffer())

  const rulesBySheet: Record<string, DataValidationRule[]> = {}
  for (const worksheet of excelWorkbook.worksheets) {
    const sheetEntry = Object.entries(normalized.sheets).find(([, sheet]) => sheet.name === worksheet.name)
    if (!sheetEntry) continue
    const [sheetId] = sheetEntry
    const model = (worksheet as unknown as { dataValidations: { model: ExcelDataValidationModel } })
      .dataValidations.model
    let ruleIndex = 0
    for (const [rangeAddresses, validation] of Object.entries(model)) {
      const ranges = rangeAddresses.split(/\s+/).map(parseValidationRange).filter((range): range is IRange => !!range)
      if (!ranges.length || validation.type === 'any') continue
      const formula1 = toFormulaString(validation.formulae?.[0], validation.type)
      const formula2 = toFormulaString(validation.formulae?.[1], validation.type)
      const rule: DataValidationRule = {
        uid: `xlsx-validation-${sheetId}-${ruleIndex++}`,
        ranges,
        type: validation.type,
        allowBlank: validation.allowBlank,
        operator: validation.operator,
        formula1,
        formula2,
        showErrorMessage: validation.showErrorMessage,
        showInputMessage: validation.showInputMessage,
        error: validation.error,
        errorStyle: excelErrorStyle(validation.errorStyle),
        errorTitle: validation.errorTitle,
        prompt: validation.prompt,
        promptTitle: validation.promptTitle,
      }
      rulesBySheet[sheetId] ??= []
      rulesBySheet[sheetId].push(rule)
    }
  }
  if (!hasValidationRules(rulesBySheet)) return normalized

  const resources = [...(normalized.resources ?? [])]
  const resourceIndex = resources.findIndex(item => item.name === DATA_VALIDATION_PLUGIN_NAME)
  const validationResource = { name: DATA_VALIDATION_PLUGIN_NAME, data: JSON.stringify(rulesBySheet) }
  if (resourceIndex >= 0) resources[resourceIndex] = validationResource
  else resources.push(validationResource)
  return { ...normalized, resources }
}

function prepareValidationResourcesForExport(workbook: IWorkbookData): IWorkbookData {
  const resourceIndex = workbook.resources?.findIndex(item => item.name === DATA_VALIDATION_PLUGIN_NAME) ?? -1
  if (resourceIndex < 0) return workbook
  const rulesBySheet = parseValidationResource(workbook)
  let changed = false
  const exportedRules = Object.fromEntries(Object.entries(rulesBySheet).map(([sheetId, rules]) => [
    sheetId,
    rules.map((rule) => {
      if (rule.type !== 'list' || !rule.formula1?.startsWith('=')) return rule
      changed = true
      return { ...rule, formula1: rule.formula1.slice(1) }
    }),
  ]))
  if (!changed) return workbook
  const resources = [...workbook.resources!]
  resources[resourceIndex] = { ...resources[resourceIndex], data: JSON.stringify(exportedRules) }
  return { ...workbook, resources }
}

function columnName(column: number): string {
  let value = column + 1
  let result = ''
  while (value > 0) {
    value--
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26)
  }
  return result
}

function serializeValidationRange(range: IRange): string {
  const start = `${columnName(range.startColumn)}${range.startRow + 1}`
  const end = `${columnName(range.endColumn)}${range.endRow + 1}`
  return start === end ? start : `${start}:${end}`
}

function excelValidationFormula(rule: DataValidationRule, formula?: string): string | undefined {
  if (formula === undefined) return undefined
  if (rule.type !== 'list') return formula.startsWith('=') ? formula.slice(1) : formula
  if (formula.startsWith('=')) return formula.slice(1)
  return `"${formula.replaceAll('"', '""')}"`
}

function excelValidationErrorStyle(style?: number): string | undefined {
  if (style === undefined) return undefined
  return ({ 0: 'information', 1: 'stop', 2: 'warning' } as Record<number, string>)[style]
}

async function writeValidationResourcesToXlsx(
  buffer: ArrayBuffer,
  workbookData: IWorkbookData,
): Promise<ArrayBuffer> {
  const rulesBySheet = parseValidationResource(workbookData)
  if (!hasValidationRules(rulesBySheet)) return buffer

  const { default: ExcelJS } = await import('@zwight/exceljs')
  const excelWorkbook = new ExcelJS.Workbook()
  await excelWorkbook.xlsx.load(buffer)
  for (const [sheetId, rules] of Object.entries(rulesBySheet)) {
    const sheetName = workbookData.sheets[sheetId]?.name
    const worksheet = sheetName ? excelWorkbook.getWorksheet(sheetName) : undefined
    if (!worksheet) continue
    const dataValidations = (worksheet as unknown as {
      dataValidations: { add: (address: string, validation: ExcelDataValidation) => void }
    }).dataValidations
    for (const rule of rules) {
      const formulae = [
        excelValidationFormula(rule, rule.formula1),
        excelValidationFormula(rule, rule.formula2),
      ].filter((formula): formula is string => formula !== undefined)
      const validation: ExcelDataValidation = {
        type: rule.type,
        formulae,
        allowBlank: rule.allowBlank,
        operator: rule.operator,
        showErrorMessage: rule.showErrorMessage,
        showInputMessage: rule.showInputMessage,
        error: rule.error,
        errorStyle: excelValidationErrorStyle(rule.errorStyle),
        errorTitle: rule.errorTitle,
        prompt: rule.prompt,
        promptTitle: rule.promptTitle,
      }
      for (const range of rule.ranges) dataValidations.add(serializeValidationRange(range), validation)
    }
  }
  return await excelWorkbook.xlsx.writeBuffer() as ArrayBuffer
}

type TextDecorationStyle = {
  ul?: { s: number }
  st?: { s: number }
}

function stripDisabledTextDecorations<T extends TextDecorationStyle>(style: T): T {
  const removeUnderline = style.ul?.s === 0
  const removeStrikethrough = style.st?.s === 0
  if (!removeUnderline && !removeStrikethrough) return style

  const sanitized = { ...style }
  if (removeUnderline) delete sanitized.ul
  if (removeStrikethrough) delete sanitized.st
  return sanitized
}

function sanitizeCell(cell: ICellData): ICellData {
  let sanitized = cell

  if (cell.s && typeof cell.s === 'object') {
    const style = stripDisabledTextDecorations(cell.s)
    if (style !== cell.s) sanitized = { ...sanitized, s: style }
  }

  const textRuns = cell.p?.body?.textRuns
  if (textRuns?.length) {
    let runsChanged = false
    const sanitizedRuns = textRuns.map((run) => {
      if (!run.ts) return run
      const textStyle = stripDisabledTextDecorations(run.ts)
      if (textStyle === run.ts) return run
      runsChanged = true
      return { ...run, ts: textStyle }
    })

    if (runsChanged) {
      sanitized = {
        ...sanitized,
        p: {
          ...cell.p!,
          body: { ...cell.p!.body!, textRuns: sanitizedRuns },
        },
      }
    }
  }

  return sanitized
}

function sanitizeSnapshotForExport(snapshot: IWorkbookData): IWorkbookData {
  snapshot = prepareValidationResourcesForExport(snapshot)
  let workbookChanged = false
  let styles = snapshot.styles

  for (const [styleId, style] of Object.entries(snapshot.styles)) {
    if (!style) continue
    const sanitizedStyle = stripDisabledTextDecorations(style)
    if (sanitizedStyle === style) continue
    if (styles === snapshot.styles) styles = { ...snapshot.styles }
    styles[styleId] = sanitizedStyle
    workbookChanged = true
  }

  let sheets = snapshot.sheets
  for (const sheetId of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[sheetId]
    if (!sheet?.cellData) continue

    let cellData = sheet.cellData
    let sheetChanged = false
    for (const [rowIndex, sourceRow] of Object.entries(sheet.cellData)) {
      if (!sourceRow) continue

      let row = sourceRow
      for (const [columnIndex, sourceCell] of Object.entries(sourceRow)) {
        const cell = sourceCell as ICellData | undefined
        if (!cell) continue
        const sanitizedCell = sanitizeCell(cell)
        if (sanitizedCell === cell) continue
        if (row === sourceRow) row = { ...sourceRow }
        row[Number(columnIndex)] = sanitizedCell
      }

      if (row !== sourceRow) {
        if (cellData === sheet.cellData) cellData = { ...sheet.cellData }
        cellData[Number(rowIndex)] = row
        sheetChanged = true
      }
    }

    if (sheetChanged) {
      if (sheets === snapshot.sheets) sheets = { ...snapshot.sheets }
      sheets[sheetId] = { ...sheet, cellData }
      workbookChanged = true
    }
  }

  const defaultStyle = snapshot.defaultStyle && typeof snapshot.defaultStyle === 'object'
    ? stripDisabledTextDecorations(snapshot.defaultStyle)
    : snapshot.defaultStyle
  if (defaultStyle !== snapshot.defaultStyle) workbookChanged = true

  return workbookChanged ? { ...snapshot, styles, sheets, defaultStyle } : snapshot
}

export function createLuckyexcelConverter(): XlsxConverter {
  return {
    async toUniver(blob: Blob, fileName: string): Promise<IWorkbookData> {
      const file = new File([blob], fileName)
      return new Promise<IWorkbookData>((resolve, reject) => {
        LuckyExcel.transformExcelToUniver(
          file,
          (workbookData: IWorkbookData) => {
            recoverValidationResources(blob, workbookData).then(resolve, reject)
          },
          (err: Error) => reject(err),
        )
      })
    },

    async toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob> {
      return new Promise<Blob>((resolve, reject) => {
        LuckyExcel.transformUniverToExcel({
          snapshot: sanitizeSnapshotForExport(snapshot),
          fileName,
          getBuffer: true,
          success: async (buffer?: ArrayBuffer) => {
            if (!buffer) {
              reject(new Error('LuckyExcel returned empty buffer'))
              return
            }
            try {
              const finalBuffer = await writeValidationResourcesToXlsx(buffer, snapshot)
              resolve(new Blob([finalBuffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              }))
            } catch (err) {
              reject(err)
            }
          },
          error: (err: Error) => reject(err),
        })
      })
    },
  }
}
