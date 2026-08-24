import type { ICellData, IWorkbookData } from '@univerjs/core'
import LuckyExcel from '@mertdeveci55/univer-import-export'
import type { XlsxConverter } from './xlsxConverter'

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
          (workbookData: IWorkbookData) => resolve(workbookData),
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
          success: (buffer?: ArrayBuffer) => {
            if (!buffer) {
              reject(new Error('LuckyExcel returned empty buffer'))
              return
            }
            resolve(new Blob([buffer], {
              type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            }))
          },
          error: (err: Error) => reject(err),
        })
      })
    },
  }
}
