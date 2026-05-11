import type { IWorkbookData } from '@univerjs/core'
import LuckyExcel from '@mertdeveci55/univer-import-export'
import type { XlsxConverter } from './xlsxConverter'

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
          snapshot,
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
