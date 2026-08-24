// 抽象接口：让 Luckyexcel 可替换（spec §5.3）。
// 默认导出 @zwight/luckyexcel 实现；如踩坑可换 SheetJS 自写或后端 Python，调用方零改动。

import type { IWorkbookData } from '@univerjs/core'
import { createLuckyexcelConverter } from './xlsxConverter.luckyexcel'

export interface XlsxConverter {
  /** .xlsx Blob → Univer IWorkbookData */
  toUniver(blob: Blob, fileName: string): Promise<IWorkbookData>
  /** Univer IWorkbookData → .xlsx Blob */
  toXlsx(snapshot: IWorkbookData, fileName: string): Promise<Blob>
}

export const xlsxConverter: XlsxConverter = createLuckyexcelConverter()
