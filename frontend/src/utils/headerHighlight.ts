import type { FUniver } from '@univerjs/core/facade'
// Side-effect import: augments FUniver with `getSheetHooks()` (sheets-ui Facade)
import '@univerjs/sheets-ui/facade'

const HEADER_BG = 'rgba(232, 244, 253, 0.7)'

/**
 * Register a cell-render hook that paints a translucent light-blue background
 * on row 0 for every cell that has content. UI-only: does NOT modify
 * IWorkbookData, so downloaded .xlsx files look identical to before.
 *
 * `getSheetHooks()` is a workbook-level interceptor — the hook applies to
 * every cell of every sheet in the active workbook automatically.
 *
 * Call once per Univer instance, after the Steady lifecycle stage so the
 * sheets-ui hook service is fully initialised.
 *
 * Returns the disposable so the caller can dispose on Univer teardown.
 */
export function registerHeaderHighlight(univerAPI: FUniver): { dispose: () => void } | null {
  const hooks = (univerAPI as any).getSheetHooks?.()
  if (!hooks?.onCellRender) return null

  return hooks.onCellRender([
    {
      // Hot path: fires for every cell paint (initial render, scroll, edit).
      // Keep tight — no allocations, no logging.
      drawWith: (ctx: CanvasRenderingContext2D, info: any) => {
        if (info.row !== 0) return
        const v = info?.data?.v
        if (v === null || v === undefined || v === '') return
        const coord = info.primaryWithCoord
        if (!coord) return
        const { startX, startY, endX, endY } = coord
        ctx.save()
        ctx.fillStyle = HEADER_BG
        ctx.fillRect(startX, startY, endX - startX, endY - startY)
        ctx.restore()
      },
      // Paint behind cell content so text/borders stay on top.
      zIndex: -1,
    },
  ])
}
