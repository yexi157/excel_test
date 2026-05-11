import type { FUniver } from '@univerjs/core/facade'
// Side-effect import: augments FUniver with `getSheetHooks()` (sheets-ui Facade)
import '@univerjs/sheets-ui/facade'
import { IRenderManagerService } from '@univerjs/engine-render'

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
 *
 * Implementation notes (gotchas discovered by tracing @univerjs/sheets-ui
 * v0.22 ES bundle):
 *
 *   1. The facade signature is `onCellRender(customRender, effect?, priority?)`.
 *      Default `effect` is already `InterceptorEffectEnum.Style` (=1) per
 *      facade.js:1050, so we don't need to pass it explicitly.
 *
 *   2. The interceptor handler skips empty cells:
 *        `if (!cell) return next(cell);`
 *      This is fine for our row-0 paint because we only target cells that
 *      already have a value (`v` defined) — those make `cell` truthy.
 *
 *   3. CRITICAL: registering the hook on `Steady` happens AFTER the first
 *      paint (Rendered = 2 < Steady = 3). The Custom render extension reads
 *      `cellData.customRender` during draw; without a dirty signal, the
 *      canvas does not redraw, so `drawWith` is never invoked. We mirror
 *      what `sheets-numfmt-ui` does after registering its preview
 *      interceptor: call `mainComponent.makeDirty()` to force a repaint.
 */
export function registerHeaderHighlight(univerAPI: FUniver): { dispose: () => void } | null {
  const hooks = (univerAPI as any).getSheetHooks?.()
  if (!hooks?.onCellRender) return null

  const disposable = hooks.onCellRender([
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
      // `zIndex` here orders our render within the Custom extension's
      // per-cell customRender array (sorted via sortRules in
      // engine-render Custom.draw). The Custom extension itself runs at
      // z=55, AFTER background (z=21), font (z=45) and border (z=50) —
      // so a translucent fill here visually overlays the cell text. The
      // 0.7 alpha keeps the text legible.
      zIndex: -1,
    },
  ])

  // Force a repaint so the freshly-registered interceptor actually runs
  // for the cells already on screen. Without this, drawWith never fires
  // until the user scrolls / edits / resizes.
  forceSpreadsheetRepaint(univerAPI)

  return disposable
}

/**
 * Walk down to the active spreadsheet's main render component and mark it
 * dirty. Mirrors the post-registration call done by sheets-numfmt-ui.
 *
 * Best-effort: if any link in the chain is missing (e.g. no active
 * workbook yet), we silently no-op rather than throw — the caller may be
 * registering the hook before any unit exists.
 */
function forceSpreadsheetRepaint(univerAPI: FUniver): void {
  try {
    const workbook = (univerAPI as any).getActiveWorkbook?.()
    const unitId: string | undefined = workbook?.getId?.()
    if (!unitId) return
    // The injector lives on the underlying Univer instance, exposed via
    // a private accessor on FUniver. We reach in deliberately because
    // there's no facade equivalent for forcing a canvas repaint in 0.22.
    const injector = (univerAPI as any)._injector
    if (!injector) return
    const renderManager = injector.get(IRenderManagerService)
    const render = renderManager?.getRenderById?.(unitId)
    render?.mainComponent?.makeDirty?.(true)
  } catch {
    // Swallow — repaint is a UX nicety, not a correctness requirement.
  }
}
