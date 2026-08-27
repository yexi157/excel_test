import {
  IUniverInstanceService,
  InterceptorEffectEnum,
  UniverInstanceType,
  type Univer,
} from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'

// Light blue header color. ARGB hex (Univer style format).
const HEADER_BG_RGB = '#e8f4fd'

/**
 * Register a CELL_CONTENT interceptor that paints a light-blue background on
 * row 0 cells with content. UI-only: never modifies IWorkbookData (the
 * intercept handler only transforms cell data passed to the renderer).
 *
 * Uses the same mechanism Univer's own sheets-numfmt-ui uses for runtime
 * cell decoration — see node_modules/@univerjs/sheets-numfmt-ui/lib/es/index.js
 * lines 960-998 for reference.
 *
 * Pass the raw `Univer` instance (not the Facade) so we can access the
 * injector via the public `__getInjector()` method.
 *
 * Returns a disposable that tears down the interceptor + forces one final
 * repaint.
 */
export function registerHeaderHighlight(univer: Univer): { dispose: () => void } | null {
  // The Univer class exposes its DI container via __getInjector.
  // FUniver Facade exposes it as `_injector` (protected) — using the raw
  // Univer instance is cleaner and survives Facade refactors.
  const injector = (univer as unknown as { __getInjector(): {
    get<T>(token: unknown): T
  } }).__getInjector()
  if (!injector) return null

  let interceptorService: SheetInterceptorService
  let renderManager: IRenderManagerService
  let instanceService: IUniverInstanceService
  try {
    interceptorService = injector.get(SheetInterceptorService)
    renderManager = injector.get(IRenderManagerService)
    instanceService = injector.get(IUniverInstanceService)
  } catch (err) {
    console.warn('[headerHighlight] DI lookup failed:', err)
    return null
  }

  const interceptorDisposable = interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    priority: 100,
    // Style effect is enough — we only override background fill, not value.
    effect: InterceptorEffectEnum.Style,
    handler: (cell, location, next) => {
      const result = next(cell) ?? {}
      if (location.row !== 0) return result

      // Check for actual content (raw cell, not the intercepted one — avoids
      // cycle and reflects the underlying data).
      const raw = location.worksheet.getCellRaw(location.row, location.col)
      const v = raw?.v
      if (v === null || v === undefined || v === '') return result

      // Merge with any existing style on the result. Result.s may be:
      //   - undefined → no existing style
      //   - string → style ID (reference into workbook.styles)
      //   - object  → inline IStyleData
      // For string IDs we'd need to resolve via workbook.getStyles(); to keep
      // this self-contained and predictable, we only override the bg field
      // and let the renderer merge with the referenced style by spreading
      // an inline object. If the cell already had an inline style we extend
      // it; if it had a string ID, the new inline object will replace it
      // (cell stays styled by referenced ID via the Univer style cache,
      // because next(cell) yields the resolved style object, not the ID).
      const existingStyle = (result as { s?: unknown }).s
      const baseStyle = existingStyle && typeof existingStyle === 'object' ? existingStyle : {}
      return {
        ...result,
        s: {
          ...baseStyle,
          bg: { rgb: HEADER_BG_RGB },
        },
      }
    },
  })

  // Force a one-shot repaint so the interceptor takes effect for cells that
  // are already on screen (Steady fires after the first paint).
  function repaint() {
    try {
      const wb = instanceService.getCurrentUnitForType(UniverInstanceType.UNIVER_SHEET) as
        | { getUnitId(): string }
        | null
      if (!wb) return
      const renderer = renderManager.getRenderById(wb.getUnitId())
      const main = (renderer as { mainComponent?: { makeDirty: (v?: boolean) => void } } | null)
        ?.mainComponent
      main?.makeDirty(true)
    } catch (err) {
      console.warn('[headerHighlight] repaint failed:', err)
    }
  }
  repaint()

  return {
    dispose: () => {
      interceptorDisposable.dispose()
      // One last repaint so the highlight goes away cleanly.
      repaint()
    },
  }
}
