/**
 * Shared helpers for assembling {@link ExportToPdfInput} from an adapter's
 * render context. Kept in core so React and Vue resolve per-page headers/footers
 * identically (parity) — mirroring the painter's `renderPage` selection of
 * default vs first-page vs even/odd variants.
 */

import type { Layout } from '../layout-engine/types';
import type { HeaderFooterContent } from '../layout-painter/renderPage';
import type { BlockLookup } from '../layout-painter/index';
import type { ExportToPdfInput, ExportToPdfOptions, PageBordersInput } from './types';

export interface HeaderFooterRenderContext {
  /** Default header/footer (used on pages 2+ when titlePg, or all pages otherwise). */
  header?: HeaderFooterContent | null;
  footer?: HeaderFooterContent | null;
  /** First-page header/footer (used on page 1 when titlePg is set). */
  firstPageHeader?: HeaderFooterContent | null;
  firstPageFooter?: HeaderFooterContent | null;
  /** Whether different first-page headers/footers are enabled (w:titlePg). */
  titlePg?: boolean;
}

/**
 * The render context an adapter captures during layout so it can later build an
 * {@link ExportToPdfInput} without re-running the pipeline. Mirrors the options
 * the adapter passes to the painter's `renderPages`.
 */
export interface AdapterPdfExportContext extends HeaderFooterRenderContext {
  blockLookup: BlockLookup;
  pageBorders?: PageBordersInput;
  backgroundColor?: string;
}

/** Build a `(pageNumber) => HeaderFooterContent | undefined` resolver for headers. */
export function headerResolver(
  ctx: HeaderFooterRenderContext
): (pageNumber: number) => HeaderFooterContent | undefined {
  return (pageNumber: number) => {
    if (ctx.titlePg && pageNumber === 1) return ctx.firstPageHeader ?? undefined;
    return ctx.header ?? undefined;
  };
}

/** Build a `(pageNumber) => HeaderFooterContent | undefined` resolver for footers. */
export function footerResolver(
  ctx: HeaderFooterRenderContext
): (pageNumber: number) => HeaderFooterContent | undefined {
  return (pageNumber: number) => {
    if (ctx.titlePg && pageNumber === 1) return ctx.firstPageFooter ?? undefined;
    return ctx.footer ?? undefined;
  };
}

/**
 * Assemble an {@link ExportToPdfInput} from a captured adapter context + layout.
 * Both adapters call this so neither can drift on which context fields flow into
 * the exporter.
 */
export function buildExportInput(
  layout: Layout,
  ctx: AdapterPdfExportContext,
  options?: ExportToPdfOptions
): ExportToPdfInput {
  return {
    layout,
    blockLookup: ctx.blockLookup,
    headerByPage: headerResolver(ctx),
    footerByPage: footerResolver(ctx),
    pageBorders: ctx.pageBorders,
    backgroundColor: ctx.backgroundColor,
    options,
  };
}

/**
 * Print a generated PDF blob via a hidden iframe (self-contained — no DOM-clone
 * style loss). Returns false if the iframe couldn't be created so the caller can
 * fall back to `window.print()`. Shared by both adapters.
 */
export function printPdfBlob(blob: Blob): boolean {
  try {
    const url = URL.createObjectURL(blob);
    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
    iframe.src = url;
    let cleaned = false;
    const cleanup = () => {
      if (cleaned) return;
      cleaned = true;
      iframe.remove();
      URL.revokeObjectURL(url);
    };
    iframe.onload = () => {
      try {
        iframe.contentWindow?.focus();
        iframe.contentWindow?.print();
      } catch {
        window.open(url, '_blank');
      }
    };
    document.body.appendChild(iframe);
    // Single cleanup timer, independent of `onload` — so the iframe/url are
    // always reclaimed even if the blob never loads. 60s covers the print dialog.
    window.setTimeout(cleanup, 60_000);
    return true;
  } catch {
    return false;
  }
}
