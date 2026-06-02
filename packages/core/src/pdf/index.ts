/**
 * Native vector PDF export.
 *
 * Walks the editor's computed `Layout` (the same one the painter draws) and emits
 * a real vector PDF — selectable text, embedded subset fonts, vector decoration,
 * embedded images — so the output matches the editor and MS Word. In-browser only
 * (text measurement and image re-encoding are canvas-bound).
 *
 * Dynamically imported (`@eigenpal/docx-editor-core/pdf`) so pdf-lib/fontkit stay
 * out of the editor's hot path and main bundle.
 *
 * @packageDocumentation
 * @public
 */

import { PDFDocument } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import { GoogleFontProvider, type FontProvider } from './fontProvider';
import { createImageEmbedder } from './pdfImage';
import { drawPage, hfFaces, pageFaces, pageImageSrcs } from './pdfPage';
import type { ExportToPdfInput, ExportToPdfOptions } from './types';

export type { ExportToPdfInput, ExportToPdfOptions, PageBordersInput } from './types';
export type { FontProvider, FontStyle } from './fontProvider';
export { GoogleFontProvider } from './fontProvider';
export {
  headerResolver,
  footerResolver,
  buildExportInput,
  printPdfBlob,
  type HeaderFooterRenderContext,
  type AdapterPdfExportContext,
} from './buildInput';

const PRODUCER = 'Eigenpal DOCX Editor';

/**
 * Export the document described by `input` to a PDF `Blob`.
 *
 * @public
 */
export async function exportToPdf(input: ExportToPdfInput): Promise<Blob> {
  const { layout, blockLookup } = input;
  const options: ExportToPdfOptions = input.options ?? {};

  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);

  // Deterministic-friendly metadata (stable output when `now` is provided).
  doc.setProducer(PRODUCER);
  doc.setCreator(PRODUCER);
  if (options.documentName) doc.setTitle(options.documentName);
  if (options.author) doc.setAuthor(options.author);
  if (options.now) {
    const d = new Date(options.now);
    doc.setCreationDate(d);
    doc.setModificationDate(d);
  }

  const fonts: FontProvider = input.fontProviderFactory
    ? input.fontProviderFactory(doc)
    : new GoogleFontProvider(doc, {
        unicodeFallbackBytes: options.unicodeFallbackBytes,
        onWarning: options.onWarning,
      });
  const embedder = createImageEmbedder(doc, options.onWarning);

  const pageNumberStart = options.pageNumberStart ?? 1;
  const totalPages = layout.pages.length;

  for (let i = 0; i < layout.pages.length; i++) {
    const page = layout.pages[i];
    const header = input.headerByPage?.(page.number);
    const footer = input.footerByPage?.(page.number);
    // Warm up every face + image this page (and its header/footer) references so
    // positioning + drawing run synchronously (run-x == glyph advance; inline
    // images draw without an async hop).
    await fonts.warmUp([...pageFaces(page, blockLookup), ...hfFaces(header), ...hfFaces(footer)]);
    await embedder.warmUp(pageImageSrcs(page, blockLookup, header, footer));
    drawPage({
      doc,
      page,
      blockLookup,
      fonts,
      embedder,
      field: { pageNumber: page.number + pageNumberStart - 1, totalPages, now: options.now },
      backgroundColor: input.backgroundColor,
      pageBorders: input.pageBorders,
      header,
      footer,
      onWarning: options.onWarning,
    });
    options.onProgress?.((i + 1) / totalPages);
  }

  const bytes = await doc.save();
  return new Blob([bytes as BlobPart], { type: 'application/pdf' });
}
