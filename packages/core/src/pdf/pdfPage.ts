/**
 * Page-level drawing for the PDF exporter: size the page, paint the background,
 * and dispatch each fragment to its renderer. Tables, text boxes, page borders,
 * and header/footer layers are staged (see TODOs) behind warnings so an
 * unsupported fragment degrades gracefully rather than crashing.
 */

import { rgb, type PDFDocument, type PDFPage } from 'pdf-lib';
import {
  type Page,
  type FlowBlock,
  type ParagraphBlock,
  type ParagraphMeasure,
  type ImageBlock,
  type TableBlock,
  type TableMeasure,
} from '../layout-engine/types';
import type { BlockLookup } from '../layout-painter/index';
import type { HeaderFooterContent } from '../layout-painter/renderPage';
import type { FieldContext } from '../layout-painter/renderParagraph/positionRuns';
import { pageYToPt, pxToPt } from './coords';
import { colorToPdf } from './pdfText';
import { drawBorderLine, normalizePageBorderSide } from './pdfBorders';
import { drawParagraphAt, drawParagraphFragment } from './pdfParagraph';
import { drawTableAt, drawTableFragment } from './pdfTable';
import { collectImageSrcs, drawImageFragment, type ImageEmbedder } from './pdfImage';
import { collectFaces, type FaceRef } from './faces';
import type { FontProvider } from './fontProvider';
import type { PageBordersInput } from './types';

export interface DrawPageArgs {
  doc: PDFDocument;
  page: Page;
  blockLookup: BlockLookup;
  fonts: FontProvider;
  embedder: ImageEmbedder;
  field: FieldContext;
  backgroundColor?: string;
  pageBorders?: PageBordersInput;
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  onWarning?: (m: string) => void;
}

/** Faces referenced by header/footer content — for warm-up. */
export function hfFaces(hf: HeaderFooterContent | undefined): FaceRef[] {
  return hf ? collectFaces(hf.blocks) : [];
}

/** Blocks referenced by a page's fragments + its header/footer. */
function pageBlocks(
  page: Page,
  blockLookup: BlockLookup,
  header?: HeaderFooterContent,
  footer?: HeaderFooterContent
): FlowBlock[] {
  const blocks: FlowBlock[] = [];
  for (const fragment of page.fragments) {
    const entry =
      fragment.blockId !== undefined ? blockLookup.get(String(fragment.blockId)) : undefined;
    if (entry) blocks.push(entry.block);
  }
  if (header) blocks.push(...header.blocks);
  if (footer) blocks.push(...footer.blocks);
  return blocks;
}

/** Faces referenced by every block on a page — for font warm-up. */
export function pageFaces(page: Page, blockLookup: BlockLookup): FaceRef[] {
  return collectFaces(pageBlocks(page, blockLookup));
}

/** Image srcs referenced by a page (body + header/footer) — for image warm-up. */
export function pageImageSrcs(
  page: Page,
  blockLookup: BlockLookup,
  header?: HeaderFooterContent,
  footer?: HeaderFooterContent
): string[] {
  return collectImageSrcs(pageBlocks(page, blockLookup, header, footer));
}

export function drawPage(args: DrawPageArgs): PDFPage {
  const {
    doc,
    page,
    blockLookup,
    fonts,
    embedder,
    field,
    backgroundColor,
    pageBorders,
    header,
    footer,
    onWarning,
  } = args;
  const wPx = page.size.w;
  const hPx = page.size.h;
  const pdfPage = doc.addPage([pxToPt(wPx), pxToPt(hPx)]);

  if (backgroundColor) {
    pdfPage.drawRectangle({
      x: 0,
      y: 0,
      width: pxToPt(wPx),
      height: pxToPt(hPx),
      color: colorToPdf(backgroundColor, rgb(1, 1, 1)),
    });
  }

  // Header / footer content.
  const contentLeft = page.margins.left;
  const contentWidth = wPx - page.margins.left - page.margins.right;
  if (header) {
    drawHfContent(
      pdfPage,
      header,
      contentLeft,
      page.margins.header ?? 48,
      contentWidth,
      hPx,
      fonts,
      field,
      embedder,
      onWarning
    );
  }
  if (footer) {
    const footerTop = hPx - (page.margins.footer ?? 48) - footer.height;
    drawHfContent(
      pdfPage,
      footer,
      contentLeft,
      footerTop,
      contentWidth,
      hPx,
      fonts,
      field,
      embedder,
      onWarning
    );
  }

  // Page borders (w:pgBorders). 'firstPage'/'notFirstPage' gate by page number.
  if (pageBorders && pageBorderShows(pageBorders.display, page.number)) {
    drawPageBorders(pdfPage, pageBorders, wPx, hPx, page.margins);
  }
  // TODO(phase-1): footnote area, column separators.

  for (const fragment of page.fragments) {
    const entry =
      fragment.blockId !== undefined ? blockLookup.get(String(fragment.blockId)) : undefined;
    switch (fragment.kind) {
      case 'paragraph': {
        if (entry?.block.kind === 'paragraph' && entry.measure.kind === 'paragraph') {
          drawParagraphFragment({
            page: pdfPage,
            block: entry.block as ParagraphBlock,
            measure: entry.measure as ParagraphMeasure,
            fragment,
            pageHpx: hPx,
            fonts,
            field,
            embedder,
          });
        }
        break;
      }
      case 'image': {
        if (entry?.block.kind === 'image') {
          drawImageFragment(pdfPage, entry.block as ImageBlock, fragment, hPx, embedder);
        }
        break;
      }
      case 'table':
        if (entry?.block.kind === 'table' && entry.measure.kind === 'table') {
          drawTableFragment({
            page: pdfPage,
            block: entry.block as TableBlock,
            measure: entry.measure as TableMeasure,
            fragment,
            x: fragment.x,
            y: fragment.y,
            fromRow: fragment.fromRow,
            toRow: fragment.toRow,
            headerRowCount: fragment.headerRowCount,
            pageHpx: hPx,
            fonts,
            field,
            embedder,
            onWarning,
          });
        }
        break;
      case 'textBox':
        // Deferred to Phase 3 (handled = intentionally not drawn).
        break;
      default: {
        // Exhaustiveness: a NEW Fragment kind must be handled above or tsc fails.
        const _exhaustive: never = fragment;
        void _exhaustive;
      }
    }
  }
  return pdfPage;
}

/** Draw header/footer flow blocks (paragraphs + tables) starting at (left, top). */
function drawHfContent(
  page: PDFPage,
  hf: HeaderFooterContent,
  left: number,
  top: number,
  width: number,
  pageHpx: number,
  fonts: FontProvider,
  field: FieldContext,
  embedder: ImageEmbedder,
  onWarning?: (m: string) => void
): void {
  let y = top;
  for (let i = 0; i < hf.blocks.length; i++) {
    const b = hf.blocks[i];
    const m = hf.measures[i];
    if (!m) continue;
    if (b.kind === 'paragraph' && m.kind === 'paragraph') {
      drawParagraphAt({
        page,
        block: b,
        measure: m,
        x: left,
        y,
        width,
        pageHpx,
        fonts,
        field,
        embedder,
      });
      y += m.totalHeight;
    } else if (b.kind === 'table' && m.kind === 'table') {
      drawTableAt({
        page,
        block: b,
        measure: m,
        x: left,
        y,
        fromRow: 0,
        toRow: b.rows.length,
        pageHpx,
        fonts,
        field,
        embedder,
        onWarning,
      });
      y += m.totalHeight;
    }
  }
}

function pageBorderShows(display: PageBordersInput['display'], pageNumber: number): boolean {
  if (display === 'firstPage') return pageNumber === 1;
  if (display === 'notFirstPage') return pageNumber !== 1;
  return true;
}

/** Draw the four page borders, inset from the page edge ('page') or margins ('text'). */
function drawPageBorders(
  page: PDFPage,
  borders: PageBordersInput,
  wPx: number,
  hPx: number,
  margins: Page['margins']
): void {
  // OOXML default page-border offset is 24pt from the page edge.
  const PAGE_INSET = 24 / 0.75; // pt → px
  const inset = borders.offsetFrom === 'text';
  const left = pxToPt(inset ? margins.left : PAGE_INSET);
  const right = pxToPt(wPx - (inset ? margins.right : PAGE_INSET));
  const top = pageYToPt(inset ? margins.top : PAGE_INSET, hPx);
  const bottom = pageYToPt(hPx - (inset ? margins.bottom : PAGE_INSET), hPx);
  const side = (
    b: PageBordersInput['top'],
    a: { x: number; y: number },
    z: { x: number; y: number }
  ) => drawBorderLine(page, normalizePageBorderSide(b), a, z);
  side(borders.top, { x: left, y: top }, { x: right, y: top });
  side(borders.bottom, { x: left, y: bottom }, { x: right, y: bottom });
  side(borders.left, { x: left, y: bottom }, { x: left, y: top });
  side(borders.right, { x: right, y: bottom }, { x: right, y: top });
}
