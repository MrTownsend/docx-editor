/**
 * Table drawing for the PDF exporter.
 *
 * Matches the painter's collapsed-border rule (first row draws top, first column
 * draws left, every cell draws right + bottom — no double-thick shared edges),
 * cell shading, vertical alignment, cell padding (defaults {1,7,1,7}), gridSpan
 * column spanning, header-row repeat on continuation fragments, and recursion
 * into cell content (paragraphs and nested tables).
 */

import type { PageSink } from './pageSink';
import type {
  TableBlock,
  TableMeasure,
  TableFragment,
  TableCell,
  Measure,
  FlowBlock,
  ParagraphBlock,
  ParagraphMeasure,
  ImageBlock,
} from '../layout-engine/types';
import type { FieldContext } from '../layout-painter/renderParagraph/positionRuns';
import { pageYToPt, pxToPt } from './coords';
import { alphaOf, colorToPdf } from './pdfText';
import { drawBorderLine } from './pdfBorders';
import { drawParagraphAt } from './pdfParagraph';
import { drawImageFragment, type ImageEmbedder } from './pdfImage';
import type { FontProvider } from './fontProvider';

const DEFAULT_PAD = { top: 1, right: 7, bottom: 1, left: 7 };

export interface DrawTableArgs {
  page: PageSink;
  block: TableBlock;
  measure: TableMeasure;
  /** Table left x (px, page coords). */
  x: number;
  /** Table top y (px, page coords). */
  y: number;
  /** Row range to draw [fromRow, toRow). */
  fromRow: number;
  toRow: number;
  /** Header rows to repeat at the top (continuation fragments). */
  headerRowCount?: number;
  pageHpx: number;
  fonts: FontProvider;
  field: FieldContext;
  embedder: ImageEmbedder;
  onWarning?: (m: string) => void;
}

/** Draw a table fragment (page-positioned). */
export function drawTableFragment(args: DrawTableArgs & { fragment: TableFragment }): void {
  const { fragment, ...rest } = args;
  drawTableAt({ ...rest, x: fragment.x, y: fragment.y });
}

/** Space before paragraph i, collapsing against the previous paragraph's after. */
function spaceBefore(blocks: FlowBlock[], i: number, prevAfter: number): number {
  const blk = blocks[i];
  const before = blk?.kind === 'paragraph' ? (blk.attrs?.spacing?.before ?? 0) : 0;
  return Math.max(prevAfter, before);
}

/** Total content height of a cell's blocks (px), including collapsed inter-paragraph spacing. */
function cellContentHeight(blocks: FlowBlock[], measures: Measure[]): number {
  let h = 0;
  let prevAfter = 0;
  for (let i = 0; i < measures.length; i++) {
    const m = measures[i];
    if (!m) continue;
    if (i > 0) h += spaceBefore(blocks, i, prevAfter);
    if (m.kind === 'paragraph') h += m.totalHeight;
    else if (m.kind === 'table') h += m.totalHeight;
    else if (m.kind === 'image') h += m.height;
    const blk = blocks[i];
    prevAfter = blk?.kind === 'paragraph' ? (blk.attrs?.spacing?.after ?? 0) : 0;
  }
  return h;
}

/** Draw a table at an explicit (x, y) — used by fragments and nested tables. */
export function drawTableAt(args: DrawTableArgs): void {
  const {
    page,
    block,
    measure,
    x,
    y,
    fromRow,
    toRow,
    headerRowCount,
    pageHpx,
    fonts,
    field,
    embedder,
    onWarning,
  } = args;
  const columnWidths = measure.columnWidths ?? block.columnWidths ?? [];

  const rowsToDraw: number[] = [];
  for (let r = 0; r < (headerRowCount ?? 0); r++) rowsToDraw.push(r);
  for (let r = fromRow; r < toRow; r++) rowsToDraw.push(r);

  // Track columns covered by a rowSpan from an earlier row: occupied[col] is the
  // number of rows still to skip. Mirrors the painter's `occupiedColumns` logic
  // so cells under a vertical merge land in the right column and the merged cell
  // grows to span its rows.
  const occupied = new Map<number, number>();
  let rowTop = y;
  for (let idx = 0; idx < rowsToDraw.length; idx++) {
    const ri = rowsToDraw[idx];
    const row = block.rows[ri];
    const rowMeasure = measure.rows[ri];
    if (!row || !rowMeasure) continue;
    const rowHeight = rowMeasure.height;

    let colIndex = 0;
    for (let ci = 0; ci < row.cells.length; ci++) {
      // Skip columns still covered by a rowSpan from above.
      while ((occupied.get(colIndex) ?? 0) > 0) colIndex++;

      const cell = row.cells[ci];
      const cellMeasure = rowMeasure.cells[ci];
      const colSpan = cell.colSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      const cellX = x + sumRange(columnWidths, 0, colIndex);
      const cellW = sumRange(columnWidths, colIndex, colIndex + colSpan) || cellMeasure?.width || 0;
      // A vertically-merged cell grows to the summed height of its spanned rows.
      let cellH = rowHeight;
      for (let k = 1; k < rowSpan; k++) cellH += measure.rows[ri + k]?.height ?? 0;

      drawCell(page, cell, cellMeasure?.blocks ?? [], cellX, rowTop, cellW, cellH, {
        firstRow: idx === 0,
        firstCol: colIndex === 0,
        pageHpx,
        fonts,
        field,
        embedder,
        onWarning,
      });

      if (rowSpan > 1) {
        // Cover this column for the span's remaining rows. Set to `rowSpan` (not
        // rowSpan-1) because the end-of-row decrement below also runs for the
        // current row, leaving exactly `rowSpan-1` future rows skipped.
        for (let c = colIndex; c < colIndex + colSpan; c++) occupied.set(c, rowSpan);
      }
      colIndex += colSpan;
    }
    // One row consumed — decrement every active rowSpan counter.
    for (const [c, n] of occupied) {
      if (n > 0) occupied.set(c, n - 1);
    }
    rowTop += rowHeight;
  }
}

interface CellCtx {
  firstRow: boolean;
  firstCol: boolean;
  pageHpx: number;
  fonts: FontProvider;
  field: FieldContext;
  embedder: ImageEmbedder;
  onWarning?: (m: string) => void;
}

function drawCell(
  page: PageSink,
  cell: TableCell,
  blockMeasures: Measure[],
  cellX: number,
  cellY: number,
  cellW: number,
  cellH: number,
  ctx: CellCtx
): void {
  // Shading.
  if (cell.background) {
    page.drawRectangle({
      x: pxToPt(cellX),
      y: pageYToPt(cellY + cellH, ctx.pageHpx),
      width: pxToPt(cellW),
      height: pxToPt(cellH),
      color: colorToPdf(cell.background),
      opacity: alphaOf(cell.background),
    });
  }

  // Collapsed borders: top only on first row, left only on first column,
  // right + bottom always (the painter's shared-edge rule).
  const b = cell.borders ?? {};
  const L = pxToPt(cellX);
  const R = pxToPt(cellX + cellW);
  const T = pageYToPt(cellY, ctx.pageHpx);
  const B = pageYToPt(cellY + cellH, ctx.pageHpx);
  if (ctx.firstRow) drawBorderLine(page, b.top, { x: L, y: T }, { x: R, y: T });
  if (ctx.firstCol) drawBorderLine(page, b.left, { x: L, y: T }, { x: L, y: B });
  drawBorderLine(page, b.right, { x: R, y: T }, { x: R, y: B });
  drawBorderLine(page, b.bottom, { x: L, y: B }, { x: R, y: B });

  // Content, vertically aligned within the cell.
  const pad = cell.padding ?? DEFAULT_PAD;
  const contentX = cellX + pad.left;
  const contentW = cellW - pad.left - pad.right;
  const innerH = cellH - pad.top - pad.bottom;
  const contentH = cellContentHeight(cell.blocks, blockMeasures);
  let contentY = cellY + pad.top;
  if (cell.verticalAlign === 'center') contentY += Math.max(0, (innerH - contentH) / 2);
  else if (cell.verticalAlign === 'bottom') contentY += Math.max(0, innerH - contentH);

  let prevAfter = 0;
  for (let i = 0; i < cell.blocks.length; i++) {
    const blk = cell.blocks[i];
    const m = blockMeasures[i];
    if (!m) continue;
    // Collapse inter-paragraph spacing (max of prev `after` and this `before`).
    if (i > 0) contentY += spaceBefore(cell.blocks, i, prevAfter);
    prevAfter = blk.kind === 'paragraph' ? (blk.attrs?.spacing?.after ?? 0) : 0;
    if (blk.kind === 'paragraph' && m.kind === 'paragraph') {
      drawParagraphAt({
        page,
        block: blk as ParagraphBlock,
        measure: m as ParagraphMeasure,
        x: contentX,
        y: contentY,
        width: contentW,
        pageHpx: ctx.pageHpx,
        fonts: ctx.fonts,
        field: ctx.field,
        embedder: ctx.embedder,
      });
      contentY += m.totalHeight;
    } else if (blk.kind === 'table' && m.kind === 'table') {
      drawTableAt({
        page,
        block: blk as TableBlock,
        measure: m,
        x: contentX,
        y: contentY,
        fromRow: 0,
        toRow: (blk as TableBlock).rows.length,
        pageHpx: ctx.pageHpx,
        fonts: ctx.fonts,
        field: ctx.field,
        embedder: ctx.embedder,
        onWarning: ctx.onWarning,
      });
      contentY += m.totalHeight;
    } else if (blk.kind === 'image' && m.kind === 'image') {
      void drawImageFragment(
        page,
        blk as ImageBlock,
        {
          kind: 'image',
          blockId: (blk as ImageBlock).id,
          x: contentX,
          y: contentY,
          width: m.width,
          height: m.height,
        },
        ctx.pageHpx,
        ctx.embedder
      );
      contentY += m.height;
    }
  }
}

function sumRange(arr: number[], from: number, to: number): number {
  let s = 0;
  for (let i = from; i < to && i < arr.length; i++) s += arr[i] || 0;
  return s;
}
