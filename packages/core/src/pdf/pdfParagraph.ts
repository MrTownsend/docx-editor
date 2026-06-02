/**
 * Paragraph drawing for the PDF exporter.
 *
 * For each measured line: position runs with the embedded-face metric (so x
 * matches the drawn glyphs), draw paragraph shading + borders, the list marker,
 * tab leaders, and the text runs. `drawParagraphAt` is the position-agnostic core
 * (also used by table cells); `drawParagraphFragment` is the page-fragment wrapper.
 */

import { rgb } from 'pdf-lib';
import type { PageSink } from './pageSink';
import type {
  ParagraphBlock,
  ParagraphMeasure,
  ParagraphFragment,
  ParagraphBorders,
  Run,
} from '../layout-engine/types';
import {
  positionRunsInLine,
  type FieldContext,
} from '../layout-painter/renderParagraph/positionRuns';
import { baselineFromTop, pageYToPt, pxToPt, textBaselinePt } from './coords';
import { alphaOf, colorToPdf, drawTextRun, faceFor, safeWidth } from './pdfText';
import { drawBorderLine } from './pdfBorders';
import { drawInlineImage, type ImageEmbedder } from './pdfImage';
import { getListMarkerInlineWidth } from '../layout-bridge/measuring/listMarkerWidth';
import type { FontProvider } from './fontProvider';

const LEADER_CHAR: Record<string, string> = {
  dot: '.',
  hyphen: '-',
  underscore: '_',
  middleDot: '·',
  heavy: '_',
};

export interface DrawParagraphArgs {
  page: PageSink;
  block: ParagraphBlock;
  measure: ParagraphMeasure;
  fragment: ParagraphFragment;
  pageHpx: number;
  fonts: FontProvider;
  field: FieldContext;
  embedder?: ImageEmbedder;
}

export interface DrawParagraphAtArgs {
  page: PageSink;
  block: ParagraphBlock;
  measure: ParagraphMeasure;
  /** Content-box left x (px, page coords) of the paragraph. */
  x: number;
  /** Top y (px, page coords) of the first drawn line. */
  y: number;
  /** Content-box width (px). */
  width: number;
  /** Line range [fromLine, toLine); defaults to the whole measure. */
  fromLine?: number;
  toLine?: number;
  pageHpx: number;
  fonts: FontProvider;
  field: FieldContext;
  /** For inline images inside the paragraph (warmed-up embedder). */
  embedder?: ImageEmbedder;
}

/** Draw a paragraph's lines at an explicit (x, y, width). Returns the y after the last line. */
export function drawParagraphAt(args: DrawParagraphAtArgs): number {
  const { page, block, measure, x, y, width, pageHpx, fonts, field, embedder } = args;
  const fromLine = args.fromLine ?? 0;
  const toLine = args.toLine ?? measure.lines.length;
  const a = block.attrs;
  const indent = a?.indent ?? {};
  const indentLeft = indent.left ?? 0;
  const indentRight = indent.right ?? 0;
  const availableWidth = width - indentLeft - indentRight;
  const firstLineIndentPx = indent.hanging ? -indent.hanging : (indent.firstLine ?? 0);
  const lineRightEdgePx = width - indentRight;
  const alignment = a?.alignment;

  const measureText = (run: Run, text: string): number => {
    const size = ('fontSize' in run && run.fontSize) || 11;
    return safeWidth(faceFor(run, text, fonts), text, size);
  };

  // Total drawn height (for shading/border boxes).
  let drawnHeight = 0;
  for (let li = fromLine; li < toLine; li++) {
    const l = measure.lines[li];
    if (l) drawnHeight += (l.floatSkipBefore ?? 0) + l.lineHeight;
  }

  // Paragraph shading behind the text.
  if (a?.shading) {
    page.drawRectangle({
      x: pxToPt(x),
      y: pageYToPt(y + drawnHeight, pageHpx),
      width: pxToPt(width),
      height: pxToPt(drawnHeight),
      color: colorToPdf(a.shading),
      opacity: alphaOf(a.shading),
    });
  }
  if (a?.borders) drawParagraphBorders(page, a.borders, x, y, width, drawnHeight, pageHpx);

  const hasMarker = !!a?.listMarker && !a?.listMarkerHidden;
  // For a hanging list the marker sits in the hung area and body text starts at
  // the left indent (firstLineIndent → 0). For a NON-hanging list the measurer
  // reserved `getListMarkerInlineWidth` on the first line, so body text must
  // start that much past the marker (else marker and text overlap).
  const hanging = (indent.hanging ?? 0) > 0;
  const markerSlot = hasMarker && !hanging ? getListMarkerInlineWidth(block) : 0;
  const markerFirstLineIndent = hasMarker
    ? hanging
      ? 0
      : firstLineIndentPx + markerSlot
    : firstLineIndentPx;
  let lineTop = y;
  for (let li = fromLine; li < toLine; li++) {
    const line = measure.lines[li];
    if (!line) continue;
    lineTop += line.floatSkipBefore ?? 0;
    const isFirst = li === 0;
    const isLast = li === measure.lines.length - 1;

    const positioned = positionRunsInLine(block, line, alignment, {
      availableWidth,
      isFirstLine: isFirst,
      isLastLine: isLast,
      paragraphEndsWithLineBreak: false,
      tabStops: a?.tabs,
      leftIndentPx: indentLeft,
      firstLineIndentPx: isFirst ? markerFirstLineIndent : firstLineIndentPx,
      lineRightEdgePx,
      field,
      measureText,
    });

    if (hasMarker && isFirst) {
      const markerX = x + indentLeft + firstLineIndentPx;
      const markerRun = {
        kind: 'text',
        text: a!.listMarker!,
        fontFamily: a!.listMarkerFontFamily,
        fontSize: a!.listMarkerFontSize,
        color: a!.listMarkerRevision
          ? undefined
          : block.runs[0] && 'color' in block.runs[0]
            ? block.runs[0].color
            : undefined,
      } as Run;
      drawTextRun({
        page,
        text: a!.listMarker!,
        xPt: pxToPt(markerX),
        baselinePt: textBaselinePt(lineTop, line, pageHpx),
        widthPx: 0,
        run: markerRun,
        fonts,
      });
    }

    for (const pr of positioned.runs) {
      const xPt = pxToPt(x + pr.x);
      if (pr.kind === 'text' || pr.kind === 'field' || pr.kind === 'other') {
        const shift = 'positionPx' in pr.run && pr.run.positionPx ? pr.run.positionPx : 0;
        drawTextRun({
          page,
          text: pr.text ?? '',
          xPt,
          baselinePt: textBaselinePt(lineTop, line, pageHpx, shift),
          widthPx: pr.width,
          lineTopPx: lineTop,
          lineHeightPx: line.lineHeight,
          pageHpx,
          wordSpacingPt: pr.kind === 'text' ? pxToPt(positioned.wordSpacingPx) : 0,
          run: pr.run,
          fonts,
        });
      } else if (pr.kind === 'image' && embedder && pr.run.kind === 'image') {
        drawInlineImage(page, pr.run, x + pr.x, lineTop, line, pageHpx, embedder);
      } else if (pr.kind === 'tab' && pr.tabLeader && pr.tabLeader !== 'none') {
        drawLeader(page, pr.tabLeader, xPt, pr.width, lineTop, line, pageHpx, fonts, pr.run);
      }
    }
    lineTop += line.lineHeight;
  }
  return lineTop;
}

export function drawParagraphFragment(args: DrawParagraphArgs): void {
  const { page, block, measure, fragment, pageHpx, fonts, field, embedder } = args;
  drawParagraphAt({
    page,
    block,
    measure,
    embedder,
    x: fragment.x,
    y: fragment.y,
    width: fragment.width,
    fromLine: fragment.fromLine,
    toLine: fragment.toLine,
    pageHpx,
    fonts,
    field,
  });
}

function drawParagraphBorders(
  page: PageSink,
  borders: ParagraphBorders,
  x: number,
  y: number,
  width: number,
  height: number,
  pageHpx: number
): void {
  const left = pxToPt(x);
  const right = pxToPt(x + width);
  const top = pageYToPt(y, pageHpx);
  const bottom = pageYToPt(y + height, pageHpx);
  drawBorderLine(page, borders.top, { x: left, y: top }, { x: right, y: top });
  drawBorderLine(page, borders.bottom, { x: left, y: bottom }, { x: right, y: bottom });
  drawBorderLine(page, borders.left, { x: left, y: top }, { x: left, y: bottom });
  drawBorderLine(page, borders.right, { x: right, y: top }, { x: right, y: bottom });
}

function drawLeader(
  page: PageSink,
  leader: string,
  xPt: number,
  widthPx: number,
  lineTop: number,
  line: { ascent: number; descent: number; lineHeight: number },
  pageHpx: number,
  fonts: FontProvider,
  run: Run
): void {
  const ch = LEADER_CHAR[leader] ?? '.';
  const size = ('fontSize' in run && run.fontSize) || 11;
  // Use the run's own face (leader glyph metrics differ per font).
  const face = faceFor(run, ch, fonts);
  const chW = safeWidth(face, ch, size);
  if (chW <= 0) return;
  const count = Math.max(0, Math.floor(pxToPt(widthPx) / chW));
  if (count === 0) return;
  page.drawText(ch.repeat(count), {
    x: xPt,
    y: pageYToPt(lineTop + baselineFromTop(line), pageHpx),
    size,
    font: face,
    color: rgb(0, 0, 0),
  });
}
