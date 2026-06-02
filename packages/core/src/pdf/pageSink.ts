/**
 * `PageSink` — the minimal subset of pdf-lib's `PDFPage` the exporter draws to.
 *
 * Abstracting the draw surface lets tests pass a {@link RecordingSink} that
 * captures every op as plain data, so golden snapshots can assert the exact
 * positions/text/rects the exporter produces. Any drift (a painter rule change,
 * a coordinate bug) shows up as a readable diff in CI — the parity tripwire.
 *
 * A real `PDFPage` is structurally assignable to `PageSink` (its draw methods are
 * a superset), so production code passes the real page unchanged.
 */

import type { PDFFont, PDFImage, RGB, Rotation } from 'pdf-lib';

export interface SinkPoint {
  x: number;
  y: number;
}

export interface SinkTextOptions {
  x: number;
  y: number;
  size: number;
  font: PDFFont;
  color?: RGB;
  opacity?: number;
}

export interface SinkLineOptions {
  start: SinkPoint;
  end: SinkPoint;
  thickness: number;
  color?: RGB;
  opacity?: number;
  dashArray?: number[];
}

export interface SinkRectOptions {
  x: number;
  y: number;
  width: number;
  height: number;
  color?: RGB;
  borderColor?: RGB;
  borderWidth?: number;
  opacity?: number;
}

export interface SinkImageOptions {
  x: number;
  y: number;
  width: number;
  height: number;
  rotate?: Rotation;
}

export interface PageSink {
  drawText(text: string, options: SinkTextOptions): void;
  drawLine(options: SinkLineOptions): void;
  drawRectangle(options: SinkRectOptions): void;
  drawImage(image: PDFImage, options: SinkImageOptions): void;
}

/** One recorded draw op (positions rounded for stable, readable golden diffs). */
export type DrawOp =
  | { op: 'text'; text: string; x: number; y: number; size: number; font?: string }
  | {
      op: 'line';
      x1: number;
      y1: number;
      x2: number;
      y2: number;
      thickness: number;
      dashed: boolean;
    }
  | {
      op: 'rect';
      x: number;
      y: number;
      width: number;
      height: number;
      filled: boolean;
      bordered: boolean;
    }
  | { op: 'image'; x: number; y: number; width: number; height: number; rotate: number };

const r2 = (n: number): number => Math.round(n * 100) / 100;

/**
 * A {@link PageSink} that records draw ops as plain data instead of writing PDF.
 * Used by golden snapshot tests — no pdf-lib document or network needed.
 */
export class RecordingSink implements PageSink {
  readonly ops: DrawOp[] = [];

  drawText(text: string, o: SinkTextOptions): void {
    const font = (o.font as { name?: string }).name;
    this.ops.push({ op: 'text', text, x: r2(o.x), y: r2(o.y), size: r2(o.size), font });
  }
  drawLine(o: SinkLineOptions): void {
    this.ops.push({
      op: 'line',
      x1: r2(o.start.x),
      y1: r2(o.start.y),
      x2: r2(o.end.x),
      y2: r2(o.end.y),
      thickness: r2(o.thickness),
      dashed: !!o.dashArray,
    });
  }
  drawRectangle(o: SinkRectOptions): void {
    this.ops.push({
      op: 'rect',
      x: r2(o.x),
      y: r2(o.y),
      width: r2(o.width),
      height: r2(o.height),
      filled: o.color !== undefined,
      bordered: o.borderWidth !== undefined,
    });
  }
  drawImage(_image: PDFImage, o: SinkImageOptions): void {
    const angle =
      o.rotate && typeof (o.rotate as { angle?: number }).angle === 'number'
        ? (o.rotate as { angle: number }).angle
        : 0;
    this.ops.push({
      op: 'image',
      x: r2(o.x),
      y: r2(o.y),
      width: r2(o.width),
      height: r2(o.height),
      rotate: r2(angle),
    });
  }
}
