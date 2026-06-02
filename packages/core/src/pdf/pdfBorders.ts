/**
 * Border → pdf-lib stroke translation, shared by paragraph and table cell
 * borders. Maps Word border styles to dash arrays and resolves the color.
 * `none`/`nil`/zero-width borders return null (nothing drawn).
 */

import { rgb, type RGB } from 'pdf-lib';
import { eighthsToPixels } from '../utils/units';
import { parseCssColor } from './cssColor';
import { pxToPt } from './coords';
import type { PageSink, SinkPoint } from './pageSink';

/** Common shape of {@link BorderStyle} and {@link CellBorderSpec}. */
export interface BorderLike {
  style?: string;
  width?: number;
  color?: string;
}

export interface Stroke {
  thickness: number;
  color: RGB;
  opacity: number;
  dashArray?: number[];
  /** Draw two parallel lines (Word `double`/`triple`/embossed styles). */
  double: boolean;
}

/** A page-border side as the adapter passes it (px width OR eighths-pt size; CSS or {rgb} color). */
export interface PageBorderSideLike {
  style?: string;
  width?: number;
  size?: number;
  color?: string | { rgb?: string } | null;
}

/** Normalize a permissive page-border side to a {@link BorderLike} (CSS color, px width). */
export function normalizePageBorderSide(b: PageBorderSideLike | undefined): BorderLike | undefined {
  if (!b) return undefined;
  const width = b.width ?? (b.size != null ? eighthsToPixels(b.size) : undefined);
  let color: string | undefined;
  if (typeof b.color === 'string') color = b.color;
  else if (b.color && typeof b.color === 'object' && b.color.rgb) color = `#${b.color.rgb}`;
  return { style: b.style, width, color };
}

/** Translate a border to a pdf-lib stroke, or null if it should not be drawn. */
export function strokeForBorder(b: BorderLike | undefined): Stroke | null {
  if (!b) return null;
  const style = (b.style ?? 'single').toLowerCase();
  if (style === 'none' || style === 'nil') return null;
  const widthPx = b.width && b.width > 0 ? b.width : 0.75;
  const thickness = pxToPt(widthPx);
  const c = parseCssColor(b.color);
  const color = c ? rgb(c.r, c.g, c.b) : rgb(0, 0, 0);
  const opacity = c?.alpha ?? 1;
  let dashArray: number[] | undefined;
  if (style === 'dashed' || style === 'dashsmallgap') dashArray = [thickness * 3, thickness * 2];
  else if (style === 'dotted') dashArray = [thickness, thickness];
  const double = style === 'double' || style === 'triple' || style.startsWith('threed');
  return { thickness, color, opacity, dashArray, double };
}

/**
 * Draw one border side. Shared by paragraph/table/page borders so the
 * stroke→drawLine shape (dash, thickness, and double-line offset) lives in one
 * place. `double` styles draw two parallel lines offset perpendicular to the run.
 */
export function drawBorderLine(
  page: PageSink,
  border: BorderLike | undefined,
  a: SinkPoint,
  z: SinkPoint
): void {
  const s = strokeForBorder(border);
  if (!s) return;
  // Perpendicular unit offset (handles horizontal and vertical sides).
  const dx = z.x - a.x;
  const dy = z.y - a.y;
  const len = Math.hypot(dx, dy) || 1;
  const nx = -dy / len;
  const ny = dx / len;
  const line = (off: number, thickness: number) =>
    page.drawLine({
      start: { x: a.x + nx * off, y: a.y + ny * off },
      end: { x: z.x + nx * off, y: z.y + ny * off },
      thickness,
      color: s.color,
      opacity: s.opacity,
      dashArray: s.dashArray,
    });
  if (s.double) {
    // Two thin lines whose total span ≈ the declared width (Word draws a double
    // border as two hairlines within the border weight, not two full-weight ones).
    const sub = s.thickness / 3;
    line(sub, sub);
    line(-sub, sub);
  } else {
    line(0, s.thickness);
  }
}
