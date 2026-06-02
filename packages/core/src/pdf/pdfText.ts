/**
 * Text + decoration drawing for the PDF exporter.
 *
 * Draws one positioned run (color, super/subscript, underline/strike) using the
 * embedded face, guarding WinAnsi-only fallback faces against non-Latin text by
 * swapping to the Unicode fallback. Line/baseline math lives in `coords`; run x
 * comes from `positionRunsInLine` (measured with this same face).
 */

import { rgb, type PDFFont, type RGB } from 'pdf-lib';
import type { PageSink } from './pageSink';
import type { Run } from '../layout-engine/types';
import { parseCssColor } from './cssColor';
import { canEncode, type FontProvider, type FontStyle } from './fontProvider';
import { pageYToPt, pxToPt } from './coords';

const BLACK = rgb(0, 0, 0);

/** CSS color string → pdf-lib RGB, defaulting to black. */
export function colorToPdf(css: string | undefined, fallback: RGB = BLACK): RGB {
  const p = parseCssColor(css);
  return p ? rgb(p.r, p.g, p.b) : fallback;
}

/** Alpha (0–1) of a CSS color, or 1 if opaque/unparsed. */
export function alphaOf(css: string | undefined): number {
  return parseCssColor(css)?.alpha ?? 1;
}

const styleOf = (run: Run): FontStyle => ({
  bold: 'bold' in run && run.bold,
  italic: 'italic' in run && run.italic,
});

/** Drop characters `font` cannot encode (last resort so the export never throws). */
function stripUnencodable(font: PDFFont, text: string): string {
  let out = '';
  for (const ch of text) {
    try {
      font.encodeText(ch);
      out += ch;
    } catch {
      /* skip un-encodable glyph */
    }
  }
  return out;
}

/** Resolve the face to draw a run with, swapping to Unicode fallback if it can't encode. */
export function faceFor(run: Run, text: string, fonts: FontProvider): PDFFont {
  const family = ('fontFamily' in run && run.fontFamily) || 'Calibri';
  const face = fonts.getFontSync(family, styleOf(run));
  if (text && !canEncode(face, text)) return fonts.getUnicodeFallbackSync();
  return face;
}

/**
 * Advance width that never throws on un-encodable glyphs. `widthOfTextAtSize`
 * encodes internally, so a WinAnsi-only fallback would throw on non-Latin; we
 * measure the encodable subset (an estimate for the dropped glyphs).
 */
export function safeWidth(font: PDFFont, text: string, sizePt: number): number {
  try {
    return font.widthOfTextAtSize(text, sizePt);
  } catch {
    let w = 0;
    for (const ch of text) {
      try {
        w += font.widthOfTextAtSize(ch, sizePt);
      } catch {
        w += sizePt * 0.5; // rough advance for a dropped glyph
      }
    }
    return w;
  }
}

export interface DrawRunArgs {
  page: PageSink;
  text: string;
  /** Absolute page x (pt). */
  xPt: number;
  /** Glyph baseline y (pt). */
  baselinePt: number;
  /** Advance width of the run in px (for underline/strike/highlight extents). */
  widthPx: number;
  /** Line box top in px from the page top (for the full-height highlight band). */
  lineTopPx?: number;
  /** Line box height in px (highlight spans the whole line, matching the painter). */
  lineHeightPx?: number;
  /** Page height in px (to flip the highlight rect into PDF coords). */
  pageHpx?: number;
  /** Extra pt added to each inter-word space when justifying (0 = none). */
  wordSpacingPt?: number;
  run: Run;
  fonts: FontProvider;
}

/** Default hyperlink color when the run carries no explicit color (Word's #0563C1). */
const HYPERLINK_BLUE = rgb(0x05 / 255, 0x63 / 255, 0xc1 / 255);

/**
 * Draw a run's glyphs, applying per-glyph letter spacing and/or per-word justify
 * spacing (pdf-lib's `drawText` supports neither, so we position each unit). A
 * single `drawText` is used when neither applies (the common, fast path).
 */
function drawRunGlyphs(
  page: PageSink,
  text: string,
  x: number,
  y: number,
  size: number,
  font: PDFFont,
  color: RGB,
  opacity: number,
  letterSpacingPt: number,
  wordSpacingPt: number
): void {
  if (letterSpacingPt === 0 && (wordSpacingPt <= 0 || !text.includes(' '))) {
    page.drawText(text, { x, y, size, font, color, opacity });
    return;
  }
  const spaceW = safeWidth(font, ' ', size);
  let cx = x;
  for (const ch of text) {
    if (ch === ' ') {
      cx += spaceW + wordSpacingPt + letterSpacingPt;
      continue;
    }
    page.drawText(ch, { x: cx, y, size, font, color, opacity });
    cx += safeWidth(font, ch, size) + letterSpacingPt;
  }
}

/**
 * Draw a single text/field run with its decorations: highlight background,
 * glyphs, underline, strike. Super/subscript shrink the size (~0.75em, matching
 * the painter) and shift the baseline. Hidden runs never reach here.
 */
export function drawTextRun(args: DrawRunArgs): void {
  const { page, text, xPt, baselinePt, widthPx, run, fonts } = args;
  if (!text) return;
  const face = faceFor(run, text, fonts);

  const basePt = ('fontSize' in run && run.fontSize) || 11;
  const isSuper = 'superscript' in run && run.superscript;
  const isSub = 'subscript' in run && run.subscript;
  // Match the painter's `font-size: 0.75em` for super/subscript.
  const sizePt = isSuper || isSub ? basePt * 0.75 : basePt;
  const shiftPt = isSuper ? basePt * 0.33 : isSub ? -basePt * 0.18 : 0;

  // Hyperlinks with no explicit run color/underline get Word's default blue +
  // underline (the painter injects this; it's not in the model's color/underline).
  const link = 'hyperlink' in run ? run.hyperlink : undefined;
  const linkDefault = !!link && !link.noDefaultStyle;
  const explicitColor = 'color' in run ? run.color : undefined;
  const color = explicitColor
    ? colorToPdf(explicitColor)
    : linkDefault
      ? HYPERLINK_BLUE
      : colorToPdf(undefined);
  const opacity = alphaOf(explicitColor);
  const letterSpacingPt =
    'letterSpacing' in run && run.letterSpacing ? pxToPt(run.letterSpacing) : 0;
  const wordSpacingPt = args.wordSpacingPt ?? 0;

  // Decoration extent = the run's box minus any trailing-space widening, so a
  // justified run's underline/strike stops at the last glyph, not in the gap.
  const trailingSpaces = text.length - text.trimEnd().length;
  const spaceAdvPt = trailingSpaces > 0 ? safeWidth(face, ' ', sizePt) : 0;
  const decorWidthPt = Math.max(
    0,
    pxToPt(widthPx) - trailingSpaces * (spaceAdvPt + wordSpacingPt + letterSpacingPt)
  );

  // Highlight background (w:highlight) — spans the full LINE box (matching the
  // painter, which pads the highlight to line height), not just the glyph band.
  const highlight = 'highlight' in run ? run.highlight : undefined;
  const hlParsed = highlight ? parseCssColor(highlight) : undefined;
  if (hlParsed && args.lineTopPx !== undefined && args.lineHeightPx !== undefined) {
    page.drawRectangle({
      x: xPt,
      y: pageYToPt(args.lineTopPx + args.lineHeightPx, args.pageHpx ?? 0),
      width: decorWidthPt,
      height: pxToPt(args.lineHeightPx),
      color: rgb(hlParsed.r, hlParsed.g, hlParsed.b),
      opacity: hlParsed.alpha,
    });
  }

  const y = baselinePt + shiftPt;
  try {
    drawRunGlyphs(page, text, xPt, y, sizePt, face, color, opacity, letterSpacingPt, wordSpacingPt);
  } catch {
    // The chosen face can't encode some glyph. Try the Unicode fallback; if even
    // that can't (no Unicode face was bundled), drop the un-encodable glyphs so
    // the export never throws.
    const fb = fonts.getUnicodeFallbackSync();
    const safe = canEncode(fb, text) ? text : stripUnencodable(fb, text);
    if (safe) page.drawText(safe, { x: xPt, y, size: sizePt, font: fb, color, opacity });
  }

  // Underline (explicit, or default for hyperlinks).
  const underline = 'underline' in run ? run.underline : undefined;
  if (underline || linkDefault) {
    const uColor =
      typeof underline === 'object' && underline.color ? colorToPdf(underline.color) : color;
    const thickness = Math.max(0.5, sizePt * 0.06);
    const uy = baselinePt + shiftPt - sizePt * 0.12;
    page.drawLine({
      start: { x: xPt, y: uy },
      end: { x: xPt + decorWidthPt, y: uy },
      thickness,
      color: uColor,
      opacity,
    });
    if (typeof underline === 'object' && underline.style === 'double') {
      page.drawLine({
        start: { x: xPt, y: uy - thickness * 1.5 },
        end: { x: xPt + decorWidthPt, y: uy - thickness * 1.5 },
        thickness,
        color: uColor,
        opacity,
      });
    }
  }
  // Strikethrough.
  if ('strike' in run && run.strike) {
    const sy = baselinePt + shiftPt + sizePt * 0.28;
    page.drawLine({
      start: { x: xPt, y: sy },
      end: { x: xPt + decorWidthPt, y: sy },
      thickness: Math.max(0.5, sizePt * 0.06),
      color,
      opacity,
    });
  }
}
