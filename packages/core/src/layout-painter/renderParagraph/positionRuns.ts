/**
 * Pure run positioning for a single line.
 *
 * Extracts the geometry half of {@link renderLine} — the `currentX` cursor, tab
 * math, indent, alignment, and justify gap distribution — into a DOM-free
 * function so the PDF exporter and the painter can position runs from ONE source
 * of logic. `measureText` is injected: the painter passes its canvas measurer;
 * the exporter passes pdf-lib's embedded-face metric so run-x and the drawn
 * glyphs share one metric source (no intra-line drift).
 *
 * Returned `x`/`width` are in content-area-relative px (i.e. they already include
 * the paragraph's left indent); the exporter adds `fragment.x` to get the page x.
 */

import type { ParagraphBlock, MeasuredLine, Run, TabStop } from '../../layout-engine/types';
import {
  calculateTabWidth,
  type TabContext,
  type TabLeader,
  type TabStop as TabCalcStop,
} from '../../prosemirror/utils/tabCalculator';
import { sliceRunsForLine } from './line';
import { isTextRun, isTabRun, isImageRun, isLineBreakRun, isFieldRun } from './shared';

export type PositionedRunKind = 'text' | 'tab' | 'image' | 'field' | 'lineBreak' | 'other';

export interface PositionedRun {
  run: Run;
  kind: PositionedRunKind;
  /** Left x in content-area-relative px (includes left indent). */
  x: number;
  /** Advance width in px (incl. letterSpacing + horizontalScale). */
  width: number;
  /** Effective text to draw (caps-transformed; field-resolved). */
  text?: string;
  /** Leader fill for a tab run. */
  tabLeader?: TabLeader;
}

export interface PositionedLine {
  runs: PositionedRun[];
  /** Whether the painter would flex-anchor this line (right-tab/TOC). */
  isFlexAnchored: boolean;
  /**
   * Extra px added to each inter-word space when justifying. Run `x`/`width`
   * already account for it across run boundaries; a renderer must additionally
   * widen the spaces drawn WITHIN a multi-word run by this amount (the painter
   * gets this for free from CSS `text-align: justify`). 0 when not justified.
   */
  wordSpacingPx: number;
}

/** Resolve a field run's display text (PAGE/NUMPAGES/DATE/TIME → value, else fallback). */
export interface FieldContext {
  pageNumber: number;
  totalPages: number;
  /** ISO date string to resolve DATE/TIME against (injected for determinism). */
  now?: string;
}

export interface PositionRunsOptions {
  availableWidth: number;
  isFirstLine: boolean;
  isLastLine: boolean;
  paragraphEndsWithLineBreak: boolean;
  tabStops?: TabStop[];
  leftIndentPx?: number;
  firstLineIndentPx?: number;
  lineRightEdgePx?: number;
  field?: FieldContext;
  /** Advance width of `text` at the run's font, in px (font size in points). */
  measureText: (run: Run, text: string) => number;
}

function convertTabStop(stop: TabStop): TabCalcStop {
  return { val: stop.val, pos: stop.pos, leader: stop.leader as TabCalcStop['leader'] };
}

function resolveFieldText(run: Run, field?: FieldContext): string {
  if (!isFieldRun(run)) return '';
  if (run.fieldType === 'PAGE' && field) return String(field.pageNumber);
  if (run.fieldType === 'NUMPAGES' && field) return String(field.totalPages);
  if ((run.fieldType === 'DATE' || run.fieldType === 'TIME') && field?.now) {
    const d = new Date(field.now);
    return run.fieldType === 'DATE' ? d.toLocaleDateString() : d.toLocaleTimeString();
  }
  return run.fallback ?? '';
}

/** Text actually drawn for a run, after caps transform. */
export function effectiveRunText(run: Run, field?: FieldContext): string {
  let text = isTextRun(run) ? run.text : isFieldRun(run) ? resolveFieldText(run, field) : '';
  // allCaps and smallCaps both render as uppercase glyphs. (smallCaps' reduced
  // cap height for source-lowercase letters is approximated as full caps — pdf
  // has no font-variant; uppercasing reads far closer to Word than mixed case.)
  if (('allCaps' in run && run.allCaps) || ('smallCaps' in run && run.smallCaps)) {
    text = text.toUpperCase();
  }
  return text;
}

/** Advance width of a run including letterSpacing + horizontalScale. */
function runAdvance(run: Run, text: string, measureText: (r: Run, t: string) => number): number {
  let w = measureText(run, text);
  const ls = 'letterSpacing' in run ? run.letterSpacing : undefined;
  if (ls && text.length > 0) w += ls * text.length;
  const hs = 'horizontalScale' in run ? run.horizontalScale : undefined;
  if (hs && hs !== 100) w *= hs / 100;
  return w;
}

/** Sum advance widths of the runs following a tab, up to the next tab/break. */
function followingWidth(
  runs: Run[],
  tabIndex: number,
  measureText: (r: Run, t: string) => number,
  field?: FieldContext
): number {
  let w = 0;
  for (let i = tabIndex + 1; i < runs.length; i++) {
    const r = runs[i];
    if (isTabRun(r) || isLineBreakRun(r)) break;
    if (isTextRun(r) || isFieldRun(r)) w += runAdvance(r, effectiveRunText(r, field), measureText);
    else if (isImageRun(r)) w += r.width || 0;
  }
  return w;
}

const RIGHT_EDGE_EPSILON_PX = 0.5;

/**
 * Position the runs of a line. Hidden runs are skipped (Word print semantics).
 */
export function positionRunsInLine(
  block: ParagraphBlock,
  line: MeasuredLine,
  alignment: 'left' | 'center' | 'right' | 'justify' | undefined,
  opts: PositionRunsOptions
): PositionedLine {
  const sliced = sliceRunsForLine(block, line).filter((r) => !('hidden' in r && r.hidden));
  const leftIndent = opts.leftIndentPx ?? 0;
  const firstIndent = opts.isFirstLine ? (opts.firstLineIndentPx ?? 0) : 0;

  const hasTabs = sliced.some(isTabRun);
  const tabContext: TabContext | undefined = hasTabs
    ? { explicitStops: opts.tabStops?.map(convertTabStop), leftIndent: Math.round(leftIndent * 15) }
    : undefined;

  let currentX = leftIndent + firstIndent;
  const positioned: PositionedRun[] = [];
  let isFlexAnchored = false;

  for (let i = 0; i < sliced.length; i++) {
    const run = sliced[i];

    if (isTabRun(run) && tabContext) {
      const follow = followingWidth(sliced, i, opts.measureText, opts.field);
      const followText = sliced
        .slice(i + 1)
        .filter(isTextRun)
        .map((r) => r.text)
        .join('');
      const decimalIndex = followText.indexOf('.');
      const decimalPrefixWidth =
        decimalIndex >= 0 ? opts.measureText(run, followText.slice(0, decimalIndex)) : 0;
      const tab = calculateTabWidth(currentX, tabContext, {
        followingWidth: follow,
        decimalPrefixWidth,
      });

      // Right-tab/TOC anchor: an end-aligned tab whose stop is at the line's
      // right edge pins trailing content to that edge.
      const rightEdge = opts.lineRightEdgePx;
      let hasFollowingTab = false;
      for (let j = i + 1; j < sliced.length; j++) {
        if (isLineBreakRun(sliced[j])) break;
        if (isTabRun(sliced[j])) {
          hasFollowingTab = true;
          break;
        }
      }
      if (
        rightEdge !== undefined &&
        tab.alignment === 'end' &&
        !hasFollowingTab &&
        currentX + tab.width + follow >= rightEdge - RIGHT_EDGE_EPSILON_PX
      ) {
        isFlexAnchored = true;
        const tabStart = currentX;
        const tabWidth = Math.max(1, rightEdge - follow - currentX);
        positioned.push({ run, kind: 'tab', x: tabStart, width: tabWidth, tabLeader: tab.leader });
        currentX = rightEdge - follow;
        // Trailing runs flush to the right edge.
        for (let j = i + 1; j < sliced.length; j++) {
          const next = sliced[j];
          if (isTabRun(next) || isLineBreakRun(next)) break;
          const text = effectiveRunText(next, opts.field);
          const w = isImageRun(next) ? next.width || 0 : runAdvance(next, text, opts.measureText);
          positioned.push({ run: next, kind: runKind(next), x: currentX, width: w, text });
          currentX += w;
        }
        return { runs: positioned, isFlexAnchored, wordSpacingPx: 0 };
      }

      let tabWidth = tab.width;
      if (rightEdge !== undefined && currentX + tabWidth + follow > rightEdge) {
        tabWidth = Math.max(1, rightEdge - currentX - follow);
      }
      positioned.push({ run, kind: 'tab', x: currentX, width: tabWidth, tabLeader: tab.leader });
      currentX += tabWidth;
    } else if (isImageRun(run)) {
      const w = run.width || 0;
      positioned.push({ run, kind: 'image', x: currentX, width: w });
      if (run.displayMode !== 'block' && run.wrapType !== 'topAndBottom') currentX += w;
    } else if (isLineBreakRun(run)) {
      positioned.push({ run, kind: 'lineBreak', x: currentX, width: 0 });
    } else {
      // text / field / other
      const text = effectiveRunText(run, opts.field);
      const w = runAdvance(run, text, opts.measureText);
      positioned.push({ run, kind: runKind(run), x: currentX, width: w, text });
      currentX += w;
    }
  }

  const wordSpacingPx = applyAlignment(
    positioned,
    currentX - (leftIndent + firstIndent),
    opts,
    alignment,
    hasTabs
  );
  return { runs: positioned, isFlexAnchored, wordSpacingPx };
}

function runKind(run: Run): PositionedRunKind {
  if (isTextRun(run)) return 'text';
  if (isFieldRun(run)) return 'field';
  if (isImageRun(run)) return 'image';
  if (isLineBreakRun(run)) return 'lineBreak';
  if (isTabRun(run)) return 'tab';
  return 'other';
}

/**
 * Shift/space runs for center/right/justify. Lines containing tab stops keep
 * their tab-driven positions (Word does not center/right tab layouts).
 */
function applyAlignment(
  runs: PositionedRun[],
  contentWidth: number,
  opts: PositionRunsOptions,
  alignment: 'left' | 'center' | 'right' | 'justify' | undefined,
  hasTabs: boolean
): number {
  if (!alignment || alignment === 'left' || hasTabs || runs.length === 0) return 0;
  const slack = opts.availableWidth - contentWidth;
  if (slack <= 0) return 0;

  if (alignment === 'center') {
    for (const r of runs) r.x += slack / 2;
    return 0;
  }
  if (alignment === 'right') {
    for (const r of runs) r.x += slack;
    return 0;
  }
  // justify (not the last line unless it ends with a break): distribute slack
  // across inter-word gaps. Run starts shift so run boundaries land correctly;
  // the returned per-gap value lets the renderer widen intra-run spaces too.
  if (opts.isLastLine && !opts.paragraphEndsWithLineBreak) return 0;
  const gaps = countWordGaps(runs);
  if (gaps <= 0) return 0;
  const per = slack / gaps;
  let shift = 0;
  for (const r of runs) {
    r.x += shift;
    if (r.kind === 'text' && r.text) {
      const spaces = (r.text.match(/ /g) || []).length;
      if (spaces > 0) {
        // Widen this run to absorb its share of the gaps; following runs shift.
        r.width += spaces * per;
        shift += spaces * per;
      }
    }
  }
  return per;
}

function countWordGaps(runs: PositionedRun[]): number {
  let gaps = 0;
  for (const r of runs) {
    if (r.kind === 'text' && r.text) gaps += (r.text.match(/ /g) || []).length;
  }
  return gaps;
}
