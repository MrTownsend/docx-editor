/**
 * Encoders for constructs markdown can't express natively: tracked changes,
 * comments, headers, footers. Three modes: `html` (default), `pandoc`, `strip`.
 *
 * The encoder takes the inner text already rendered and returns the wrapped
 * form. Order of operations: render children first, then wrap.
 */

import type { TrackedChangeInfo } from '../types/document';
import type { RenderContext } from './types';

function escapeAttr(value: string): string {
  return value.replace(/"/g, '&quot;');
}

function attrs(info: TrackedChangeInfo): string {
  const parts: string[] = [];
  if (info.author) parts.push(`author="${escapeAttr(info.author)}"`);
  if (info.date) parts.push(`date="${escapeAttr(info.date)}"`);
  if (typeof info.id === 'number') parts.push(`id="${info.id}"`);
  return parts.length ? ' ' + parts.join(' ') : '';
}

/** Wrap an insertion. `<ins>` is real HTML; pandoc uses bracketed span. */
export function wrapInsertion(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  if (ctx.opts.annotations === 'strip') return inner;
  if (ctx.opts.annotations === 'pandoc') {
    const author = info.author ? ` author="${escapeAttr(info.author)}"` : '';
    return `[${inner}]{.ins${author}}`;
  }
  return `<ins${attrs(info)}>${inner}</ins>`;
}

/** Wrap a deletion. */
export function wrapDeletion(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  if (ctx.opts.annotations === 'strip') return '';
  if (ctx.opts.annotations === 'pandoc') {
    const author = info.author ? ` author="${escapeAttr(info.author)}"` : '';
    return `[${inner}]{.del${author}}`;
  }
  return `<del${attrs(info)}>${inner}</del>`;
}

/** Wrap a move-from (deletion side of a move). */
export function wrapMoveFrom(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  if (ctx.opts.annotations === 'strip') return '';
  if (ctx.opts.annotations === 'pandoc') {
    return `[${inner}]{.move-from author="${escapeAttr(info.author ?? '')}"}`;
  }
  return `<del data-move="from"${attrs(info)}>${inner}</del>`;
}

/** Wrap a move-to (insertion side of a move). */
export function wrapMoveTo(ctx: RenderContext, info: TrackedChangeInfo, inner: string): string {
  if (ctx.opts.annotations === 'strip') return inner;
  if (ctx.opts.annotations === 'pandoc') {
    return `[${inner}]{.move-to author="${escapeAttr(info.author ?? '')}"}`;
  }
  return `<ins data-move="to"${attrs(info)}>${inner}</ins>`;
}

/** Wrap commented text in inline mode. */
export function wrapComment(
  ctx: RenderContext,
  meta: { id: number; author?: string },
  inner: string
): string {
  if (ctx.opts.annotations === 'strip') return inner;
  if (ctx.opts.annotations === 'pandoc') {
    const author = meta.author ? ` author="${escapeAttr(meta.author)}"` : '';
    return `[${inner}]{.comment id="${meta.id}"${author}}`;
  }
  const author = meta.author ? ` author="${escapeAttr(meta.author)}"` : '';
  return `<comment id="${meta.id}"${author}>${inner}</comment>`;
}

/**
 * Block wrapper for headers/footers.
 *
 * Uses blank lines between the wrapper tag and the inner content so each
 * side parses as its own HTML block (CommonMark type-7). Markdown inside
 * (bold, links, lists, images, tables) is then parsed normally by GFM
 * renderers; the surrounding `<header>` / `<footer>` tags pass through as
 * raw HTML.
 */
export function wrapHeaderFooter(
  ctx: RenderContext,
  kind: 'header' | 'footer',
  inner: string
): string {
  if (ctx.opts.annotations === 'strip') return '';
  if (ctx.opts.annotations === 'pandoc') {
    return `:::${kind}\n\n${inner}\n\n:::`;
  }
  return `<${kind}>\n\n${inner}\n\n</${kind}>`;
}
