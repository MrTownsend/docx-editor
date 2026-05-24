/**
 * Render inline content (runs, hyperlinks, comment-range markers, tracked
 * changes) to markdown. Operates on the `ParagraphContent[]` of a paragraph,
 * or the equivalent inline lists inside hyperlinks and tracked-change wrappers.
 *
 * Inline marks are rendered as character-precise wrappers around the affected
 * runs. Comments and tracked changes are turned into configurable annotation
 * tags via `./annotations.ts`. Footnote and image references are emitted as
 * virtual identifiers; the caller collects the actual content from the
 * surrounding render context.
 */

import type { DocxPackage } from '../types/document';
import type {
  Run,
  Hyperlink,
  CommentRangeStart,
  CommentRangeEnd,
  Comment,
  Insertion,
  Deletion,
  MoveFrom,
  MoveTo,
  RunContent,
  ParagraphContent,
} from '../types/document';
import { registerImage } from './images';
import { escapeInline, escapeLinkUrl, escapeAltText } from './escape';
import { wrapInsertion, wrapDeletion, wrapMoveFrom, wrapMoveTo, wrapComment } from './annotations';
import type { RenderContext } from './types';

/**
 * Inline marks we recognize. Order matters: we open from outermost to
 * innermost so the output reads cleanly (`***bold italic***`), not the reverse.
 */
type MarkKey = 'bold' | 'italic' | 'code' | 'strike';

const MARK_DELIMS: Record<MarkKey, string> = {
  bold: '**',
  italic: '*',
  code: '`',
  strike: '~~',
};

// Word has no `code` run property. We infer it from a small whitelist of
// monospace font families so prose set in fonts like `Monotype Corsiva` is
// not wrapped in backticks.
const MONOSPACE_FONTS = new Set([
  'consolas',
  'courier',
  'courier new',
  'menlo',
  'monaco',
  'sf mono',
  'jetbrains mono',
  'fira code',
  'fira mono',
  'source code pro',
  'roboto mono',
  'inconsolata',
  'lucida console',
  'monospace',
]);

function marksFor(run: Run): MarkKey[] {
  const f = run.formatting;
  if (!f) return [];
  const out: MarkKey[] = [];
  if (f.bold) out.push('bold');
  if (f.italic) out.push('italic');
  if (f.strike) out.push('strike');
  const ascii = f.fontFamily?.ascii?.toLowerCase();
  if (ascii && MONOSPACE_FONTS.has(ascii)) out.push('code');
  return out;
}

function applyMarks(text: string, marks: MarkKey[]): string {
  if (!text) return text;
  // Code overrides other marks: code is literal.
  if (marks.includes('code')) {
    const delim = '`';
    // If text contains backticks, use a longer fence.
    if (text.includes('`')) return `\`\`${text}\`\``;
    return `${delim}${text}${delim}`;
  }
  let out = text;
  for (const m of marks) {
    const d = MARK_DELIMS[m];
    out = `${d}${out}${d}`;
  }
  return out;
}

/** Render a single run's RunContent array into the inline text fragment. */
function renderRunContent(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  content: RunContent[],
  paraId: string | undefined
): string {
  let out = '';
  for (const item of content) {
    switch (item.type) {
      case 'text':
        out += escapeInline(item.text);
        break;
      case 'tab':
        out += '    ';
        break;
      case 'break':
        if (item.breakType === 'page') {
          // Page break inside text; in unpaged we emit a thematic break-ish marker.
          out += '\n\n';
        } else {
          // Soft break: emit a two-space line break (markdown's hard wrap).
          out += '  \n';
        }
        break;
      case 'symbol':
        out += escapeInline(item.char);
        break;
      case 'softHyphen':
        // U+00AD soft hyphen. Word displays it only when needed for line
        // breaks; drop it from the markdown output.
        break;
      case 'noBreakHyphen':
        out += '‑';
        break;
      case 'footnoteRef':
      case 'endnoteRef': {
        const markerNumber = ctx.footnoteRefs.length + 1;
        ctx.footnoteRefs.push({ refId: item.id, markerNumber });
        out += `[^${markerNumber}]`;
        break;
      }
      case 'drawing': {
        // Preferred path: resolve via the package's `rels -> media` chain.
        // That returns raw bytes, so we can register a stable virtual path
        // and expose the image in `result.images`.
        const ref = pkg?.relationships?.get(item.image.rId);
        const media = ref ? pkg?.media?.get(ref.target) : undefined;
        if (media) {
          const reg = registerImage(ctx, media, item.image, paraId);
          const alt = reg.alt ? escapeAltText(reg.alt) : '';
          out += `![${alt}](${reg.virtualPath})`;
          break;
        }
        // Fallback: header/footer images use a separate rels file
        // (word/_rels/header1.xml.rels) that does not live in
        // `pkg.relationships`. The parser inlines the bytes into
        // `image.src` (typically a data URL) — emit that directly.
        if (item.image.src) {
          const alt = item.image.alt ?? item.image.title ?? item.image.filename ?? '';
          out += `![${escapeAltText(alt)}](${item.image.src})`;
          break;
        }
        ctx.warnings.push(`image rId=${item.image.rId} not resolvable`);
        break;
      }
      case 'shape':
        ctx.warnings.push('shape not representable in markdown');
        break;
      case 'fieldChar':
      case 'instrText':
        // Field chrome. Skip: the field result text lives in surrounding runs.
        break;
      default:
        break;
    }
  }
  return out;
}

/** Render a single Run with its formatting applied. */
function renderRun(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  run: Run,
  paraId: string | undefined
): string {
  const inner = renderRunContent(ctx, pkg, run.content, paraId);
  if (!inner) return '';
  // Markdown can't carry whitespace at the boundaries of a mark, so we split
  // leading/trailing whitespace out of the wrapped text.
  const match = inner.match(/^(\s*)(.*?)(\s*)$/s);
  if (!match) return applyMarks(inner, marksFor(run));
  const [, lead, core, trail] = match;
  if (!core) return inner;
  return `${lead}${applyMarks(core, marksFor(run))}${trail}`;
}

function renderHyperlink(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  link: Hyperlink,
  paraId: string | undefined
): string {
  const inner = link.children
    .map((child) => (child.type === 'run' ? renderRun(ctx, pkg, child, paraId) : ''))
    .join('');
  if (!inner) return '';
  const href = link.href ?? (link.anchor ? `#${link.anchor}` : '');
  if (!href) {
    ctx.warnings.push('hyperlink missing href and anchor; rendered as plain text');
    return inner;
  }
  if (ctx.opts.hyperlinks === 'reference') {
    const refNumber = ctx.hyperlinkRefs.length + 1;
    ctx.hyperlinkRefs.push({ href, refNumber });
    return `[${inner}][${refNumber}]`;
  }
  return `[${inner}](${escapeLinkUrl(href)})`;
}

function renderTrackedWrapper(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  wrapper: Insertion | Deletion | MoveFrom | MoveTo,
  paraId: string | undefined
): string {
  if (ctx.opts.trackedChanges === 'clean') {
    // Insertions become real text; deletions vanish.
    if (wrapper.type === 'insertion' || wrapper.type === 'moveTo') {
      return wrapper.content
        .map((child) =>
          child.type === 'run'
            ? renderRun(ctx, pkg, child, paraId)
            : renderHyperlink(ctx, pkg, child, paraId)
        )
        .join('');
    }
    return '';
  }
  const inner = wrapper.content
    .map((child) =>
      child.type === 'run'
        ? renderRun(ctx, pkg, child, paraId)
        : renderHyperlink(ctx, pkg, child, paraId)
    )
    .join('');
  switch (wrapper.type) {
    case 'insertion':
      return wrapInsertion(ctx, wrapper.info, inner);
    case 'deletion':
      return wrapDeletion(ctx, wrapper.info, inner);
    case 'moveFrom':
      return wrapMoveFrom(ctx, wrapper.info, inner);
    case 'moveTo':
      return wrapMoveTo(ctx, wrapper.info, inner);
  }
}

interface CommentSlot {
  start: number;
  comment?: Comment;
}

/** Look up a comment by id in the document body. */
function commentById(pkg: DocxPackage | undefined, id: number): Comment | undefined {
  return pkg?.document.comments?.find((c) => c.id === id);
}

/**
 * Render the full inline content of a paragraph, tracking comment-range
 * boundaries to apply the configured wrapper.
 */
export function renderParagraphInline(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  content: ParagraphContent[],
  paraId: string | undefined
): string {
  let out = '';
  // Stack of open comment ranges (in document order) so nested comments wrap correctly.
  const openComments: CommentSlot[] = [];

  for (const item of content) {
    switch (item.type) {
      case 'run':
        out += renderRun(ctx, pkg, item, paraId);
        break;
      case 'hyperlink':
        out += renderHyperlink(ctx, pkg, item, paraId);
        break;
      case 'insertion':
      case 'deletion':
      case 'moveFrom':
      case 'moveTo':
        out += renderTrackedWrapper(ctx, pkg, item, paraId);
        break;
      case 'commentRangeStart':
        out += handleCommentStart(ctx, pkg, item, openComments, out.length);
        break;
      case 'commentRangeEnd':
        out = handleCommentEnd(ctx, pkg, item, openComments, out);
        break;
      case 'bookmarkStart':
      case 'bookmarkEnd':
      case 'moveFromRangeStart':
      case 'moveFromRangeEnd':
      case 'moveToRangeStart':
      case 'moveToRangeEnd':
        // Markers without inline payload.
        break;
      case 'simpleField':
      case 'complexField': {
        // Render the visible result content.
        const runs = item.type === 'simpleField' ? item.content : item.fieldResult;
        for (const child of runs) {
          out +=
            child.type === 'run'
              ? renderRun(ctx, pkg, child, paraId)
              : renderHyperlink(ctx, pkg, child, paraId);
        }
        break;
      }
      case 'inlineSdt': {
        const sdtContent = (item as { content?: ParagraphContent[] }).content;
        if (Array.isArray(sdtContent)) {
          out += renderParagraphInline(ctx, pkg, sdtContent, paraId);
        }
        break;
      }
      default:
        break;
    }
  }

  // Close any still-open comment ranges defensively.
  while (openComments.length) {
    const slot = openComments.pop();
    if (!slot) break;
    if (!slot.comment || ctx.opts.comments === 'strip') continue;
    out = applyCommentWrapping(ctx, slot, out);
  }

  return out;
}

function handleCommentStart(
  ctx: RenderContext,
  pkg: DocxPackage | undefined,
  marker: CommentRangeStart,
  openComments: CommentSlot[],
  startPos: number
): string {
  if (ctx.opts.comments === 'strip') {
    openComments.push({ start: startPos, comment: undefined });
    return '';
  }
  const comment = commentById(pkg, marker.id);
  openComments.push({ start: startPos, comment });
  return '';
}

function handleCommentEnd(
  ctx: RenderContext,
  _pkg: DocxPackage | undefined,
  _marker: CommentRangeEnd,
  openComments: CommentSlot[],
  current: string
): string {
  const slot = openComments.pop();
  if (!slot || !slot.comment || ctx.opts.comments === 'strip') return current;
  return applyCommentWrapping(ctx, slot, current);
}

function applyCommentWrapping(ctx: RenderContext, slot: CommentSlot, current: string): string {
  if (!slot.comment) return current;
  const before = current.slice(0, slot.start);
  const inner = current.slice(slot.start);
  if (ctx.opts.comments === 'sidecar') {
    const markerNumber = ctx.commentRefs.length + 1;
    ctx.commentRefs.push({ commentId: slot.comment.id, markerNumber });
    return `${before}${inner}[^c${markerNumber}]`;
  }
  return before + wrapComment(ctx, { id: slot.comment.id, author: slot.comment.author }, inner);
}
