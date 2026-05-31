/**
 * Document to ProseMirror Conversion
 *
 * Converts our Document type (from DOCX parsing) to a ProseMirror document.
 * Preserves all formatting attributes for round-trip fidelity.
 *
 * Style Resolution:
 * When styles are provided, paragraph properties are resolved from the style chain:
 * - Document defaults (docDefaults)
 * - Normal style (if no explicit styleId)
 * - Style chain (basedOn inheritance)
 * - Inline properties (highest priority)
 *
 * This file owns the top-level entry points (toProseDoc, headerFooterToProseDoc,
 * footnoteToProseDoc, createEmptyDoc). Per-domain converters live under
 * ./toProseDoc/ (marks, runs, paragraph, tables, textbox) — symmetric to
 * the fromProseDoc/ split.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../schema';
import type {
  Document,
  Paragraph,
  Table,
  BlockContent,
  StyleDefinitions,
  Theme,
} from '../../types/document';
import { createStyleResolver } from '../styles';
import { paragraphHasPageBreak } from './toProseDoc/paragraph';

/**
 * Flatten block content for PM conversion, expanding block-level SDTs into
 * their child blocks. Block SDTs are preserved in the document model for
 * round-trip (see `BlockSdt`), but are not yet a distinct editable PM node
 * (Phase 1) — rendering their children inline keeps the content visible and
 * editable, matching the pre-existing behavior where the wrapper was
 * flattened at parse time. Nested block SDTs are expanded recursively.
 */
function flattenBlockContent(blocks: BlockContent[]): Array<Paragraph | Table> {
  const out: Array<Paragraph | Table> = [];
  for (const block of blocks) {
    if (block.type === 'blockSdt') {
      out.push(...flattenBlockContent(block.content));
    } else {
      out.push(block);
    }
  }
  return out;
}
import { convertTable } from './toProseDoc/tables';
import { convertParagraphWithTextBoxes } from './toProseDoc/textbox';

/**
 * Options for document conversion
 */
export interface ToProseDocOptions {
  /** Style definitions for resolving paragraph styles */
  styles?: StyleDefinitions;
  /**
   * Doc-level `w:defaultTabStop` (§17.6.13) in twips, stamped onto the PM
   * doc node so `toFlowBlocks` picks it up. The body entry point reads
   * this from the parsed package; HF/footnote callers must pass it
   * through explicitly since their input is a content array, not a full
   * `Document`. Falls back to the OOXML default (720 twips) when null.
   */
  defaultTabStopTwips?: number | null;
}

/**
 * Convert a Document to a ProseMirror document
 *
 * @param document - The Document to convert
 * @param options - Conversion options including style definitions
 */
export function toProseDoc(document: Document, options?: ToProseDocOptions): PMNode {
  const paragraphs = flattenBlockContent(document.package.document.content);
  const nodes: PMNode[] = [];
  const theme = document.package.theme ?? null;

  // Create style resolver if styles are provided
  const styleResolver = options?.styles ? createStyleResolver(options.styles) : null;

  for (const block of paragraphs) {
    if (block.type === 'paragraph') {
      // Convert paragraph and extract text boxes as sibling nodes
      nodes.push(...convertParagraphWithTextBoxes(block, styleResolver));
      // If any run in this paragraph contains a page break, emit a pageBreak node after
      if (paragraphHasPageBreak(block)) {
        nodes.push(schema.node('pageBreak'));
      }
    } else if (block.type === 'table') {
      const pmTable = convertTable(block, styleResolver, theme);
      nodes.push(pmTable);
    }
  }

  // Ensure we have at least one paragraph
  if (nodes.length === 0) {
    nodes.push(schema.node('paragraph', {}, []));
  }

  return schema.node(
    'doc',
    { defaultTabStopTwips: document.package.settings?.defaultTabStop ?? null },
    nodes
  );
}

/**
 * Convert HeaderFooter content (array of Paragraph/Table blocks) to a ProseMirror document.
 * Used for editing headers/footers in their own ProseMirror editor and for the
 * unified header/footer render pipeline. `theme` must be threaded for themeColor
 * resolution in cell shading (`<w:shd w:themeFill=...>`) — without it, themed
 * fills in HF tables fall back to the unresolved theme key.
 */
export function headerFooterToProseDoc(
  content: BlockContent[],
  options?: ToProseDocOptions & { theme?: Theme | null }
): PMNode {
  const nodes: PMNode[] = [];
  const styleResolver = options?.styles ? createStyleResolver(options.styles) : null;
  const theme = options?.theme ?? null;

  for (const block of flattenBlockContent(content)) {
    if (block.type === 'paragraph') {
      nodes.push(...convertParagraphWithTextBoxes(block, styleResolver));
    } else if (block.type === 'table') {
      nodes.push(convertTable(block, styleResolver, theme));
    }
  }

  if (nodes.length === 0) {
    nodes.push(schema.node('paragraph', {}, []));
  }

  return schema.node('doc', { defaultTabStopTwips: options?.defaultTabStopTwips ?? null }, nodes);
}

/**
 * Convert footnote/endnote content (array of Paragraph/Table blocks) to a
 * ProseMirror document. Mirrors `headerFooterToProseDoc` so footnotes flow
 * through the same body pipeline (toFlowBlocks → measureBlocks →
 * renderFragment) and inherit its block support — paragraph + table + image
 * + textBox + fields. Pre-PR, footnoteLayout's `convertFootnoteToContent`
 * re-implemented run/paragraph conversion by hand and silently dropped
 * tables, images, and fields nested inside a footnote.
 */
export function footnoteToProseDoc(
  content: BlockContent[],
  options?: ToProseDocOptions & { theme?: Theme | null }
): PMNode {
  return headerFooterToProseDoc(content, options);
}

/**
 * Create an empty ProseMirror document
 */
export function createEmptyDoc(): PMNode {
  return schema.node('doc', null, [schema.node('paragraph', {}, [])]);
}
