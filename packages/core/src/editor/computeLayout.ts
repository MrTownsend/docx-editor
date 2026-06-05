/**
 * The pure layout COMPUTE pass shared by the React and Vue adapters — issue
 * #696 Tier 2, the clean half of the engine spine.
 *
 * This is the 6-step pass from React's `useLayoutPipeline` minus the DOM paint
 * + scroll/event side-effects (which stay adapter-side, where the framework
 * timing lives): PM doc → flow blocks → measure → header/footer resolve →
 * margin extension → `layoutDocument` (+ two-pass footnote stabilization) →
 * footnote render items. It is pure (no DOM, no refs, no rAF) and returns
 * everything the adapter needs to paint.
 *
 * The one injected seam is `measureBlocks` — each adapter passes its own
 * measurer (React's is caching), same pattern as `measureBlocksWithFloats`.
 * `getHfPmDoc` is the HF-unification seam (prefer the persistent PM doc over
 * re-parsing `HeaderFooter.content`).
 */

import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';

import {
  layoutDocument,
  type ColumnLayout,
  type FlowBlock,
  type FootnoteContent,
  type Layout,
  type Measure,
  type PageMargins,
  type SectionBreakBlock,
} from '../layout-engine';
import {
  toFlowBlocks,
  computePerBlockWidths,
  collectFootnoteRefs,
  convertHeaderFooterToContent,
  convertHeaderFooterPmDocToContent,
  buildFootnoteContentMap,
  buildFootnoteRenderItems,
  stabilizeFootnoteLayout,
  twipsToPixels,
  type FloatPageGeometry,
} from '../layout-bridge';
import {
  pageGeometryFromPage,
  type FootnoteRenderItem,
  type HeaderFooterContent,
} from '../layout-painter';
import type {
  Document,
  HeaderFooter,
  SectionProperties,
  StyleDefinitions,
  Theme,
  Watermark,
} from '../types/document';

interface PageSizePx {
  w: number;
  h: number;
}

/** Adapter-supplied block measurer (React's is caching). */
export type MeasureBlocksFn = (
  blocks: FlowBlock[],
  contentWidth: number | number[],
  pageGeometry?: FloatPageGeometry
) => Measure[];

export interface ComputeLayoutInputs {
  state: EditorState;
  document: Document | null;
  pageSize: PageSizePx;
  margins: PageMargins;
  columns: ColumnLayout | undefined;
  finalPageSize: PageSizePx;
  finalMargins: PageMargins;
  finalColumns: ColumnLayout | undefined;
  pageGap: number;
  contentWidth: number;
  theme: Theme | null | undefined;
  styles: StyleDefinitions | null | undefined;
  sectionProperties: SectionProperties | null | undefined;
  finalSectionProperties: SectionProperties | null | undefined;
  /** Resolved HF objects for the section (default + first-page). */
  headerContent: HeaderFooter | null | undefined;
  footerContent: HeaderFooter | null | undefined;
  firstPageHeaderContent: HeaderFooter | null | undefined;
  firstPageFooterContent: HeaderFooter | null | undefined;
  measureBlocks: MeasureBlocksFn;
  /** HF unification: the persistent PM doc for an HF, or null to re-parse content. */
  getHfPmDoc: (hf: HeaderFooter) => PMNode | null | undefined;
}

export interface LayoutComputation {
  blocks: FlowBlock[];
  measures: Measure[];
  layout: Layout;
  headerContentForRender: HeaderFooterContent | undefined;
  footerContentForRender: HeaderFooterContent | undefined;
  firstPageHeaderForRender: HeaderFooterContent | undefined;
  firstPageFooterForRender: HeaderFooterContent | undefined;
  hasTitlePg: boolean;
  watermark: Watermark | undefined;
  headerDistancePx: number | undefined;
  footerDistancePx: number | undefined;
  pageBorders: SectionProperties['pageBorders'] | undefined;
  footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined;
}

/**
 * Run the pure layout compute pass (the 6 steps in this file's header), lifted
 * verbatim from `useLayoutPipeline`. The adapter performs the DOM paint
 * (`renderPages`), scroll-restore, `painter:painted`, and state writeback with
 * the returned values.
 */
export function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation {
  const {
    state,
    document,
    pageSize,
    margins,
    columns,
    finalPageSize,
    finalMargins,
    finalColumns,
    pageGap,
    contentWidth,
    theme,
    styles,
    sectionProperties,
    finalSectionProperties,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    measureBlocks,
    getHfPmDoc,
  } = inputs;

  // Step 1: PM doc → flow blocks.
  const pageContentHeight = pageSize.h - margins.top - margins.bottom;
  const blocks = toFlowBlocks(state.doc, { theme, pageContentHeight });

  // Step 2: Measure all blocks (per-section widths; full measure for float context).
  const blockWidths = computePerBlockWidths(
    blocks,
    { pageSize, margins, columns },
    { pageSize: finalPageSize, margins: finalMargins, columns: finalColumns }
  );
  const measures = measureBlocks(
    blocks,
    blockWidths,
    pageGeometryFromPage({ size: pageSize, margins })
  );

  // Step 2.5: Footnote references.
  const footnoteRefs = collectFootnoteRefs(blocks);
  const hasFootnotes = footnoteRefs.length > 0 && !!document?.package?.footnotes;

  // Step 2.75: Header/footer content for rendering (needed before layout to
  // compute effective margins when HF content exceeds available space).
  const hfMetricsHeader = { section: 'header' as const, pageSize, margins };
  const hfMetricsFooter = { section: 'footer' as const, pageSize, margins };
  const defaultTabStopTwips = state.doc.attrs?.defaultTabStopTwips as number | null;
  const hfOptions = { styles, theme, measureBlocks, defaultTabStopTwips };

  // HF unification phase 1: prefer the persistent PM doc when mounted.
  const convertHf = (
    hf: HeaderFooter | null | undefined,
    metrics: typeof hfMetricsHeader | typeof hfMetricsFooter
  ): HeaderFooterContent | undefined => {
    if (!hf) return undefined;
    const pmDoc = getHfPmDoc(hf);
    if (pmDoc) {
      return convertHeaderFooterPmDocToContent(pmDoc, contentWidth, metrics, hfOptions);
    }
    return convertHeaderFooterToContent(hf, contentWidth, metrics, hfOptions);
  };

  const headerContentForRender = convertHf(headerContent, hfMetricsHeader);
  const footerContentForRender = convertHf(footerContent, hfMetricsFooter);
  const hasTitlePg = sectionProperties?.titlePg === true;
  const firstPageHeaderForRender = hasTitlePg
    ? convertHf(firstPageHeaderContent, hfMetricsHeader)
    : undefined;
  const firstPageFooterForRender = hasTitlePg
    ? convertHf(firstPageFooterContent, hfMetricsFooter)
    : undefined;

  // Watermark rides PM state as a doc attr (so it's undoable).
  const watermark = (state.doc.attrs?.watermark as Watermark | null) ?? undefined;

  // Margin extension: push body clear of header/footer when they overflow.
  const headerDistance = margins.header ?? 48;
  const footerDistance = margins.footer ?? 48;
  const availableHeaderSpace = margins.top - headerDistance;
  const availableFooterSpace = margins.bottom - footerDistance;
  const hfHeight = (hf: HeaderFooterContent | undefined) =>
    hf ? (hf.visualBottom ?? hf.height) : 0;
  const hfFooterHeight = (hf: HeaderFooterContent | undefined) =>
    hf ? Math.max((hf.visualBottom ?? hf.height) - (hf.visualTop ?? 0), hf.height) : 0;
  const headerContentHeight = Math.max(
    hfHeight(headerContentForRender),
    hfHeight(firstPageHeaderForRender)
  );
  const footerContentHeight = Math.max(
    hfFooterHeight(footerContentForRender),
    hfFooterHeight(firstPageFooterForRender)
  );

  const extendHeader = headerContentHeight > availableHeaderSpace;
  const extendFooter = footerContentHeight > availableFooterSpace;
  let effectiveMargins = margins;
  let effectiveFinalMargins = finalMargins;
  if (extendHeader || extendFooter) {
    const extend = (m: PageMargins): PageMargins => {
      const out = { ...m };
      if (extendHeader) out.top = Math.max(m.top, headerDistance + headerContentHeight);
      if (extendFooter) out.bottom = Math.max(m.bottom, footerDistance + footerContentHeight);
      return out;
    };
    effectiveMargins = extend(margins);
    effectiveFinalMargins = extend(finalMargins);
    for (const block of blocks) {
      if (block.kind !== 'sectionBreak') continue;
      const sb = block as SectionBreakBlock;
      if (sb.margins) sb.margins = extend(sb.margins);
    }
  }

  // Step 3: Layout onto pages (two-pass when footnotes exist).
  const bodyBreakType = finalSectionProperties?.sectionStart as
    | 'continuous'
    | 'nextPage'
    | 'evenPage'
    | 'oddPage'
    | undefined;
  const layoutOpts = {
    pageSize,
    margins: effectiveMargins,
    finalPageSize,
    finalMargins: effectiveFinalMargins,
    columns: finalColumns,
    bodyBreakType,
    pageGap,
  };

  let layout: Layout;
  let pageFootnoteMap = new Map<number, number[]>();
  let footnoteContentMap = new Map<number, FootnoteContent>();

  if (hasFootnotes) {
    const pass1Layout = layoutDocument(blocks, measures, layoutOpts);
    footnoteContentMap = buildFootnoteContentMap(
      document!.package.footnotes!,
      footnoteRefs,
      contentWidth,
      {
        styles: styles ?? undefined,
        theme: theme ?? null,
        measureBlocks,
        defaultTabStopTwips,
      }
    );
    const stabilized = stabilizeFootnoteLayout({
      blocks,
      measures,
      layoutOpts,
      footnoteRefs,
      footnoteContentMap,
      initialLayout: pass1Layout,
    });
    layout = stabilized.layout;
    pageFootnoteMap = stabilized.pageFootnoteMap;
  } else {
    layout = layoutDocument(blocks, measures, layoutOpts);
  }

  const footnotesByPage = hasFootnotes
    ? buildFootnoteRenderItems(pageFootnoteMap, footnoteContentMap, document)
    : undefined;

  return {
    blocks,
    measures,
    layout,
    headerContentForRender,
    footerContentForRender,
    firstPageHeaderForRender,
    firstPageFooterForRender,
    hasTitlePg,
    watermark,
    headerDistancePx: sectionProperties?.headerDistance
      ? twipsToPixels(sectionProperties.headerDistance)
      : undefined,
    footerDistancePx: sectionProperties?.footerDistance
      ? twipsToPixels(sectionProperties.footerDistance)
      : undefined,
    pageBorders: sectionProperties?.pageBorders,
    footnotesByPage: footnotesByPage?.size ? footnotesByPage : undefined,
  };
}
