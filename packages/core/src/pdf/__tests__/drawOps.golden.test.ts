/**
 * Golden draw-op snapshots — the parity DRIFT TRIPWIRE.
 *
 * Renders fixture paragraphs/tables to a RecordingSink (no pdf-lib page, no
 * network) and snapshots the exact draw ops (text positions, rects, lines). Any
 * change to the exporter's geometry or to a painter rule it mirrors shows up as a
 * readable snapshot diff in CI — so screen↔PDF drift can't ship silently. Update
 * the snapshot ONLY when a change is intentional and reviewed.
 */
import { describe, test, expect } from 'bun:test';
import { PDFDocument, StandardFonts, type PDFFont } from 'pdf-lib';
import { RecordingSink } from '../pageSink';
import { drawParagraphAt } from '../pdfParagraph';
import { drawTableAt } from '../pdfTable';
import { createImageEmbedder } from '../pdfImage';
import type { FontProvider } from '../fontProvider';
import type { FieldContext } from '../../layout-painter/renderParagraph/positionRuns';
import type {
  ParagraphBlock,
  ParagraphMeasure,
  MeasuredLine,
  Run,
  TableBlock,
  TableMeasure,
} from '../../layout-engine/types';

// Deterministic fonts: one embedded Helvetica (its glyph metrics are fixed, so
// positions are stable across machines — no network, no system-font variance).
async function fixtures() {
  const doc = await PDFDocument.create();
  const helv = doc.embedStandardFont(StandardFonts.Helvetica);
  const fonts: FontProvider = {
    async warmUp() {},
    getFontSync: () => helv as PDFFont,
    getUnicodeFallbackSync: () => helv as PDFFont,
  };
  return { doc, fonts };
}

const field: FieldContext = { pageNumber: 1, totalPages: 1 };
const line = (runs: Run[], over: Partial<MeasuredLine> = {}): MeasuredLine => ({
  fromRun: 0,
  fromChar: 0,
  toRun: runs.length - 1,
  toChar: 999,
  width: 200,
  ascent: 12,
  descent: 4,
  lineHeight: 18,
  ...over,
});

describe('PDF draw-op golden snapshots (drift tripwire)', () => {
  test('paragraph: formatted + highlighted + justified runs', async () => {
    const { fonts } = await fixtures();
    const runs: Run[] = [
      { kind: 'text', text: 'Bold ', bold: true, fontSize: 12 } as Run,
      { kind: 'text', text: 'highlighted words ', highlight: 'yellow', fontSize: 12 } as Run,
      { kind: 'text', text: 'and underlined', underline: true, fontSize: 12 } as Run,
    ];
    const block: ParagraphBlock = {
      kind: 'paragraph',
      id: 1 as never,
      runs,
      attrs: { alignment: 'justify' },
    };
    const measure: ParagraphMeasure = {
      kind: 'paragraph',
      lines: [line(runs, { width: 300 })],
      totalHeight: 18,
    };
    const sink = new RecordingSink();
    drawParagraphAt({
      page: sink,
      block,
      measure,
      x: 96,
      y: 96,
      width: 400,
      pageHpx: 1056,
      fonts,
      field,
    });
    expect(sink.ops).toMatchSnapshot();
  });

  test('paragraph: inline image + hyperlink + letter spacing + small caps', async () => {
    const { fonts, doc } = await fixtures();
    const embedder = createImageEmbedder(doc);
    // A 1x1 PNG so the inline image embeds and draws (not a placeholder).
    const png =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
    await embedder.warmUp([png]);
    const runs: Run[] = [
      { kind: 'text', text: 'LINK', fontSize: 12, hyperlink: { href: 'https://x' } } as Run,
      { kind: 'text', text: ' spaced', fontSize: 12, letterSpacing: 2 } as Run,
      { kind: 'image', src: png, width: 10, height: 10 } as Run,
      { kind: 'text', text: 'caps', fontSize: 12, smallCaps: true } as Run,
    ];
    const block: ParagraphBlock = { kind: 'paragraph', id: 2 as never, runs, attrs: {} };
    const measure: ParagraphMeasure = {
      kind: 'paragraph',
      lines: [line(runs, { width: 200 })],
      totalHeight: 18,
    };
    const sink = new RecordingSink();
    drawParagraphAt({
      page: sink,
      block,
      measure,
      x: 96,
      y: 96,
      width: 400,
      pageHpx: 1056,
      fonts,
      field,
      embedder,
    });
    // Locks: hyperlink underline (a line op), letter-spaced per-char draws,
    // inline image op, small-caps uppercased text.
    expect(sink.ops).toMatchSnapshot();
    expect(sink.ops.some((o) => o.op === 'image')).toBe(true);
    expect(sink.ops.some((o) => o.op === 'text' && o.text === 'CAPS')).toBe(true);
  });

  test('table: 2x2 with shading, borders, and a vertical merge', async () => {
    const { fonts, doc } = await fixtures();
    const embedder = createImageEmbedder(doc);
    const cellPara = (text: string) => {
      const runs: Run[] = [{ kind: 'text', text, fontSize: 11 } as Run];
      return {
        block: { kind: 'paragraph', id: 0 as never, runs, attrs: {} } as ParagraphBlock,
        measure: { kind: 'paragraph', lines: [line(runs)], totalHeight: 18 } as ParagraphMeasure,
      };
    };
    const border = { style: 'single', width: 1, color: '#000000' };
    const borders = { top: border, bottom: border, left: border, right: border };
    const mkCell = (text: string, opts: Partial<Run> & { bg?: string; rowSpan?: number } = {}) => {
      const cp = cellPara(text);
      return {
        cell: {
          id: 0 as never,
          blocks: [cp.block],
          borders,
          background: opts.bg,
          rowSpan: opts.rowSpan,
        },
        measure: { blocks: [cp.measure], width: 150, height: 30 },
      };
    };
    // Row 0: [A1 (rowSpan 2)] [B1];  Row 1: [B2]  (A2 covered by A1's merge)
    const r0 = [mkCell('A1', { bg: '#eeeeee', rowSpan: 2 }), mkCell('B1')];
    const r1 = [mkCell('B2')];
    const table = {
      kind: 'table' as const,
      id: 5 as never,
      columnWidths: [150, 150],
      rows: [
        { id: 0 as never, cells: r0.map((c) => c.cell) },
        { id: 1 as never, cells: r1.map((c) => c.cell) },
      ],
    } as unknown as TableBlock;
    const measure = {
      kind: 'table' as const,
      columnWidths: [150, 150],
      totalWidth: 300,
      totalHeight: 60,
      rows: [
        { cells: r0.map((c) => c.measure), height: 30 },
        { cells: r1.map((c) => c.measure), height: 30 },
      ],
    } as unknown as TableMeasure;

    const sink = new RecordingSink();
    drawTableAt({
      page: sink,
      block: table,
      measure,
      x: 96,
      y: 96,
      fromRow: 0,
      toRow: 2,
      pageHpx: 1056,
      fonts,
      field,
      embedder,
    });
    expect(sink.ops).toMatchSnapshot();
  });
});
