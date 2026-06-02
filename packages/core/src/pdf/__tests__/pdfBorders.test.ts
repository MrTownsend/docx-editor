import { describe, test, expect } from 'bun:test';
import { strokeForBorder, normalizePageBorderSide, drawBorderLine } from '../pdfBorders';
import { RecordingSink } from '../pageSink';

describe('pdf/pdfBorders', () => {
  test('strokeForBorder maps styles', () => {
    expect(strokeForBorder(undefined)).toBeNull();
    expect(strokeForBorder({ style: 'none' })).toBeNull();
    expect(strokeForBorder({ style: 'single', width: 1 })?.double).toBe(false);
    expect(strokeForBorder({ style: 'double', width: 1 })?.double).toBe(true);
    expect(strokeForBorder({ style: 'dotted', width: 1 })?.dashArray).toBeDefined();
  });

  test('normalizePageBorderSide handles eighths-pt size and {rgb} color', () => {
    const n = normalizePageBorderSide({ style: 'single', size: 8, color: { rgb: 'FF0000' } });
    expect(n?.color).toBe('#FF0000');
    expect(n?.width).toBeCloseTo(8 / 8 / 72 / (1 / 96), 2); // 1pt → px
  });

  test('drawBorderLine draws one line for single, two parallel for double', () => {
    const single = new RecordingSink();
    drawBorderLine(single, { style: 'single', width: 1 }, { x: 0, y: 10 }, { x: 100, y: 10 });
    expect(single.ops.filter((o) => o.op === 'line')).toHaveLength(1);

    const dbl = new RecordingSink();
    drawBorderLine(dbl, { style: 'double', width: 1 }, { x: 0, y: 10 }, { x: 100, y: 10 });
    const lines = dbl.ops.filter((o) => o.op === 'line');
    expect(lines).toHaveLength(2);
    // Two horizontal lines offset above and below y=10.
    const ys = lines.map((l) => (l.op === 'line' ? l.y1 : 0)).sort((a, b) => a - b);
    expect(ys[0]).toBeLessThan(10);
    expect(ys[1]).toBeGreaterThan(10);
    // Each sub-line is ~1/3 the single thickness (total span ≈ declared width),
    // so a double border isn't ~3× heavier than Word's.
    const singleThickness = strokeForBorder({ style: 'single', width: 1 })!.thickness;
    const dblThickness = lines[0].op === 'line' ? lines[0].thickness : 0;
    expect(dblThickness).toBeCloseTo(singleThickness / 3, 4);
  });

  test('drawBorderLine draws nothing for a none border', () => {
    const sink = new RecordingSink();
    drawBorderLine(sink, { style: 'none' }, { x: 0, y: 0 }, { x: 10, y: 0 });
    expect(sink.ops).toHaveLength(0);
  });
});
