import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';

import { getSelectionRectsFromDom } from '../clickToPositionDom';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function buildBlankRow(pmPos: number): HTMLElement {
  document.body.innerHTML = `
    <div class="layout-page" data-page-number="1">
      <div class="layout-page-content">
        <div class="layout-paragraph">
          <div class="layout-line">
            <span class="layout-run" data-pm-start="${pmPos}" data-pm-end="${pmPos}">\u200B</span>
          </div>
        </div>
      </div>
    </div>`;
  return document.body.firstElementChild as HTMLElement;
}

describe('getSelectionRectsFromDom — blank-line marker', () => {
  test('emits a fixed-width sliver for a zero-width marker inside the selection', () => {
    const container = buildBlankRow(7);

    const rects = getSelectionRectsFromDom(container, 5, 10, new DOMRect(0, 0, 0, 0));

    expect(rects).toHaveLength(1);
    expect(rects[0].width).toBe(4);
  });

  test('emits nothing for a blank-line marker outside the selection', () => {
    const container = buildBlankRow(7);

    const rects = getSelectionRectsFromDom(container, 0, 3, new DOMRect(0, 0, 0, 0));

    expect(rects).toHaveLength(0);
  });
});
