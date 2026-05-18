/**
 * Schema-free table-delete commands plus the two keymap commands that gate
 * the global Backspace/Delete bindings on tables (`deleteTableIfSelected`,
 * `preventTableMergeAtGap`). None of these touch the schema — they only
 * mutate transactions on existing nodes — so all five are top-level
 * `Command` constants. The keymap commands could have been zero-arg
 * factories matching the schema-binding `make*` pattern elsewhere, but
 * since they capture no state the wrap would be dead work.
 */

import { type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import { getTableContext } from '../context';

export const deleteRow: Command = (
  state: EditorState,
  dispatch?: (tr: Transaction) => void
): boolean => {
  const context = getTableContext(state);
  if (
    !context.isInTable ||
    context.rowIndex === undefined ||
    !context.table ||
    context.tablePos === undefined ||
    (context.rowCount || 0) <= 1
  )
    return false;

  if (dispatch) {
    const tr = state.tr;
    let rowStart = context.tablePos + 1;
    for (let i = 0; i < context.rowIndex; i++) {
      rowStart += context.table.child(i).nodeSize;
    }
    const rowEnd = rowStart + context.table.child(context.rowIndex).nodeSize;
    tr.delete(rowStart, rowEnd);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

export const deleteColumn: Command = (
  state: EditorState,
  dispatch?: (tr: Transaction) => void
): boolean => {
  const context = getTableContext(state);
  if (
    !context.isInTable ||
    context.columnIndex === undefined ||
    !context.table ||
    context.tablePos === undefined ||
    (context.columnCount || 0) <= 1
  )
    return false;

  if (dispatch) {
    let tr = state.tr;
    const newColumnCount = (context.columnCount || 2) - 1;
    // Width is stored as 50ths of a percent per ECMA-376 §17.18.111
    // (5000 = 100%) so resolveTableWidthPx can apply it directly.
    const newColWidthPercent = Math.floor(5000 / newColumnCount);

    const deleteOps: { start: number; end: number }[] = [];
    let rowPos = context.tablePos + 1;

    context.table.forEach((row) => {
      if (row.type.name === 'tableRow') {
        let cellPos = rowPos + 1;
        let colIdx = 0;

        row.forEach((cell) => {
          const cellStart = cellPos;
          const cellEnd = cellPos + cell.nodeSize;
          const cellColspan = cell.attrs.colspan || 1;

          if (colIdx <= context.columnIndex! && context.columnIndex! < colIdx + cellColspan) {
            deleteOps.push({ start: cellStart, end: cellEnd });
          }

          cellPos = cellEnd;
          colIdx += cellColspan;
        });
      }
      rowPos += row.nodeSize;
    });

    deleteOps.reverse().forEach(({ start, end }) => {
      tr = tr.delete(start, end);
    });

    const updatedTable = tr.doc.nodeAt(context.tablePos);
    if (updatedTable && updatedTable.type.name === 'table') {
      const firstRow = updatedTable.child(0);
      if (firstRow && firstRow.type.name === 'tableRow') {
        let cellPos = context.tablePos + 2;
        firstRow.forEach((cell) => {
          if (cell.type.name === 'tableCell' || cell.type.name === 'tableHeader') {
            tr = tr.setNodeMarkup(cellPos, undefined, {
              ...cell.attrs,
              width: newColWidthPercent,
              widthType: 'pct',
            });
          }
          cellPos += cell.nodeSize;
        });
      }

      // Update table columnWidths to match new column count.
      const colCount = firstRow?.childCount ?? newColumnCount;
      const tableWidthTwips = (updatedTable.attrs.width as number) || 9360;
      const colWidthTwips = Math.floor(tableWidthTwips / Math.max(1, colCount));
      tr = tr.setNodeMarkup(context.tablePos, undefined, {
        ...updatedTable.attrs,
        columnWidths: Array(colCount).fill(colWidthTwips),
      });
    }

    dispatch(tr.scrollIntoView());
  }
  return true;
};

export const deleteTable: Command = (
  state: EditorState,
  dispatch?: (tr: Transaction) => void
): boolean => {
  const context = getTableContext(state);
  if (!context.isInTable || context.tablePos === undefined || !context.table) return false;

  if (dispatch) {
    const tr = state.tr;
    tr.delete(context.tablePos, context.tablePos + context.table.nodeSize);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Delete the entire table when a CellSelection covers every cell — used as
 * the first link of the Backspace/Delete keymap chain so a user can select
 * an entire table and remove it with one keypress.
 */
export const deleteTableIfSelected: Command = (state, dispatch) => {
  const selection = state.selection as CellSelection;
  const isCellSel = '$anchorCell' in selection && typeof selection.forEachCell === 'function';
  if (!isCellSel) return false;

  const context = getTableContext(state);
  if (!context.isInTable || context.tablePos === undefined || !context.table) return false;

  let totalCells = 0;
  context.table.descendants((node) => {
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      totalCells += 1;
    }
  });

  let selectedCells = 0;
  selection.forEachCell(() => {
    selectedCells += 1;
  });

  const isFullTable = totalCells > 0 && selectedCells >= totalCells;

  if (!isFullTable) return false;

  if (dispatch) {
    const tr = state.tr.delete(context.tablePos, context.tablePos + context.table.nodeSize);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Keep the spacer paragraph between two adjacent tables — when the cursor
 * is in an empty paragraph that sits between tables, intercept Backspace/
 * Delete so the paragraph isn't merged into the surrounding tables.
 */
export const preventTableMergeAtGap: Command = (state) => {
  const { $from, empty } = state.selection;
  if (!empty) return false;

  const parent = $from.parent;
  if (parent.type.name !== 'paragraph') return false;
  if (parent.textContent.length > 0) return false;

  const depth = $from.depth;
  if (depth < 1) return false;
  const container = $from.node(depth - 1);
  const index = $from.index(depth - 1);
  const before = index > 0 ? container.child(index - 1) : null;
  const after = index + 1 < container.childCount ? container.child(index + 1) : null;
  const beforeIsTable = before?.type.name === 'table';
  const afterIsTable = after?.type.name === 'table';
  if (beforeIsTable || afterIsTable) {
    // Keep the spacer paragraph adjacent to tables so they can't visually merge.
    return true;
  }

  return false;
};
