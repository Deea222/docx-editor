import { describe, expect, test } from 'bun:test';
import type { Table, TableCell, TableRow } from '@docx-editor.dev/core/types/document';
import { addRow } from './operations';

const cell = (vMerge?: 'restart' | 'continue'): TableCell => ({
  type: 'tableCell',
  content: vMerge === 'continue' ? [] : [{ type: 'paragraph', content: [], formatting: {} }],
  formatting: vMerge ? { vMerge } : {},
});

const row = (...cells: TableCell[]): TableRow => ({ type: 'tableRow', cells });

// Column 0 is vertically merged across all three rows; column 1 is independent.
const mergedTable = (): Table => ({
  type: 'table',
  rows: [
    row(cell('restart'), cell()),
    row(cell('continue'), cell()),
    row(cell('continue'), cell()),
  ],
});

const mergeColumn = (table: Table): (string | null)[] =>
  table.rows.map((r) => r.cells[0]?.formatting?.vMerge ?? null);

describe('addRow', () => {
  test('extends a vertical merge when inserting inside it', () => {
    const result = addRow(mergedTable(), 1, 'after');
    expect(mergeColumn(result)).toEqual(['restart', 'continue', 'continue', 'continue']);
  });

  test('does not extend a vertical merge when inserting below it', () => {
    const result = addRow(mergedTable(), 2, 'after');
    expect(mergeColumn(result)).toEqual(['restart', 'continue', 'continue', null]);
  });
});
