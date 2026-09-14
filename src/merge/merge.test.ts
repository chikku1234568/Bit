import { describe, expect, it } from 'vitest';
import type { Cell, WorkbookSnapshot } from '../xlsx/types.js';
import { applyResolutions, cellsEqual, mergeSnapshots } from './merge.js';

function cell(v: Cell['v'], f: Cell['f'] = null, fmt?: Cell['fmt']): Cell {
  return fmt ? { v, f, fmt } : { v, f };
}

function book(
  cells: Record<string, Cell>,
  sheet = 'Budget',
  extra?: WorkbookSnapshot['sheets'],
  order?: string[],
): WorkbookSnapshot {
  return {
    sheetOrder: order ?? [sheet, ...(extra ? Object.keys(extra) : [])].filter(
      (v, i, a) => a.indexOf(v) === i,
    ),
    sheets: {
      [sheet]: { dimensions: { rows: 10, cols: 5 }, cells },
      ...extra,
    },
  };
}

describe('M5 mergeSnapshots', () => {
  it('no-op merge', () => {
    const s = book({ A1: cell(1) });
    const r = mergeSnapshots(s, s, s);
    expect(r.conflicts).toHaveLength(0);
    expect(r.snapshot.sheets.Budget.cells.A1.v).toBe(1);
  });

  it('disjoint clean combine', () => {
    const base = book({ A1: cell(1), B1: cell(2) });
    const ours = book({ A1: cell(10), B1: cell(2) });
    const theirs = book({ A1: cell(1), B1: cell(20) });
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts).toHaveLength(0);
    expect(r.snapshot.sheets.Budget.cells.A1.v).toBe(10);
    expect(r.snapshot.sheets.Budget.cells.B1.v).toBe(20);
    expect(r.autoChangeCount).toBeGreaterThan(0);
  });

  it('conflict values', () => {
    const base = book({ A1: cell(1) });
    const ours = book({ A1: cell(2) });
    const theirs = book({ A1: cell(3) });
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts).toHaveLength(1);
    expect(r.conflicts[0]).toMatchObject({
      sheet: 'Budget',
      address: 'A1',
      reason: 'value',
    });
  });

  it('formula vs value → conflict', () => {
    const base = book({ A1: cell(1, null) });
    const ours = book({ A1: cell(null, '=B1') });
    const theirs = book({ A1: cell(99, null) });
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts.length).toBeGreaterThanOrEqual(1);
    expect(r.conflicts[0].address).toBe('A1');
  });

  it('formatting-only change on one side → auto-take', () => {
    const base = book({ A1: cell(1, null, { bold: false }) });
    const ours = book({ A1: cell(1, null, { bold: false }) });
    const theirs = book({ A1: cell(1, null, { bold: true }) });
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts).toHaveLength(0);
    expect(r.snapshot.sheets.Budget.cells.A1.fmt?.bold).toBe(true);
  });

  it('delete vs edit → conflict', () => {
    const base = book({ A1: cell(1), B1: cell(2) });
    const ours = book({ B1: cell(2) }); // deleted A1
    const theirs = book({ A1: cell(9), B1: cell(2) }); // edited A1
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts.some((c) => c.address === 'A1' && c.reason === 'delete-vs-edit')).toBe(
      true,
    );
  });

  it('edit vs delete → conflict', () => {
    const base = book({ A1: cell(1) });
    const ours = book({ A1: cell(5) });
    const theirs = book({}); // deleted
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts.some((c) => c.reason === 'edit-vs-delete')).toBe(true);
  });

  it('sheet added on one side → clean take', () => {
    const base = book({ A1: cell(1) });
    const ours = book({ A1: cell(1) });
    const theirs: WorkbookSnapshot = {
      sheetOrder: ['Budget', 'Tax'],
      sheets: {
        Budget: base.sheets.Budget,
        Tax: { dimensions: { rows: 2, cols: 2 }, cells: { A1: cell('rate') } },
      },
    };
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts).toHaveLength(0);
    expect(r.snapshot.sheetOrder).toContain('Tax');
    expect(r.snapshot.sheets.Tax.cells.A1.v).toBe('rate');
  });

  it('both same new value → take Y', () => {
    const base = book({ A1: cell(1) });
    const ours = book({ A1: cell(7) });
    const theirs = book({ A1: cell(7) });
    const r = mergeSnapshots(base, ours, theirs);
    expect(r.conflicts).toHaveLength(0);
    expect(r.snapshot.sheets.Budget.cells.A1.v).toBe(7);
  });

  it('applyResolutions clears conflicts', () => {
    const base = book({ A1: cell(1) });
    const ours = book({ A1: cell(2) });
    const theirs = book({ A1: cell(3) });
    const r = mergeSnapshots(base, ours, theirs);
    const applied = applyResolutions(r, {
      'Budget!A1': { action: 'keep-theirs' },
    });
    expect(applied.conflicts).toHaveLength(0);
    expect(applied.snapshot.sheets.Budget.cells.A1.v).toBe(3);
  });

  it('cellsEqual normalises fmt', () => {
    expect(cellsEqual(cell(1, null, { bold: true }), cell(1, null, { bold: true }))).toBe(
      true,
    );
    expect(cellsEqual(cell(1), cell(2))).toBe(false);
  });
});
