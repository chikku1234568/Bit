import { describe, expect, it } from 'vitest';
import type { WorkbookSnapshot } from '../xlsx/types.js';
import { diffSnapshots } from './diff.js';

function snap(
  cells: WorkbookSnapshot['sheets'][string]['cells'],
  sheetOrder: string[] = ['Budget'],
  extraSheets: WorkbookSnapshot['sheets'] = {},
): WorkbookSnapshot {
  return {
    sheetOrder,
    sheets: {
      Budget: {
        dimensions: { rows: 10, cols: 5 },
        cells,
      },
      ...extraSheets,
    },
  };
}

describe('M4 diffSnapshots', () => {
  it('identical empty → no changes', () => {
    const a = snap({});
    expect(diffSnapshots(a, a).changes).toEqual([]);
    expect(diffSnapshots(snap({}), snap({})).changes).toEqual([]);
  });

  it('value change', () => {
    const base = snap({ A1: { v: 1, f: null } });
    const compare = snap({ A1: { v: 2, f: null } });
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'cell-value' && c.address === 'A1')).toBe(true);
    expect(r.changes.find((c) => c.kind === 'cell-value')).toMatchObject({
      before: 1,
      after: 2,
    });
  });

  it('formula change', () => {
    const base = snap({ B1: { v: null, f: '=A1' } });
    const compare = snap({ B1: { v: null, f: '=A1*2' } });
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'cell-formula')).toBe(true);
  });

  it('format-only change', () => {
    const base = snap({ A1: { v: 1, f: null, fmt: { bold: true } } });
    const compare = snap({ A1: { v: 1, f: null, fmt: { bold: false } } });
    const r = diffSnapshots(base, compare);
    expect(r.changes).toHaveLength(1);
    expect(r.changes[0].kind).toBe('cell-format');
  });

  it('sheet add', () => {
    const base = snap({ A1: { v: 1, f: null } });
    const compare: WorkbookSnapshot = {
      sheetOrder: ['Budget', 'Assumptions'],
      sheets: {
        Budget: base.sheets.Budget,
        Assumptions: {
          dimensions: { rows: 2, cols: 2 },
          cells: { A1: { v: 'tax', f: null } },
        },
      },
    };
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'sheet-add' && c.sheet === 'Assumptions')).toBe(
      true,
    );
  });

  it('cell delete', () => {
    const base = snap({ A1: { v: 1, f: null }, B1: { v: 2, f: null } });
    const compare = snap({ A1: { v: 1, f: null } });
    const r = diffSnapshots(base, compare);
    expect(
      r.changes.some((c) => c.kind === 'cell-remove' && c.address === 'B1'),
    ).toBe(true);
  });

  it('same formula with v null vs 123 is NOT a value change', () => {
    const base = snap({ B3: { v: null, f: '=B2*0.2' } });
    const compare = snap({ B3: { v: 123, f: '=B2*0.2' } });
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'cell-value')).toBe(false);
    expect(r.changes).toEqual([]);
  });

  it('column width and freeze show as layout changes', () => {
    const base = snap({ A1: { v: 1, f: null } });
    const compare: WorkbookSnapshot = {
      sheetOrder: ['Budget'],
      sheets: {
        Budget: {
          dimensions: { rows: 10, cols: 5 },
          cells: { A1: { v: 1, f: null } },
          columnWidths: { A: 24 },
          freeze: { row: 1, col: 0 },
        },
      },
    };
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'col-width' && c.address === 'col:A')).toBe(
      true,
    );
    expect(r.changes.some((c) => c.kind === 'sheet-freeze')).toBe(true);
  });
  it('hyperlink URL change', () => {
    const base = snap({ A1: { v: 'Docs', f: null, hyperlink: 'https://a.example' } });
    const compare = snap({ A1: { v: 'Docs', f: null, hyperlink: 'https://b.example' } });
    const r = diffSnapshots(base, compare);
    expect(r.changes.some((c) => c.kind === 'cell-hyperlink')).toBe(true);
    expect(r.changes.find((c) => c.kind === 'cell-hyperlink')).toMatchObject({
      before: 'https://a.example',
      after: 'https://b.example',
    });
  });
});
