import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, beforeAll } from 'vitest';
import type { Cell, WorkbookSnapshot } from './types.js';
import { fmtEqual, normalizeHex } from './format.js';
import { parseXlsx } from './parse.js';
import { writeXlsx } from './write.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURES = join(__dirname, '../../fixtures');
const SAMPLE_PATH = join(FIXTURES, 'sample-budget.xlsx');

/** FP&A-like budget assumptions + formulas */
function buildBudgetSnapshot(): WorkbookSnapshot {
  return {
    sheetOrder: ['Budget', 'Assumptions'],
    sheets: {
      Budget: {
        dimensions: { rows: 12, cols: 6 },
        columnWidths: { A: 18, B: 14, C: 14, D: 16 },
        rowHeights: { '1': 22 },
        freeze: { row: 1, col: 1 },
        merges: ['A8:B8'],
        hiddenRows: [3],
        tabColor: '#4472C4',
        cells: {
          A1: {
            v: 'Line Item',
            f: null,
            fmt: {
              bold: true,
              fill: '#4472C4',
              fontColor: '#FFFFFF',
              alignment: { horizontal: 'center', vertical: 'middle' },
            },
          },
          B1: {
            v: 'Q1',
            f: null,
            fmt: { bold: true, fill: '#4472C4', fontColor: '#FFFFFF' },
          },
          C1: {
            v: 'Q2',
            f: null,
            fmt: { bold: true, fill: '#4472C4', fontColor: '#FFFFFF' },
          },
          D1: {
            v: 'H1 Total',
            f: null,
            fmt: { bold: true, fill: '#4472C4', fontColor: '#FFFFFF' },
          },
          A2: {
            v: 'Revenue',
            f: null,
            fmt: {
              bold: true,
              underline: true,
              alignment: { wrapText: true },
            },
          },
          B2: {
            v: 100000,
            f: null,
            fmt: {
              numFmt: '#,##0.00',
              borders: {
                top: { style: 'thin', color: '#000000' },
                bottom: { style: 'thin', color: '#000000' },
                left: { style: 'thin', color: '#000000' },
                right: { style: 'thin', color: '#000000' },
              },
            },
          },
          C2: { v: 110000, f: null, fmt: { numFmt: '#,##0.00' } },
          D2: {
            v: null,
            f: '=B2+C2',
            fmt: { numFmt: '#,##0.00', bold: true },
          },
          A3: { v: 'COGS', f: null },
          B3: { v: 40000, f: null, fmt: { numFmt: '#,##0.00' } },
          C3: { v: 44000, f: null, fmt: { numFmt: '#,##0.00' } },
          D3: { v: null, f: '=B3+C3', fmt: { numFmt: '#,##0.00' } },
          A4: { v: 'Gross Profit', f: null, fmt: { bold: true, italic: true } },
          B4: {
            v: null,
            f: '=B2-B3',
            fmt: { numFmt: '#,##0.00', bold: true, fill: '#FFFF00' },
          },
          C4: {
            v: null,
            f: '=C2-C3',
            fmt: { numFmt: '#,##0.00', bold: true, fill: '#FFFF00' },
          },
          D4: {
            v: null,
            f: '=D2-D3',
            fmt: { numFmt: '#,##0.00', bold: true, fill: '#FFFF00' },
          },
          A6: { v: 'Tax rate', f: null },
          B6: {
            v: null,
            f: '=Assumptions!B1',
            fmt: { numFmt: '0.00%' },
          },
          A7: { v: 'Tax', f: null },
          B7: {
            v: null,
            f: '=B4*B6',
            fmt: { numFmt: '#,##0.00', fontColor: '#FF0000', strike: true },
          },
          A8: { v: 'Source notes', f: null, fmt: { italic: true, fontSize: 9 } },
          C6: { v: null, f: null, fmt: { fill: '#F4B183' } },
        },
      },
      Assumptions: {
        dimensions: { rows: 3, cols: 2 },
        columnWidths: { C: 12 },
        hiddenColumns: ['C'],
        cells: {
          A1: { v: 'Tax Rate', f: null, fmt: { bold: true } },
          B1: { v: 0.21, f: null, fmt: { numFmt: '0.00%', fill: '#E2EFDA' } },
          A2: { v: 'Notes', f: null, fmt: { italic: true } },
          B2: { v: 'FY26 planning assumptions', f: null },
        },
      },
    },
  };
}



function cellEqual(a: Cell, b: Cell): boolean {
  // Formula is authoritative; compare f and tracked fmt.
  // Values: if both have formulas, v may be null/cached differently — compare when both non-null.
  if ((a.f ?? null) !== (b.f ?? null)) return false;
  if (!fmtEqual(a.fmt, b.fmt)) return false;
  if (a.f) {
    // Formula cells: allow null vs computed mismatch on v
    if (a.v !== null && b.v !== null && a.v !== b.v) return false;
    return true;
  }
  return a.v === b.v;
}

function assertSnapshotRoundTrip(
  original: WorkbookSnapshot,
  parsed: WorkbookSnapshot,
): void {
  expect(parsed.sheetOrder).toEqual(original.sheetOrder);
  expect(Object.keys(parsed.sheets).sort()).toEqual(
    Object.keys(original.sheets).sort(),
  );

  for (const name of original.sheetOrder) {
    const oSheet = original.sheets[name];
    const pSheet = parsed.sheets[name];
    expect(pSheet, `missing sheet ${name}`).toBeTruthy();

    const oAddrs = Object.keys(oSheet.cells).sort();
    const pAddrs = Object.keys(pSheet.cells).sort();
    expect(pAddrs).toEqual(oAddrs);
    expect(pSheet.columnWidths).toEqual(oSheet.columnWidths);
    expect(pSheet.rowHeights).toEqual(oSheet.rowHeights);
    expect(pSheet.merges).toEqual(oSheet.merges);
    expect(pSheet.freeze).toEqual(oSheet.freeze);
    expect(pSheet.hiddenRows).toEqual(oSheet.hiddenRows);
    expect(!!pSheet.hidden).toBe(!!oSheet.hidden);
    if (oSheet.tabColor) {
      expect(normalizeHex(pSheet.tabColor)).toBe(normalizeHex(oSheet.tabColor));
    }

    for (const addr of oAddrs) {
      const ok = cellEqual(oSheet.cells[addr], pSheet.cells[addr]);
      if (!ok) {
        expect.fail(
          `Cell ${name}!${addr} mismatch:\n` +
            `  expected: ${JSON.stringify(oSheet.cells[addr])}\n` +
            `  actual:   ${JSON.stringify(pSheet.cells[addr])}`,
        );
      }
    }
  }
}

describe('xlsx bridge round-trip', () => {
  const original = buildBudgetSnapshot();
  let firstWrite: Buffer;
  let afterParse: WorkbookSnapshot;
  let secondWrite: Buffer;
  let afterSecondParse: WorkbookSnapshot;

  beforeAll(async () => {
    await mkdir(FIXTURES, { recursive: true });
    firstWrite = await writeXlsx(original);
    await writeFile(SAMPLE_PATH, firstWrite);
    afterParse = await parseXlsx(firstWrite);
    secondWrite = await writeXlsx(afterParse);
    afterSecondParse = await parseXlsx(secondWrite);
  });

  it('writes a non-empty xlsx buffer', () => {
    expect(firstWrite.length).toBeGreaterThan(100);
    // ZIP signature
    expect(firstWrite[0]).toBe(0x50);
    expect(firstWrite[1]).toBe(0x4b);
  });

  it('preserves sheet order and names', () => {
    expect(afterParse.sheetOrder).toEqual(['Budget', 'Assumptions']);
  });

  it('round-trips values, formulas, and basic formatting', () => {
    assertSnapshotRoundTrip(original, afterParse);
  });

  it('is stable across a second write→parse cycle', () => {
    assertSnapshotRoundTrip(afterParse, afterSecondParse);
  });

  it('tracks bold, italic, fill, fontColor, and numFmt on key cells', () => {
    const b4 = afterParse.sheets.Budget.cells.B4;
    expect(b4.f).toBe('=B2-B3');
    expect(b4.fmt?.bold).toBe(true);
    expect(normalizeHex(b4.fmt?.fill)).toBe('FFFF00');
    expect(b4.fmt?.numFmt).toBe('#,##0.00');

    const a4 = afterParse.sheets.Budget.cells.A4;
    expect(a4.fmt?.italic).toBe(true);
    expect(a4.fmt?.bold).toBe(true);

    const b7 = afterParse.sheets.Budget.cells.B7;
    expect(b7.f).toBe('=B4*B6');
    expect(normalizeHex(b7.fmt?.fontColor)).toBe('FF0000');

    const tax = afterParse.sheets.Assumptions.cells.B1;
    expect(tax.v).toBe(0.21);
    expect(tax.fmt?.numFmt).toBe('0.00%');
    expect(normalizeHex(tax.fmt?.fill)).toBe('E2EFDA');
  });

  it('parses from filesystem path', async () => {
    const fromPath = await parseXlsx(SAMPLE_PATH);
    expect(fromPath.sheetOrder).toEqual(original.sheetOrder);
    expect(fromPath.sheets.Budget.cells.D2.f).toBe('=B2+C2');
  });

  it('commits fixture sample-budget.xlsx for CLI demos', async () => {
    const again = await parseXlsx(SAMPLE_PATH);
    expect(Object.keys(again.sheets.Budget.cells).length).toBeGreaterThan(10);
  });

  it('round-trips borders, alignment, underline, strike, layout', () => {
    const b2 = afterParse.sheets.Budget.cells.B2;
    expect(b2.fmt?.borders?.top?.style).toBe('thin');
    expect(normalizeHex(b2.fmt?.borders?.top?.color)).toBe('000000');

    const a2 = afterParse.sheets.Budget.cells.A2;
    expect(a2.fmt?.underline).toBeTruthy();
    expect(a2.fmt?.alignment?.wrapText).toBe(true);

    const b7 = afterParse.sheets.Budget.cells.B7;
    expect(b7.fmt?.strike).toBe(true);

    const c6 = afterParse.sheets.Budget.cells.C6;
    expect(c6.v).toBeNull();
    expect(c6.f).toBeNull();
    expect(normalizeHex(c6.fmt?.fill)).toBe('F4B183');

    const budget = afterParse.sheets.Budget;
    expect(budget.columnWidths?.A).toBeGreaterThan(10);
    expect(budget.rowHeights?.['1']).toBe(22);
    expect(budget.merges).toContain('A8:B8');
    expect(budget.freeze).toEqual({ row: 1, col: 1 });
    expect(budget.hiddenRows).toContain(3);
    expect(normalizeHex(budget.tabColor)).toBe('4472C4');
    expect(afterParse.sheets.Assumptions.hiddenColumns).toContain('C');
  });
});

describe('hyperlink URL tracking', () => {
  it('round-trips cell hyperlink URL', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['Links'],
      sheets: {
        Links: {
          dimensions: { rows: 2, cols: 2 },
          cells: {
            A1: { v: 'Docs', f: null, hyperlink: 'https://example.com/docs' },
            B1: { v: 1, f: null },
          },
        },
      },
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.sheets.Links.cells.A1.v).toBe('Docs');
    expect(parsed.sheets.Links.cells.A1.hyperlink).toBe('https://example.com/docs');
  });
});

describe('data validation tracking', () => {
  it('round-trips sheet validations', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['V'],
      sheets: {
        V: {
          dimensions: { rows: 2, cols: 2 },
          cells: {
            A1: { v: 5, f: null },
            B1: { v: 'Yes', f: null },
          },
          validations: [
            {
              sqref: 'A1',
              type: 'whole',
              operator: 'between',
              formulae: ['1', '10'],
              allowBlank: true,
              showErrorMessage: true,
              errorTitle: 'Out of range',
              error: 'Enter 1-10',
            },
            {
              sqref: 'B1',
              type: 'list',
              formulae: ['"Yes,No"'],
              allowBlank: true,
            },
          ],
        },
      },
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.sheets.V.validations).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sqref: 'A1',
          type: 'whole',
          operator: 'between',
          formulae: ['1', '10'],
          allowBlank: true,
        }),
        expect.objectContaining({
          sqref: 'B1',
          type: 'list',
          formulae: ['"Yes,No"'],
        }),
      ]),
    );
  });
});

describe('named range tracking', () => {
  it('round-trips workbook named ranges', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['Budget'],
      sheets: {
        Budget: {
          dimensions: { rows: 2, cols: 2 },
          cells: { A1: { v: 1, f: null }, B1: { v: 2, f: null } },
        },
      },
      names: [
        { name: 'TaxRate', refersTo: 'Budget!$A$1', scope: null },
        { name: 'LocalAmt', refersTo: 'Budget!$B$1', scope: 'Budget' },
      ],
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.names).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ name: 'TaxRate', refersTo: 'Budget!$A$1' }),
        expect.objectContaining({
          name: 'LocalAmt',
          refersTo: 'Budget!$B$1',
          scope: 'Budget',
        }),
      ]),
    );
  });
});

describe('cell comment tracking', () => {
  it('round-trips cell comment text', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['Notes'],
      sheets: {
        Notes: {
          dimensions: { rows: 1, cols: 1 },
          cells: {
            A1: { v: 'x', f: null, comment: { text: 'Check with FP&A' } },
          },
        },
      },
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.sheets.Notes.cells.A1.comment).toEqual({
      text: 'Check with FP&A',
    });
  });
});

describe('tables and autofilter tracking', () => {
  it('round-trips Excel table metadata', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['T'],
      sheets: {
        T: {
          dimensions: { rows: 3, cols: 2 },
          cells: {
            A1: { v: 'Name', f: null },
            B1: { v: 'Amt', f: null },
            A2: { v: 'x', f: null },
            B2: { v: 10, f: null },
          },
          tables: [
            { name: 'MyTable', ref: 'A1:B2', headerRow: true, totalsRow: false },
          ],
        },
      },
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.sheets.T.tables).toEqual([
      expect.objectContaining({
        name: 'MyTable',
        ref: 'A1:B2',
        headerRow: true,
        totalsRow: false,
      }),
    ]);
  });

  it('round-trips sheet autoFilter range', async () => {
    const snap: WorkbookSnapshot = {
      sheetOrder: ['F'],
      sheets: {
        F: {
          dimensions: { rows: 3, cols: 2 },
          cells: {
            A1: { v: 'A', f: null },
            B1: { v: 'B', f: null },
            A2: { v: 1, f: null },
            B2: { v: 2, f: null },
          },
          autoFilter: 'A1:B2',
        },
      },
    };
    const buf = await writeXlsx(snap);
    const parsed = await parseXlsx(buf);
    expect(parsed.sheets.F.autoFilter).toBe('A1:B2');
  });
});
