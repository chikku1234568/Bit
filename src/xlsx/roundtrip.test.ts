import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, beforeAll } from 'vitest';
import type { Cell, CellFormat, WorkbookSnapshot } from './types.js';
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
        cells: {
          A1: {
            v: 'Line Item',
            f: null,
            fmt: { bold: true, fill: '#4472C4', fontColor: '#FFFFFF' },
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
          A2: { v: 'Revenue', f: null, fmt: { bold: true } },
          B2: { v: 100000, f: null, fmt: { numFmt: '#,##0.00' } },
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
            fmt: { numFmt: '#,##0.00', fontColor: '#FF0000' },
          },
        },
      },
      Assumptions: {
        dimensions: { rows: 3, cols: 2 },
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

function normalizeHex(c: string | undefined): string | undefined {
  if (!c) return undefined;
  return c.replace(/^#/, '').toUpperCase();
}

function fmtEqual(a?: CellFormat, b?: CellFormat): boolean {
  if (!a && !b) return true;
  if (!a || !b) {
    // Treat missing vs empty-equivalent as equal for unset flags
    const left = a ?? {};
    const right = b ?? {};
    return (
      left.numFmt === right.numFmt &&
      !!left.bold === !!right.bold &&
      !!left.italic === !!right.italic &&
      normalizeHex(left.fill) === normalizeHex(right.fill) &&
      normalizeHex(left.fontColor) === normalizeHex(right.fontColor)
    );
  }
  return (
    (a.numFmt ?? undefined) === (b.numFmt ?? undefined) &&
    !!a.bold === !!b.bold &&
    !!a.italic === !!b.italic &&
    normalizeHex(a.fill) === normalizeHex(b.fill) &&
    normalizeHex(a.fontColor) === normalizeHex(b.fontColor)
  );
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
    // already written in beforeAll
    const again = await parseXlsx(SAMPLE_PATH);
    expect(Object.keys(again.sheets.Budget.cells).length).toBeGreaterThan(10);
  });
});
