/**
 * Pure snapshot diff (M4). Snapshots in → DiffResult out.
 * No Fastify / ExcelJS / fs.
 */

import type { Cell, CellFormat, WorkbookSnapshot } from '../xlsx/types.js';

export type DiffKind =
  | 'sheet-add'
  | 'sheet-remove'
  | 'sheet-order'
  | 'cell-add'
  | 'cell-remove'
  | 'cell-value'
  | 'cell-formula'
  | 'cell-format';

export interface DiffEntry {
  kind: DiffKind;
  sheet?: string;
  address?: string;
  before?: unknown;
  after?: unknown;
}

export interface DiffResult {
  changes: DiffEntry[];
}

function normalizeFmt(fmt: CellFormat | undefined): CellFormat | null {
  if (!fmt) return null;
  const out: CellFormat = {};
  if (fmt.numFmt !== undefined) out.numFmt = fmt.numFmt;
  if (fmt.bold !== undefined) out.bold = fmt.bold;
  if (fmt.italic !== undefined) out.italic = fmt.italic;
  if (fmt.fill !== undefined) out.fill = fmt.fill;
  if (fmt.fontColor !== undefined) out.fontColor = fmt.fontColor;
  return Object.keys(out).length === 0 ? null : out;
}

function fmtEqual(a: CellFormat | undefined, b: CellFormat | undefined): boolean {
  return JSON.stringify(normalizeFmt(a)) === JSON.stringify(normalizeFmt(b));
}

/**
 * Compare two cells. Formula is authoritative for value:
 * same `f` with different cached `v` (e.g. null vs 123) is NOT a value change.
 */
function diffCell(
  sheet: string,
  address: string,
  base: Cell | undefined,
  compare: Cell | undefined,
): DiffEntry[] {
  const out: DiffEntry[] = [];
  if (!base && !compare) return out;
  if (!base && compare) {
    out.push({ kind: 'cell-add', sheet, address, after: compare });
    return out;
  }
  if (base && !compare) {
    out.push({ kind: 'cell-remove', sheet, address, before: base });
    return out;
  }
  const b = base!;
  const c = compare!;

  const baseF = b.f ?? null;
  const compareF = c.f ?? null;
  if (baseF !== compareF) {
    out.push({ kind: 'cell-formula', sheet, address, before: baseF, after: compareF });
  }

  // Value: only if formulas differ in a way that isn't "same formula", or both have no formula.
  // Same non-null formula → ignore v divergence (cached vs null).
  const sameFormula = baseF === compareF;
  if (sameFormula && baseF !== null) {
    // formula authoritative — skip value
  } else if (b.v !== c.v) {
    out.push({ kind: 'cell-value', sheet, address, before: b.v, after: c.v });
  }

  if (!fmtEqual(b.fmt, c.fmt)) {
    out.push({
      kind: 'cell-format',
      sheet,
      address,
      before: normalizeFmt(b.fmt),
      after: normalizeFmt(c.fmt),
    });
  }

  return out;
}

export function diffSnapshots(
  base: WorkbookSnapshot,
  compare: WorkbookSnapshot,
): DiffResult {
  const changes: DiffEntry[] = [];

  const baseOrder = base.sheetOrder ?? [];
  const compareOrder = compare.sheetOrder ?? [];
  const baseSet = new Set(baseOrder);
  const compareSet = new Set(compareOrder);

  for (const name of compareOrder) {
    if (!baseSet.has(name)) {
      changes.push({ kind: 'sheet-add', sheet: name });
    }
  }
  for (const name of baseOrder) {
    if (!compareSet.has(name)) {
      changes.push({ kind: 'sheet-remove', sheet: name });
    }
  }

  // Order change among shared sheets (compare relative order)
  const sharedBase = baseOrder.filter((n) => compareSet.has(n));
  const sharedCompare = compareOrder.filter((n) => baseSet.has(n));
  if (JSON.stringify(sharedBase) !== JSON.stringify(sharedCompare)) {
    changes.push({
      kind: 'sheet-order',
      before: sharedBase,
      after: sharedCompare,
    });
  }

  const allSheets = new Set([...baseOrder, ...compareOrder, ...Object.keys(base.sheets), ...Object.keys(compare.sheets)]);
  for (const sheetName of allSheets) {
    if (!baseSet.has(sheetName) && compareSet.has(sheetName)) {
      // New sheet: report cell-adds for its cells (optional — sheet-add is enough structurally;
      // still emit cell-adds so UI can show content).
      const cells = compare.sheets[sheetName]?.cells ?? {};
      for (const [addr, cell] of Object.entries(cells)) {
        changes.push({ kind: 'cell-add', sheet: sheetName, address: addr, after: cell });
      }
      continue;
    }
    if (baseSet.has(sheetName) && !compareSet.has(sheetName)) {
      const cells = base.sheets[sheetName]?.cells ?? {};
      for (const [addr, cell] of Object.entries(cells)) {
        changes.push({ kind: 'cell-remove', sheet: sheetName, address: addr, before: cell });
      }
      continue;
    }
    // Both sides (or present in sheets map)
    const baseCells = base.sheets[sheetName]?.cells ?? {};
    const compareCells = compare.sheets[sheetName]?.cells ?? {};
    const addrs = new Set([...Object.keys(baseCells), ...Object.keys(compareCells)]);
    for (const addr of [...addrs].sort()) {
      changes.push(...diffCell(sheetName, addr, baseCells[addr], compareCells[addr]));
    }
  }

  return { changes };
}

// silence unused in case of tree-shake
