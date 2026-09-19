/**
 * Pure snapshot diff (M4). Snapshots in → DiffResult out.
 * No Fastify / ExcelJS / fs.
 */

import { fmtEqual, normalizeFmt } from '../xlsx/format.js';
import type {
  Cell,
  CellComment,
  NamedRange,
  Sheet,
  SheetTable,
  ValidationRule,
  WorkbookSnapshot,
} from '../xlsx/types.js';

export type DiffKind =
  | 'sheet-add'
  | 'sheet-remove'
  | 'sheet-order'
  | 'sheet-hidden'
  | 'sheet-tab-color'
  | 'sheet-freeze'
  | 'col-width'
  | 'row-height'
  | 'col-hidden'
  | 'row-hidden'
  | 'merge'
  | 'cell-add'
  | 'cell-remove'
  | 'cell-value'
  | 'cell-formula'
  | 'cell-format'
  | 'cell-hyperlink'
  | 'cell-comment'
  | 'validation'
  | 'named-range'
  | 'table'
  | 'auto-filter';

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

function commentKey(c: CellComment | undefined): string | null {
  if (!c) return null;
  return JSON.stringify({
    text: c.text,
    ...(c.author != null ? { author: c.author } : {}),
  });
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

  const bh = b.hyperlink ?? null;
  const ch = c.hyperlink ?? null;
  if (bh !== ch) {
    out.push({ kind: 'cell-hyperlink', sheet, address, before: bh, after: ch });
  }

  const bc = commentKey(b.comment);
  const cc = commentKey(c.comment);
  if (bc !== cc) {
    out.push({
      kind: 'cell-comment',
      sheet,
      address,
      before: b.comment ?? null,
      after: c.comment ?? null,
    });
  }

  return out;
}

function mapDiff(
  kind: DiffKind,
  sheet: string,
  before: Record<string, unknown> | undefined,
  after: Record<string, unknown> | undefined,
  addressPrefix: string,
  changes: DiffEntry[],
): void {
  const b = before ?? {};
  const a = after ?? {};
  const keys = new Set([...Object.keys(b), ...Object.keys(a)]);
  for (const k of [...keys].sort()) {
    if (JSON.stringify(b[k] ?? null) === JSON.stringify(a[k] ?? null)) continue;
    changes.push({
      kind,
      sheet,
      address: `${addressPrefix}${k}`,
      before: b[k],
      after: a[k],
    });
  }
}

function listDiff(
  kind: DiffKind,
  sheet: string,
  before: Array<string | number> | undefined,
  after: Array<string | number> | undefined,
  changes: DiffEntry[],
): void {
  const b = [...(before ?? [])].map(String).sort();
  const a = [...(after ?? [])].map(String).sort();
  if (JSON.stringify(b) === JSON.stringify(a)) return;
  changes.push({ kind, sheet, before: b, after: a });
}

function validationMap(
  rules: ValidationRule[] | undefined,
): Record<string, ValidationRule> {
  const out: Record<string, ValidationRule> = {};
  for (const r of rules ?? []) {
    out[r.sqref] = r;
  }
  return out;
}

function tableMap(tables: SheetTable[] | undefined): Record<string, SheetTable> {
  const out: Record<string, SheetTable> = {};
  for (const t of tables ?? []) {
    out[t.name] = t;
  }
  return out;
}

function namedRangeKey(n: NamedRange): string {
  return `${n.scope ?? ''}|${n.name}`;
}

function diffValidations(
  sheet: string,
  before: ValidationRule[] | undefined,
  after: ValidationRule[] | undefined,
  changes: DiffEntry[],
): void {
  const b = validationMap(before);
  const a = validationMap(after);
  const keys = new Set([...Object.keys(b), ...Object.keys(a)]);
  for (const sqref of [...keys].sort()) {
    if (JSON.stringify(b[sqref] ?? null) === JSON.stringify(a[sqref] ?? null)) continue;
    changes.push({
      kind: 'validation',
      sheet,
      address: sqref,
      before: b[sqref] ?? null,
      after: a[sqref] ?? null,
    });
  }
}

function diffTables(
  sheet: string,
  before: SheetTable[] | undefined,
  after: SheetTable[] | undefined,
  changes: DiffEntry[],
): void {
  const b = tableMap(before);
  const a = tableMap(after);
  const keys = new Set([...Object.keys(b), ...Object.keys(a)]);
  for (const name of [...keys].sort()) {
    if (JSON.stringify(b[name] ?? null) === JSON.stringify(a[name] ?? null)) continue;
    changes.push({
      kind: 'table',
      sheet,
      address: name,
      before: b[name] ?? null,
      after: a[name] ?? null,
    });
  }
}

function diffNamedRanges(
  before: NamedRange[] | undefined,
  after: NamedRange[] | undefined,
  changes: DiffEntry[],
): void {
  const bMap: Record<string, NamedRange> = {};
  const aMap: Record<string, NamedRange> = {};
  for (const n of before ?? []) bMap[namedRangeKey(n)] = n;
  for (const n of after ?? []) aMap[namedRangeKey(n)] = n;
  const keys = new Set([...Object.keys(bMap), ...Object.keys(aMap)]);
  for (const k of [...keys].sort()) {
    if (JSON.stringify(bMap[k] ?? null) === JSON.stringify(aMap[k] ?? null)) continue;
    const name = (aMap[k] ?? bMap[k])?.name;
    changes.push({
      kind: 'named-range',
      address: name,
      before: bMap[k] ?? null,
      after: aMap[k] ?? null,
    });
  }
}

function diffSheetLayout(sheet: string, base: Sheet | undefined, compare: Sheet | undefined, changes: DiffEntry[]): void {
  if (!base && !compare) return;
  mapDiff('col-width', sheet, base?.columnWidths, compare?.columnWidths, 'col:', changes);
  mapDiff('row-height', sheet, base?.rowHeights, compare?.rowHeights, 'row:', changes);
  listDiff('col-hidden', sheet, base?.hiddenColumns, compare?.hiddenColumns, changes);
  listDiff('row-hidden', sheet, base?.hiddenRows, compare?.hiddenRows, changes);
  listDiff('merge', sheet, base?.merges, compare?.merges, changes);

  const bHidden = !!(base?.hidden || base?.veryHidden);
  const cHidden = !!(compare?.hidden || compare?.veryHidden);
  if (bHidden !== cHidden) {
    changes.push({
      kind: 'sheet-hidden',
      sheet,
      before: base?.veryHidden ? 'veryHidden' : base?.hidden ? 'hidden' : 'visible',
      after: compare?.veryHidden ? 'veryHidden' : compare?.hidden ? 'hidden' : 'visible',
    });
  }
  if ((base?.tabColor ?? null) !== (compare?.tabColor ?? null)) {
    changes.push({
      kind: 'sheet-tab-color',
      sheet,
      before: base?.tabColor,
      after: compare?.tabColor,
    });
  }
  if (JSON.stringify(base?.freeze ?? null) !== JSON.stringify(compare?.freeze ?? null)) {
    changes.push({
      kind: 'sheet-freeze',
      sheet,
      before: base?.freeze,
      after: compare?.freeze,
    });
  }

  diffValidations(sheet, base?.validations, compare?.validations, changes);
  diffTables(sheet, base?.tables, compare?.tables, changes);

  const bAf = base?.autoFilter ?? null;
  const cAf = compare?.autoFilter ?? null;
  if (bAf !== cAf) {
    changes.push({
      kind: 'auto-filter',
      sheet,
      before: bAf,
      after: cAf,
    });
  }
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
    diffSheetLayout(sheetName, base.sheets[sheetName], compare.sheets[sheetName], changes);
  }

  diffNamedRanges(base.names, compare.names, changes);

  return { changes };
}
