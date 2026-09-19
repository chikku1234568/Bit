/**
 * Pure three-way merge (M5). Snapshots in → snapshot + conflicts out.
 * No Fastify / ExcelJS / fs. Address-stable only.
 */

import { fmtEqual } from '../xlsx/format.js';
import type { Cell, FreezePane, Sheet, WorkbookSnapshot } from '../xlsx/types.js';

export interface CellConflict {
  sheet: string;
  address: string;
  base: Cell | null;
  ours: Cell | null;
  theirs: Cell | null;
  reason: 'value' | 'delete-vs-edit' | 'edit-vs-delete' | 'both-added';
}

export type ResolutionChoice =
  | { action: 'keep-ours' }
  | { action: 'keep-theirs' }
  | { action: 'edit'; cell: Cell };

export interface MergeResult {
  snapshot: WorkbookSnapshot;
  conflicts: CellConflict[];
  autoChangeCount: number;
}

/** Cell equality: v, f, and tracked fmt match (normalised). */
export function cellsEqual(a: Cell | null | undefined, b: Cell | null | undefined): boolean {
  if (!a && !b) return true;
  if (!a || !b) return false;
  if ((a.f ?? null) !== (b.f ?? null)) return false;
  if (a.v !== b.v) return false;
  return fmtEqual(a.fmt, b.fmt);
}

function cloneCell(c: Cell | null | undefined): Cell | null {
  if (!c) return null;
  return JSON.parse(JSON.stringify(c)) as Cell;
}

type Side = 'ours' | 'theirs' | 'conflict' | 'delete' | 'base' | 'none';

/**
 * Per-cell three-way outcome from design §9.2.
 * Returns which side to take, or conflict.
 */
function resolveCell(
  base: Cell | null | undefined,
  ours: Cell | null | undefined,
  theirs: Cell | null | undefined,
): { side: Side; reason?: CellConflict['reason'] } {
  const B = base ?? null;
  const O = ours ?? null;
  const T = theirs ?? null;
  const hasB = B !== null;
  const hasO = O !== null;
  const hasT = T !== null;

  // All missing
  if (!hasB && !hasO && !hasT) return { side: 'none' };

  // missing | Y | missing → Ours
  if (!hasB && hasO && !hasT) return { side: 'ours' };
  // missing | missing | Y → Theirs
  if (!hasB && !hasO && hasT) return { side: 'theirs' };
  // missing | Y | Z → conflict if Y≠Z else Y
  if (!hasB && hasO && hasT) {
    if (cellsEqual(O, T)) return { side: 'ours' };
    return { side: 'conflict', reason: 'both-added' };
  }

  // From here base exists (X)
  // X | X | X → no change
  if (cellsEqual(B, O) && cellsEqual(B, T)) return { side: 'base' };
  // X | Y | X → Ours
  if (!cellsEqual(B, O) && cellsEqual(B, T) && hasO) return { side: 'ours' };
  // X | X | Y → Theirs
  if (cellsEqual(B, O) && !cellsEqual(B, T) && hasT) return { side: 'theirs' };
  // X | Y | Y → Y (both same)
  if (hasO && hasT && cellsEqual(O, T) && !cellsEqual(B, O)) return { side: 'ours' };

  // X | missing | X → delete on Ours
  if (hasB && !hasO && cellsEqual(B, T)) return { side: 'delete' };
  // X | X | missing → delete on Theirs
  if (hasB && cellsEqual(B, O) && !hasT) return { side: 'delete' };
  // X | missing | Y → conflict delete vs edit
  if (hasB && !hasO && hasT && !cellsEqual(B, T)) {
    return { side: 'conflict', reason: 'delete-vs-edit' };
  }
  // X | Y | missing → conflict edit vs delete
  if (hasB && hasO && !hasT && !cellsEqual(B, O)) {
    return { side: 'conflict', reason: 'edit-vs-delete' };
  }

  // X | Y | Z (Y≠Z) → conflict
  if (hasO && hasT && !cellsEqual(O, T)) {
    return { side: 'conflict', reason: 'value' };
  }

  // Fallback
  if (hasO && !hasT) return { side: 'ours' };
  if (!hasO && hasT) return { side: 'theirs' };
  return { side: 'base' };
}

function eqJson(a: unknown, b: unknown): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/** Conflict on layout → keep ours (Main) so Combine is not blocked on column width. */
function threeWayScalar<T>(base: T | undefined, ours: T | undefined, theirs: T | undefined): T | undefined {
  if (eqJson(ours, theirs)) return ours;
  if (eqJson(base, ours)) return theirs;
  if (eqJson(base, theirs)) return ours;
  return ours;
}

function threeWayMap(
  base?: Record<string, number>,
  ours?: Record<string, number>,
  theirs?: Record<string, number>,
): Record<string, number> | undefined {
  const keys = new Set([
    ...Object.keys(base ?? {}),
    ...Object.keys(ours ?? {}),
    ...Object.keys(theirs ?? {}),
  ]);
  const out: Record<string, number> = {};
  for (const k of keys) {
    const v = threeWayScalar(base?.[k], ours?.[k], theirs?.[k]);
    if (v != null) out[k] = v;
  }
  return Object.keys(out).length ? out : undefined;
}

function threeWaySet<T extends string | number>(
  base?: T[],
  ours?: T[],
  theirs?: T[],
): T[] | undefined {
  const b = new Set(base ?? []);
  const o = new Set(ours ?? []);
  const t = new Set(theirs ?? []);
  const all = new Set<T>([...(base ?? []), ...(ours ?? []), ...(theirs ?? [])]);
  const out: T[] = [];
  for (const item of all) {
    const inB = b.has(item);
    const inO = o.has(item);
    const inT = t.has(item);
    let keep = false;
    if (inO === inT) keep = inO;
    else if (inB === inO) keep = inT;
    else if (inB === inT) keep = inO;
    else keep = inO;
    if (keep) out.push(item);
  }
  if (out.length === 0) return undefined;
  return out.sort((a, b2) => (a < b2 ? -1 : a > b2 ? 1 : 0));
}

function mergeSheetLayout(base?: Sheet, ours?: Sheet, theirs?: Sheet): Partial<Sheet> {
  const freeze = threeWayScalar<FreezePane | undefined>(
    base?.freeze,
    ours?.freeze,
    theirs?.freeze,
  );
  const layout: Partial<Sheet> = {
    columnWidths: threeWayMap(base?.columnWidths, ours?.columnWidths, theirs?.columnWidths),
    rowHeights: threeWayMap(base?.rowHeights, ours?.rowHeights, theirs?.rowHeights),
    hiddenColumns: threeWaySet(base?.hiddenColumns, ours?.hiddenColumns, theirs?.hiddenColumns),
    hiddenRows: threeWaySet(base?.hiddenRows, ours?.hiddenRows, theirs?.hiddenRows),
    merges: threeWaySet(base?.merges, ours?.merges, theirs?.merges),
    hidden: threeWayScalar(base?.hidden, ours?.hidden, theirs?.hidden),
    veryHidden: threeWayScalar(base?.veryHidden, ours?.veryHidden, theirs?.veryHidden),
    tabColor: threeWayScalar(base?.tabColor, ours?.tabColor, theirs?.tabColor),
    freeze,
  };
  return layout;
}

function pick(
  side: Side,
  base: Cell | null | undefined,
  ours: Cell | null | undefined,
  theirs: Cell | null | undefined,
): Cell | null {
  if (side === 'ours') return cloneCell(ours);
  if (side === 'theirs') return cloneCell(theirs);
  if (side === 'base') return cloneCell(base);
  if (side === 'delete' || side === 'none') return null;
  return null;
}

/**
 * Three-way merge. Sheets: only on one side → take; both → cell-wise.
 * Rename treated as delete+add (V0 limitation).
 */
export function mergeSnapshots(
  base: WorkbookSnapshot,
  ours: WorkbookSnapshot,
  theirs: WorkbookSnapshot,
): MergeResult {
  const conflicts: CellConflict[] = [];
  let autoChangeCount = 0;

  const baseSheets = new Set(base.sheetOrder);
  const oursSheets = new Set(ours.sheetOrder);
  const theirsSheets = new Set(theirs.sheetOrder);

  // Sheet presence: union with order preferring ours order then theirs additions
  const resultOrder: string[] = [];
  const seen = new Set<string>();

  // Sheets only on theirs (added on scenario) — take
  // Sheets only on ours — take
  // Sheets on both — merge cell-wise
  // Sheets only on base (deleted both?) — drop
  // Sheets on base+ours not theirs — keep ours (deleted on theirs only if cells handled)
  // V0: sheet only on one side of ours/theirs → take that side

  for (const name of ours.sheetOrder) {
    if (!seen.has(name)) {
      // If sheet only on theirs? handled below
      if (oursSheets.has(name) || baseSheets.has(name)) {
        resultOrder.push(name);
        seen.add(name);
      }
    }
  }
  for (const name of theirs.sheetOrder) {
    if (!seen.has(name) && !baseSheets.has(name)) {
      // Added only on theirs
      resultOrder.push(name);
      seen.add(name);
    } else if (!seen.has(name) && oursSheets.has(name)) {
      resultOrder.push(name);
      seen.add(name);
    }
  }

  // Also include sheets that exist only on theirs already done; only on ours done.
  // Sheet removed on ours but present on theirs+base: still in theirs order — add if not seen
  for (const name of theirs.sheetOrder) {
    if (!seen.has(name) && baseSheets.has(name) && !oursSheets.has(name)) {
      // Deleted on ours entirely — if theirs still has it equal to base, delete sheet;
      // if theirs edited, take theirs sheet (simplified: take theirs if present)
      resultOrder.push(name);
      seen.add(name);
    }
  }

  const resultSheets: WorkbookSnapshot['sheets'] = {};

  const allSheetNames = new Set([
    ...base.sheetOrder,
    ...ours.sheetOrder,
    ...theirs.sheetOrder,
  ]);

  for (const sheetName of allSheetNames) {
    const inBase = baseSheets.has(sheetName);
    const inOurs = oursSheets.has(sheetName);
    const inTheirs = theirsSheets.has(sheetName);

    // Sheet only on ours (not base, not theirs) → take ours
    if (!inBase && inOurs && !inTheirs) {
      resultSheets[sheetName] = JSON.parse(JSON.stringify(ours.sheets[sheetName]));
      if (!seen.has(sheetName)) {
        resultOrder.push(sheetName);
        seen.add(sheetName);
      }
      autoChangeCount += 1;
      continue;
    }
    // Sheet only on theirs → take theirs
    if (!inBase && !inOurs && inTheirs) {
      resultSheets[sheetName] = JSON.parse(JSON.stringify(theirs.sheets[sheetName]));
      if (!seen.has(sheetName)) {
        resultOrder.push(sheetName);
        seen.add(sheetName);
      }
      autoChangeCount += 1;
      continue;
    }
    // Sheet deleted on both sides (in base only)
    if (inBase && !inOurs && !inTheirs) {
      autoChangeCount += 1;
      continue;
    }
    // Sheet deleted on ours only, unchanged on theirs → delete
    if (inBase && !inOurs && inTheirs) {
      const bSheet = base.sheets[sheetName];
      const tSheet = theirs.sheets[sheetName];
      if (JSON.stringify(bSheet) === JSON.stringify(tSheet)) {
        autoChangeCount += 1;
        continue; // deleted
      }
      // Theirs edited whole sheet while ours deleted → treat as take theirs for V0
      // (cell-wise would be heavy); still cell-merge with empty ours
    }
    // Sheet deleted on theirs only, unchanged on ours → delete
    if (inBase && inOurs && !inTheirs) {
      const bSheet = base.sheets[sheetName];
      const oSheet = ours.sheets[sheetName];
      if (JSON.stringify(bSheet) === JSON.stringify(oSheet)) {
        autoChangeCount += 1;
        continue;
      }
    }

    // Both sides (or one side with edits) → cell-wise
    const bCells = base.sheets[sheetName]?.cells ?? {};
    const oCells = ours.sheets[sheetName]?.cells ?? {};
    const tCells = theirs.sheets[sheetName]?.cells ?? {};
    const addrs = new Set([...Object.keys(bCells), ...Object.keys(oCells), ...Object.keys(tCells)]);
    const outCells: Record<string, Cell> = {};
    let maxRow = 0;
    let maxCol = 0;

    for (const addr of addrs) {
      const b = bCells[addr];
      const o = inOurs ? oCells[addr] : undefined;
      const t = inTheirs ? tCells[addr] : undefined;
      // When sheet missing on a side, treat all cells as missing on that side
      const oCell = inOurs ? (o ?? null) : null;
      const tCell = inTheirs ? (t ?? null) : null;
      const bCell = inBase ? (b ?? null) : null;

      const { side, reason } = resolveCell(bCell, oCell, tCell);
      if (side === 'conflict') {
        conflicts.push({
          sheet: sheetName,
          address: addr,
          base: cloneCell(bCell),
          ours: cloneCell(oCell),
          theirs: cloneCell(tCell),
          reason: reason ?? 'value',
        });
        // Keep base (or ours) as placeholder until resolved
        const placeholder = cloneCell(oCell) ?? cloneCell(bCell);
        if (placeholder) outCells[addr] = placeholder;
        continue;
      }
      const chosen = pick(side, bCell, oCell, tCell);
      if (side !== 'base' && side !== 'none') {
        // Count auto change when result differs from base
        if (!cellsEqual(chosen, bCell)) autoChangeCount += 1;
      }
      if (chosen) outCells[addr] = chosen;
    }

    // dimensions: take max of sides present
    const dims = [
      inOurs ? ours.sheets[sheetName]?.dimensions : null,
      inTheirs ? theirs.sheets[sheetName]?.dimensions : null,
      inBase ? base.sheets[sheetName]?.dimensions : null,
    ].filter(Boolean) as { rows: number; cols: number }[];
    maxRow = Math.max(1, ...dims.map((d) => d.rows));
    maxCol = Math.max(1, ...dims.map((d) => d.cols));

    if (inOurs || inTheirs) {
      const layout = mergeSheetLayout(
        base.sheets[sheetName],
        ours.sheets[sheetName],
        theirs.sheets[sheetName],
      );
      resultSheets[sheetName] = {
        dimensions: { rows: maxRow, cols: maxCol },
        cells: outCells,
        ...layout,
      };
      if (!seen.has(sheetName)) {
        resultOrder.push(sheetName);
        seen.add(sheetName);
      }
    }
  }

  // Ensure order only includes sheets we kept
  const finalOrder = resultOrder.filter((n) => n in resultSheets);

  return {
    snapshot: { sheetOrder: finalOrder, sheets: resultSheets },
    conflicts,
    autoChangeCount,
  };
}

/**
 * Apply conflict resolutions to a merge result. Pure.
 */
export function applyResolutions(
  mergeResult: MergeResult,
  resolutions: Record<string, ResolutionChoice>,
): MergeResult {
  const snapshot: WorkbookSnapshot = JSON.parse(JSON.stringify(mergeResult.snapshot));
  const remaining: CellConflict[] = [];

  for (const conflict of mergeResult.conflicts) {
    const key = `${conflict.sheet}!${conflict.address}`;
    const res = resolutions[key];
    if (!res) {
      remaining.push(conflict);
      continue;
    }
    const sheet = snapshot.sheets[conflict.sheet];
    if (!sheet) {
      snapshot.sheets[conflict.sheet] = {
        dimensions: { rows: 1, cols: 1 },
        cells: {},
      };
      if (!snapshot.sheetOrder.includes(conflict.sheet)) {
        snapshot.sheetOrder.push(conflict.sheet);
      }
    }
    const cells = snapshot.sheets[conflict.sheet].cells;
    if (res.action === 'keep-ours') {
      if (conflict.ours) cells[conflict.address] = cloneCell(conflict.ours)!;
      else delete cells[conflict.address];
    } else if (res.action === 'keep-theirs') {
      if (conflict.theirs) cells[conflict.address] = cloneCell(conflict.theirs)!;
      else delete cells[conflict.address];
    } else if (res.action === 'edit') {
      cells[conflict.address] = cloneCell(res.cell)!;
    }
  }

  return {
    snapshot,
    conflicts: remaining,
    autoChangeCount: mergeResult.autoChangeCount,
  };
}
