/**
 * Canonical Bit workbook snapshot types (M1).
 * Tracked: values, formulas, sheet structure, basic formatting.
 */

export interface CellFormat {
  /** Excel number format string, e.g. "0.00", "$#,##0" */
  numFmt?: string;
  bold?: boolean;
  italic?: boolean;
  /** Fill colour as #RRGGBB (ARGB stripped to RGB when present) */
  fill?: string;
  /** Font colour as #RRGGBB */
  fontColor?: string;
}

export interface Cell {
  /** Cached / typed value. null when formula-only or empty. */
  v: string | number | boolean | null;
  /** Formula string including leading "=", or null if none. */
  f: string | null;
  /** Basic formatting subset; omit or empty when none tracked. */
  fmt?: CellFormat;
}

export interface SheetDimensions {
  rows: number;
  cols: number;
}

export interface Sheet {
  dimensions: SheetDimensions;
  cells: Record<string, Cell>;
}

export interface WorkbookSnapshot {
  sheets: Record<string, Sheet>;
  /** Sheet names in workbook order */
  sheetOrder: string[];
}
