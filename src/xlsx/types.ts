/**
 * Canonical Bit workbook snapshot types.
 * Tracked: values, formulas, sheet structure, cell formatting, layout.
 */

export interface BorderEdge {
  style?: string;
  /** #RRGGBB */
  color?: string;
}

export interface CellBorders {
  top?: BorderEdge;
  left?: BorderEdge;
  bottom?: BorderEdge;
  right?: BorderEdge;
}

export interface CellAlignment {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  indent?: number;
  textRotation?: number;
}

export interface CellFormat {
  /** Excel number format string, e.g. "0.00", "$#,##0" */
  numFmt?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strike?: boolean;
  fontName?: string;
  fontSize?: number;
  /** Fill colour as #RRGGBB (theme/indexed resolved when possible) */
  fill?: string;
  /** Font colour as #RRGGBB */
  fontColor?: string;
  borders?: CellBorders;
  alignment?: CellAlignment;
}

export interface Cell {
  /** Cached / typed value. null when formula-only or empty. */
  v: string | number | boolean | null;
  /** Formula string including leading "=", or null if none. */
  f: string | null;
  /** Tracked formatting; omit when none. */
  fmt?: CellFormat;
}

export interface SheetDimensions {
  rows: number;
  cols: number;
}

export interface FreezePane {
  /** Number of frozen rows (ySplit). */
  row: number;
  /** Number of frozen columns (xSplit). */
  col: number;
}

export interface Sheet {
  dimensions: SheetDimensions;
  cells: Record<string, Cell>;
  /** Column letter → width in Excel character units. Sparse; defaults omitted. */
  columnWidths?: Record<string, number>;
  /** Row number (string key) → height in points. Sparse; defaults omitted. */
  rowHeights?: Record<string, number>;
  hiddenColumns?: string[];
  hiddenRows?: number[];
  /** Merge ranges, e.g. "A1:D1". */
  merges?: string[];
  hidden?: boolean;
  veryHidden?: boolean;
  tabColor?: string;
  freeze?: FreezePane;
}

export interface WorkbookSnapshot {
  sheets: Record<string, Sheet>;
  /** Sheet names in workbook order */
  sheetOrder: string[];
}
