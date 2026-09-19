/**
 * Canonical Bit workbook snapshot types.
 * Tracked: values, formulas, sheet structure, cell formatting, layout,
 * hyperlinks, validation, named ranges, comments, tables / autofilter.
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

export interface CellComment {
  text: string;
  /** ExcelJS does not round-trip authors reliably. */
  author?: string;
}

export interface Cell {
  /** Cached / typed value. null when formula-only or empty. */
  v: string | number | boolean | null;
  /** Formula string including leading "=", or null if none. */
  f: string | null;
  /** Tracked formatting; omit when none. */
  fmt?: CellFormat;
  /** Hyperlink target URL (external or internal), when present. */
  hyperlink?: string;
  /** Cell note / comment. */
  comment?: CellComment;
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

/** Data validation rule (ExcelJS expands multi-cell sqref to per-cell on read). */
export interface ValidationRule {
  sqref: string;
  type: string;
  operator?: string;
  formulae?: string[];
  allowBlank?: boolean;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}

export interface SheetTable {
  name: string;
  ref: string;
  headerRow?: boolean;
  totalsRow?: boolean;
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
  /** Data validation rules on this sheet. */
  validations?: ValidationRule[];
  /** Excel tables (ListObjects). */
  tables?: SheetTable[];
  /** Sheet autoFilter range, e.g. "A1:D10". */
  autoFilter?: string | null;
}

/** Workbook- or sheet-scoped defined name. */
export interface NamedRange {
  name: string;
  refersTo: string;
  /** Sheet name when sheet-scoped; null/omit = workbook. */
  scope?: string | null;
}

export interface WorkbookSnapshot {
  sheets: Record<string, Sheet>;
  /** Sheet names in workbook order */
  sheetOrder: string[];
  /** Defined names (named ranges). */
  names?: NamedRange[];
}
