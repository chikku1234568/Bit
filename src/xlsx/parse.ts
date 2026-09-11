import ExcelJS from 'exceljs';
import { readFile } from 'node:fs/promises';
import type { Cell, CellFormat, Sheet, WorkbookSnapshot } from './types.js';

function argbToHex(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  const cleaned = argb.replace(/^#/, '').toUpperCase();
  // ExcelJS often uses AARRGGBB
  if (cleaned.length === 8) {
    return `#${cleaned.slice(2)}`;
  }
  if (cleaned.length === 6) {
    return `#${cleaned}`;
  }
  return undefined;
}

function extractFill(cell: ExcelJS.Cell): string | undefined {
  const fill = cell.fill;
  if (!fill || fill.type !== 'pattern') return undefined;
  const pattern = fill as ExcelJS.FillPattern;
  if (pattern.pattern === 'none') return undefined;
  const fg = pattern.fgColor;
  if (!fg) return undefined;
  if ('argb' in fg && typeof fg.argb === 'string') {
    return argbToHex(fg.argb);
  }
  return undefined;
}

function extractFontColor(cell: ExcelJS.Cell): string | undefined {
  const color = cell.font?.color;
  if (!color) return undefined;
  if ('argb' in color && typeof color.argb === 'string') {
    return argbToHex(color.argb);
  }
  return undefined;
}

function normalizeFormula(formula: string | undefined): string | null {
  if (!formula) return null;
  const trimmed = formula.trim();
  if (!trimmed) return null;
  return trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
}

function cellValueToV(
  value: ExcelJS.CellValue,
): string | number | boolean | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === 'object') {
    // Formula result object: { formula, result?, shareType?, ... }
    if ('result' in value) {
      const result = (value as { result?: ExcelJS.CellValue }).result;
      return cellValueToV(result ?? null);
    }
    // Rich text
    if ('richText' in value && Array.isArray((value as ExcelJS.CellRichTextValue).richText)) {
      return (value as ExcelJS.CellRichTextValue).richText.map((p) => p.text).join('');
    }
    // Hyperlink
    if ('text' in value && typeof (value as ExcelJS.CellHyperlinkValue).text === 'string') {
      return (value as ExcelJS.CellHyperlinkValue).text;
    }
    // Error
    if ('error' in value) {
      return String((value as ExcelJS.CellErrorValue).error);
    }
  }
  return null;
}

function extractFormat(excelCell: ExcelJS.Cell): CellFormat | undefined {
  const fmt: CellFormat = {};
  const numFmt = excelCell.numFmt;
  if (numFmt && numFmt !== 'General') {
    fmt.numFmt = numFmt;
  }
  if (excelCell.font?.bold) fmt.bold = true;
  if (excelCell.font?.italic) fmt.italic = true;
  const fill = extractFill(excelCell);
  if (fill) fmt.fill = fill;
  const fontColor = extractFontColor(excelCell);
  if (fontColor) fmt.fontColor = fontColor;

  if (Object.keys(fmt).length === 0) return undefined;
  return fmt;
}

function colToLetter(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseWorksheet(ws: ExcelJS.Worksheet): Sheet {
  const cells: Record<string, Cell> = {};
  let maxRow = 0;
  let maxCol = 0;

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (excelCell, colNumber) => {
      maxRow = Math.max(maxRow, rowNumber);
      maxCol = Math.max(maxCol, colNumber);

      const address = `${colToLetter(colNumber)}${rowNumber}`;
      let formula: string | null = null;
      let v: string | number | boolean | null = null;

      const raw = excelCell.value;
      if (raw && typeof raw === 'object' && 'formula' in raw) {
        formula = normalizeFormula((raw as ExcelJS.CellFormulaValue).formula);
        v = cellValueToV(raw);
      } else if (raw && typeof raw === 'object' && 'sharedFormula' in raw) {
        // Prefer master formula if present; sharedFormula alone is the ref
        const shared = raw as ExcelJS.CellSharedFormulaValue;
        formula = normalizeFormula(shared.formula ?? shared.sharedFormula);
        v = cellValueToV(raw);
      } else {
        v = cellValueToV(raw);
      }

      // Also check excelCell.formula getter when value didn't expose it
      if (!formula && excelCell.formula) {
        formula = normalizeFormula(excelCell.formula);
      }

      const fmt = extractFormat(excelCell);

      // Skip truly empty cells (no value, no formula, no tracked fmt)
      if (v === null && formula === null && !fmt) return;

      const cell: Cell = { v, f: formula };
      if (fmt) cell.fmt = fmt;
      cells[address] = cell;
    });
  });

  // Prefer ExcelJS reported dimensions when larger
  const dim = ws.dimensions;
  if (dim) {
    maxRow = Math.max(maxRow, dim.bottom || 0);
    maxCol = Math.max(maxCol, dim.right || 0);
  }

  return {
    dimensions: {
      rows: maxRow || 1,
      cols: maxCol || 1,
    },
    cells,
  };
}

/**
 * Parse an .xlsx file from a Buffer or filesystem path into a WorkbookSnapshot.
 */
export async function parseXlsx(
  input: Buffer | string,
): Promise<WorkbookSnapshot> {
  const workbook = new ExcelJS.Workbook();

  if (typeof input === 'string') {
    await workbook.xlsx.readFile(input);
  } else {
    // ExcelJS accepts Buffer via stream-like load
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await workbook.xlsx.load(input as any);
  }

  const sheets: Record<string, Sheet> = {};
  const sheetOrder: string[] = [];

  workbook.eachSheet((ws) => {
    sheetOrder.push(ws.name);
    sheets[ws.name] = parseWorksheet(ws);
  });

  return { sheets, sheetOrder };
}

/** Convenience: read path to buffer then parse (same as parseXlsx(path)). */
export async function parseXlsxFromPath(path: string): Promise<WorkbookSnapshot> {
  const buf = await readFile(path);
  return parseXlsx(buf);
}
