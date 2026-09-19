import ExcelJS from 'exceljs';
import { readFile } from 'node:fs/promises';
import {
  DEFAULT_COL_WIDTHS,
  DEFAULT_FONT_NAME,
  DEFAULT_FONT_SIZE,
  DEFAULT_ROW_HEIGHT,
  colToLetter,
  excelColorToHex,
  isMergeSlave,
  type ExcelColorLike,
} from './format.js';
import type {
  Cell,
  CellAlignment,
  CellBorders,
  CellFormat,
  FreezePane,
  Sheet,
  WorkbookSnapshot,
} from './types.js';

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
    if ('result' in value) {
      const result = (value as { result?: ExcelJS.CellValue }).result;
      return cellValueToV(result ?? null);
    }
    if ('richText' in value && Array.isArray((value as ExcelJS.CellRichTextValue).richText)) {
      return (value as ExcelJS.CellRichTextValue).richText.map((p) => p.text).join('');
    }
    if ('text' in value && typeof (value as ExcelJS.CellHyperlinkValue).text === 'string') {
      return (value as ExcelJS.CellHyperlinkValue).text;
    }
    if ('error' in value) {
      return String((value as ExcelJS.CellErrorValue).error);
    }
  }
  return null;
}

function asColor(c: unknown): ExcelColorLike | undefined {
  if (!c || typeof c !== 'object') return undefined;
  return c as ExcelColorLike;
}

function extractFill(cell: ExcelJS.Cell): string | undefined {
  const fill = cell.fill;
  if (!fill || fill.type !== 'pattern') return undefined;
  const pattern = fill as ExcelJS.FillPattern;
  if (pattern.pattern === 'none') return undefined;
  return excelColorToHex(asColor(pattern.fgColor));
}

function extractFontColor(cell: ExcelJS.Cell): string | undefined {
  return excelColorToHex(asColor(cell.font?.color));
}

function extractBorders(cell: ExcelJS.Cell): CellBorders | undefined {
  const b = cell.border;
  if (!b) return undefined;
  const edge = (
    side: ExcelJS.Border | undefined,
  ): { style?: string; color?: string } | undefined => {
    if (!side || (!side.style && !side.color)) return undefined;
    const out: { style?: string; color?: string } = {};
    if (side.style) out.style = String(side.style);
    const color = excelColorToHex(asColor(side.color));
    if (color) out.color = color;
    return out;
  };
  const out: CellBorders = {};
  const top = edge(b.top);
  const left = edge(b.left);
  const bottom = edge(b.bottom);
  const right = edge(b.right);
  if (top) out.top = top;
  if (left) out.left = left;
  if (bottom) out.bottom = bottom;
  if (right) out.right = right;
  return Object.keys(out).length ? out : undefined;
}

function extractAlignment(cell: ExcelJS.Cell): CellAlignment | undefined {
  const a = cell.alignment;
  if (!a) return undefined;
  const out: CellAlignment = {};
  if (a.horizontal) out.horizontal = String(a.horizontal);
  if (a.vertical) out.vertical = String(a.vertical);
  if (a.wrapText) out.wrapText = true;
  if (a.indent) out.indent = a.indent;
  if (a.textRotation) out.textRotation = a.textRotation;
  return Object.keys(out).length ? out : undefined;
}

function extractFormat(excelCell: ExcelJS.Cell): CellFormat | undefined {
  const fmt: CellFormat = {};
  const numFmt = excelCell.numFmt;
  if (numFmt && numFmt !== 'General') {
    fmt.numFmt = numFmt;
  }
  const font = excelCell.font;
  if (font?.bold) fmt.bold = true;
  if (font?.italic) fmt.italic = true;
  if (font?.strike) fmt.strike = true;
  if (font?.underline && font.underline !== false) {
    fmt.underline = font.underline === true ? true : String(font.underline);
  }
  if (font?.name && font.name !== DEFAULT_FONT_NAME) {
    fmt.fontName = font.name;
  }
  if (font?.size != null && font.size !== DEFAULT_FONT_SIZE) {
    fmt.fontSize = font.size;
  }
  const fill = extractFill(excelCell);
  if (fill) fmt.fill = fill;
  const fontColor = extractFontColor(excelCell);
  if (fontColor && fontColor.replace(/^#/, '').toUpperCase() !== '000000') {
    fmt.fontColor = fontColor;
  }
  const borders = extractBorders(excelCell);
  if (borders) fmt.borders = borders;
  const alignment = extractAlignment(excelCell);
  if (alignment) fmt.alignment = alignment;

  if (Object.keys(fmt).length === 0) return undefined;
  return fmt;
}

function extractFreeze(ws: ExcelJS.Worksheet): FreezePane | undefined {
  const views = ws.views;
  if (!views || views.length === 0) return undefined;
  const view = views[0];
  if (view.state !== 'frozen' && view.state !== 'split') return undefined;
  const row = view.ySplit ?? 0;
  const col = view.xSplit ?? 0;
  if (!row && !col) return undefined;
  return { row, col };
}

function extractMerges(ws: ExcelJS.Worksheet): string[] | undefined {
  const model = ws.model as { merges?: string[] } | undefined;
  const list = model?.merges;
  if (!Array.isArray(list) || list.length === 0) return undefined;
  return [...list].sort();
}

function parseWorksheet(ws: ExcelJS.Worksheet): Sheet {
  const cells: Record<string, Cell> = {};
  let maxRow = 0;
  let maxCol = 0;
  const columnWidths: Record<string, number> = {};
  const rowHeights: Record<string, number> = {};
  const hiddenColumns: string[] = [];
  const hiddenRows: number[] = [];
  const merges = extractMerges(ws);

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    maxRow = Math.max(maxRow, rowNumber);
    if (row.height != null && row.height !== DEFAULT_ROW_HEIGHT) {
      rowHeights[String(rowNumber)] = row.height;
    }
    if (row.hidden) hiddenRows.push(rowNumber);

    row.eachCell({ includeEmpty: true }, (excelCell, colNumber) => {
      maxCol = Math.max(maxCol, colNumber);

      const address = `${colToLetter(colNumber)}${rowNumber}`;
      let formula: string | null = null;
      let v: string | number | boolean | null = null;

      const raw = excelCell.value;
      if (raw && typeof raw === 'object' && 'formula' in raw) {
        formula = normalizeFormula((raw as ExcelJS.CellFormulaValue).formula);
        v = cellValueToV(raw);
      } else if (raw && typeof raw === 'object' && 'sharedFormula' in raw) {
        const shared = raw as ExcelJS.CellSharedFormulaValue;
        formula = normalizeFormula(shared.formula ?? shared.sharedFormula);
        v = cellValueToV(raw);
      } else {
        v = cellValueToV(raw);
      }

      if (!formula && excelCell.formula) {
        formula = normalizeFormula(excelCell.formula);
      }

      const fmt = extractFormat(excelCell);
      let hyperlink: string | undefined;
      if (raw && typeof raw === 'object' && 'hyperlink' in raw) {
        const href = (raw as ExcelJS.CellHyperlinkValue).hyperlink;
        if (typeof href === 'string' && href.trim()) hyperlink = href.trim();
      } else if (excelCell.hyperlink) {
        const href = String(excelCell.hyperlink);
        if (href.trim()) hyperlink = href.trim();
      }
      if (isMergeSlave(address, merges)) return;
      if (v === null && formula === null && !fmt && !hyperlink) return;

      const cell: Cell = { v, f: formula };
      if (fmt) cell.fmt = fmt;
      if (hyperlink) cell.hyperlink = hyperlink;
      cells[address] = cell;
    });
  });

  const cols = ws.columns ?? [];
  for (const col of cols) {
    if (!col || col.number == null) continue;
    const letter = colToLetter(col.number);
    maxCol = Math.max(maxCol, col.number);
    if (col.width != null && !DEFAULT_COL_WIDTHS.has(col.width)) {
      columnWidths[letter] = col.width;
    }
    if (col.hidden) hiddenColumns.push(letter);
  }

  const dim = ws.dimensions;
  if (dim) {
    maxRow = Math.max(maxRow, dim.bottom || 0);
    maxCol = Math.max(maxCol, dim.right || 0);
  }

  const sheet: Sheet = {
    dimensions: {
      rows: maxRow || 1,
      cols: maxCol || 1,
    },
    cells,
  };

  if (Object.keys(columnWidths).length) sheet.columnWidths = columnWidths;
  if (Object.keys(rowHeights).length) sheet.rowHeights = rowHeights;
  if (hiddenColumns.length) sheet.hiddenColumns = hiddenColumns.sort();
  if (hiddenRows.length) sheet.hiddenRows = hiddenRows.sort((a, b) => a - b);

  if (merges) sheet.merges = merges;

  if (ws.state === 'hidden') sheet.hidden = true;
  if (ws.state === 'veryHidden') sheet.veryHidden = true;

  const tab = excelColorToHex(asColor(ws.properties?.tabColor));
  if (tab) sheet.tabColor = tab;

  const freeze = extractFreeze(ws);
  if (freeze) sheet.freeze = freeze;

  return sheet;
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
