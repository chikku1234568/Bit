import ExcelJS from 'exceljs';
import { colToLetter, hexToArgb } from './format.js';
import type {
  BorderEdge,
  Cell,
  CellAlignment,
  CellBorders,
  CellFormat,
  Sheet,
  WorkbookSnapshot,
} from './types.js';

function applyBorderEdge(edge: BorderEdge | undefined): ExcelJS.Border | undefined {
  if (!edge || (!edge.style && !edge.color)) return undefined;
  const out: ExcelJS.Border = {};
  if (edge.style) out.style = edge.style as ExcelJS.BorderStyle;
  if (edge.color) out.color = { argb: hexToArgb(edge.color) };
  return out;
}

function applyFormat(excelCell: ExcelJS.Cell, fmt: CellFormat): void {
  if (fmt.numFmt) {
    excelCell.numFmt = fmt.numFmt;
  }

  const font: Partial<ExcelJS.Font> = { ...(excelCell.font || {}) };
  let fontDirty = false;
  if (fmt.bold) {
    font.bold = true;
    fontDirty = true;
  }
  if (fmt.italic) {
    font.italic = true;
    fontDirty = true;
  }
  if (fmt.strike) {
    font.strike = true;
    fontDirty = true;
  }
  if (fmt.underline) {
    font.underline = fmt.underline as ExcelJS.Font['underline'];
    fontDirty = true;
  }
  if (fmt.fontName) {
    font.name = fmt.fontName;
    fontDirty = true;
  }
  if (fmt.fontSize != null) {
    font.size = fmt.fontSize;
    fontDirty = true;
  }
  if (fmt.fontColor) {
    font.color = { argb: hexToArgb(fmt.fontColor) };
    fontDirty = true;
  }
  if (fontDirty) excelCell.font = font as ExcelJS.Font;

  if (fmt.fill) {
    excelCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: hexToArgb(fmt.fill) },
    };
  }

  if (fmt.borders) {
    const b: CellBorders = fmt.borders;
    excelCell.border = {
      top: applyBorderEdge(b.top),
      left: applyBorderEdge(b.left),
      bottom: applyBorderEdge(b.bottom),
      right: applyBorderEdge(b.right),
    };
  }

  if (fmt.alignment) {
    const a: CellAlignment = fmt.alignment;
    const align: Partial<ExcelJS.Alignment> = {};
    if (a.horizontal) align.horizontal = a.horizontal as ExcelJS.Alignment['horizontal'];
    if (a.vertical) align.vertical = a.vertical as ExcelJS.Alignment['vertical'];
    if (a.wrapText) align.wrapText = true;
    if (a.indent) align.indent = a.indent;
    if (a.textRotation) align.textRotation = a.textRotation;
    excelCell.alignment = align;
  }
}

function setCell(excelCell: ExcelJS.Cell, cell: Cell): void {
  if (cell.f) {
    const formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
    if (cell.v !== null && cell.v !== undefined) {
      excelCell.value = {
        formula,
        result: cell.v as string | number | boolean,
      };
    } else {
      excelCell.value = { formula };
    }
  } else if (cell.hyperlink) {
    excelCell.value = {
      text: cell.v != null ? String(cell.v) : cell.hyperlink,
      hyperlink: cell.hyperlink,
    };
  } else if (cell.v !== null && cell.v !== undefined) {
    excelCell.value = cell.v;
  }

  if (cell.fmt) {
    applyFormat(excelCell, cell.fmt);
  }
}

function applySheetLayout(ws: ExcelJS.Worksheet, sheet: Sheet): void {
  if (sheet.hidden) ws.state = 'hidden';
  else if (sheet.veryHidden) ws.state = 'veryHidden';

  if (sheet.tabColor) {
    ws.properties.tabColor = { argb: hexToArgb(sheet.tabColor) };
  }

  if (sheet.columnWidths) {
    for (const [letter, width] of Object.entries(sheet.columnWidths)) {
      ws.getColumn(letter).width = width;
    }
  }
  if (sheet.hiddenColumns) {
    for (const letter of sheet.hiddenColumns) {
      ws.getColumn(letter).hidden = true;
    }
  }

  if (sheet.rowHeights) {
    for (const [row, height] of Object.entries(sheet.rowHeights)) {
      ws.getRow(Number(row)).height = height;
    }
  }
  if (sheet.hiddenRows) {
    for (const row of sheet.hiddenRows) {
      const r = ws.getRow(row);
      r.hidden = true;
      if (r.height == null) r.height = 15;
    }
  }

  if (sheet.merges) {
    for (const range of sheet.merges) {
      try {
        ws.mergeCells(range);
      } catch {
        // Overlapping merge on reconstruct — skip
      }
    }
  }

  if (sheet.freeze && (sheet.freeze.row > 0 || sheet.freeze.col > 0)) {
    const topLeft = `${colToLetter(Math.max(1, sheet.freeze.col + 1))}${Math.max(1, sheet.freeze.row + 1)}`;
    ws.views = [
      {
        state: 'frozen',
        xSplit: sheet.freeze.col,
        ySplit: sheet.freeze.row,
        topLeftCell: topLeft,
        activeCell: 'A1',
      },
    ];
  }
}

/**
 * Write a WorkbookSnapshot to an .xlsx Buffer.
 */
export async function writeXlsx(snapshot: WorkbookSnapshot): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bit';
  workbook.created = new Date();

  const order =
    snapshot.sheetOrder.length > 0
      ? snapshot.sheetOrder
      : Object.keys(snapshot.sheets);

  for (const name of order) {
    const sheet = snapshot.sheets[name];
    if (!sheet) continue;
    const ws = workbook.addWorksheet(name);

    for (const [address, cell] of Object.entries(sheet.cells)) {
      const excelCell = ws.getCell(address);
      setCell(excelCell, cell);
    }

    applySheetLayout(ws, sheet);
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
