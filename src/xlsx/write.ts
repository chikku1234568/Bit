import ExcelJS from 'exceljs';
import type { Cell, CellFormat, WorkbookSnapshot } from './types.js';

function hexToArgb(hex: string): string {
  const cleaned = hex.replace(/^#/, '').toUpperCase();
  if (cleaned.length === 6) return `FF${cleaned}`;
  if (cleaned.length === 8) return cleaned;
  return `FF${cleaned.padStart(6, '0').slice(0, 6)}`;
}

function applyFormat(excelCell: ExcelJS.Cell, fmt: CellFormat): void {
  if (fmt.numFmt) {
    excelCell.numFmt = fmt.numFmt;
  }
  if (fmt.bold || fmt.italic || fmt.fontColor) {
    excelCell.font = {
      ...(excelCell.font || {}),
      ...(fmt.bold ? { bold: true } : {}),
      ...(fmt.italic ? { italic: true } : {}),
      ...(fmt.fontColor
        ? { color: { argb: hexToArgb(fmt.fontColor) } }
        : {}),
    };
  }
  if (fmt.fill) {
    excelCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: hexToArgb(fmt.fill) },
    };
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
  } else if (cell.v !== null && cell.v !== undefined) {
    excelCell.value = cell.v;
  }

  if (cell.fmt) {
    applyFormat(excelCell, cell.fmt);
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
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
