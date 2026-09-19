import ExcelJS from 'exceljs';
import { colToLetter, parseAddress, hexToArgb } from './format.js';
import type {
  BorderEdge,
  Cell,
  CellAlignment,
  CellBorders,
  CellFormat,
  NamedRange,
  Sheet,
  SheetTable,
  ValidationRule,
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

  if (cell.comment?.text) {
    excelCell.note = cell.comment.text;
  }
}

function applyValidations(ws: ExcelJS.Worksheet, rules: ValidationRule[]): void {
  const store = (
    ws as ExcelJS.Worksheet & {
      dataValidations: { add: (address: string, v: ExcelJS.DataValidation) => void };
    }
  ).dataValidations;
  for (const rule of rules) {
    if (!rule.sqref || !rule.type) continue;
    const dv: ExcelJS.DataValidation = {
      type: rule.type as ExcelJS.DataValidation['type'],
      formulae: (rule.formulae ?? []).map((f) => {
        // Restore numeric formulae for whole/decimal when possible
        if (rule.type === 'whole' || rule.type === 'textLength') {
          const n = Number(f);
          if (Number.isFinite(n) && String(n) === f.trim()) return n;
        }
        if (rule.type === 'decimal') {
          const n = Number(f);
          if (Number.isFinite(n)) return n;
        }
        return f;
      }),
    };
    if (rule.operator) dv.operator = rule.operator as ExcelJS.DataValidationOperator;
    if (rule.allowBlank) dv.allowBlank = true;
    if (rule.showErrorMessage) dv.showErrorMessage = true;
    if (rule.showInputMessage) dv.showInputMessage = true;
    if (rule.errorTitle) dv.errorTitle = rule.errorTitle;
    if (rule.error) dv.error = rule.error;
    if (rule.promptTitle) dv.promptTitle = rule.promptTitle;
    if (rule.prompt) dv.prompt = rule.prompt;
    store.add(rule.sqref, dv);
  }
}

function parseRangeBounds(ref: string): {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
} | null {
  const parts = ref.split(':');
  const start = parseAddress(parts[0]);
  if (!start) return null;
  const end = parts[1] ? parseAddress(parts[1]) : start;
  if (!end) return null;
  return {
    c1: Math.min(start.col, end.col),
    r1: Math.min(start.row, end.row),
    c2: Math.max(start.col, end.col),
    r2: Math.max(start.row, end.row),
  };
}

function buildTablePayload(
  ws: ExcelJS.Worksheet,
  t: SheetTable,
): ExcelJS.TableProperties | null {
  const bounds = parseRangeBounds(t.ref);
  if (!bounds) return null;
  const headerRow = t.headerRow !== false;
  const totalsRow = !!t.totalsRow;
  const colCount = bounds.c2 - bounds.c1 + 1;
  const columns: ExcelJS.TableColumnProperties[] = [];
  for (let i = 0; i < colCount; i++) {
    const col = bounds.c1 + i;
    let name = `Column${i + 1}`;
    if (headerRow) {
      const v = ws.getCell(bounds.r1, col).value;
      if (v != null && v !== '') name = String(typeof v === 'object' ? (v as { text?: string }).text ?? v : v);
    }
    columns.push({ name });
  }
  const dataStart = headerRow ? bounds.r1 + 1 : bounds.r1;
  const dataEnd = totalsRow ? bounds.r2 - 1 : bounds.r2;
  const rows: unknown[][] = [];
  for (let r = dataStart; r <= dataEnd; r++) {
    const row: unknown[] = [];
    for (let c = bounds.c1; c <= bounds.c2; c++) {
      const cell = ws.getCell(r, c);
      const raw = cell.value;
      if (raw && typeof raw === 'object' && 'result' in raw) {
        row.push((raw as ExcelJS.CellFormulaValue).result ?? null);
      } else if (raw && typeof raw === 'object' && 'text' in raw) {
        row.push((raw as ExcelJS.CellHyperlinkValue).text);
      } else {
        row.push(raw ?? null);
      }
    }
    rows.push(row);
  }
  // ExcelJS requires at least an empty rows array; ensure column names unique
  const seen = new Set<string>();
  for (const col of columns) {
    let n = col.name;
    let i = 2;
    while (seen.has(n)) {
      n = `${col.name}_${i++}`;
    }
    seen.add(n);
    col.name = n;
  }
  return {
    name: t.name,
    ref: t.ref,
    headerRow,
    totalsRow,
    columns,
    rows: rows as ExcelJS.TableProperties['rows'],
  };
}

function applyTables(ws: ExcelJS.Worksheet, tables: SheetTable[]): void {
  for (const t of tables) {
    const payload = buildTablePayload(ws, t);
    if (!payload) continue;
    try {
      ws.addTable(payload);
    } catch {
      // Overlapping / invalid table — skip
    }
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

  if (sheet.validations?.length) {
    applyValidations(ws, sheet.validations);
  }

  if (sheet.tables?.length) {
    applyTables(ws, sheet.tables);
  }

  if (sheet.autoFilter) {
    ws.autoFilter = sheet.autoFilter;
  }
}

function applyNamedRanges(workbook: ExcelJS.Workbook, names: NamedRange[]): void {
  for (const n of names) {
    if (!n.name || !n.refersTo) continue;
    const definedName = n.scope ? `${n.scope}!${n.name}` : n.name;
    try {
      // definedNames.add(locStr, name) — locStr is the refersTo range
      const parts = n.refersTo.split(',').map((s) => s.trim()).filter(Boolean);
      for (const loc of parts) {
        workbook.definedNames.add(loc, definedName);
      }
    } catch {
      // Invalid name / ref — skip
    }
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

  if (snapshot.names?.length) {
    applyNamedRanges(workbook, snapshot.names);
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

