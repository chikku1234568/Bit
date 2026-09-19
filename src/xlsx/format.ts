/**
 * Format normalisation and colour helpers shared by parse, write, diff, merge.
 */

import type {
  BorderEdge,
  CellAlignment,
  CellBorders,
  CellFormat,
} from './types.js';

/** Office theme scheme (lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink). */
const OFFICE_THEME = [
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  '0563C1',
  '954F72',
];

const INDEXED: Record<number, string> = {
  0: '000000',
  1: 'FFFFFF',
  2: 'FF0000',
  3: '00FF00',
  4: '0000FF',
  5: 'FFFF00',
  6: 'FF00FF',
  7: '00FFFF',
  8: '000000',
  9: 'FFFFFF',
  10: 'FF0000',
  11: '00FF00',
  12: '0000FF',
  13: 'FFFF00',
  14: 'FF00FF',
  15: '00FFFF',
  64: '000000',
  65: 'FFFFFF',
};

export function colToLetter(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function letterToCol(letter: string): number {
  let n = 0;
  for (const ch of letter.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

export function parseAddress(addr: string): { col: number; row: number } | null {
  const m = addr.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return null;
  return { col: letterToCol(m[1]), row: Number(m[2]) };
}

/** True if address is inside a merge but is not the top-left (master) cell. */
export function isMergeSlave(address: string, merges: string[] | undefined): boolean {
  if (!merges?.length) return false;
  const pos = parseAddress(address);
  if (!pos) return false;
  for (const range of merges) {
    const [start, end] = range.split(':');
    if (!start || !end) continue;
    const a = parseAddress(start);
    const b = parseAddress(end);
    if (!a || !b) continue;
    const c1 = Math.min(a.col, b.col);
    const c2 = Math.max(a.col, b.col);
    const r1 = Math.min(a.row, b.row);
    const r2 = Math.max(a.row, b.row);
    if (pos.col >= c1 && pos.col <= c2 && pos.row >= r1 && pos.row <= r2) {
      const masterCol = Math.min(a.col, b.col);
      const masterRow = Math.min(a.row, b.row);
      if (pos.col !== masterCol || pos.row !== masterRow) return true;
    }
  }
  return false;
}

export function normalizeHex(c: string | undefined): string | undefined {
  if (!c) return undefined;
  const cleaned = c.replace(/^#/, '').toUpperCase();
  if (cleaned.length === 8) return cleaned.slice(2);
  if (cleaned.length === 6) return cleaned;
  return undefined;
}

export function hexToArgb(hex: string): string {
  const cleaned = hex.replace(/^#/, '').toUpperCase();
  if (cleaned.length === 6) return `FF${cleaned}`;
  if (cleaned.length === 8) return cleaned;
  return `FF${cleaned.padStart(6, '0').slice(0, 6)}`;
}

export function argbToHex(argb: string | undefined): string | undefined {
  const n = normalizeHex(argb);
  return n ? `#${n}` : undefined;
}

function applyTint(rgb: string, tint: number): string {
  const ch = (hex: string): number => {
    const v = parseInt(hex, 16);
    if (tint < 0) return Math.round(v * (1 + tint));
    return Math.round(v * (1 - tint) + 255 * tint);
  };
  const r = ch(rgb.slice(0, 2));
  const g = ch(rgb.slice(2, 4));
  const b = ch(rgb.slice(4, 6));
  const to = (n: number) =>
    Math.max(0, Math.min(255, n)).toString(16).padStart(2, '0').toUpperCase();
  return `${to(r)}${to(g)}${to(b)}`;
}

export interface ExcelColorLike {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

export function excelColorToHex(color: ExcelColorLike | undefined): string | undefined {
  if (!color) return undefined;
  if (typeof color.argb === 'string') {
    const hex = argbToHex(color.argb);
    if (hex) return hex;
  }
  if (typeof color.theme === 'number' && color.theme >= 0 && color.theme < OFFICE_THEME.length) {
    let rgb = OFFICE_THEME[color.theme];
    if (typeof color.tint === 'number' && color.tint !== 0) {
      rgb = applyTint(rgb, color.tint);
    }
    return `#${rgb}`;
  }
  if (typeof color.indexed === 'number' && INDEXED[color.indexed]) {
    return `#${INDEXED[color.indexed]}`;
  }
  return undefined;
}

function normalizeBorders(b: CellBorders | undefined): CellBorders | undefined {
  if (!b) return undefined;
  const edge = (e: BorderEdge | undefined): BorderEdge | undefined => {
    if (!e) return undefined;
    const out: BorderEdge = {};
    if (e.style) out.style = e.style;
    const c = normalizeHex(e.color);
    if (c) out.color = `#${c}`;
    return Object.keys(out).length ? out : undefined;
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

function normalizeAlign(a: CellAlignment | undefined): CellAlignment | undefined {
  if (!a) return undefined;
  const out: CellAlignment = {};
  if (a.horizontal) out.horizontal = a.horizontal;
  if (a.vertical) out.vertical = a.vertical;
  if (a.wrapText) out.wrapText = true;
  if (a.indent) out.indent = a.indent;
  if (a.textRotation) out.textRotation = a.textRotation;
  return Object.keys(out).length ? out : undefined;
}

/** Drop empty / default-equivalent format so equality is stable. */
export function normalizeFmt(fmt: CellFormat | undefined): CellFormat | null {
  if (!fmt) return null;
  const out: CellFormat = {};
  if (fmt.numFmt) out.numFmt = fmt.numFmt;
  if (fmt.bold) out.bold = true;
  if (fmt.italic) out.italic = true;
  if (fmt.underline) out.underline = fmt.underline;
  if (fmt.strike) out.strike = true;
  if (fmt.fontName) out.fontName = fmt.fontName;
  if (fmt.fontSize != null) out.fontSize = fmt.fontSize;
  const fill = normalizeHex(fmt.fill);
  if (fill) out.fill = `#${fill}`;
  const fontColor = normalizeHex(fmt.fontColor);
  // Default Excel text colour — omit so it doesn't look like a change.
  if (fontColor && fontColor !== '000000') out.fontColor = `#${fontColor}`;
  const borders = normalizeBorders(fmt.borders);
  if (borders) out.borders = borders;
  const alignment = normalizeAlign(fmt.alignment);
  if (alignment) out.alignment = alignment;
  return Object.keys(out).length === 0 ? null : out;
}

export function fmtEqual(
  a: CellFormat | undefined,
  b: CellFormat | undefined,
): boolean {
  return JSON.stringify(normalizeFmt(a)) === JSON.stringify(normalizeFmt(b));
}

export const DEFAULT_FONT_NAME = 'Calibri';
export const DEFAULT_FONT_SIZE = 11;
export const DEFAULT_ROW_HEIGHT = 15;
export const DEFAULT_COL_WIDTHS = new Set([8, 8.43, 9, 10]);
