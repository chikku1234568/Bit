#!/usr/bin/env node
/**
 * CLI: npm run roundtrip -- sample.xlsx
 * Parses → writes sample.roundtrip.xlsx and prints counts.
 */
import { writeFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { parseXlsx, writeXlsx } from '../src/xlsx/index.js';

async function main(): Promise<void> {
  const inputArg = process.argv[2];
  if (!inputArg) {
    console.error('Usage: npm run roundtrip -- <file.xlsx>');
    process.exit(1);
  }

  const inputPath = resolve(inputArg);
  console.log(`Reading: ${inputPath}`);

  const snapshot = await parseXlsx(inputPath);

  let cellCount = 0;
  let formulaCount = 0;
  for (const name of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[name];
    const cells = Object.values(sheet.cells);
    cellCount += cells.length;
    formulaCount += cells.filter((c) => c.f).length;
  }

  console.log(`Sheets: ${snapshot.sheetOrder.length} (${snapshot.sheetOrder.join(', ')})`);
  console.log(`Cells:  ${cellCount}`);
  console.log(`Formulas: ${formulaCount}`);

  for (const name of snapshot.sheetOrder) {
    const sheet = snapshot.sheets[name];
    const n = Object.keys(sheet.cells).length;
    console.log(
      `  - ${name}: ${n} cells, dims ${sheet.dimensions.rows}x${sheet.dimensions.cols}`,
    );
  }

  const buf = await writeXlsx(snapshot);
  const dir = dirname(inputPath);
  const base = basename(inputPath, '.xlsx');
  const outPath = join(dir, `${base}.roundtrip.xlsx`);
  await writeFile(outPath, buf);
  console.log(`Wrote: ${outPath} (${buf.length} bytes)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
