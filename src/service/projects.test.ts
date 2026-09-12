import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { writeXlsx, parseXlsx } from '../xlsx/index.js';
import type { WorkbookSnapshot } from '../xlsx/types.js';
import { BitStore } from '../store/store.js';
import { ProjectService } from './projects.js';
import { buildApp } from '../api/app.js';

function budgetSnapshot(
  overrides?: Partial<WorkbookSnapshot['sheets']['Budget']['cells']>,
): WorkbookSnapshot {
  return {
    sheetOrder: ['Budget'],
    sheets: {
      Budget: {
        dimensions: { rows: 5, cols: 3 },
        cells: {
          A1: { v: 'Item', f: null, fmt: { bold: true } },
          B1: { v: 'Amount', f: null, fmt: { bold: true } },
          A2: { v: 'Revenue', f: null },
          B2: { v: 1000, f: null, fmt: { numFmt: '#,##0.00' } },
          A3: { v: 'Tax', f: null },
          B3: { v: null, f: '=B2*0.2', fmt: { numFmt: '#,##0.00' } },
          ...overrides,
        },
      },
    },
  };
}

function multipartPayload(
  fields: Record<string, string>,
  file: { field: string; filename: string; buffer: Buffer },
): { payload: Buffer; contentType: string } {
  const boundary = '----BitTestBoundary' + Math.random().toString(16).slice(2);
  const parts: Buffer[] = [];
  for (const [name, value] of Object.entries(fields)) {
    parts.push(
      Buffer.from(
        `--${boundary}\r\nContent-Disposition: form-data; name="${name}"\r\n\r\n${value}\r\n`,
      ),
    );
  }
  parts.push(
    Buffer.from(
      `--${boundary}\r\nContent-Disposition: form-data; name="${file.field}"; filename="${file.filename}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`,
    ),
  );
  parts.push(file.buffer);
  parts.push(Buffer.from(`\r\n--${boundary}--\r\n`));
  return {
    payload: Buffer.concat(parts),
    contentType: `multipart/form-data; boundary=${boundary}`,
  };
}

describe('M2 ProjectService', () => {
  let dataDir: string;
  let store: BitStore;
  let service: ProjectService;

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m2-'));
    store = new BitStore(dataDir);
    await store.init();
    service = new ProjectService(store);
  });

  afterEach(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  it('create project → save version → list history → download round-trips tracked fields', async () => {
    const v1Snap = budgetSnapshot();
    const v1Buf = await writeXlsx(v1Snap);

    const project = await service.createProject({
      name: 'FY27 Budget',
      author: 'Alex',
      message: 'Initial version',
      xlsxBuffer: v1Buf,
    });

    expect(project.name).toBe('FY27 Budget');
    expect(project.main.name).toBe('Main');
    expect(project.main.isMain).toBe(true);
    expect(project.tipVersion?.author).toBe('Alex');
    expect(project.tipVersion?.message).toBe('Initial version');
    expect(project.tipVersion?.parentIds).toEqual([]);

    const v2Snap = budgetSnapshot({
      B2: { v: 1200, f: null, fmt: { numFmt: '#,##0.00' } },
    });
    const v2Buf = await writeXlsx(v2Snap);

    const v2 = await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'Updated revenue assumption',
      xlsxBuffer: v2Buf,
    });

    expect(v2.parentIds).toEqual([project.tipVersion!.id]);
    expect(v2.message).toBe('Updated revenue assumption');

    const history = await service.listVersions(project.id);
    expect(history).toHaveLength(2);
    expect(history[0].id).toBe(v2.id);
    expect(history[0].scenarioName).toBe('Main');
    expect(history[1].id).toBe(project.tipVersion!.id);

    const refreshed = await service.getProject(project.id);
    expect(refreshed.tipVersion?.id).toBe(v2.id);

    const { buffer, filename } = await service.getVersionXlsx(v2.id);
    expect(filename).toMatch(/\.xlsx$/);
    expect(buffer.length).toBeGreaterThan(100);

    const parsed = await parseXlsx(buffer);
    expect(parsed.sheetOrder).toEqual(['Budget']);
    expect(parsed.sheets.Budget.cells.B2.v).toBe(1200);
    expect(parsed.sheets.Budget.cells.B3.f).toBe('=B2*0.2');
    expect(parsed.sheets.Budget.cells.A1.fmt?.bold).toBe(true);
    expect(parsed.sheets.Budget.cells.B2.fmt?.numFmt).toBe('#,##0.00');
  });
});

describe('M2 API', () => {
  let dataDir: string;
  let app: Awaited<ReturnType<typeof buildApp>>['app'];

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m2-api-'));
    const built = await buildApp({ dataDir, logger: false });
    app = built.app;
    await app.ready();
  });

  afterEach(async () => {
    await app.close();
    await rm(dataDir, { recursive: true, force: true });
  });

  it('HTTP: create → save → history → download', async () => {
    const buf = await writeXlsx(budgetSnapshot());
    const createMp = multipartPayload(
      { name: 'API Budget', author: 'Jordan', message: 'Kickoff' },
      { field: 'file', filename: 'budget.xlsx', buffer: buf },
    );

    const createRes = await app.inject({
      method: 'POST',
      url: '/projects',
      headers: { 'content-type': createMp.contentType },
      payload: createMp.payload,
    });
    expect(createRes.statusCode).toBe(201);
    const project = createRes.json();
    expect(project.name).toBe('API Budget');
    expect(project.main.name).toBe('Main');
    const version1Id = project.tipVersion.id as string;

    const getRes = await app.inject({ method: 'GET', url: `/projects/${project.id}` });
    expect(getRes.statusCode).toBe(200);
    expect(getRes.json().name).toBe('API Budget');

    const v2Buf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 1500, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const saveMp = multipartPayload(
      { message: 'Revenue up', author: 'Jordan' },
      { field: 'file', filename: 'budget-v2.xlsx', buffer: v2Buf },
    );
    const saveRes = await app.inject({
      method: 'POST',
      url: `/projects/${project.id}/versions`,
      headers: {
        'content-type': saveMp.contentType,
        'x-bit-author': 'Jordan',
      },
      payload: saveMp.payload,
    });
    expect(saveRes.statusCode).toBe(201);
    const saved = saveRes.json();
    expect(saved.message).toBe('Revenue up');
    expect(saved.parentIds).toEqual([version1Id]);

    const histRes = await app.inject({
      method: 'GET',
      url: `/projects/${project.id}/versions`,
    });
    expect(histRes.statusCode).toBe(200);
    expect(histRes.json()).toHaveLength(2);

    const dl = await app.inject({
      method: 'GET',
      url: `/versions/${saved.id}/xlsx`,
    });
    expect(dl.statusCode).toBe(200);
    expect(String(dl.headers['content-type'])).toContain('spreadsheetml');
    const downloaded = Buffer.from(dl.rawPayload);
    const round = await parseXlsx(downloaded);
    expect(round.sheets.Budget.cells.B2.v).toBe(1500);
    expect(round.sheets.Budget.cells.B3.f).toBe('=B2*0.2');
  });
});
