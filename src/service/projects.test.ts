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

describe('M3 Scenarios', () => {
  let dataDir: string;
  let store: BitStore;
  let service: ProjectService;

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m3-'));
    store = new BitStore(dataDir);
    await store.init();
    service = new ProjectService(store);
  });

  afterEach(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  it('create scenario tip equals Main tip; save on scenario leaves Main tip; Main advances independently', async () => {
    const v1Buf = await writeXlsx(budgetSnapshot());
    const project = await service.createProject({
      name: 'Scenario Budget',
      author: 'Alex',
      message: 'Initial',
      xlsxBuffer: v1Buf,
    });
    const mainTipId = project.tipVersion!.id;

    const scenario = await service.createScenario({
      projectId: project.id,
      name: 'Tax update',
      author: 'Jordan',
    });
    expect(scenario.tipVersionId).toBe(mainTipId);
    expect(scenario.isMain).toBe(false);
    expect(scenario.name).toBe('Tax update');

    const listed = await service.listScenarios(project.id);
    expect(listed[0].isMain).toBe(true);
    expect(listed.map((s) => s.name)).toEqual(['Main', 'Tax update']);

    const scenBuf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 1100, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const scenVer = await service.saveVersion({
      projectId: project.id,
      scenarioId: scenario.id,
      author: 'Jordan',
      message: 'Tax scenario edit',
      xlsxBuffer: scenBuf,
    });
    expect(scenVer.parentIds).toEqual([mainTipId]);

    const afterScen = await service.getProject(project.id);
    expect(afterScen.main.tipVersionId).toBe(mainTipId);
    expect(afterScen.tipVersion?.id).toBe(mainTipId);

    const scenFresh = await service.getScenario(scenario.id);
    expect(scenFresh.tipVersionId).toBe(scenVer.id);

    const mainBuf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 2000, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const mainVer = await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'Main moved on',
      xlsxBuffer: mainBuf,
    });
    expect(mainVer.parentIds).toEqual([mainTipId]);

    const final = await service.getProject(project.id);
    expect(final.main.tipVersionId).toBe(mainVer.id);
    const scenStill = await service.getScenario(scenario.id);
    expect(scenStill.tipVersionId).toBe(scenVer.id);
  });

  it('rejects blank and Main scenario names', async () => {
    const project = await service.createProject({
      name: 'Names',
      author: 'Alex',
      xlsxBuffer: await writeXlsx(budgetSnapshot()),
    });
    await expect(
      service.createScenario({ projectId: project.id, name: '  ' }),
    ).rejects.toThrow(/required/i);
    await expect(
      service.createScenario({ projectId: project.id, name: 'Main' }),
    ).rejects.toThrow(/Main/i);
    await expect(
      service.createScenario({ projectId: project.id, name: 'main' }),
    ).rejects.toThrow(/Main/i);
  });
});

describe('M3 API', () => {
  let dataDir: string;
  let app: Awaited<ReturnType<typeof buildApp>>['app'];

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m3-api-'));
    const built = await buildApp({ dataDir, logger: false });
    app = built.app;
    await app.ready();
  });

  afterEach(async () => {
    await app.close();
    await rm(dataDir, { recursive: true, force: true });
  });

  it('HTTP POST scenario + POST versions + GET project shows both tips', async () => {
    const buf = await writeXlsx(budgetSnapshot());
    const createMp = multipartPayload(
      { name: 'M3 API', author: 'Alex', message: 'Start' },
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
    const mainTip = project.tipVersion.id as string;

    const scenRes = await app.inject({
      method: 'POST',
      url: `/projects/${project.id}/scenarios`,
      headers: { 'content-type': 'application/json', 'x-bit-author': 'Jordan' },
      payload: JSON.stringify({ name: 'Tax update' }),
    });
    expect(scenRes.statusCode).toBe(201);
    const scenario = scenRes.json();
    expect(scenario.tipVersionId).toBe(mainTip);

    const v2Buf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 1300, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const saveMp = multipartPayload(
      { message: 'Scenario save', author: 'Jordan' },
      { field: 'file', filename: 's.xlsx', buffer: v2Buf },
    );
    const saveRes = await app.inject({
      method: 'POST',
      url: `/scenarios/${scenario.id}/versions`,
      headers: { 'content-type': saveMp.contentType, 'x-bit-author': 'Jordan' },
      payload: saveMp.payload,
    });
    expect(saveRes.statusCode).toBe(201);
    const saved = saveRes.json();
    expect(saved.parentIds).toEqual([mainTip]);

    const getRes = await app.inject({ method: 'GET', url: `/projects/${project.id}` });
    expect(getRes.statusCode).toBe(200);
    const detail = getRes.json();
    expect(detail.main.tipVersionId).toBe(mainTip);
    const tax = detail.scenarios.find((s: { name: string }) => s.name === 'Tax update');
    expect(tax.tipVersionId).toBe(saved.id);
  });
});

describe('M4 Diff API', () => {
  let dataDir: string;
  let app: Awaited<ReturnType<typeof buildApp>>['app'];

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m4-api-'));
    const built = await buildApp({ dataDir, logger: false });
    app = built.app;
    await app.ready();
  });

  afterEach(async () => {
    await app.close();
    await rm(dataDir, { recursive: true, force: true });
  });

  it('GET /diff returns changes; 404 missing; 400 different projects', async () => {
    const buf = await writeXlsx(budgetSnapshot());
    const createMp = multipartPayload(
      { name: 'Diff Proj', author: 'Alex', message: 'v1' },
      { field: 'file', filename: 'b.xlsx', buffer: buf },
    );
    const createRes = await app.inject({
      method: 'POST',
      url: '/projects',
      headers: { 'content-type': createMp.contentType },
      payload: createMp.payload,
    });
    const project = createRes.json();
    const v1 = project.tipVersion.id as string;

    const v2Buf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 999, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const saveMp = multipartPayload(
      { message: 'v2', author: 'Alex' },
      { field: 'file', filename: 'b2.xlsx', buffer: v2Buf },
    );
    const saveRes = await app.inject({
      method: 'POST',
      url: `/projects/${project.id}/versions`,
      headers: { 'content-type': saveMp.contentType },
      payload: saveMp.payload,
    });
    const v2 = saveRes.json().id as string;

    const diffRes = await app.inject({
      method: 'GET',
      url: `/diff?base=${v1}&compare=${v2}`,
    });
    expect(diffRes.statusCode).toBe(200);
    const body = diffRes.json();
    expect(body.changes.some((c: { kind: string }) => c.kind === 'cell-value')).toBe(true);

    const miss = await app.inject({
      method: 'GET',
      url: `/diff?base=${v1}&compare=00000000-0000-0000-0000-000000000000`,
    });
    expect(miss.statusCode).toBe(404);

    const otherMp = multipartPayload(
      { name: 'Other', author: 'Alex', message: 'o1' },
      { field: 'file', filename: 'o.xlsx', buffer: buf },
    );
    const other = await app.inject({
      method: 'POST',
      url: '/projects',
      headers: { 'content-type': otherMp.contentType },
      payload: otherMp.payload,
    });
    const otherId = other.json().tipVersion.id as string;
    const bad = await app.inject({
      method: 'GET',
      url: `/diff?base=${v1}&compare=${otherId}`,
    });
    expect(bad.statusCode).toBe(400);
  });
});

describe('M6 Review + Combine', () => {
  let dataDir: string;
  let store: BitStore;
  let service: ProjectService;

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m6-'));
    store = new BitStore(dataDir);
    await store.init();
    service = new ProjectService(store);
  });

  afterEach(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  it('clean combine → Main version with two parents', async () => {
    const project = await service.createProject({
      name: 'Review Budget',
      author: 'Alex',
      message: 'Initial',
      xlsxBuffer: await writeXlsx(budgetSnapshot()),
    });
    const mainTip = project.tipVersion!.id;
    const scenario = await service.createScenario({
      projectId: project.id,
      name: 'Tax update',
      author: 'Jordan',
    });
    const scenVer = await service.saveVersion({
      projectId: project.id,
      scenarioId: scenario.id,
      author: 'Jordan',
      message: 'Tax bump',
      xlsxBuffer: await writeXlsx(
        budgetSnapshot({
          B2: { v: 1500, f: null, fmt: { numFmt: '#,##0.00' } },
        }),
      ),
    });

    const review = await service.createReview({
      scenarioId: scenario.id,
      author: 'Jordan',
      note: 'Please review tax',
    });
    expect(review.baseVersionId).toBe(mainTip);
    expect(review.compareVersionId).toBe(scenVer.id);

    const { version, review: combined } = await service.combineReview({
      reviewId: review.id,
      author: 'Alex',
      message: 'Combined tax update',
    });
    expect(combined.status).toBe('combined');
    expect(version.parentIds).toEqual([mainTip, scenVer.id]);
    expect(version.scenarioId).toBe(project.mainScenarioId);
    const refreshed = await service.getProject(project.id);
    expect(refreshed.main.tipVersionId).toBe(version.id);
  });

  it('conflict without resolutions throws ConflictError', async () => {
    const { ConflictError } = await import('./projects.js');
    const project = await service.createProject({
      name: 'Conflict',
      author: 'Alex',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 1000, f: null } })),
    });
    const scenario = await service.createScenario({
      projectId: project.id,
      name: 'Alt',
    });
    await service.saveVersion({
      projectId: project.id,
      scenarioId: scenario.id,
      author: 'Jordan',
      message: 'Scenario 1111',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 1111, f: null } })),
    });
    const review = await service.createReview({
      scenarioId: scenario.id,
      author: 'Jordan',
    });
    // Main moves after review opened → overlapping edit
    await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'Main 2222',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 2222, f: null } })),
    });

    await expect(
      service.combineReview({ reviewId: review.id, author: 'Alex' }),
    ).rejects.toBeInstanceOf(ConflictError);
  });

  it('conflict with resolutions applied → combine succeeds', async () => {
    const project = await service.createProject({
      name: 'Resolve',
      author: 'Alex',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 1000, f: null } })),
    });
    const scenario = await service.createScenario({
      projectId: project.id,
      name: 'Alt',
    });
    const scenVer = await service.saveVersion({
      projectId: project.id,
      scenarioId: scenario.id,
      author: 'Jordan',
      message: 'Scenario 1111',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 1111, f: null } })),
    });
    const review = await service.createReview({
      scenarioId: scenario.id,
      author: 'Jordan',
    });
    await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'Main 2222',
      xlsxBuffer: await writeXlsx(budgetSnapshot({ B2: { v: 2222, f: null } })),
    });

    const { version } = await service.combineReview({
      reviewId: review.id,
      author: 'Alex',
      resolutions: {
        'Budget!B2': { action: 'keep-theirs' },
      },
    });
    expect(version.parentIds).toHaveLength(2);
    expect(version.parentIds).toContain(scenVer.id);
    const snap = await service.getVersionSnapshot(version.id);
    expect(snap.sheets.Budget.cells.B2.v).toBe(1111);
  });

  it('cannot ask for review from Main', async () => {
    const project = await service.createProject({
      name: 'NoMainReview',
      author: 'Alex',
      xlsxBuffer: await writeXlsx(budgetSnapshot()),
    });
    await expect(
      service.createReview({ scenarioId: project.mainScenarioId, author: 'Alex' }),
    ).rejects.toThrow(/Main/i);
  });
});

describe('M6 API', () => {
  let dataDir: string;
  let app: Awaited<ReturnType<typeof buildApp>>['app'];

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-m6-api-'));
    const built = await buildApp({ dataDir, logger: false });
    app = built.app;
    await app.ready();
  });

  afterEach(async () => {
    await app.close();
    await rm(dataDir, { recursive: true, force: true });
  });

  it('HTTP ask for review + combine', async () => {
    const buf = await writeXlsx(budgetSnapshot());
    const createMp = multipartPayload(
      { name: 'M6 API', author: 'Alex', message: 'Start' },
      { field: 'file', filename: 'b.xlsx', buffer: buf },
    );
    const project = (
      await app.inject({
        method: 'POST',
        url: '/projects',
        headers: { 'content-type': createMp.contentType },
        payload: createMp.payload,
      })
    ).json();

    const scenario = (
      await app.inject({
        method: 'POST',
        url: `/projects/${project.id}/scenarios`,
        headers: { 'content-type': 'application/json' },
        payload: JSON.stringify({ name: 'Tax update' }),
      })
    ).json();

    const v2Buf = await writeXlsx(
      budgetSnapshot({
        B2: { v: 1600, f: null, fmt: { numFmt: '#,##0.00' } },
      }),
    );
    const saveMp = multipartPayload(
      { message: 'Scenario save', author: 'Jordan' },
      { field: 'file', filename: 's.xlsx', buffer: v2Buf },
    );
    await app.inject({
      method: 'POST',
      url: `/scenarios/${scenario.id}/versions`,
      headers: { 'content-type': saveMp.contentType },
      payload: saveMp.payload,
    });

    const reviewRes = await app.inject({
      method: 'POST',
      url: `/scenarios/${scenario.id}/reviews`,
      headers: { 'content-type': 'application/json', 'x-bit-author': 'Jordan' },
      payload: JSON.stringify({ note: 'Please look' }),
    });
    expect(reviewRes.statusCode).toBe(201);
    const review = reviewRes.json();

    const combineRes = await app.inject({
      method: 'POST',
      url: `/reviews/${review.id}/combine`,
      headers: { 'content-type': 'application/json', 'x-bit-author': 'Alex' },
      payload: JSON.stringify({ message: 'Looks good' }),
    });
    expect(combineRes.statusCode).toBe(200);
    const body = combineRes.json();
    expect(body.review.status).toBe('combined');
    expect(body.version.parentIds).toHaveLength(2);
  });
});
