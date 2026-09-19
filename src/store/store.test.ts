import { mkdtemp, rm, readFile, writeFile, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { writeXlsx } from '../xlsx/index.js';
import type { WorkbookSnapshot } from '../xlsx/types.js';
import { BitStore, StoreConflictError, CAS_MESSAGE } from './store.js';
import { ProjectService, ConflictError } from '../service/projects.js';
import { buildApp } from '../api/app.js';

function snap(n = 1): WorkbookSnapshot {
  return {
    sheetOrder: ['Budget'],
    sheets: {
      Budget: {
        dimensions: { rows: 2, cols: 2 },
        cells: {
          A1: { v: 'Item', f: null },
          B1: { v: n, f: null },
        },
      },
    },
  };
}

describe('append-only store + CAS', () => {
  let dataDir: string;
  let store: BitStore;
  let service: ProjectService;

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-store-'));
    store = new BitStore(dataDir);
    await store.init();
    service = new ProjectService(store);
  });

  afterEach(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  it('writes project.json + versions/ files; never mutates an old version file', async () => {
    const buf = await writeXlsx(snap(1));
    const project = await service.createProject({
      name: 'CAS Demo',
      author: 'Alex',
      message: 'v1',
      xlsxBuffer: buf,
    });
    const v1Id = project.tipVersion!.id;
    const v1Path = store.versionPath(v1Id);
    const before = await readFile(v1Path, 'utf8');

    const v2 = await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'v2',
      xlsxBuffer: await writeXlsx(snap(2)),
    });

    const after = await readFile(v1Path, 'utf8');
    expect(after).toBe(before);
    expect(v2.parentIds).toEqual([v1Id]);

    const catalogRaw = await readFile(join(dataDir, 'project.json'), 'utf8');
    const catalog = JSON.parse(catalogRaw);
    expect(catalog.projects).toHaveLength(1);
    expect(catalog.versions).toBeUndefined();
    expect(catalog.scenarios[0].tipVersionId).toBe(v2.id);

    const refreshed = await service.getProject(project.id);
    expect(refreshed.tipVersion?.id).toBe(v2.id);
  });

  it('CAS race on save returns 409 ConflictError', async () => {
    const buf = await writeXlsx(snap(1));
    const project = await service.createProject({
      name: 'Race',
      author: 'Alex',
      message: 'v1',
      xlsxBuffer: buf,
    });
    const tip = project.tipVersion!.id;

    // First save advances tip
    await service.saveVersion({
      projectId: project.id,
      author: 'Alex',
      message: 'winner',
      xlsxBuffer: await writeXlsx(snap(2)),
      expectedTipVersionId: tip,
    });

    // Second save still expecting old tip → 409
    await expect(
      service.saveVersion({
        projectId: project.id,
        author: 'Jordan',
        message: 'loser',
        xlsxBuffer: await writeXlsx(snap(3)),
        expectedTipVersionId: tip,
      }),
    ).rejects.toMatchObject({
      name: 'ConflictError',
      message: CAS_MESSAGE,
    });

    const main = await service.getProject(project.id);
    expect(main.tipVersion?.message).toBe('winner');
  });

  it('successful save advances tip via CAS', async () => {
    const project = await service.createProject({
      name: 'Advance',
      author: 'A',
      message: 'start',
      xlsxBuffer: await writeXlsx(snap(1)),
    });
    const tip1 = project.tipVersion!.id;
    const v2 = await service.saveVersion({
      projectId: project.id,
      author: 'A',
      message: 'next',
      xlsxBuffer: await writeXlsx(snap(9)),
      expectedTipVersionId: tip1,
    });
    const cat = await store.readCatalog();
    const main = cat.scenarios.find((s) => s.isMain)!;
    expect(main.tipVersionId).toBe(v2.id);
  });

  it('migrate-on-read from fat meta.json', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'bit-migrate-'));
    try {
      await mkdir(join(dir, 'blobs'), { recursive: true });
      await mkdir(join(dir, 'xlsx'), { recursive: true });
      const projectId = 'p1';
      const scenarioId = 's1';
      const versionId = 'v1';
      const hash = 'abc';
      await writeFile(
        join(dir, 'blobs', `${hash}.json`),
        JSON.stringify(snap(1)),
        'utf8',
      );
      await writeFile(
        join(dir, 'meta.json'),
        JSON.stringify(
          {
            projects: [
              {
                id: projectId,
                name: 'Legacy',
                createdAt: '2020-01-01T00:00:00.000Z',
                mainScenarioId: scenarioId,
              },
            ],
            scenarios: [
              {
                id: scenarioId,
                projectId,
                name: 'Main',
                tipVersionId: versionId,
                isMain: true,
              },
            ],
            versions: [
              {
                id: versionId,
                projectId,
                scenarioId,
                parentIds: [],
                author: 'old',
                timestamp: '2020-01-01T00:00:00.000Z',
                message: 'legacy',
                snapshotHash: hash,
              },
            ],
            reviews: [],
          },
          null,
          2,
        ),
        'utf8',
      );

      const migrated = new BitStore(dir);
      await migrated.init();
      const catalog = await migrated.readCatalog();
      expect(catalog.projects[0].name).toBe('Legacy');
      expect(catalog.scenarios[0].tipVersionId).toBe(versionId);
      const v = await migrated.getVersion(versionId);
      expect(v?.message).toBe('legacy');
      // Slim catalog has no versions key on disk
      const pj = JSON.parse(await readFile(join(dir, 'project.json'), 'utf8'));
      expect(pj.versions).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('fresh project.lock causes conflict when acquiring lock', async () => {
    const project = await service.createProject({
      name: 'Lock',
      author: 'A',
      message: 'v1',
      xlsxBuffer: await writeXlsx(snap(1)),
    });
    await writeFile(
      join(dataDir, 'project.lock'),
      JSON.stringify({
        holder: 'other',
        expiresAt: new Date(Date.now() + 60_000).toISOString(),
      }),
      'utf8',
    );
    await expect(
      store.withLock(async () => 'ok', 'bit', 15_000, 120),
    ).rejects.toBeInstanceOf(StoreConflictError);
    // Tip unchanged
    const cat = await store.readCatalog();
    expect(cat.scenarios.find((s) => s.id === project.mainScenarioId)?.tipVersionId).toBe(
      project.tipVersion!.id,
    );
  });
});

describe('promote to Main', () => {
  let dataDir: string;
  let store: BitStore;
  let service: ProjectService;

  beforeEach(async () => {
    dataDir = await mkdtemp(join(tmpdir(), 'bit-promote-'));
    store = new BitStore(dataDir);
    await store.init();
    service = new ProjectService(store);
  });

  afterEach(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  it('promote creates Main tip reusing snapshot; CAS on promote race', async () => {
    const project = await service.createProject({
      name: 'Promo',
      author: 'Alex',
      message: 'main-v1',
      xlsxBuffer: await writeXlsx(snap(1)),
    });
    const mainTip = project.tipVersion!.id;

    const scenario = await service.createScenario({
      projectId: project.id,
      name: 'Jordan draft',
    });
    const jordan = await service.saveVersion({
      projectId: project.id,
      scenarioId: scenario.id,
      author: 'Jordan',
      message: 'jordan edit',
      xlsxBuffer: await writeXlsx(snap(42)),
    });

    const promoted = await service.promoteToMain({
      versionId: jordan.id,
      author: 'Alex',
      message: 'Make Jordan Main',
      expectedMainTip: mainTip,
    });

    expect(promoted.scenarioId).toBe(project.mainScenarioId);
    expect(promoted.snapshotHash).toBe(jordan.snapshotHash);
    expect(promoted.promotedFromVersionId).toBe(jordan.id);
    expect(promoted.parentIds).toEqual([mainTip]);

    const after = await service.getProject(project.id);
    expect(after.tipVersion?.id).toBe(promoted.id);

    // Race: another promote still expecting old tip
    await expect(
      service.promoteToMain({
        versionId: jordan.id,
        author: 'Other',
        expectedMainTip: mainTip,
      }),
    ).rejects.toBeInstanceOf(ConflictError);
  });

  it('HTTP POST /versions/:id/promote', async () => {
    const built = await buildApp({ dataDir, logger: false });
    await built.app.ready();
    try {
      const project = await service.createProject({
        name: 'HTTP Promo',
        author: 'A',
        message: 'v1',
        xlsxBuffer: await writeXlsx(snap(1)),
      });
      const sc = await service.createScenario({ projectId: project.id, name: 'Draft' });
      const draft = await service.saveVersion({
        projectId: project.id,
        scenarioId: sc.id,
        author: 'B',
        message: 'draft',
        xlsxBuffer: await writeXlsx(snap(7)),
      });

      const res = await built.app.inject({
        method: 'POST',
        url: `/versions/${draft.id}/promote`,
        headers: { 'content-type': 'application/json', 'x-bit-author': 'A' },
        payload: JSON.stringify({ message: 'Make this Main' }),
      });
      expect(res.statusCode).toBe(201);
      const body = res.json();
      expect(body.promotedFromVersionId).toBe(draft.id);
      expect(body.message).toBe('Make this Main');

      const tip = await service.getProject(project.id);
      expect(tip.tipVersion?.id).toBe(body.id);
    } finally {
      await built.app.close();
    }
  });
});
