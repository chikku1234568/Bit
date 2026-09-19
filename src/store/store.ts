import { createHash, randomUUID } from 'node:crypto';
import {
  mkdir,
  readFile,
  writeFile,
  readdir,
  access,
  rename,
  unlink,
  open,
} from 'node:fs/promises';
import { join } from 'node:path';
import type { WorkbookSnapshot } from '../xlsx/types.js';
import type { Project, Scenario, Version, Review } from '../domain/types.js';

export interface StorePaths {
  root: string;
  /** Slim manifest: projects + scenario tip refs + reviews (no versions). */
  projectJson: string;
  /** Legacy fat meta (migrate-on-read). */
  meta: string;
  versions: string;
  blobs: string;
  xlsx: string;
  lock: string;
}

/** Slim on-disk catalog — versions live as append-only files under versions/. */
export interface ProjectCatalog {
  projects: Project[];
  scenarios: Scenario[];
  reviews: Review[];
}

/** @deprecated Compatibility shape; versions are loaded from files. */
export interface MetaDb extends ProjectCatalog {
  versions: Version[];
}

export class StoreConflictError extends Error {
  constructor(message: string, public readonly details?: unknown) {
    super(message);
    this.name = 'StoreConflictError';
  }
}

const CAS_MESSAGE = 'Project was updated — Fetch and try again.';
const DEFAULT_LOCK_MS = 15_000;

function emptyCatalog(): ProjectCatalog {
  return { projects: [], scenarios: [], reviews: [] };
}

export function hashSnapshot(snapshot: WorkbookSnapshot): string {
  const json = JSON.stringify(snapshot);
  return createHash('sha256').update(json, 'utf8').digest('hex');
}

interface LockPayload {
  holder: string;
  expiresAt: string;
}

export class BitStore {
  readonly paths: StorePaths;

  constructor(rootDir: string) {
    this.paths = {
      root: rootDir,
      projectJson: join(rootDir, 'project.json'),
      meta: join(rootDir, 'meta.json'),
      versions: join(rootDir, 'versions'),
      blobs: join(rootDir, 'blobs'),
      xlsx: join(rootDir, 'xlsx'),
      lock: join(rootDir, 'project.lock'),
    };
  }

  async init(): Promise<void> {
    await mkdir(this.paths.blobs, { recursive: true });
    await mkdir(this.paths.xlsx, { recursive: true });
    await mkdir(this.paths.versions, { recursive: true });
    await this.ensureCatalog();
  }

  private async fileExists(path: string): Promise<boolean> {
    try {
      await access(path);
      return true;
    } catch {
      return false;
    }
  }

  /** Create slim catalog, migrating fat meta.json when present. */
  private async ensureCatalog(): Promise<void> {
    if (await this.fileExists(this.paths.projectJson)) {
      return;
    }
    if (await this.fileExists(this.paths.meta)) {
      await this.migrateFatMeta();
      return;
    }
    await this.writeCatalog(emptyCatalog());
  }

  private async migrateFatMeta(): Promise<void> {
    const raw = await readFile(this.paths.meta, 'utf8');
    const parsed = JSON.parse(raw) as MetaDb;
    const versions = parsed.versions ?? [];
    for (const v of versions) {
      const path = join(this.paths.versions, `${v.id}.json`);
      if (!(await this.fileExists(path))) {
        await writeFile(path, JSON.stringify(v, null, 2), 'utf8');
      }
    }
    const catalog: ProjectCatalog = {
      projects: parsed.projects ?? [],
      scenarios: parsed.scenarios ?? [],
      reviews: parsed.reviews ?? [],
    };
    await this.writeCatalog(catalog);
    // Keep meta.json as a backup; slim project.json is now source of truth for refs.
    try {
      await rename(this.paths.meta, join(this.paths.root, 'meta.json.bak'));
    } catch {
      /* ignore if rename fails (e.g. already migrated) */
    }
  }

  async readCatalog(): Promise<ProjectCatalog> {
    await this.ensureCatalog();
    const raw = await readFile(this.paths.projectJson, 'utf8');
    const parsed = JSON.parse(raw) as ProjectCatalog;
    if (!parsed.reviews) parsed.reviews = [];
    if (!parsed.projects) parsed.projects = [];
    if (!parsed.scenarios) parsed.scenarios = [];
    return parsed;
  }

  async writeCatalog(catalog: ProjectCatalog): Promise<void> {
    const payload: ProjectCatalog = {
      projects: catalog.projects,
      scenarios: catalog.scenarios,
      reviews: catalog.reviews ?? [],
    };
    const tmp = `${this.paths.projectJson}.${randomUUID()}.tmp`;
    await writeFile(tmp, JSON.stringify(payload, null, 2), 'utf8');
    await rename(tmp, this.paths.projectJson);
  }

  /**
   * Compatibility: catalog + all version files.
   * Prefer readCatalog / getVersion / listVersionRecords.
   */
  async readMeta(): Promise<MetaDb> {
    const catalog = await this.readCatalog();
    const versions = await this.listVersionRecords();
    return { ...catalog, versions };
  }

  /** @deprecated Prefer writeCatalog + putVersion + casScenarioTip. */
  async writeMeta(meta: MetaDb): Promise<void> {
    for (const v of meta.versions ?? []) {
      const path = join(this.paths.versions, `${v.id}.json`);
      if (!(await this.fileExists(path))) {
        await writeFile(path, JSON.stringify(v, null, 2), 'utf8');
      }
    }
    await this.writeCatalog({
      projects: meta.projects,
      scenarios: meta.scenarios,
      reviews: meta.reviews ?? [],
    });
  }

  async putSnapshot(snapshot: WorkbookSnapshot): Promise<string> {
    const hash = hashSnapshot(snapshot);
    const path = join(this.paths.blobs, `${hash}.json`);
    if (!(await this.fileExists(path))) {
      await writeFile(path, JSON.stringify(snapshot), 'utf8');
    }
    return hash;
  }

  async getSnapshot(hash: string): Promise<WorkbookSnapshot> {
    const path = join(this.paths.blobs, `${hash}.json`);
    const raw = await readFile(path, 'utf8');
    return JSON.parse(raw) as WorkbookSnapshot;
  }

  /**
   * Append-only version write. Never overwrites an existing version file.
   * Uses exclusive create (wx) so a race cannot clobber.
   */
  async putVersion(version: Version): Promise<void> {
    const path = join(this.paths.versions, `${version.id}.json`);
    try {
      const fh = await open(path, 'wx');
      try {
        await fh.writeFile(JSON.stringify(version, null, 2), 'utf8');
      } finally {
        await fh.close();
      }
    } catch (err: unknown) {
      const code = (err as NodeJS.ErrnoException)?.code;
      if (code === 'EEXIST') {
        throw new StoreConflictError(`Version already exists: ${version.id}`);
      }
      throw err;
    }
  }

  async getVersion(versionId: string): Promise<Version | null> {
    const path = join(this.paths.versions, `${versionId}.json`);
    try {
      const raw = await readFile(path, 'utf8');
      return JSON.parse(raw) as Version;
    } catch {
      return null;
    }
  }

  async listVersionRecords(projectId?: string): Promise<Version[]> {
    await mkdir(this.paths.versions, { recursive: true });
    const files = await readdir(this.paths.versions);
    const versions: Version[] = [];
    for (const f of files) {
      if (!f.endsWith('.json')) continue;
      try {
        const raw = await readFile(join(this.paths.versions, f), 'utf8');
        const v = JSON.parse(raw) as Version;
        if (!projectId || v.projectId === projectId) versions.push(v);
      } catch {
        /* skip corrupt */
      }
    }
    return versions;
  }

  /**
   * Compare-and-swap a scenario tip. Reads expected tip; writes new tip only if
   * it still matches. Throws StoreConflictError (→ 409) on mismatch.
   */
  async casScenarioTip(
    scenarioId: string,
    expectedTip: string | null,
    newTip: string,
  ): Promise<Scenario> {
    return this.withLock(async () => {
      const catalog = await this.readCatalog();
      const scenario = catalog.scenarios.find((s) => s.id === scenarioId);
      if (!scenario) {
        throw new StoreConflictError(`Scenario not found: ${scenarioId}`);
      }
      if (scenario.tipVersionId !== expectedTip) {
        throw new StoreConflictError(CAS_MESSAGE, {
          expectedTip,
          actualTip: scenario.tipVersionId,
          scenarioId,
        });
      }
      scenario.tipVersionId = newTip;
      await this.writeCatalog(catalog);
      return { ...scenario };
    });
  }

  async putXlsxArtefact(versionId: string, buffer: Buffer): Promise<void> {
    await writeFile(join(this.paths.xlsx, `${versionId}.xlsx`), buffer);
  }

  async getXlsxArtefact(versionId: string): Promise<Buffer | null> {
    try {
      return await readFile(join(this.paths.xlsx, `${versionId}.xlsx`));
    } catch {
      return null;
    }
  }

  newId(): string {
    return randomUUID();
  }

  async reset(): Promise<void> {
    await this.init();
    await this.writeCatalog(emptyCatalog());
  }

  async listBlobHashes(): Promise<string[]> {
    const files = await readdir(this.paths.blobs);
    return files.filter((f) => f.endsWith('.json')).map((f) => f.replace(/\.json$/, ''));
  }

  /** Absolute path to a version metadata file (for tests / append-only checks). */
  versionPath(versionId: string): string {
    return join(this.paths.versions, `${versionId}.json`);
  }

  async withLock<T>(
    fn: () => Promise<T>,
    holder = 'bit',
    ttlMs = DEFAULT_LOCK_MS,
    acquireTimeoutMs = 5_000,
  ): Promise<T> {
    await this.acquireLock(holder, ttlMs, acquireTimeoutMs);
    try {
      return await fn();
    } finally {
      await this.releaseLock(holder);
    }
  }

  private async acquireLock(holder: string, ttlMs: number, acquireTimeoutMs = 5_000): Promise<void> {
    const deadline = Date.now() + acquireTimeoutMs;
    while (Date.now() < deadline) {
      if (await this.tryAcquireLock(holder, ttlMs)) return;
      await new Promise((r) => setTimeout(r, 20));
    }
    throw new StoreConflictError(CAS_MESSAGE, { reason: 'lock-timeout' });
  }

  private async tryAcquireLock(holder: string, ttlMs: number): Promise<boolean> {
    const now = Date.now();
    if (await this.fileExists(this.paths.lock)) {
      try {
        const raw = await readFile(this.paths.lock, 'utf8');
        const lock = JSON.parse(raw) as LockPayload;
        const expires = Date.parse(lock.expiresAt);
        if (Number.isFinite(expires) && expires > now) {
          return false; // fresh lock held by someone else
        }
        // stale — remove and retry create
        await unlink(this.paths.lock).catch(() => undefined);
      } catch {
        await unlink(this.paths.lock).catch(() => undefined);
      }
    }
    const payload: LockPayload = {
      holder,
      expiresAt: new Date(now + ttlMs).toISOString(),
    };
    try {
      const fh = await open(this.paths.lock, 'wx');
      try {
        await fh.writeFile(JSON.stringify(payload), 'utf8');
      } finally {
        await fh.close();
      }
      return true;
    } catch (err: unknown) {
      if ((err as NodeJS.ErrnoException)?.code === 'EEXIST') return false;
      throw err;
    }
  }

  private async releaseLock(holder: string): Promise<void> {
    try {
      const raw = await readFile(this.paths.lock, 'utf8');
      const lock = JSON.parse(raw) as LockPayload;
      if (lock.holder === holder) {
        await unlink(this.paths.lock);
      }
    } catch {
      /* already gone */
    }
  }
}

export { CAS_MESSAGE };
