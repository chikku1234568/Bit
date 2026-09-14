import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, writeFile, readdir, access } from 'node:fs/promises';
import { join } from 'node:path';
import type { WorkbookSnapshot } from '../xlsx/types.js';
import type { Project, Scenario, Version, Review } from '../domain/types.js';

export interface StorePaths {
  root: string;
  meta: string;
  blobs: string;
  xlsx: string;
}

export interface MetaDb {
  projects: Project[];
  scenarios: Scenario[];
  versions: Version[];
  reviews: Review[];
}

function emptyMeta(): MetaDb {
  return { projects: [], scenarios: [], versions: [], reviews: [] };
}

export function hashSnapshot(snapshot: WorkbookSnapshot): string {
  const json = JSON.stringify(snapshot);
  return createHash('sha256').update(json, 'utf8').digest('hex');
}

export class BitStore {
  readonly paths: StorePaths;

  constructor(rootDir: string) {
    this.paths = {
      root: rootDir,
      meta: join(rootDir, 'meta.json'),
      blobs: join(rootDir, 'blobs'),
      xlsx: join(rootDir, 'xlsx'),
    };
  }

  async init(): Promise<void> {
    await mkdir(this.paths.blobs, { recursive: true });
    await mkdir(this.paths.xlsx, { recursive: true });
    try {
      await access(this.paths.meta);
    } catch {
      await this.writeMeta(emptyMeta());
    }
  }

  async readMeta(): Promise<MetaDb> {
    const raw = await readFile(this.paths.meta, 'utf8');
    const parsed = JSON.parse(raw) as MetaDb;
    if (!parsed.reviews) parsed.reviews = [];
    return parsed;
  }

  async writeMeta(meta: MetaDb): Promise<void> {
    await writeFile(this.paths.meta, JSON.stringify(meta, null, 2), 'utf8');
  }

  async putSnapshot(snapshot: WorkbookSnapshot): Promise<string> {
    const hash = hashSnapshot(snapshot);
    const path = join(this.paths.blobs, `${hash}.json`);
    try {
      await access(path);
    } catch {
      await writeFile(path, JSON.stringify(snapshot), 'utf8');
    }
    return hash;
  }

  async getSnapshot(hash: string): Promise<WorkbookSnapshot> {
    const path = join(this.paths.blobs, `${hash}.json`);
    const raw = await readFile(path, 'utf8');
    return JSON.parse(raw) as WorkbookSnapshot;
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

  /** Test helper: wipe and re-init (used by tests with temp dirs). */
  async reset(): Promise<void> {
    await this.init();
    await this.writeMeta(emptyMeta());
  }

  async listBlobHashes(): Promise<string[]> {
    const files = await readdir(this.paths.blobs);
    return files.filter((f) => f.endsWith('.json')).map((f) => f.replace(/\.json$/, ''));
  }
}
