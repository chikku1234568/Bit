import type { WorkbookSnapshot } from '../xlsx/types.js';
import { diffSnapshots, type DiffResult } from '../diff/diff.js';
import {
  mergeSnapshots,
  applyResolutions,
  type MergeResult,
  type ResolutionChoice,
} from '../merge/merge.js';
import { parseXlsx, writeXlsx } from '../xlsx/index.js';
import type {
  Project,
  ProjectDetail,
  Scenario,
  Version,
  VersionListItem,
  Review,
} from '../domain/types.js';
import type { BitStore } from '../store/store.js';

export class NotFoundError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'NotFoundError';
  }
}

export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ValidationError';
  }
}

export class ConflictError extends Error {
  constructor(message: string, public readonly details?: unknown) {
    super(message);
    this.name = 'ConflictError';
  }
}

export class ProjectService {
  constructor(private readonly store: BitStore) {}

  async listProjects(): Promise<Project[]> {
    const meta = await this.store.readMeta();
    return [...meta.projects].sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  async getProject(id: string): Promise<ProjectDetail> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === id);
    if (!project) throw new NotFoundError(`Project not found: ${id}`);
    const scenarios = await this.listScenarios(id);
    const main = scenarios.find((s) => s.id === project.mainScenarioId);
    if (!main) throw new NotFoundError(`Main scenario missing for project ${id}`);
    const tipVersion = main.tipVersionId
      ? meta.versions.find((v) => v.id === main.tipVersionId) ?? null
      : null;
    return { ...project, main, scenarios, tipVersion };
  }


  async listScenarios(projectId: string): Promise<Scenario[]> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === projectId);
    if (!project) throw new NotFoundError(`Project not found: ${projectId}`);
    const scenarios = meta.scenarios.filter((s) => s.projectId === projectId);
    const main = scenarios.find((s) => s.id === project.mainScenarioId);
    const rest = scenarios
      .filter((s) => s.id !== project.mainScenarioId)
      .sort((a, b) => a.name.localeCompare(b.name));
    return main ? [main, ...rest] : rest;
  }

  async getScenario(scenarioId: string): Promise<Scenario> {
    const meta = await this.store.readMeta();
    const scenario = meta.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) throw new NotFoundError(`Scenario not found: ${scenarioId}`);
    return scenario;
  }

  async createScenario(opts: {
    projectId: string;
    name: string;
    author?: string;
  }): Promise<Scenario> {
    const name = opts.name.trim();
    if (!name) throw new ValidationError('Scenario name is required');
    if (name.toLowerCase() === 'main') {
      throw new ValidationError('Cannot name a scenario "Main"');
    }

    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === opts.projectId);
    if (!project) throw new NotFoundError(`Project not found: ${opts.projectId}`);

    const existing = meta.scenarios.filter((s) => s.projectId === opts.projectId);
    if (existing.some((s) => s.name.toLowerCase() === name.toLowerCase())) {
      throw new ValidationError(`Scenario name already used: ${name}`);
    }

    const main = existing.find((s) => s.id === project.mainScenarioId);
    if (!main) throw new NotFoundError(`Main scenario missing for project ${opts.projectId}`);
    if (!main.tipVersionId) {
      throw new ValidationError('Main has no tip version yet');
    }

    const scenario: Scenario = {
      id: this.store.newId(),
      projectId: opts.projectId,
      name,
      tipVersionId: main.tipVersionId, // same version id — do not duplicate blob
      isMain: false,
    };
    meta.scenarios.push(scenario);
    await this.store.writeMeta(meta);
    return scenario;
  }

  async createProject(opts: {
    name: string;
    author: string;
    message?: string;
    xlsxBuffer: Buffer;
  }): Promise<ProjectDetail> {
    const name = opts.name.trim();
    if (!name) throw new ValidationError('Project name is required');
    const author = opts.author.trim() || 'anonymous';
    const message = (opts.message ?? 'Initial version').trim() || 'Initial version';

    let snapshot: WorkbookSnapshot;
    try {
      snapshot = await parseXlsx(opts.xlsxBuffer);
    } catch (err) {
      throw new ValidationError(
        `Couldn't read this file — is it .xlsx? (${err instanceof Error ? err.message : String(err)})`,
      );
    }

    const snapshotHash = await this.store.putSnapshot(snapshot);
    const now = new Date().toISOString();
    const projectId = this.store.newId();
    const scenarioId = this.store.newId();
    const versionId = this.store.newId();

    const project: Project = {
      id: projectId,
      name,
      createdAt: now,
      mainScenarioId: scenarioId,
    };
    const scenario: Scenario = {
      id: scenarioId,
      projectId,
      name: 'Main',
      tipVersionId: versionId,
      isMain: true,
    };
    const version: Version = {
      id: versionId,
      projectId,
      scenarioId,
      parentIds: [],
      author,
      timestamp: now,
      message,
      snapshotHash,
    };

    // Cache export artefact from original upload when possible; also write from snapshot for fidelity.
    const exportBuf = await writeXlsx(snapshot);
    await this.store.putXlsxArtefact(versionId, exportBuf);

    const meta = await this.store.readMeta();
    meta.projects.push(project);
    meta.scenarios.push(scenario);
    meta.versions.push(version);
    await this.store.writeMeta(meta);

    return this.getProject(projectId);
  }

  async listVersions(projectId: string): Promise<VersionListItem[]> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === projectId);
    if (!project) throw new NotFoundError(`Project not found: ${projectId}`);
    const scenarios = new Map(
      meta.scenarios.filter((s) => s.projectId === projectId).map((s) => [s.id, s]),
    );
    return meta.versions
      .filter((v) => v.projectId === projectId)
      .map((v) => ({
        ...v,
        scenarioName: scenarios.get(v.scenarioId)?.name ?? 'Unknown',
      }))
      .sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  }

  async getProjectGraph(projectId: string): Promise<{
    nodes: Array<{
      id: string;
      message: string;
      author: string;
      timestamp: string;
      scenarioId: string;
      scenarioName: string;
      isMain: boolean;
      isTip: boolean;
      parentIds: string[];
    }>;
    edges: Array<{ from: string; to: string }>;
    scenarios: Array<{ id: string; name: string; isMain: boolean; tipVersionId: string | null }>;
  }> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === projectId);
    if (!project) throw new NotFoundError(`Project not found: ${projectId}`);
    const scenarios = meta.scenarios.filter((s) => s.projectId === projectId);
    const scenarioById = new Map(scenarios.map((s) => [s.id, s]));
    const tipIds = new Set(
      scenarios.map((s) => s.tipVersionId).filter((id): id is string => !!id),
    );
    const versions = meta.versions.filter((v) => v.projectId === projectId);
    const nodes = versions.map((v) => {
      const sc = scenarioById.get(v.scenarioId);
      return {
        id: v.id,
        message: v.message,
        author: v.author,
        timestamp: v.timestamp,
        scenarioId: v.scenarioId,
        scenarioName: sc?.name ?? 'Unknown',
        isMain: sc?.isMain ?? false,
        isTip: tipIds.has(v.id),
        parentIds: v.parentIds,
      };
    });
    const edges: Array<{ from: string; to: string }> = [];
    for (const v of versions) {
      for (const parent of v.parentIds) {
        edges.push({ from: parent, to: v.id });
      }
    }
    return {
      nodes,
      edges,
      scenarios: scenarios.map((s) => ({
        id: s.id,
        name: s.name,
        isMain: s.isMain,
        tipVersionId: s.tipVersionId,
      })),
    };
  }

  async getVersion(versionId: string): Promise<Version & { scenarioName: string }> {
    const meta = await this.store.readMeta();
    const version = meta.versions.find((v) => v.id === versionId);
    if (!version) throw new NotFoundError(`Version not found: ${versionId}`);
    const scenario = meta.scenarios.find((s) => s.id === version.scenarioId);
    return { ...version, scenarioName: scenario?.name ?? 'Unknown' };
  }

  async saveVersion(opts: {
    projectId: string;
    /** Defaults to Main scenario for M2. */
    scenarioId?: string;
    author: string;
    message: string;
    xlsxBuffer: Buffer;
  }): Promise<Version> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === opts.projectId);
    if (!project) throw new NotFoundError(`Project not found: ${opts.projectId}`);

    const scenarioId = opts.scenarioId ?? project.mainScenarioId;
    const scenario = meta.scenarios.find(
      (s) => s.id === scenarioId && s.projectId === opts.projectId,
    );
    if (!scenario) throw new NotFoundError(`Scenario not found: ${scenarioId}`);

    const author = opts.author.trim() || 'anonymous';
    const message = opts.message.trim();
    if (!message) throw new ValidationError('Version message is required');

    let snapshot: WorkbookSnapshot;
    try {
      snapshot = await parseXlsx(opts.xlsxBuffer);
    } catch (err) {
      throw new ValidationError(
        `Couldn't read this file — is it .xlsx? (${err instanceof Error ? err.message : String(err)})`,
      );
    }

    const snapshotHash = await this.store.putSnapshot(snapshot);
    const now = new Date().toISOString();
    const versionId = this.store.newId();
    const parentIds = scenario.tipVersionId ? [scenario.tipVersionId] : [];

    const version: Version = {
      id: versionId,
      projectId: opts.projectId,
      scenarioId: scenario.id,
      parentIds,
      author,
      timestamp: now,
      message,
      snapshotHash,
    };

    const exportBuf = await writeXlsx(snapshot);
    await this.store.putXlsxArtefact(versionId, exportBuf);

    scenario.tipVersionId = versionId;
    meta.versions.push(version);
    await this.store.writeMeta(meta);

    return version;
  }

  async getVersionXlsx(versionId: string): Promise<{ buffer: Buffer; filename: string }> {
    const meta = await this.store.readMeta();
    const version = meta.versions.find((v) => v.id === versionId);
    if (!version) throw new NotFoundError(`Version not found: ${versionId}`);
    const project = meta.projects.find((p) => p.id === version.projectId);
    const safeName = (project?.name ?? 'workbook').replace(/[^\w\-]+/g, '_');

    let buffer = await this.store.getXlsxArtefact(versionId);
    if (!buffer) {
      const snapshot = await this.store.getSnapshot(version.snapshotHash);
      buffer = await writeXlsx(snapshot);
      await this.store.putXlsxArtefact(versionId, buffer);
    }

    return {
      buffer,
      filename: `${safeName}-${versionId.slice(0, 8)}.xlsx`,
    };
  }

  async getVersionSnapshot(versionId: string): Promise<WorkbookSnapshot> {
    const version = await this.getVersion(versionId);
    return this.store.getSnapshot(version.snapshotHash);
  }

  async diffVersions(baseId: string, compareId: string): Promise<DiffResult & { baseId: string; compareId: string }> {
    const meta = await this.store.readMeta();
    const base = meta.versions.find((v) => v.id === baseId);
    const compare = meta.versions.find((v) => v.id === compareId);
    if (!base) throw new NotFoundError(`Version not found: ${baseId}`);
    if (!compare) throw new NotFoundError(`Version not found: ${compareId}`);
    if (base.projectId !== compare.projectId) {
      throw new ValidationError('Versions must belong to the same project');
    }
    const baseSnap = await this.store.getSnapshot(base.snapshotHash);
    const compareSnap = await this.store.getSnapshot(compare.snapshotHash);
    const result = diffSnapshots(baseSnap, compareSnap);
    return { ...result, baseId, compareId };
  }

  async previewMerge(opts: {
    baseId: string;
    oursId: string;
    theirsId: string;
  }): Promise<{
    conflictCount: number;
    conflicts: MergeResult['conflicts'];
    autoChangeCount: number;
    baseId: string;
    oursId: string;
    theirsId: string;
  }> {
    const meta = await this.store.readMeta();
    const base = meta.versions.find((v) => v.id === opts.baseId);
    const ours = meta.versions.find((v) => v.id === opts.oursId);
    const theirs = meta.versions.find((v) => v.id === opts.theirsId);
    if (!base) throw new NotFoundError(`Version not found: ${opts.baseId}`);
    if (!ours) throw new NotFoundError(`Version not found: ${opts.oursId}`);
    if (!theirs) throw new NotFoundError(`Version not found: ${opts.theirsId}`);
    const projectIds = new Set([base.projectId, ours.projectId, theirs.projectId]);
    if (projectIds.size !== 1) {
      throw new ValidationError('Versions must belong to the same project');
    }
    const [baseSnap, oursSnap, theirsSnap] = await Promise.all([
      this.store.getSnapshot(base.snapshotHash),
      this.store.getSnapshot(ours.snapshotHash),
      this.store.getSnapshot(theirs.snapshotHash),
    ]);
    const result = mergeSnapshots(baseSnap, oursSnap, theirsSnap);
    return {
      conflictCount: result.conflicts.length,
      conflicts: result.conflicts,
      autoChangeCount: result.autoChangeCount,
      baseId: opts.baseId,
      oursId: opts.oursId,
      theirsId: opts.theirsId,
    };
  }

  async createReview(opts: {
    scenarioId: string;
    author: string;
    note?: string;
  }): Promise<Review> {
    const meta = await this.store.readMeta();
    const scenario = meta.scenarios.find((s) => s.id === opts.scenarioId);
    if (!scenario) throw new NotFoundError(`Scenario not found: ${opts.scenarioId}`);
    if (scenario.isMain) {
      throw new ValidationError('Cannot ask for review from Main — create a scenario first');
    }
    const project = meta.projects.find((p) => p.id === scenario.projectId);
    if (!project) throw new NotFoundError(`Project not found: ${scenario.projectId}`);
    const main = meta.scenarios.find((s) => s.id === project.mainScenarioId);
    if (!main?.tipVersionId) throw new ValidationError('Main has no tip version');
    if (!scenario.tipVersionId) throw new ValidationError('Scenario has no tip version');

    const now = new Date().toISOString();
    const review: Review = {
      id: this.store.newId(),
      projectId: scenario.projectId,
      scenarioId: scenario.id,
      baseVersionId: main.tipVersionId,
      compareVersionId: scenario.tipVersionId,
      author: opts.author.trim() || 'anonymous',
      note: opts.note?.trim() || undefined,
      status: 'open',
      createdAt: now,
      updatedAt: now,
    };
    meta.reviews.push(review);
    await this.store.writeMeta(meta);
    return review;
  }

  async listReviews(projectId: string): Promise<Review[]> {
    const meta = await this.store.readMeta();
    const project = meta.projects.find((p) => p.id === projectId);
    if (!project) throw new NotFoundError(`Project not found: ${projectId}`);
    return meta.reviews
      .filter((r) => r.projectId === projectId)
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  async getReview(reviewId: string): Promise<Review> {
    const meta = await this.store.readMeta();
    const review = meta.reviews.find((r) => r.id === reviewId);
    if (!review) throw new NotFoundError(`Review not found: ${reviewId}`);
    return review;
  }

  async updateReviewStatus(
    reviewId: string,
    status: 'changes-requested' | 'closed',
    note?: string,
  ): Promise<Review> {
    const meta = await this.store.readMeta();
    const review = meta.reviews.find((r) => r.id === reviewId);
    if (!review) throw new NotFoundError(`Review not found: ${reviewId}`);
    if (review.status === 'combined') {
      throw new ValidationError('Review already combined');
    }
    review.status = status;
    review.updatedAt = new Date().toISOString();
    if (note?.trim()) review.note = note.trim();
    await this.store.writeMeta(meta);
    return review;
  }

  async combineReview(opts: {
    reviewId: string;
    author: string;
    message?: string;
    resolutions?: Record<string, ResolutionChoice>;
  }): Promise<{ review: Review; version: Version; merge: MergeResult }> {
    const meta = await this.store.readMeta();
    const review = meta.reviews.find((r) => r.id === opts.reviewId);
    if (!review) throw new NotFoundError(`Review not found: ${opts.reviewId}`);
    if (review.status === 'combined') {
      throw new ValidationError('Review already combined');
    }
    if (review.status === 'closed') {
      throw new ValidationError('Review is closed');
    }

    const project = meta.projects.find((p) => p.id === review.projectId);
    if (!project) throw new NotFoundError(`Project not found: ${review.projectId}`);
    const main = meta.scenarios.find((s) => s.id === project.mainScenarioId);
    if (!main) throw new NotFoundError('Main scenario missing');

    const baseVer = meta.versions.find((v) => v.id === review.baseVersionId);
    const compareVer = meta.versions.find((v) => v.id === review.compareVersionId);
    // Ours = current Main tip (or frozen base); design: Ours = Main tip, Theirs = scenario tip
    const mainTipId = main.tipVersionId ?? review.baseVersionId;
    const oursVer = meta.versions.find((v) => v.id === mainTipId) ?? baseVer;
    if (!baseVer || !compareVer || !oursVer) {
      throw new NotFoundError('Review versions missing');
    }

    const [baseSnap, oursSnap, theirsSnap] = await Promise.all([
      this.store.getSnapshot(baseVer.snapshotHash),
      this.store.getSnapshot(oursVer.snapshotHash),
      this.store.getSnapshot(compareVer.snapshotHash),
    ]);

    let merge = mergeSnapshots(baseSnap, oursSnap, theirsSnap);
    if (opts.resolutions && Object.keys(opts.resolutions).length > 0) {
      merge = applyResolutions(merge, opts.resolutions);
    }
    if (merge.conflicts.length > 0) {
      throw new ConflictError('Needs a decision on conflicting cells before combine', {
        conflictCount: merge.conflicts.length,
        conflicts: merge.conflicts,
      });
    }

    const snapshotHash = await this.store.putSnapshot(merge.snapshot);
    const now = new Date().toISOString();
    const versionId = this.store.newId();
    const version: Version = {
      id: versionId,
      projectId: review.projectId,
      scenarioId: main.id,
      parentIds: [mainTipId, review.compareVersionId],
      author: opts.author.trim() || 'anonymous',
      timestamp: now,
      message: (opts.message ?? 'Combined into Main').trim() || 'Combined into Main',
      snapshotHash,
    };

    const exportBuf = await writeXlsx(merge.snapshot);
    await this.store.putXlsxArtefact(versionId, exportBuf);

    main.tipVersionId = versionId;
    meta.versions.push(version);
    review.status = 'combined';
    review.updatedAt = now;
    await this.store.writeMeta(meta);

    return { review, version, merge };
  }
}
