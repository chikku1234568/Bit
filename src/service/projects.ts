import type { WorkbookSnapshot } from '../xlsx/types.js';
import { parseXlsx, writeXlsx } from '../xlsx/index.js';
import type {
  Project,
  ProjectDetail,
  Scenario,
  Version,
  VersionListItem,
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
}
