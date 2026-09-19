const API_BASE = import.meta.env.VITE_API_BASE ?? '/api';

export interface Project {
  id: string;
  name: string;
  createdAt: string;
  mainScenarioId: string;
}

export interface Scenario {
  id: string;
  projectId: string;
  name: string;
  tipVersionId: string | null;
  isMain: boolean;
}

export interface Version {
  id: string;
  projectId: string;
  scenarioId: string;
  parentIds: string[];
  author: string;
  timestamp: string;
  message: string;
  snapshotHash: string;
  scenarioName?: string;
  promotedFromVersionId?: string;
}

export interface ProjectDetail extends Project {
  main: Scenario;
  scenarios: Scenario[];
  tipVersion: Version | null;
}

function authorHeader(): Record<string, string> {
  const author = localStorage.getItem('bit-author') || 'demo-user';
  return { 'X-Bit-Author': author };
}

export function getAuthor(): string {
  return localStorage.getItem('bit-author') || 'demo-user';
}

export function setAuthor(name: string): void {
  localStorage.setItem('bit-author', name.trim() || 'demo-user');
}

async function json<T>(res: Response): Promise<T> {
  if (!res.ok) {
    let msg = res.statusText;
    try {
      const body = await res.json();
      if (body?.error) msg = body.error;
    } catch {
      /* ignore */
    }
    throw new Error(msg);
  }
  return res.json() as Promise<T>;
}

export async function listProjects(): Promise<Project[]> {
  return json(await fetch(`${API_BASE}/projects`));
}

export async function getProject(id: string): Promise<ProjectDetail> {
  return json(await fetch(`${API_BASE}/projects/${id}`));
}

export async function listVersions(projectId: string): Promise<Version[]> {
  return json(await fetch(`${API_BASE}/projects/${projectId}/versions`));
}

export async function listScenarios(projectId: string): Promise<Scenario[]> {
  return json(await fetch(`${API_BASE}/projects/${projectId}/scenarios`));
}

export async function createScenario(opts: {
  projectId: string;
  name: string;
}): Promise<Scenario> {
  return json(
    await fetch(`${API_BASE}/projects/${opts.projectId}/scenarios`, {
      method: 'POST',
      headers: {
        ...authorHeader(),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ name: opts.name, author: getAuthor() }),
    }),
  );
}

export async function createProject(opts: {
  name: string;
  message: string;
  file: File;
}): Promise<ProjectDetail> {
  const form = new FormData();
  form.append('name', opts.name);
  form.append('message', opts.message || 'Initial version');
  form.append('author', getAuthor());
  form.append('file', opts.file);
  return json(
    await fetch(`${API_BASE}/projects`, {
      method: 'POST',
      headers: authorHeader(),
      body: form,
    }),
  );
}

export async function saveVersion(opts: {
  projectId: string;
  scenarioId?: string;
  message: string;
  file: File;
}): Promise<Version> {
  const form = new FormData();
  form.append('message', opts.message);
  form.append('author', getAuthor());
  form.append('file', opts.file);
  const url = opts.scenarioId
    ? `${API_BASE}/scenarios/${opts.scenarioId}/versions`
    : `${API_BASE}/projects/${opts.projectId}/versions`;
  return json(
    await fetch(url, {
      method: 'POST',
      headers: authorHeader(),
      body: form,
    }),
  );
}

export function downloadVersionUrl(versionId: string): string {
  return `${API_BASE}/versions/${versionId}/xlsx`;
}

export interface DiffEntry {
  kind: string;
  sheet?: string;
  address?: string;
  before?: unknown;
  after?: unknown;
}

export interface DiffResult {
  changes: DiffEntry[];
  baseId: string;
  compareId: string;
}

export async function getDiff(base: string, compare: string): Promise<DiffResult> {
  const qs = new URLSearchParams({ base, compare });
  return json(await fetch(`${API_BASE}/diff?${qs}`));
}

export type ReviewStatus = 'open' | 'changes-requested' | 'combined' | 'closed';

export interface Review {
  id: string;
  projectId: string;
  scenarioId: string;
  baseVersionId: string;
  compareVersionId: string;
  author: string;
  note?: string;
  status: ReviewStatus;
  createdAt: string;
  updatedAt: string;
}

export interface CellConflict {
  sheet: string;
  address: string;
  base: unknown;
  ours: unknown;
  theirs: unknown;
  reason: string;
}

export async function createReview(opts: {
  scenarioId: string;
  note?: string;
}): Promise<Review> {
  return json(
    await fetch(`${API_BASE}/scenarios/${opts.scenarioId}/reviews`, {
      method: 'POST',
      headers: { ...authorHeader(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ note: opts.note, author: getAuthor() }),
    }),
  );
}

export async function listReviews(projectId: string): Promise<Review[]> {
  return json(await fetch(`${API_BASE}/projects/${projectId}/reviews`));
}

export async function getReview(id: string): Promise<Review> {
  return json(await fetch(`${API_BASE}/reviews/${id}`));
}

export async function getMergePreview(opts: {
  base: string;
  ours: string;
  theirs: string;
}): Promise<{
  conflictCount: number;
  conflicts: CellConflict[];
  autoChangeCount: number;
}> {
  const qs = new URLSearchParams(opts);
  return json(await fetch(`${API_BASE}/merge?${qs}`));
}

export async function combineReview(opts: {
  reviewId: string;
  message?: string;
  resolutions?: Record<string, { action: string; cell?: unknown }>;
}): Promise<{ review: Review; version: Version }> {
  return json(
    await fetch(`${API_BASE}/reviews/${opts.reviewId}/combine`, {
      method: 'POST',
      headers: { ...authorHeader(), 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: opts.message,
        resolutions: opts.resolutions,
        author: getAuthor(),
      }),
    }),
  );
}

export async function requestReviewChanges(opts: {
  reviewId: string;
  note?: string;
}): Promise<Review> {
  return json(
    await fetch(`${API_BASE}/reviews/${opts.reviewId}/request-changes`, {
      method: 'POST',
      headers: { ...authorHeader(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ note: opts.note }),
    }),
  );
}

export async function closeReview(opts: {
  reviewId: string;
  note?: string;
}): Promise<Review> {
  return json(
    await fetch(`${API_BASE}/reviews/${opts.reviewId}/close`, {
      method: 'POST',
      headers: { ...authorHeader(), 'Content-Type': 'application/json' },
      body: JSON.stringify({ note: opts.note }),
    }),
  );
}

export interface AgentStatus {
  ok: boolean;
  dataDir: string;
  projectCount: number;
  addin: boolean;
}

export async function getAgentStatus(): Promise<AgentStatus> {
  return json(await fetch(`${API_BASE}/agent/status`));
}

export async function setDataDir(dataDir: string): Promise<AgentStatus & { dataDir: string }> {
  return json(
    await fetch(`${API_BASE}/agent/data-dir`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dataDir }),
    }),
  );
}

export async function pickFolder(): Promise<{ ok: boolean; dataDir: string }> {
  return json(
    await fetch(`${API_BASE}/agent/pick-folder`, {
      method: 'POST',
    }),
  );
}

export interface GraphNode {
  id: string;
  message: string;
  author: string;
  timestamp: string;
  scenarioId: string;
  scenarioName: string;
  isMain: boolean;
  isTip: boolean;
  parentIds: string[];
}

export interface VersionGraph {
  nodes: GraphNode[];
  edges: Array<{ from: string; to: string }>;
  scenarios: Array<{ id: string; name: string; isMain: boolean; tipVersionId: string | null }>;
}

export async function getProjectGraph(projectId: string): Promise<VersionGraph> {
  return json(await fetch(`${API_BASE}/projects/${projectId}/graph`));
}


export async function promoteVersion(opts: {
  versionId: string;
  message?: string;
  expectedMainTip?: string;
}): Promise<Version> {
  return json(
    await fetch(`${API_BASE}/versions/${opts.versionId}/promote`, {
      method: 'POST',
      headers: { ...authorHeader(), 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: opts.message,
        expectedMainTip: opts.expectedMainTip,
        author: getAuthor(),
      }),
    }),
  );
}

export async function getVersionXlsxBase64(
  versionId: string,
): Promise<{ filename: string; base64: string }> {
  return json(await fetch(`${API_BASE}/versions/${versionId}/xlsx/base64`));
}
