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
