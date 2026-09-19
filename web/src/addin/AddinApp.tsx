import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  createProject,
  createScenario,
  getAgentStatus,
  getAuthor,
  getDiff,
  getProject,
  getProjectGraph,
  getVersionXlsxBase64,
  listProjects,
  listScenarios,
  listVersions,
  pickFolder,
  saveVersion,
  setAuthor,
  setDataDir,
  type DiffEntry,
  type Project,
  type ProjectDetail,
  type Scenario,
  type Version,
  type VersionGraph,
} from '../api';
import { GraphView } from './GraphView';
import { getOpenWorkbookFile, isInExcel, openWorkbookFromBase64, waitOfficeReady } from './office';

type Tab = 'home' | 'history' | 'graph' | 'changed';

export function AddinApp() {
  const [ready, setReady] = useState(false);
  const [inExcel, setInExcel] = useState(false);
  const [agentOk, setAgentOk] = useState(false);
  const [dataDir, setDataDirState] = useState('');
  const [folderDraft, setFolderDraft] = useState('');
  const [author, setAuthorState] = useState(getAuthor());
  const [projects, setProjects] = useState<Project[]>([]);
  const [project, setProject] = useState<ProjectDetail | null>(null);
  const [scenarios, setScenarios] = useState<Scenario[]>([]);
  const [scenarioId, setScenarioId] = useState<string | null>(null);
  const [versions, setVersions] = useState<Version[]>([]);
  const [graph, setGraph] = useState<VersionGraph | null>(null);
  const [changes, setChanges] = useState<DiffEntry[]>([]);
  const [note, setNote] = useState('');
  const [newScenario, setNewScenario] = useState('');
  const [newProjectName, setNewProjectName] = useState('');
  const [tab, setTab] = useState<Tab>('home');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setAuthor(author);
  }, [author]);

  const selected = useMemo(
    () => scenarios.find((s) => s.id === scenarioId) ?? scenarios.find((s) => s.isMain) ?? null,
    [scenarios, scenarioId],
  );

  const refreshAgent = useCallback(async () => {
    const s = await getAgentStatus();
    setAgentOk(s.ok);
    setDataDirState(s.dataDir);
    setFolderDraft(s.dataDir);
    return s;
  }, []);

  const refreshProject = useCallback(async (id: string) => {
    const [p, sc, vs] = await Promise.all([getProject(id), listScenarios(id), listVersions(id)]);
    setProject(p);
    setScenarios(sc);
    setVersions(vs);
    setScenarioId((cur) => cur ?? p.mainScenarioId);
    return p;
  }, []);

  const boot = useCallback(async () => {
    setError(null);
    try {
      await waitOfficeReady();
      setInExcel(isInExcel());
      await refreshAgent();
      const list = await listProjects();
      setProjects(list);
      const stored = localStorage.getItem('bit-addin-project');
      const pick = list.find((p) => p.id === stored) ?? list[0];
      if (pick) {
        localStorage.setItem('bit-addin-project', pick.id);
        await refreshProject(pick.id);
      }
      setReady(true);
    } catch (err) {
      setAgentOk(false);
      setError(err instanceof Error ? err.message : String(err));
      setReady(true);
    }
  }, [refreshAgent, refreshProject]);

  useEffect(() => {
    void boot();
  }, [boot]);

  async function onPickFolder() {
    setBusy(true);
    setError(null);
    try {
      const r = await pickFolder();
      setDataDirState(r.dataDir);
      setFolderDraft(r.dataDir);
      const list = await listProjects();
      setProjects(list);
      setProject(null);
      setScenarios([]);
      setVersions([]);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onSetFolder() {
    setBusy(true);
    setError(null);
    try {
      const r = await setDataDir(folderDraft);
      setDataDirState(r.dataDir);
      const list = await listProjects();
      setProjects(list);
      setProject(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onSelectProject(id: string) {
    localStorage.setItem('bit-addin-project', id);
    setGraph(null);
    setChanges([]);
    await refreshProject(id);
  }

  async function onCreateProject() {
    setBusy(true);
    setError(null);
    try {
      const file = await getOpenWorkbookFile();
      const created = await createProject({
        name: newProjectName.trim() || file.name.replace(/\.xlsx$/i, ''),
        message: 'Initial version',
        file,
      });
      setNewProjectName('');
      localStorage.setItem('bit-addin-project', created.id);
      const list = await listProjects();
      setProjects(list);
      await refreshProject(created.id);
      setTab('home');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onSave() {
    if (!project || !selected) return;
    if (!note.trim()) {
      setError('Add a short note for this version');
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const file = await getOpenWorkbookFile();
      await saveVersion({
        projectId: project.id,
        scenarioId: selected.id,
        message: note.trim(),
        file,
      });
      setNote('');
      await refreshProject(project.id);
      setTab('history');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onCreateScenario(e: React.FormEvent) {
    e.preventDefault();
    if (!project || !newScenario.trim()) return;
    setBusy(true);
    setError(null);
    try {
      const s = await createScenario({ projectId: project.id, name: newScenario.trim() });
      setNewScenario('');
      await refreshProject(project.id);
      setScenarioId(s.id);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onOpenVersion(id: string) {
    setBusy(true);
    setError(null);
    try {
      const { base64 } = await getVersionXlsxBase64(id);
      await openWorkbookFromBase64(base64);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function loadGraph() {
    if (!project) return;
    setBusy(true);
    setError(null);
    try {
      setGraph(await getProjectGraph(project.id));
      setTab('graph');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function loadChanged() {
    if (!project || !selected?.tipVersionId) return;
    const mainTip = project.main.tipVersionId;
    const compare = selected.tipVersionId;
    const base =
      !selected.isMain && mainTip && mainTip !== compare
        ? mainTip
        : versions.find((v) => v.id !== compare)?.id ?? compare;
    setBusy(true);
    setError(null);
    try {
      const diff = await getDiff(base, compare);
      setChanges(diff.changes);
      setTab('changed');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  if (!ready) {
    return <p className="addin-muted">Starting Bit…</p>;
  }

  return (
    <div className="addin">
      <header className="addin-head">
        <strong>Bit</strong>
        <span className={agentOk ? 'pill ok' : 'pill bad'}>
          {agentOk ? 'agent' : 'offline'}
        </span>
      </header>

      {!inExcel ? (
        <p className="banner">Sideload this page in Excel. Browser preview cannot read the workbook.</p>
      ) : null}

      {error ? <div className="addin-error">{error}</div> : null}

      <section className="block">
        <label>
          You are
          <input value={author} onChange={(e) => setAuthorState(e.target.value)} />
        </label>
        <label>
          Storage folder
          <input value={folderDraft} onChange={(e) => setFolderDraft(e.target.value)} />
        </label>
        <div className="row">
          <button type="button" disabled={busy} onClick={() => void onPickFolder()}>
            Choose folder
          </button>
          <button type="button" className="secondary" disabled={busy} onClick={() => void onSetFolder()}>
            Use path
          </button>
        </div>
        <p className="addin-muted path">{dataDir}</p>
      </section>

      <section className="block">
        <label>
          Project
          <select
            value={project?.id ?? ''}
            onChange={(e) => {
              if (e.target.value) void onSelectProject(e.target.value);
            }}
          >
            <option value="">— none —</option>
            {projects.map((p) => (
              <option key={p.id} value={p.id}>
                {p.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          New from this workbook
          <input
            value={newProjectName}
            onChange={(e) => setNewProjectName(e.target.value)}
            placeholder="FY27 Budget"
          />
        </label>
        <button type="button" disabled={busy || !inExcel} onClick={() => void onCreateProject()}>
          Create project
        </button>
      </section>

      {project ? (
        <>
          <section className="block">
            <p className="addin-muted">
              Scenario:{' '}
              <strong>{selected?.name ?? 'Main'}</strong>
            </p>
            <div className="chips">
              {scenarios.map((s) => (
                <button
                  key={s.id}
                  type="button"
                  className={s.id === selected?.id ? 'chip on' : 'chip'}
                  onClick={() => setScenarioId(s.id)}
                >
                  {s.name}
                </button>
              ))}
            </div>
            <form onSubmit={(e) => void onCreateScenario(e)} className="row">
              <input
                value={newScenario}
                onChange={(e) => setNewScenario(e.target.value)}
                placeholder="New scenario"
              />
              <button type="submit" disabled={busy || !newScenario.trim()}>
                Add
              </button>
            </form>
          </section>

          <section className="block">
            <label>
              Note
              <input
                value={note}
                onChange={(e) => setNote(e.target.value)}
                placeholder="Updated tax rate"
              />
            </label>
            <button type="button" disabled={busy || !inExcel} onClick={() => void onSave()}>
              {busy ? 'Working…' : 'Save version'}
            </button>
          </section>

          <nav className="tabs">
            <button type="button" className={tab === 'home' ? 'on' : ''} onClick={() => setTab('home')}>
              Home
            </button>
            <button
              type="button"
              className={tab === 'history' ? 'on' : ''}
              onClick={() => setTab('history')}
            >
              History
            </button>
            <button type="button" className={tab === 'graph' ? 'on' : ''} onClick={() => void loadGraph()}>
              Graph
            </button>
            <button
              type="button"
              className={tab === 'changed' ? 'on' : ''}
              onClick={() => void loadChanged()}
            >
              Changed
            </button>
          </nav>

          {tab === 'home' ? (
            <p className="addin-muted">
              Edit in Excel, then <strong>Save version</strong>. Switch scenario to branch. Graph
              shows the version tree.
            </p>
          ) : null}

          {tab === 'history' ? (
            <ul className="hist">
              {versions
                .filter((v) => !selected || v.scenarioId === selected.id || v.id === selected.tipVersionId)
                .map((v) => (
                  <li key={v.id}>
                    <div>
                      <strong>{v.message}</strong>
                      <div className="addin-muted">
                        {v.scenarioName ?? ''} · {v.author} ·{' '}
                        {new Date(v.timestamp).toLocaleString()}
                      </div>
                    </div>
                    <button type="button" className="secondary" onClick={() => void onOpenVersion(v.id)}>
                      Open
                    </button>
                  </li>
                ))}
            </ul>
          ) : null}

          {tab === 'graph' && graph ? <GraphView graph={graph} /> : null}

          {tab === 'changed' ? (
            <ul className="hist">
              {changes.length === 0 ? (
                <li className="addin-muted">No tracked changes.</li>
              ) : (
                changes.slice(0, 80).map((c, i) => (
                  <li key={i}>
                    <strong>{c.kind}</strong>{' '}
                    <span className="addin-muted">
                      {c.sheet}
                      {c.address ? `!${c.address}` : ''}
                    </span>
                  </li>
                ))
              )}
            </ul>
          ) : null}
        </>
      ) : (
        <p className="addin-muted">Create a project from this workbook to start.</p>
      )}
    </div>
  );
}
