import { useCallback, useEffect, useMemo, useState } from 'react';
import { Link, useNavigate, useParams, useSearchParams } from 'react-router-dom';
import {
  createScenario,
  downloadVersionUrl,
  getProject,
  listVersions,
  saveVersion,
  type ProjectDetail,
  type Scenario,
  type Version,
} from '../api';

export function ProjectPage() {
  const { id } = useParams<{ id: string }>();
  const [searchParams] = useSearchParams();
  const navigate = useNavigate();
  const [project, setProject] = useState<ProjectDetail | null>(null);
  const [versions, setVersions] = useState<Version[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [newScenarioName, setNewScenarioName] = useState('');

  const refresh = useCallback(async () => {
    if (!id) return;
    try {
      const [p, v] = await Promise.all([getProject(id), listVersions(id)]);
      setProject(p);
      setVersions(v);
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }, [id]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const selectedScenario: Scenario | null = useMemo(() => {
    if (!project) return null;
    const q = searchParams.get('scenario');
    if (q) {
      const found = project.scenarios.find((s) => s.id === q);
      if (found) return found;
    }
    return project.main;
  }, [project, searchParams]);

  const selectedTipId = selectedScenario?.tipVersionId ?? null;
  const selectedTip = versions.find((v) => v.id === selectedTipId) ?? null;

  const filteredVersions = useMemo(() => {
    if (!selectedScenario) return versions;
    // Versions saved on this scenario, plus the shared tip when still pointing at Main.
    return versions.filter(
      (v) =>
        v.scenarioId === selectedScenario.id ||
        v.id === selectedScenario.tipVersionId,
    );
  }, [versions, selectedScenario]);

  function switchScenario(scenarioId: string) {
    if (!id) return;
    navigate(`/projects/${id}?scenario=${scenarioId}`);
  }

  async function onCreateScenario(e: React.FormEvent) {
    e.preventDefault();
    if (!id || !newScenarioName.trim()) {
      setError('Enter a scenario name');
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const s = await createScenario({ projectId: id, name: newScenarioName.trim() });
      setNewScenarioName('');
      await refresh();
      navigate(`/projects/${id}?scenario=${s.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onSave(e: React.FormEvent) {
    e.preventDefault();
    if (!id || !file || !selectedScenario) {
      setError('Choose an .xlsx file to save a version');
      return;
    }
    if (!message.trim()) {
      setError('Add a short note for this version');
      return;
    }
    setBusy(true);
    setError(null);
    try {
      await saveVersion({
        projectId: id,
        scenarioId: selectedScenario.id,
        message: message.trim(),
        file,
      });
      setMessage('');
      setFile(null);
      await refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  if (!project && !error) return <p className="muted">Loading…</p>;
  if (!project) {
    return (
      <div className="page">
        <Link to="/">← Projects</Link>
        <div className="error">{error}</div>
      </div>
    );
  }

  return (
    <div className="page">
      <Link to="/" className="back">
        ← Projects
      </Link>
      <h1>{project.name}</h1>
      <p className="muted">
        Scenario: <strong>{selectedScenario?.name ?? 'Main'}</strong>
        {selectedTip ? (
          <>
            {' '}
            · tip saved {new Date(selectedTip.timestamp).toLocaleString()} by{' '}
            {selectedTip.author}
          </>
        ) : null}
      </p>

      {error && <div className="error">{error}</div>}

      <section className="card">
        <h2>Scenarios</h2>
        <ul className="list">
          {project.scenarios.map((s) => (
            <li key={s.id}>
              <div>
                <strong>{s.name}</strong>
                {s.isMain ? <span className="muted"> · default</span> : null}
              </div>
              <button
                type="button"
                className="button secondary"
                disabled={selectedScenario?.id === s.id}
                onClick={() => switchScenario(s.id)}
              >
                {selectedScenario?.id === s.id ? 'Current' : 'Switch scenario'}
              </button>
            </li>
          ))}
        </ul>
        <form onSubmit={onCreateScenario} className="form inline">
          <label>
            New scenario
            <input
              value={newScenarioName}
              onChange={(e) => setNewScenarioName(e.target.value)}
              placeholder="Tax update"
            />
          </label>
          <button type="submit" disabled={busy || !newScenarioName.trim()}>
            Create scenario
          </button>
        </form>
      </section>

      <section className="card">
        <h2>{selectedScenario?.isMain ? 'Main tip' : `${selectedScenario?.name} tip`}</h2>
        {selectedTip ? (
          <div className="tip">
            <div>
              <strong>{selectedTip.message}</strong>
              <div className="muted">
                {selectedTip.author} · {new Date(selectedTip.timestamp).toLocaleString()}
              </div>
            </div>
            <a className="button" href={downloadVersionUrl(selectedTip.id)}>
              Download .xlsx
            </a>
          </div>
        ) : selectedTipId ? (
          <div className="tip">
            <p className="muted">Tip version {selectedTipId.slice(0, 8)}…</p>
            <a className="button" href={downloadVersionUrl(selectedTipId)}>
              Download .xlsx
            </a>
          </div>
        ) : (
          <p className="muted">No versions yet.</p>
        )}
      </section>

      <section className="card">
        <h2>Save version</h2>
        <p className="muted">
          Edit in Excel, then upload the workbook back to{' '}
          <strong>{selectedScenario?.name ?? 'Main'}</strong> with a note.
        </p>
        <form onSubmit={onSave} className="form">
          <label>
            Note
            <input
              value={message}
              onChange={(e) => setMessage(e.target.value)}
              placeholder="Updated tax rate"
              required
            />
          </label>
          <label>
            Workbook (.xlsx)
            <input
              type="file"
              accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={(e) => setFile(e.target.files?.[0] ?? null)}
            />
          </label>
          <button type="submit" disabled={busy || !file}>
            {busy ? 'Saving…' : 'Save version'}
          </button>
        </form>
      </section>

      <section className="card">
        <h2>Version history</h2>
        <p className="muted">Showing versions for {selectedScenario?.name ?? 'Main'}.</p>
        {filteredVersions.length === 0 ? (
          <p className="muted">No versions.</p>
        ) : (
          <ul className="list versions">
            {filteredVersions.map((v) => (
              <li key={v.id}>
                <div>
                  <strong>{v.message}</strong>
                  <div className="muted">
                    {v.author} · {v.scenarioName ?? 'Main'} ·{' '}
                    {new Date(v.timestamp).toLocaleString()}
                  </div>
                </div>
                <a className="button secondary" href={downloadVersionUrl(v.id)}>
                  Download .xlsx
                </a>
              </li>
            ))}
          </ul>
        )}
      </section>
    </div>
  );
}
