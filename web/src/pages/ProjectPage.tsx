import { useCallback, useEffect, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import {
  downloadVersionUrl,
  getProject,
  listVersions,
  saveVersion,
  type ProjectDetail,
  type Version,
} from '../api';

export function ProjectPage() {
  const { id } = useParams<{ id: string }>();
  const [project, setProject] = useState<ProjectDetail | null>(null);
  const [versions, setVersions] = useState<Version[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [file, setFile] = useState<File | null>(null);

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

  async function onSave(e: React.FormEvent) {
    e.preventDefault();
    if (!id || !file) {
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
      await saveVersion({ projectId: id, message: message.trim(), file });
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

  const tip = project.tipVersion;

  return (
    <div className="page">
      <Link to="/" className="back">
        ← Projects
      </Link>
      <h1>{project.name}</h1>
      <p className="muted">
        Scenario: <strong>Main</strong>
        {tip ? (
          <>
            {' '}
            · tip saved {new Date(tip.timestamp).toLocaleString()} by {tip.author}
          </>
        ) : null}
      </p>

      {error && <div className="error">{error}</div>}

      <section className="card">
        <h2>Main tip</h2>
        {tip ? (
          <div className="tip">
            <div>
              <strong>{tip.message}</strong>
              <div className="muted">
                {tip.author} · {new Date(tip.timestamp).toLocaleString()}
              </div>
            </div>
            <a className="button" href={downloadVersionUrl(tip.id)}>
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
          Edit in Excel, then upload the workbook back to Main with a note.
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
        {versions.length === 0 ? (
          <p className="muted">No versions.</p>
        ) : (
          <ul className="list versions">
            {versions.map((v) => (
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
