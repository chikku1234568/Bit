import { useEffect, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { createProject, listProjects, type Project } from '../api';

export function HomePage() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [name, setName] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const navigate = useNavigate();

  async function refresh() {
    try {
      setProjects(await listProjects());
      setError(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }

  useEffect(() => {
    void refresh();
  }, []);

  async function onCreate(e: React.FormEvent) {
    e.preventDefault();
    if (!file) {
      setError('Choose an .xlsx file');
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const project = await createProject({
        name: name || file.name.replace(/\.xlsx$/i, ''),
        message: 'Initial version',
        file,
      });
      navigate(`/projects/${project.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="page">
      <h1>Projects</h1>
      <p className="muted">One workbook = one Bit project. Upload an .xlsx to start.</p>

      {error && <div className="error">{error}</div>}

      <section className="card">
        <h2>Create project</h2>
        <form onSubmit={onCreate} className="form">
          <label>
            Name
            <input
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="FY27 Budget"
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
            {busy ? 'Creating…' : 'Create project'}
          </button>
        </form>
      </section>

      <section className="card">
        <h2>Your projects</h2>
        {projects.length === 0 ? (
          <p className="muted">No projects yet.</p>
        ) : (
          <ul className="list">
            {projects.map((p) => (
              <li key={p.id}>
                <Link to={`/projects/${p.id}`}>{p.name}</Link>
                <span className="muted">
                  {new Date(p.createdAt).toLocaleString()}
                </span>
              </li>
            ))}
          </ul>
        )}
      </section>
    </div>
  );
}
