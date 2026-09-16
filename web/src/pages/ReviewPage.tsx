import { useEffect, useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import {
  closeReview,
  combineReview,
  downloadVersionUrl,
  getDiff,
  getMergePreview,
  getProject,
  getReview,
  requestReviewChanges,
  type CellConflict,
  type DiffEntry,
  type ProjectDetail,
  type Review,
  type Version,
} from '../api';

type Choice = 'keep-ours' | 'keep-theirs' | 'edit';

export function ReviewPage() {
  const { reviewId } = useParams<{ reviewId: string }>();
  const navigate = useNavigate();
  const [review, setReview] = useState<Review | null>(null);
  const [project, setProject] = useState<ProjectDetail | null>(null);
  const [changes, setChanges] = useState<DiffEntry[]>([]);
  const [conflicts, setConflicts] = useState<CellConflict[]>([]);
  const [autoChangeCount, setAutoChangeCount] = useState(0);
  const [choices, setChoices] = useState<Record<string, Choice>>({});
  const [edits, setEdits] = useState<Record<string, string>>({});
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [combinedVersion, setCombinedVersion] = useState<Version | null>(null);

  useEffect(() => {
    if (!reviewId) return;
    void (async () => {
      try {
        const r = await getReview(reviewId);
        setReview(r);
        const p = await getProject(r.projectId);
        setProject(p);
        const diff = await getDiff(r.baseVersionId, r.compareVersionId);
        setChanges(diff.changes);
        const ours = p.main.tipVersionId ?? r.baseVersionId;
        const preview = await getMergePreview({
          base: r.baseVersionId,
          ours,
          theirs: r.compareVersionId,
        });
        setConflicts(preview.conflicts);
        setAutoChangeCount(preview.autoChangeCount);
        const init: Record<string, Choice> = {};
        for (const c of preview.conflicts) {
          init[`${c.sheet}!${c.address}`] = 'keep-ours';
        }
        setChoices(init);
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      }
    })();
  }, [reviewId]);

  const scenarioName = useMemo(() => {
    if (!project || !review) return '';
    return project.scenarios.find((s) => s.id === review.scenarioId)?.name ?? 'Scenario';
  }, [project, review]);

  async function onCombine() {
    if (!review) return;
    setBusy(true);
    setError(null);
    try {
      const resolutions: Record<string, { action: string; cell?: { v: string | number | null; f: null } }> =
        {};
      for (const c of conflicts) {
        const key = `${c.sheet}!${c.address}`;
        const choice = choices[key] ?? 'keep-ours';
        if (choice === 'edit') {
          const raw = edits[key] ?? '';
          const num = Number(raw);
          resolutions[key] = {
            action: 'edit',
            cell: { v: raw === '' ? null : Number.isFinite(num) && raw.trim() !== '' ? num : raw, f: null },
          };
        } else {
          resolutions[key] = { action: choice };
        }
      }
      const result = await combineReview({
        reviewId: review.id,
        message: `Combined ${scenarioName} into Main`,
        resolutions: conflicts.length ? resolutions : undefined,
      });
      setCombinedVersion(result.version);
      setReview(result.review);
      setConflicts([]);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onRequestChanges() {
    if (!review) return;
    setBusy(true);
    try {
      const r = await requestReviewChanges({ reviewId: review.id, note: 'Please revise' });
      setReview(r);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  async function onClose() {
    if (!review) return;
    setBusy(true);
    try {
      const r = await closeReview({ reviewId: review.id });
      setReview(r);
      navigate(`/projects/${r.projectId}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  if (!review && !error) return <p className="muted">Loading…</p>;
  if (!review) {
    return (
      <div className="page">
        <div className="error">{error}</div>
      </div>
    );
  }

  return (
    <div className="page">
      <Link to={`/projects/${review.projectId}`} className="back">
        ← Project
      </Link>
      <h1>Review</h1>
      <p className="muted">
        {scenarioName} → Main · {review.status}
        {review.note ? ` · “${review.note}”` : ''}
      </p>
      {error && <div className="error">{error}</div>}

      {combinedVersion ? (
        <section className="card">
          <h2>Combined into Main</h2>
          <p>
            New Main version: <strong>{combinedVersion.message}</strong>
          </p>
          <a className="button" href={downloadVersionUrl(combinedVersion.id)}>
            Download .xlsx
          </a>
        </section>
      ) : null}

      <section className="card">
        <h2>Summary</h2>
        <p>
          What changed: <strong>{changes.length}</strong> · Auto-merged:{' '}
          <strong>{autoChangeCount}</strong> · Needs a decision:{' '}
          <strong>{conflicts.length}</strong>
        </p>
        <p>
          <Link
            to={`/projects/${review.projectId}/what-changed?base=${review.baseVersionId}&compare=${review.compareVersionId}`}
          >
            What changed
          </Link>
        </p>
      </section>

      {conflicts.length > 0 && review.status !== 'combined' ? (
        <section className="card">
          <h2>Needs a decision</h2>
          <table className="diff-table">
            <thead>
              <tr>
                <th>Cell</th>
                <th>Main</th>
                <th>Scenario</th>
                <th>Decision</th>
              </tr>
            </thead>
            <tbody>
              {conflicts.map((c) => {
                const key = `${c.sheet}!${c.address}`;
                return (
                  <tr key={key}>
                    <td>
                      {c.sheet}!{c.address}
                    </td>
                    <td>
                      <code>{JSON.stringify(c.ours)}</code>
                    </td>
                    <td>
                      <code>{JSON.stringify(c.theirs)}</code>
                    </td>
                    <td>
                      <select
                        value={choices[key] ?? 'keep-ours'}
                        onChange={(e) =>
                          setChoices((prev) => ({
                            ...prev,
                            [key]: e.target.value as Choice,
                          }))
                        }
                      >
                        <option value="keep-ours">Keep Main</option>
                        <option value="keep-theirs">Keep scenario</option>
                        <option value="edit">Edit</option>
                      </select>
                      {choices[key] === 'edit' ? (
                        <input
                          placeholder="Value"
                          value={edits[key] ?? ''}
                          onChange={(e) =>
                            setEdits((prev) => ({ ...prev, [key]: e.target.value }))
                          }
                        />
                      ) : null}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </section>
      ) : null}

      {review.status !== 'combined' && review.status !== 'closed' ? (
        <section className="card actions">
          <button type="button" className="button" disabled={busy} onClick={() => void onCombine()}>
            Combine into Main
          </button>
          <button
            type="button"
            className="button secondary"
            disabled={busy}
            onClick={() => void onRequestChanges()}
          >
            Request changes
          </button>
          <button
            type="button"
            className="button secondary"
            disabled={busy}
            onClick={() => void onClose()}
          >
            Close
          </button>
        </section>
      ) : null}
    </div>
  );
}
