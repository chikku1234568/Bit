import { useEffect, useMemo, useState } from 'react';
import { Link, useParams, useSearchParams } from 'react-router-dom';
import {
  getDiff,
  getProject,
  listVersions,
  type DiffEntry,
  type ProjectDetail,
  type Version,
} from '../api';

const KIND_LABELS: Record<string, string> = {
  'sheet-add': 'Sheet added',
  'sheet-remove': 'Sheet removed',
  'sheet-order': 'Sheet order',
  'sheet-hidden': 'Sheet hidden',
  'sheet-tab-color': 'Tab colour',
  'sheet-freeze': 'Freeze panes',
  'col-width': 'Column width',
  'row-height': 'Row height',
  'col-hidden': 'Hidden columns',
  'row-hidden': 'Hidden rows',
  merge: 'Merged cells',
  'cell-add': 'Cell added',
  'cell-remove': 'Cell removed',
  'cell-value': 'Value',
  'cell-formula': 'Formula',
  'cell-format': 'Format',
  'cell-hyperlink': 'Hyperlink',
  'cell-comment': 'Comment',
  validation: 'Validation',
  'named-range': 'Named range',
  table: 'Table',
  'auto-filter': 'AutoFilter',
};

function fmt(v: unknown): string {
  if (v === undefined) return '—';
  if (v === null) return 'null';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

export function WhatChangedPage() {
  const { id } = useParams<{ id: string }>();
  const [searchParams, setSearchParams] = useSearchParams();
  const [project, setProject] = useState<ProjectDetail | null>(null);
  const [versions, setVersions] = useState<Version[]>([]);
  const [changes, setChanges] = useState<DiffEntry[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [kindFilter, setKindFilter] = useState('all');
  const [sheetFilter, setSheetFilter] = useState('all');

  const base = searchParams.get('base') ?? '';
  const compare = searchParams.get('compare') ?? '';

  useEffect(() => {
    if (!id) return;
    void (async () => {
      try {
        const [p, v] = await Promise.all([getProject(id), listVersions(id)]);
        setProject(p);
        setVersions(v);

        let b = searchParams.get('base');
        let c = searchParams.get('compare');
        if (!b || !c) {
          const scenarioId = searchParams.get('scenario');
          const scenario =
            (scenarioId && p.scenarios.find((s) => s.id === scenarioId)) || null;
          if (scenario && !scenario.isMain && scenario.tipVersionId && p.main.tipVersionId) {
            b = p.main.tipVersionId;
            c = scenario.tipVersionId;
          } else {
            const mainVersions = v
              .filter((x) => x.scenarioId === p.mainScenarioId)
              .sort((a, x) => x.timestamp.localeCompare(a.timestamp));
            c = p.main.tipVersionId ?? mainVersions[0]?.id ?? '';
            b = mainVersions[1]?.id ?? c;
          }
          if (b && c) {
            const next = new URLSearchParams(searchParams);
            next.set('base', b);
            next.set('compare', c);
            setSearchParams(next, { replace: true });
          }
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [id]);

  useEffect(() => {
    if (!base || !compare) return;
    void (async () => {
      try {
        const r = await getDiff(base, compare);
        setChanges(r.changes);
        setError(null);
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      }
    })();
  }, [base, compare]);

  const sheets = useMemo(() => {
    const s = new Set<string>();
    for (const ch of changes) if (ch.sheet) s.add(ch.sheet);
    return [...s].sort();
  }, [changes]);

  const filtered = useMemo(() => {
    return changes.filter((ch) => {
      if (kindFilter !== 'all' && ch.kind !== kindFilter) return false;
      if (sheetFilter !== 'all' && ch.sheet !== sheetFilter) return false;
      return true;
    });
  }, [changes, kindFilter, sheetFilter]);

  function onSelect(which: 'base' | 'compare', versionId: string) {
    const next = new URLSearchParams(searchParams);
    next.set(which, versionId);
    setSearchParams(next);
  }

  if (!project && !error) return <p className="muted">Loading…</p>;

  return (
    <div className="page">
      <Link to={id ? `/projects/${id}` : '/'} className="back">
        ← Project
      </Link>
      <h1>What changed</h1>
      {project ? <p className="muted">{project.name}</p> : null}
      {error && <div className="error">{error}</div>}

      <section className="card">
        <div className="form inline">
          <label>
            Base
            <select value={base} onChange={(e) => onSelect('base', e.target.value)}>
              <option value="">Select…</option>
              {versions.map((v) => (
                <option key={v.id} value={v.id}>
                  {v.message} ({v.scenarioName ?? ''}) · {v.id.slice(0, 8)}
                </option>
              ))}
            </select>
          </label>
          <label>
            Compare
            <select value={compare} onChange={(e) => onSelect('compare', e.target.value)}>
              <option value="">Select…</option>
              {versions.map((v) => (
                <option key={v.id} value={v.id}>
                  {v.message} ({v.scenarioName ?? ''}) · {v.id.slice(0, 8)}
                </option>
              ))}
            </select>
          </label>
          <label>
            Kind
            <select value={kindFilter} onChange={(e) => setKindFilter(e.target.value)}>
              <option value="all">All</option>
              {Object.entries(KIND_LABELS).map(([k, label]) => (
                <option key={k} value={k}>
                  {label}
                </option>
              ))}
            </select>
          </label>
          <label>
            Sheet
            <select value={sheetFilter} onChange={(e) => setSheetFilter(e.target.value)}>
              <option value="all">All</option>
              {sheets.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </label>
        </div>
      </section>

      <section className="card">
        <h2>
          Changes ({filtered.length}
          {filtered.length !== changes.length ? ` of ${changes.length}` : ''})
        </h2>
        {filtered.length === 0 ? (
          <p className="muted">No differences.</p>
        ) : (
          <table className="diff-table">
            <thead>
              <tr>
                <th>Kind</th>
                <th>Sheet</th>
                <th>Cell</th>
                <th>Before</th>
                <th>After</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((ch, i) => (
                <tr key={`${ch.kind}-${ch.sheet}-${ch.address}-${i}`}>
                  <td>{KIND_LABELS[ch.kind] ?? ch.kind}</td>
                  <td>{ch.sheet ?? '—'}</td>
                  <td>{ch.address ?? '—'}</td>
                  <td>
                    <code>{fmt(ch.before)}</code>
                  </td>
                  <td>
                    <code>{fmt(ch.after)}</code>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>
    </div>
  );
}
