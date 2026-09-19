import type { VersionGraph } from '../api';

const COLORS = [
  '#0f172a',
  '#0369a1',
  '#15803d',
  '#c2410c',
  '#7c3aed',
  '#be123c',
];

export function GraphView({
  graph,
  onMakeMain,
  mainTipId,
}: {
  graph: VersionGraph;
  onMakeMain?: (versionId: string) => void;
  mainTipId?: string | null;
}) {
  const scenarioIndex = new Map(graph.scenarios.map((s, i) => [s.id, i]));
  const sorted = [...graph.nodes].sort((a, b) => a.timestamp.localeCompare(b.timestamp));
  const colOf = (scenarioId: string) => scenarioIndex.get(scenarioId) ?? 0;
  const rowOf = new Map(sorted.map((n, i) => [n.id, i]));
  const cols = Math.max(1, graph.scenarios.length);
  const rows = Math.max(1, sorted.length);
  const cw = 140;
  const rh = 56;
  const pad = 24;
  const width = pad * 2 + cols * cw;
  const height = pad * 2 + rows * rh;

  function pos(id: string): { x: number; y: number } {
    const n = graph.nodes.find((x) => x.id === id);
    const col = n ? colOf(n.scenarioId) : 0;
    const row = rowOf.get(id) ?? 0;
    return { x: pad + col * cw + cw / 2, y: pad + row * rh + 18 };
  }

  return (
    <div className="graph-wrap">
      <svg
        className="graph-svg"
        viewBox={`0 0 ${width} ${height}`}
        width={width}
        height={Math.min(height, 360)}
      >
        {graph.edges.map((e) => {
          const a = pos(e.from);
          const b = pos(e.to);
          return (
            <line
              key={`${e.from}-${e.to}`}
              x1={a.x}
              y1={a.y}
              x2={b.x}
              y2={b.y}
              stroke="#94a3b8"
              strokeWidth={1.5}
            />
          );
        })}
        {sorted.map((n) => {
          const p = pos(n.id);
          const fill = COLORS[colOf(n.scenarioId) % COLORS.length];
          return (
            <g key={n.id}>
              <circle cx={p.x} cy={p.y} r={n.isTip ? 8 : 6} fill={fill} />
              {n.parentIds.length > 1 ? (
                <circle cx={p.x} cy={p.y} r={11} fill="none" stroke={fill} strokeWidth={1.5} />
              ) : null}
            </g>
          );
        })}
      </svg>
      <ol className="graph-legend">
        {sorted.map((n) => {
          const isMainTip = mainTipId === n.id;
          return (
            <li key={n.id}>
              <span
                className="dot"
                style={{ background: COLORS[colOf(n.scenarioId) % COLORS.length] }}
              />
              <span>
                <strong>{n.message}</strong>
                <span className="muted">
                  {' '}
                  · {n.scenarioName}
                  {n.isTip ? ' · tip' : ''}
                  {isMainTip ? ' · Main' : ''}
                  {n.parentIds.length > 1 ? ' · combined' : ''} · {n.author}
                </span>
              </span>
              {onMakeMain && !isMainTip ? (
                <button type="button" className="secondary" onClick={() => onMakeMain(n.id)}>
                  Make this Main
                </button>
              ) : null}
            </li>
          );
        })}
      </ol>
    </div>
  );
}
