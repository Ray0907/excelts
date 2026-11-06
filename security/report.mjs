const pad = (s, n) => String(s).padEnd(n);
const fmt = (n, d = 0) => (n === undefined || Number.isNaN(n) ? '-' : n.toFixed(d));

export function attackTable(rows) {
  const head = ['attack', 'mode', 'verdict', 'outcome', 'time ms', 'peak MB', 'rss MB', 'inflated MB', 'rows', 'detail'];
  const body = rows.map((x) => [x.id, x.mode, x.verdict.toUpperCase(), x.outcome, fmt(x.wallMs), fmt(x.peakMb), fmt(x.rssMb), fmt(x.inflateOutMb), x.rows ?? '-', (x.reasons?.join('; ') || x.err?.message || '').slice(0, 70)]);
  const w = head.map((h, i) => Math.max(h.length, ...body.map((r) => String(r[i]).length)));
  const line = (r) => r.map((c, i) => pad(c, w[i])).join('  ');
  return [line(head), line(w.map((n) => '-'.repeat(n))), ...body.map(line)].join('\n');
}

export function attackMarkdown(rows) {
  const esc = (s) => String(s).replace(/\|/g, '\\|');
  const out = ['| attack | mode | verdict | outcome | time ms | peak MB | rss MB | inflated MB | rows | detail |', '|---|---|---|---|---:|---:|---:|---:|---:|---|'];
  for (const x of rows) out.push(`| ${x.id} | ${x.mode} | ${x.verdict} | ${x.outcome} | ${fmt(x.wallMs)} | ${fmt(x.peakMb)} | ${fmt(x.rssMb)} | ${fmt(x.inflateOutMb)} | ${x.rows ?? '-'} | ${esc((x.reasons?.join('; ') || x.err?.message || '').slice(0, 120))} |`);
  return out.join('\n');
}

export function fuzzMarkdown(s) {
  const ms = [...s.ms].sort((a, b) => a - b);
  const q = (p) => ms[Math.min(ms.length - 1, Math.floor(ms.length * p))] ?? 0;
  const out = [`cases ${s.cases}, pass ${s.pass}, warn ${s.warn}, fail ${s.fail} (process aborts ${s.aborts}), seeds ${s.seeds}, ${s.seconds.toFixed(1)} s, wall ms p50 ${q(0.5).toFixed(0)} p99 ${q(0.99).toFixed(0)} max ${q(1).toFixed(0)}`, `outcomes ${JSON.stringify(s.byOutcome)}`, ''];
  out.push('| count | signature | worst ms | worst MB | example (replay with --seed S --case I) | saved |', '|---:|---|---:|---:|---|---|');
  for (const [sig, g] of [...s.groups.entries()].sort((a, b) => b[1].count - a[1].count)) out.push(`| ${g.count} | ${sig.replace(/\|/g, '/')} | ${fmt(g.worst.wallMs)} | ${fmt(g.worst.memMb)} | case ${g.example.index} | ${g.hashes.join(', ')} |`);
  return out.join('\n');
}
