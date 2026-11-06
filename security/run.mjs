import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { attacks } from './attacks.mjs';
import { fuzz } from './fuzz.mjs';
import { ROOT, buildReader } from './harness.mjs';
import { abortReason, spawnChild } from './lib/spawn.mjs';
import { attackMarkdown, attackTable, fuzzMarkdown } from './report.mjs';

const args = process.argv.slice(2);
const opt = (name, dflt) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] : dflt;
};
const has = (name) => args.includes(`--${name}`);
const cmd = args.find((a) => !a.startsWith('--') && !args[args.indexOf(a) - 1]?.startsWith('--')) ?? 'all';

if (has('help')) {
  console.log(`node security/run.mjs [attacks|fuzz|all] [--seed N] [--iterations N] [--jobs N] [--mode collect|stream]
  [--only SUBSTRING] [--bomb-mb N] [--case I] [--replay FILE]
  attacks   targeted fixtures, sequential, each run in both consumer modes
  fuzz      seeded mutation fuzzer (--case I replays one case of --seed)
  --replay  run one file (.xlsx) under the budget and print the full result`);
  process.exit(0);
}

const seed = Number(opt('seed', 1));
const iterations = Number(opt('iterations', 2000));
const jobs = Number(opt('jobs', 4));
const crashDir = path.join(ROOT, 'security', 'crashes');
const bundle = buildReader();

if (has('replay')) {
  for (const mode of ['collect', 'stream']) {
    const r = await spawnChild(['run', path.resolve(opt('replay')), mode, `--bundle=${bundle}`]);
    const line = r.stdout.split('\n').find((l) => l.startsWith('RESULT '));
    console.log(mode, line ? JSON.stringify(JSON.parse(line.slice(7)), null, 2) : abortReason(r));
  }
  process.exit(0);
}

const report = [`# Reader security run`, '', `generated ${new Date().toISOString()}, budget 5000 ms and 512 MB, node ${process.version}`, ''];
let failures = 0;

if (cmd === 'attacks' || cmd === 'all') {
  const only = opt('only');
  const list = attacks({ bombMb: Number(opt('bomb-mb', 1024)) }).filter((a) => !only || a.id.includes(only));
  const rows = [];
  const tmp = path.join(ROOT, 'security', '.build');
  const bombMb = Number(opt('bomb-mb', 1024));
  for (const a of list) {
    const file = path.join(tmp, `attack-${a.id}.xlsx`);
    const built = await spawnChild(['build', a.id, file, `--bomb-mb=${bombMb}`], { timeoutMs: 300_000, heapMb: 8192 });
    const size = /BUILT (\w+)/.exec(built.stdout)?.[1];
    if (!size) {
      rows.push({ id: a.id, group: a.group, mode: '-', verdict: 'fail', outcome: 'build-error', reasons: [abortReason(built)] });
      continue;
    }
    if (size === 'null' || size === '0') {
      rows.push({ id: a.id, group: a.group, mode: '-', verdict: 'skip', outcome: 'skipped', reasons: ['input absent'] });
      continue;
    }
    for (const mode of opt('mode') ? [opt('mode')] : ['stream', 'collect']) {
      const r = await spawnChild(['run', file, mode, a.id, `--bundle=${bundle}`, `--bomb-mb=${bombMb}`], { timeoutMs: 180_000 });
      const line = r.stdout.split('\n').find((l) => l.startsWith('RESULT '));
      const res = line ? JSON.parse(line.slice(7)) : { outcome: 'abort', verdict: 'fail', reasons: [abortReason(r)] };
      rows.push({ id: a.id, group: a.group, note: a.note, mode, ...res });
      process.stdout.write(res.verdict === 'pass' ? '.' : res.verdict === 'warn' ? 'w' : 'F');
    }
  }
  process.stdout.write('\n');
  console.log(attackTable(rows));
  failures += rows.filter((r) => r.verdict === 'fail').length;
  report.push('## Attacks', '', attackMarkdown(rows), '', '### Notes', '', ...list.map((a) => `- ${a.id}: ${a.note}`), '');
  if (!has('case') && !has('only')) writeFileSync(path.join(ROOT, 'security', 'results.md'), report.join('\n'));
}

if (cmd === 'fuzz' || cmd === 'all') {
  const onlyCase = has('case') ? Number(opt('case')) : undefined;
  const s = await fuzz({ seed, iterations, jobs, bundle, mode: opt('mode', 'collect'), crashDir, onlyCase });
  const md = fuzzMarkdown(s);
  console.log(md);
  failures += s.fail;
  report.push('## Fuzz', '', `seed ${seed}, iterations ${iterations}`, '', md, '');
  if (!has('case') && !has('only')) writeFileSync(path.join(ROOT, 'security', 'results.md'), report.join('\n'));
}

console.log(`\nresults: security/results.md; failures: ${failures}`);
process.exit(failures ? 1 : 0);
