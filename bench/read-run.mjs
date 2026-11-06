import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, statSync } from 'node:fs';
import { cpus, loadavg } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Checksum, sameResult } from './checksum.mjs';
import { HEADERS, rows } from './data.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const arg = (k, d) => process.argv.find((a) => a.startsWith(`--${k}=`))?.split('=')[1] ?? d;
const sizes = arg('sizes', '10000,100000').split(',').map(Number);
const arms = arg('arms', 'checksum-floor,exceljs-read-mem,exceljs-read-stream,next-read').split(',');
const sources = arg('sources', 'exceljs').split(',');
const reps = Number(arg('reps', '5'));
const med = (a) => [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)];
const frozen = createHash('sha256').update(['data.mjs', 'checksum.mjs', 'read-child.mjs', 'read-run.mjs'].map((f) => readFileSync(join(here, f))).join('')).digest('hex').slice(0, 12);
console.log(`read harness ${frozen} | node ${process.version} | cores ${cpus().length} | load ${loadavg().map((x) => x.toFixed(1)).join(',')} | reps ${reps}`);

const results = [];
for (const n of sizes) {
  const exp = new Checksum();
  exp.row(HEADERS);
  for (const r of rows(n)) exp.row(r);
  const want = exp.result();
  for (const src of sources) {
    const path = join(here, 'inputs', `${src}-${n}.xlsx`);
    for (let rep = 0; rep < reps; rep++) {
      for (let k = 0; k < arms.length; k++) {
        const arm = arms[(k + rep) % arms.length];
        const p = spawnSync(process.execPath, ['--max-old-space-size=8192', join(here, 'read-child.mjs'), arm, path, String(n)], { encoding: 'utf8', maxBuffer: 1 << 26 });
        let r;
        try { r = JSON.parse(p.stdout.trim().split('\n').pop()); } catch { r = { arm, error: `child failed: ${(p.stderr || '').split('\n')[0]} (signal ${p.signal})` }; }
        r.n = n; r.src = src; r.rep = rep;
        r.bad = r.error ? [] : sameResult(r.result, want);
        results.push(r);
        process.stdout.write(`  n=${n} ${src} rep=${rep} ${arm.padEnd(20)} ${r.error ? 'ERROR ' + r.error : `${r.ms.toFixed(0)}ms rss ${r.rssMB.toFixed(0)}MB`}${r.bad.length ? ' WRONG ' + r.bad.join('; ') : ''}\n`);
      }
    }
  }
}
console.log('\narm                  source    rows     median ms  range ms        cpu/wall  median rss MB  errors  wrong  runs');
for (const n of sizes) for (const src of sources) for (const arm of arms) {
  const rs = results.filter((r) => r.arm === arm && r.n === n && r.src === src);
  const ok = rs.filter((r) => !r.error);
  const wrong = ok.filter((r) => r.bad.length).length;
  if (!ok.length) { console.log(`${arm.padEnd(20)} ${src.padEnd(9)} ${String(n).padEnd(8)} all errored: ${rs[0]?.error}`); continue; }
  const ms = ok.map((r) => r.ms);
  console.log(`${arm.padEnd(20)} ${src.padEnd(9)} ${String(n).padEnd(8)} ${med(ms).toFixed(0).padStart(9)}  ${(Math.min(...ms).toFixed(0) + '-' + Math.max(...ms).toFixed(0)).padEnd(14)}  ${med(ok.map((r) => r.cpuMs / r.ms)).toFixed(2).padStart(8)}  ${med(ok.map((r) => r.rssMB)).toFixed(0).padStart(13)}  ${String(rs.length - ok.length).padStart(6)}  ${String(wrong).padStart(5)}  ${ok.length}`);
}
const file = join(here, 'results', `read-${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
writeFileSync(file, JSON.stringify({ frozen, results }, null, 1));
console.log(`\nraw: ${file}`);
