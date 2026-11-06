import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { cpus, loadavg } from 'node:os';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { verify } from './verify.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const arg = (k, d) => process.argv.find((a) => a.startsWith(`--${k}=`))?.split('=')[1] ?? d;
const sizes = arg('sizes', '10000,100000').split(',').map(Number);
const arms = arg('arms', 'noop,exceljs-mem,exceljs-stream,next').split(',');
const reps = Number(arg('reps', '5'));
const heapMB = 8192;
const outDir = join(here, 'out');
mkdirSync(outDir, { recursive: true });
mkdirSync(join(here, 'results'), { recursive: true });

const med = (a) => [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)];
const frozen = createHash('sha256').update(['data.mjs', 'child.mjs', 'verify.mjs', 'run.mjs'].map((f) => readFileSync(join(here, f))).join('')).digest('hex').slice(0, 12);
console.log(`harness ${frozen} | node ${process.version} | cores ${cpus().length} | load ${loadavg().map((x) => x.toFixed(1)).join(',')} | reps ${reps} | heap ${heapMB}MB`);

const results = [];
for (const n of sizes) {
  for (let rep = 0; rep < reps; rep++) {
    const order = arms.map((_, i) => arms[(i + rep) % arms.length]); // rotate so no arm always runs first
    for (const arm of order) {
      const out = join(outDir, `${arm}-${n}.xlsx`);
      rmSync(out, { force: true });
      const p = spawnSync(process.execPath, [`--max-old-space-size=${heapMB}`, join(here, 'child.mjs'), arm, String(n), out], { encoding: 'utf8', maxBuffer: 1 << 26 });
      let r;
      try { r = JSON.parse(p.stdout.trim().split('\n').pop()); } catch { r = { arm, n, error: `child failed: ${(p.stderr || '').split('\n')[0]} (signal ${p.signal})` }; }
      r.verify = r.error || arm === 'noop' ? { ok: arm === 'noop' } : (() => { try { return verify(out, n); } catch (e) { return { ok: false, why: String(e.message) }; } })();
      r.rep = rep;
      results.push(r);
      process.stdout.write(`  n=${n} rep=${rep} ${arm.padEnd(15)} ${r.error ? 'ERROR ' + r.error : `${r.ms.toFixed(0)}ms rss ${r.rssMB.toFixed(0)}MB`}${r.verify.ok ? '' : ' VERIFY-FAIL ' + r.verify.why}\n`);
    }
  }
}

console.log('\narm              rows     median ms  range ms        cpu/wall  median rss MB  file MB  errors  verify-fail  runs');
for (const n of sizes) for (const arm of arms) {
  const rs = results.filter((r) => r.arm === arm && r.n === n);
  const ok = rs.filter((r) => !r.error);
  const errors = rs.length - ok.length;
  const vf = ok.filter((r) => !r.verify.ok).length;
  if (!ok.length) { console.log(`${arm.padEnd(16)} ${String(n).padEnd(8)} all runs errored: ${rs[0]?.error}`); continue; }
  const ms = ok.map((r) => r.ms);
  console.log(`${arm.padEnd(16)} ${String(n).padEnd(8)} ${med(ms).toFixed(0).padStart(9)}  ${Math.min(...ms).toFixed(0)}-${Math.max(...ms).toFixed(0)}`.padEnd(54) + `${(med(ok.map((r) => r.cpuMs / r.ms))).toFixed(2).padStart(8)}  ${med(ok.map((r) => r.rssMB)).toFixed(0).padStart(13)}  ${(med(ok.map((r) => r.bytes)) / 1048576).toFixed(1).padStart(7)}  ${String(errors).padStart(6)}  ${String(vf).padStart(11)}  ${ok.length}`);
}
const file = join(here, 'results', `${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
writeFileSync(file, JSON.stringify({ frozen, node: process.version, load: loadavg(), results }, null, 1));
console.log(`\nraw: ${file}`);
