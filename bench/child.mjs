import { statSync } from 'node:fs';
import { performance } from 'node:perf_hooks';
import { rows } from './data.mjs';

const [arm, nStr, out] = process.argv.slice(2);
const n = Number(nStr);
const { default: write } = await import(`./arms/${arm}.mjs`);
const cpu0 = process.cpuUsage();
const t0 = performance.now();
let error = null;
try {
  await write(rows(n), out);
} catch (e) {
  error = String(e && e.stack ? e.stack.split('\n')[0] : e);
}
const ms = performance.now() - t0;
const cpu = process.cpuUsage(cpu0);
let bytes = 0;
if (arm !== 'noop' && !error) bytes = statSync(out).size;
console.log(JSON.stringify({ arm, n, ms, cpuMs: (cpu.user + cpu.system) / 1000, rssMB: process.resourceUsage().maxRSS / 1024, bytes, error }));
