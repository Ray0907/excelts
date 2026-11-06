import { performance } from 'node:perf_hooks';
import { Checksum } from './checksum.mjs';

const [arm, path, nStr] = process.argv.slice(2);
const { default: read } = await import(`./read-arms/${arm}.mjs`);
const sum = new Checksum();
const cpu0 = process.cpuUsage();
const t0 = performance.now();
let error = null;
try { await read(path, sum, Number(nStr)); } catch (e) { error = String(e && e.stack ? e.stack.split('\n').slice(0, 2).join(' ') : e); }
const ms = performance.now() - t0;
const cpu = process.cpuUsage(cpu0);
console.log(JSON.stringify({ arm, ms, cpuMs: (cpu.user + cpu.system) / 1000, rssMB: process.resourceUsage().maxRSS / 1024, result: sum.result(), error }));
