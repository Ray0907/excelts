import { Worker } from 'node:worker_threads';
import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { mkdirSync } from 'node:fs';
import path from 'node:path';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const BUDGET = { ms: 5000, heapMb: 512 };
const WORKER = path.join(ROOT, 'security', 'worker.mjs');

export function buildReader() {
  const outfile = path.join(ROOT, 'security', '.build', 'reader.mjs');
  mkdirSync(path.dirname(outfile), { recursive: true });
  execFileSync(path.join(ROOT, 'node_modules', '.bin', 'esbuild'), [path.join(ROOT, 'src', 'index.ts'), '--bundle', '--platform=node', '--format=esm', `--outfile=${outfile}`, '--log-level=error'], { stdio: 'inherit' });
  return outfile;
}

/**
 * Runs the reader over `bytes` in a fresh worker under the budget. Peak memory is sampled inside the worker between rows and after each
 * inflate (a synchronous reader cannot be sampled from outside), so a single huge allocation is seen only through the inflate accounting and the OOM kill.
 */
export function runCase(bytes, { bundle, mode = 'collect', ms = BUDGET.ms, heapMb = BUDGET.heapMb, trackRss = false, check } = {}) {
  return new Promise((resolve) => {
    const rssStart = trackRss ? process.memoryUsage.rss() : 0;
    let rssPeak = rssStart;
    let poll;
    if (trackRss) poll = setInterval(() => (rssPeak = Math.max(rssPeak, process.memoryUsage.rss())), 20);
    const worker = new Worker(WORKER, { workerData: { bytes, bundle: pathToFileURL(bundle).href, mode }, resourceLimits: { maxOldGenerationSizeMb: heapMb } });
    let t0 = performance.now();
    let settled = false;
    let message;
    let timer;
    let terminating = false;
    let hardTimer;
    const finish = (r) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      clearTimeout(hardTimer);
      clearInterval(poll);
      const wall = performance.now() - t0;
      r.wallMs = wall;
      r.rssMb = trackRss ? Math.max(0, rssPeak - rssStart) / 1048576 : undefined;
      resolve(judge(r, { ms, heapMb, check }));
    };
    worker.on('online', () => {
      t0 = performance.now();
      timer = setTimeout(() => {
        terminating = true;
        worker.terminate();
        // A native zlib call cannot be interrupted; do not wait forever for the exit event.
        hardTimer = setTimeout(() => finish({ outcome: 'timeout', hung: true }), 5000);
      }, ms);
    });
    worker.on('message', (m) => (message = m));
    worker.on('error', (e) => {
      if (e && e.code === 'ERR_WORKER_OUT_OF_MEMORY') finish({ outcome: 'oom', err: { name: e.name, message: e.message } });
      else finish({ outcome: 'crash', err: { name: e?.name, message: String(e?.message ?? e).slice(0, 200) } });
    });
    worker.on('exit', (code) => {
      if (message) return finish(message);
      if (settled) return;
      finish({ outcome: terminating || performance.now() - t0 >= ms ? 'timeout' : 'crash', err: terminating ? undefined :  { name: 'exit', message: `worker exited ${code} without a result` } });
    });
  });
}

function judge(r, { ms, heapMb, check }) {
  const reasons = [];
  let verdict = 'pass';
  const fail = (why) => {
    verdict = 'fail';
    reasons.push(why);
  };
  if (r.outcome === 'timeout') fail(`exceeded ${ms} ms${r.hung ? ' (uninterruptible)' : ''}`);
  if (r.outcome === 'oom') fail(`heap limit ${heapMb} MB hit`);
  if (r.outcome === 'crash') fail('worker crashed');
  if (r.outcome === 'threw-nonerror') fail('threw a non-Error');
  if (r.outcome === 'threw-internal' && verdict === 'pass') {
    verdict = 'warn';
    reasons.push('engine/internal error instead of a clear message');
  }
  const mem = Math.max(r.peakMb ?? 0, r.rssMb ?? 0);
  r.memMb = mem;
  if (mem > heapMb && r.outcome !== 'oom') fail(`peak ${mem.toFixed(0)} MB over ${heapMb} MB`);
  if ((r.ms ?? 0) > ms || (r.wallMs ?? 0) > ms * 1.2) if (r.outcome !== 'timeout') fail(`took ${(r.wallMs ?? r.ms).toFixed(0)} ms`);
  const extra = check?.(r);
  if (extra) fail(extra);
  r.verdict = verdict;
  r.reasons = reasons;
  return r;
}

export async function pool(items, jobs, fn) {
  const results = new Array(items.length);
  let next = 0;
  const lane = async () => {
    for (;;) {
      const i = next++;
      if (i >= items.length) return;
      results[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: jobs }, lane));
  return results;
}
