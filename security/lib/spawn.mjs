import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const CHILD = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'child.mjs');

/**
 * Runs `child.mjs` in its own process. A worker that exhausts memory can abort the whole process (V8 fatal OOM), which no
 * in-process handler can survive, so every case that might do that runs behind this boundary.
 * Do not pass a heap flag to a child that runs workers: it silently overrides the workers' resourceLimits. `onLine` sees stdout line by line.
 */
export function spawnChild(args, { timeoutMs = 120_000, heapMb, onLine } = {}) {
  return new Promise((resolve) => {
    const p = spawn(process.execPath, [...(heapMb ? [`--max-old-space-size=${heapMb}`] : []), CHILD, ...args], { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '';
    let err = '';
    let pending = '';
    let killed = false;
    const timer = setTimeout(() => {
      killed = true;
      p.kill('SIGKILL');
    }, timeoutMs);
    p.stdout.on('data', (d) => {
      const s = d.toString();
      if (onLine) {
        pending += s;
        let i;
        while ((i = pending.indexOf('\n')) >= 0) {
          onLine(pending.slice(0, i));
          pending = pending.slice(i + 1);
        }
      } else out += s;
    });
    p.stderr.on('data', (d) => (err = (err + d.toString()).slice(-4000)));
    p.on('close', (code, signal) => {
      clearTimeout(timer);
      if (onLine && pending) onLine(pending);
      resolve({ code, signal, killed, stdout: out, stderr: err });
    });
  });
}

/** One-line reason a child process died, taken from V8's fatal error banner when there is one. */
export function abortReason(r) {
  const fatal = /FATAL ERROR: ([^\n]+)/.exec(r.stderr)?.[1];
  if (r.killed) return 'child killed after the hard timeout (uninterruptible)';
  return `process aborted${r.signal ? ` by ${r.signal}` : ` with exit ${r.code}`}${fatal ? `: ${fatal}` : ''}`;
}
