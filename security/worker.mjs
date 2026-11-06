import { parentPort, workerData } from 'node:worker_threads';
import { createRequire } from 'node:module';
import v8 from 'node:v8';

const require = createRequire(import.meta.url);
const zlib = require('node:zlib');
const MB = 1048576;

const acct = { inflateCalls: 0, inflateOut: 0, inflateMax: 0, inflateMs: 0 };
let peakBytes = 0;
const sample = () => {
  const s = v8.getHeapStatistics();
  const total = s.used_heap_size + s.external_memory;
  if (total > peakBytes) peakBytes = total;
};

// The reader destructures zlib at import time, so the wrapper must be installed before it loads.
const inflate = zlib.inflateRawSync;
zlib.inflateRawSync = function (...args) {
  const t0 = performance.now();
  const out = inflate.apply(this, args);
  acct.inflateMs += performance.now() - t0;
  acct.inflateCalls++;
  acct.inflateOut += out.length;
  if (out.length > acct.inflateMax) acct.inflateMax = out.length;
  sample();
  return out;
};

const INTERNAL = /Maximum call stack|Invalid (string|array|typed array) length|Invalid code point|Cannot read propert|is not a function|is not iterable|is not defined|Array buffer allocation|Cannot convert|Cannot create a (Buffer|string)|Invalid count value|before initialization|of undefined|of null/i;

function classify(e) {
  if (!(e instanceof Error)) return 'threw-nonerror';
  if (e instanceof TypeError || e instanceof ReferenceError || e instanceof SyntaxError || INTERNAL.test(e.message)) return 'threw-internal';
  if (e instanceof RangeError && !/zip64|exceed|limit|too (large|many)/i.test(e.message)) return 'threw-internal';
  return 'threw-error';
}

const { bytes, bundle, mode } = workerData;
const { readXlsx } = await import(bundle);

const res = { rows: 0, cells: 0, sheets: 0, maxStr: 0, maxRowLen: 0 };
const t0 = performance.now();
let outcome = 'returned';
let err;
try {
  const wb = readXlsx(bytes);
  sample();
  res.sheets = wb.sheets.length;
  const kept = [];
  for (const sh of wb.sheets) {
    const acc = mode === 'collect' ? [] : undefined;
    for (const r of sh.rows()) {
      res.rows++;
      if (r.length > res.maxRowLen) res.maxRowLen = r.length;
      if (r.length) {
        res.cells += r.length;
        r.forEach((v) => {
          if (typeof v === 'string' && v.length > res.maxStr) res.maxStr = v.length;
        });
      }
      if (acc) acc.push(r);
      if ((res.rows & 0xfff) === 0) sample();
    }
    if (acc) kept.push(acc);
  }
  sample();
  res.retained = kept.length;
} catch (e) {
  outcome = classify(e);
  err = e instanceof Error ? { name: e.name, message: String(e.message).slice(0, 300) } : { name: typeof e, message: String(e).slice(0, 300) };
  sample();
}
const ms = performance.now() - t0;
const hs = v8.getHeapStatistics();
parentPort.postMessage({
  outcome,
  err,
  ms,
  ...res,
  ...acct,
  peakMb: Math.max(peakBytes, hs.peak_malloced_memory) / MB,
  inflateMaxMb: acct.inflateMax / MB,
  inflateOutMb: acct.inflateOut / MB,
});
