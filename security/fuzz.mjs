import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { ROOT } from './harness.mjs';
import { abortReason, spawnChild } from './lib/spawn.mjs';
import { entryData, zipBuild, zipFields, zipParse } from './lib/zip.mjs';
import { int, pick } from './fuzz-rng.mjs';
import { xmlMutation } from './fuzz-xml.mjs';

export function mulberry32(a) {
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const mix = (seed, i) => {
  let h = (seed ^ Math.imul(i + 1, 0x9e3779b1)) >>> 0;
  h = Math.imul(h ^ (h >>> 16), 0x85ebca6b);
  h = Math.imul(h ^ (h >>> 13), 0xc2b2ae35);
  return (h ^ (h >>> 16)) >>> 0;
};

const INTERESTING = [0x00, 0xff, 0x3c, 0x3e, 0x26, 0x22, 0x27, 0x2f, 0x3d, 0x7f, 0x80, 0x20, 0x0a];
const INTERESTING_NAMES = ['NUL', 'FF', '<', '>', '&', '"', "'", '/', '=', 'DEL', 'C1', 'SP', 'LF'];

/** Seeds: ~60 small corpus files stratified by source, plus the fixtures writeXlsx produces. */
export async function loadSeeds({ seed, bundle, target = 60 }) {
  const seeds = [];
  const manifest = path.join(ROOT, 'corpus', 'manifest.tsv');
  if (existsSync(manifest)) {
    const lines = readFileSync(manifest, 'utf8').split('\n');
    const head = lines[0].split('\t');
    const col = (n) => head.indexOf(n);
    const bySource = new Map();
    for (const line of lines.slice(1)) {
      const f = line.split('\t');
      if (f.length < head.length - 2) continue;
      const bytes = Number(f[col('bytes')]);
      if (f[col('valid')] !== 'ok' || !(bytes < 300_000) || f[col('dup_of')] || !(Number(f[col('sheets')]) >= 1)) continue;
      const file = path.join(ROOT, 'corpus', f[col('file')]);
      if (!existsSync(file)) continue;
      const src = f[col('source')];
      if (!bySource.has(src)) bySource.set(src, []);
      bySource.get(src).push(file);
    }
    const rng = mulberry32(mix(seed, 7));
    const lists = [...bySource.entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([, l]) => l.sort().sort(() => rng() - 0.5));
    for (let round = 0; seeds.length < target && lists.some((l) => l.length > round); round++) {
      for (const l of lists) if (l[round] && seeds.length < target) seeds.push({ id: path.relative(ROOT, l[round]), bytes: new Uint8Array(readFileSync(l[round])) });
    }
  }
  const { writeXlsx } = await import(pathToFileURL(bundle).href);
  const { FIXTURES } = await import('../verify/fixtures.mjs');
  for (const f of FIXTURES.filter((x) => !x.big)) seeds.push({ id: `fixture:${f.name}`, bytes: writeXlsx({ name: 'Sheet1', columns: f.columns, rows: f.rows }) });
  seeds.push({
    id: 'fixture:tiny',
    bytes: writeXlsx({ name: 'T', columns: [{ header: 'a' }, { header: 'd', format: 'yyyy-mm-dd' }], rows: [['x', new Date(0)], [1.5, null], [true, new Date(86400000)]] }),
  });
  return seeds;
}

const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

const BYTE_OPS = {
  bitflip(b, rng, r) {
    const n = 1 + int(rng, 8);
    for (let k = 0; k < n; k++) {
      const at = int(rng, b.length);
      const bit = int(rng, 8);
      b[at] ^= 1 << bit;
      r.push(`flip@${at}.${bit}`);
    }
    return b;
  },
  subst(b, rng, r) {
    const n = 1 + int(rng, 16);
    for (let k = 0; k < n; k++) {
      const at = int(rng, b.length);
      const v = int(rng, INTERESTING.length);
      b[at] = INTERESTING[v];
      r.push(`set@${at}=${INTERESTING_NAMES[v]}`);
    }
    return b;
  },
  chunkDelete(b, rng, r) {
    const len = 1 + int(rng, Math.max(1, Math.min(4096, b.length >> 2)));
    const at = int(rng, b.length - len + 1);
    r.push(`delete@${at}+${len}`);
    return Buffer.concat([b.subarray(0, at), b.subarray(at + len)]);
  },
  chunkDup(b, rng, r) {
    const len = 1 + int(rng, Math.max(1, Math.min(4096, b.length >> 2)));
    const at = int(rng, b.length - len + 1);
    const times = 1 + int(rng, 8);
    r.push(`dup@${at}+${len}x${times}`);
    return Buffer.concat([b.subarray(0, at + len), ...Array.from({ length: times }, () => b.subarray(at, at + len)), b.subarray(at + len)]);
  },
  chunkSwap(b, rng, r) {
    const len = 1 + int(rng, Math.max(1, Math.min(2048, b.length >> 3)));
    const x = int(rng, b.length - len + 1);
    const y = int(rng, b.length - len + 1);
    r.push(`swap@${x}<->${y}+${len}`);
    const a = Buffer.from(b.subarray(x, x + len));
    b.copy(b, x, y, y + len);
    a.copy(b, y);
    return b;
  },
  truncate(b, rng, r) {
    const at = rng() < 0.5 ? int(rng, b.length) : clamp(b.length - 1 - int(rng, 200), 0, b.length);
    r.push(`truncate@${at}`);
    return b.subarray(0, at);
  },
};

const FIELD_VALUES = (rng, orig, size) => {
  const max = size === 2 ? 0xffff : 0xffffffff;
  return pick(rng, [0, 1, orig + 1, Math.max(0, orig - 1), max, max - 1, 0x7fffffff & max, 0x8000 & max, int(rng, max), 65536 & max, 1 << 24]);
};

function zipFieldMutation(b, rng, r) {
  const fields = zipFields(b);
  if (!fields || !fields.length) return BYTE_OPS.bitflip(b, rng, r);
  const n = 1 + int(rng, 3);
  for (let k = 0; k < n; k++) {
    const f = pick(rng, fields);
    const orig = f.size === 2 ? b.readUInt16LE(f.at) : b.readUInt32LE(f.at);
    const v = f.label.endsWith('method') ? pick(rng, [0, 1, 6, 8, 9, 12, 14, 93, 95, 98, 99]) : FIELD_VALUES(rng, orig, f.size);
    if (f.size === 2) b.writeUInt16LE(v & 0xffff, f.at);
    else b.writeUInt32LE(v >>> 0, f.at);
    r.push(`${f.label}@${f.at}:${orig}->${v >>> 0}`);
  }
  return b;
}

function zipRebuild(b, rng, r) {
  const entries = zipParse(b);
  if (!entries || !entries.length) return undefined;
  const ents = entries.map((e) => ({ name: e.name, comp: e.comp, method: e.method, usize: e.usize, crc: e.crc }));
  const opts = {};
  const op = pick(rng, ['collide', 'traversal', 'absolute', 'nul', 'dup', 'drop', 'reorder', 'hugeUsize', 'method', 'overlap', 'comment', 'zip64', 'longName', 'case']);
  const i = int(rng, ents.length);
  const j = int(rng, ents.length);
  r.push(`rebuild:${op}[${i},${j}]`);
  switch (op) {
    case 'collide':
      ents[i].name = ents[j].name;
      break;
    case 'traversal':
      ents[i].name = pick(rng, ['../', '..\\', 'xl/../../', '../../../tmp/', 'xl/worksheets/../../..//']) + ents[i].name;
      break;
    case 'absolute':
      ents[i].name = pick(rng, ['/', 'C:\\', '\\\\host\\share\\', 'file:///']) + ents[i].name;
      break;
    case 'nul':
      ents[i].name = ents[i].name.replace(/\.xml$/, '\u0000.xml');
      break;
    case 'dup':
      ents.push({ ...ents[i] });
      break;
    case 'drop':
      ents.splice(i, 1);
      break;
    case 'reorder':
      ents.reverse();
      break;
    case 'hugeUsize':
      ents[i].usize = pick(rng, [0xfffffffe, 0x7fffffff, 0]);
      break;
    case 'method':
      ents[i].method = pick(rng, [0, 9, 12, 99]);
      break;
    case 'overlap':
      ents[i] = { ...ents[j], name: ents[i].name, pointTo: j };
      break;
    case 'comment':
      opts.comment = Buffer.alloc(int(rng, 65536), 0x50);
      break;
    case 'zip64':
      opts.zip64Locator = true;
      break;
    case 'longName':
      ents[i].name = ents[i].name + 'a'.repeat(int(rng, 65535 - ents[i].name.length));
      break;
    case 'case':
      ents[i].name = ents[i].name.toUpperCase();
      break;
  }
  return zipBuild(ents, opts);
}


const CLASSES = [
  ['byte', 4],
  ['zipfield', 3],
  ['zipbuild', 3],
  ['xml', 6],
];

/** Builds case `index` from the master seed alone, so `--seed S --case I` reproduces it exactly. */
export function makeCase(seed, index, seeds) {
  const rng = mulberry32(mix(seed, index));
  const s = pick(rng, seeds);
  const recipe = [`seed=${seed}`, `case=${index}`, `file=${s.id}`];
  let total = CLASSES.reduce((a, [, w]) => a + w, 0);
  let x = rng() * total;
  let cls = CLASSES[0][0];
  for (const [c, w] of CLASSES) {
    if ((x -= w) < 0) {
      cls = c;
      break;
    }
  }
  recipe.push(`class=${cls}`);
  let b = Buffer.from(s.bytes);
  let out;
  if (cls === 'byte') {
    const steps = 1 + int(rng, 3);
    for (let k = 0; k < steps && b.length > 8; k++) b = Buffer.from(BYTE_OPS[pick(rng, Object.keys(BYTE_OPS))](b, rng, recipe));
    out = b;
  } else if (cls === 'zipfield') out = zipFieldMutation(b, rng, recipe);
  else if (cls === 'zipbuild') out = zipRebuild(b, rng, recipe);
  else out = xmlMutation(b, rng, recipe);
  if (!out) {
    recipe.push('(mutation not applicable, byte flip instead)');
    out = BYTE_OPS.bitflip(Buffer.from(s.bytes), rng, recipe);
  }
  return { bytes: new Uint8Array(out.buffer, out.byteOffset, out.byteLength), recipe, cls };
}

export function saveCrash(dir, bytes, info) {
  mkdirSync(dir, { recursive: true });
  const hash = createHash('sha1').update(bytes).digest('hex').slice(0, 12);
  writeFileSync(path.join(dir, `${hash}.xlsx`), bytes);
  writeFileSync(path.join(dir, `${hash}.json`), JSON.stringify(info, null, 2));
  return hash;
}

/** Runs inside a child process: announces each case before and after so a supervisor can tell which one aborted the process. */
export async function fuzzChunk({ seed, indices, jobs, bundle, mode, pool, runCase }) {
  const seeds = await loadSeeds({ seed, bundle });
  await pool(indices, jobs, async (index) => {
    process.stdout.write(`S ${index}\n`);
    const c = makeCase(seed, index, seeds);
    const r = await runCase(c.bytes, { bundle, mode });
    process.stdout.write(`R ${index} ${JSON.stringify({ outcome: r.outcome, verdict: r.verdict, reasons: r.reasons, wallMs: r.wallMs, memMb: r.memMb, err: r.err })}\n`);
  });
}

export async function fuzz({ seed, iterations, jobs, bundle, mode = 'collect', crashDir, onlyCase, log = console.log }) {
  const seeds = await loadSeeds({ seed, bundle });
  log(`fuzz: seed=${seed} iterations=${iterations} jobs=${jobs} mode=${mode} seeds=${seeds.length}`);
  let todo = onlyCase !== undefined ? [onlyCase] : Array.from({ length: iterations }, (_, i) => i);
  const summary = { cases: 0, pass: 0, warn: 0, fail: 0, aborts: 0, byOutcome: {}, groups: new Map(), ms: [], t0: performance.now() };

  const record = (index, r) => {
    const c = makeCase(seed, index, seeds);
    summary.cases++;
    summary[r.verdict]++;
    summary.byOutcome[r.outcome] = (summary.byOutcome[r.outcome] ?? 0) + 1;
    summary.ms.push(r.wallMs ?? 0);
    if (r.verdict === 'pass') return;
    const family = c.recipe.find((x) => x.startsWith('xml:'))?.split(':').slice(2).join(':') ?? c.recipe.find((x) => x.startsWith('rebuild:'))?.replace(/\[.*$/, '') ?? c.cls;
    const sig = `${r.verdict} ${r.outcome} ${(r.err?.message ?? r.reasons?.join(';') ?? '').replace(/\d+/g, 'N').slice(0, 80)} | ${family}`;
    const g = summary.groups.get(sig) ?? { count: 0, worst: r, example: undefined, hashes: [] };
    g.count++;
    if ((r.wallMs ?? 0) >= (g.worst.wallMs ?? 0)) g.worst = r;
    if (g.hashes.length < 2 && r.verdict === 'fail') {
      g.hashes.push(saveCrash(crashDir, c.bytes, { signature: sig, recipe: c.recipe, result: r, replay: `node security/run.mjs fuzz --seed ${seed} --case ${index}   or   node security/run.mjs --replay security/crashes/<hash>.xlsx` }));
    }
    g.example ??= { recipe: c.recipe.join(' '), index };
    summary.groups.set(sig, g);
  };

  const abort = (index, r) => {
    summary.aborts++;
    record(index, { outcome: 'abort', verdict: 'fail', reasons: [abortReason(r)], wallMs: 0 });
  };

  const CHUNK = 250;
  while (todo.length) {
    const batch = todo.slice(0, CHUNK);
    todo = todo.slice(CHUNK);
    const started = new Set();
    const done = new Set();
    const res = await spawnChild(['fuzz', `--seed=${seed}`, `--indices=${JSON.stringify(batch)}`, `--jobs=${jobs}`, `--bundle=${bundle}`, `--mode=${mode}`], {
      timeoutMs: 600_000,
      onLine(line) {
        const m = /^([SR]) (\d+)(?: (.*))?$/.exec(line);
        if (!m) return;
        const i = Number(m[2]);
        if (m[1] === 'S') started.add(i);
        else {
          done.add(i);
          record(i, JSON.parse(m[3]));
        }
      },
    });
    if (res.code === 0 && done.size === batch.length) continue;
    // The process died with cases in flight; replay each alone so only the real culprit is blamed.
    const inflight = [...started].filter((i) => !done.has(i));
    for (const i of inflight) {
      const solo = await spawnChild(['fuzz', `--seed=${seed}`, `--indices=[${i}]`, '--jobs=1', `--bundle=${bundle}`, `--mode=${mode}`], {
        timeoutMs: 120_000,
        onLine(line) {
          const m = /^R (\d+) (.*)$/.exec(line);
          if (m) {
            done.add(i);
            record(i, JSON.parse(m[2]));
          }
        },
      });
      if (!done.has(i)) abort(i, solo);
    }
    todo = batch.filter((i) => !started.has(i) && !done.has(i)).concat(todo);
    if (!inflight.length && !batch.some((i) => done.has(i))) {
      abort(batch[0], res);
      todo = todo.filter((i) => i !== batch[0]);
    }
  }
  summary.seconds = (performance.now() - summary.t0) / 1000;
  summary.seeds = seeds.length;
  return summary;
}
