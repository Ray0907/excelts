import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { writeXlsx } from '../dist/index.js';
import { FIXTURES, expectedGrid } from './fixtures.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const out = join(here, 'out');
mkdirSync(out, { recursive: true });
const selfTest = process.argv.includes('--self-test');

const same = (a, b) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= Math.max(1e-9 * Math.abs(b), 1.5) * (b > 1e11 && b < 1e13 ? 1 : 1e-6) : a === b);
const trim = (g) => g.map((r) => { const c = [...r]; while (c.length && c[c.length - 1] == null) c.pop(); return c; });

function compare(got, want) {
  let cells = 0, bad = [];
  const g = trim(got), w = trim(want);
  if (g.length !== w.length) bad.push(`rows ${g.length} != ${w.length}`);
  for (let i = 0; i < Math.min(g.length, w.length) && bad.length < 5; i++) {
    const n = Math.max(g[i].length, w[i].length);
    for (let j = 0; j < n; j++) {
      cells++;
      const a = g[i][j] ?? null, b = w[i][j] ?? null;
      if (!(a === b || (a !== null && b !== null && same(a, b)))) { bad.push(`r${i + 1}c${j + 1} got ${JSON.stringify(a)} want ${JSON.stringify(b)}`); if (bad.length >= 5) break; }
    }
  }
  return { cells, bad };
}

const readers = {
  async exceljs(path) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const ws = wb.worksheets[0];
    const grid = [];
    ws.eachRow({ includeEmpty: true }, (row) => {
      const r = [];
      row.eachCell({ includeEmpty: true }, (c, col) => { const v = c.value; r[col - 1] = v instanceof Date ? v.getTime() : v; });
      grid.push(r);
    });
    return grid;
  },
  async sheetjs(path) {
    const wb = XLSX.read(readFileSync(path), { cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null, blankrows: true });
    return rows.map((r) => r.map((v) => (v instanceof Date ? Math.round(v.getTime()) : v)));
  },
  async openpyxl(path) {
    const jf = path + '.openpyxl.json';
    const p = spawnSync(join(here, '..', '.venv/bin/python'), [join(here, 'read_openpyxl.py'), path, jf], { encoding: 'utf8' });
    if (p.status !== 0) throw new Error(p.stderr.trim().split('\n').pop());
    return JSON.parse(readFileSync(jf, 'utf8'));
  },
};

let failed = false;
const tally = Object.fromEntries(Object.keys(readers).map((k) => [k, { cells: 0, fixtures: 0 }]));
for (const f of FIXTURES) {
  const path = join(out, `${f.name}.xlsx`);
  const bytes = writeXlsx({ name: 'Sheet1', columns: f.columns, rows: f.rows });
  writeFileSync(path, bytes);
  let want = expectedGrid(f);
  if (selfTest) want = want.map((r, i) => (i === 1 ? r.map((v, j) => (j === 0 ? 'MUTATED' : v)) : r)); // sensitivity: expectation no longer matches the file
  if (selfTest && f.name === 'edge') writeFileSync(path, bytes.subarray(0, Math.floor(bytes.length / 2))); // truncated zip must be rejected
  for (const [name, read] of Object.entries(readers)) {
    let line;
    try {
      const { cells, bad } = compare(await read(path), want);
      tally[name].cells += cells;
      tally[name].fixtures++;
      line = bad.length ? `MISMATCH ${bad.join(' | ')}` : `ok ${cells} cells`;
      if (bad.length) failed = true;
    } catch (e) {
      line = `READ-ERROR ${String(e.message).split('\n')[0].slice(0, 100)}`;
      failed = true;
    }
    console.log(`${f.name.padEnd(10)} ${name.padEnd(9)} ${line}`);
  }
}
console.log('\n' + Object.entries(tally).map(([k, v]) => `${k}: ${v.cells} cells in ${v.fixtures} fixtures`).join(' | '));
if (selfTest) {
  console.log(failed ? 'self-test PASS: corrupted input was detected' : 'self-test FAIL: corruption went undetected');
  process.exit(failed ? 0 : 1);
}
process.exit(failed ? 1 : 0);
