// Untimed: writes the read-benchmark input files with two unrelated generators (exceljs, openpyxl) so no reader is tested only on its own writer's output.
import { spawnSync } from 'node:child_process';
import { createWriteStream, existsSync, mkdirSync } from 'node:fs';
import { once } from 'node:events';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { COLUMNS, rows } from './data.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, 'inputs');
mkdirSync(dir, { recursive: true });
const sizes = (process.argv[2] ?? '10000,100000,500000').split(',').map(Number);

for (const n of sizes) {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: join(dir, `exceljs-${n}.xlsx`), useSharedStrings: true, useStyles: true });
  const ws = wb.addWorksheet('Sheet1');
  ws.columns = COLUMNS.map((c) => ({ header: c.header, key: c.header, style: c.format ? { numFmt: c.format } : {} }));
  for (const r of rows(n)) ws.addRow(r).commit();
  await ws.commit();
  await wb.commit();

  const nd = join(dir, `data-${n}.ndjson`);
  const out = createWriteStream(nd);
  for (const r of rows(n)) if (!out.write(JSON.stringify(r.map((v) => (v instanceof Date ? { d: v.getTime() } : v))) + '\n')) await once(out, 'drain');
  out.end();
  await once(out, 'finish');
  const py = join(here, '..', '.venv/bin/python');
  if (existsSync(py)) {
    const p = spawnSync(py, [join(here, 'gen_openpyxl.py'), nd, join(dir, `openpyxl-${n}.xlsx`)], { encoding: 'utf8' });
    if (p.status !== 0) throw new Error(p.stderr);
  } else console.log('skipping the openpyxl input (create .venv with openpyxl to enable it)');
  console.log(`inputs for ${n} rows written`);
}
