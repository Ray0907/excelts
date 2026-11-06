// Writes the verify fixtures with writeXlsx and validates every SpreadsheetML part against the ECMA-376 XSDs.
import { mkdirSync, writeFileSync } from 'node:fs';
import { writeXlsx } from '../dist/index.js';
import { FIXTURES } from '../verify/fixtures.mjs';
import { validateXlsx } from './xsd.mjs';
import { compatFixtures } from './compat-fixtures.mjs';

mkdirSync('validate/out', { recursive: true });
let failed = 0;
for (const f of FIXTURES) {
  const path = `validate/out/${f.name}.xlsx`;
  writeFileSync(path, writeXlsx({ name: 'Sheet1', columns: f.columns, rows: f.rows }));
  const res = validateXlsx(path);
  const bad = res.filter((r) => !r.ok);
  failed += bad.length;
  console.log(`${bad.length ? 'FAIL' : 'ok  '} ${f.name} (${res.length} parts)`);
  for (const b of bad) console.log(`     ${b.part}: ${b.errors.join(' | ').slice(0, 400)}`);
}
for (const f of compatFixtures()) {
  const path = `validate/out/${f.name}.xlsx`;
  writeFileSync(path, await f.wb.xlsx.writeBuffer());
  const res = validateXlsx(path);
  const bad = res.filter((r) => !r.ok);
  failed += bad.length;
  console.log(`${bad.length ? 'FAIL' : 'ok  '} ${f.name} (${res.length} parts)`);
  for (const b of bad) console.log(`     ${b.part}: ${b.errors.join(' | ').slice(0, 400)}`);
}
process.exit(failed ? 1 : 0);
