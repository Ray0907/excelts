import { writeFileSync } from 'node:fs';
import { writeXlsx } from '../../dist/index.js';
import { COLUMNS } from '../data.mjs';

export default async function write(rows, path) {
  writeFileSync(path, writeXlsx({ name: 'Sheet1', columns: COLUMNS, rows }));
}
