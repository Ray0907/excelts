import { readFileSync } from 'node:fs';
import { readXlsx } from '../../dist/index.js';
export default async function read(path, sum) { const wb = readXlsx(readFileSync(path)); for (const row of wb.sheets[0].rows()) sum.row(row, 0); }
