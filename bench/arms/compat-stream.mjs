import ExcelJS from '../../dist/compat/index.js';
import { COLUMNS } from '../data.mjs';

export default async function write(rows, path) {
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path, useSharedStrings: true, useStyles: true });
  const ws = wb.addWorksheet('Sheet1');
  ws.columns = COLUMNS.map((c) => ({ header: c.header, key: c.header, style: c.format ? { numFmt: c.format } : {} }));
  for (const r of rows) ws.addRow(r).commit();
  await ws.commit();
  await wb.commit();
}
