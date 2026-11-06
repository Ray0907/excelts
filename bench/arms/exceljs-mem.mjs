import ExcelJS from 'exceljs';
import { COLUMNS } from '../data.mjs';

export default async function write(rows, path) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.columns = COLUMNS.map((c) => ({ header: c.header, key: c.header, style: c.format ? { numFmt: c.format } : {} }));
  for (const r of rows) ws.addRow(r);
  await wb.xlsx.writeFile(path);
}
