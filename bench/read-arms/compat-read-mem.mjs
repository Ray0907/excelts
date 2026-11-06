import ExcelJS from '../../dist/compat/index.js';
export default async function read(path, sum) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  wb.worksheets[0].eachRow({ includeEmpty: false }, (row) => sum.row(row.values, 1));
}
