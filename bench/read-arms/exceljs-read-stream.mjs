import ExcelJS from 'exceljs';
export default async function read(path, sum) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(path, { sharedStrings: 'cache', styles: 'cache', hyperlinks: 'ignore', worksheets: 'emit', entries: 'ignore' });
  for await (const ws of reader) {
    for await (const row of ws) sum.row(row.values, 1);
    break;
  }
}
