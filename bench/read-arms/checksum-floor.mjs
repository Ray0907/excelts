import { rows, HEADERS } from '../data.mjs';
export default async function read(_path, sum, n) {
  sum.row(HEADERS);
  for (const r of rows(n)) sum.row(r);
}
