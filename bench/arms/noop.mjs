export default async function write(rows) {
  let n = 0;
  for (const r of rows) n += r.length;
  return n;
}
