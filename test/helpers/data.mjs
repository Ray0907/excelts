export const HEADERS = ['id', 'user', 'uniq', 'amount', 'at', 'flag', 'qty', 'cat', 'price', 'note'];
export const DATE_FORMAT = 'yyyy-mm-dd hh:mm:ss';
export const COLUMNS = HEADERS.map((header) => ({ header, format: header === 'at' ? DATE_FORMAT : undefined }));
const CATS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j'];
const T0 = Date.UTC(2020, 0, 1);

export function row(i) {
  return [
    i,
    `user_${i % 1000}`,
    `uniq_${i}_${(Math.imul(i, 2654435761) >>> 0).toString(36)}`,
    Math.round(((i * 37) % 100000)) / 100,
    new Date(T0 + i * 60000),
    i % 3 === 0,
    i % 97,
    CATS[i % 10],
    Math.round(((i * 11) % 500000)) / 100,
    `note & <${i % 50}>`,
  ];
}

export function* rows(n) {
  for (let i = 0; i < n; i++) yield row(i);
}
