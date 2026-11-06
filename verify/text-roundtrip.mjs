// Bounded-exhaustive check: every string over a small alphabet, up to a length bound, survives write-side escaping and read-side decoding.
// Usage: node verify/text-roundtrip.mjs [--quick | --full | --if-changed]   (default --quick)
import { execFileSync } from 'node:child_process';
import { escapeText } from '../dist/xml.js';
import { decodeText } from '../dist/xmlRead.js';
import { writeXlsx, readXlsx } from '../dist/index.js';

// Covered code and the commit at which this check was last confirmed to describe it (see .claude/skills/verify-text-roundtrip/SKILL.md).
const COVERS = ['src/xml.ts', 'src/xmlRead.ts', 'src/sharedStrings.ts', 'src/readSharedStrings.ts'];
const SYNC_COMMIT = 'afe392a';

const arg = process.argv[2] ?? '--quick';
if (arg === '--if-changed') {
  const diff = execFileSync('git', ['diff', '--name-only', SYNC_COMMIT, '--', ...COVERS], { encoding: 'utf8' }).trim();
  if (!diff) {
    console.log(`skip: ${COVERS.join(', ')} unchanged since ${SYNC_COMMIT}`);
    process.exit(0);
  }
  console.log(`covered files changed since ${SYNC_COMMIT}:\n${diff}`);
}
const full = arg === '--full';

// Tokens are chosen so the shortest interesting sequences fit: `_x005F_` is 7 tokens, `_x000` + control char is 6.
const ALPHABETS = [
  { name: 'underscore-escapes', tokens: ['_', 'x', '0', '5', 'F', '\u0001', '\r', '&'], quick: 7, full: 8 },
  { name: 'markup-and-newlines', tokens: ['&', '<', '>', '"', '\r', '\n', ']', '_x0001_'], quick: 6, full: 7 },
  { name: 'cdata-and-noncharacters', tokens: ['<![CDATA[', ']]>', '￾', '￿', '_', 'x', '0', 'D'], quick: 4, full: 6 },
];

let checked = 0;
const failures = [];
function check(s) {
  checked++;
  const back = decodeText(escapeText(s));
  if (back !== s && failures.length < 10) failures.push({ input: JSON.stringify(s), escaped: JSON.stringify(escapeText(s)), back: JSON.stringify(back) });
}

for (const { name, tokens, quick, full: fullLen } of ALPHABETS) {
  const max = full ? fullLen : quick;
  const t0 = Date.now();
  const before = checked;
  const walk = (prefix, depth) => {
    check(prefix);
    if (depth === max) return;
    for (const t of tokens) walk(prefix + t, depth + 1);
  };
  walk('', 0);
  console.log(`${name}: ${checked - before} strings up to ${max} tokens, ${Date.now() - t0} ms`);
}

// The same property through the real writer and reader, on the shortest strings that exercise each rule.
const SAMPLES = ['_x005F_', '_x0001_', '_x000\u0001', '_x005F_x005F_', 'a\r\nb', '\r', '<![CDATA[x]]>', ']]>', '\u0001', '&amp;', '_x00', '_x000D_', '￾'];
const book = writeXlsx({ name: 'S', columns: [{ header: 'h' }], rows: SAMPLES.map((s) => [s]) });
const rows = [...readXlsx(book).sheets[0].rows()].slice(1).map((r) => r[0]);
SAMPLES.forEach((s, i) => {
  checked++;
  if (rows[i] !== s && failures.length < 10) failures.push({ input: JSON.stringify(s), via: 'writeXlsx/readXlsx', back: JSON.stringify(rows[i]) });
});

console.log(`checked ${checked} strings, ${failures.length ? 'FAIL' : 'ok'}`);
for (const f of failures) console.log(f);
process.exit(failures.length ? 1 : 0);
