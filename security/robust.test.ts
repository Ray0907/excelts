import { describe, expect, it } from 'vitest';
import * as fc from 'fast-check';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readXlsx, writeXlsx } from '../src/index.js';

const RUNS = Number(process.env.FC_RUNS ?? 100);
const SEED = process.env.FC_SEED === undefined ? undefined : Number(process.env.FC_SEED);
const params = { numRuns: RUNS, seed: SEED } as const;
const BUDGET_MS = 2000;

const isError = (e: unknown): boolean => e instanceof Error || Object.prototype.toString.call(e) === '[object Error]';

/** Consumes every row of every sheet; returns normally or throws an Error. Anything else, or taking too long, is a failure. */
function readAll(bytes: Uint8Array): void {
  const t0 = performance.now();
  try {
    for (const sheet of readXlsx(bytes).sheets) for (const _ of sheet.rows()) void _;
  } catch (e) {
    expect(isError(e), `threw a non-Error: ${String(e)}`).toBe(true);
    expect((e as Error).message.length).toBeGreaterThan(0);
  }
  expect(performance.now() - t0).toBeLessThan(BUDGET_MS);
}

const valid = writeXlsx({
  name: 'Seed',
  columns: [{ header: 'text' }, { header: 'n' }, { header: 'when', format: 'yyyy-mm-dd' }],
  rows: [['a & <b>', 1.5, new Date(0)], ['', -2, null], [true, 1e21, new Date(86400000)]],
});

const eocd = (count: number, cdSize: number, cdOffset: number, commentLen: number): Uint8Array => {
  const b = new Uint8Array(22);
  const v = new DataView(b.buffer);
  v.setUint32(0, 0x06054b50, true);
  v.setUint16(8, count, true);
  v.setUint16(10, count, true);
  v.setUint32(12, cdSize, true);
  v.setUint32(16, cdOffset, true);
  v.setUint16(20, commentLen, true);
  return b;
};

describe('reader on hostile input never throws a non-Error and returns quickly', () => {
  it('arbitrary bytes', () => {
    fc.assert(fc.property(fc.uint8Array({ maxLength: 4000 }), readAll), params);
  });

  it('arbitrary bytes ending in a plausible end-of-central-directory record', () => {
    fc.assert(
      fc.property(fc.uint8Array({ maxLength: 4000 }), fc.integer({ min: 0, max: 0xffff }), fc.nat({ max: 0xffffffff }), fc.nat({ max: 0xffffffff }), (body, count, size, offset) => {
        const tail = eocd(count, size, offset % (body.length + 1), 0);
        const all = new Uint8Array(body.length + tail.length);
        all.set(body);
        all.set(tail, body.length);
        readAll(all);
      }),
      params,
    );
  });

  it('a valid file with substituted and truncated bytes', () => {
    const edit = fc.record({ at: fc.nat({ max: valid.length - 1 }), value: fc.oneof(fc.integer({ min: 0, max: 255 }), fc.constantFrom(0, 255, 0x3c, 0x3e, 0x26, 0x22)) });
    fc.assert(
      fc.property(fc.array(edit, { minLength: 1, maxLength: 20 }), fc.option(fc.nat({ max: valid.length }), { nil: undefined }), (edits, cut) => {
        const b = Uint8Array.from(valid);
        for (const { at, value } of edits) b[at] = value;
        readAll(cut === undefined ? b : b.subarray(0, cut));
      }),
      params,
    );
  });

  // Payloads that need real resources (huge row numbers, millions of nodes) are left to security/fuzz.mjs, where a worker enforces a budget; here a hang would stall the test run.
  const SNIPPETS = [
    '<![CDATA[', '<![CDATA[x]]>]]>', '<!--', '<?pi', '&#xFFFFFFFF;', '&#x110000;', '&#0;', '&#xD800;', '&bogus;', '&amp', '</row>', '</c>', '<c', '<row', '<is>', '</is>',
    '<c r="" s="-1" t="s"><v>-1</v></c>', '<c r="A0" s="99999999" t="s"><v>4294967296</v></c>', '<c r="a1" t="b"><v>x</v></c>', '<c r="ZZZ1" t="e"><v>#REF!</v></c>',
    '<c t="inlineStr"><is><r><t>x</t></r><rPh><t>y</t></rPh></is></c>', '<c t="d"><v>not a date</v></c>', '<c t="n"><v>1e999</v></c>', '<c t="str"><v></v></c>', '<v/>', '<row r="3"><c/></row>',
  ];
  it('a valid sheet part with hostile snippets spliced in', () => {
    const files = unzipSync(valid);
    const names = Object.keys(files).filter((n) => n.endsWith('.xml') || n.endsWith('.rels'));
    fc.assert(
      fc.property(fc.array(fc.record({ file: fc.nat({ max: names.length - 1 }), at: fc.double({ min: 0, max: 1, noNaN: true }), snippet: fc.constantFrom(...SNIPPETS) }), { minLength: 1, maxLength: 4 }), (edits) => {
        const mutated: Record<string, Uint8Array> = { ...files };
        for (const { file, at, snippet } of edits) {
          const name = names[file]!;
          const text = strFromU8(mutated[name]!);
          const i = Math.floor(at * text.length);
          mutated[name] = strToU8(text.slice(0, i) + snippet + text.slice(i));
        }
        readAll(zipSync(mutated));
      }),
      params,
    );
  });

  it('is deterministic about the error: the same bytes fail the same way twice', () => {
    fc.assert(
      fc.property(fc.uint8Array({ maxLength: 500 }), (b) => {
        const run = (): string => {
          try {
            for (const s of readXlsx(b).sheets) for (const _ of s.rows()) void _;
            return 'ok';
          } catch (e) {
            return (e as Error).message;
          }
        };
        expect(run()).toBe(run());
      }),
      params,
    );
  });
});
