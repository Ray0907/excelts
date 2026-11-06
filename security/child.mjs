import { readFileSync, writeFileSync } from 'node:fs';
import { attacks } from './attacks.mjs';
import { fuzzChunk } from './fuzz.mjs';
import { pool, runCase } from './harness.mjs';

const [cmd, ...rest] = process.argv.slice(2);
const flags = Object.fromEntries(rest.filter((a) => a.startsWith('--')).map((a) => a.slice(2).split('=')));
const positional = rest.filter((a) => !a.startsWith('--'));

if (cmd === 'build') {
  const [id, out] = positional;
  const a = attacks({ bombMb: Number(flags['bomb-mb'] ?? 1024) }).find((x) => x.id === id);
  const bytes = a?.build();
  writeFileSync(out, bytes ?? Buffer.alloc(0));
  process.stdout.write(`BUILT ${bytes ? bytes.length : 'null'}\n`);
} else if (cmd === 'run') {
  const [file, mode, id] = positional;
  const check = id ? attacks({ bombMb: Number(flags['bomb-mb'] ?? 1024) }).find((x) => x.id === id)?.check : undefined;
  const r = await runCase(new Uint8Array(readFileSync(file)), { bundle: flags.bundle, mode, trackRss: true, check });
  process.stdout.write(`RESULT ${JSON.stringify(r)}\n`);
} else if (cmd === 'fuzz') {
  await fuzzChunk({ seed: Number(flags.seed), indices: JSON.parse(flags.indices), jobs: Number(flags.jobs), bundle: flags.bundle, mode: flags.mode, pool, runCase });
}
process.exit(0);
