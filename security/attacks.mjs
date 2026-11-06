import { bombAttacks } from './attacks-bombs.mjs';
import { xmlAttacks } from './attacks-xml.mjs';
import { zipAttacks } from './attacks-zip.mjs';

/**
 * Each attack: { id, group, note, build(opts) -> Buffer | null, modes?, check? }.
 * `modes` defaults to both consumer styles: 'stream' drops each row, 'collect' keeps every row of every sheet like `[...rows()]`.
 */
export function attacks(opts = {}) {
  return [...bombAttacks(opts), ...xmlAttacks(opts), ...zipAttacks(opts)];
}
