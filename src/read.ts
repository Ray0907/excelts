import type { CellValue } from './index.js';
import { readXfKinds } from './readStyles.js';
import { readSharedStrings } from './readSharedStrings.js';
import { scanRows, type GapBudget, type SheetContext } from './readSheet.js';
import { unzip } from './unzip.js';
import { readWorkbookParts } from './workbookParts.js';
import { defaultCompressor, type Compressor } from './zip.js';
import { resolveLimits, XlsxLimitError, type ReadOptions } from './limits.js';

export interface ReadSheet {
  name: string;
  /** Row N (1-based) is the Nth yielded array; gaps and trailing empties are not padded. */
  rows(): IterableIterator<CellValue[]>;
}
export interface ReadWorkbook {
  sheets: ReadSheet[];
}

const decoder = new TextDecoder();

/** @internal tests pass `fallbackCompressor` to exercise the non-Node path. */
export function readXlsx(bytes: Uint8Array, options: ReadOptions = {}, compressor: Compressor = defaultCompressor): ReadWorkbook {
  const limits = resolveLimits(options.limits);
  const zip = unzip(bytes, compressor, limits);
  const parts = readWorkbookParts(zip, limits);
  // The shared string table and the styles are parsed by the first cell that needs them: a numeric cell with no style reads neither.
  let sst: string[] | undefined;
  let kinds: Uint8Array | undefined;
  const ctx: SheetContext = {
    get sharedStrings() {
      if (!sst) {
        const bytes = zip.get(parts.sharedStringsPath);
        sst = bytes ? readSharedStrings(decoder.decode(bytes), limits) : [];
      }
      return sst;
    },
    get xfKind() {
      if (!kinds) {
        const bytes = zip.get(parts.stylesPath);
        kinds = bytes ? readXfKinds(decoder.decode(bytes), limits) : new Uint8Array(0);
      }
      return kinds;
    },
    date1904: parts.date1904,
    limits,
  };
  // Gap rows are charged per sheet at the most any one iteration of it produced, so reading a sheet again costs nothing extra.
  const gapPeak: number[] = parts.sheets.map(() => 0);
  let gapSum = 0;
  return {
    sheets: parts.sheets.map(({ name, path }, index) => ({
      name,
      rows() {
        const data = zip.get(path);
        if (!data) throw new Error(`invalid xlsx: missing worksheet part ${path}`);
        let mine = 0;
        const gaps: GapBudget = {
          add(count) {
            mine += count;
            if (mine > gapPeak[index]!) {
              gapSum += mine - gapPeak[index]!;
              gapPeak[index] = mine;
              if (gapSum > limits.maxGapRows) throw new XlsxLimitError('maxGapRows', `row numbers skip ${gapSum} empty rows in total`);
            }
          },
        };
        return scanRows(data, ctx, gaps);
      },
    })),
  };
}
