/** Ceilings that keep a hostile package from exhausting memory or time. Every field is optional on input; omitted ones take the default. */
export interface XlsxLimits {
  /** Entries in the zip central directory. */
  maxEntries: number;
  /** Inflated bytes of any one entry. */
  maxEntryBytes: number;
  /** Inflated bytes summed over the distinct entries one workbook object reads. */
  maxTotalBytes: number;
  /** Largest inflated/compressed ratio allowed for an entry whose inflated size exceeds `ratioMinBytes`. */
  maxRatio: number;
  /** Entries up to this inflated size are exempt from `maxRatio`. */
  ratioMinBytes: number;
  /** Worksheets listed in the workbook. */
  maxSheets: number;
  /** Entries of the shared string table. */
  maxSharedStrings: number;
  /** Total characters of the shared string table. */
  maxSharedStringBytes: number;
  /** Entries of cellXfs (and of custom number formats). */
  maxCellXfs: number;
  /** Columns in a row; also bounds the column letters of a cell reference. */
  maxColumns: number;
  /** Highest row number, and the number of rows a sheet may yield. */
  maxRows: number;
  /** Empty rows a workbook may synthesize for row-number gaps, summed over its sheets (each sheet counted once, at its largest). Only gaps count: rows present in the part never do. */
  maxGapRows: number;
  /** Characters of one cell's text. */
  maxCellChars: number;
  /** Cells covered by merged ranges, per sheet (compat loader; each merged cell becomes a stored slot). */
  maxMergedCells: number;
  /** Pictures in a drawing part, and notes in a comments part, per sheet (compat loader). */
  maxDrawingObjects: number;
}

export const DEFAULT_LIMITS: Readonly<XlsxLimits> = Object.freeze({
  maxEntries: 10_000,
  maxEntryBytes: 256 * 1024 * 1024,
  maxTotalBytes: 1024 * 1024 * 1024,
  maxRatio: 200,
  ratioMinBytes: 1024 * 1024,
  maxSheets: 1000,
  maxSharedStrings: 5_000_000,
  maxSharedStringBytes: 256 * 1024 * 1024,
  maxCellXfs: 64_000,
  maxColumns: 16_384,
  maxRows: 1_048_576,
  maxGapRows: 4_194_304,
  maxCellChars: 32_767,
  maxMergedCells: 1_048_576,
  maxDrawingObjects: 100_000,
});

export interface ReadOptions {
  limits?: Partial<XlsxLimits>;
}

/** Thrown when the input is larger or more repetitive than the configured limits allow. `limit` names the `XlsxLimits` field that was hit. */
export class XlsxLimitError extends Error {
  constructor(
    readonly limit: keyof XlsxLimits,
    message: string,
  ) {
    super(`${message} (limit ${limit})`);
    this.name = 'XlsxLimitError';
  }
}

export function resolveLimits(partial?: Partial<XlsxLimits>): Readonly<XlsxLimits> {
  if (!partial) return DEFAULT_LIMITS;
  const out: XlsxLimits = { ...DEFAULT_LIMITS };
  for (const k of Object.keys(DEFAULT_LIMITS) as (keyof XlsxLimits)[]) {
    const v = partial[k];
    if (v === undefined) continue;
    if (typeof v !== 'number' || Number.isNaN(v) || v < 0) throw new TypeError(`limits.${k} must be a non-negative number`);
    out[k] = v;
  }
  return out;
}

/** True for the error Node's zlib raises when `maxOutputLength` is exceeded, and for the one the fallback inflater raises. */
export function isOutputOverflow(e: unknown): boolean {
  return e instanceof OutputOverflow || (typeof e === 'object' && e !== null && (e as { code?: unknown }).code === 'ERR_BUFFER_TOO_LARGE');
}

/** Raised by an inflater that stopped at its output cap; unzip turns it into an XlsxLimitError that says which cap. */
export class OutputOverflow extends Error {
  constructor() {
    super('inflated output exceeds the cap');
    this.name = 'OutputOverflow';
  }
}
