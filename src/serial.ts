const MS_PER_DAY = 86400000;
const UNIX_EPOCH_SERIAL = 25569; // 1970-01-01 in the 1900 date system
const FIRST_REAL_DATE = 2; // linear serial of 1900-01-01; below it only times of day and durations exist
const FIRST_LINEAR_SERIAL = 61; // 1900-03-01; Excel counts a nonexistent 1900-02-29 before it

/** Excel serial (1900 system) of a UTC instant, inverse of `serialToDate`. */
export function dateToSerial1900(ms: number): number {
  const linear = ms / MS_PER_DAY + UNIX_EPOCH_SERIAL;
  return linear >= FIRST_REAL_DATE && linear < FIRST_LINEAR_SERIAL ? linear - 1 : linear;
}
