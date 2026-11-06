import { StyleBook } from './styleBook.js';
import type { Worksheet } from './worksheet.js';

let fallback: StyleBook | undefined;

/** The style book of a sheet's workbook; a detached sheet (a hand-made stand-in with no workbook) shares one scratch book. */
export function bookOf(ws: Worksheet): StyleBook {
  return ws.workbook ? ws.workbook._styles : (fallback ??= new StyleBook());
}
