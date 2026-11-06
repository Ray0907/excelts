// XML 1.0 forbids these code points even when escaped as character references.
const ILLEGAL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;

const ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  '\r': '&#13;', // a bare CR is normalized to LF by parsers
};

export function escapeXml(s: string): string {
  return s.replace(ILLEGAL, '').replace(/[&<>"\r]/g, (c) => ESCAPES[c]!);
}

const CONTROL = '\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF';
// An underscore needs escaping when it would start a literal `_xHHHH_`; the final `_` may also be the first character of the escape that follows (`_xABCD` + control character).
const NO_ESCAPE_NEEDED = new RegExp(`[&<>"\\r${CONTROL}]`);
const TEXT_SPECIAL = new RegExp(`[&<>"\\r${CONTROL}]|_(?=x[0-9a-fA-F]{4}(?:_|[${CONTROL}]))`, 'g');

/** Escapes element text of a string item (ST_Xstring): control characters XML cannot carry become `_xHHHH_`, and a literal `_xHHHH_` has its underscore escaped, so every string reads back unchanged. */
export function escapeText(s: string): string {
  // One cheap class scan plus a native substring search settles the common case; the lookahead alternation below is several times slower per character.
  if (!NO_ESCAPE_NEEDED.test(s) && s.indexOf('_x') < 0) return s;
  return s.replace(TEXT_SPECIAL, (c) => {
    if (c === '_') return '_x005F_';
    const code = c.charCodeAt(0);
    return ESCAPES[c] ?? `_x${code.toString(16).toUpperCase().padStart(4, '0')}_`;
  });
}

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
