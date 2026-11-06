---
name: verify-text-roundtrip
description: Use when a change touches src/xml.ts, src/xmlRead.ts, src/sharedStrings.ts or src/readSharedStrings.ts (cell text escaping and decoding). Checks that every string written to an xlsx reads back unchanged.
---

# Text round trip check

**Property.** For every string `s`, `decodeText(escapeText(s)) === s`, and a string written with `writeXlsx` reads back identical with `readXlsx`.
**Harm if broken.** Silent data corruption: a user's text with control characters, CR, or a literal `_x005F_`-style sequence comes back different, with no error.

## Run

```
npx tsc -p . && node verify/text-roundtrip.mjs            # quick, about 1 s
node verify/text-roundtrip.mjs --if-changed               # quick, but skips when the covered files are unchanged since the sync commit
node verify/text-roundtrip.mjs --full                     # about 10 s; before merging a change to the covered files
```

You do not have to run `--full` on every edit. Run `--quick` while iterating, `--full` once before committing a change to the covered files, and nothing when `--if-changed` says skip.

## What it covers

Exhaustive over every sequence of tokens from three small alphabets (underscore/hex/control, markup/newlines, CDATA/non-characters), up to 7 tokens quick and 8 full, through the real `escapeText` and `decodeText` from `dist/`. Plus 13 hand-picked strings through the real `writeXlsx` and `readXlsx`.
It does NOT cover: attribute escaping (`escapeXml`), lone surrogates, strings over 8 tokens, characters outside the alphabets, rich-text runs, phonetic runs, or any other reader's view of the file. A pass means no counterexample exists inside these bounds.

Covered files: `src/xml.ts`, `src/xmlRead.ts`, `src/sharedStrings.ts`, `src/readSharedStrings.ts`.
Sync commit: `b66d880` (the commit at which the script was confirmed to describe these files; the later sharedStrings.ts change only added a rich-text path, which this check does not cover). Move it forward, in `verify/text-roundtrip.mjs`, only after you confirm the alphabets and the sample list still describe the code at the new commit.

## Reading a failure

The output prints the input, the escaped form and what came back. Tell it as steps: "text `X` was written as `Y`, the reader returned `Z`." That is a real bug when the covered code changed and the escaped form is not what the OOXML rule says: a control character must be `_xHHHH_`, and a literal `_xHHHH_` must have its underscore escaped as `_x005F_` (ECMA-376 ST_Xstring). It is a modeling error only if the failing token is outside what the writer promises to accept; then change the alphabet and say why in your commit message, never loosen the comparison.

## When you change the rules on purpose

Update the alphabets so a new special sequence has enough tokens to appear, update `SAMPLES`, and show the check failing on the old code before you move the sync commit.

## Proven to catch breaks

Three mutations of `dist/xml.js` were each caught in quick mode: dropping the control-character case of the underscore lookahead (counterexample `_x0500` + U+0001), dropping the underscore lookahead entirely (`_x0055` + U+0001), and removing the CR escape (`_x00F5` + CR). Redo one of these after any edit to the escape rules, to confirm the check still fails.
