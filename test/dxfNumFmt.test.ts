import { expect, it } from 'vitest';
import { readXfKinds, XF } from '../src/readStyles.js';

it('a numFmt inside dxfs does not redefine a built-in id used by cellXfs', () => {
  const xml = '<styleSheet><cellXfs count="1"><xf numFmtId="14"/></cellXfs><dxfs count="1"><dxf><numFmt numFmtId="14" formatCode="0.00%"/></dxf></dxfs></styleSheet>';
  expect(Array.from(readXfKinds(xml))).toEqual([XF.DATE]);
});
