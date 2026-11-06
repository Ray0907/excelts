import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Workbook, XlsxLimitError } from '../../src/compat/index.js';

const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));

async function roundTrip(wb: Workbook): Promise<Workbook> {
  return new Workbook().xlsx.load(await wb.xlsx.writeBuffer());
}

describe('pictures', () => {
  it('a cell range, a one-cell anchor with a link, and a background survive a round trip', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'PNG' });
    const ws = wb.addWorksheet('S');
    ws.addImage(id, 'C3:E6');
    ws.addImage(id, { tl: { col: 0.5, row: 1.25 }, ext: { width: 100, height: 50 }, hyperlinks: { hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'go' } });
    ws.addBackgroundImage(id);
    const out = await roundTrip(wb);
    const s = out.getWorksheet('S')!;
    const [a, b] = s.getImages();
    expect([a!.range.tl.col, a!.range.tl.row, a!.range.br!.col, a!.range.br!.row]).toEqual([2, 2, 5, 6]);
    expect(b!.range.ext).toEqual({ width: 100, height: 50 });
    expect(b!.range.tl.col).toBeCloseTo(0.5, 2);
    expect(b!.range.hyperlinks).toEqual({ hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'go' });
    expect(Buffer.compare(Buffer.from(out.getImage(a!.imageId)!.buffer!), Buffer.from(PNG))).toBe(0);
    expect(out.getImage(s.getBackgroundImageId()!)!.extension).toBe('png');
    expect(out.media).toHaveLength(1);
  });

  it('editAs on a one-cell anchor round-trips through a private ignorable attribute; absent or default writes none', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    const ws = wb.addWorksheet('S');
    ws.addImage(id, { tl: { col: 0, row: 0 }, ext: { width: 10, height: 10 }, editAs: 'absolute' });
    ws.addImage(id, { tl: { col: 1, row: 1 }, ext: { width: 10, height: 10 } });
    ws.addImage(id, { tl: { col: 2, row: 2 }, ext: { width: 10, height: 10 }, editAs: 'bogus' });
    ws.addImage(id, 'C3:E6');
    ws.addImage(id, { tl: { col: 0, row: 8 }, br: { col: 2, row: 10 }, editAs: 'twoCell' });
    const bytes = await wb.xlsx.writeBuffer();
    const xml = strFromU8(unzipSync(new Uint8Array(bytes))['xl/drawings/drawing1.xml']!);
    expect(xml.match(/xn:editAs="/g)).toHaveLength(1);
    expect(xml).toContain('xmlns:xn="urn:xlsx-next:1"');
    expect(xml).toContain('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"');
    expect(xml).toContain('mc:Ignorable="xn"');
    expect(xml).toContain('<xdr:oneCellAnchor xn:editAs="absolute">');
    expect(xml).toContain('<xdr:twoCellAnchor editAs="twoCell">');
    const got = (await new Workbook().xlsx.load(bytes)).getWorksheet('S')!.getImages().map((i) => i.range.editAs);
    expect(got).toEqual(['absolute', 'oneCell', 'oneCell', 'oneCell', 'twoCell']);
  });

  it('a drawing with no private attribute has no extra namespaces', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    wb.addWorksheet('S').addImage(id, { tl: { col: 1, row: 1 }, ext: { width: 10, height: 10 } });
    const xml = strFromU8(unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()))['xl/drawings/drawing1.xml']!);
    expect(xml).not.toMatch(/xn:|mc:|Ignorable/);
  });

  it('the picture bytes are written once per image, with a lower-case extension and content type', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'JPEG' });
    wb.addWorksheet('A').addImage(id, 'A1:B2');
    wb.addWorksheet('B').addImage(id, 'A1:B2');
    const files = unzipSync(await wb.xlsx.writeBuffer());
    expect(Object.keys(files).filter((n) => n.startsWith('xl/media/'))).toEqual(['xl/media/image1.jpeg']);
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('<Default Extension="jpeg" ContentType="image/jpeg"/>');
    expect(strFromU8(files['xl/drawings/drawing2.xml']!)).toContain('r:embed="rId1"');
  });

  it('sheet elements follow the schema order: drawing, legacyDrawing, picture', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    const ws = wb.addWorksheet('S');
    ws.addImage(id, 'A1:B2');
    ws.addBackgroundImage(id);
    ws.getCell('A1').note = 'n';
    ws.getCell('A2').value = { text: 'x', hyperlink: 'https://example.com' };
    const xml = strFromU8(unzipSync(await wb.xlsx.writeBuffer())['xl/worksheets/sheet1.xml']!);
    const at = (t: string): number => xml.indexOf(`<${t} `);
    expect(at('hyperlinks')).toBeLessThan(at('drawing'));
    expect(at('drawing')).toBeLessThan(at('legacyDrawing'));
    expect(at('legacyDrawing')).toBeLessThan(at('picture'));
    const rels = strFromU8(unzipSync(await wb.xlsx.writeBuffer())['xl/worksheets/_rels/sheet1.xml.rels']!);
    expect([...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1])).toEqual(['rId1', 'rId2', 'rId3', 'rId4', 'rId5']);
  });

  it('a missing file rejects the write; an unknown image id rejects it too', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ filename: '/nonexistent/dir/image.jpg', extension: 'jpeg' });
    wb.addWorksheet('S').addImage(id, 'B2:D6');
    await expect(wb.xlsx.writeBuffer()).rejects.toThrow();
    const wb2 = new Workbook();
    wb2.addWorksheet('S').addImage(7, 'B2:D6');
    await expect(wb2.xlsx.writeBuffer()).rejects.toThrow(/image 7/);
  });

  it('rejects a bad range and a bad extension', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    expect(() => ws.addImage(0, { tl: { col: 1, row: 1 } } as never)).toThrow(TypeError);
    const id = wb.addImage({ buffer: PNG, extension: '../x' });
    ws.addImage(id, 'A1:B2');
    await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/extension/);
  });

  it('works in the streaming writer: picture, background and a note set before the sheet is committed', async () => {
    const { stream } = await import('../../src/compat/index.js');
    const chunks: Uint8Array[] = [];
    const sink = { write: (b: Uint8Array, cb?: (e?: Error | null) => void) => { chunks.push(b); cb?.(); return true; }, end: (cb?: () => void) => { cb?.(); } };
    const wb = new stream.xlsx.WorkbookWriter({ stream: sink });
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.getCell('A1').note = 'streamed note';
    ws.addImage(id, 'B2:C3');
    ws.addBackgroundImage(id);
    ws.getRow(1).commit();
    await wb.commit();
    const files = unzipSync(chunks[0]!);
    expect(strFromU8(files['xl/comments1.xml']!)).toContain('streamed note');
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/media/image1.png', 'xl/drawings/vmlDrawing1.vml']));
    const xml = strFromU8(files['xl/worksheets/sheet1.xml']!);
    expect(xml.indexOf('<drawing ')).toBeLessThan(xml.indexOf('<legacyDrawing '));
    const back = await new Workbook().xlsx.load(chunks[0]!);
    expect(back.getWorksheet('S')!.getCell('A1').note).toBe('streamed note');
    expect(back.getWorksheet('S')!.getImages()).toHaveLength(1);
  });
});

describe('loading pictures', () => {
  async function withDrawing(inner: string): Promise<Uint8Array> {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.addImage(id, 'A1:B2');
    const files = unzipSync(await wb.xlsx.writeBuffer());
    files['xl/drawings/drawing1.xml'] = strToU8(
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + inner + '</xdr:wsDr>',
    );
    return zipSync(files);
  }

  it('a chart or shape in a drawing is reported as lost, and saving refuses unless told to ignore it', async () => {
    const bytes = await withDrawing('<xdr:twoCellAnchor><xdr:from/><xdr:to/><xdr:graphicFrame/><xdr:clientData/></xdr:twoCellAnchor>');
    const wb = await new Workbook().xlsx.load(bytes);
    expect(wb.getWorksheet('S')!.getImages()).toHaveLength(0);
    await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/drawings \(charts or shapes\)/);
    expect((await wb.xlsx.writeBuffer({ ignoreUnsupported: true })).length).toBeGreaterThan(100);
  });

  it('an absolute anchor is reported as lost', async () => {
    const wb = await new Workbook().xlsx.load(await withDrawing('<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/><xdr:pic/><xdr:clientData/></xdr:absoluteAnchor>'));
    await expect(wb.xlsx.writeBuffer()).rejects.toThrow(/absolute-anchored/);
  });

  it('maxDrawingObjects bounds pictures and notes', async () => {
    const wb = new Workbook();
    const id = wb.addImage({ buffer: PNG, extension: 'png' });
    const ws = wb.addWorksheet('S');
    for (let i = 1; i <= 5; i++) {
      ws.addImage(id, `A${i}:B${i + 1}`);
      ws.getCell(i, 3).note = `n${i}`;
    }
    const bytes = await wb.xlsx.writeBuffer();
    await expect(new Workbook().xlsx.load(bytes, { limits: { maxDrawingObjects: 4 } })).rejects.toBeInstanceOf(XlsxLimitError);
    expect((await new Workbook().xlsx.load(bytes)).getWorksheet('S')!.getImages()).toHaveLength(5);
  });

  it('the picture bytes are read only when asked for', async () => {
    const out = await new Workbook().xlsx.load(await (async () => {
      const wb = new Workbook();
      wb.addWorksheet('S').addImage(wb.addImage({ buffer: PNG, extension: 'png' }), 'A1:B2');
      return wb.xlsx.writeBuffer();
    })());
    const d = Object.getOwnPropertyDescriptor(out.media[0], 'buffer')!;
    expect(typeof d.get).toBe('function');
    expect(out.media[0]!.buffer!.length).toBe(PNG.length);
  });
});

describe('notes', () => {
  it('a string note and a full note survive, with the placement details', async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('B2').note = 'five';
    const full = {
      texts: [{ font: { size: 12, color: { argb: 'FFFF6600' }, name: 'Calibri', scheme: 'minor' }, text: 'seven & <more>\nline' }],
      margins: { insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35] },
      protection: { locked: 'False', lockText: 'False' },
      editAs: 'twoCells',
    };
    ws.getCell('D2').note = full;
    ws.getCell('D2').value = 7;
    const s = (await roundTrip(wb)).getWorksheet('S')!;
    expect(s.getCell('B2').note).toBe('five');
    expect(s.getCell('D2').note).toEqual(full);
    expect(s.getCell('D2').value).toBe(7);
  });

  it('notes on the same cells in two sheets keep their own text and unique shape ids', async () => {
    const wb = new Workbook();
    wb.addWorksheet('A').getCell('A1').note = 'a';
    wb.addWorksheet('B').getCell('A1').note = 'b';
    const files = unzipSync(await wb.xlsx.writeBuffer());
    expect(strFromU8(files['xl/drawings/vmlDrawing1.vml']!)).toContain('_x0000_s1025');
    expect(strFromU8(files['xl/drawings/vmlDrawing2.vml']!)).toContain('_x0000_s2049');
    const out = await new Workbook().xlsx.load(await wb.xlsx.writeBuffer());
    expect([out.getWorksheet('A')!.getCell('A1').note, out.getWorksheet('B')!.getCell('A1').note]).toEqual(['a', 'b']);
  });

  it('notes move with their cells when rows or columns are spliced', () => {
    const ws = new Workbook().addWorksheet('S');
    ws.getCell('B2').value = 1;
    ws.getCell('B2').note = 'x';
    ws.getCell('C3').note = 'y';
    ws.insertRow(1, ['new']);
    expect(ws.getCell('B3').note).toBe('x');
    expect(ws.getCell('B2').note).toBeUndefined();
    ws.spliceColumns(1, 0, ['ins']);
    expect(ws.getCell('C3').note).toBe('x');
    expect(ws.getCell('D4').note).toBe('y');
    ws.spliceColumns(3, 1);
    expect(ws.getCell('C3').note).toBeUndefined();
    ws.spliceRows(4, 1);
    expect(ws.getCell('C4').note).toBeUndefined();
  });

  it('a note can be removed, and bad input is refused', () => {
    const c = new Workbook().addWorksheet('S').getCell('A1');
    c.note = 'x';
    c.note = undefined;
    expect(c.note).toBeUndefined();
    expect(() => { c.note = 5 as never; }).toThrow(TypeError);
  });

  it('shapes in a legacy drawing that are not notes are reported as lost', async () => {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').note = 'n';
    const files = unzipSync(await wb.xlsx.writeBuffer());
    files['xl/drawings/vmlDrawing1.vml'] = strToU8(strFromU8(files['xl/drawings/vmlDrawing1.vml']!).replace('</xml>', '<v:shape id="b"><x:ClientData ObjectType="Button"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>'));
    const out = await new Workbook().xlsx.load(zipSync(files));
    expect(out.getWorksheet('S')!.getCell('A1').note).toBe('n');
    await expect(out.xlsx.writeBuffer()).rejects.toThrow(/legacy drawing shapes/);
  });
});
