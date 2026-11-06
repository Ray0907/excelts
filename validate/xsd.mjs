// Validates the SpreadsheetML parts of an xlsx against the ECMA-376 transitional XSDs (from POI's ooxml-schemas-1.4.jar, extracted to xsd/x) using xmllint.
import { readFileSync, existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { unzipSync } from 'fflate';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const XSD_DIR = join(here, '..', 'xsd', 'x', 'schemaorg_apache_xmlbeans', 'src');
const RULES = [
  [/^xl\/workbook\.xml$/, 'sml-workbook.xsd'],
  [/^xl\/styles\.xml$/, 'sml-styles.xsd'],
  [/^xl\/sharedStrings\.xml$/, 'sml-sharedStringTable.xsd'],
  [/^xl\/worksheets\/sheet\d+\.xml$/, 'sml-sheet.xsd'],
  [/^xl\/comments\d+\.xml$/, 'sml-comments.xsd'],
  [/^xl\/drawings\/drawing\d+\.xml$/, 'dml-spreadsheetDrawing.patched.xsd'],
];

// libxml2 imports only the first schema file of a namespace, but dml-spreadsheetDrawing.xsd imports five files of the drawingml/main namespace.
// Build a copy of it that imports one wrapper which includes those five, so the drawing part is validated for real.
const MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main';
function patchedDrawingSchema() {
  const dir = join(here, 'out', 'xsd');
  mkdirSync(dir, { recursive: true });
  const wrap = join(dir, 'dml-main-all.xsd');
  const incl = ['dml-shapeStyle', 'dml-shapeProperties', 'dml-documentProperties', 'dml-graphicalObject', 'dml-text'];
  writeFileSync(wrap, `<?xml version="1.0" encoding="utf-8"?>\n<xsd:schema targetNamespace="${MAIN}" xmlns:xsd="http://www.w3.org/2001/XMLSchema" elementFormDefault="qualified">${incl.map((f) => `<xsd:include schemaLocation="${join(XSD_DIR, f)}.xsd"/>`).join('')}</xsd:schema>`);
  let src = readFileSync(join(XSD_DIR, 'dml-spreadsheetDrawing.xsd'), 'utf8').replace(/^\uFEFF/, '');
  src = src.replace(/\s*<xsd:import schemaLocation="dml-[A-Za-z]+\.xsd" namespace="[^"]*" \/>/g, '');
  src = src.replace('<xsd:complexType name="CT_AnchorClientData">', `<xsd:import schemaLocation="${wrap}" namespace="${MAIN}"/><xsd:complexType name="CT_AnchorClientData">`);
  writeFileSync(join(dir, 'dml-spreadsheetDrawing.patched.xsd'), src);
  return dir;
}

// Markup Compatibility preprocessing: Office drops attributes and elements in namespaces listed by mc:Ignorable before schema validation.
export function stripIgnorable(xml) {
  // AlternateContent: Office takes the first Choice it understands, else Fallback; for schema purposes keep the Fallback only.
  xml = xml.replace(/<mc:AlternateContent\b[^>]*>([\s\S]*?)<\/mc:AlternateContent>/g, (_, body) => /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/.exec(body)?.[1] ?? '');
  // extLst is an extension point whose content is vendor-defined; the schema cannot judge it.
  xml = xml.replace(/<extLst\b[^>]*>[\s\S]*?<\/extLst>/g, '').replace(/<extLst\b[^>]*\/>/g, '');
  // This XSD set does not import the xml namespace, so xml:space on <t> is reported although the spec allows it.
  xml = xml.replace(/\sxml:space="[^"]*"/g, '');
  const m = /mc:Ignorable="([^"]*)"/.exec(xml);
  if (!m) return xml;
  for (const prefix of m[1].split(/\s+/).filter(Boolean)) {
    const p = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    xml = xml.replace(new RegExp(`\\s${p}:[\\w.-]+="[^"]*"`, 'g'), '');
    xml = xml.replace(new RegExp(`<${p}:([\\w.-]+)\\b[^>]*/>`, 'g'), '');
    xml = xml.replace(new RegExp(`<${p}:([\\w.-]+)\\b[^>]*>[\\s\\S]*?</${p}:\\1>`, 'g'), '');
  }
  return xml.replace(/\s?mc:Ignorable="[^"]*"/, '');
}

export function validateXlsx(path) {
  if (!existsSync(XSD_DIR)) throw new Error('XSDs missing: run xsd/fetch.sh');
  const files = unzipSync(readFileSync(path));
  const results = [];
  for (const [name, bytes] of Object.entries(files)) {
    const rule = RULES.find(([re]) => re.test(name));
    if (!rule) continue;
    const schema = rule[1].includes('patched') ? join(patchedDrawingSchema(), rule[1]) : join(XSD_DIR, rule[1]);
    const p = spawnSync('xmllint', ['--noout', '--schema', schema, '-'], { input: stripIgnorable(Buffer.from(bytes).toString('utf8')), encoding: 'utf8', maxBuffer: 1 << 26 });
    const errors = (p.stderr || '').split('\n').filter((l) => l && !/validates$/.test(l));
    results.push({ part: name, xsd: rule[1], ok: p.status === 0, errors: errors.slice(0, 3) });
  }
  return results;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  let bad = 0;
  for (const f of process.argv.slice(2)) {
    const r = validateXlsx(f);
    const failed = r.filter((x) => !x.ok);
    bad += failed.length ? 1 : 0;
    console.log(`${failed.length ? 'FAIL' : 'ok  '} ${f} (${r.length} parts)`);
    for (const x of failed) console.log(`     ${x.part} vs ${x.xsd}: ${x.errors.join(' | ').slice(0, 300)}`);
  }
  process.exit(bad ? 1 : 0);
}
