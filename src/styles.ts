import { XML_DECLARATION, escapeXml } from './xml.js';

const FIRST_CUSTOM_NUM_FMT_ID = 164;
const BUILTIN_DATE_TIME_NUM_FMT_ID = 22; // m/d/yy h:mm

export class Styles {
  /** xf 0 is the default (General); xf 1 is the built-in date-time format. */
  static readonly DEFAULT_XF = 0;
  static readonly DATE_XF = 1;

  private readonly xfByFormat = new Map<string, number>();
  private readonly numFmtIds: number[] = [0, BUILTIN_DATE_TIME_NUM_FMT_ID];
  private readonly customCodes: string[] = [];

  xfForFormat(code: string | undefined): number {
    if (!code) return Styles.DEFAULT_XF;
    let xf = this.xfByFormat.get(code);
    if (xf === undefined) {
      xf = this.numFmtIds.length;
      this.xfByFormat.set(code, xf);
      this.numFmtIds.push(FIRST_CUSTOM_NUM_FMT_ID + this.customCodes.length);
      this.customCodes.push(code);
    }
    return xf;
  }

  toXml(): string {
    const numFmts = this.customCodes.length
      ? `<numFmts count="${this.customCodes.length}">` +
        this.customCodes
          .map((c, i) => `<numFmt numFmtId="${FIRST_CUSTOM_NUM_FMT_ID + i}" formatCode="${escapeXml(c)}"/>`)
          .join('') +
        '</numFmts>'
      : '';
    const xfs = this.numFmtIds
      .map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0"${id ? ' applyNumberFormat="1"' : ''}/>`)
      .join('');
    return (
      XML_DECLARATION +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmts +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.numFmtIds.length}">${xfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'
    );
  }
}
