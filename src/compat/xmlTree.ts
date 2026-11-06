import { decodeEntities, decodeText } from '../xmlRead.js';

export interface XNode {
  name: string;
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

const local = (n: string): string => {
  const i = n.indexOf(':');
  return i < 0 || n.startsWith('xml:') ? n : n.slice(i + 1);
};

const ATTR = /([^\s=/>"']+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** A small DOM for the little parts of a package (styles, workbook, rels, sheet tails). Prefixes are dropped; entities are decoded. Not for sheetData. */
export function parseXml(xml: string, maxNodes = 5_000_000): XNode {
  const root: XNode = { name: '#root', attrs: {}, children: [], text: '' };
  const stack: XNode[] = [root];
  let nodes = 0;
  let i = 0;
  const n = xml.length;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    const top = stack[stack.length - 1]!;
    if (lt > i && top !== root) top.text += xml.slice(i, lt);
    if (xml.startsWith('<!--', lt)) {
      const e = xml.indexOf('-->', lt + 4);
      i = e < 0 ? n : e + 3;
    } else if (xml.startsWith('<![CDATA[', lt)) {
      const e = xml.indexOf(']]>', lt + 9);
      const end = e < 0 ? n : e;
      top.text += xml.slice(lt + 9, end).replace(/&/g, '&amp;').replace(/</g, '&lt;');
      i = e < 0 ? n : e + 3;
    } else if (xml.charCodeAt(lt + 1) === 63 || xml.charCodeAt(lt + 1) === 33) {
      const e = xml.indexOf('>', lt);
      i = e < 0 ? n : e + 1;
    } else if (xml.charCodeAt(lt + 1) === 47) {
      const e = xml.indexOf('>', lt);
      if (stack.length > 1) stack.pop();
      i = e < 0 ? n : e + 1;
    } else {
      // find the end of the tag, skipping quoted '>'
      let j = lt + 1;
      let q = 0;
      for (; j < n; j++) {
        const c = xml.charCodeAt(j);
        if (q) {
          if (c === q) q = 0;
        } else if (c === 34 || c === 39) q = c;
        else if (c === 62) break;
      }
      const selfClose = xml.charCodeAt(j - 1) === 47;
      const body = xml.slice(lt + 1, selfClose ? j - 1 : j);
      let k = 0;
      while (k < body.length && !/[\s]/.test(body[k]!)) k++;
      const node: XNode = { name: local(body.slice(0, k)), attrs: {}, children: [], text: '' };
      if (++nodes > maxNodes) throw new Error('xml has too many elements');
      if (k < body.length) {
        ATTR.lastIndex = k;
        let m: RegExpExecArray | null;
        const rest = body;
        while ((m = ATTR.exec(rest))) {
          const key = local(m[1]!);
          if (!(key in node.attrs)) node.attrs[key] = decodeEntities(m[2] ?? m[3] ?? '');
        }
      }
      top.children.push(node);
      if (!selfClose) stack.push(node);
      i = j + 1;
    }
  }
  return root.children[0] ?? root;
}

/** Decoded text of a node (entities and _xHHHH_ escapes). */
export function textOf(n: XNode): string {
  return decodeText(n.text);
}

export function child(n: XNode, name: string): XNode | undefined {
  return n.children.find((c) => c.name === name);
}
export function kids(n: XNode | undefined, name: string): XNode[] {
  return n ? n.children.filter((c) => c.name === name) : [];
}
