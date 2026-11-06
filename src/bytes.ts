/** Search and text access over a sheet part held as bytes. */
export interface ByteView {
  /** Bytes the scanner indexes directly; a Buffer view in Node, the input otherwise. */
  readonly data: Uint8Array;
  indexOfSeq(pattern: Uint8Array, from: number): number;
  lastIndexOfSeq(pattern: Uint8Array): number;
  sliceText(from: number, to: number): string;
}

export const ascii = (s: string): Uint8Array => Uint8Array.from(s, (c) => c.charCodeAt(0));

interface Utf8Buffer extends Uint8Array {
  indexOf(value: Uint8Array | number, byteOffset?: number): number;
  lastIndexOf(value: Uint8Array | number, byteOffset?: number): number;
  utf8Slice(start: number, end: number): string;
}

class NodeView implements ByteView {
  readonly data: Utf8Buffer;
  constructor(bytes: Uint8Array) {
    this.data = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength) as unknown as Utf8Buffer;
  }
  indexOfSeq(pattern: Uint8Array, from: number): number {
    return this.data.indexOf(pattern, from);
  }
  lastIndexOfSeq(pattern: Uint8Array): number {
    return this.data.lastIndexOf(pattern);
  }
  sliceText(from: number, to: number): string {
    return this.data.utf8Slice(from, to);
  }
}

const decoder = new TextDecoder();

class PlainView implements ByteView {
  constructor(readonly data: Uint8Array) {}
  indexOfSeq(pattern: Uint8Array, from: number): number {
    const data = this.data;
    const first = pattern[0]!;
    const last = data.length - pattern.length;
    let at = data.indexOf(first, from);
    outer: while (at >= 0 && at <= last) {
      for (let k = 1; k < pattern.length; k++) {
        if (data[at + k] !== pattern[k]) {
          at = data.indexOf(first, at + 1);
          continue outer;
        }
      }
      return at;
    }
    return -1;
  }
  lastIndexOfSeq(pattern: Uint8Array): number {
    const data = this.data;
    outer: for (let at = data.length - pattern.length; at >= 0; at--) {
      for (let k = 0; k < pattern.length; k++) if (data[at + k] !== pattern[k]) continue outer;
      return at;
    }
    return -1;
  }
  sliceText(from: number, to: number): string {
    return decoder.decode(this.data.subarray(from, to));
  }
}

/** Picked once per sheet scan: Node's native Buffer search and utf8Slice when a global Buffer exists, a plain loop and TextDecoder otherwise. */
export function viewOf(bytes: Uint8Array): ByteView {
  return typeof Buffer === 'function' ? new NodeView(bytes) : new PlainView(bytes);
}
