import { NotSupportedError } from './errors.js';
import type { Style, StyleKey } from './styleTypes.js';
import { STYLE_KEYS } from './styleTypes.js';

const clone = <T>(v: T): T => (v === undefined ? v : (JSON.parse(JSON.stringify(v)) as T));

/** A read-only view for objects nested below a style component: assigning into them cannot be written through, so it fails loudly. */
function guard(v: unknown, path: string): unknown {
  if (v === null || typeof v !== 'object') return v;
  return new Proxy(v as object, {
    get(t, k, r) {
      const x = Reflect.get(t, k, r);
      return typeof k === 'string' && x && typeof x === 'object' ? guard(x, `${path}.${k}`) : x;
    },
    set(_t, k) {
      throw new NotSupportedError(`mutating ${path}.${String(k)} in place`, `assign the whole object instead, e.g. ${path.split('.')[0]}.${path.split('.')[1] ?? 'x'} = { ... }`);
    },
    deleteProperty(_t, k) {
      throw new NotSupportedError(`deleting ${path}.${String(k)} in place`, 'assign a whole replacement object instead');
    },
  });
}

/** `cell.font`, `cell.fill` ...: assigning a property (cell.font.bold = true) rewrites the cell's style; one level deep. */
export function componentView<T extends object>(cur: T | undefined, name: string, write: (next: T | undefined) => void): T | undefined {
  if (cur === undefined) return undefined;
  const target = clone(cur) as Record<string, unknown>;
  return new Proxy(target, {
    get(t, k, r) {
      const x = Reflect.get(t, k, r);
      return typeof k === 'string' && x && typeof x === 'object' ? guard(x, `${name}.${k}`) : x;
    },
    set(t, k, v) {
      if (typeof k !== 'string') return Reflect.set(t, k, v);
      const next = { ...t, [k]: clone(v) };
      write(next as T);
      t[k] = clone(v);
      return true;
    },
    deleteProperty(t, k) {
      const next = { ...t };
      delete next[k as string];
      write(next as T);
      return Reflect.deleteProperty(t, k);
    },
  }) as T;
}

/** `cell.style`: a plain-looking object whose component assignments write through. */
export function styleView(read: () => Style, write: (key: StyleKey, v: unknown) => void): Style {
  const target = {} as Record<string, unknown>;
  const snap = read();
  for (const k of STYLE_KEYS) if (snap[k] !== undefined) target[k] = clone(snap[k]);
  return new Proxy(target, {
    get(t, k, r) {
      const x = Reflect.get(t, k, r);
      if (typeof k === 'string' && x && typeof x === 'object' && (STYLE_KEYS as readonly string[]).includes(k)) {
        return componentView(x as object, k, (n) => {
          write(k as StyleKey, n);
          t[k] = n;
        });
      }
      return x;
    },
    set(t, k, v) {
      if (typeof k === 'string' && (STYLE_KEYS as readonly string[]).includes(k)) {
        write(k as StyleKey, v);
        if (v === undefined) delete t[k];
        else t[k] = clone(v);
        return true;
      }
      return Reflect.set(t, k, v);
    },
    deleteProperty(t, k) {
      if (typeof k === 'string' && (STYLE_KEYS as readonly string[]).includes(k)) write(k as StyleKey, undefined);
      return Reflect.deleteProperty(t, k);
    },
  }) as Style;
}

export { clone };
