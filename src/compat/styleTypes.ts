export interface Color { argb?: string; theme?: number; tint?: number; indexed?: number; auto?: boolean }
export interface Font {
  name?: string; size?: number; family?: number; scheme?: string; charset?: number; color?: Color;
  bold?: boolean; italic?: boolean; underline?: boolean | string; strike?: boolean; outline?: boolean; shadow?: boolean;
  condense?: boolean; extend?: boolean; vertAlign?: string;
}
export interface BorderEdge { style?: string; color?: Color }
export interface Border {
  top?: BorderEdge; left?: BorderEdge; bottom?: BorderEdge; right?: BorderEdge;
  diagonal?: BorderEdge & { up?: boolean; down?: boolean };
}
export interface PatternFill { type: 'pattern'; pattern: string; fgColor?: Color; bgColor?: Color }
export interface GradientFill {
  type: 'gradient'; gradient: 'angle' | 'path'; degree?: number;
  center?: { left?: number; top?: number; right?: number; bottom?: number };
  stops: { position: number; color: Color }[];
}
export type Fill = PatternFill | GradientFill;
export interface Alignment {
  horizontal?: string; vertical?: string; wrapText?: boolean; shrinkToFit?: boolean; indent?: number;
  textRotation?: number | 'vertical'; readingOrder?: string | number;
}
export interface Protection { locked?: boolean; hidden?: boolean }
export interface Style {
  numFmt?: string; font?: Font; alignment?: Alignment; border?: Border; fill?: Fill; protection?: Protection;
}
export type StyleKey = keyof Style;
export const STYLE_KEYS: readonly StyleKey[] = ['numFmt', 'font', 'alignment', 'border', 'fill', 'protection'];
