import type { EditorNumberingFormat } from './model.js';
import {
  allNativeParagraphs,
  findNativeParagraphLocation,
  type NativeWordDocument,
  type NativeWordParagraphNode,
} from './native-model.js';

export type NativeWordListType = 'bullet' | 'numbered';

export function nativeListType(numbering: EditorNumberingFormat | undefined): NativeWordListType | null {
  if (!numbering) return null;
  return numbering.numFmt === 'bullet' ? 'bullet' : 'numbered';
}

export function createNativeListNumbering(
  doc: NativeWordDocument,
  type: NativeWordListType,
  ilvl = 0,
): EditorNumberingFormat {
  const level = clampListLevel(ilvl);
  return {
    numId: createNativeNumId(doc, type),
    ilvl: level,
    numFmt: type === 'bullet' ? 'bullet' : numberingFormatForLevel(level),
    lvlText: type === 'bullet' ? bulletForLevel(level) : `%${level + 1}.`,
    start: 1,
  };
}

export function nativeNumberingAtLevel(
  numbering: EditorNumberingFormat,
  ilvl: number,
): EditorNumberingFormat {
  const level = clampListLevel(ilvl);
  const type = nativeListType(numbering) ?? 'numbered';
  return {
    ...numbering,
    ilvl: level,
    numFmt: type === 'bullet' ? 'bullet' : numberingFormatForLevel(level),
    lvlText: type === 'bullet' ? bulletForLevel(level) : `%${level + 1}.`,
    start: numbering.start ?? 1,
  };
}

export function refreshNativeNumberingLabels(doc: NativeWordDocument): void {
  const counters = new Map<string, number>();
  for (const paragraph of allNativeParagraphs(doc)) {
    const numbering = paragraph.format?.numbering;
    if (!paragraph.format?.hasNumbering || !numbering) continue;
    const key = `${paragraph.story}:${numbering.numId}:${numbering.ilvl}`;
    if (nativeListType(numbering) === 'bullet') {
      numbering.label = normalizeBulletLabel(numbering.lvlText);
      continue;
    }
    const current = (counters.get(key) ?? ((numbering.start ?? 1) - 1)) + 1;
    counters.set(key, current);
    numbering.label = formatNumberingLabel(current, numbering.numFmt, numbering.lvlText);
  }
}

export function nativeListIndent(ilvl: number): { left: number; firstLine: number } {
  const level = clampListLevel(ilvl);
  return { left: (level + 1) * 720, firstLine: -360 };
}

export function adjacentNativeListNumbering(
  doc: NativeWordDocument,
  paraIds: string[],
  type: NativeWordListType,
): EditorNumberingFormat | null {
  const selected = new Set(paraIds);
  for (const paraId of paraIds) {
    const location = findNativeParagraphLocation(doc, paraId);
    if (!location) continue;
    const candidates = [location.container[location.index - 1], location.container[location.index + 1]];
    for (const candidate of candidates) {
      if (candidate?.kind !== 'paragraph' || selected.has(candidate.paraId)) continue;
      const numbering = reusableNativeNumbering(candidate, type);
      if (numbering) return { ...numbering };
    }
  }
  return null;
}

function createNativeNumId(doc: NativeWordDocument, type: NativeWordListType): string {
  const used = new Set(
    allNativeParagraphs(doc)
      .map((paragraph) => paragraph.format?.numbering?.numId)
      .filter((value): value is string => !!value),
  );
  let suffix = 1;
  let candidate = '';
  do {
    candidate = `coco-${type}-${suffix}`;
    suffix += 1;
  } while (used.has(candidate));
  return candidate;
}

function reusableNativeNumbering(
  paragraph: NativeWordParagraphNode,
  type: NativeWordListType,
): EditorNumberingFormat | null {
  const numbering = paragraph.format?.hasNumbering ? paragraph.format.numbering : undefined;
  if (!numbering || nativeListType(numbering) !== type) return null;
  return numbering.numId.startsWith(`coco-${type}-`) ? numbering : null;
}

function clampListLevel(ilvl: number): number {
  return Math.max(0, Math.min(8, Math.round(ilvl)));
}

function bulletForLevel(ilvl: number): string {
  return ['•', '◦', '▪'][ilvl % 3]!;
}

function numberingFormatForLevel(ilvl: number): string {
  return ['decimal', 'lowerLetter', 'lowerRoman'][ilvl % 3]!;
}

function normalizeBulletLabel(lvlText: string | undefined): string {
  const label = lvlText?.replace(/%\d+/g, '').trim();
  return label || '•';
}

function formatNumberingLabel(value: number, numFmt: string | undefined, lvlText: string | undefined): string {
  const formatted =
    numFmt === 'lowerLetter'
      ? alphaLabel(value, false)
      : numFmt === 'upperLetter'
        ? alphaLabel(value, true)
        : numFmt === 'lowerRoman'
          ? romanLabel(value).toLowerCase()
          : numFmt === 'upperRoman'
            ? romanLabel(value)
            : String(value);
  return lvlText ? lvlText.replace(/%\d+/g, formatted) : `${formatted}.`;
}

function alphaLabel(value: number, upper: boolean): string {
  let remaining = Math.max(1, value);
  let label = '';
  while (remaining > 0) {
    remaining -= 1;
    label = String.fromCharCode((upper ? 65 : 97) + (remaining % 26)) + label;
    remaining = Math.floor(remaining / 26);
  }
  return label;
}

function romanLabel(value: number): string {
  let remaining = Math.max(1, Math.min(3999, value));
  let label = '';
  const tokens: Array<[number, string]> = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  for (const [amount, token] of tokens) {
    while (remaining >= amount) {
      label += token;
      remaining -= amount;
    }
  }
  return label;
}
