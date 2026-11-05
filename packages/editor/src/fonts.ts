export interface OfficeFontFallback {
  families: readonly string[];
  generic?: 'serif' | 'sans-serif';
}

// One registry for Word DOM, Spreadsheet Canvas/measurement, and PPT text/thumbnails.
const OFFICE_FONT_FALLBACKS: Array<OfficeFontFallback & { matches: RegExp }> = [
  {
    matches: /(?:楷体|kaiti|stkaiti)/i,
    families: ['Kaiti SC', 'STKaiti', 'KaiTi', 'KaiTi_GB2312', 'DFKai-SB'],
    generic: 'serif',
  },
  {
    matches: /(?:仿宋|fangsong|stfangsong)/i,
    families: ['STFangsong', 'FangSong', 'FangSong_GB2312', 'Noto Serif CJK SC'],
    generic: 'serif',
  },
  {
    matches: /(?:黑体|heiti|simhei|微软雅黑|microsoft yahei|pingfang)/i,
    families: ['STHeiti', 'Heiti SC', 'PingFang SC', 'Microsoft YaHei', 'Noto Sans CJK SC'],
    generic: 'sans-serif',
  },
  {
    matches: /(?:宋体|simsun|songti|stsong)/i,
    families: ['SimSun', 'Songti SC', 'STSong', 'Noto Serif CJK SC'],
    generic: 'serif',
  },
];

const DEFAULT_CJK_FALLBACK = OFFICE_FONT_FALLBACKS[3]!;

/** Keep the document's own family first; normalize platform aliases only in the view. */
export function resolveOfficeFontFamily(
  fontFamily: string,
  defaultFallback: OfficeFontFallback = DEFAULT_CJK_FALLBACK,
): string {
  const families = fontFamily.split(',').map((family) => family.trim()).filter(Boolean);
  const requested = families.join(' ');
  const fallback = OFFICE_FONT_FALLBACKS.find((item) => item.matches.test(requested)) ?? defaultFallback;
  const uniqueFamilies = [...new Set([...families, ...fallback.families])];
  return [
    ...uniqueFamilies.map((family) =>
      `"${family.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n\f]/g, ' ')}"`,
    ),
    ...(fallback.generic ? [fallback.generic] : []),
  ].join(', ');
}

/** Wait only for the fonts used by this view, not unrelated fonts elsewhere on the page. */
export async function loadOfficeFonts(fonts: Iterable<string>, sample = '汉字Aa0'): Promise<void> {
  if (typeof document === 'undefined' || !document.fonts?.load) return;
  await Promise.all([...new Set(fonts)].map((font) => document.fonts.load(font, sample).catch(() => [])));
}
