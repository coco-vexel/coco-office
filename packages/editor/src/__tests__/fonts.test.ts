import { describe, expect, test } from 'vitest';
import { resolveOfficeFontFamily } from '../fonts.js';

describe('shared Office font family resolution', () => {
  test.each(['KaiTi', '楷体', '楷体_GB2312', 'STKaiti', 'Kaiti SC'])('maps %s to platform equivalents', (family) => {
    const css = resolveOfficeFontFamily(family);
    expect(css.startsWith(`"${family}",`)).toBe(true);
    expect(css).toContain('"STKaiti"');
    expect(css).toContain('"Kaiti SC"');
    expect(css.endsWith(', serif')).toBe(true);
  });

  test('keeps the same CJK mapping regardless of the renderer default stack', () => {
    expect(resolveOfficeFontFamily('KaiTi', { families: [] })).toBe(
      resolveOfficeFontFamily('KaiTi', { families: ['Calibri'], generic: 'sans-serif' }),
    );
    expect(resolveOfficeFontFamily('Carlito', { families: [] })).toBe('"Carlito"');
  });
});
