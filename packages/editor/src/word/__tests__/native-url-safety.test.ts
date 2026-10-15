import { describe, expect, it } from 'vitest';
import { safeEmbeddedRasterImageUrl, safeRenderedHyperlinkHref } from '../native-url-safety.js';

describe('native document URL safety', () => {
  it('accepts embedded raster images and rejects active or remote image sources', () => {
    expect(safeEmbeddedRasterImageUrl('data:image/png;base64,AA==')).toBe(
      'data:image/png;base64,AA==',
    );
    expect(safeEmbeddedRasterImageUrl('data:image/svg+xml;base64,PHN2Zz4=')).toBeNull();
    expect(safeEmbeddedRasterImageUrl('javascript:alert(1)')).toBeNull();
    expect(safeEmbeddedRasterImageUrl('https://example.test/tracker.png')).toBeNull();
  });

  it('allows explicit web, mail and telephone links and neutralizes unsafe targets', () => {
    expect(safeRenderedHyperlinkHref('https://coco.example/docs', undefined)).toBe(
      'https://coco.example/docs',
    );
    expect(safeRenderedHyperlinkHref('mailto:team@example.test', undefined)).toBe(
      'mailto:team@example.test',
    );
    expect(safeRenderedHyperlinkHref('javascript:alert(1)', undefined)).toBe('#');
    expect(safeRenderedHyperlinkHref('//evil.example/path', undefined)).toBe('#');
    expect(safeRenderedHyperlinkHref(undefined, 'Section 1')).toBe('#Section%201');
  });
});
