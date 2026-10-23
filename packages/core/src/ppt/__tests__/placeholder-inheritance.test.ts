import { describe, expect, it } from 'vitest';
import type { PptShapeModel, PptSlideLayoutModel } from '../model.js';
import {
  findMatchingPptPlaceholder,
  resolvePptEffectivePlaceholderProperties,
} from '../placeholder-inheritance.js';

describe('PPT placeholder inheritance', () => {
  it('resolves master, layout, and slide properties without changing the source shape', () => {
    const masterBody: PptShapeModel = {
      id: 'master-body',
      kind: 'placeholder',
      placeholder: { type: 'body', index: '1' },
      geometry: { left: 100, top: 200, width: 800, height: 600 },
      style: { fillColor: '#EEEEEE', lineColor: '#111111' },
      text: {
        text: 'Master body',
        style: { fontSize: 32, color: '#222222' },
        body: { insetLeft: 40, insetTop: 20, autoFit: 'normal' },
      },
    };
    const layoutBody: PptShapeModel = {
      id: 'layout-body',
      kind: 'placeholder',
      placeholder: { index: '1' },
      geometry: { top: 250 },
      style: { fillColor: '#FFFFFF' },
      text: { text: 'Layout body', body: { insetTop: 30 } },
    };
    const slideBody: PptShapeModel = {
      id: 'slide-body',
      kind: 'placeholder',
      placeholder: { index: '1' },
      geometry: { left: 120 },
      text: { text: 'Slide body', style: { bold: true }, body: { autoFit: 'shape' } },
    };
    const layout: PptSlideLayoutModel = {
      relationshipId: 'rId1',
      path: 'ppt/slideLayouts/slideLayout1.xml',
      shapes: [layoutBody],
      master: {
        relationshipId: 'rId1',
        path: 'ppt/slideMasters/slideMaster1.xml',
        shapes: [masterBody],
        textStyles: { body: { fontFamily: 'Aptos', fontSize: 30 } },
      },
    };

    expect(resolvePptEffectivePlaceholderProperties(slideBody, layout)).toEqual({
      placeholder: { type: 'body', index: '1' },
      geometry: { left: 120, top: 250, width: 800, height: 600 },
      style: { fillColor: '#FFFFFF', lineColor: '#111111' },
      textStyle: {
        fontFamily: 'Aptos',
        fontSize: 32,
        color: '#222222',
        bold: true,
      },
      textBody: { insetLeft: 40, insetTop: 30, autoFit: 'shape' },
    });
    expect(slideBody).toEqual({
      id: 'slide-body',
      kind: 'placeholder',
      placeholder: { index: '1' },
      geometry: { left: 120 },
      text: { text: 'Slide body', style: { bold: true }, body: { autoFit: 'shape' } },
    });
  });

  it('falls back to placeholder type when indices differ between layout and master', () => {
    const shape: PptShapeModel = {
      id: 'layout-date',
      kind: 'placeholder',
      placeholder: { type: 'dt', index: '10' },
      geometry: {},
    };
    const candidate: PptShapeModel = {
      id: 'master-date',
      kind: 'placeholder',
      placeholder: { type: 'dt', index: '2' },
      geometry: { left: 10, top: 20, width: 30, height: 40 },
    };

    expect(findMatchingPptPlaceholder([candidate], shape)).toBe(candidate);
  });
});
