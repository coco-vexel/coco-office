// @vitest-environment happy-dom

import { describe, expect, test } from 'vitest';
import { renderNativeParagraph } from '../native-renderer.js';

describe('native Word fixed tab leaders', () => {
  test('renders an underscore leader row at the fixed right-tab width', () => {
    const paragraph = renderNativeParagraph({
      kind: 'paragraph',
      paraId: 'TAB00001',
      story: 'body',
      styleId: null,
      format: {
        tabStops: [{ alignment: 'right', position: 7200, leader: 'underscore' }],
      },
      content: [{ kind: 'text', text: '学生姓名：\u00a0张三\t', format: null }],
    });
    const row = paragraph.querySelector<HTMLElement>('.coco-word-tab-row');

    expect(row?.classList.contains('is-right-tab')).toBe(true);
    expect(row?.classList.contains('has-leader-underscore')).toBe(true);
    expect(row?.dataset.tabLeader).toBe('underscore');
    expect(row?.style.getPropertyValue('--coco-word-tab-stop')).toBe('360pt');
  });

  test('renders one continuous line for an underlined tab without a leader', () => {
    const paragraph = renderNativeParagraph({
      kind: 'paragraph',
      paraId: 'TAB00002',
      story: 'body',
      styleId: null,
      format: {
        tabStops: [{ alignment: 'right', position: 7200 }],
      },
      content: [
        { kind: 'text', text: '学生姓名：', format: null },
        { kind: 'text', text: '\u00a0张三', format: { underline: true } },
        { kind: 'text', text: '\t', format: { underline: true } },
      ],
    });
    const row = paragraph.querySelector<HTMLElement>('.coco-word-tab-row');
    const underlinedRun = paragraph.querySelector<HTMLElement>('.coco-word-underlined-tab-run');

    expect(row?.classList.contains('has-underlined-tab')).toBe(true);
    expect(row?.className).not.toContain('has-leader-');
    expect(row?.style.getPropertyValue('--coco-word-tab-stop')).toBe('360pt');
    expect(underlinedRun?.textContent).toBe('\u00a0张三');
    expect(row?.querySelector('.coco-word-tab-segment.is-segment-0')?.lastChild).toBe(underlinedRun);
  });
});
