import { describe, expect, it } from 'vitest';
import {
  nativeBlockClipboardText,
  parseNativeBlockClipboard,
  serializeNativeBlockClipboard,
  serializeNativeClipboardHtml,
} from '../native-clipboard.js';

describe('native clipboard HTML', () => {
  it('round trips native block payloads and rejects malformed payloads', () => {
    const blocks = [{
      kind: 'paragraph' as const,
      paraId: 'P1',
      story: 'body',
      styleId: null,
      format: null,
      content: [{
        kind: 'text' as const,
        text: 'Rich text',
        format: { bold: true },
        hyperlink: { target: 'https://coco.example/docs' },
      }],
    }];

    expect(parseNativeBlockClipboard(serializeNativeBlockClipboard(blocks))).toEqual(blocks);
    expect(nativeBlockClipboardText(blocks)).toBe('Rich text');
    expect(parseNativeBlockClipboard('{"version":1,"blocks":[{"kind":"script"}]}')).toBeNull();
  });

  it('serializes run formatting and consecutive lists', () => {
    const html = serializeNativeClipboardHtml([
      {
        text: 'Bold red',
        runs: [{ from: 0, to: 4, format: { bold: true, color: 'CC0000' } }],
        listType: 'bullet',
      },
      {
        text: 'Highlighted',
        runs: [{ from: 0, to: 11, format: { highlight: 'FFFF00', fontSize: 14 } }],
        listType: 'bullet',
      },
      { text: 'Plain', runs: [] },
    ]);

    expect(html).toBe(
      '<ul><li><span style="font-weight:700;color:#CC0000">Bold</span> red</li>' +
        '<li><span style="font-size:14pt;background-color:#FFFF00">Highlighted</span></li></ul>' +
        '<p>Plain</p>',
    );
  });

  it('escapes clipboard text and inline styles', () => {
    const html = serializeNativeClipboardHtml([
      {
        text: '<A>\nB',
        runs: [{ from: 0, to: 3, format: { fontFamily: 'A&B"' } }],
      },
    ]);

    expect(html).toBe('<p><span style="font-family:A&amp;B&quot;">&lt;A&gt;</span><br>B</p>');
  });

  it('serializes external links and document anchors with their run formatting', () => {
    const html = serializeNativeClipboardHtml([{
      text: 'Docs section',
      runs: [
        {
          from: 0,
          to: 4,
          format: { bold: true },
          hyperlink: { target: 'https://coco.example/docs', tooltip: 'Coco docs' },
        },
        { from: 5, to: 12, format: {}, hyperlink: { anchor: 'Section_1' } },
      ],
    }]);

    expect(html).toBe(
      '<p><a href="https://coco.example/docs" title="Coco docs"><span style="font-weight:700">Docs</span></a> ' +
        '<a href="#Section_1">section</a></p>',
    );
  });
});
