import { jsonToOoxml, ooxmlToJson, jsonToFlatOpc } from '@coco-vexel/coco-office-oojson';

export function runOojsonSample(text = 'Hello OOJSON') {
  const input = {
    type: 'document',
    content: [
      {
        type: 'paragraph',
        pPr: {},
        rPr: {},
        changes: [],
        content: [
          {
            type: 'run',
            rPr: {},
            content: [text],
            footnotes: [],
            endnotes: [],
            reviewType: 'common',
          },
        ],
      },
    ],
  };
  const xml = jsonToOoxml(input);
  const result = ooxmlToJson(xml);
  const actualText = result.doc.content
    .flatMap((p) => p.content ?? [])
    .flatMap((run) => run.content ?? [])
    .filter((value) => typeof value === 'string')
    .join('');
  if (actualText !== text) throw new Error('OOJSON text round-trip mismatch');
  return { input, xml, flatOpc: jsonToFlatOpc(input), output: result.doc, text: actualText };
}
