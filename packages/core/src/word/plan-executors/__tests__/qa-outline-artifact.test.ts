import { describe, test, expect } from 'vitest';
import JSZip from 'jszip';
import { Kernel } from '../../index.js';
import type { InsertParagraphSpec } from '../../plan.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

function makeDocXml() {
  const ns = W_NS,
    ns14 = W14_NS;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' +
    ns +
    '" xmlns:w14="' +
    ns14 +
    '">' +
    '<w:body>' +
    '<w:p w14:paraId="ANCH0001"><w:r><w:t>anchor</w:t></w:r></w:p>' +
    '<w:p w14:paraId="SRCE0001">' +
    '<w:pPr>' +
    '<w:pStyle w:val="Heading1"/>' +
    '<w:spacing w:before="240" w:after="0"/>' +
    '<w:rPr><w:b/></w:rPr>' +
    '</w:pPr>' +
    '<w:r><w:t>first-chapter</w:t></w:r>' +
    '</w:p>' +
    '</w:body></w:document>'
  );
}

async function buildDocx(): Promise<Uint8Array> {
  const ct =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const wrels =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', rels);
  zip.file('word/document.xml', makeDocXml());
  zip.file('word/_rels/document.xml.rels', wrels);
  return zip.generateAsync({ type: 'uint8array' });
}

async function getDocXml(bytes: Uint8Array): Promise<string> {
  const z = await JSZip.loadAsync(bytes);
  const f = z.file('word/document.xml');
  if (!f) throw new Error('missing word/document.xml');
  return f.async('string');
}

function getParaXml(xml: string, id: string): string {
  const tag = 'paraId="' + id + '"';
  const idx = xml.indexOf(tag);
  if (idx < 0) throw new Error('paraId not found: ' + id);
  let s = idx;
  while (s > 0 && !(xml[s] === '<' && xml.slice(s, s + 4) === '<w:p')) s--;
  return xml.slice(s, xml.indexOf('</w:p>', s) + 6);
}

const OPTS = { trackChanges: false };

describe('QA artifact: outlineLevel real docx verification', () => {
  test('[QA-1] merger-scenario: styleSource pPr has spacing+rPr(no numPr) -> outlineLvl before rPr', async () => {
    const docx = await buildDocx();
    const k = new Kernel();
    await k.load({ source: { type: 'zip', bytes: docx }, docId: 'qa-1' });
    const r = await k.applyPlan({
      ops: [
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'ANCH0001',
          spec: {
            text: 'ch2',
            styleSource: 'SRCE0001',
            outlineLevel: 1,
          } as unknown as InsertParagraphSpec,
        },
      ],
      options: OPTS,
    });
    const nid = r.insertedParagraphs[0]!.newParaId;
    expect(nid).toBeTruthy();
    const zipBytes = await k.toZip();
    const docXml = await getDocXml(zipBytes);
    const px = getParaXml(docXml, nid);
    const oi = px.indexOf('outlineLvl');
    const ri = px.indexOf('<w:rPr');
    expect(oi, 'outlineLvl must exist in pPr').toBeGreaterThanOrEqual(0);
    expect(ri, 'rPr must exist (from styleSource)').toBeGreaterThanOrEqual(0);
    expect(oi, 'outlineLvl must come BEFORE rPr -- schema compliance').toBeLessThan(ri);
    const vm = px.match(/outlineLvl[^>]*val="(\d+)"/);
    expect(vm?.[1]).toBe('1');
  });

  test('[QA-2] round-trip: toZip then fromZip does not throw', async () => {
    const docx = await buildDocx();
    const k = new Kernel();
    await k.load({ source: { type: 'zip', bytes: docx }, docId: 'qa-2a' });
    await k.applyPlan({
      ops: [
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'ANCH0001',
          spec: {
            text: 'ch2',
            styleSource: 'SRCE0001',
            outlineLevel: 1,
          } as unknown as InsertParagraphSpec,
        },
      ],
      options: OPTS,
    });
    const zipBytes = await k.toZip();
    const k2 = new Kernel();
    await expect(
      k2.load({ source: { type: 'zip', bytes: zipBytes }, docId: 'qa-2b' }),
    ).resolves.not.toThrow();
  });

  test('[QA-3] omit outlineLevel -> no outlineLvl injected', async () => {
    const docx = await buildDocx();
    const k = new Kernel();
    await k.load({ source: { type: 'zip', bytes: docx }, docId: 'qa-3' });
    const r = await k.applyPlan({
      ops: [
        {
          type: 'insertParagraphAfter',
          anchorParaId: 'ANCH0001',
          spec: { text: 'x', styleSource: 'SRCE0001' } as unknown as InsertParagraphSpec,
        },
      ],
      options: OPTS,
    });
    const zip = await k.toZip();
    const xml = await getDocXml(zip);
    const px = getParaXml(xml, r.insertedParagraphs[0]!.newParaId);
    expect(px).not.toContain('outlineLvl');
  });

  test('[QA-4] outlineLevel=-1 -> INVALID_ARGUMENT', async () => {
    const docx = await buildDocx();
    const k = new Kernel();
    await k.load({ source: { type: 'zip', bytes: docx }, docId: 'qa-4' });
    await expect(
      k.applyPlan({
        ops: [
          {
            type: 'insertParagraphAfter',
            anchorParaId: 'ANCH0001',
            spec: { text: 'x', outlineLevel: -1 } as unknown as InsertParagraphSpec,
          },
        ],
        options: OPTS,
      }),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  });

  test('[QA-5] outlineLevel=1.5 -> INVALID_ARGUMENT', async () => {
    const docx = await buildDocx();
    const k = new Kernel();
    await k.load({ source: { type: 'zip', bytes: docx }, docId: 'qa-5' });
    await expect(
      k.applyPlan({
        ops: [
          {
            type: 'insertParagraphAfter',
            anchorParaId: 'ANCH0001',
            spec: { text: 'x', outlineLevel: 1.5 } as unknown as InsertParagraphSpec,
          },
        ],
        options: OPTS,
      }),
    ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  });
});
