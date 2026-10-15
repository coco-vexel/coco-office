/**
 * pack 往返不变性测试套件(#11)。
 *
 * 验证四条核心路径:
 *   1. fromZip → toZip → fromZip          语义等价
 *   2. fromFlatOpc → toFlatOpc → fromFlatOpc 语义等价
 *   3. fromZip → toFlatOpc → fromFlatOpc → toZip 交叉等价
 *   4. xml:space="preserve" / <mc:AlternateContent> / rels Target 守恒
 *
 * 本文件是 #9 / #10 单测的超集:后两者验证单向正确,本文件验证多跳不变。
 */
import JSZip from 'jszip';
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  fromFlatOpc,
  fromZip,
  getPart,
  listParts,
  toFlatOpc,
  toZip,
  type Package,
} from '../index.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

/** 覆盖关键保留点:xml:space / mc:AlternateContent / 外层嵌套 */
const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:mc="${MC_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">  leading spaces  </w:t></w:r></w:p>
    <w:p>
      <mc:AlternateContent>
        <mc:Choice Requires="w14">
          <w:r><w:t>choice-body</w:t></w:r>
        </mc:Choice>
        <mc:Fallback>
          <w:r><w:t>fallback-body</w:t></w:r>
        </mc:Fallback>
      </mc:AlternateContent>
    </w:p>
  </w:body>
</w:document>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"/>`;

/** rels 的 Target 以 "word/document.xml" 这种 non-URL-encoded 字面路径存在 */
const MAIN_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const PNG_BYTES = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0xde, 0xad, 0xbe, 0xef,
]);

async function makeFixtureZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS_XML);
  zip.file('word/document.xml', DOC_XML);
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/media/image1.png', PNG_BYTES);
  return zip.generateAsync({ type: 'uint8array' });
}

/** 提取 Package 的语义指纹:path + contentType + 二进制 bytes 或 XML 文本 */
async function fingerprint(pkg: Package): Promise<Record<string, string>> {
  const fp: Record<string, string> = {};
  for (const part of listParts(pkg)) {
    fp[`${part.path}:ct`] = part.contentType;
    if (part.xmlDom) {
      // 用结构化字段刻画,避免 serializer 的 prefix/spacing 差异
      fp[`${part.path}:w:t`] = textsOfLocal(part.xmlDom, W_NS, 't').join('|');
      fp[`${part.path}:xml:space`] = textsOfLocal(part.xmlDom, W_NS, 't')
        .map(
          (_, i) =>
            getAttrNS(
              part.xmlDom!,
              W_NS,
              't',
              i,
              'http://www.w3.org/XML/1998/namespace',
              'space',
            ) ?? '',
        )
        .join('|');
      fp[`${part.path}:mc:Choice`] = textsOfLocal(part.xmlDom, MC_NS, 'Choice').join('|');
      fp[`${part.path}:mc:Fallback`] = textsOfLocal(part.xmlDom, MC_NS, 'Fallback').join('|');
      fp[`${part.path}:rel-targets`] = targetsOfRels(part.xmlDom).join('|');
    } else if (part.bytes) {
      fp[`${part.path}:bytes-hex`] = Array.from(part.bytes)
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    }
  }
  return fp;
}

function textsOfLocal(dom: Document, ns: string, local: string): string[] {
  const els = dom.getElementsByTagNameNS(ns, local);
  const out: string[] = [];
  for (let i = 0; i < els.length; i++) {
    out.push(els[i]?.textContent ?? '');
  }
  return out;
}

function getAttrNS(
  dom: Document,
  elNs: string,
  elLocal: string,
  idx: number,
  attrNs: string,
  attrLocal: string,
): string | null {
  const el = dom.getElementsByTagNameNS(elNs, elLocal)[idx];
  return el?.getAttributeNS(attrNs, attrLocal) ?? null;
}

function targetsOfRels(dom: Document): string[] {
  const ns = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const rels = dom.getElementsByTagNameNS(ns, 'Relationship');
  const out: string[] = [];
  for (let i = 0; i < rels.length; i++) {
    const t = rels[i]?.getAttribute('Target');
    if (t !== null && t !== undefined) out.push(t);
  }
  return out;
}

describe('path 1:fromZip → toZip → fromZip', () => {
  test('paths / contentType 守恒', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const zip2 = await toZip(pkg1);
    const pkg2 = await fromZip(zip2);

    const paths1 = listParts(pkg1)
      .map((p) => p.path)
      .sort();
    const paths2 = listParts(pkg2)
      .map((p) => p.path)
      .sort();
    expect(paths2).toEqual(paths1);
    for (const p of paths1) {
      expect(getPart(pkg2, p)?.contentType).toBe(getPart(pkg1, p)?.contentType);
    }
  });

  test('语义指纹守恒', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const zip2 = await toZip(pkg1);
    const pkg2 = await fromZip(zip2);
    expect(await fingerprint(pkg2)).toEqual(await fingerprint(pkg1));
  });

  test('二进制字节严格守恒', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const zip2 = await toZip(pkg1);
    const pkg2 = await fromZip(zip2);
    const img1 = getPart(pkg1, 'word/media/image1.png');
    const img2 = getPart(pkg2, 'word/media/image1.png');
    expect(Array.from(img2!.bytes!)).toEqual(Array.from(img1!.bytes!));
  });
});

describe('path 2:fromFlatOpc → toFlatOpc → fromFlatOpc', () => {
  test('paths(去 Content_Types)/ contentType / 语义指纹守恒', async () => {
    // 以 zip 起手构造 Flat OPC 源
    const pkg0 = await fromZip(await makeFixtureZip());
    const flat1 = toFlatOpc(pkg0);
    const pkg1 = fromFlatOpc(flat1);
    const flat2 = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat2);

    const paths1 = listParts(pkg1)
      .map((p) => p.path)
      .sort();
    const paths2 = listParts(pkg2)
      .map((p) => p.path)
      .sort();
    expect(paths2).toEqual(paths1);
    // Flat OPC 不含 [Content_Types].xml
    expect(paths1).not.toContain('[Content_Types].xml');

    for (const p of paths1) {
      expect(getPart(pkg2, p)?.contentType).toBe(getPart(pkg1, p)?.contentType);
    }

    expect(await fingerprint(pkg2)).toEqual(await fingerprint(pkg1));
  });

  test('mso-application PI 贯穿两次往返', async () => {
    const pkg0 = await fromZip(await makeFixtureZip());
    const flat1 = toFlatOpc(pkg0);
    expect(flat1).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
    const pkg1 = fromFlatOpc(flat1);
    const flat2 = toFlatOpc(pkg1);
    expect(flat2).toMatch(/<\?mso-application progid="Word\.Document"\?>/);
  });
});

describe('path 3:zip ↔ Flat OPC 交叉', () => {
  test('fromZip → toFlatOpc → fromFlatOpc → toZip:语义指纹守恒', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);

    const flat = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat);

    const zip2 = await toZip(pkg2);
    const pkg3 = await fromZip(zip2);

    // 交叉后:Flat OPC 会丢掉 [Content_Types].xml(从 pkg1 有 → pkg2 无);
    // 最终 pkg3 经 toZip 回注后,ct 根据 part 自身 contentType 回填时并不会自动
    // 补出 CT part(本期 toZip 只回写已有 parts)。所以比较 pkg1 去掉 CT 后的指纹。
    const fp1 = await fingerprint(pkg1);
    const fp3 = await fingerprint(pkg3);
    const fp1NoCT = omitKeysStartingWith(fp1, '[Content_Types].xml');
    const fp3NoCT = omitKeysStartingWith(fp3, '[Content_Types].xml');
    expect(fp3NoCT).toEqual(fp1NoCT);
  });
});

describe('path 4:关键元素保留', () => {
  test('xml:space="preserve" 三跳后仍在', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const flat = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat);
    const zip2 = await toZip(pkg2);
    const pkg3 = await fromZip(zip2);

    const doc = getPart(pkg3, 'word/document.xml')!.xmlDom!;
    const ts = doc.getElementsByTagNameNS(W_NS, 't');
    const firstT = ts[0];
    expect(firstT?.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space')).toBe(
      'preserve',
    );
    expect(firstT?.textContent).toBe('  leading spaces  ');
  });

  test('<mc:AlternateContent> 的 Choice + Fallback 都保留', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const flat = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat);
    const zip2 = await toZip(pkg2);
    const pkg3 = await fromZip(zip2);

    const doc = getPart(pkg3, 'word/document.xml')!.xmlDom!;
    const choice = doc.getElementsByTagNameNS(MC_NS, 'Choice');
    const fallback = doc.getElementsByTagNameNS(MC_NS, 'Fallback');
    expect(choice.length).toBe(1);
    expect(fallback.length).toBe(1);
    expect(choice[0]?.getAttribute('Requires')).toBe('w14');
    expect(choice[0]?.textContent).toContain('choice-body');
    expect(fallback[0]?.textContent).toContain('fallback-body');
  });

  test('rels Target 路径原样保留(不做 path 规范化)', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const flat = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat);
    const zip2 = await toZip(pkg2);
    const pkg3 = await fromZip(zip2);

    const rels = getPart(pkg3, '_rels/.rels')!.xmlDom!;
    const targets = targetsOfRels(rels);
    expect(targets).toContain('word/document.xml');
    expect(targets).toContain('docProps/app.xml');
  });

  test('PNG 二进制字节三跳后仍一致', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const flat = toFlatOpc(pkg1);
    const pkg2 = fromFlatOpc(flat);
    const zip2 = await toZip(pkg2);
    const pkg3 = await fromZip(zip2);

    const img3 = getPart(pkg3, 'word/media/image1.png');
    expect(Array.from(img3!.bytes!)).toEqual(Array.from(PNG_BYTES));
  });
});

describe('dirty 复用路径在 roundtrip 里零 serialize', () => {
  test('未改任何 part 时:fromZip → toZip 的字节与"用 _originalBytes 重打包"一致', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const zip2 = await toZip(pkg1);

    // 读回 zip2,检查 document.xml 的字节 === 源 zip1 里的字节
    const z1 = await JSZip.loadAsync(zip1);
    const z2 = await JSZip.loadAsync(zip2);
    const doc1 = await z1.file('word/document.xml')!.async('uint8array');
    const doc2 = await z2.file('word/document.xml')!.async('uint8array');
    expect(Array.from(doc2)).toEqual(Array.from(doc1));
  });
});

function omitKeysStartingWith(obj: Record<string, string>, prefix: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(obj)) {
    if (!k.startsWith(prefix)) out[k] = v;
  }
  return out;
}

describe('sanity:XML 经 parse → serialize → parse 始终合法', () => {
  test('Flat OPC 输出可再 parse', async () => {
    const zip1 = await makeFixtureZip();
    const pkg1 = await fromZip(zip1);
    const flat = toFlatOpc(pkg1);
    expect(() => parseXmlDocument(flat)).not.toThrow();
  });
});
