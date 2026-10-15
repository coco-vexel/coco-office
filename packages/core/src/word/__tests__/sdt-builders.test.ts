import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { W_NS } from '../namespaces.js';
import { buildInlineSdt, hashSdtId, type SdtSpec } from '../sdt-builders.js';

function makeDoc(): Document {
  return parseXmlDocument(`<root xmlns:w="${W_NS}"/>`) as unknown as Document;
}

function serialize(node: Element): string {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return serializeToWellFormedString(node as any);
}

describe('buildInlineSdt', () => {
  test('plain text SDT 含 alias / tag / id / placeholder / showingPlcHdr / w:text', () => {
    const doc = makeDoc();
    const sdt = buildInlineSdt(doc, { kind: 'text', alias: '编制人' }, 42);
    const xml = serialize(sdt);

    expect(xml).toContain('<w:sdt');
    expect(xml).toContain('<w:sdtPr');
    expect(xml).toContain('<w:alias w:val="编制人"');
    expect(xml).toContain('<w:tag w:val="编制人"');
    expect(xml).toContain('<w:id w:val="42"');
    expect(xml).toContain('<w:placeholder');
    expect(xml).toContain('DefaultPlaceholder_-1854013440');
    expect(xml).toContain('<w:showingPlcHdr');
    expect(xml).toContain('<w:text');
    expect(xml).toContain('<w:sdtContent');
    expect(xml).toContain('请输入编制人');
  });

  test('date SDT 含 w:date / dateFormat / lid / calendar', () => {
    const doc = makeDoc();
    const sdt = buildInlineSdt(doc, { kind: 'date', alias: '签字日期' }, 100);
    const xml = serialize(sdt);

    expect(xml).toContain('<w:date');
    expect(xml).toContain('<w:dateFormat w:val="yyyy-M-d"');
    expect(xml).toContain('<w:lid w:val="zh-CN"');
    expect(xml).toContain('<w:calendar w:val="gregorian"');
    expect(xml).toContain('请输入签字日期');
  });

  test('date SDT 接受自定义 format / locale', () => {
    const doc = makeDoc();
    const sdt = buildInlineSdt(
      doc,
      { kind: 'date', alias: 'd', dateFormat: 'yyyy年M月d日', locale: 'zh-CN' },
      1,
    );
    const xml = serialize(sdt);
    expect(xml).toContain('<w:dateFormat w:val="yyyy年M月d日"');
  });

  test('rich text SDT 含 w:richText 类型标记', () => {
    const doc = makeDoc();
    const sdt = buildInlineSdt(doc, { kind: 'richText', alias: '说明内容' }, 7);
    const xml = serialize(sdt);
    expect(xml).toContain('<w:richText');
    expect(xml).not.toContain('<w:text/>');
  });

  test('自定义 tag 与 placeholderText 写入', () => {
    const doc = makeDoc();
    const spec: SdtSpec = {
      kind: 'text',
      alias: '编制人',
      tag: 'drafter_name',
      placeholderText: '<在此填写编制人>',
    };
    const xml = serialize(buildInlineSdt(doc, spec, 1));
    expect(xml).toContain('<w:tag w:val="drafter_name"');
    expect(xml).toContain('&lt;在此填写编制人&gt;');
  });

  test('rPr 透传到 sdtPr 与 sdtContent run', () => {
    const doc = makeDoc();
    // 模拟原 run 的 rPr (字体颜色)
    const rPr = doc.createElementNS(W_NS, 'w:rPr');
    const color = doc.createElementNS(W_NS, 'w:color');
    color.setAttributeNS(W_NS, 'w:val', 'FF0000');
    rPr.appendChild(color);

    const sdt = buildInlineSdt(doc, { kind: 'text', alias: '名称' }, 1, rPr);
    const xml = serialize(sdt);

    // rPr 应在 sdtPr 内出现一次, sdtContent 的 run 内出现一次
    const rPrCount = (xml.match(/<w:rPr/g) ?? []).length;
    expect(rPrCount).toBe(2);
    const colorCount = (xml.match(/<w:color w:val="FF0000"/g) ?? []).length;
    expect(colorCount).toBe(2);
  });

  test('alias 含空白时 tag 自动空格变下划线', () => {
    const doc = makeDoc();
    const sdt = buildInlineSdt(doc, { kind: 'text', alias: '编制 单位' }, 1);
    const xml = serialize(sdt);
    expect(xml).toContain('<w:tag w:val="编制_单位"');
  });
});

describe('hashSdtId', () => {
  test('幂等: 同一组输入产同一 id', () => {
    const id1 = hashSdtId('tbl-0001', 0, 1, '编制人');
    const id2 = hashSdtId('tbl-0001', 0, 1, '编制人');
    expect(id1).toBe(id2);
  });

  test('输入不同产不同 id (avalanche, 抽样)', () => {
    const ids = new Set<number>();
    ids.add(hashSdtId('tbl-0001', 0, 1, '编制人'));
    ids.add(hashSdtId('tbl-0001', 0, 1, '审核人'));
    ids.add(hashSdtId('tbl-0001', 1, 1, '编制人'));
    ids.add(hashSdtId('tbl-0001', 0, 2, '编制人'));
    ids.add(hashSdtId('tbl-0002', 0, 1, '编制人'));
    expect(ids.size).toBe(5);
  });

  test('id 永远非负 (32-bit signed 转正)', () => {
    for (const alias of ['a', 'A', '编', '测试', '12345']) {
      for (let row = 0; row < 5; row++) {
        for (let col = 0; col < 5; col++) {
          const id = hashSdtId('tbl-0001', row, col, alias);
          expect(id).toBeGreaterThanOrEqual(0);
          expect(Number.isInteger(id)).toBe(true);
        }
      }
    }
  });
});
