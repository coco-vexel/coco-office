/**
 * WordPlan.apply 端到端测试 —— #190(replaceDrawingWithPlaceholder / deleteDrawing)。
 *
 * 走整条链:doc.plan() → builder → apply() → kernel.applyPlan → dispatchOp →
 * executeReplaceDrawingWithPlaceholder / executeDeleteDrawing → runTransaction →
 * (commit 后)sweepOrphanMediaParts → toZip。
 *
 * 重点验收:
 *   - A1 布局不变:wp:extent / wrap / position 字段在替换前后完全相等
 *   - A2 显示尺寸:wp:extent.cx / cy 不动
 *   - A3 原始二进制不残留:apply 后 toZip 输出包内 `word/media/<原图>` 不存在
 *         word/media/placeholder.png 在
 *   - deleteDrawing:w:drawing 节点消失,父空 w:r 也清除
 *   - 多 drawing 共享 media 时引用计数正确(只在最后一处替换 / 删除后才删 binary)
 *
 * inline 模式即可,无需 worker(逻辑全在主线程)。
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { A_NS, PIC_NS, R_NS, W14_NS, W_NS, WP_NS, WPS_NS } from '../namespaces.js';

const DOC_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="${IMG_REL}" Target="media/original.jpg"/>
</Relationships>`;

function makeDrawingXml(paraId: string, rId = 'rId7'): string {
  return `<w:p w14:paraId="${paraId}"><w:r><w:drawing>
    <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}" distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="3000000" cy="2000000"/>
      <wp:docPr id="1" name="P"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr><pic:cNvPr id="0" name="orig.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
}

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

const ORIGINAL_JPG_BYTES = new Uint8Array([
  // 极简 4 字节"假"JPEG 占位 + magic 头 —— 真假不重要,关键是字节身份
  0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46,
]);

async function makeFixtureWithPicture(bodyInner: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/_rels/document.xml.rels', DOC_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  zip.file('word/media/original.jpg', ORIGINAL_JPG_BYTES);
  return zip.generateAsync({ type: 'uint8array' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

describe('replaceDrawingWithPlaceholder —— 端到端(#190)', () => {
  test('单图替换:toZip 输出布局不变 + 占位 PNG 入包 + 原 jpg 不残留', async () => {
    const buf = await makeFixtureWithPicture(makeDrawingXml('AAAAAAA1'));
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-replace', mode: 'inline' });

    const before = await doc.drawings.list();
    expect(before).toHaveLength(1);
    const drawingId = before[0]!.drawingId;
    const beforeExtent = before[0]!.extentEMU;

    await doc.plan().replaceDrawingWithPlaceholder(drawingId).apply();

    // 验:A2 wp:extent 不动
    const after = await doc.drawings.list();
    expect(after).toHaveLength(1);
    expect(after[0]!.extentEMU).toEqual(beforeExtent);
    expect(after[0]!.anchoringMode).toBe('inline');

    // 序列化后验:输出包内的 media 替换情况
    const out = await doc.toZip();
    const zip = await JSZip.loadAsync(out);

    // placeholder.png 已入包
    const placeholderFile = zip.file('word/media/placeholder.png');
    expect(placeholderFile).not.toBeNull();

    // 原图二进制不残留(A3)
    const origFile = zip.file('word/media/original.jpg');
    expect(origFile).toBeNull();

    // document.xml.rels 内不再含 original.jpg
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsXml).not.toContain('original.jpg');
    expect(relsXml).toContain('media/placeholder.png');

    // [Content_Types] 含 png Default
    const ctXml = await zip.file('[Content_Types].xml')!.async('string');
    expect(ctXml).toMatch(/Extension="png"/);

    // document.xml 含 docPr/@descr = "[图片占位]"
    const docXmlOut = await zip.file('word/document.xml')!.async('string');
    expect(docXmlOut).toContain('descr="[图片占位]"');

    await doc.dispose();
  });

  test('两图共享同一 jpg:都替换后 jpg 才删,中间状态保留', async () => {
    const buf = await makeFixtureWithPicture(
      makeDrawingXml('AAAAAAA1') + makeDrawingXml('AAAAAAA2'),
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-share', mode: 'inline' });
    const ids = (await doc.drawings.list()).map((d) => d.drawingId);
    expect(ids).toHaveLength(2);

    // 先只替换第一张:中间产物里 jpg 应仍在
    await doc.plan().replaceDrawingWithPlaceholder(ids[0]!).apply();
    let zip = await JSZip.loadAsync(await doc.toZip());
    expect(zip.file('word/media/original.jpg')).not.toBeNull();

    // 再替换第二张:此时 jpg 应已删除
    await doc.plan().replaceDrawingWithPlaceholder(ids[1]!).apply();
    zip = await JSZip.loadAsync(await doc.toZip());
    expect(zip.file('word/media/original.jpg')).toBeNull();
    expect(zip.file('word/media/placeholder.png')).not.toBeNull();

    await doc.dispose();
  });
});

describe('deleteDrawing —— 端到端(#190)', () => {
  test('删 picture drawing:w:drawing 消失 + 原 jpg 不残留', async () => {
    const buf = await makeFixtureWithPicture(makeDrawingXml('AAAAAAA1'));
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-del', mode: 'inline' });
    const ids = (await doc.drawings.list()).map((d) => d.drawingId);

    await doc.plan().deleteDrawing(ids[0]!).apply();
    expect(await doc.drawings.list()).toHaveLength(0);

    const zip = await JSZip.loadAsync(await doc.toZip());
    const docXmlOut = await zip.file('word/document.xml')!.async('string');
    expect(docXmlOut).not.toContain('<w:drawing');
    expect(zip.file('word/media/original.jpg')).toBeNull();

    // 不写 placeholder.png(deleteDrawing 不应注入占位)
    expect(zip.file('word/media/placeholder.png')).toBeNull();

    await doc.dispose();
  });
});

describe('replaceShapeInnerText —— 端到端(#210)', () => {
  test('textBox apply 后:外壳保留,内部文字变占位,kind 仍为 textBox', async () => {
    const TEXTBOX_BODY = `<w:p w14:paraId="TXT00001"><w:r><w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}" distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="2000000" cy="1000000"/>
        <wp:docPr id="1" name="TextBox 1"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:cNvSpPr/>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="FFFFCC"/></a:solidFill>
                <a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  <w:p w14:paraId="TXT00002">
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:t>原始文本不应残留</w:t></w:r>
                  </w:p>
                  <w:p w14:paraId="TXT00003">
                    <w:r><w:t>第二行也要清掉</w:t></w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`;

    // textBox 不引用 media,DOC_RELS 用空 fixture 避开 jpg 残留干扰
    const zipIn = new JSZip();
    zipIn.file('[Content_Types].xml', CT_XML);
    zipIn.file('_rels/.rels', MAIN_RELS);
    zipIn.file(
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
    );
    zipIn.file('word/document.xml', docXml(TEXTBOX_BODY));
    const buf = await zipIn.generateAsync({ type: 'uint8array' });

    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-textbox', mode: 'inline' });
    const before = await doc.drawings.list();
    expect(before).toHaveLength(1);
    expect(before[0]!.kind).toBe('textBox');
    expect(before[0]!.hasInnerText).toBe(true);
    expect(before[0]!.innerTextAlign).toBe('center');
    const drawingId = before[0]!.drawingId;
    const beforeExtent = before[0]!.extentEMU;

    await doc.plan().replaceShapeInnerText(drawingId, '[Shape 内容占位]').apply();

    // 索引重建后:仍为 textBox,外壳尺寸不变,首段 jc 保留
    const after = await doc.drawings.list();
    expect(after).toHaveLength(1);
    expect(after[0]!.kind).toBe('textBox');
    expect(after[0]!.extentEMU).toEqual(beforeExtent);
    expect(after[0]!.innerTextAlign).toBe('center');
    expect(after[0]!.hasInnerText).toBe(true);

    // 序列化后:document.xml 中原文本不残留,占位文字写入
    const out = await doc.toZip();
    const zipOut = await JSZip.loadAsync(out);
    const docXmlOut = await zipOut.file('word/document.xml')!.async('string');
    expect(docXmlOut).not.toContain('原始文本不应残留');
    expect(docXmlOut).not.toContain('第二行也要清掉');
    expect(docXmlOut).toContain('[Shape 内容占位]');
    // 外壳关键属性序列化后仍在
    expect(docXmlOut).toContain('FFFFCC');
    expect(docXmlOut).toContain('w="6350"');
    expect(docXmlOut).toContain('cx="2000000"');

    await doc.dispose();
  });
});

describe('索引同步 —— apply 后再 list/get 正确', () => {
  test('apply replaceDrawingWithPlaceholder 后 drawingIdx 仍可解析同一 drawingId', async () => {
    const buf = await makeFixtureWithPicture(makeDrawingXml('AAAAAAA1'));
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-idx', mode: 'inline' });
    const ids = (await doc.drawings.list()).map((d) => d.drawingId);

    await doc.plan().replaceDrawingWithPlaceholder(ids[0]!).apply();

    // 索引重建后,同一 drawingId 仍能查到(物理位置不变,只是内部 blip 改了)
    const after = await doc.drawings.get(ids[0]!);
    expect(after).not.toBeNull();
    expect(after!.kind).toBe('picture');

    await doc.dispose();
  });

  test('apply deleteDrawing 后 drawings.list 为空', async () => {
    const buf = await makeFixtureWithPicture(makeDrawingXml('AAAAAAA1'));
    const doc = await WordDocument.fromZip(buf, { docId: 'e2e-del-idx', mode: 'inline' });
    const ids = (await doc.drawings.list()).map((d) => d.drawingId);

    await doc.plan().deleteDrawing(ids[0]!).apply();
    expect(await doc.drawings.list()).toHaveLength(0);

    await doc.dispose();
  });
});
