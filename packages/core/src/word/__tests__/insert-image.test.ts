/**
 * insertImageAfter + getImageBytes 端到端测试（task-insert-image）。
 *
 * §4 测试矩阵 16 条：
 *   insertImageAfter:
 *     1  基本插入：anchor 后新增段，段内 <w:drawing> 含 wp:inline + blip@r:embed
 *     2  media part 新建：hasPart('word/media/image1.png')；bytes 一致
 *     3  rel + rId：rels 有 image Relationship，rId 不撞
 *     4  content-type：CT 含 Default Extension="png"
 *     5  EMU 尺寸：wp:extent@cx/cy + a:ext@cx/cy == widthEmu/heightEmu
 *     6  新段 paraId：有 w14:paraId 且全局唯一；重载后 paragraphs.list 含该段
 *     7  多次插入：image1/image2 不撞，rId 不撞，docPr id 不撞
 *     8  anchor 不存在 → PARA_NOT_FOUND
 *     9  不支持 contentType → UNSUPPORTED_OOXML_FEATURE
 *     10 jpeg：image1.jpeg + Default Extension=jpeg
 *     11 round-trip docx：apply→toZip→fromZip，drawing 在、media 在、getImageBytes bytes 一致
 *     12 ★跨入口 round-trip flatOpc（R3）：apply→toFlatOpc→fromFlatOpc→getImageBytes bytes 一致
 *     13 事务回滚（R1）：多-op plan 中途失败 → 新建 media part 被删（不残留）
 *   getImageBytes:
 *     14 读 source 图：getImageBytes 返回 bytes（长度>0）+ contentType + extentEMU
 *     15 非 picture / 不存在：返回 null
 *     16 inline 模式可用：mode='inline' 下走通
 */
import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS, WP_NS, A_NS, R_NS } from '../namespaces.js';

// ─────────── 最小 1×1 透明 PNG（8 字节签名 + IHDR + IDAT + IEND）───────────
// 标准 1×1 RGBA 透明 PNG，已编码为字节数组
const TINY_PNG = new Uint8Array([
  0x89,
  0x50,
  0x4e,
  0x47,
  0x0d,
  0x0a,
  0x1a,
  0x0a, // PNG signature
  0x00,
  0x00,
  0x00,
  0x0d,
  0x49,
  0x48,
  0x44,
  0x52, // IHDR chunk len + type
  0x00,
  0x00,
  0x00,
  0x01,
  0x00,
  0x00,
  0x00,
  0x01, // 1x1
  0x08,
  0x06,
  0x00,
  0x00,
  0x00,
  0x1f,
  0x15,
  0xc4, // 8bit RGBA + CRC
  0x89,
  0x00,
  0x00,
  0x00,
  0x0b,
  0x49,
  0x44,
  0x41, // IDAT chunk
  0x54,
  0x78,
  0x9c,
  0x62,
  0x00,
  0x01,
  0x00,
  0x00,
  0x05,
  0x00,
  0x01,
  0x0d,
  0x0a,
  0x2d,
  0xb4,
  0x00, // compressed data + CRC
  0x00,
  0x00,
  0x00,
  0x49,
  0x45,
  0x4e,
  0x44,
  0xae, // IEND
  0x42,
  0x60,
  0x82,
]);

const TINY_JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46]);

// ─────────── fixture 构造 ───────────

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

function docXml(bodyInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${bodyInner}</w:body>
</w:document>`;
}

async function makeFixtureZip(bodyInner: string, withDocRels = true): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(bodyInner));
  if (withDocRels) {
    zip.file('word/_rels/document.xml.rels', DOC_RELS);
  }
  return zip.generateAsync({ type: 'uint8array' });
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── insertImageAfter ───────────

describe('plan.insertImageAfter', () => {
  test('1 基本插入：anchor 后新增段，段内 <w:drawing> 含 wp:inline + blip@r:embed', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0001"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-1', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0001', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 914400,
        heightEmu: 685800,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const xml = await z.file('word/document.xml')!.async('string');
    // 新段在 anchor 后
    expect(xml).toContain('AABB0001');
    // drawing 结构
    expect(xml).toContain('w:drawing');
    expect(xml).toContain('wp:inline');
    expect(xml).toContain('a:blip');
    expect(xml).toMatch(/r:embed="rId\d+"/);
  });

  test('2 media part 新建：hasPart("word/media/image1.png")；bytes 一致', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0002"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-2', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0002', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const mediaFile = z.file('word/media/image1.png');
    expect(mediaFile).not.toBeNull();
    const mediaBytes = new Uint8Array(await mediaFile!.async('arraybuffer'));
    expect(mediaBytes).toEqual(TINY_PNG);
  });

  test('3 rel + rId：rels 有 image Relationship，rId 不与现有 rId1 冲突', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0003"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-3', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0003', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const relsXml = await z.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsXml).toContain('relationships/image');
    expect(relsXml).toContain('media/image1.png');
    // rId2 = max(existing rId1) + 1
    expect(relsXml).toContain('rId2');
  });

  test('4 content-type：CT 含 Default Extension="png" ContentType="image/png"', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0004"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-4', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0004', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const ctXml = await z.file('[Content_Types].xml')!.async('string');
    expect(ctXml).toContain('Extension="png"');
    expect(ctXml).toContain('image/png');
  });

  test('5 EMU 尺寸：wp:extent@cx/cy == widthEmu/heightEmu；a:ext 同', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0005"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-5', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0005', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 1234567,
        heightEmu: 7654321,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const xml = await z.file('word/document.xml')!.async('string');
    expect(xml).toContain('cx="1234567"');
    expect(xml).toContain('cy="7654321"');
  });

  test('6 新段 paraId：有 w14:paraId 且全局唯一；重载后 paragraphs.list 含该段', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0006"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-6', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0006', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    // 重载后检查段落数（2 = anchor + 新图段）
    const doc2 = await WordDocument.fromZip(out, { docId: 'img-6-r', mode: 'inline' });
    const list = await doc2.paragraphs.list({ pagination: { mode: 'all' } });
    expect(list.paragraphs.length).toBeGreaterThanOrEqual(2);

    // paraId 唯一
    const ids = list.paragraphs.map((p) => p.paraId);
    expect(new Set(ids).size).toBe(ids.length);
    await doc2.dispose();
  });

  test('7 多次插入：image1/image2 不撞，rId 不撞，docPr id 不撞', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0007"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-7', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0007', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const paraList = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    // 第 2 段（image段）的 paraId
    const imgParaId = paraList.paragraphs[1]!.paraId;

    await doc
      .plan()
      .insertImageAfter(imgParaId, {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 200000,
        heightEmu: 200000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    // 两个 media 文件
    expect(z.file('word/media/image1.png')).not.toBeNull();
    expect(z.file('word/media/image2.png')).not.toBeNull();

    const docXmlStr = await z.file('word/document.xml')!.async('string');
    const relsXml = await z.file('word/_rels/document.xml.rels')!.async('string');
    // 两个不同 rId
    expect(relsXml).toContain('rId2');
    expect(relsXml).toContain('rId3');
    // docPr id 不撞：有两个不同的 wp:docPr id
    const docPrMatches = [...docXmlStr.matchAll(/docPr[^>]*id="(\d+)"/g)];
    const docPrIds = docPrMatches.map((m) => m[1]);
    expect(new Set(docPrIds).size).toBe(docPrIds.length);
  });

  test('8 anchor 不存在 → PARA_NOT_FOUND', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0008"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-8', mode: 'inline' });
    await expect(
      doc
        .plan()
        .insertImageAfter('FFFFFFFF', {
          bytes: TINY_PNG,
          contentType: 'image/png',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .apply(),
    ).rejects.toThrow(WordError);
    await doc.dispose();
  });

  test('9 不支持的 contentType → UNSUPPORTED_OOXML_FEATURE', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0009"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-9', mode: 'inline' });
    await expect(
      doc
        .plan()
        .insertImageAfter('AABB0009', {
          bytes: TINY_PNG,
          contentType: 'image/webp',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .apply(),
    ).rejects.toThrow(/UNSUPPORTED_OOXML_FEATURE|unsupported contentType/i);
    await doc.dispose();
  });

  test('10 jpeg：image1.jpeg + Default Extension=jpeg', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0010"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-10', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0010', {
        bytes: TINY_JPEG,
        contentType: 'image/jpeg',
        widthEmu: 100000,
        heightEmu: 100000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    expect(z.file('word/media/image1.jpeg')).not.toBeNull();
    const ctXml = await z.file('[Content_Types].xml')!.async('string');
    expect(ctXml).toContain('Extension="jpeg"');
    expect(ctXml).toContain('image/jpeg');
  });

  test('11 round-trip docx：apply→toZip→fromZip，drawing 在、media 在、getImageBytes bytes 一致', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0011"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-11', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0011', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 914400,
        heightEmu: 685800,
      })
      .apply();

    const out = await doc.toZip();

    // 重载
    const doc2 = await WordDocument.fromZip(out, { docId: 'img-11-r', mode: 'inline' });
    const drawings = await doc2.drawings.list();
    expect(drawings.length).toBeGreaterThanOrEqual(1);

    const drawing = drawings[drawings.length - 1]!;
    const result = await drawing.getImageBytes();
    expect(result).not.toBeNull();
    expect(result!.bytes).toEqual(TINY_PNG);
    expect(result!.contentType).toBe('image/png');
    expect(result!.extentEMU.cx).toBe(914400);
    expect(result!.extentEMU.cy).toBe(685800);

    await doc.dispose();
    await doc2.dispose();
  });

  test('12 ★跨入口 round-trip flatOpc（R3）：apply→toFlatOpc→fromFlatOpc→getImageBytes bytes 一致', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0012"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-12', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('AABB0012', {
        bytes: TINY_PNG,
        contentType: 'image/png',
        widthEmu: 500000,
        heightEmu: 400000,
      })
      .apply();

    // toFlatOpc（wps 入口）
    const flatOpc = await doc.toFlatOpc();
    await doc.dispose();

    // fromFlatOpc 重载
    const doc2 = await WordDocument.fromFlatOpc(flatOpc, { docId: 'img-12-r', mode: 'inline' });
    const drawings = await doc2.drawings.list();
    expect(drawings.length).toBeGreaterThanOrEqual(1);

    const drawing = drawings[drawings.length - 1]!;
    const result = await drawing.getImageBytes();
    expect(result).not.toBeNull();
    // 字节完全一致（base64 往返不损）
    expect(result!.bytes).toEqual(TINY_PNG);
    expect(result!.contentType).toBe('image/png');
    expect(result!.extentEMU.cx).toBe(500000);
    expect(result!.extentEMU.cy).toBe(400000);

    await doc2.dispose();
  });

  test('13 事务回滚（R1）：plan 第 2 个 op 失败 → image1.png 被删，不残留', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB0013"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-13', mode: 'inline' });

    // plan: insertImageAfter（成功）+ 一个会 throw 的 op（deleteParagraph 引用不存在段）
    await expect(
      doc
        .plan()
        .insertImageAfter('AABB0013', {
          bytes: TINY_PNG,
          contentType: 'image/png',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .deleteParagraph('FFFFFFFF') // 不存在 → PARA_NOT_FOUND → 整个 plan 回滚
        .apply(),
    ).rejects.toThrow(WordError);

    // 回滚后 media part 不残留
    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    expect(z.file('word/media/image1.png')).toBeNull();
  });

  test('13b R1 多张图回滚：2 张 png + 失败 op → image1.png + image2.png 都被删（多图盲区）', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB013B"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-13b', mode: 'inline' });

    // plan: 2 × insertImageAfter（均会成功）+ 1 × deleteParagraph 不存在 → 回滚全部
    await expect(
      doc
        .plan()
        .insertImageAfter('AABB013B', {
          bytes: TINY_PNG,
          contentType: 'image/png',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .insertImageAfter('AABB013B', {
          bytes: TINY_PNG,
          contentType: 'image/png',
          widthEmu: 200000,
          heightEmu: 200000,
        })
        .deleteParagraph('FFFFFFFF') // PARA_NOT_FOUND → 整个 plan 回滚
        .apply(),
    ).rejects.toThrow(WordError);

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    // 两张图都不残留（R1 盲区修复验证）
    expect(z.file('word/media/image1.png')).toBeNull();
    expect(z.file('word/media/image2.png')).toBeNull();
  });

  test('13c R1 混合 ext 回滚：1 png + 1 jpeg + 失败 op → 两个 media 文件都被删', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="AABB013C"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-13c', mode: 'inline' });

    await expect(
      doc
        .plan()
        .insertImageAfter('AABB013C', {
          bytes: TINY_PNG,
          contentType: 'image/png',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .insertImageAfter('AABB013C', {
          bytes: TINY_JPEG,
          contentType: 'image/jpeg',
          widthEmu: 100000,
          heightEmu: 100000,
        })
        .deleteParagraph('FFFFFFFF')
        .apply(),
    ).rejects.toThrow(WordError);

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    expect(z.file('word/media/image1.png')).toBeNull();
    expect(z.file('word/media/image1.jpeg')).toBeNull();
  });

  // ─────────── EMF/WMF 矢量图透传保留（task-emf-image）───────────
  // 真实文档（poi3-VariousPictures.docx）含 EMF/WMF 矢量图（粘贴的 Office 图表/示意图、
  // 老文档插图）。typeset 透传源图：getImageBytes 读出 image/x-emf 原始字节 →
  // insertImageAfter 重新引用。白名单此前不含 EMF/WMF → 真实文档排版崩溃。
  // 修复：把 image/x-emf/.emf、image/x-wmf/.wmf 加进支持列表，矢量图原样保留进产物。

  // 最小 EMF 字节（EMF header magic：record type 0x00000001 + 'EMF '/0x464D4520 签名）。
  // 仅作 part 字节透传验证，无需是合法可渲染 EMF。
  const TINY_EMF = new Uint8Array([
    0x01,
    0x00,
    0x00,
    0x00, // iType = EMR_HEADER
    0x6c,
    0x00,
    0x00,
    0x00, // nSize
    0x20,
    0x45,
    0x4d,
    0x46, // ' EMF' signature
  ]);
  // 最小 WMF 字节（placeable WMF magic：0x9AC6CDD7）。
  const TINY_WMF = new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a, 0x00, 0x00, 0x00, 0x00]);

  test('emf 透传：插入 image/x-emf → 产物保留 image1.emf（字节一致）+ CT Default Extension=emf', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="EEMF0001"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-emf', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('EEMF0001', {
        bytes: TINY_EMF,
        contentType: 'image/x-emf',
        widthEmu: 914400,
        heightEmu: 685800,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    // 矢量图 part 原样保留，扩展名 .emf
    const mediaFile = z.file('word/media/image1.emf');
    expect(mediaFile).not.toBeNull();
    const mediaBytes = new Uint8Array(await mediaFile!.async('arraybuffer'));
    expect(mediaBytes).toEqual(TINY_EMF);
    // content-type Default 登记
    const ctXml = await z.file('[Content_Types].xml')!.async('string');
    expect(ctXml).toContain('Extension="emf"');
    expect(ctXml).toContain('image/x-emf');
    // drawing 引用到位
    const relsXml = await z.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsXml).toContain('media/image1.emf');
  });

  test('wmf 透传：插入 image/x-wmf → 产物保留 image1.wmf（字节一致）+ CT Default Extension=wmf', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="EWMF0001"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-wmf', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('EWMF0001', {
        bytes: TINY_WMF,
        contentType: 'image/x-wmf',
        widthEmu: 500000,
        heightEmu: 400000,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const z = await JSZip.loadAsync(out);
    const mediaFile = z.file('word/media/image1.wmf');
    expect(mediaFile).not.toBeNull();
    const mediaBytes = new Uint8Array(await mediaFile!.async('arraybuffer'));
    expect(mediaBytes).toEqual(TINY_WMF);
    const ctXml = await z.file('[Content_Types].xml')!.async('string');
    expect(ctXml).toContain('Extension="wmf"');
    expect(ctXml).toContain('image/x-wmf');
  });

  test('emf round-trip：apply→toZip→fromZip→getImageBytes 字节 + contentType 一致（透传不损）', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="EEMF0002"><w:r><w:t>anchor</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'img-emf-rt', mode: 'inline' });
    await doc
      .plan()
      .insertImageAfter('EEMF0002', {
        bytes: TINY_EMF,
        contentType: 'image/x-emf',
        widthEmu: 914400,
        heightEmu: 685800,
      })
      .apply();

    const out = await doc.toZip();
    await doc.dispose();

    const doc2 = await WordDocument.fromZip(out, { docId: 'img-emf-rt-r', mode: 'inline' });
    const drawings = await doc2.drawings.list();
    expect(drawings.length).toBeGreaterThanOrEqual(1);
    const drawing = drawings[drawings.length - 1]!;
    const result = await drawing.getImageBytes();
    expect(result).not.toBeNull();
    expect(result!.bytes).toEqual(TINY_EMF);
    expect(result!.contentType).toBe('image/x-emf');
    await doc2.dispose();
  });
});

// ─────────── getImageBytes ───────────

describe('Drawing.getImageBytes', () => {
  /**
   * 构造含 inline picture 的 fixture（预先有图片的 zip）。
   * 用一个真实 drawing 结构（简化）让 drawingIdx 能索引到。
   */
  async function makeFixtureWithImage(): Promise<Uint8Array> {
    const docPicXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"
            xmlns:wp="${WP_NS}"
            xmlns:a="${A_NS}"
            xmlns:r="${R_NS}"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p w14:paraId="GG000001">
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="685800"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="1" name="image1"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="1" name="image1"/>
                    <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId2"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="685800"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

    const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

    const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;

    const zip = new JSZip();
    zip.file('[Content_Types].xml', ctXml);
    zip.file('_rels/.rels', MAIN_RELS);
    zip.file('word/document.xml', docPicXml);
    zip.file('word/_rels/document.xml.rels', docRelsXml);
    zip.file('word/media/image1.png', TINY_PNG);
    return zip.generateAsync({ type: 'uint8array' });
  }

  test('14 读 source 图：getImageBytes 返回 bytes（长度>0）+ contentType + extentEMU', async () => {
    const buf = await makeFixtureWithImage();
    const doc = await WordDocument.fromZip(buf, { docId: 'gib-14', mode: 'inline' });

    const drawings = await doc.drawings.list();
    expect(drawings.length).toBe(1);
    const drawing = drawings[0]!;

    const result = await drawing.getImageBytes();
    expect(result).not.toBeNull();
    expect(result!.bytes.length).toBeGreaterThan(0);
    expect(result!.bytes).toEqual(TINY_PNG);
    expect(result!.contentType).toBe('image/png');
    expect(result!.extentEMU.cx).toBe(914400);
    expect(result!.extentEMU.cy).toBe(685800);

    await doc.dispose();
  });

  test('14b drawings.list/get 返回 live Drawing，可读取图片并删除', async () => {
    const buf = await makeFixtureWithImage();
    const doc = await WordDocument.fromZip(buf, { docId: 'gib-14b', mode: 'inline' });

    const drawings = await doc.drawings.list();
    expect(drawings).toHaveLength(1);
    const drawing = drawings[0]!;
    expect(drawing.id).toBe(drawing.drawingId);
    expect(typeof drawing.getInfo).toBe('function');
    expect(typeof drawing.getImageBytes).toBe('function');
    expect(typeof drawing.delete).toBe('function');

    const sameDrawing = await doc.drawings.get(drawing.drawingId);
    expect(sameDrawing?.id).toBe(drawing.drawingId);
    expect((await sameDrawing?.getInfo())?.drawingId).toBe(drawing.drawingId);

    const image = await drawing.getImageBytes();
    expect(image?.bytes).toEqual(TINY_PNG);
    expect(image?.contentType).toBe('image/png');

    await drawing.delete();
    expect(await doc.drawings.list()).toHaveLength(0);

    await doc.dispose();
  });

  test('15 非 picture / 不存在：返回 null', async () => {
    const buf = await makeFixtureZip(
      `<w:p w14:paraId="GIB0015"><w:r><w:t>no drawing</w:t></w:r></w:p>`,
    );
    const doc = await WordDocument.fromZip(buf, { docId: 'gib-15', mode: 'inline' });

    const drawing = await doc.drawings.get('dr-9999');
    expect(drawing).toBeNull();

    await doc.dispose();
  });

  test('16 inline 模式可用：mode="inline" 下 getImageBytes 走通（merger 依赖）', async () => {
    const buf = await makeFixtureWithImage();
    const doc = await WordDocument.fromZip(buf, { docId: 'gib-16', mode: 'inline' });

    const drawings = await doc.drawings.list();
    const drawing = drawings[0]!;
    const result = await drawing.getImageBytes();

    // inline 模式下 Uint8Array 直接引用传递，不经 structuredClone
    expect(result).not.toBeNull();
    expect(result!.bytes).toBeInstanceOf(Uint8Array);
    expect(result!.bytes.length).toBeGreaterThan(0);

    await doc.dispose();
  });
});
