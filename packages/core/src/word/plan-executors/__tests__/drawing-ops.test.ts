/**
 * executeReplaceDrawingWithPlaceholder 单元测试(#190)。
 *
 * 覆盖:
 *   - 基本:inline picture → 替换 r:embed 指向共享 placeholder rId
 *   - 布局保持:wp:extent / wp:inline 包装 / wp:docPr 其它属性不动
 *   - 中文标注:wp:docPr/@descr / @title 被设为 "[图片占位]"
 *   - anchor(浮动)与 inline 同样工作
 *   - 多次调用幂等:placeholder.png 只建一次,rId 在该 part 复用
 *   - kind 校验:chart / smartArt / shape 抛 UNSUPPORTED_OOXML_FEATURE
 *   - drawingId 不存在抛 DRAWING_NOT_FOUND
 *   - trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE
 *   - [Content_Types].xml 自动补 `<Default Extension="png" .../>`
 *   - 占位 PNG bytes 正确入包(8-bit grayscale,839 字节)
 *
 * 引用计数清理(S4)、deleteDrawing(S5)在各自 PR 的独立 test 文件覆盖。
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createPackage,
  createXmlPart,
  getPart,
  hasPart,
  isXmlPart,
  setPart,
  type Package,
} from '../../../pack/index.js';
import { buildIndices } from '../../indices.js';
import { A_NS, MC_NS, PIC_NS, R_NS, W_NS, W14_NS, WP_NS, WPS_NS } from '../../namespaces.js';
import type { ApplyResult } from '../../plan.js';
import { WordError } from '../../word-error.js';
import type { ExecutorContext } from '../executor-context.js';
import {
  executeReplaceDrawingWithPlaceholder,
  executeDeleteDrawing,
  executeReplaceShapeInnerText,
} from '../drawing-ops.js';
import {
  PLACEHOLDER_DESCR,
  PLACEHOLDER_MEDIA_PATH,
  IMAGE_RELATIONSHIP_TYPE,
} from '../placeholder-image.js';
import { PLACEHOLDER_PNG_BYTES } from '../placeholder-image-bytes.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HEADER_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const CONTENT_TYPES_CTYPE = 'application/vnd.openxmlformats-package.content-types+xml';
const RELS_CTYPE = 'application/vnd.openxmlformats-package.relationships+xml';

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CTYPE}"/>
</Types>`;

const DOC_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/original.jpg"/>
</Relationships>`;

function putXml(pkg: Package, path: string, xml: string, contentType: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

function emptyApplyResult(): ApplyResult {
  return {
    insertedParagraphs: [],
    insertedBodyStartParaIds: [],
    createdComments: [],
    createdBookmarks: [],
    acceptedRevisions: 0,
    rejectedRevisions: 0,
    replacedTableCells: 0,
  };
}

/** 含 inline picture drawing 的最小 body */
const INLINE_PICTURE = `
<w:p w14:paraId="AAAAAAA1">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}" distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="3000000" cy="2000000"/>
        <wp:docPr id="1" name="Picture 1"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="img.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId7"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** 含 anchor 浮动 picture drawing */
const ANCHOR_PICTURE = `
<w:p w14:paraId="AAAAAAA2">
  <w:r>
    <w:drawing>
      <wp:anchor xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}"
                 distT="0" distB="0" distL="0" distR="0" behindDoc="0" locked="0"
                 layoutInCell="1" allowOverlap="1" simplePos="0" relativeHeight="251658240">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>100</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>200</wp:posOffset></wp:positionV>
        <wp:extent cx="2000000" cy="1500000"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <wp:docPr id="2" name="Anchor 2"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr><pic:cNvPr id="0" name="img2.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </w:r>
</w:p>`;

/** 含 chart drawing —— uri 指 chart,不是 picture */
const INLINE_CHART = `
<w:p w14:paraId="AAAAAAA3">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="3000000" cy="2000000"/>
        <wp:docPr id="3" name="Chart 1"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId7"/>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

function buildPkg(
  bodyInner: string,
  opts: {
    contentTypes?: string;
    docRels?: string;
  } = {},
): Package {
  const pkg = createPackage();
  putXml(pkg, '[Content_Types].xml', opts.contentTypes ?? CONTENT_TYPES_XML, CONTENT_TYPES_CTYPE);
  putXml(pkg, 'word/_rels/document.xml.rels', opts.docRels ?? DOC_RELS_XML, RELS_CTYPE);
  putXml(
    pkg,
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${bodyInner}</w:body></w:document>`,
    DOC_CTYPE,
  );
  return pkg;
}

function buildCtx(pkg: Package, opts: { trackChanges?: boolean } = {}): ExecutorContext {
  const indices = buildIndices(pkg);
  return {
    pkg,
    indices,
    options: {
      trackChanges: opts.trackChanges ?? false,
      author: '',
      date: new Date('2026-05-22T00:00:00Z'),
    },
    result: emptyApplyResult(),
    opIndex: 0,
    docId: 'drawing-test',
    orphanMediaCandidates: new Set<string>(),
  };
}

function firstDrawingId(ctx: ExecutorContext): string {
  const id = [...ctx.indices.drawingIdx.keys()][0];
  if (!id) throw new Error('no drawing in fixture');
  return id;
}

// ─────────── tests ───────────

describe('executeReplaceDrawingWithPlaceholder —— 基本路径(#190 S3)', () => {
  test('inline picture:r:embed 改写为 placeholder rId,wp:extent 不动', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    // a:blip/@r:embed 改了
    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    const blip = drawing.getElementsByTagNameNS(A_NS, 'blip')[0]!;
    const newRId = blip.getAttributeNS(R_NS, 'embed');
    expect(newRId).toBeTruthy();
    expect(newRId).not.toBe('rId7');

    // wp:extent 完全保留
    const extent = drawing.getElementsByTagNameNS(WP_NS, 'extent')[0]!;
    expect(extent.getAttribute('cx')).toBe('3000000');
    expect(extent.getAttribute('cy')).toBe('2000000');

    // 仍然是 inline(没有变成 anchor)
    expect(drawing.getElementsByTagNameNS(WP_NS, 'inline')).toHaveLength(1);
    expect(drawing.getElementsByTagNameNS(WP_NS, 'anchor')).toHaveLength(0);
  });

  test('wp:docPr/@descr + @title 被设为 "[图片占位]"', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    const docPr = drawing.getElementsByTagNameNS(WP_NS, 'docPr')[0]!;
    expect(docPr.getAttribute('descr')).toBe(PLACEHOLDER_DESCR);
    expect(docPr.getAttribute('title')).toBe(PLACEHOLDER_DESCR);
    // 原有 id / name 不变
    expect(docPr.getAttribute('id')).toBe('1');
    expect(docPr.getAttribute('name')).toBe('Picture 1');
  });

  test('placeholder.png 被入包(binary part) + 字节正确', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    expect(hasPart(pkg, PLACEHOLDER_MEDIA_PATH)).toBe(false);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const mediaPart = getPart(pkg, PLACEHOLDER_MEDIA_PATH);
    expect(mediaPart).toBeDefined();
    expect(mediaPart!.contentType).toBe('image/png');
    expect(mediaPart!.bytes).toBe(PLACEHOLDER_PNG_BYTES);
    expect(mediaPart!._dirty).toBe(true);
  });

  test('[Content_Types].xml 自动补 Default Extension="png"', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const ctPart = getPart(pkg, '[Content_Types].xml')!;
    expect(isXmlPart(ctPart)).toBe(true);
    const defaults = ctPart.xmlDom!.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/package/2006/content-types',
      'Default',
    );
    let hasPng = false;
    for (let i = 0; i < defaults.length; i++) {
      if (defaults[i]!.getAttribute('Extension')?.toLowerCase() === 'png') {
        expect(defaults[i]!.getAttribute('ContentType')).toBe('image/png');
        hasPng = true;
      }
    }
    expect(hasPng).toBe(true);
  });

  test('document.xml.rels 添加 image rel 指向 placeholder.png', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const blip = ctx.indices.drawingIdx.get(drawingId)!.getElementsByTagNameNS(A_NS, 'blip')[0]!;
    const newRId = blip.getAttributeNS(R_NS, 'embed')!;

    const relsPart = getPart(pkg, 'word/_rels/document.xml.rels')!;
    const rels = relsPart.xmlDom!.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/package/2006/relationships',
      'Relationship',
    );
    let found = false;
    for (let i = 0; i < rels.length; i++) {
      if (rels[i]!.getAttribute('Id') === newRId) {
        expect(rels[i]!.getAttribute('Type')).toBe(IMAGE_RELATIONSHIP_TYPE);
        expect(rels[i]!.getAttribute('Target')).toBe('media/placeholder.png');
        found = true;
      }
    }
    expect(found).toBe(true);
  });

  test('anchor 浮动 picture:r:embed 改写,wp:wrapSquare 与 positionH/V 保留', () => {
    const pkg = buildPkg(ANCHOR_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    expect(drawing.getElementsByTagNameNS(WP_NS, 'anchor')).toHaveLength(1);
    expect(drawing.getElementsByTagNameNS(WP_NS, 'wrapSquare')).toHaveLength(1);
    expect(drawing.getElementsByTagNameNS(WP_NS, 'positionH')).toHaveLength(1);
    expect(drawing.getElementsByTagNameNS(WP_NS, 'positionV')).toHaveLength(1);

    const blip = drawing.getElementsByTagNameNS(A_NS, 'blip')[0]!;
    expect(blip.getAttributeNS(R_NS, 'embed')).not.toBe('rId7');
  });

  test('多次调用幂等:同 part 内 placeholder rId 复用,binary part 不重复创建', () => {
    const pkg = buildPkg(INLINE_PICTURE + INLINE_PICTURE.replace('AAAAAAA1', 'AAAAAAA9'));
    const ctx = buildCtx(pkg);
    const ids = [...ctx.indices.drawingIdx.keys()];
    expect(ids).toHaveLength(2);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId: ids[0]!,
    });
    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId: ids[1]!,
    });

    // 两张 drawing 都指向同一个 placeholder rId
    const blips = pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(A_NS, 'blip');
    expect(blips).toHaveLength(2);
    const r1 = (blips[0]! as Element).getAttributeNS(R_NS, 'embed');
    const r2 = (blips[1]! as Element).getAttributeNS(R_NS, 'embed');
    expect(r1).toBe(r2);

    // rels 只多了一条 image rel(原 rId7 暂留待 S4 清理)
    const rels = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    let placeholderCount = 0;
    for (let i = 0; i < rels.length; i++) {
      if (
        rels[i]!.getAttribute('Type') === IMAGE_RELATIONSHIP_TYPE &&
        rels[i]!.getAttribute('Target') === 'media/placeholder.png'
      ) {
        placeholderCount++;
      }
    }
    expect(placeholderCount).toBe(1);
  });
});

describe('executeReplaceDrawingWithPlaceholder —— 校验与错误', () => {
  test('drawingId 不存在抛 DRAWING_NOT_FOUND', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    expect(() =>
      executeReplaceDrawingWithPlaceholder(ctx, {
        type: 'replaceDrawingWithPlaceholder',
        drawingId: 'dr-9999',
      }),
    ).toThrow(WordError);
    try {
      executeReplaceDrawingWithPlaceholder(ctx, {
        type: 'replaceDrawingWithPlaceholder',
        drawingId: 'dr-9999',
      });
    } catch (e) {
      expect((e as WordError).code).toBe('DRAWING_NOT_FOUND');
    }
  });

  test('chart kind 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = buildPkg(INLINE_CHART);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);
    try {
      executeReplaceDrawingWithPlaceholder(ctx, {
        type: 'replaceDrawingWithPlaceholder',
        drawingId,
      });
      expect.fail('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toContain('picture');
    }
  });

  test('trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg, { trackChanges: true });
    const drawingId = firstDrawingId(ctx);
    expect(() =>
      executeReplaceDrawingWithPlaceholder(ctx, {
        type: 'replaceDrawingWithPlaceholder',
        drawingId,
      }),
    ).toThrowError(/trackChanges=true/);
  });
});

describe('executeReplaceDrawingWithPlaceholder —— header part 场景', () => {
  test('header part 内 picture:rels 写到 word/_rels/header1.xml.rels', () => {
    const pkg = buildPkg('');
    // 追加 header part + 其 rels
    const HEADER_XML = `<w:hdr xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">${INLINE_PICTURE.replace('AAAAAAA1', 'HHHHHHH1')}</w:hdr>`;
    putXml(pkg, 'word/header1.xml', HEADER_XML, HEADER_CTYPE);
    const HEADER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/header-orig.png"/>
</Relationships>`;
    putXml(pkg, 'word/_rels/header1.xml.rels', HEADER_RELS, RELS_CTYPE);

    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);
    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    const headerRels = pkg.parts
      .get('word/_rels/header1.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    let placeholderInHeaderRels = false;
    for (let i = 0; i < headerRels.length; i++) {
      if (
        headerRels[i]!.getAttribute('Type') === IMAGE_RELATIONSHIP_TYPE &&
        headerRels[i]!.getAttribute('Target') === 'media/placeholder.png'
      ) {
        placeholderInHeaderRels = true;
      }
    }
    expect(placeholderInHeaderRels).toBe(true);

    // document.xml.rels 不应被改(header 的 rId 隔离于 header.xml.rels)
    const docRels = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    for (let i = 0; i < docRels.length; i++) {
      expect(docRels[i]!.getAttribute('Target')).not.toBe('media/placeholder.png');
    }
  });
});

describe('executeReplaceDrawingWithPlaceholder —— S4 引用计数清理', () => {
  test('单 drawing 替换:旧 rId 在 host 内归零 → 删 rels 条目 + 加入 orphan 候选', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId,
    });

    // rels DOM 里旧 rId7 已被移除
    const rels = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    for (let i = 0; i < rels.length; i++) {
      expect(rels[i]!.getAttribute('Id')).not.toBe('rId7');
    }

    // orphan 候选已记录 media/original.jpg
    expect([...ctx.orphanMediaCandidates!]).toContain('word/media/original.jpg');
  });

  test('两张 drawing 共享同一原图:第一张替换后旧 rId 残余引用 > 0,rels 不动', () => {
    // 两张 picture drawing 都用 rId7
    const pkg = buildPkg(INLINE_PICTURE + INLINE_PICTURE.replace('AAAAAAA1', 'AAAAAAA9'));
    const ctx = buildCtx(pkg);
    const ids = [...ctx.indices.drawingIdx.keys()];
    expect(ids).toHaveLength(2);

    // 只替换第一张
    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId: ids[0]!,
    });

    // host 内 rId7 仍被第二张 drawing 引用 → rels 条目不动
    const relsAfter1 = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    let stillHasRId7 = false;
    for (let i = 0; i < relsAfter1.length; i++) {
      if (relsAfter1[i]!.getAttribute('Id') === 'rId7') stillHasRId7 = true;
    }
    expect(stillHasRId7).toBe(true);
    expect(ctx.orphanMediaCandidates!.size).toBe(0);

    // 替换第二张后 rId7 在 host 内归零 → rels 条目删除 + 候选记录
    executeReplaceDrawingWithPlaceholder(ctx, {
      type: 'replaceDrawingWithPlaceholder',
      drawingId: ids[1]!,
    });
    const relsAfter2 = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    for (let i = 0; i < relsAfter2.length; i++) {
      expect(relsAfter2[i]!.getAttribute('Id')).not.toBe('rId7');
    }
    expect([...ctx.orphanMediaCandidates!]).toContain('word/media/original.jpg');
  });

  test('countMediaReferences 与 sweepOrphanMediaParts:正确处理跨 part 共享 media', async () => {
    const { countMediaReferences, sweepOrphanMediaParts } = await import('../placeholder-image.js');

    // 构造:document.xml + header1.xml 两个 part 都引用 media/shared.png
    const pkg = createPackage();
    putXml(
      pkg,
      'word/_rels/document.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/shared.png"/>
</Relationships>`,
      RELS_CTYPE,
    );
    putXml(
      pkg,
      'word/_rels/header1.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${IMAGE_RELATIONSHIP_TYPE}" Target="media/shared.png"/>
</Relationships>`,
      RELS_CTYPE,
    );

    // 添加 binary part
    const { createBinaryPart, setPart } = await import('../../../pack/index.js');
    setPart(
      pkg,
      createBinaryPart({
        path: 'word/media/shared.png',
        contentType: 'image/png',
        bytes: new Uint8Array([1, 2, 3, 4]),
      }),
    );

    // 跨 part 引用数 = 2
    expect(countMediaReferences(pkg, 'word/media/shared.png')).toBe(2);

    // 用 sweep 删除候选:此时未归零,不删
    const removed1 = sweepOrphanMediaParts(pkg, ['word/media/shared.png']);
    expect(removed1).toEqual([]);
    expect(pkg.parts.has('word/media/shared.png')).toBe(true);

    // 移除 document.xml.rels 的条目后 ref=1,仍不删
    const doc = pkg.parts.get('word/_rels/document.xml.rels')!.xmlDom!;
    const docRel = doc.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/package/2006/relationships',
      'Relationship',
    )[0]!;
    docRel.parentNode!.removeChild(docRel);
    expect(countMediaReferences(pkg, 'word/media/shared.png')).toBe(1);
    expect(sweepOrphanMediaParts(pkg, ['word/media/shared.png'])).toEqual([]);
    expect(pkg.parts.has('word/media/shared.png')).toBe(true);

    // 移除 header1 的条目后 ref=0,sweep 真的删
    const hdr = pkg.parts.get('word/_rels/header1.xml.rels')!.xmlDom!;
    const hdrRel = hdr.getElementsByTagNameNS(
      'http://schemas.openxmlformats.org/package/2006/relationships',
      'Relationship',
    )[0]!;
    hdrRel.parentNode!.removeChild(hdrRel);
    expect(countMediaReferences(pkg, 'word/media/shared.png')).toBe(0);
    expect(sweepOrphanMediaParts(pkg, ['word/media/shared.png'])).toEqual([
      'word/media/shared.png',
    ]);
    expect(pkg.parts.has('word/media/shared.png')).toBe(false);
  });

  test('resolveRelTarget:常见路径形态', async () => {
    const { resolveRelTarget } = await import('../placeholder-image.js');
    expect(resolveRelTarget('word/_rels/document.xml.rels', 'media/x.png')).toBe(
      'word/media/x.png',
    );
    expect(resolveRelTarget('word/_rels/header1.xml.rels', 'media/x.png')).toBe('word/media/x.png');
    expect(resolveRelTarget('word/_rels/document.xml.rels', './media/x.png')).toBe(
      'word/media/x.png',
    );
    expect(resolveRelTarget('word/_rels/document.xml.rels', '/word/media/x.png')).toBe(
      'word/media/x.png',
    );
    // 顶层 .rels 基为包根
    expect(resolveRelTarget('_rels/.rels', 'word/document.xml')).toBe('word/document.xml');
  });
});

describe('executeDeleteDrawing(#190 S5)', () => {
  test('picture drawing 删除:w:drawing 节点消失,父 w:r 仅含 drawing 时整段删', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId });

    // 整 part 内不应再含任何 w:drawing
    const drawings = pkg.parts
      .get('word/document.xml')!
      .xmlDom!.getElementsByTagNameNS(W_NS, 'drawing');
    expect(drawings).toHaveLength(0);

    // 父段落不应残留空 w:r(整段 w:r 已删)
    const runs = pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'r');
    expect(runs).toHaveLength(0);

    // drawingIdx 同步删条目
    expect(ctx.indices.drawingIdx.has(drawingId)).toBe(false);
  });

  test('父 w:r 含 drawing + 文本时,删 drawing 但保留 w:r(因为还有内容)', () => {
    const MIXED = `
<w:p w14:paraId="AAAAAAA5">
  <w:r>
    <w:t>before</w:t>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:pic="${PIC_NS}">
        <wp:extent cx="1" cy="1"/>
        <wp:docPr id="9" name="pic"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
    <w:t>after</w:t>
  </w:r>
</w:p>`;
    const pkg = buildPkg(MIXED);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId });

    // drawing 删了
    const drawings = pkg.parts
      .get('word/document.xml')!
      .xmlDom!.getElementsByTagNameNS(W_NS, 'drawing');
    expect(drawings).toHaveLength(0);

    // w:r 仍在(还有 w:t)
    const runs = pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'r');
    expect(runs).toHaveLength(1);
    const ts = runs[0]!.getElementsByTagNameNS(W_NS, 't');
    expect(ts).toHaveLength(2);
  });

  test('chart kind 也可删除(deleteDrawing 不限制 kind)', () => {
    const pkg = buildPkg(INLINE_CHART);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    expect(() => executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId })).not.toThrow();
    expect(
      pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'drawing'),
    ).toHaveLength(0);
  });

  test('drawingId 不存在抛 DRAWING_NOT_FOUND', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    try {
      executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId: 'dr-9999' });
      expect.fail('should throw');
    } catch (e) {
      expect((e as WordError).code).toBe('DRAWING_NOT_FOUND');
    }
  });

  test('trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg, { trackChanges: true });
    const drawingId = firstDrawingId(ctx);
    expect(() => executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId })).toThrowError(
      /trackChanges=true/,
    );
  });

  test('rels / media 引用清理:独占 media 路径成为 orphan 候选', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeDeleteDrawing(ctx, { type: 'deleteDrawing', drawingId });

    // rels 内 rId7 删除
    const rels = pkg.parts
      .get('word/_rels/document.xml.rels')!
      .xmlDom!.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/package/2006/relationships',
        'Relationship',
      );
    for (let i = 0; i < rels.length; i++) {
      expect(rels[i]!.getAttribute('Id')).not.toBe('rId7');
    }
    expect([...ctx.orphanMediaCandidates!]).toContain('word/media/original.jpg');
  });
});

describe('placeholder.png 字节自校验', () => {
  test('839 字节,PNG signature + IHDR 200x150 8-bit grayscale', () => {
    const b = PLACEHOLDER_PNG_BYTES;
    expect(b.length).toBe(839);
    // PNG signature
    expect(b[0]).toBe(0x89);
    expect(b[1]).toBe(0x50);
    expect(b[2]).toBe(0x4e);
    expect(b[3]).toBe(0x47);
    // IHDR width / height(big-endian)
    const width = (b[16]! << 24) | (b[17]! << 16) | (b[18]! << 8) | b[19]!;
    const height = (b[20]! << 24) | (b[21]! << 16) | (b[22]! << 8) | b[23]!;
    expect(width).toBe(200);
    expect(height).toBe(150);
    // bit depth = 8, color type = 0 (grayscale)
    expect(b[24]).toBe(8);
    expect(b[25]).toBe(0);
  });
});

// ─────────── #210 replaceShapeInnerText fixtures ───────────

/** 单段单 run 的 textBox(最简形态) */
const INLINE_TEXTBOX_SIMPLE = `
<w:p w14:paraId="TXT00001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}" distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="2000000" cy="1000000"/>
        <wp:docPr id="11" name="TextBox 1"/>
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
                    <w:r><w:t>Original text inside the textbox</w:t></w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** 带艺术字 rPr（红色 + sz56）的 textBox —— 红头机关名一类，验证 rPr 继承 */
const INLINE_TEXTBOX_ARTISTIC = `
<w:p w14:paraId="ART00001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}">
        <wp:extent cx="3000000" cy="800000"/>
        <wp:docPr id="21" name="RedHead Org"/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:spPr>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  <w:p w14:paraId="ART00002">
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:color w:val="FF0000"/>
                        <w:sz w:val="56"/>
                        <w:szCs w:val="56"/>
                      </w:rPr>
                      <w:t>江临市滨河区人民政府</w:t>
                    </w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** rPr 仅在段落 pPr/rPr（run 自身无 rPr）的 textBox —— 验证 pickFirstRunRPr 兜底 */
const INLINE_TEXTBOX_PPRRPR = `
<w:p w14:paraId="PPR00001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}">
        <wp:extent cx="3000000" cy="800000"/>
        <wp:docPr id="22" name="PprRpr Box"/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  <w:p w14:paraId="PPR00002">
                    <w:pPr><w:rPr><w:sz w:val="80"/></w:rPr></w:pPr>
                    <w:r><w:t>专题会议纪要</w:t></w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** 多段 + 首段 jc=center 的 textBox */
const INLINE_TEXTBOX_MULTI_PARA = `
<w:p w14:paraId="TXT10001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}">
        <wp:extent cx="2000000" cy="1500000"/>
        <wp:docPr id="12" name="TextBox 2"/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  <w:p w14:paraId="TXT10002">
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r><w:t>line 1</w:t></w:r>
                  </w:p>
                  <w:p w14:paraId="TXT10003">
                    <w:r><w:t>line 2</w:t></w:r>
                  </w:p>
                  <w:p w14:paraId="TXT10004">
                    <w:r><w:t>line 3</w:t></w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** textBox 内含 hyperlink(r:id 指向 rId7) */
const INLINE_TEXTBOX_HYPERLINK = `
<w:p w14:paraId="TXT20001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}">
        <wp:extent cx="2000000" cy="1000000"/>
        <wp:docPr id="13" name="TextBox 3"/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  <w:p w14:paraId="TXT20002">
                    <w:hyperlink r:id="rId7">
                      <w:r><w:t>click me</w:t></w:r>
                    </w:hyperlink>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

/** shape(无 txbxContent,仅几何) */
const INLINE_SHAPE_NO_TEXT = `
<w:p w14:paraId="SHP00001">
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:wps="${WPS_NS}">
        <wp:extent cx="500000" cy="500000"/>
        <wp:docPr id="14" name="Shape 1"/>
        <a:graphic>
          <a:graphicData uri="${WPS_NS}">
            <wps:wsp>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="500000" cy="500000"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </wps:spPr>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>`;

// ─────────── #210 replaceShapeInnerText 测试 ───────────

describe('executeReplaceShapeInnerText —— 基本路径(#210)', () => {
  test('单段 textBox:txbxContent 内容被替换为单段单 run 占位文字', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '[Shape 内容占位]',
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    const txbxContents = drawing.getElementsByTagNameNS(W_NS, 'txbxContent');
    expect(txbxContents).toHaveLength(1);
    const txbx = txbxContents[0]!;

    // 单段
    const ps = snapshot(txbx.getElementsByTagNameNS(W_NS, 'p'));
    expect(ps).toHaveLength(1);
    // 段内单 run 单 w:t,内容为占位
    const runs = snapshot(ps[0]!.getElementsByTagNameNS(W_NS, 'r'));
    expect(runs).toHaveLength(1);
    const ts = snapshot(runs[0]!.getElementsByTagNameNS(W_NS, 't'));
    expect(ts).toHaveLength(1);
    expect(ts[0]!.textContent).toBe('[Shape 内容占位]');
    // 原文本不残留
    const allTexts = snapshot(txbx.getElementsByTagNameNS(W_NS, 't'));
    for (const t of allTexts) {
      expect(t.textContent).not.toContain('Original text');
    }
  });

  test('外壳完全保留:wp:inline / wp:extent / wps:spPr / a:solidFill / a:ln 字节级不变', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '[Shape 内容占位]',
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    // wp:inline 仍是 inline 而非 anchor
    expect(drawing.getElementsByTagNameNS(WP_NS, 'inline')).toHaveLength(1);
    expect(drawing.getElementsByTagNameNS(WP_NS, 'anchor')).toHaveLength(0);

    // wp:extent 不动
    const extent = drawing.getElementsByTagNameNS(WP_NS, 'extent')[0]!;
    expect(extent.getAttribute('cx')).toBe('2000000');
    expect(extent.getAttribute('cy')).toBe('1000000');

    // wps:spPr 内的填充 / 边框 / 几何全部保留
    const solidFill = drawing.getElementsByTagNameNS(A_NS, 'solidFill');
    expect(solidFill.length).toBeGreaterThan(0);
    const srgbClr = drawing.getElementsByTagNameNS(A_NS, 'srgbClr');
    expect(srgbClr.length).toBeGreaterThan(0);
    expect(srgbClr[0]!.getAttribute('val')).toBe('FFFFCC');
    const ln = drawing.getElementsByTagNameNS(A_NS, 'ln')[0]!;
    expect(ln.getAttribute('w')).toBe('6350');
    const prstGeom = drawing.getElementsByTagNameNS(A_NS, 'prstGeom')[0]!;
    expect(prstGeom.getAttribute('prst')).toBe('rect');
  });

  test('多段 textBox:压缩为单段,首段 w:jc=center 保留', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_MULTI_PARA);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '占位',
    });

    const txbx = ctx.indices.drawingIdx
      .get(drawingId)!
      .getElementsByTagNameNS(W_NS, 'txbxContent')[0]!;
    const ps = snapshot(txbx.getElementsByTagNameNS(W_NS, 'p'));
    expect(ps).toHaveLength(1);

    // jc 保留为 center
    const jcs = snapshot(ps[0]!.getElementsByTagNameNS(W_NS, 'jc'));
    expect(jcs).toHaveLength(1);
    expect(jcs[0]!.getAttributeNS(W_NS, 'val')).toBe('center');

    // 原 line 2 / line 3 / line 1 内容都消失
    const text = txbx.textContent ?? '';
    expect(text).toBe('占位');
  });

  test('rPr 继承:替换后新 run 保留原首 run 的 color/sz(红头艺术字字段化,解除默认字体限制)', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_ARTISTIC);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '{发文机关标志}',
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    const txbx = drawing.getElementsByTagNameNS(W_NS, 'txbxContent')[0]!;

    // 单段单 run，文字为占位
    const ps = snapshot(txbx.getElementsByTagNameNS(W_NS, 'p'));
    expect(ps).toHaveLength(1);
    const runs = snapshot(ps[0]!.getElementsByTagNameNS(W_NS, 'r'));
    expect(runs).toHaveLength(1);
    expect(runs[0]!.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('{发文机关标志}');

    // 关键：新 run 含继承自原 run 的 rPr，color=FF0000 / sz=56 未丢
    const rPr = snapshot(runs[0]!.getElementsByTagNameNS(W_NS, 'rPr'));
    expect(rPr.length, 'run 应含继承的 rPr').toBe(1);
    const color = rPr[0]!.getElementsByTagNameNS(W_NS, 'color')[0]!;
    expect(color.getAttributeNS(W_NS, 'val') ?? color.getAttribute('w:val')).toBe('FF0000');
    const sz = rPr[0]!.getElementsByTagNameNS(W_NS, 'sz')[0]!;
    expect(sz.getAttributeNS(W_NS, 'val') ?? sz.getAttribute('w:val')).toBe('56');

    // rPr 必须是 run 首子（schema：w:rPr 在 w:t 之前）
    const firstElChild = Array.from(runs[0]!.childNodes).find((n) => n.nodeType === 1) as Element;
    expect(firstElChild.localName).toBe('rPr');
  });

  test('rPr 兜底:run 无 rPr 但段落 pPr/rPr 有 → 占位 run 继承 pPr/rPr', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_PPRRPR);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '{标题}',
    });

    const txbx = ctx.indices.drawingIdx
      .get(drawingId)!
      .getElementsByTagNameNS(W_NS, 'txbxContent')[0]!;
    const run = txbx.getElementsByTagNameNS(W_NS, 'r')[0]!;
    const rPr = snapshot(run.getElementsByTagNameNS(W_NS, 'rPr'));
    expect(rPr.length, 'pickFirstRunRPr 兜底 pPr/rPr').toBe(1);
    const sz = rPr[0]!.getElementsByTagNameNS(W_NS, 'sz')[0]!;
    expect(sz.getAttributeNS(W_NS, 'val') ?? sz.getAttribute('w:val')).toBe('80');
    const pPr = snapshot(txbx.getElementsByTagNameNS(W_NS, 'pPr'));
    expect(pPr.length, 'shape first paragraph pPr should be preserved').toBe(1);
    const pPrSz = pPr[0]!.getElementsByTagNameNS(W_NS, 'sz')[0]!;
    expect(pPrSz.getAttributeNS(W_NS, 'val') ?? pPrSz.getAttribute('w:val')).toBe('80');
  });

  test('无 rPr 来源:裸 run（向后兼容,语义不变,不报错）', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE); // run 无 rPr，pPr 也无 rPr
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '占位',
    });

    const txbx = ctx.indices.drawingIdx
      .get(drawingId)!
      .getElementsByTagNameNS(W_NS, 'txbxContent')[0]!;
    const run = txbx.getElementsByTagNameNS(W_NS, 'r')[0]!;
    // 无可继承 rPr → 裸 run（与旧行为一致）
    expect(snapshot(run.getElementsByTagNameNS(W_NS, 'rPr'))).toHaveLength(0);
    expect(run.getElementsByTagNameNS(W_NS, 't')[0]!.textContent).toBe('占位');
  });

  test('textBox 含 hyperlink:替换后 rId7 在 host 内归零 → rels 清理 + orphan 候选记录', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_HYPERLINK);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    // 替换前:rels 内 rId7 存在
    const relsBefore = snapshot(
      pkg.parts
        .get('word/_rels/document.xml.rels')!
        .xmlDom!.getElementsByTagNameNS(
          'http://schemas.openxmlformats.org/package/2006/relationships',
          'Relationship',
        ),
    );
    expect(relsBefore.some((r) => r.getAttribute('Id') === 'rId7')).toBe(true);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '占位',
    });

    // 替换后:hyperlink 没了
    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    expect(drawing.getElementsByTagNameNS(W_NS, 'hyperlink')).toHaveLength(0);

    // rels 内 rId7 也删了(host 内已归零)
    const relsAfter = snapshot(
      pkg.parts
        .get('word/_rels/document.xml.rels')!
        .xmlDom!.getElementsByTagNameNS(
          'http://schemas.openxmlformats.org/package/2006/relationships',
          'Relationship',
        ),
    );
    for (const r of relsAfter) {
      expect(r.getAttribute('Id')).not.toBe('rId7');
    }
    expect([...ctx.orphanMediaCandidates!]).toContain('word/media/original.jpg');
  });

  test('placeholder 含首尾空格:w:t/@xml:space=preserve 保留', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '  [带前后空格的占位]  ',
    });

    const drawing = ctx.indices.drawingIdx.get(drawingId)!;
    const t = drawing.getElementsByTagNameNS(W_NS, 't')[0]!;
    expect(t.textContent).toBe('  [带前后空格的占位]  ');
    // xml:space="preserve" 一律挂(_shared.createT 的契约)
    expect(t.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space')).toBe('preserve');
  });

  test('hostPart markDirty 触发,后续序列化可见', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);

    executeReplaceShapeInnerText(ctx, {
      type: 'replaceShapeInnerText',
      drawingId,
      placeholder: '[占位]',
    });

    const hostPart = pkg.parts.get('word/document.xml')!;
    expect(hostPart._dirty).toBe(true);
  });
});

describe('executeReplaceShapeInnerText —— 校验与错误', () => {
  test('drawingId 不存在抛 DRAWING_NOT_FOUND', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg);
    try {
      executeReplaceShapeInnerText(ctx, {
        type: 'replaceShapeInnerText',
        drawingId: 'dr-9999',
        placeholder: 'x',
      });
      expect.fail('should throw');
    } catch (e) {
      expect((e as WordError).code).toBe('DRAWING_NOT_FOUND');
    }
  });

  test('picture(kind=picture)抛 UNSUPPORTED_OOXML_FEATURE,信息指向 replaceDrawingWithPlaceholder', () => {
    const pkg = buildPkg(INLINE_PICTURE);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);
    try {
      executeReplaceShapeInnerText(ctx, {
        type: 'replaceShapeInnerText',
        drawingId,
        placeholder: 'x',
      });
      expect.fail('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toContain('replaceDrawingWithPlaceholder');
    }
  });

  test('chart(kind=chart)抛 UNSUPPORTED_OOXML_FEATURE,信息指向 deleteDrawing', () => {
    const pkg = buildPkg(INLINE_CHART);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);
    try {
      executeReplaceShapeInnerText(ctx, {
        type: 'replaceShapeInnerText',
        drawingId,
        placeholder: 'x',
      });
      expect.fail('should throw');
    } catch (e) {
      expect(e).toBeInstanceOf(WordError);
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toContain('deleteDrawing');
    }
  });

  test('shape 无 txbxContent 抛 UNSUPPORTED_OOXML_FEATURE,信息指向 deleteDrawing', () => {
    const pkg = buildPkg(INLINE_SHAPE_NO_TEXT);
    const ctx = buildCtx(pkg);
    const drawingId = firstDrawingId(ctx);
    try {
      executeReplaceShapeInnerText(ctx, {
        type: 'replaceShapeInnerText',
        drawingId,
        placeholder: 'x',
      });
      expect.fail('should throw');
    } catch (e) {
      expect((e as WordError).code).toBe('UNSUPPORTED_OOXML_FEATURE');
      expect((e as WordError).message).toContain('deleteDrawing');
    }
  });

  test('trackChanges=true 抛 UNSUPPORTED_OOXML_FEATURE', () => {
    const pkg = buildPkg(INLINE_TEXTBOX_SIMPLE);
    const ctx = buildCtx(pkg, { trackChanges: true });
    const drawingId = firstDrawingId(ctx);
    expect(() =>
      executeReplaceShapeInnerText(ctx, {
        type: 'replaceShapeInnerText',
        drawingId,
        placeholder: 'x',
      }),
    ).toThrowError(/trackChanges=true/);
  });
});

// 防 unused-import 警告(MC_NS 在跨 ns 测试场景可能用到,这里 keep import for symmetry)
void MC_NS;

// 工具:用 snapshot 包装 NodeList,避免 TypeScript 类型与 forEach 兼容问题
function snapshot<T extends Element>(col: { length: number; [idx: number]: T }): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) out.push(col[i]!);
  return out;
}
