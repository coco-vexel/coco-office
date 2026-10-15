/**
 * Drawing 枚举测试（#189）。
 *
 * 验证：
 *  - inline / anchor 锚定方式区分
 *  - picture / chart / smartArt / shape / textBox / group / canvas 形态识别
 *  - extent / wrap / position 字段解析
 *  - anchorParaId / runIndex 段落引用
 *  - isInHeaderFooter（页眉页脚 part 识别）
 *  - Shape 形态信号 hasBorder / hasFill / hasInnerText / innerTextAlign
 *  - drawingIdx 装载期分配 'dr-NNNN' ID
 *  - scope 'body' vs 'all' 行为
 */
import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../../pack/index.js';
import { buildIndices } from '../indices.js';
import { listDrawings, getDrawingById } from '../list-drawings.js';
import {
  W_NS,
  W14_NS,
  WP_NS,
  A_NS,
  PIC_NS,
  CHART_NS,
  DGM_NS,
  WPS_NS,
  WPG_NS,
  MC_NS,
} from '../namespaces.js';

const DOC_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const HDR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FTR_CTYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const COMMENTS_CTYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

/** 标准命名空间声明：方便组装测试 fixture */
const NS_DECL =
  ` xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}"` +
  ` xmlns:pic="${PIC_NS}" xmlns:c="${CHART_NS}" xmlns:dgm="${DGM_NS}"` +
  ` xmlns:wps="${WPS_NS}" xmlns:wpg="${WPG_NS}" xmlns:mc="${MC_NS}"`;

function putXmlPart(pkg: Package, path: string, xml: string, contentType = DOC_CTYPE): void {
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

/** 包一个内联 picture wp:inline，extent cx/cy、graphicData uri 都可参数化 */
function inlinePictureDrawing(cx = 914400, cy = 685800): string {
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="${cx}" cy="${cy}"/>
      <a:graphic>
        <a:graphicData uri="${PIC_NS}">
          <pic:pic/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

/** 浮动 picture wp:anchor，含 wrap + position */
function anchorPictureDrawing(
  opts: {
    cx?: number;
    cy?: number;
    wrap?: 'wrapNone' | 'wrapSquare' | 'wrapTight' | 'wrapTopAndBottom' | 'wrapThrough';
    behindDoc?: '0' | '1';
    posH?: { relativeFrom: string; offset?: number; align?: string };
    posV?: { relativeFrom: string; offset?: number; align?: string };
  } = {},
): string {
  const { cx = 914400, cy = 685800, wrap = 'wrapSquare', behindDoc, posH, posV } = opts;

  const wrapEl = `<wp:${wrap}/>`;
  const posHEl = posH
    ? `<wp:positionH relativeFrom="${posH.relativeFrom}">${posH.offset !== undefined ? `<wp:posOffset>${posH.offset}</wp:posOffset>` : `<wp:align>${posH.align ?? 'center'}</wp:align>`}</wp:positionH>`
    : '';
  const posVEl = posV
    ? `<wp:positionV relativeFrom="${posV.relativeFrom}">${posV.offset !== undefined ? `<wp:posOffset>${posV.offset}</wp:posOffset>` : `<wp:align>${posV.align ?? 'top'}</wp:align>`}</wp:positionV>`
    : '';
  const behindAttr = behindDoc !== undefined ? ` behindDoc="${behindDoc}"` : '';

  return `<w:drawing>
    <wp:anchor${behindAttr}>
      <wp:extent cx="${cx}" cy="${cy}"/>
      ${posHEl}
      ${posVEl}
      ${wrapEl}
      <a:graphic>
        <a:graphicData uri="${PIC_NS}">
          <pic:pic/>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;
}

function chartDrawing(): string {
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="2000000" cy="1500000"/>
      <a:graphic>
        <a:graphicData uri="${CHART_NS}">
          <c:chart/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

function smartArtDrawing(): string {
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="3000000" cy="2000000"/>
      <a:graphic>
        <a:graphicData uri="${DGM_NS}">
          <dgm:relIds/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

/** Shape（无文字 txbxContent）—— hasBorder/hasFill 可配置 */
function shapeDrawing(opts: { hasBorder?: boolean; hasFill?: boolean } = {}): string {
  const { hasBorder = true, hasFill = true } = opts;
  const lnEl = hasBorder ? '<a:ln/>' : '<a:ln><a:noFill/></a:ln>';
  const fillEl = hasFill ? '<a:solidFill/>' : '<a:noFill/>';
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="1500000" cy="1000000"/>
      <a:graphic>
        <a:graphicData uri="${WPS_NS}">
          <wps:wsp>
            <wps:spPr>
              ${fillEl}
              ${lnEl}
            </wps:spPr>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

/** TextBox：含 w:txbxContent + 内段，可配置文字与对齐 */
function textBoxDrawing(opts: { text?: string; jc?: string } = {}): string {
  const { text = '说明文字', jc } = opts;
  const jcEl = jc ? `<w:pPr><w:jc w:val="${jc}"/></w:pPr>` : '';
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="2500000" cy="1500000"/>
      <a:graphic>
        <a:graphicData uri="${WPS_NS}">
          <wps:wsp>
            <wps:spPr>
              <a:solidFill/>
              <a:ln/>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p>${jcEl}<w:r><w:t>${text}</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

/**
 * Shape，prstGeom 预设几何可配置（用于线条/连接线识别测试）。
 * 默认 prst="rect"（普通矩形装饰，非线条）。
 */
function shapeWithGeom(opts: { prst?: string; connector?: boolean } = {}): string {
  const { prst = 'rect', connector = false } = opts;
  const cnvPr = connector ? '<wps:cNvCnPr/>' : '<wps:cNvSpPr/>';
  return `<w:drawing>
    <wp:anchor>
      <wp:extent cx="5518150" cy="0"/>
      <wp:wrapNone/>
      <a:graphic>
        <a:graphicData uri="${WPS_NS}">
          <wps:wsp>
            ${cnvPr}
            <wps:spPr>
              <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
              <a:noFill/>
              <a:ln w="9525"><a:solidFill/></a:ln>
            </wps:spPr>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;
}

function groupDrawing(): string {
  return `<w:drawing>
    <wp:inline>
      <wp:extent cx="4000000" cy="3000000"/>
      <a:graphic>
        <a:graphicData uri="${WPG_NS}">
          <wpg:wgp/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

/** 把 drawing XML 内嵌到一个 body 段落里 */
function bodyDocWithDrawings(drawingsXml: string[]): string {
  const paraXml = drawingsXml
    .map((d, i) => `<w:p w14:paraId="${String(i + 1).padStart(8, '0')}"><w:r>${d}</w:r></w:p>`)
    .join('\n');
  return `<w:document${NS_DECL}>
    <w:body>${paraXml}</w:body>
  </w:document>`;
}

describe('buildIndices — drawingIdx 装载期分配', () => {
  test('按文档顺序分配 dr-0001 起的 ID', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([inlinePictureDrawing(), anchorPictureDrawing(), chartDrawing()]),
    );
    const { drawingIdx } = buildIndices(pkg);
    expect(drawingIdx.size).toBe(3);
    const ids = [...drawingIdx.keys()];
    expect(ids).toEqual(['dr-0001', 'dr-0002', 'dr-0003']);
  });

  test('跨 part 顺序：document.xml → header → footer 串行分配', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr${NS_DECL}><w:p w14:paraId="HHHHHHHH"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:hdr>`,
      HDR_CTYPE,
    );
    putXmlPart(
      pkg,
      'word/footer1.xml',
      `<w:ftr${NS_DECL}><w:p w14:paraId="FFFFFFFF"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:ftr>`,
      FTR_CTYPE,
    );
    const { drawingIdx } = buildIndices(pkg);
    expect([...drawingIdx.keys()]).toHaveLength(3);
  });
});

describe('listDrawings — 锚定方式 / kind 识别', () => {
  test('inline picture：anchoringMode=inline / kind=picture / hasPicReference=true / 无 wrap、position', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([inlinePictureDrawing(914400, 685800)]),
    );
    const idx = buildIndices(pkg);
    const list = listDrawings(pkg, idx);
    expect(list).toHaveLength(1);
    const d = list[0]!;
    expect(d.drawingId).toBe('dr-0001');
    expect(d.anchoringMode).toBe('inline');
    expect(d.kind).toBe('picture');
    expect(d.hasPicReference).toBe(true);
    expect(d.extentEMU).toEqual({ cx: 914400, cy: 685800 });
    expect(d.wrap).toBeUndefined();
    expect(d.position).toBeUndefined();
    expect(d.isInHeaderFooter).toBe(false);
    expect(d.partPath).toBe('word/document.xml');
    expect(d.anchorParaId).toBe('00000001');
    expect(d.runIndex).toBe(0);
  });

  test('anchor picture：anchoringMode=anchor / wrap=square / position 解析', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([
        anchorPictureDrawing({
          wrap: 'wrapSquare',
          posH: { relativeFrom: 'page', offset: 1000000 },
          posV: { relativeFrom: 'margin', align: 'center' },
        }),
      ]),
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    const d = list[0]!;
    expect(d.anchoringMode).toBe('anchor');
    expect(d.wrap).toBe('square');
    expect(d.position?.horizontal).toEqual({ relativeFrom: 'page', posOffset: 1000000 });
    expect(d.position?.vertical).toEqual({ relativeFrom: 'margin', align: 'center' });
  });

  test('anchor + wrapNone + behindDoc=1 → wrap=behind；behindDoc=0 → inFront', () => {
    const pkg1 = createPackage();
    putXmlPart(
      pkg1,
      'word/document.xml',
      bodyDocWithDrawings([anchorPictureDrawing({ wrap: 'wrapNone', behindDoc: '1' })]),
    );
    expect(listDrawings(pkg1, buildIndices(pkg1))[0]!.wrap).toBe('behind');

    const pkg2 = createPackage();
    putXmlPart(
      pkg2,
      'word/document.xml',
      bodyDocWithDrawings([anchorPictureDrawing({ wrap: 'wrapNone', behindDoc: '0' })]),
    );
    expect(listDrawings(pkg2, buildIndices(pkg2))[0]!.wrap).toBe('inFront');

    const pkg3 = createPackage();
    putXmlPart(
      pkg3,
      'word/document.xml',
      bodyDocWithDrawings([anchorPictureDrawing({ wrap: 'wrapTopAndBottom' })]),
    );
    expect(listDrawings(pkg3, buildIndices(pkg3))[0]!.wrap).toBe('topAndBottom');
  });

  test('chart kind 识别 + hasPicReference=false', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([chartDrawing()]));
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('chart');
    expect(d.hasPicReference).toBe(false);
    expect(d.extentEMU).toEqual({ cx: 2000000, cy: 1500000 });
  });

  test('smartArt kind 识别（uri = .../diagram）', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([smartArtDrawing()]));
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.kind).toBe('smartArt');
  });

  test('group kind 识别（uri = .../wordprocessingGroup）', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([groupDrawing()]));
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.kind).toBe('group');
  });
});

describe('listDrawings — Shape / TextBox 形态信号', () => {
  test('shape（无 txbxContent）→ kind=shape，hasBorder/hasFill 反映 a:ln / 填充元素', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([
        shapeDrawing({ hasBorder: true, hasFill: true }),
        shapeDrawing({ hasBorder: false, hasFill: false }),
      ]),
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list).toHaveLength(2);
    expect(list[0]!.kind).toBe('shape');
    expect(list[0]!.hasBorder).toBe(true);
    expect(list[0]!.hasFill).toBe(true);
    // textBox 专属字段不应出现在 shape 上
    expect(list[0]!.hasInnerText).toBeUndefined();
    expect(list[0]!.innerTextAlign).toBeUndefined();

    expect(list[1]!.hasBorder).toBe(false);
    expect(list[1]!.hasFill).toBe(false);
  });

  test('textBox（含 txbxContent）→ kind=textBox，hasInnerText + innerTextAlign', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([
        textBoxDrawing({ text: '说明文字', jc: 'center' }),
        textBoxDrawing({ text: '', jc: undefined }), // 空文字
      ]),
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list[0]!.kind).toBe('textBox');
    expect(list[0]!.hasInnerText).toBe(true);
    expect(list[0]!.innerText).toBe('说明文字');
    expect(list[0]!.innerTextAlign).toBe('center');

    expect(list[1]!.kind).toBe('textBox');
    expect(list[1]!.hasInnerText).toBe(false);
    expect(list[1]!.innerText).toBeUndefined(); // 空文字 → innerText 不设
    expect(list[1]!.innerTextAlign).toBeUndefined();
  });

  test('textBox innerText：多 run / 多段按文档顺序拼接 + trim（供红头字段化读原文）', () => {
    const pkg = createPackage();
    // 机关名（contiguous）：单段多 run，原文须按顺序拼成「江临市滨河区人民政府」
    const orgName = `<w:drawing><wp:inline><wp:extent cx="2000000" cy="800000"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:spPr/>
        <wps:txbx><w:txbxContent>
          <w:p><w:r><w:t>江临市</w:t></w:r><w:r><w:t>滨河区</w:t></w:r><w:r><w:t>人民政府</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx>
      </wps:wsp></a:graphicData></a:graphic>
    </wp:inline></w:drawing>`;
    // 文种（公文字间距：每字之间空格）：「专 题 会 议 纪 要」原样保留，trim 仅去首尾空白
    const docType = `<w:drawing><wp:inline><wp:extent cx="2000000" cy="800000"/>
      <a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:spPr/>
        <wps:txbx><w:txbxContent>
          <w:p><w:r><w:t xml:space="preserve">  专 题 会 议 纪 要  </w:t></w:r></w:p>
        </w:txbxContent></wps:txbx>
      </wps:wsp></a:graphicData></a:graphic>
    </wp:inline></w:drawing>`;
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([orgName, docType]));
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list[0]!.innerText).toBe('江临市滨河区人民政府');
    // 首尾空白 trim 掉，字间空格保留（原文）
    expect(list[1]!.innerText).toBe('专 题 会 议 纪 要');
  });

  test('非 shape/textBox 的 hasBorder/hasFill/hasInnerText/innerText 都为 undefined', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([inlinePictureDrawing(), chartDrawing()]),
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    for (const d of list) {
      expect(d.hasBorder).toBeUndefined();
      expect(d.hasFill).toBeUndefined();
      expect(d.hasInnerText).toBeUndefined();
      expect(d.innerText).toBeUndefined();
      expect(d.innerTextAlign).toBeUndefined();
    }
  });
});

describe('listDrawings — 线条/连接线识别（isLineOrConnector）', () => {
  test('shape prstGeom prst="line" → isLineOrConnector=true（母版版记横线）', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([shapeWithGeom({ prst: 'line', connector: true })]),
    );
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('shape');
    expect(d.isLineOrConnector).toBe(true);
  });

  test('shape prst="straightConnector1" → isLineOrConnector=true', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([shapeWithGeom({ prst: 'straightConnector1', connector: true })]),
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.isLineOrConnector).toBe(true);
  });

  test.each(['bentConnector3', 'curvedConnector2', 'bentConnector5', 'curvedConnector5'])(
    'shape prst="%s"（折线/曲线连接符）→ isLineOrConnector=true',
    (prst) => {
      const pkg = createPackage();
      putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([shapeWithGeom({ prst })]));
      expect(listDrawings(pkg, buildIndices(pkg))[0]!.isLineOrConnector).toBe(true);
    },
  );

  test('shape 含 wps:cNvCnPr 但 prst 非线条预设 → 仍 isLineOrConnector=true（connector 信号兜底）', () => {
    const pkg = createPackage();
    // 极端：prst 是未知值，但 cNvCnPr 在 → 仍判线条
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([shapeWithGeom({ prst: 'someUnknown', connector: true })]),
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.isLineOrConnector).toBe(true);
  });

  test('shape prst="rect" + wps:cNvSpPr（普通矩形装饰）→ isLineOrConnector=false', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([shapeWithGeom({ prst: 'rect', connector: false })]),
    );
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('shape');
    expect(d.isLineOrConnector).toBe(false);
  });

  test('shape prst="ellipse"（椭圆装饰）→ isLineOrConnector=false', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([shapeWithGeom({ prst: 'ellipse' })]));
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.isLineOrConnector).toBe(false);
  });

  test('VML v:line 兜底：container 内嵌 v:line → isLineOrConnector=true', () => {
    const pkg = createPackage();
    // 构造 shape kind（graphicData uri = WPS）但 spPr 内无 prstGeom，靠内嵌 v:line 命中信号 3
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL} xmlns:v="urn:schemas-microsoft-com:vml">
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:anchor>
              <wp:extent cx="100" cy="0"/>
              <wp:wrapNone/>
              <a:graphic>
                <a:graphicData uri="${WPS_NS}">
                  <wps:wsp>
                    <wps:spPr><v:line/></wps:spPr>
                  </wps:wsp>
                </a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('shape');
    expect(d.isLineOrConnector).toBe(true);
  });

  test('非 shape kind（picture/chart/textBox）→ isLineOrConnector 一律 undefined', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([
        inlinePictureDrawing(),
        chartDrawing(),
        textBoxDrawing({ text: '说明' }),
      ]),
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    for (const d of list) {
      expect(d.isLineOrConnector).toBeUndefined();
    }
  });

  test('普通 shape（shapeDrawing，无 prstGeom / 无 cNvCnPr）→ isLineOrConnector=false', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([shapeDrawing()]));
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('shape');
    expect(d.isLineOrConnector).toBe(false);
  });
});

describe('listDrawings — 段落引用与 partPath', () => {
  test('anchorParaId / runIndex 指向承载 drawing 的 w:r 与 w:p', () => {
    const pkg = createPackage();
    // 一个段落里两个 drawing：前一个 inline 图片，后一个 chart
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA">
            <w:r>${inlinePictureDrawing()}</w:r>
            <w:r><w:t>正文</w:t></w:r>
            <w:r>${chartDrawing()}</w:r>
          </w:p>
        </w:body>
      </w:document>`,
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list).toHaveLength(2);
    expect(list[0]!.anchorParaId).toBe('AAAAAAAA');
    expect(list[0]!.runIndex).toBe(0);
    expect(list[1]!.anchorParaId).toBe('AAAAAAAA');
    expect(list[1]!.runIndex).toBe(2);
  });
});

describe('listDrawings — scope 行为 / 页眉页脚识别', () => {
  test("scope: 'body'（默认）不返回页眉页脚内 drawing", () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr${NS_DECL}><w:p w14:paraId="HHHHHHHH"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:hdr>`,
      HDR_CTYPE,
    );
    const idx = buildIndices(pkg);
    const bodyList = listDrawings(pkg, idx);
    expect(bodyList).toHaveLength(1);
    expect(bodyList[0]!.partPath).toBe('word/document.xml');
    expect(bodyList[0]!.isInHeaderFooter).toBe(false);
  });

  // CR 跟进：scope:'all' 在文档中承诺包含 comments 内 drawing；
  // 旧实现 buildDrawingIndex 漏过 .comments+xml 导致永远枚不到。
  test("scope: 'body' 不收 comments part 内的 drawing", () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/comments.xml',
      `<w:comments${NS_DECL}>
        <w:comment w:id="0" w:author="t" w:date="2026-05-20T00:00:00Z">
          <w:p w14:paraId="CCCCCCCC"><w:r>${inlinePictureDrawing()}</w:r></w:p>
        </w:comment>
      </w:comments>`,
      COMMENTS_CTYPE,
    );
    const idx = buildIndices(pkg);
    // drawingIdx 应收两个（含 comments part 内的），但 scope='body' 不返回
    expect(idx.drawingIdx.size).toBe(2);
    const bodyList = listDrawings(pkg, idx);
    expect(bodyList).toHaveLength(1);
    expect(bodyList[0]!.partPath).toBe('word/document.xml');
  });

  test("scope: 'all' 收 comments part 内的 drawing，partPath=word/comments.xml", () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/comments.xml',
      `<w:comments${NS_DECL}>
        <w:comment w:id="0" w:author="t" w:date="2026-05-20T00:00:00Z">
          <w:p w14:paraId="CCCCCCCC"><w:r>${inlinePictureDrawing()}</w:r></w:p>
        </w:comment>
      </w:comments>`,
      COMMENTS_CTYPE,
    );
    const list = listDrawings(pkg, buildIndices(pkg), { scope: 'all' });
    expect(list).toHaveLength(2);
    const inComment = list.find((d) => d.partPath === 'word/comments.xml');
    expect(inComment).toBeDefined();
    expect(inComment!.kind).toBe('picture');
    expect(inComment!.anchorParaId).toBe('CCCCCCCC');
    expect(inComment!.isInHeaderFooter).toBe(false);
  });

  test("scope: 'all' 返回页眉页脚 drawing，isInHeaderFooter=true", () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr${NS_DECL}><w:p w14:paraId="HHHHHHHH"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:hdr>`,
      HDR_CTYPE,
    );
    putXmlPart(
      pkg,
      'word/footer1.xml',
      `<w:ftr${NS_DECL}><w:p w14:paraId="FFFFFFFF"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:ftr>`,
      FTR_CTYPE,
    );
    const list = listDrawings(pkg, buildIndices(pkg), { scope: 'all' });
    expect(list).toHaveLength(3);
    const inHdrFtr = list.filter((d) => d.isInHeaderFooter);
    expect(inHdrFtr).toHaveLength(2);
    // 主文档的不应标 isInHeaderFooter
    const body = list.find((d) => d.partPath === 'word/document.xml')!;
    expect(body.isInHeaderFooter).toBe(false);
  });
});

describe('getDrawingById — O(1) 取单个', () => {
  test('存在的 drawingId → 返回完整 DrawingInfo', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([inlinePictureDrawing(), chartDrawing()]),
    );
    const idx = buildIndices(pkg);
    const d2 = getDrawingById(pkg, idx, 'dr-0002');
    expect(d2).not.toBeNull();
    expect(d2!.kind).toBe('chart');
    expect(d2!.drawingId).toBe('dr-0002');
  });

  test('不存在的 drawingId → 返回 null', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    const idx = buildIndices(pkg);
    expect(getDrawingById(pkg, idx, 'dr-9999')).toBeNull();
  });

  // CR BUG-1：getDrawingById 旧实现返回 index:0（list-tables.ts copy-paste），
  // 与 listDrawings 的 1-based 序号不一致；改为返回全局 drawingIdx 1-based 序号。
  test('index 字段返回全局 drawingIdx 中的 1-based 真实序号（非 0 哨兵）', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([
        inlinePictureDrawing(),
        chartDrawing(),
        smartArtDrawing(),
        groupDrawing(),
      ]),
    );
    const idx = buildIndices(pkg);

    expect(getDrawingById(pkg, idx, 'dr-0001')!.index).toBe(1);
    expect(getDrawingById(pkg, idx, 'dr-0002')!.index).toBe(2);
    expect(getDrawingById(pkg, idx, 'dr-0003')!.index).toBe(3);
    expect(getDrawingById(pkg, idx, 'dr-0004')!.index).toBe(4);
  });

  test('getDrawingById 跨 part 的 index 与 listDrawings({scope:"all"}) 一致', () => {
    const pkg = createPackage();
    putXmlPart(pkg, 'word/document.xml', bodyDocWithDrawings([inlinePictureDrawing()]));
    putXmlPart(
      pkg,
      'word/header1.xml',
      `<w:hdr${NS_DECL}><w:p w14:paraId="HHHHHHHH"><w:r>${inlinePictureDrawing()}</w:r></w:p></w:hdr>`,
      HDR_CTYPE,
    );
    const idx = buildIndices(pkg);
    const all = listDrawings(pkg, idx, { scope: 'all' });
    for (const d of all) {
      const got = getDrawingById(pkg, idx, d.drawingId)!;
      expect(got.index).toBe(d.index);
    }
  });
});

describe('listDrawings — CR 修复：on/off 属性解析（ECMA-376 ST_OnOff）', () => {
  // 旧实现只匹配 behindDoc="1" / "0"，会把 true/false/on/off 误判
  test.each([
    ['1', 'behind'],
    ['true', 'behind'],
    ['True', 'behind'],
    ['on', 'behind'],
    ['ON', 'behind'],
    ['0', 'inFront'],
    ['false', 'inFront'],
    ['False', 'inFront'],
    ['off', 'inFront'],
  ] as const)('behindDoc=%s 应解析为 wrap=%s', (behindVal, expectedWrap) => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:anchor behindDoc="${behindVal}">
              <wp:extent cx="100" cy="100"/>
              <wp:wrapNone/>
              <a:graphic>
                <a:graphicData uri="${PIC_NS}"><pic:pic/></a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.wrap).toBe(expectedWrap);
  });

  test('behindDoc 缺失 + wrapNone → wrap=none（不假设默认值）', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:anchor>
              <wp:extent cx="100" cy="100"/>
              <wp:wrapNone/>
              <a:graphic>
                <a:graphicData uri="${PIC_NS}"><pic:pic/></a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.wrap).toBe('none');
  });

  test('behindDoc 字面无效（如 "yes"）→ 视作未声明，wrap=none', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:anchor behindDoc="yes">
              <wp:extent cx="100" cy="100"/>
              <wp:wrapNone/>
              <a:graphic>
                <a:graphicData uri="${PIC_NS}"><pic:pic/></a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.wrap).toBe('none');
  });
});

describe('listDrawings — CR BUG-2 修复：a:ln w="0" 零宽边框', () => {
  test('<a:ln w="0"/>（零宽线）→ hasBorder=false', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:inline>
              <wp:extent cx="100" cy="100"/>
              <a:graphic>
                <a:graphicData uri="${WPS_NS}">
                  <wps:wsp>
                    <wps:spPr>
                      <a:solidFill/>
                      <a:ln w="0"/>
                    </wps:spPr>
                  </wps:wsp>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const d = listDrawings(pkg, buildIndices(pkg))[0]!;
    expect(d.kind).toBe('shape');
    expect(d.hasBorder).toBe(false);
    expect(d.hasFill).toBe(true);
  });

  test('<a:ln w="12700"/>（标准 1pt 宽）→ hasBorder=true', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:inline>
              <wp:extent cx="100" cy="100"/>
              <a:graphic>
                <a:graphicData uri="${WPS_NS}">
                  <wps:wsp>
                    <wps:spPr>
                      <a:noFill/>
                      <a:ln w="12700"/>
                    </wps:spPr>
                  </wps:wsp>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.hasBorder).toBe(true);
  });

  test('<a:ln w="0"><a:noFill/></a:ln>（双重声明无边框）→ hasBorder=false', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:inline>
              <wp:extent cx="100" cy="100"/>
              <a:graphic>
                <a:graphicData uri="${WPS_NS}">
                  <wps:wsp>
                    <wps:spPr>
                      <a:ln w="0"><a:noFill/></a:ln>
                    </wps:spPr>
                  </wps:wsp>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    expect(listDrawings(pkg, buildIndices(pkg))[0]!.hasBorder).toBe(false);
  });
});

describe('listDrawings — CR BUG-3 修复：mc:AlternateContent Fallback 过滤', () => {
  test('mc:Fallback 内的 w:drawing 不进 drawingIdx，避免 mc:Choice / Fallback 双计', () => {
    const pkg = createPackage();
    // Word 2010+ 把 SmartArt 包在 mc:AlternateContent 里：
    //   - mc:Choice 用 w14 / dgm14 渲染（含真实 drawing）
    //   - mc:Fallback 给老版本用（不识别 Choice 时用），也可能包含 drawing
    // 旧实现深遍历会把两个 drawing 都收，导致双计。
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wpg">
                ${chartDrawing()}
              </mc:Choice>
              <mc:Fallback>
                ${inlinePictureDrawing()}
              </mc:Fallback>
            </mc:AlternateContent>
          </w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const idx = buildIndices(pkg);
    expect(idx.drawingIdx.size).toBe(1);
    const list = listDrawings(pkg, idx);
    expect(list).toHaveLength(1);
    // 留下来的应该是 mc:Choice 里的 chart，不是 Fallback 里的 picture
    expect(list[0]!.kind).toBe('chart');
  });

  test('普通 drawing（不在 Fallback 内）正常计入', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      bodyDocWithDrawings([inlinePictureDrawing(), chartDrawing()]),
    );
    expect(buildIndices(pkg).drawingIdx.size).toBe(2);
  });
});

describe('listDrawings — txbxContent 内 drawing 的 anchorParaId 主动留空（设计）', () => {
  // 文本框内 drawing 不应把外层段落 paraId 当作承载段（避免删段落操作误伤）
  test('TextBox 内的图片，anchorParaId/runIndex 都为 undefined', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="OUTERPID"><w:r><w:drawing>
            <wp:inline>
              <wp:extent cx="100" cy="100"/>
              <a:graphic>
                <a:graphicData uri="${WPS_NS}">
                  <wps:wsp>
                    <wps:spPr><a:ln/></wps:spPr>
                    <wps:txbx>
                      <w:txbxContent>
                        <w:p w14:paraId="INNERPID"><w:r>${inlinePictureDrawing()}</w:r></w:p>
                      </w:txbxContent>
                    </wps:txbx>
                  </wps:wsp>
                </a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list).toHaveLength(2);
    // 外层 textBox 自己挂在 OUTERPID 上
    const outerTextBox = list.find((d) => d.kind === 'textBox')!;
    expect(outerTextBox.anchorParaId).toBe('OUTERPID');
    // 文本框内的 picture：anchorParaId 主动留空（设计：避免误指外层段落）
    const innerPic = list.find((d) => d.kind === 'picture')!;
    expect(innerPic.anchorParaId).toBeUndefined();
    expect(innerPic.runIndex).toBeUndefined();
  });
});

describe('listDrawings — 兜底场景', () => {
  test('extent 缺失 → extentEMU = { cx: 0, cy: 0 }', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:inline>
              <a:graphic>
                <a:graphicData uri="${PIC_NS}"><pic:pic/></a:graphicData>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list[0]!.extentEMU).toEqual({ cx: 0, cy: 0 });
  });

  test('graphicData uri 不在已知列表中 → kind=other', () => {
    const pkg = createPackage();
    putXmlPart(
      pkg,
      'word/document.xml',
      `<w:document${NS_DECL}>
        <w:body>
          <w:p w14:paraId="AAAAAAAA"><w:r><w:drawing>
            <wp:inline>
              <wp:extent cx="100" cy="100"/>
              <a:graphic>
                <a:graphicData uri="http://example.com/unknown"/>
              </a:graphic>
            </wp:inline>
          </w:drawing></w:r></w:p>
        </w:body>
      </w:document>`,
    );
    const list = listDrawings(pkg, buildIndices(pkg));
    expect(list[0]!.kind).toBe('other');
  });
});
