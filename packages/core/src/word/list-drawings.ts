/**
 * Drawing 对象枚举（#189）
 *
 * 为骨架化 M3 阶段（图片分类/Shape 分类/占位图替换）提供 drawing 数据。
 *
 * 边界（实施前澄清，见 #189 评论）：
 * - **本模块只做"枚举 + 形态信号"**，不下 A / B / C 业务分类结论
 * - C 类细分（水印 / 装饰线 / 背景图）需语义判断，属于 #191
 * - 分组键计算、代表样本选取属于 #191 / #192
 * - 占位图替换属于 #190
 *
 * 设计原则（对齐 #185 list-tables.ts）：
 * - 只走 getElementsByTagNameNS / 直接子节点遍历，禁用 XPath（性能约束）
 * - drawingId 由 buildIndices 装载期分配（'dr-0001' 起），存入 drawingIdx
 * - 默认枚举主文档 body（scope: 'body'）
 * - 提供 isInHeaderFooter 字段，让 #191 直接识别 C 类页眉页脚 logo（无需自己查 part 路径）
 */

import type { Package } from '../pack/index.js';
import {
  W_NS,
  WP_NS,
  A_NS,
  PIC_NS,
  CHART_NS,
  DGM_NS,
  WPS_NS,
  WPG_NS,
  WPC_NS,
  W14_NS,
  MC_NS,
  VML_NS,
} from './namespaces.js';
import type { PackageIndices } from './indices.js';
import { getRunText, type RunTextOptions } from './runs.js';
import { resolveScope, type Story } from './stories.js';

// ─────────── 公开类型 ───────────

/** 锚定方式：inline（行内嵌入）/ anchor（浮动定位）*/
export type DrawingAnchoringMode = 'inline' | 'anchor';

/**
 * Drawing 内容形态（**OOXML 子树类型**，非 A/B/C 业务分类）。
 *
 * 识别优先级：a:graphicData @uri → 命中预设 URI 之一时按表分发；都不命中时落 'other'。
 * 'textBox' 是 'shape' 的特化（wps:wsp 内含 w:txbxContent），单独抽出便于 #192 处理。
 *
 * | kind     | URI 或子元素信号                                                          | 典型场景 |
 * |----------|---------------------------------------------------------------------------|----------|
 * | picture  | uri = .../picture                                                          | 嵌入/浮动图片 |
 * | chart    | uri = .../chart                                                            | Word 原生图表 |
 * | smartArt | uri = .../diagram                                                          | SmartArt（1.0/2010+）|
 * | shape    | uri = .../wordprocessingShape，且 **无** w:txbxContent                     | 形状（无文字） |
 * | textBox  | uri = .../wordprocessingShape，且 **有** w:txbxContent                     | 文本框 / 标注 |
 * | group    | uri = .../wordprocessingGroup                                              | 形状组合 |
 * | canvas   | uri = .../wordprocessingCanvas                                             | 绘图画布 |
 * | other    | 都不命中                                                                   | 未来扩展占位 |
 */
export type DrawingKind =
  | 'picture'
  | 'chart'
  | 'smartArt'
  | 'shape'
  | 'textBox'
  | 'group'
  | 'canvas'
  | 'other';

export type TextBoxBlockInfo =
  | { type: 'paragraph'; paraId: string; text: string }
  | { type: 'drawing'; drawingId: string; paraId?: string }
  | { type: 'table'; tableId: string };

/**
 * 文字环绕方式（仅 anchor 有）。
 *
 * - none：wp:wrapNone（不环绕；配合 behindDoc 形成"衬于文字下方/上方"）
 * - square：wp:wrapSquare（四周型 / 矩形环绕）
 * - tight：wp:wrapTight（紧密型）
 * - through：wp:wrapThrough（穿越型）
 * - topAndBottom：wp:wrapTopAndBottom（上下型）
 * - behind：wp:wrapNone + behindDoc="1"（衬于文字下方）
 * - inFront：wp:wrapNone + behindDoc="0"（浮于文字上方）
 */
export type DrawingWrapType =
  | 'none'
  | 'square'
  | 'tight'
  | 'through'
  | 'topAndBottom'
  | 'behind'
  | 'inFront';

/** 浮动定位的单方向描述（原值透传，由 #191 自行二次解释）*/
export interface DrawingPositionAxis {
  /** wp:positionH/V @relativeFrom 原值（如 'page' / 'margin' / 'column' / 'character' / 'paragraph' / 'line'）*/
  relativeFrom?: string;
  /** wp:posOffset 文本内容（EMU 整数）。和 align 二选一。*/
  posOffset?: number;
  /** wp:align 文本内容（'left' / 'center' / 'right' / 'top' / 'bottom' / 'inside' / 'outside'）。和 posOffset 二选一。*/
  align?: string;
}

/** 浮动位置（仅 anchor 有）*/
export interface DrawingPosition {
  horizontal?: DrawingPositionAxis;
  vertical?: DrawingPositionAxis;
}

export interface DrawingInfo {
  /** 装载期分配的 ID（'dr-NNNN'，从 0001 起），供 #190 / #191 / #192 引用 */
  drawingId: string;
  /** 所在 part 路径（'word/document.xml' / 'word/header1.xml' 等）*/
  partPath: string;
  /** 枚举结果顺序（1-based，与文档顺序一致）*/
  index: number;
  /** 锚定方式：inline 行内嵌入 / anchor 浮动定位 */
  anchoringMode: DrawingAnchoringMode;
  /** 内容形态（OOXML 层级，非业务分类）*/
  kind: DrawingKind;
  /**
   * 显示尺寸（EMU；1 inch = 914400 EMU）。
   * 来源：wp:extent @cx @cy。
   * 未声明或解析失败时为 { cx: 0, cy: 0 }（保持字段必存在，避免下游需要判 undefined）。
   */
  extentEMU: { cx: number; cy: number };
  /** 文字环绕方式（仅 anchor，inline 不含此字段）*/
  wrap?: DrawingWrapType;
  /** 浮动位置（仅 anchor，inline 不含此字段）*/
  position?: DrawingPosition;
  /**
   * 承载此 drawing 的段落 paraId（w:p @w14:paraId）。
   *
   * 多数情况下，drawing 挂在 w:r 下、该 w:r 属于某个 w:p，本字段就是这个 w:p 的 paraId。
   *
   * **关键例外（CR 设计说明）**：嵌套在 **TextBox 内**（即位于 `w:txbxContent` 子树）
   * 的 drawing，本字段为 `undefined`。这是**主动设计**：textBox 内段落与外层段落是两套
   * 独立的段落流，把外层段落 paraId 挂给文本框内 drawing 会让"删段落 → 顺带删图"等
   * 操作误伤外层段落。下游需要文本框内图片的段落引用时，应：
   *  1. 用 `runIndex === undefined && anchorParaId === undefined` 识别该情形
   *  2. 自行向上找最近的 `w:txbxContent`，再用其外层 anchor 段落定位
   *
   * 缺 paraId 时也为 undefined（合法但少见；由 ensureParaIds 在装载期补齐，但本字段
   * 不依赖这一保证）。
   */
  anchorParaId?: string;
  /**
   * 在所在段落里的 run 序号（0-based，承载 drawing 的 w:r 在父 w:p 直接子 w:r 中的位置）。
   * 同段多 drawing 时用于区分；用于题注绑定（#193）/ 占位图替换（#190）的精确引用。
   */
  runIndex?: number;
  /**
   * 是否位于页眉页脚 part（contentType 以 .header+xml / .footer+xml 结尾）。
   *
   * 这是**形态信号**，不是 C 类结论。#191 在做 A/B/C 业务分类时：
   * - true + kind = picture → 大概率 C 类页眉页脚 logo，由 .dot 原生保留
   * - 其余情况由 #191 自行判断
   */
  isInHeaderFooter: boolean;
  /**
   * 形态信号：是否含 pic:pic 子树（A 类正向信号）。
   * - kind = 'picture' 时必为 true
   * - 理论上 chart/shape 也可能套 pic:pic，此时 kind 不变但本字段为 true
   */
  hasPicReference: boolean;
  /**
   * Shape 专属：是否为「线条 / 连接线（line / connector）」几何。仅 kind = 'shape' 时可能为 true，
   * 其余 kind（picture/chart/.../textBox/group/canvas/other）一律 undefined。
   *
   * **形态信号，非业务结论**。下游（skeleton #192 shape-role）用它把版记横线一类「纯几何线条」
   * 与矩形/椭圆等装饰噪声区分开：线条原样保留，其余几何仍按现状裁掉。
   *
   * 识别（命中任一即 true，从强到弱）：
   *  1. DrawingML 预设几何：`<a:prstGeom prst="...">` 的 prst 命中线条/连接线预设
   *     （line / straightConnector1 / bentConnectorN / curvedConnectorN，N=2..5；见 LINE_CONNECTOR_PRESETS）
   *  2. 连接线非可视属性：wps:wsp 内含 `<wps:cNvCnPr>`（connector non-visual props；
   *     普通 shape 用 wps:cNvSpPr，连接线用 wps:cNvCnPr）
   *  3. VML 回退：子树含 `<v:line>`（mc:Fallback 内的旧版降级线条）
   *
   * 母版实测（生成的公文__会议纪要.docx）：版记 3 根横线均 prst="line" + wps:cNvCnPr + v:line 三信号齐全；
   * 红头艺术字文本框是 prst="rect" + wps:cNvSpPr，不命中。
   */
  isLineOrConnector?: boolean;
  /**
   * Shape / TextBox 专属：是否有边框（a:spPr/a:ln 存在且未声明 a:noFill）。
   * 非 shape/textBox 一律 undefined。
   */
  hasBorder?: boolean;
  /**
   * Shape / TextBox 专属：是否有填充（a:spPr 下含 a:solidFill / a:gradFill / a:blipFill / a:pattFill）。
   * 显式 a:noFill 视作无填充。非 shape/textBox 一律 undefined。
   */
  hasFill?: boolean;
  /**
   * TextBox 专属：是否含实质内部文字（w:txbxContent 内 w:t 文本拼接后非空）。
   * 非 textBox 一律 undefined。
   */
  hasInnerText?: boolean;
  /**
   * TextBox 专属：内部文字原文（w:txbxContent 内全部 w:t 文本按文档顺序拼接后 trim）。
   *
   * 供下游（skeleton #192 / typeset 投影）读取红头艺术字等内容型文本框的原文做语义命名/字段化
   * （如「江临市滨河区人民政府」「专题会议纪要」）。
   *
   * 语义：
   *  - 仅 kind = 'textBox' 时可能有值；非 textBox 一律 undefined。
   *  - trim 后为空（无文字 / 仅空白）→ undefined（与 hasInnerText=false 对应）。
   *  - 多段 / 多 run 文本框按 w:t 出现顺序直接拼接，**不插入换行/空格**（保持与 hasInnerText
   *    同源的拼接口径）；下游需要分段语义时应自行用 paragraphs API 处理，本字段只给原文快照。
   */
  innerText?: string;
  /**
   * TextBox 专属：文本框内容流里的段落 paraId，按文档顺序排列。
   *
   * 这些 paraId 指向 `w:txbxContent` 内部段落。调用方可通过
   * `doc.paragraphs.get(paraId)` 或 `TextBoxDrawing.getParagraphs()` 继续进行段落/文字级操作。
   * 非 textBox 一律 undefined。
   */
  innerParagraphIds?: string[];
  /**
   * TextBox 专属：文本框内容流的直接内容块，按 OOXML child-node 顺序排列。
   *
   * 目前包含 direct paragraph、paragraph 内 direct drawing，以及 direct table。
   * 非 textBox 一律 undefined。
   */
  innerBlocks?: TextBoxBlockInfo[];
  /**
   * TextBox 专属：首个内部段落的对齐（w:txbxContent 第一段 w:pPr/w:jc @val）。
   *
   * 注：文本框内可能多段且对齐不同，#192 需自行决定如何取代表值；本字段只提供首段值便于快速分组。
   * 非 textBox 或无 jc 设置一律 undefined。
   */
  innerTextAlign?: string;
}

export interface ListDrawingsOptions {
  /**
   * 枚举范围（默认 'body'）。
   * - 'body'：仅主文档 body（不含页眉页脚 / 注释 / 脚注尾注）
   * - 'all'：所有 stories（body + headers + footers + footnotes + endnotes + comments）
   *
   * 注意：#189 的 isInHeaderFooter 字段只在 'all' 模式下才可能为 true；
   * 'body' 模式下页眉页脚 drawing 不会被枚出。
   */
  scope?: 'body' | 'all';
  /**
   * 启用结构空白保真读（task-20260623-core-tab-faithful-read）。
   * `true` 时 TextBox 的 `innerText` 含 `\t`（制表符）和 `\n`（软换行）——
   * 是 `appendSegmentsToRun` 的精确逆运算。
   * 默认 `false`：只拼 `<w:t>`（与历史行为一致）。
   */
  structuralWhitespace?: boolean;
}

// ─────────── 内部工具 ───────────

/** 把 HTMLCollection-ish 枚举结果快照成数组（避免 live 行为）*/
function snapshot<T extends Element>(col: { length: number } & Record<number, T>): T[] {
  const out: T[] = [];
  for (let i = 0; i < col.length; i++) {
    const el = col[i];
    if (el) out.push(el);
  }
  return out;
}

/**
 * OOXML on/off 属性解析。
 *
 * 同一布尔语义在 OOXML 文档中可序列化为多种字面形态（ECMA-376 §17.18.93 ST_OnOff）：
 *   - true:  "1" / "true" / "on"
 *   - false: "0" / "false" / "off"
 *
 * 缺值（null）/ 未识别字面 → undefined（区分"未声明"与"显式 false"，由调用方决策默认值）。
 *
 * 触发场景：`<wp:anchor behindDoc="true">`（部分模板/第三方工具用 true 而非 1）
 * —— 旧实现只匹配 "1" / "0"，会把 behindDoc="true" 误判为 inFront。
 */
function parseOnOff(value: string | null): boolean | undefined {
  if (value === null) return undefined;
  const v = value.trim().toLowerCase();
  if (v === '1' || v === 'true' || v === 'on') return true;
  if (v === '0' || v === 'false' || v === 'off') return false;
  return undefined;
}

/**
 * 判断 element 是否嵌套在 mc:AlternateContent / mc:Fallback 之中。
 *
 * 触发场景：第三方工具或某些 Word 旧版会把 SmartArt / Chart 的"降级 fallback"
 * 也用 `<w:drawing>` 包装放进 `<mc:Fallback>`，再被 mc:Choice 里的真版本"覆盖"。
 * 此时同一对象会被深遍历枚两次（Choice + Fallback），导致 drawingIdx 双计。
 *
 * 防御方式：装载期主动跳过 Fallback 内部的 drawing —— mc:Choice 已经是规范的
 * "首选呈现"，Fallback 只在 reader 不识别 Choice 时使用，不应进枚举结果。
 *
 * 实测：Word / WPS 生成的真实文档 Fallback 用 VML `<w:pict>`，不命中本守卫；
 * 加上这一层只为防御第三方工具。
 */
function hasFallbackAncestor(el: Element): boolean {
  let p = el.parentElement;
  while (p) {
    if (p.namespaceURI === MC_NS && p.localName === 'Fallback') return true;
    p = p.parentElement;
  }
  return false;
}

/** 取 element 直接子中首个匹配 (ns, local) 的，未找到返回 null */
function firstDirectChild(parent: Element, ns: string, local: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (
      child &&
      child.nodeType === 1 /* ELEMENT_NODE */ &&
      child.namespaceURI === ns &&
      child.localName === local
    ) {
      return child;
    }
  }
  return null;
}

/** 取 element 子孙中首个匹配 (ns, local) 的，未找到返回 null */
function firstDescendant(parent: Element, ns: string, local: string): Element | null {
  const col = parent.getElementsByTagNameNS(ns, local);
  return col.length > 0 ? col[0]! : null;
}

/** 解析 number-like attribute；缺/非数为 undefined */
function readNumAttr(el: Element, ns: string | null, local: string): number | undefined {
  const v = ns === null ? el.getAttribute(local) : el.getAttributeNS(ns, local);
  if (v === null) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** 取元素直接文本（不递归）*/
function textContentOfDirect(el: Element): string {
  let s = '';
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes[i] as Node & { nodeValue?: string };
    if (child.nodeType === 3 /* TEXT_NODE */) {
      s += child.nodeValue ?? '';
    }
  }
  return s;
}

/** 取 wp:positionH / wp:positionV 元素的 axis 描述 */
function readPositionAxis(axisEl: Element): DrawingPositionAxis {
  const out: DrawingPositionAxis = {};
  const relativeFrom = axisEl.getAttribute('relativeFrom');
  if (relativeFrom !== null) out.relativeFrom = relativeFrom;

  const posOffsetEl = firstDirectChild(axisEl, WP_NS, 'posOffset');
  if (posOffsetEl) {
    const txt = textContentOfDirect(posOffsetEl).trim();
    const n = Number(txt);
    if (Number.isFinite(n)) out.posOffset = n;
  } else {
    const alignEl = firstDirectChild(axisEl, WP_NS, 'align');
    if (alignEl) {
      const txt = textContentOfDirect(alignEl).trim();
      if (txt) out.align = txt;
    }
  }
  return out;
}

/** 从 wp:anchor 解析 wrap 类型 */
function readWrap(anchor: Element): DrawingWrapType | undefined {
  // 优先 topAndBottom / square / tight / through
  if (firstDirectChild(anchor, WP_NS, 'wrapTopAndBottom')) return 'topAndBottom';
  if (firstDirectChild(anchor, WP_NS, 'wrapSquare')) return 'square';
  if (firstDirectChild(anchor, WP_NS, 'wrapTight')) return 'tight';
  if (firstDirectChild(anchor, WP_NS, 'wrapThrough')) return 'through';
  if (firstDirectChild(anchor, WP_NS, 'wrapNone')) {
    // wp:wrapNone 自身存在；behindDoc 在 anchor 元素属性上
    // ECMA-376 ST_OnOff：支持 "1"/"true"/"on" 与 "0"/"false"/"off" 多种字面（见 parseOnOff）
    const behind = parseOnOff(anchor.getAttribute('behindDoc'));
    if (behind === true) return 'behind';
    if (behind === false) return 'inFront';
    return 'none';
  }
  return undefined;
}

/** 从 wp:inline / wp:anchor 解析 extent */
function readExtent(container: Element): { cx: number; cy: number } {
  const ext = firstDirectChild(container, WP_NS, 'extent');
  if (!ext) return { cx: 0, cy: 0 };
  const cx = readNumAttr(ext, null, 'cx') ?? 0;
  const cy = readNumAttr(ext, null, 'cy') ?? 0;
  return { cx, cy };
}

/** 从 wp:anchor 解析 position */
function readPosition(anchor: Element): DrawingPosition {
  const out: DrawingPosition = {};
  const h = firstDirectChild(anchor, WP_NS, 'positionH');
  if (h) out.horizontal = readPositionAxis(h);
  const v = firstDirectChild(anchor, WP_NS, 'positionV');
  if (v) out.vertical = readPositionAxis(v);
  return out;
}

/**
 * 从 wp:inline / wp:anchor 子树识别 kind。
 *
 * 一级判定：a:graphicData @uri；命中 6 个预设 URI 之一即落对应 kind。
 * 二级判定：shape kind 内若有 w:txbxContent → 重定为 textBox。
 * 都不命中 → 'other'。
 */
function detectKind(container: Element): DrawingKind {
  const graphic = firstDescendant(container, A_NS, 'graphic');
  if (!graphic) return 'other';
  const gd = firstDescendant(graphic, A_NS, 'graphicData');
  if (!gd) return 'other';
  const uri = gd.getAttribute('uri') ?? '';

  if (uri === PIC_NS) return 'picture';
  if (uri === CHART_NS) return 'chart';
  if (uri === DGM_NS) return 'smartArt';
  if (uri === WPG_NS) return 'group';
  if (uri === WPC_NS) return 'canvas';
  if (uri === WPS_NS) {
    // 区分 shape vs textBox：wps:wsp 子树里若有 w:txbxContent，认定为 textBox
    if (gd.getElementsByTagNameNS(W_NS, 'txbxContent').length > 0) return 'textBox';
    return 'shape';
  }
  return 'other';
}

/** 检测整个 drawing 子树是否含 pic:pic（A 类正向信号；包括套在其他对象内的情况）*/
function hasPicReferenceIn(container: Element): boolean {
  return container.getElementsByTagNameNS(PIC_NS, 'pic').length > 0;
}

/**
 * 线条 / 连接线的 DrawingML 预设几何（a:prstGeom @prst）白名单。
 *
 * - `line`：经典直线（Word「插入→形状→直线」生成，母版版记横线即此）
 * - `straightConnector1`：直线连接符
 * - `bentConnector2..5` / `curvedConnector2..5`：折线 / 曲线连接符（流程图连线）
 *
 * 仅收"视觉上是一根线/连线"的预设；矩形 rect / 椭圆 ellipse / 标注 callout 等不在内。
 * 收紧白名单而非黑名单：未知新预设默认按"非线条"处理，保守不误保。
 */
const LINE_CONNECTOR_PRESETS = new Set<string>([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
]);

/**
 * 是否为「线条 / 连接线」几何（只对 kind = 'shape' 有意义）。
 *
 * 命中任一即 true（从强到弱）：
 *  1. a:prstGeom @prst ∈ LINE_CONNECTOR_PRESETS
 *  2. 含 wps:cNvCnPr（连接线非可视属性；普通 shape 是 wps:cNvSpPr）
 *  3. 含 VML v:line（旧版降级线条；防御性兜底，通常在 mc:Fallback 内，
 *     而本函数收到的 container 是 mc:Choice 内的 wp:inline/anchor，故此项多数情况不命中，
 *     仅覆盖 container 内嵌 VML 的边角情形）
 */
function isLineOrConnectorShape(container: Element): boolean {
  // 信号 1：a:prstGeom @prst 命中线条/连接线预设
  const prstGeoms = container.getElementsByTagNameNS(A_NS, 'prstGeom');
  for (let i = 0; i < prstGeoms.length; i++) {
    const prst = prstGeoms[i]?.getAttribute('prst');
    if (prst && LINE_CONNECTOR_PRESETS.has(prst)) return true;
  }
  // 信号 2：wps:cNvCnPr（connector non-visual props）
  if (container.getElementsByTagNameNS(WPS_NS, 'cNvCnPr').length > 0) return true;
  // 信号 3：VML v:line（防御性兜底）
  if (container.getElementsByTagNameNS(VML_NS, 'line').length > 0) return true;
  return false;
}

/**
 * Shape / TextBox 形态信号：边框 / 填充 / 内部文字 / 首段对齐。
 *
 * 只在 kind = shape | textBox 时调用。返回顺序与 DrawingInfo 字段一致：
 *   { hasBorder, hasFill, hasInnerText?, innerText?, innerParagraphIds?, innerTextAlign? }
 *
 * @param runTextOpts 传 `{ structuralWhitespace: true }` 时 innerText 保真读（含 \t / \n）
 */
function readShapeSignals(
  container: Element,
  isTextBox: boolean,
  runTextOpts?: RunTextOptions,
  drawingToId: Map<Element, string> = new Map(),
  tableToId: Map<Element, string> = new Map(),
): {
  hasBorder: boolean;
  hasFill: boolean;
  hasInnerText?: boolean;
  innerText?: string;
  innerParagraphIds?: string[];
  innerBlocks?: TextBoxBlockInfo[];
  innerTextAlign?: string;
} {
  // wps:spPr 是 wps namespace 下的 spPr（与 a:spPr 不同的命名空间，但元素结构一致）
  // 也可能在 a:spPr 下 —— 实际生成的 docx 主要走 wps:spPr，我们两个 NS 都查
  const wspSpPr = firstDescendant(container, WPS_NS, 'spPr');
  const aSpPr = firstDescendant(container, A_NS, 'spPr');
  const spPr = wspSpPr ?? aSpPr;

  let hasBorder = false;
  let hasFill = false;
  if (spPr) {
    // 边框：a:ln 存在 且 未声明 a:noFill 子元素 且 w 属性 != "0"
    //
    // 三个否定条件：
    //   1. a:ln 不存在     → 无线条定义     → hasBorder = false
    //   2. <a:ln><a:noFill/></a:ln> → 显式声明无填充 → hasBorder = false
    //   3. <a:ln w="0"/>   → 零 EMU 宽度    → 视觉无边框 → hasBorder = false
    const ln = firstDirectChild(spPr, A_NS, 'ln');
    if (ln) {
      const noLineFill = firstDirectChild(ln, A_NS, 'noFill');
      // w 是 a:ln 的属性（非命名空间属性），可能为 "0" 表示零宽线
      const widthAttr = ln.getAttribute('w');
      const isZeroWidth = widthAttr !== null && widthAttr.trim() === '0';
      hasBorder = !noLineFill && !isZeroWidth;
    }
    // 填充：a:solidFill / a:gradFill / a:blipFill / a:pattFill 之一存在，且不是 noFill
    const noFill = firstDirectChild(spPr, A_NS, 'noFill');
    if (!noFill) {
      hasFill =
        !!firstDirectChild(spPr, A_NS, 'solidFill') ||
        !!firstDirectChild(spPr, A_NS, 'gradFill') ||
        !!firstDirectChild(spPr, A_NS, 'blipFill') ||
        !!firstDirectChild(spPr, A_NS, 'pattFill');
    }
  }

  const out: {
    hasBorder: boolean;
    hasFill: boolean;
    hasInnerText?: boolean;
    innerText?: string;
    innerParagraphIds?: string[];
    innerBlocks?: TextBoxBlockInfo[];
    innerTextAlign?: string;
  } = { hasBorder, hasFill };

  if (isTextBox) {
    const txbx = firstDescendant(container, W_NS, 'txbxContent');
    if (txbx) {
      const paragraphs = snapshot(txbx.getElementsByTagNameNS(W_NS, 'p'));
      const ids: string[] = [];
      for (const p of paragraphs) {
        const id = p.getAttributeNS(W14_NS, 'paraId');
        if (id) ids.push(id.toUpperCase());
      }
      out.innerParagraphIds = ids;
      out.innerBlocks = collectTextBoxBlocks(txbx, drawingToId, tableToId, runTextOpts);

      // 内部文字：按 runTextOpts 决定拼接口径：
      //   - 默认（无 opts）：所有 w:t 节点文本按文档顺序拼接（与历史行为一致，hasInnerText 同源）
      //   - structuralWhitespace: true：按 run 子节点顺序走保真读（\t/\n 还原）
      let joined = '';
      if (runTextOpts?.structuralWhitespace) {
        // 保真读：枚举 txbxContent 内所有直属段落的直属 run，逐个 getRunText 保真
        const runs = snapshot(txbx.getElementsByTagNameNS(W_NS, 'r'));
        for (const r of runs) {
          joined += getRunText(r, runTextOpts);
        }
      } else {
        // 默认：只拼 w:t（历史行为，hasInnerText 计算依据不变）
        const tEls = snapshot(txbx.getElementsByTagNameNS(W_NS, 't'));
        for (const t of tEls) {
          joined += textContentOfDirect(t);
        }
      }
      out.hasInnerText = joined.length > 0;
      const trimmed = joined.trim();
      if (trimmed.length > 0) out.innerText = trimmed;

      // 首段对齐：第一个 w:p 的 w:pPr/w:jc @val
      const firstP = firstDescendant(txbx, W_NS, 'p');
      if (firstP) {
        const pPr = firstDirectChild(firstP, W_NS, 'pPr');
        if (pPr) {
          const jc = firstDirectChild(pPr, W_NS, 'jc');
          if (jc) {
            const val = jc.getAttributeNS(W_NS, 'val');
            if (val) out.innerTextAlign = val;
          }
        }
      }
    } else {
      out.hasInnerText = false;
    }
  }

  return out;
}

/**
 * 取承载 drawing 的 w:p 元素和在父 w:p 下的 run 序号；找不到时返回 nulls。
 *
 * **两段式**（CR 设计澄清）：
 *
 *   1. 向上找最近的 `w:r` 和最近的 `w:p`（这一步遇到 `txbxContent` 前会先撞到内层 `w:p`，
 *      旧实现把 txbxContent 守卫放在这一步是死代码 —— 永远触发不到）
 *
 *   2. 找到 `w:p` 后**继续向上**检查它是否落在 `w:txbxContent` 子树里。若是，返回
 *      `{ para: null, runIndex: undefined }`，让 `DrawingInfo.anchorParaId` 留空。
 *
 * 为什么这样设计：文本框内的 `w:p` 与外层 `w:p` 是**两套独立段落流**（不在同一
 * paraIdx scope）。如果把"文本框内段落 paraId"挂给 drawing：
 *   - 下游"删 paraId" 命令找不到（文本框内段落不进 paraIdx body 视图）
 *   - 误以为"删这个段落 = 删这张图"，结果删的是文本框内的整段，留下空文本框
 *
 * 文档站根（body / hdr / ftr / footnote / endnote / comment）作为向上扫描的边界，
 * 避免穿过 story 边界。
 */
function findAnchorParagraph(drawing: Element): {
  para: Element | null;
  runIndex: number | undefined;
} {
  // 第 1 步：向上找最近的 w:r 与最近的 w:p
  let p: Element | null = drawing.parentElement;
  let runEl: Element | null = null;
  let containerP: Element | null = null;
  while (p) {
    if (p.namespaceURI === W_NS) {
      if (p.localName === 'r' && !runEl) runEl = p;
      if (p.localName === 'p') {
        containerP = p;
        break;
      }
    }
    p = p.parentElement;
  }
  if (!containerP) return { para: null, runIndex: undefined };

  // 第 2 步：检查 containerP 是否落在 txbxContent 内 —— 若是，主动返回 null
  // 防御 story 根（body / hdr / ftr / footnote / endnote / comment）作为扫描边界
  let q: Element | null = containerP.parentElement;
  while (q) {
    if (q.namespaceURI === W_NS) {
      if (q.localName === 'txbxContent') {
        return { para: null, runIndex: undefined };
      }
      if (
        q.localName === 'body' ||
        q.localName === 'hdr' ||
        q.localName === 'ftr' ||
        q.localName === 'footnote' ||
        q.localName === 'endnote' ||
        q.localName === 'comment'
      ) {
        break;
      }
    }
    q = q.parentElement;
  }

  // 第 3 步：在 containerP 直接子中找 runEl 的序号
  let runIndex: number | undefined;
  if (runEl) {
    let n = 0;
    for (let i = 0; i < containerP.childNodes.length; i++) {
      const child = containerP.childNodes[i] as Element;
      if (child && child.nodeType === 1 && child.namespaceURI === W_NS && child.localName === 'r') {
        if (child === runEl) {
          runIndex = n;
          break;
        }
        n++;
      }
    }
  }

  return { para: containerP, runIndex };
}

/** 判断 drawing 是否嵌套在另一个 drawing 内（用于 group / canvas 内子图过滤的预判）*/
function hasDrawingAncestor(drawing: Element): boolean {
  let p = drawing.parentElement;
  while (p) {
    if (p.namespaceURI === W_NS && p.localName === 'drawing' && p !== drawing) return true;
    p = p.parentElement;
  }
  return false;
}

function nearestAncestor(el: Element, ns: string, localName: string): Element | null {
  let p = el.parentElement;
  while (p) {
    if (p.namespaceURI === ns && p.localName === localName) return p;
    p = p.parentElement;
  }
  return null;
}

function paragraphText(p: Element, runTextOpts?: RunTextOptions): string {
  let text = '';
  for (let i = 0; i < p.childNodes.length; i++) {
    const child = p.childNodes[i] as Element;
    if (!child || child.nodeType !== 1 || child.namespaceURI !== W_NS) continue;
    if (child.localName === 'r') text += getRunText(child, runTextOpts);
  }
  return text;
}

function directDrawingIdsInTextBoxParagraph(
  p: Element,
  ownerTxbx: Element,
  drawingToId: Map<Element, string>,
): string[] {
  const ids: string[] = [];
  const drawings = p.getElementsByTagNameNS(W_NS, 'drawing');
  for (let i = 0; i < drawings.length; i++) {
    const drawing = drawings[i];
    if (!drawing || nearestAncestor(drawing, W_NS, 'txbxContent') !== ownerTxbx) continue;
    const id = drawingToId.get(drawing);
    if (id) ids.push(id);
  }
  return ids;
}

function collectTextBoxBlocks(
  txbx: Element,
  drawingToId: Map<Element, string>,
  tableToId: Map<Element, string>,
  runTextOpts?: RunTextOptions,
): TextBoxBlockInfo[] {
  const blocks: TextBoxBlockInfo[] = [];
  for (let i = 0; i < txbx.childNodes.length; i++) {
    const child = txbx.childNodes[i] as Element;
    if (!child || child.nodeType !== 1 || child.namespaceURI !== W_NS) continue;

    if (child.localName === 'p') {
      const paraId = child.getAttributeNS(W14_NS, 'paraId') ?? undefined;
      const text = paragraphText(child, runTextOpts);
      const drawingIds = directDrawingIdsInTextBoxParagraph(child, txbx, drawingToId);
      if (paraId && (text.length > 0 || drawingIds.length === 0)) {
        blocks.push({ type: 'paragraph', paraId: paraId.toUpperCase(), text });
      }
      for (const drawingId of drawingIds) {
        blocks.push(
          paraId
            ? { type: 'drawing', drawingId, paraId: paraId.toUpperCase() }
            : { type: 'drawing', drawingId },
        );
      }
      continue;
    }

    if (child.localName === 'tbl') {
      const tableId = tableToId.get(child);
      if (tableId) blocks.push({ type: 'table', tableId });
    }
  }
  return blocks;
}

function reverseElementIndex(index: Map<string, Element>): Map<Element, string> {
  const reversed = new Map<Element, string>();
  for (const [id, el] of index) reversed.set(el, id);
  return reversed;
}

/** 取所在 story（用 contains 检测，stories 数量通常 < 20）*/
function findStoryForEl(el: Element, stories: Story[]): Story | undefined {
  for (const s of stories) {
    if (s.root.contains(el as unknown as Node)) return s;
  }
  return undefined;
}

/** 判 partPath 对应的 contentType 是否是 header / footer */
function isHeaderFooterPart(pkg: Package, partPath: string): boolean {
  const part = pkg.parts.get(partPath);
  if (!part) return false;
  const ct = part.contentType;
  return ct.endsWith('.header+xml') || ct.endsWith('.footer+xml');
}

/** 分析单个 w:drawing → DrawingInfo（不含 drawingId / partPath / index / isInHeaderFooter）*/
function analyzeDrawing(
  drawing: Element,
  runTextOpts?: RunTextOptions,
  drawingToId?: Map<Element, string>,
  tableToId?: Map<Element, string>,
): Omit<DrawingInfo, 'drawingId' | 'partPath' | 'index' | 'isInHeaderFooter'> {
  // 容器：drawing 的直接子是 wp:inline 或 wp:anchor
  const inlineEl = firstDirectChild(drawing, WP_NS, 'inline');
  const anchorEl = firstDirectChild(drawing, WP_NS, 'anchor');
  const container = inlineEl ?? anchorEl;

  const anchoringMode: DrawingAnchoringMode = anchorEl ? 'anchor' : 'inline';

  const kind: DrawingKind = container ? detectKind(container) : 'other';
  const extentEMU = container ? readExtent(container) : { cx: 0, cy: 0 };
  const hasPicReference = container ? hasPicReferenceIn(container) : false;

  const { para, runIndex } = findAnchorParagraph(drawing);
  const anchorParaId = para?.getAttributeNS(W14_NS, 'paraId') ?? undefined;

  const info: Omit<DrawingInfo, 'drawingId' | 'partPath' | 'index' | 'isInHeaderFooter'> = {
    anchoringMode,
    kind,
    extentEMU,
    hasPicReference,
  };

  if (anchorEl) {
    const wrap = readWrap(anchorEl);
    if (wrap !== undefined) info.wrap = wrap;
    const pos = readPosition(anchorEl);
    if (pos.horizontal || pos.vertical) info.position = pos;
  }

  if (anchorParaId !== undefined) info.anchorParaId = anchorParaId;
  if (runIndex !== undefined) info.runIndex = runIndex;

  if (kind === 'shape' || kind === 'textBox') {
    const sig = container
      ? readShapeSignals(container, kind === 'textBox', runTextOpts, drawingToId, tableToId)
      : { hasBorder: false, hasFill: false };
    info.hasBorder = sig.hasBorder;
    info.hasFill = sig.hasFill;
    if (sig.hasInnerText !== undefined) info.hasInnerText = sig.hasInnerText;
    if (sig.innerText !== undefined) info.innerText = sig.innerText;
    if (sig.innerParagraphIds !== undefined) info.innerParagraphIds = sig.innerParagraphIds;
    if (sig.innerBlocks !== undefined) info.innerBlocks = sig.innerBlocks;
    if (sig.innerTextAlign !== undefined) info.innerTextAlign = sig.innerTextAlign;
  }

  // 线条/连接线信号仅对纯几何 shape 有意义（textBox 是内容载体，不参与线条判定）
  if (kind === 'shape' && container) {
    info.isLineOrConnector = isLineOrConnectorShape(container);
  }

  return info;
}

// ─────────── 公开函数 ───────────

/**
 * 枚举文档 drawing 对象，按文档顺序返回（与 drawingIdx 插入顺序一致）。
 *
 * - `scope: 'body'`（默认）—— 仅主文档 body 内的 drawing
 * - `scope: 'all'` —— 包含 headers / footers / footnotes / endnotes / comments
 *
 * **不做**业务分类（A/B/C）/ 分组键计算 / 代表样本选取 —— 这些属于 #191/#192。
 * 嵌套在另一个 drawing 内的子 drawing（如 group 的成员、canvas 的成员）**会被一并枚出**，
 * 由调用方按 hasDrawingAncestor 自行决定是否合并到外层（保持枚举层"完整暴露"的原则）。
 */
export function listDrawings(
  pkg: Package,
  indices: PackageIndices,
  opts?: ListDrawingsOptions,
): DrawingInfo[] {
  const scope = opts?.scope ?? 'body';
  const stories = resolveScope(pkg, scope);
  const result: DrawingInfo[] = [];
  let idx = 1;
  const runTextOpts: RunTextOptions | undefined = opts?.structuralWhitespace
    ? { structuralWhitespace: true }
    : undefined;
  const drawingToId = reverseElementIndex(indices.drawingIdx);
  const tableToId = reverseElementIndex(indices.tblIdx);

  // 直接遍历 drawingIdx（插入顺序 = 文档顺序）
  for (const [drawingId, drawing] of indices.drawingIdx) {
    const story = findStoryForEl(drawing, stories);
    if (!story) continue; // 不在请求的 scope 内

    const info: DrawingInfo = {
      drawingId,
      partPath: story.partPath,
      index: idx++,
      isInHeaderFooter: isHeaderFooterPart(pkg, story.partPath),
      ...analyzeDrawing(drawing, runTextOpts, drawingToId, tableToId),
    };
    result.push(info);
  }

  return result;
}

/**
 * 按 drawingId 精确取单个 drawing。不存在返回 null。
 *
 * **index 字段语义**：返回 drawing 在 **全局 drawingIdx 插入顺序**中的 1-based 位置
 *  —— 与 `listDrawings({ scope: 'all' })` 的 index 一致；与 `listDrawings({ scope: 'body' })`
 * 的 1-based 位置**不一定**相等（后者可能过滤掉了页眉页脚 drawing）。
 *
 * 复杂度：drawingIdx 查找 O(1)；全局序号扫描 O(N)，N = 包内 drawing 总数。
 * 调用方需要严格按 scope 取序号时应直接走 listDrawings。
 *
 * 旧实现复用了 list-tables.ts 的 `index: 0` 哨兵，下游用 `index` 排序/引用会拿到
 * 错误值（CR BUG-1）—— 改为真实序号。
 */
export function getDrawingById(
  pkg: Package,
  indices: PackageIndices,
  drawingId: string,
): DrawingInfo | null {
  const drawing = indices.drawingIdx.get(drawingId);
  if (!drawing) return null;

  // 找所属 story 以获取 partPath
  const stories = resolveScope(pkg, 'all');
  const story = findStoryForEl(drawing, stories);
  if (!story) return null;
  const drawingToId = reverseElementIndex(indices.drawingIdx);
  const tableToId = reverseElementIndex(indices.tblIdx);

  // 计算全局 1-based 序号（drawingIdx 是 Map，迭代序 = 文档顺序）
  let globalIndex = 0;
  let i = 0;
  for (const id of indices.drawingIdx.keys()) {
    i++;
    if (id === drawingId) {
      globalIndex = i;
      break;
    }
  }

  return {
    drawingId,
    partPath: story.partPath,
    index: globalIndex,
    isInHeaderFooter: isHeaderFooterPart(pkg, story.partPath),
    ...analyzeDrawing(drawing, undefined, drawingToId, tableToId),
  };
}

/**
 * 装载期构建 drawingIdx：枚举 Package 中所有 Word 内容 part 的 <w:drawing>，
 * 按文档顺序分配 'dr-NNNN' ID。由 buildIndices 调用。
 *
 * 内部函数风格对齐 buildTableIndex（list-tables.ts）。
 */
export function buildDrawingIndex(
  parts: Iterable<{
    contentType: string;
    xmlDom?: {
      getElementsByTagNameNS(
        ns: string,
        local: string,
      ): { length: number } & Record<number, Element>;
    };
  }>,
): Map<string, Element> {
  const drawingIdx = new Map<string, Element>();
  let n = 1;

  for (const part of parts) {
    const ct = part.contentType;
    // 与 ListDrawingsOptions / DrawingsApi.list 的 scope:'all' 语义对齐：
    // body + headers + footers + footnotes + endnotes + comments
    //
    // CR 修复：旧实现注释「comments.xml 内通常不含 drawing，跳过」是错误假设
    // —— 评论 part 内合法可含 drawing（如批注图示），漏掉会让 scope:'all' 与文档承诺不一致。
    if (
      !ct.includes('wordprocessingml') &&
      !ct.endsWith('.main+xml') &&
      !ct.endsWith('.header+xml') &&
      !ct.endsWith('.footer+xml') &&
      !ct.endsWith('.footnotes+xml') &&
      !ct.endsWith('.endnotes+xml') &&
      !ct.endsWith('.comments+xml')
    ) {
      continue;
    }
    if (!part.xmlDom) continue;

    const drawings = part.xmlDom.getElementsByTagNameNS(W_NS, 'drawing');
    for (let i = 0; i < drawings.length; i++) {
      const dr = drawings[i];
      if (!dr) continue;
      // CR BUG-3：跳过 mc:Fallback 内的 drawing —— mc:Choice 已是首选呈现，
      // Fallback 是降级方案，把它也计入会导致同一对象双计。
      if (hasFallbackAncestor(dr as Element)) continue;
      // 嵌套 drawing：group / canvas 的成员 drawing 仍是 w:drawing，但通常 group/canvas
      // 的 child shapes 在 wpg:grpSp / wpc:wpc 下，不会用 w:drawing 包装。不做主动过滤
      // —— 让上层（#191 / #192）按需用 hasDrawingAncestor 自己判。
      const id = `dr-${String(n).padStart(4, '0')}`;
      drawingIdx.set(id, dr);
      n++;
    }
  }

  return drawingIdx;
}

/** 暴露内部工具：判断元素是否嵌套在另一个 drawing 内（#191 合并 group 子图时用）*/
export { hasDrawingAncestor };
