/**
 * SDT (Structured Document Tag, Content Control) XML 构建器 —— #188 表格域化。
 *
 * 设计:
 *   - 三种子类型: Plain Text / Date Picker / Rich Text
 *   - 输出 inline SDT (作 <w:r> 兄弟出现在 <w:p> 内), 段落 pPr 由调用方保留
 *   - SDT id 由调用方传入 (典型: hash 输入对象保证幂等), 避免 builder 维护全局状态
 *   - rPr 可由调用方拷贝原 run 的 rPr 传入, 保留字体 / 颜色 / 字号等字符属性
 *
 * SDT id 取值: OOXML 规定 32-bit signed integer; Word 加载时容错并按需重分配,
 * 调用方不必严格保证全局唯一, hash 取值即可。
 *
 * 参考: ECMA-376 §17.5 Structured Document Tags。
 */

import { W_NS, XML_NS } from './namespaces.js';

// ─────────── 公开类型 ───────────

/**
 * SDT 配置 (JSON-serializable, 可走 plan op RPC)。
 *
 * `rPr` 字符属性透传不在 spec 内 —— DOM Element 不能 JSON 序列化, 由 executor 在
 * 替换时从原 run 直接克隆并传给 `buildInlineSdt` 的 `rPr` 参数。
 */
interface SdtSpecBase {
  /** UI 显示名 (Word "Click here..." 上方的提示标签) */
  alias: string;
  /**
   * 程序匹配 tag (后续模板填充阶段按 tag 定位 SDT)。
   * 缺省时按 `alias` trim + 空白转 `_`。
   */
  tag?: string;
  /** 占位提示文本; 缺省 `"请输入" + alias` */
  placeholderText?: string;
}

export interface PlainTextSdtSpec extends SdtSpecBase {
  kind: 'text';
}

export interface DateSdtSpec extends SdtSpecBase {
  kind: 'date';
  /** 日期格式 (e.g. "yyyy-M-d"), 缺省 "yyyy-M-d" */
  dateFormat?: string;
  /** Locale, 缺省 "zh-CN" */
  locale?: string;
}

export interface RichTextSdtSpec extends SdtSpecBase {
  kind: 'richText';
}

export type SdtSpec = PlainTextSdtSpec | DateSdtSpec | RichTextSdtSpec;

// ─────────── 主入口 ───────────

/**
 * 构建一个 inline SDT 元素 (`<w:sdt>`), 用于段落内 run 替换场景。
 *
 * 结构:
 *   <w:sdt>
 *     <w:sdtPr>
 *       (rPr?)
 *       <w:alias w:val="..."/>
 *       <w:tag   w:val="..."/>
 *       <w:id    w:val="..."/>
 *       <w:placeholder><w:docPart w:val="DefaultPlaceholder_-1854013440"/></w:placeholder>
 *       <w:showingPlcHdr/>
 *       (类型标记: <w:text/> | <w:date.../> | <w:richText/>)
 *     </w:sdtPr>
 *     <w:sdtContent>
 *       <w:r>(rPr?)<w:t xml:space="preserve">placeholderText</w:t></w:r>
 *     </w:sdtContent>
 *   </w:sdt>
 *
 * @param doc       owner Document (来自 paragraph.ownerDocument)
 * @param spec      SDT 配置 (JSON-serializable)
 * @param id        32-bit 整数 ID; 推荐用 hashSdtId(tableId, row, col, alias) 算
 * @param rPr       可选, 从原 run 克隆的 `<w:rPr>` —— 字符属性透传, 同时挂到
 *                  sdtPr 与 sdtContent 内 placeholder run 上
 */
export function buildInlineSdt(doc: Document, spec: SdtSpec, id: number, rPr?: Element): Element {
  const sdt = createW(doc, 'sdt');

  // ─── sdtPr ───
  const sdtPr = createW(doc, 'sdtPr');
  sdt.appendChild(sdtPr);

  // rPr (字符属性透传, 影响 sdt 内 placeholder run 的默认字符样式)
  if (rPr) {
    sdtPr.appendChild(rPr.cloneNode(true));
  }

  // alias
  const alias = createW(doc, 'alias');
  alias.setAttributeNS(W_NS, 'w:val', spec.alias);
  sdtPr.appendChild(alias);

  // tag
  const tag = createW(doc, 'tag');
  tag.setAttributeNS(W_NS, 'w:val', spec.tag ?? deriveTag(spec.alias));
  sdtPr.appendChild(tag);

  // id
  const idEl = createW(doc, 'id');
  idEl.setAttributeNS(W_NS, 'w:val', String(normalizeSdtId(id)));
  sdtPr.appendChild(idEl);

  // placeholder → docPart 引用 (Word 默认占位)
  const placeholder = createW(doc, 'placeholder');
  const docPart = createW(doc, 'docPart');
  docPart.setAttributeNS(W_NS, 'w:val', 'DefaultPlaceholder_-1854013440');
  placeholder.appendChild(docPart);
  sdtPr.appendChild(placeholder);

  // showingPlcHdr (当前显示占位符状态)
  sdtPr.appendChild(createW(doc, 'showingPlcHdr'));

  // 类型标记
  switch (spec.kind) {
    case 'text':
      sdtPr.appendChild(createW(doc, 'text'));
      break;
    case 'richText':
      sdtPr.appendChild(createW(doc, 'richText'));
      break;
    case 'date':
      sdtPr.appendChild(buildDateChild(doc, spec));
      break;
  }

  // ─── sdtContent ───
  const sdtContent = createW(doc, 'sdtContent');
  sdt.appendChild(sdtContent);

  const placeholderText = spec.placeholderText ?? `请输入${spec.alias}`;
  const run = createW(doc, 'r');
  if (rPr) {
    run.appendChild(rPr.cloneNode(true));
  }
  const t = createW(doc, 't');
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = placeholderText;
  run.appendChild(t);
  sdtContent.appendChild(run);

  return sdt;
}

// ─────────── ID hash ───────────

/**
 * 由 (tableId, row, col, alias) 算 SDT id (djb2 hash → 32-bit signed)。
 *
 * 幂等性保证: 同一组输入永远产出同一 id, 重跑骨架化不会因 id 变化制造无意义 diff。
 */
export function hashSdtId(tableId: string, row: number, col: number, alias: string): number {
  const input = `${tableId}|${row}|${col}|${alias}`;
  let h = 5381;
  for (let i = 0; i < input.length; i++) {
    h = ((h << 5) + h + input.charCodeAt(i)) | 0;
  }
  return normalizeSdtId(h);
}

/** SDT id 取 32-bit signed 范围内的正数, 避免负号引起 XML 兼容性问题。 */
function normalizeSdtId(id: number): number {
  const n = id | 0;
  return n < 0 ? -n : n;
}

// ─────────── 内部工具 ───────────

function createW(doc: Document, localName: string): Element {
  return doc.createElementNS(W_NS, `w:${localName}`);
}

function buildDateChild(doc: Document, spec: DateSdtSpec): Element {
  const date = createW(doc, 'date');

  const dateFormat = createW(doc, 'dateFormat');
  dateFormat.setAttributeNS(W_NS, 'w:val', spec.dateFormat ?? 'yyyy-M-d');
  date.appendChild(dateFormat);

  const lid = createW(doc, 'lid');
  lid.setAttributeNS(W_NS, 'w:val', spec.locale ?? 'zh-CN');
  date.appendChild(lid);

  const storeAs = createW(doc, 'storeMappedDataAs');
  storeAs.setAttributeNS(W_NS, 'w:val', 'dateTime');
  date.appendChild(storeAs);

  const calendar = createW(doc, 'calendar');
  calendar.setAttributeNS(W_NS, 'w:val', 'gregorian');
  date.appendChild(calendar);

  return date;
}

/**
 * 从 alias 派生 tag: 中文 alias 第一版直接返回原值 (OOXML 允许 tag 含 Unicode);
 * 仅做 trim + 去内部空白。后续可加 pinyin 转换或 hash 标识。
 */
function deriveTag(alias: string): string {
  return alias.trim().replace(/\s+/g, '_');
}
