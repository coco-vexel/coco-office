/**
 * TOC 域查询（#197）。
 *
 * 识别文档中的 TOC（目录）域：
 *   - complex field：`<w:fldChar w:fldCharType="begin">` … `end`，depth 计数跳过嵌套子域
 *   - simple field：`<w:fldSimple w:instr="TOC ...">` 单段闭合
 *
 * 只识别域名前缀匹配 `/^TOC\b/i` 的域，PAGEREF / HYPERLINK 等子域不收集。
 *
 * 嵌套处理：depth 计数器跟踪 fldChar 深度；instr 只在 depth=1（最外层）时收集。
 * complex TOC 的 firstParaId/lastParaId 是 begin/end fldChar 所在段；中间段也在区间内。
 */
import type { Package } from '../pack/index.js';
import { listStories } from './stories.js';
import { W_NS, W14_NS } from './namespaces.js';

export interface TocFieldRange {
  /** 文档顺序索引（0-based） */
  tocIndex: number;
  /** TOC 域第一段的 paraId（complex field = fldChar begin 段；simple field = fldSimple 段）*/
  firstParaId: string;
  /** TOC 域最后一段的 paraId（complex field = fldChar end 段；simple field = firstParaId）*/
  lastParaId: string;
  /**
   * 域指令文本（拼接所有 instrText / fldSimple@w:instr，trim 后）。
   * 示例：`'TOC \\o "1-3" \\h \\z \\u'`
   */
  instr: string;
}

/** 每个 body 直属 `<w:p>` 里所有有序节点事件（fldChar begin/end、instrText） */
interface ParaEvent {
  kind: 'begin' | 'end' | 'instr';
  text?: string; // instrText 的内容
}

/**
 * 按文档顺序产出 body 下的"段级" `<w:p>` —— **下钻 `<w:sdt>`/`<w:sdtContent>`，但不下钻 `<w:tbl>`**。
 *
 * root-cause（task-toc-leak，2026-06-14 坐实）：Word 2016+ 默认把目录（TOC）整体包进一个
 *   `<w:sdt>`（结构化文档标签 / content control，给"更新目录"按钮用）。真实文档
 *   `sample-files.com-formatted-report.docx` 的 TOC begin 段祖先链是
 *   `p > sdtContent > sdt > body`——深了两层。旧实现只遍历 `body.childNodes`（**直属**子节点），
 *   完全看不到 sdtContent 内的段 → fldChar begin/end/instrText 一个都收不到 → `listTocFields`
 *   返回 0 → 排版侧 inToc 全 false → TOC 14 条目泄漏成重复标题（标题数翻倍）。
 *
 * 修法：把"直属 p 遍历"改为这个生成器。它在 body 子树上按文档顺序递归：
 *   - `<w:p>` → 产出（段级单元，complex/simple field 逻辑照旧按段处理）；
 *   - `<w:sdt>` → 下钻其 `<w:sdtContent>`（可嵌套 sdt），把里面的 p 当作同级段继续产出；
 *   - `<w:sdtContent>` → 下钻（sdt 的直接子；也兼容直接遇到的情况）；
 *   - `<w:tbl>` → **跳过**（表格内段是独立的 cell 段流，跨进来会打乱段顺序、且表格内嵌 TOC
 *     极罕见，超本修范围）；其余元素跳过。
 *
 * 不回归既有用例：无 sdt 的文档里，该生成器产出的序列与"遍历 body 直属 p"完全一致
 *   （sectPr/其它非 p 非 sdt 元素本来就被跳过），simple/complex/嵌套子域/多域/跨节所有现有测试
 *   的段序、firstParaId/lastParaId 不变。
 */
function* iterTocParas(root: Element): Generator<Element> {
  for (let i = 0; i < root.childNodes.length; i++) {
    const node = root.childNodes[i];
    if (!node || node.nodeType !== 1) continue;
    const el = node as Element;
    if (el.namespaceURI !== W_NS) continue;
    if (el.localName === 'p') {
      yield el;
    } else if (el.localName === 'sdt' || el.localName === 'sdtContent') {
      // 下钻 sdt（→ 找其 sdtContent → 内部 p）/ sdtContent（→ 内部 p），保持文档顺序。
      yield* iterTocParas(el);
    }
    // <w:tbl> 及其它元素：跳过（不下钻）。
  }
}

function collectParaEvents(para: Element): ParaEvent[] {
  const events: ParaEvent[] = [];
  // 按 DOM 文档顺序遍历段内所有 run
  const runs = para.getElementsByTagNameNS(W_NS, 'r');
  for (let ri = 0; ri < runs.length; ri++) {
    const run = runs[ri];
    if (!run) continue;

    // fldChar
    const fldChars = run.getElementsByTagNameNS(W_NS, 'fldChar');
    for (let fi = 0; fi < fldChars.length; fi++) {
      const fc = fldChars[fi];
      if (!fc) continue;
      const ct = fc.getAttributeNS(W_NS, 'fldCharType');
      if (ct === 'begin') events.push({ kind: 'begin' });
      else if (ct === 'end') events.push({ kind: 'end' });
    }

    // instrText
    const instrTexts = run.getElementsByTagNameNS(W_NS, 'instrText');
    for (let it = 0; it < instrTexts.length; it++) {
      const itEl = instrTexts[it];
      if (!itEl) continue;
      events.push({ kind: 'instr', text: itEl.textContent ?? '' });
    }
  }
  return events;
}

export function listTocFields(pkg: Package): TocFieldRange[] {
  const stories = listStories(pkg, { kinds: ['body'] });
  if (stories.length === 0) return [];
  const body = stories[0]!.root;

  const out: TocFieldRange[] = [];
  let tocIndex = 0;

  // complex field state
  let depth = 0; // 全局 fldChar 嵌套深度（0 = 无活动域）
  let outerInstrParts: string[] = []; // depth=1 时收集的 instrText
  let outerFirstParaId: string | undefined; // begin 所在段
  let outerIsToc = false; // instrParts join 后匹配 TOC

  // 按文档顺序遍历 body 下的段级 p——下钻 sdt/sdtContent（覆盖被 content control 包裹的 TOC）。
  for (const el of iterTocParas(body)) {
    const paraId = (el.getAttributeNS(W14_NS, 'paraId') ?? '').toUpperCase();

    // ── simple field ──────────────────────────────────────────────
    // 只处理 depth=0 时遇到的 fldSimple（不在 complex field 内部）
    if (depth === 0) {
      const simples = el.getElementsByTagNameNS(W_NS, 'fldSimple');
      for (let si = 0; si < simples.length; si++) {
        const fs = simples[si];
        if (!fs) continue;
        const instr = (fs.getAttributeNS(W_NS, 'instr') ?? '').trim();
        if (/^TOC\b/i.test(instr) && paraId.length > 0) {
          out.push({
            tocIndex: tocIndex++,
            firstParaId: paraId,
            lastParaId: paraId,
            instr,
          });
        }
      }
    }

    // ── complex field events ───────────────────────────────────────
    const events = collectParaEvents(el);

    // 若 depth > 0，当前段已是活动 TOC 域内的一段（begin 段已在之前处理过）
    // 只需在处理事件前记录"当前段进入域"
    const wasInDomain = depth > 0;
    void wasInDomain; // 仅用于语义说明

    for (const ev of events) {
      if (ev.kind === 'begin') {
        depth++;
        if (depth === 1) {
          // 开始最外层域
          outerFirstParaId = paraId.length > 0 ? paraId : undefined;
          outerInstrParts = [];
          outerIsToc = false;
        }
      } else if (ev.kind === 'end') {
        if (depth === 1) {
          // 关闭最外层域
          const fullInstr = outerInstrParts.join('').trim();
          const isToc = outerIsToc || /^TOC\b/i.test(fullInstr);
          if (isToc && outerFirstParaId && paraId.length > 0) {
            out.push({
              tocIndex: tocIndex++,
              firstParaId: outerFirstParaId,
              lastParaId: paraId,
              instr: fullInstr,
            });
          }
          outerFirstParaId = undefined;
          outerInstrParts = [];
          outerIsToc = false;
        }
        if (depth > 0) depth--;
      } else if (ev.kind === 'instr' && depth === 1) {
        // 只收集最外层域的 instrText
        outerInstrParts.push(ev.text ?? '');
        if (!outerIsToc) {
          const joined = outerInstrParts.join('').trim();
          if (/^TOC\b/i.test(joined)) outerIsToc = true;
        }
      }
    }
  }

  return out;
}
