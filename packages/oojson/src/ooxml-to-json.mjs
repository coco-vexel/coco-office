// 工具②：OOXML(WordprocessingML body) → ONLYOFFICE-JSON。
// 写链路用：agent 改完的 OOXML → 本工具转 ONLYOFFICE-JSON → FromJSON+RemoveAll+InsertContent 灌回编辑器。
// FromJSON 严格：段落必须有 changes:[]、顶层必须有 paraPr。
// 无损保证：每个属性袋/元素 = { ...oo:x 兜底, ...本侧模型解析 }，模型键覆盖兜底（反映 agent 编辑）。
import { parseXml, kid, kidsOf } from "./xml.mjs";
import { colorToJson, highlightToJson, mergeAutoRgb } from "./color.mjs";
import { readExtra } from "./extra.mjs";
import { ooxmlDrawingToJson } from "./ooxml-drawing-to-json.mjs";

// 数值用 Number 而非 parseInt：ONLYOFFICE 宽度等可为小数(如 tcW.w=22.36)，parseInt 会截断丢精度。
const a2i = (v) => (v == null || v === "" ? null : Number(v));

// 缺口收集：解析时凡遇本侧未解析、且未被兜底属性覆盖的元素即记入（理论上应为空）。
export class Gaps {
  constructor() { this.counts = new Map(); }
  add(key) { this.counts.set(key, (this.counts.get(key) || 0) + 1); }
  merge(other) { for (const [k, v] of other.counts) this.counts.set(k, (this.counts.get(k) || 0) + v); }
  top(n = 50) { return [...this.counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, n); }
  get size() { return this.counts.size; }
}

function xmlShd(node) {
  if (!node) return undefined;
  const o = { val: node.attrs["w:val"], color: mergeAutoRgb(colorToJson(node.attrs["w:color"]), node.attrs["oo:crgb"]) };
  if (node.attrs["w:fill"]) o.fill = mergeAutoRgb(colorToJson(node.attrs["w:fill"]), node.attrs["oo:frgb"]);
  return { ...readExtra(node), ...o }; // themeColor/themeFill 等经 oo:x 兜底
}
function xmlPpToJson(pPrNode, gaps) {
  const o = { bFromDocument: true, type: "paraPr" };
  if (!pPrNode) return o;
  const g = (t) => kid(pPrNode, t);
  if (g("w:pStyle")) o.pStyle = g("w:pStyle").attrs["w:val"];
  for (const t of ["widowControl", "keepNext", "keepLines", "pageBreakBefore", "contextualSpacing", "suppressLineNumbers"])
    if (g("w:" + t)) o[t] = true;
  if (g("w:shd")) o.shd = xmlShd(g("w:shd"));
  const sp = g("w:spacing");
  if (sp) {
    o.spacing = {};
    if (sp.attrs["w:before"]) o.spacing.before = a2i(sp.attrs["w:before"]);
    if (sp.attrs["w:after"]) o.spacing.after = a2i(sp.attrs["w:after"]);
    if (sp.attrs["w:line"]) o.spacing.line = a2i(sp.attrs["w:line"]);
    if (sp.attrs["w:lineRule"]) o.spacing.lineRule = sp.attrs["w:lineRule"];
    if (sp.attrs["w:beforeAutospacing"]) o.spacing.beforeAutoSpacing = sp.attrs["w:beforeAutospacing"] === "1" ? "on" : "off";
    if (sp.attrs["w:afterAutospacing"]) o.spacing.afterAutoSpacing = sp.attrs["w:afterAutospacing"] === "1" ? "on" : "off";
  }
  const ind = g("w:ind");
  if (ind) {
    o.ind = {};
    if (ind.attrs["w:left"]) o.ind.left = a2i(ind.attrs["w:left"]);
    if (ind.attrs["w:right"]) o.ind.right = a2i(ind.attrs["w:right"]);
    if (ind.attrs["w:firstLine"]) o.ind.firstLine = a2i(ind.attrs["w:firstLine"]);
  }
  if (g("w:jc")) o.jc = g("w:jc").attrs["w:val"];
  return { ...readExtra(pPrNode), ...o };
}
function xmlRpToJson(rPrNode, gaps) {
  const o = { bFromDocument: true, type: "textPr" };
  if (!rPrNode) return o;
  const g = (t) => kid(rPrNode, t);
  const rf = g("w:rFonts");
  if (rf) { o.rFonts = {}; for (const k of ["ascii", "hAnsi", "eastAsia", "cs", "hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"]) if (rf.attrs["w:" + k]) o.rFonts[k] = rf.attrs["w:" + k]; }
  const tgl = (name) => { const e = g("w:" + name); if (!e) return undefined; const v = e.attrs["w:val"]; return !(v === "false" || v === "0"); };
  for (const t of ["b", "bCs", "i", "iCs"]) { const v = tgl(t); if (v !== undefined) o[t] = v; }
  if (g("w:color")) o.color = mergeAutoRgb(colorToJson(g("w:color").attrs["w:val"]), g("w:color").attrs["oo:rgb"]);
  if (g("w:spacing")) o.spacing = a2i(g("w:spacing").attrs["w:val"]);
  if (g("w:sz")) o.sz = a2i(g("w:sz").attrs["w:val"]);
  if (g("w:szCs")) o.szCs = a2i(g("w:szCs").attrs["w:val"]);
  if (g("w:shd")) o.shd = xmlShd(g("w:shd"));
  // highlight 不在此解析：OOXML 仅 16 命名色，非精确 RGB 会被最近邻映射丢精度；由 rPr 的 oo:x 兜底精确还原。
  if (g("w:vertAlign")) o.vertAlign = g("w:vertAlign").attrs["w:val"];
  const lang = g("w:lang");
  if (lang) { o.lang = {}; for (const k of ["val", "eastAsia", "bidi"]) if (lang.attrs["w:" + k]) o.lang[k] = lang.attrs["w:" + k]; }
  return { ...readExtra(rPrNode), ...o };
}
function runToJson(rNode, gaps) {
  const content = [];
  for (const c of rNode.kids) {
    // w:delText：被 <w:del> 追踪删除的文字载体（语义同 w:t，只是位于删除修订里）。
    if (c.tag === "w:t" || c.tag === "w:delText") content.push((c.kids || []).filter((k) => k.text != null).map((k) => k.text).join(""));
    else if (c.tag === "w:br") { const ty = c.attrs["w:type"]; content.push({ type: "break", breakType: ty === "column" ? "column" : "page" }); }
    else if (c.tag === "w:tab") content.push({ type: "tab" });
    else if (c.tag === "oo:ci") { try { content.push(JSON.parse(c.attrs["oo:v"])); } catch { /* 兜底损坏则跳过 */ } }
    else if (c.tag === "w:drawing" || c.tag === "mc:AlternateContent") {
      const drawing = ooxmlDrawingToJson(c);
      if (drawing) content.push(drawing);
      else gaps && gaps.add("w:r>" + c.tag);
    } else if (c.tag && c.tag !== "w:rPr") gaps && gaps.add("w:r>" + c.tag);
  }
  const ex = readExtra(rNode);
  return { ...ex, bFromDocument: true, rPr: xmlRpToJson(kid(rNode, "w:rPr"), gaps), content, footnotes: ex.footnotes || [], endnotes: ex.endnotes || [], reviewType: ex.reviewType || "common", type: "run" };
}
// <w:ins>/<w:del>（及 move 的 <w:moveTo>/<w:moveFrom>）是 run 的**修订包裹器**：包住一段
// 被插入/删除的 run。ONLYOFFICE-JSON 不用包裹器，而是给每个 run 打 reviewType + reviewInfo。
// 这里把包裹器拆开、递归内部 run 并标注——否则 SDK 的 trackChanges（校对）产出的真实
// <w:ins>/<w:del> 会被当未建模元素整个丢掉 → 插入/删除的文字在写回时全部丢失。
const REV_TAG_TYPE = { "w:ins": "add", "w:del": "remove", "w:moveTo": "add", "w:moveFrom": "remove" };
function revInfoFromEl(el) {
  return {
    userId: "",
    author: el.attrs["w:author"] || "",
    date: el.attrs["w:date"] || "",
    moveType: "noMove",
    prevType: -1,
  };
}
/** 把一个修订包裹器里的 run 递归收进 content（每个 run 标 reviewType/reviewInfo）。支持一层嵌套（move）。 */
function pushRevisionRuns(el, reviewType, reviewInfo, content, gaps) {
  for (const c of el.kids) {
    if (c.tag === "w:r") content.push({ ...runToJson(c, gaps), reviewType, reviewInfo });
    else if (REV_TAG_TYPE[c.tag]) pushRevisionRuns(c, REV_TAG_TYPE[c.tag], revInfoFromEl(c), content, gaps);
    else if (c.tag) gaps && gaps.add(el.tag + ">" + c.tag);
  }
}
function hyperlinkToJson(node, gaps, rels) {
  const content = [];
  for (const c of node.kids || []) {
    if (c.tag === "w:r") content.push(runToJson(c, gaps));
    else if (REV_TAG_TYPE[c.tag]) pushRevisionRuns(c, REV_TAG_TYPE[c.tag], revInfoFromEl(c), content, gaps);
    else if (c.tag === "oo:pi") { try { content.push(JSON.parse(c.attrs["oo:v"])); } catch { /* skip invalid fallback payload */ } }
    else if (c.tag) gaps && gaps.add("w:hyperlink>" + c.tag);
  }
  const relId = node.attrs["r:id"];
  const rel = relId && rels && typeof rels.get === "function" ? rels.get(relId) : null;
  const value = rel?.target || node.attrs["w:anchor"] || "";
  const tooltip = node.attrs["w:tooltip"] || "";
  return { ...readExtra(node), ...(tooltip ? { tooltip } : {}), ...(value ? { value } : {}), content, type: "hyperlink" };
}

function applyParagraphTextPr(item, paragraphRPr) {
  if (!item || !paragraphRPr) return item;
  if (item.type === "run") {
    return {
      ...item,
      rPr: {
        ...(item.rPr || {}),
        ...paragraphRPr,
        rFonts: {
          ...(item.rPr?.rFonts || {}),
          ...(paragraphRPr.rFonts || {}),
        },
      },
    };
  }
  if (item.type === "hyperlink" && Array.isArray(item.content)) {
    return {
      ...item,
      content: item.content.map((child) => applyParagraphTextPr(child, paragraphRPr)),
    };
  }
  return item;
}

function paraToJson(pNode, gaps, rels) {
  const pPrNode = kid(pNode, "w:pPr");
  // 段落 mark rPr 在 <w:pPr><w:rPr> 里：既是段落 rPr，也是末尾 endRun 的 rPr
  const markRpr = xmlRpToJson(pPrNode && kid(pPrNode, "w:rPr"), gaps);
  // 按文档顺序走内容项：run / oo:pi(commentRangeStart 等未建模项)
  const content = [];
  for (const c of pNode.kids) {
    if (c.tag === "w:r") content.push(runToJson(c, gaps));
    else if (c.tag === "w:hyperlink") content.push(hyperlinkToJson(c, gaps, rels));
    else if (REV_TAG_TYPE[c.tag]) pushRevisionRuns(c, REV_TAG_TYPE[c.tag], revInfoFromEl(c), content, gaps);
    else if (c.tag === "oo:pi") { try { content.push(JSON.parse(c.attrs["oo:v"])); } catch { /* 兜底损坏则跳过 */ } }
    else if (c.tag === "w:commentRangeStart") { content.push({ type: "commentRangeStart", id: Number(c.attrs["w:id"]) }); }
    else if (c.tag === "w:commentRangeEnd") { content.push({ type: "commentRangeEnd", id: Number(c.attrs["w:id"]) }); }
    else if (c.tag === "w:commentReference") { content.push({ type: "commentReference", id: Number(c.attrs["w:id"]) }); }
    else if (c.tag && c.tag !== "w:pPr") gaps && gaps.add("w:p>" + c.tag);
  }
  // 补回段落标记 endRun（ONLYOFFICE 每段 content 末尾都有）；oo:er 还原其修订等非默认字段。
  // WPS stores effective paragraph text formatting in pPr/rPr while leaving stale,
  // partial values on individual runs. ONLYOFFICE FromJSON does not apply the
  // paragraph-level textPr to those runs, so materialize the effective properties.
  for (let i = 0; i < content.length; i += 1) {
    content[i] = applyParagraphTextPr(content[i], markRpr);
  }
  const er = readExtra(pNode, "oo:er");
  content.push({ bFromDocument: true, footnotes: [], endnotes: [], reviewType: "common", ...er, rPr: markRpr, content: [], type: "endRun" });
  const ex = readExtra(pNode);
  return {
    ...ex,
    bFromDocument: true,
    pPr: xmlPpToJson(pPrNode, gaps),
    rPr: markRpr,
    content,
    changes: ex.changes || [],
    type: "paragraph",
  };
}

// 块级内容数组(body 或单元格)：w:p→段落、w:tbl→表格、oo:blk→未建模块(blockLvlSdt 等)无损还原。
function blocksFromKids(kids, gaps, rels) {
  const out = [];
  for (const c of kids || []) {
    if (c.tag === "w:p") out.push(paraToJson(c, gaps, rels));
    else if (c.tag === "w:tbl") out.push(tableToJson(c, gaps, rels));
    else if (c.tag === "oo:blk") { try { out.push(JSON.parse(c.attrs["oo:v"])); } catch { /* 兜底损坏则跳过 */ } }
    else if (c.tag && c.tag !== "w:sectPr" && c.tag !== "w:tcPr" && c.tag !== "w:trPr" && c.tag !== "w:tblPr" && c.tag !== "w:tblGrid") gaps && gaps.add("blk>" + c.tag);
  }
  return out;
}

// ── 表格 ──
const BDR_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV", "start", "end", "between", "bar"];
function xmlBdrEl(node) {
  const o = { color: mergeAutoRgb(colorToJson(node.attrs["w:color"]), node.attrs["oo:rgb"]), sz: a2i(node.attrs["w:sz"]) || 0, space: a2i(node.attrs["w:space"]) || 0, value: node.attrs["w:val"] };
  return { ...readExtra(node), ...o }; // themeColor 等经 oo:x 兜底
}
function xmlBorders(node) {
  if (!node) return undefined;
  const o = {};
  for (const c of node.kids) { const s = (c.tag || "").replace(/^w:/, ""); if (BDR_SIDES.includes(s)) o[s] = xmlBdrEl(c); }
  return o; // 空元素 → {}（保留"显式无边框"语义，不丢键）
}
function xmlWidth(node) {
  if (!node) return undefined;
  return { type: node.attrs["w:type"] || "auto", w: a2i(node.attrs["w:w"]) || 0 };
}
function xmlMar(node) {
  if (!node) return undefined;
  const o = {};
  for (const c of node.kids) { const s = (c.tag || "").replace(/^w:/, ""); if (["top", "left", "bottom", "right", "start", "end"].includes(s)) o[s] = xmlWidth(c); }
  return Object.keys(o).length ? o : undefined;
}
function tblPrToJson(node, gaps) {
  const o = { type: "tablePr" };
  if (!node) return o;
  const g = (t) => kid(node, t);
  if (g("w:tblStyle")) o.tblStyle = g("w:tblStyle").attrs["w:val"];
  if (g("w:tblW")) o.tblW = xmlWidth(g("w:tblW"));
  if (g("w:jc")) o.jc = g("w:jc").attrs["w:val"];
  if (g("w:tblBorders")) o.tblBorders = xmlBorders(g("w:tblBorders"));
  if (g("w:tblCellMar")) o.tblCellMar = xmlMar(g("w:tblCellMar"));
  if (g("w:tblLayout")) o.tblLayout = g("w:tblLayout").attrs["w:type"];
  if (g("w:tblLook")) { const a = g("w:tblLook").attrs; o.tblLook = {}; for (const k of ["firstColumn", "firstRow", "lastColumn", "lastRow", "noHBand", "noVBand"]) if (a["w:" + k] != null) o.tblLook[k] = a["w:" + k] === "1" || a["w:" + k] === "true"; }
  // tblInd(数字而非宽度对象)/tblOverlap/tblpPr/inline 等 → 由 oo:x 兜底，不在此解析以免写错形态
  return { ...readExtra(node), ...o };
}
function tcPrToJson(node, gaps) {
  const o = { type: "tableCellPr" };
  if (!node) return o;
  const g = (t) => kid(node, t);
  if (g("w:tcW")) o.tcW = xmlWidth(g("w:tcW"));
  if (g("w:gridSpan")) o.gridSpan = a2i(g("w:gridSpan").attrs["w:val"]);
  if (g("w:vMerge")) o.vMerge = g("w:vMerge").attrs["w:val"] || "continue";
  if (g("w:tcBorders")) o.tcBorders = xmlBorders(g("w:tcBorders"));
  if (g("w:shd")) o.shd = xmlShd(g("w:shd"));
  if (g("w:vAlign")) o.vAlign = g("w:vAlign").attrs["w:val"];
  return { ...readExtra(node), ...o };
}
function trPrToJson(node, gaps) {
  const o = { type: "tableRowPr" };
  if (!node) return o;
  const g = (t) => kid(node, t);
  if (g("w:tblHeader")) o.tblHeader = true;
  if (g("w:jc")) o.jc = g("w:jc").attrs["w:val"];
  if (g("w:trHeight")) o.trHeight = { val: a2i(g("w:trHeight").attrs["w:val"]), hRule: g("w:trHeight").attrs["w:hRule"] };
  return { ...readExtra(node), ...o };
}
let _cellId = 1;
function cellToJson(tc, gaps, rels) {
  const blocks = blocksFromKids(tc.kids, gaps, rels);
  const inner = blocks.length ? blocks : [paraToJson({ tag: "w:p", attrs: {}, kids: [] }, gaps, rels)];
  const ex = readExtra(tc);
  // cell.content 是 DocumentContent 壳；id 优先用兜底里的原值，缺失才用计数器（agent 新增单元格）
  return { ...ex, content: { bPresentation: false, content: inner, type: "docContent" }, tcPr: tcPrToJson(kid(tc, "w:tcPr"), gaps), id: ex.id != null ? ex.id : String(_cellId++), type: "tblCell" };
}
function rowToJson(tr, gaps, rels) {
  const ex = readExtra(tr);
  return {
    ...ex,
    content: kidsOf(tr, "w:tc").map((tc) => cellToJson(tc, gaps, rels)),
    reviewInfo: ex.reviewInfo || { userId: "", author: "", date: "", moveType: "noMove", prevType: -1 },
    reviewType: ex.reviewType || "common",
    trPr: trPrToJson(kid(tr, "w:trPr"), gaps),
    type: "tblRow",
  };
}
function tableToJson(tbl, gaps, rels) {
  const tblGrid = kidsOf(kid(tbl, "w:tblGrid"), "w:gridCol").map((g) => ({ w: a2i(g.attrs["w:w"]), type: "gridCol" }));
  const ex = readExtra(tbl);
  return {
    ...ex,
    bPresentation: false,
    tblGrid,
    tblPr: tblPrToJson(kid(tbl, "w:tblPr"), gaps),
    content: kidsOf(tbl, "w:tr").map((tr) => rowToJson(tr, gaps, rels)),
    changes: ex.changes || [],
    type: "table",
  };
}

/**
 * OOXML(document.xml) → ONLYOFFICE-JSON 文档对象
 * @returns { doc, gaps }
 */
export function ooxmlToJson(xml, gaps = new Gaps(), options = {}) {
  _cellId = 1;
  const docNode = parseXml(xml);
  const body = kid(docNode, "w:body");
  const empty = { type: "document", paraPr: { bFromDocument: true, type: "paraPr" }, content: [] };
  if (!body) return { doc: empty, gaps };
  const content = blocksFromKids(body.kids, gaps, options.relationships || options.rels);
  const ex = readExtra(body); // 文档级 paraPr 等
  return { doc: { type: "document", paraPr: { bFromDocument: true, type: "paraPr" }, ...ex, content }, gaps };
}
