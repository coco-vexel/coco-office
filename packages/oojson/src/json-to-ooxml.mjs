// 工具①：ONLYOFFICE-JSON → OOXML(WordprocessingML body)。
// 读链路用：编辑器 ToJSON 出 ONLYOFFICE-JSON → 本工具转 OOXML → 喂 agent。
// 无损保证：每个属性袋/元素把未建模的剩余键塞进 oo:x 兜底属性（见 extra.mjs）。
import { esc, attrStr } from "./xml.mjs";
import { colorToOoxml, highlightToOoxml, autoRgbAttr } from "./color.mjs";
import { extraAttr, onlyKeys } from "./extra.mjs";

function shdXml(shd) {
  if (!shd) return "";
  const fill = shd.fill ? ` w:fill="${colorToOoxml(shd.fill)}"${autoRgbAttr("oo:frgb", shd.fill)}` : "";
  return `<w:shd w:val="${shd.val || "clear"}" w:color="${colorToOoxml(shd.color)}"${autoRgbAttr("oo:crgb", shd.color)}${fill}${extraAttr(shd, ["val", "color", "fill"])}/>`;
}
function ppToXml(pPr) {
  if (!pPr) return "";
  let x = "";
  if (pPr.pStyle != null) x += `<w:pStyle w:val="${esc(pPr.pStyle)}"/>`;
  for (const t of ["keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing", "suppressLineNumbers"])
    if (pPr[t] === true) x += `<w:${t}/>`; // false 由兜底还原（编辑器恒输出的默认键）
  x += shdXml(pPr.shd);
  if (pPr.spacing) {
    const s = pPr.spacing, a = {};
    if (s.before != null) a["w:before"] = s.before;
    if (s.after != null) a["w:after"] = s.after;
    if (s.line != null) a["w:line"] = s.line;
    if (s.lineRule != null) a["w:lineRule"] = s.lineRule;
    if (s.beforeAutoSpacing != null) a["w:beforeAutospacing"] = s.beforeAutoSpacing === "off" ? 0 : 1;
    if (s.afterAutoSpacing != null) a["w:afterAutospacing"] = s.afterAutoSpacing === "off" ? 0 : 1;
    x += `<w:spacing${attrStr(a)}/>`;
  }
  if (pPr.ind) {
    const d = pPr.ind, a = {};
    if (d.left != null) a["w:left"] = d.left;
    if (d.right != null) a["w:right"] = d.right;
    if (d.firstLine != null) a["w:firstLine"] = d.firstLine;
    x += `<w:ind${attrStr(a)}/>`;
  }
  if (pPr.jc != null) x += `<w:jc w:val="${pPr.jc}"/>`;
  return x; // 返回 pPr 内层（不含 <w:pPr> 包裹），由 paraToXml 拼上段落 mark rPr + oo:x 后统一包裹
}
function rpInner(rPr) {
  let x = "";
  if (rPr.rFonts) {
    const f = rPr.rFonts, a = {};
    for (const k of ["ascii", "hAnsi", "eastAsia", "cs", "hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"]) if (f[k] != null) a["w:" + k] = f[k];
    x += `<w:rFonts${attrStr(a)}/>`;
  }
  const tgl = (name, v) => (v === true ? `<w:${name}/>` : v === false ? `<w:${name} w:val="false"/>` : "");
  x += tgl("b", rPr.b) + tgl("bCs", rPr.bCs) + tgl("i", rPr.i) + tgl("iCs", rPr.iCs);
  if (rPr.color) x += `<w:color w:val="${colorToOoxml(rPr.color)}"${autoRgbAttr("oo:rgb", rPr.color)}/>`;
  if (rPr.spacing != null) x += `<w:spacing w:val="${rPr.spacing}"/>`;
  if (rPr.sz != null) x += `<w:sz w:val="${rPr.sz}"/>`;
  if (rPr.szCs != null) x += `<w:szCs w:val="${rPr.szCs}"/>`;
  if (rPr.highlight) { const h = highlightToOoxml(rPr.highlight); if (h) x += `<w:highlight w:val="${h}"/>`; }
  if (rPr.vertAlign != null) x += `<w:vertAlign w:val="${typeof rPr.vertAlign === "object" ? rPr.vertAlign.val || rPr.vertAlign.value : rPr.vertAlign}"/>`;
  x += shdXml(rPr.shd);
  if (rPr.lang) {
    const a = {};
    for (const k of ["val", "eastAsia", "bidi"]) if (rPr.lang[k] != null) a["w:" + k] = rPr.lang[k];
    x += `<w:lang${attrStr(a)}/>`;
  }
  return x;
}
function rpToXml(rPr) {
  if (!rPr) return "";
  const inner = rpInner(rPr);
  const ex = extraAttr(rPr);
  return inner || ex ? `<w:rPr${ex}>${inner}</w:rPr>` : "";
}
// run 内容项：文本/分页符/制表符 → 真 OOXML；其余（drawing 等）→ oo:ci 无损兜底。
function drawingDocPrAttrs(d) {
  const docPr = d?.docPr || {};
  const attrs = { id: docPr.id ?? 1, name: docPr.name ?? "Picture" };
  if (docPr.descr != null) attrs.descr = docPr.descr;
  if (docPr.title != null) attrs.title = docPr.title;
  return attrStr(attrs);
}

function srcRectXml(blipFill) {
  const rect = blipFill?.srcRect;
  if (!rect || typeof rect !== "object") return "";
  const attrs = {};
  for (const key of ["l", "t", "r", "b"]) if (rect[key] != null) attrs[key] = rect[key];
  return Object.keys(attrs).length ? `<a:srcRect${attrStr(attrs)}/>` : "";
}

function presetGeometryXml(spPr) {
  const raw =
    spPr?.prstGeom?.prst ??
    spPr?.custGeom?.preset ??
    spPr?.geometry?.prst ??
    spPr?.geometry?.preset ??
    spPr?.prstGeom;
  const prst = typeof raw === "string" ? raw : "rect";
  return `<a:prstGeom prst="${esc(prst)}"><a:avLst/></a:prstGeom>`;
}

function effectExtentXml(effectExtent) {
  if (!effectExtent || typeof effectExtent !== "object") return "";
  const attrs = {};
  for (const key of ["l", "t", "r", "b"]) if (effectExtent[key] != null) attrs[key] = effectExtent[key];
  return Object.keys(attrs).length ? `<wp:effectExtent${attrStr(attrs)}/>` : "";
}

function drawingPositionXml(tag, pos, fallbackRelativeFrom) {
  if (!pos || typeof pos !== "object") return `<wp:${tag} relativeFrom="${fallbackRelativeFrom}"><wp:posOffset>0</wp:posOffset></wp:${tag}>`;
  const relativeFrom = pos.relativeFrom || fallbackRelativeFrom;
  const child = pos.align != null ? `<wp:align>${esc(pos.align)}</wp:align>` : `<wp:posOffset>${esc(pos.posOffset ?? pos.offset ?? 0)}</wp:posOffset>`;
  return `<wp:${tag} relativeFrom="${esc(relativeFrom)}">${child}</wp:${tag}>`;
}

function anchorWrapXml(d) {
  const type = d?.wrapType || d?.wrap?.type || d?.wrappingType;
  const wrapText = d?.wrapText || d?.wrap?.wrapText || "bothSides";
  if (type === "none") return "<wp:wrapNone/>";
  if (type === "topAndBottom" || type === "topBottom") return "<wp:wrapTopAndBottom/>";
  if (type === "tight") return `<wp:wrapTight wrapText="${esc(wrapText)}"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`;
  if (type === "through") {
    const edited = boolAttr(d?.wrapThrough?.edited ?? d?.wrap?.edited ?? false);
    return `<wp:wrapThrough wrapText="${esc(wrapText)}"><wp:wrapPolygon edited="${esc(edited)}"><wp:start x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>`;
  }
  return `<wp:wrapSquare wrapText="${esc(wrapText)}"/>`;
}

function drawingContainerXml(d, cx, cy, docPrAttrs, graphicXml) {
  const distAttrs = { distT: d?.distT ?? 0, distB: d?.distB ?? 0, distL: d?.distL ?? 0, distR: d?.distR ?? 0 };
  if (d?.drawingType !== "anchor" && d?.inline !== false && d?.anchor !== true) {
    return (
      `<wp:inline${attrStr(distAttrs)}>` +
      `<wp:extent cx="${esc(cx)}" cy="${esc(cy)}"/>` +
      effectExtentXml(d?.effectExtent) +
      `<wp:docPr${docPrAttrs}/>` +
      graphicXml +
      `</wp:inline>`
    );
  }
  const simplePos = d?.simplePos && typeof d.simplePos === "object" ? d.simplePos : {};
  const anchorAttrs = {
    ...distAttrs,
    simplePos: boolAttr(d?.simplePos === true || d?.simplePos?.enabled === true),
    relativeHeight: d?.relativeHeight ?? 0,
    behindDoc: boolAttr(d?.behindDoc ?? false),
    locked: boolAttr(d?.locked ?? false),
    layoutInCell: boolAttr(d?.layoutInCell ?? true),
    allowOverlap: boolAttr(d?.allowOverlap ?? true),
  };
  return (
    `<wp:anchor${attrStr(anchorAttrs)}>` +
    `<wp:simplePos x="${esc(simplePos.x ?? 0)}" y="${esc(simplePos.y ?? 0)}"/>` +
    drawingPositionXml("positionH", d?.positionH, "column") +
    drawingPositionXml("positionV", d?.positionV, "paragraph") +
    `<wp:extent cx="${esc(cx)}" cy="${esc(cy)}"/>` +
    effectExtentXml(d?.effectExtent) +
    anchorWrapXml(d) +
    `<wp:docPr${docPrAttrs}/>` +
    `<wp:cNvGraphicFramePr/>` +
    graphicXml +
    `</wp:anchor>`
  );
}
function drawingXml(d) {
  const graphic = d?.graphic;
  const blipFill = graphic?.blipFill || {};
  const spPr = graphic?.spPr || {};
  const relId = blipFill.rasterImageId;
  const cx = d?.extent?.cx ?? spPr?.xfrm?.ext?.cx ?? 1;
  const cy = d?.extent?.cy ?? spPr?.xfrm?.ext?.cy ?? 1;
  const docPrAttrs = drawingDocPrAttrs(d);
  if (
    (graphic?.type === "shape" || graphic?.type === "connectShape") &&
    !relId
  ) {
    const content = d?.content || graphic?.content;
    const inner = Array.isArray(content?.content) ? content.content : [];
    const text = inner.map(blockPlainText).join("");
    const graphicXml =
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
      `<wps:wsp>${shapeNonVisualXml(graphic)}<wps:spPr>${shapePropertiesXml(d, cx, cy)}</wps:spPr>${drawingTextboxXml(text, "wps")}</wps:wsp>` +
      `</a:graphicData></a:graphic>`;
    return `<w:drawing>${drawingContainerXml(d, cx, cy, docPrAttrs, graphicXml)}</w:drawing>`;
  }
  if (!relId || !String(relId).startsWith("rId")) return null;
  const graphicXml =
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr${docPrAttrs}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${esc(relId)}"/>${srcRectXml(blipFill)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${esc(cx)}" cy="${esc(cy)}"/></a:xfrm>${presetGeometryXml(spPr)}</pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  return `<w:drawing>${drawingContainerXml(d, cx, cy, docPrAttrs, graphicXml)}</w:drawing>`;
}
function noteReferenceXml(tag, item) {
  const custom = item.customMarkFollows === true ? ` w:customMarkFollows="1"` : "";
  return `<w:${tag} w:id="${esc(item.footnote)}"${custom}/>`;
}
function runItemXml(it) {
  if (typeof it === "string") return `<w:t xml:space="preserve">${esc(it)}</w:t>`;
  if (it && it.type === "break" && onlyKeys(it, ["type", "breakType"]) && (it.breakType === "page" || it.breakType === "column"))
    return `<w:br w:type="${it.breakType}"/>`; // 仅 page/column 确定可逆；line/textWrapping 等走 oo:ci 兜底
  if (it && it.type === "tab" && onlyKeys(it, ["type"])) return `<w:tab/>`;
  if (it && it.type === "separator" && onlyKeys(it, ["type"])) return `<w:separator/>`;
  if (it && it.type === "continuationSeparator" && onlyKeys(it, ["type"])) return `<w:continuationSeparator/>`;
  if (it && it.type === "footnoteRef" && it.footnote != null) return noteReferenceXml("footnoteReference", it);
  if (it && it.type === "endnoteRef" && it.footnote != null) return noteReferenceXml("endnoteReference", it);
  if (it && it.type === "paraDrawing") {
    const xml = drawingXml(it);
    if (xml) return xml;
  }
  return `<oo:ci oo:v="${esc(JSON.stringify(it))}"/>`;
}
function runToXml(run) {
  const items = (run.content || []).map(runItemXml).join("");
  const ex = extraAttr(run, ["rPr", "content", "footnotes", "endnotes"]);
  return `<w:r${ex}>${rpToXml(run.rPr)}${items}</w:r>`;
}
function hyperlinkToXml(link) {
  const attrs = [];
  const relId = link.rId || link.id || link["r:id"];
  if (relId) attrs.push(`r:id="${esc(relId)}"`);
  else if (link.value) attrs.push(`w:anchor="${esc(String(link.value).replace(/^#/, ""))}"`);
  if (link.tooltip) attrs.push(`w:tooltip="${esc(link.tooltip)}"`);
  const inner = (link.content || []).map((item) => (item?.type === "run" ? runToXml(item) : `<oo:pi oo:v="${esc(JSON.stringify(item))}"/>`)).join("");
  return `<w:hyperlink${attrs.length ? " " + attrs.join(" ") : ""}>${inner}</w:hyperlink>`;
}
// 段落内容项：run→<w:r>；endRun(段落标记)→不输出(体现在 pPr/rPr)；
// 其它(commentRangeStart/End、hyperlink、math 等)→ oo:pi 按序无损占位。
function paraItemXml(it) {
  if (it.type === "endRun") return "";
  if (it.type === "run") return runToXml(it);
  if (it.type === "hyperlink") return hyperlinkToXml(it);
  if (it.type === "commentRangeStart" && it.id != null) return `<w:commentRangeStart w:id="${esc(it.id)}"/>`;
  if (it.type === "commentRangeEnd" && it.id != null) {
    return `<w:commentRangeEnd w:id="${esc(it.id)}"/><w:r><w:commentReference w:id="${esc(it.id)}"/></w:r>`;
  }
  if (it.type === "commentReference" && it.id != null) return `<w:r><w:commentReference w:id="${esc(it.id)}"/></w:r>`;
  return `<oo:pi oo:v="${esc(JSON.stringify(it))}"/>`;
}
function paraToXml(p) {
  const inner = ppToXml(p.pPr) + rpToXml(p.rPr); // pPr 属性 + 段落 mark rPr（= endRun）
  const ppEx = extraAttr(p.pPr);
  const pPrXml = inner || ppEx ? `<w:pPr${ppEx}>${inner}</w:pPr>` : "";
  const ex = extraAttr(p, ["pPr", "rPr", "content"]);
  // endRun(段落标记)可带修订信息(reviewInfo/reviewType=remove,即段落标记的删除被追踪)，rPr 已在 pPr 里；
  // 其余非默认字段挂到 oo:er 由读侧还原到合成的 endRun。
  const endRun = (p.content || []).find((x) => x.type === "endRun");
  const erEx = endRun ? extraAttr(endRun, ["rPr", "content"], "oo:er") : "";
  return `<w:p${ex}${erEx}>${pPrXml}${(p.content || []).map(paraItemXml).join("")}</w:p>`;
}

// 块级内容数组(body 或单元格)：table/paragraph→真元素；其它(blockLvlSdt 等)→ oo:blk 无损占位。
export function blocksToXml(arr) {
  return (arr || []).map((el) => (el.type === "table" ? tableToXml(el) : el.type === "paragraph" ? paraToXml(el) : `<oo:blk oo:v="${esc(JSON.stringify(el))}"/>`)).join("");
}

// ── 表格 ──
function bdrElXml(tag, b) {
  if (!b) return "";
  const a = {};
  const v = b.val != null ? b.val : b.value; if (v != null) a["w:val"] = v;
  if (b.sz != null) a["w:sz"] = b.sz;
  if (b.space != null) a["w:space"] = b.space;
  if (b.color) a["w:color"] = colorToOoxml(b.color);
  return `<w:${tag}${attrStr(a)}${b.color ? autoRgbAttr("oo:rgb", b.color) : ""}${extraAttr(b, ["color", "sz", "space", "value", "val"])}/>`;
}
const BDR_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV", "start", "end", "between", "bar"];
function bordersXml(tag, o) {
  if (!o) return "";
  let x = "";
  for (const s of BDR_SIDES) if (o[s]) x += bdrElXml(s, o[s]);
  // 空 {} 也要发出包裹元素，让读侧产出 {} 而非缺失
  return Object.keys(o).length ? `<w:${tag}>${x}</w:${tag}>` : `<w:${tag}/>`;
}
function widthXml(tag, w) {
  if (!w) return "";
  const a = {};
  if (w.type != null) a["w:type"] = w.type;
  const v = w.w != null ? w.w : w.value; if (v != null) a["w:w"] = v;
  return `<w:${tag}${attrStr(a)}/>`;
}
function marXml(tag, m) {
  if (!m) return "";
  let x = "";
  for (const s of ["top", "left", "bottom", "right", "start", "end"]) if (m[s]) x += widthXml(s, m[s]);
  return x ? `<w:${tag}>${x}</w:${tag}>` : "";
}
function tblPrXml(t) {
  if (!t) return "<w:tblPr/>";
  let x = "";
  if (t.tblStyle != null) x += `<w:tblStyle w:val="${esc(t.tblStyle)}"/>`;
  x += widthXml("tblW", t.tblW);
  if (t.jc != null) x += `<w:jc w:val="${t.jc}"/>`;
  x += bordersXml("tblBorders", t.tblBorders);
  x += marXml("tblCellMar", t.tblCellMar);
  if (t.tblLayout != null) x += `<w:tblLayout w:type="${typeof t.tblLayout === "object" ? t.tblLayout.type : t.tblLayout}"/>`;
  if (t.tblLook != null && typeof t.tblLook === "object") { const a = {}; for (const k of Object.keys(t.tblLook)) if (k !== "type") a["w:" + k] = t.tblLook[k] === true ? 1 : t.tblLook[k] === false ? 0 : t.tblLook[k]; x += `<w:tblLook${attrStr(a)}/>`; }
  // tblInd/tblOverlap/tblpPr/inline 等 → 走 oo:x 兜底（形态古怪/非 OOXML 标准，建模易错）
  return `<w:tblPr${extraAttr(t)}>${x}</w:tblPr>`;
}
function cellToXml(c) {
  const p = c.tcPr || {};
  let pr = "";
  pr += widthXml("tcW", p.tcW);
  if (p.gridSpan != null) pr += `<w:gridSpan w:val="${p.gridSpan}"/>`;
  if (p.vMerge != null) pr += `<w:vMerge${p.vMerge && p.vMerge !== "continue" ? ` w:val="${p.vMerge}"` : ""}/>`;
  if (p.tcBorders) pr += bordersXml("tcBorders", p.tcBorders);
  pr += shdXml(p.shd);
  if (p.vAlign != null) pr += `<w:vAlign w:val="${typeof p.vAlign === "object" ? p.vAlign.val || p.vAlign.value : p.vAlign}"/>`;
  // tcMar/noWrap 等 → oo:x 兜底
  const prXml = `<w:tcPr${extraAttr(p)}>${pr}</w:tcPr>`;
  // cell.content 是 DocumentContent 壳 {type:"docContent", content:[段落]}
  const cont = c.content && Array.isArray(c.content.content) ? c.content.content : Array.isArray(c.content) ? c.content : [];
  const body = blocksToXml(cont) || "<w:p/>";
  const ex = extraAttr(c, ["tcPr", "content"]); // id 也进兜底，保内部计数器精确还原
  return `<w:tc${ex}>${prXml}${body}</w:tc>`;
}
function rowToXml(r) {
  const p = r.trPr || {};
  let pr = "";
  if (p.tblHeader) pr += `<w:tblHeader/>`;
  if (p.jc != null) pr += `<w:jc w:val="${p.jc}"/>`;
  if (p.trHeight) pr += `<w:trHeight${attrStr({ "w:val": p.trHeight.value != null ? p.trHeight.value : p.trHeight.val, "w:hRule": p.trHeight.hRule })}/>`;
  const trPrXml = pr || extraAttr(p) ? `<w:trPr${extraAttr(p)}>${pr}</w:trPr>` : "";
  const cells = Array.isArray(r.content) ? r.content : [];
  const ex = extraAttr(r, ["trPr", "content"]);
  return `<w:tr${ex}>${trPrXml}${cells.map(cellToXml).join("")}</w:tr>`;
}
function tableToXml(t) {
  const grid = (t.tblGrid || []).map((g) => `<w:gridCol w:w="${g.w}"/>`).join("");
  const rows = Array.isArray(t.content) ? t.content : [];
  const ex = extraAttr(t, ["tblPr", "tblGrid", "content"]);
  return `<w:tbl${ex}>${tblPrXml(t.tblPr)}<w:tblGrid>${grid}</w:tblGrid>${rows.map(rowToXml).join("")}</w:tbl>`;
}

function blockPlainText(block) {
  const out = [];
  const visit = (value) => {
    if (typeof value === "string") out.push(value);
    else if (Array.isArray(value)) value.forEach(visit);
    else if (value && typeof value === "object") {
      if (value.type === "run" || value.type === "paragraph") visit(value.content);
    }
  };
  visit(block);
  return out.join("");
}

function drawingTextboxXml(text, prefix) {
  if (!text) return "";
  return `<${prefix}:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p></w:txbxContent></${prefix}:txbx>`;
}

function shouldEmitTopLevelDrawing(drawing) {
  return drawing?.graphic?.type !== "image" && !drawing?.graphic?.blipFill?.rasterImageId;
}

function emuToPt(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return null;
  return Math.max(0, n / 12700).toFixed(2).replace(/\.00$/, "");
}

function rgbaColor(value) {
  const rgba =
    value?.color?.color?.rgba ??
    value?.fill?.color?.color?.rgba ??
    value?.fill?.fill?.color?.color?.rgba ??
    value?.rgba ??
    value?.color?.rgba ??
    value?.fill?.color?.rgba ??
    value?.fill?.fill?.color?.rgba;
  if (!rgba || typeof rgba !== "object") return null;
  return { auto: false, r: rgba.red ?? rgba.r ?? 0, g: rgba.green ?? rgba.g ?? 0, b: rgba.blue ?? rgba.b ?? 0 };
}

function fillKind(value) {
  return value?.type ?? value?.fill?.type ?? value?.fill?.fill?.type;
}

function shapeFillXml(fill) {
  if (fillKind(fill) === "noFill") return "<a:noFill/>";
  const color = rgbaColor(fill);
  return color ? `<a:solidFill><a:srgbClr val="${colorToOoxml(color)}"/></a:solidFill>` : "";
}

function shapeLineXml(line) {
  if (!line || typeof line !== "object") return "";
  const attrs = {};
  if (line.w != null) attrs.w = line.w;
  if (line.cap != null) attrs.cap = line.cap;
  if (line.cmpd != null) attrs.cmpd = line.cmpd;
  const fill = shapeFillXml(line.fill);
  const dash = line.prstDash ? `<a:prstDash val="${esc(line.prstDash)}"/>` : "";
  const join = line.lineJoin?.type === "round" ? "<a:round/>" : "";
  return `<a:ln${attrStr(attrs)}>${fill}${dash}${join}</a:ln>`;
}

function shapePropertiesXml(drawing, cx, cy) {
  const spPr = drawing?.graphic?.spPr ?? {};
  const xfrm = spPr.xfrm ?? {};
  const xfrmAttrs = {};
  for (const key of ["flipH", "flipV", "rot"]) {
    if (xfrm[key] != null) xfrmAttrs[key] = boolAttr(xfrm[key]);
  }
  const off = xfrm.off ?? {};
  const ext = xfrm.ext ?? {};
  return (
    `<a:xfrm${attrStr(xfrmAttrs)}>` +
    `<a:off x="${esc(off.x ?? 0)}" y="${esc(off.y ?? 0)}"/>` +
    `<a:ext cx="${esc(ext.cx ?? cx)}" cy="${esc(ext.cy ?? cy)}"/>` +
    `</a:xfrm>` +
    presetGeometryXml(spPr) +
    shapeFillXml(spPr.fill) +
    shapeLineXml(spPr.ln)
  );
}

function shapeNonVisualXml(graphic) {
  if (graphic?.nvCxnSpPr) {
    const locks = graphic.nvCxnSpPr?.cNvCxnSpPr?.cxnSpLocks ?? {};
    return `<wps:cNvCnPr><a:cxnSpLocks${attrStr({
      noChangeShapeType:
        locks.noChangeShapeType == null
          ? undefined
          : boolAttr(locks.noChangeShapeType),
    })}/></wps:cNvCnPr>`;
  }
  return "<wps:cNvSpPr/>";
}

function vmlShapeAttrs(name, cx, cy, spPr) {
  const width = emuToPt(cx);
  const height = emuToPt(cy);
  const attrs = { id: name };
  if (width != null && height != null) attrs.style = `width:${width}pt;height:${height}pt`;
  const fill = spPr?.fill;
  if (fillKind(fill) === "noFill") attrs.filled = "f";
  else {
    const fillColor = rgbaColor(fill);
    if (fillColor) attrs.fillcolor = `#${colorToOoxml(fillColor)}`;
  }
  const ln = spPr?.ln;
  if (fillKind(ln?.fill) === "noFill") attrs.stroked = "f";
  else {
    const strokeColor = rgbaColor(ln?.fill);
    if (strokeColor) attrs.strokecolor = `#${colorToOoxml(strokeColor)}`;
    const strokeWeight = emuToPt(ln?.w);
    if (strokeWeight != null) attrs.strokeweight = `${strokeWeight}pt`;
  }
  return attrStr(attrs);
}
function topLevelDrawingXml(drawing, index) {
  const content = drawing && (drawing.content || drawing.graphic?.content);
  const inner = Array.isArray(content?.content) ? content.content : [];
  const text = inner.map(blockPlainText).join("");
  const id = drawing?.docPr?.id ?? index + 1;
  const name = drawing?.docPr?.name ?? `Drawing ${index + 1}`;
  const cx = drawing?.extent?.cx ?? drawing?.graphic?.spPr?.xfrm?.ext?.cx ?? 1;
  const cy = drawing?.extent?.cy ?? drawing?.graphic?.spPr?.xfrm?.ext?.cy ?? 1;
  const drawingText = drawingTextboxXml(text, "wps");
  const pictText = drawingTextboxXml(text, "v");
  const graphicXml =
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp>${shapeNonVisualXml(drawing?.graphic)}<wps:spPr>${shapePropertiesXml(drawing, cx, cy)}</wps:spPr>${drawingText}</wps:wsp>` +
    `</a:graphicData></a:graphic>`;
  return (
    `<w:p oo:drawingObject="${esc(index)}"><w:r>` +
    `<w:drawing>${drawingContainerXml(drawing, cx, cy, attrStr({ id, name }), graphicXml)}</w:drawing>` +
    `<w:pict><v:shape${vmlShapeAttrs(name, cx, cy, drawing?.graphic?.spPr)}>${pictText}</v:shape></w:pict>` +
    `</w:r></w:p>`
  );
}

function drawingTextBlocksToXml(doc) {
  const drawings = Array.isArray(doc?.drawingObjects) ? doc.drawingObjects : [];
  const bodyDrawingIds = new Set();
  const collect = (value) => {
    if (!value || typeof value !== "object") return;
    if (value.type === "paraDrawing" && value.docPr?.id != null) {
      bodyDrawingIds.add(String(value.docPr.id));
    }
    if (Array.isArray(value)) {
      value.forEach(collect);
      return;
    }
    Object.values(value).forEach(collect);
  };
  collect(doc?.content);
  return drawings
    .filter(
      (drawing) =>
        shouldEmitTopLevelDrawing(drawing) &&
        (drawing?.docPr?.id == null ||
          !bodyDrawingIds.has(String(drawing.docPr.id))),
    )
    .map((drawing, index) => topLevelDrawingXml(drawing, index))
    .join("");
}function boolAttr(v) {
  return v === true ? 1 : v === false ? 0 : v;
}

function sectRefXml(tag, refs) {
  if (!refs || typeof refs !== "object") return "";
  let out = "";
  for (const type of ["first", "default", "even"]) {
    const ref = refs[type];
    if (!ref) continue;
    const id = typeof ref === "string" ? ref : ref.id || ref.rId || ref["r:id"];
    if (!id) continue;
    out += `<w:${tag} w:type="${type}" r:id="${esc(id)}"/>`;
  }
  return out;
}

function sectBorderXml(tag, b) {
  if (!b) return "";
  const a = {};
  const v = b.val != null ? b.val : b.value;
  if (v != null) a["w:val"] = v;
  if (b.sz != null) a["w:sz"] = b.sz;
  if (b.space != null) a["w:space"] = b.space;
  if (b.color) a["w:color"] = colorToOoxml(b.color);
  return `<w:${tag}${attrStr(a)}${b.color ? autoRgbAttr("oo:rgb", b.color) : ""}${extraAttr(b, ["color", "sz", "space", "value", "val"])}/>`;
}

function sectValOf(value) {
  return value && typeof value === "object" ? value.val ?? value.value ?? value.type : value;
}

function notePrChildXml(tag, value) {
  const v = sectValOf(value);
  return v == null ? "" : `<w:${tag} w:val="${esc(String(v))}"/>`;
}

function notePrXml(tag, pr) {
  if (!pr || typeof pr !== "object" || !Object.keys(pr).length) return "";
  let inner = "";
  inner += notePrChildXml("pos", pr.pos ?? pr.position);
  inner += notePrChildXml("numFmt", pr.numFmt);
  inner += notePrChildXml("numStart", pr.numStart ?? pr.start);
  inner += notePrChildXml("numRestart", pr.numRestart ?? pr.restart);
  return `<w:${tag}${extraAttr(pr, ["pos", "position", "numFmt", "numStart", "start", "numRestart", "restart"])}>${inner}</w:${tag}>`;
}
function sectPrXml(sectPr) {
  if (!sectPr || typeof sectPr !== "object") return "";
  let x = "";
  x += sectRefXml("headerReference", sectPr.headerReference);
  x += sectRefXml("footerReference", sectPr.footerReference);
  x += notePrXml("footnotePr", sectPr.footnotePr);
  x += notePrXml("endnotePr", sectPr.endnotePr);
  if (sectPr.type != null) x += `<w:type w:val="${esc(sectPr.type)}"/>`;
  if (sectPr.pgSz) {
    const p = sectPr.pgSz;
    const a = {};
    if (p.w != null) a["w:w"] = p.w;
    if (p.h != null) a["w:h"] = p.h;
    const orient = p.orient ?? p.orinet;
    if (orient != null) a["w:orient"] = orient;
    x += `<w:pgSz${attrStr(a)}${extraAttr(p, ["w", "h", "orient", "orinet"])}/>`;
  }
  if (sectPr.pgMar) {
    const m = sectPr.pgMar;
    const a = {};
    for (const k of ["top", "right", "bottom", "left", "header", "footer", "gutter"]) if (m[k] != null) a["w:" + k] = m[k];
    x += `<w:pgMar${attrStr(a)}${extraAttr(m, ["top", "right", "bottom", "left", "header", "footer", "gutter"])}/>`;
  }
  if (sectPr.pgBorders) {
    const b = sectPr.pgBorders;
    let inner = "";
    for (const side of ["top", "left", "bottom", "right"]) inner += sectBorderXml(side, b[side]);
    const a = {};
    for (const k of ["display", "offsetFrom", "zOrder"]) if (b[k] != null) a["w:" + k] = b[k];
    x += `<w:pgBorders${attrStr(a)}${extraAttr(b, ["top", "left", "bottom", "right", "display", "offsetFrom", "zOrder"])}>${inner}</w:pgBorders>`;
  }
  if (sectPr.cols) {
    const c = sectPr.cols;
    const a = {};
    for (const k of ["equalWidth", "num", "sep", "space"]) if (c[k] != null) a["w:" + k] = boolAttr(c[k]);
    const cols = Array.isArray(c.col) ? c.col.map((col) => {
      const ca = {};
      for (const [k, v] of Object.entries(col)) ca["w:" + k] = boolAttr(v);
      return `<w:col${attrStr(ca)}/>`;
    }).join("") : "";
    x += `<w:cols${attrStr(a)}${extraAttr(c, ["col", "equalWidth", "num", "sep", "space"])}>${cols}</w:cols>`;
  }
  if (sectPr.titlePg) x += `<w:titlePg/>`;
  if (sectPr.pgNumType) {
    const n = sectPr.pgNumType;
    const a = {};
    for (const k of ["start", "fmt", "chapStyle", "chapSep"]) if (n[k] != null && n[k] !== -1) a["w:" + k] = n[k];
    if (Object.keys(a).length || Object.keys(n).some((k) => !["start", "fmt", "chapStyle", "chapSep"].includes(k))) {
      x += `<w:pgNumType${attrStr(a)}${extraAttr(n, ["start", "fmt", "chapStyle", "chapSep"])}/>`;
    }
  }
  if (sectPr.rtlGutter) x += `<w:rtlGutter/>`;
  return `<w:sectPr${extraAttr(sectPr, ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "pgBorders", "cols", "titlePg", "pgNumType", "rtlGutter"])}>${x}</w:sectPr>`;
}
/** ONLYOFFICE-JSON 文档对象 → OOXML(document.xml) 字符串 */
export function jsonToOoxml(doc) {
  const body = drawingTextBlocksToXml(doc) + blocksToXml(doc.content) + sectPrXml(doc.sectPr || doc.sectionPr);
  // 文档级默认属性 paraPr(源自 styles.xml docDefaults,不在 body 里)→ 挂 <w:body> 的 oo:x 无损保留
  const docEx = extraAttr(doc, ["content", "sectPr", "sectionPr", "drawingObjects", "commentMetadata", "theme", "styles", "numbering"]);
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:oo="urn:coco:oojson">` +
    `<w:body${docEx}>${body}</w:body></w:document>`
  );
}
