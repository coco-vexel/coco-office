// Tool 3: ONLYOFFICE style definitions -> OOXML styles.xml.
// Stable source: Api.GetDocument().ToJSON(false, true, true, true, true, true).
// The sixth argument exports the native styles map keyed by ONLYOFFICE numeric styleId.
import { esc, attrStr } from "./xml.mjs";
import { colorToOoxml, autoRgbAttr, highlightToOoxml } from "./color.mjs";
import { extraAttr } from "./extra.mjs";

function shdXml(shd) {
  if (!shd) return "";
  const fill = shd.fill ? ` w:fill="${colorToOoxml(shd.fill)}"${autoRgbAttr("oo:frgb", shd.fill)}` : "";
  return `<w:shd w:val="${shd.val || "clear"}" w:color="${colorToOoxml(shd.color)}"${autoRgbAttr("oo:crgb", shd.color)}${fill}${extraAttr(shd, ["val", "color", "fill"])}/>`;
}

function bdrElXml(tag, b) {
  if (!b) return "";
  const a = {};
  const v = b.val != null ? b.val : b.value;
  if (v != null) a["w:val"] = v;
  if (b.sz != null) a["w:sz"] = b.sz;
  if (b.space != null) a["w:space"] = b.space;
  if (b.color) a["w:color"] = colorToOoxml(b.color);
  return `<w:${tag}${attrStr(a)}${b.color ? autoRgbAttr("oo:rgb", b.color) : ""}${extraAttr(b, ["color", "sz", "space", "value", "val"])}/>`;
}

const BDR_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV", "start", "end", "between", "bar", "tl2br", "tr2bl"];
function bordersXml(tag, o) {
  if (!o) return "";
  let x = "";
  for (const side of BDR_SIDES) if (o[side]) x += bdrElXml(side, o[side]);
  return Object.keys(o).length ? `<w:${tag}>${x}</w:${tag}>` : `<w:${tag}/>`;
}

function widthXml(tag, w) {
  if (!w) return "";
  const a = {};
  if (w.type != null) a["w:type"] = w.type;
  const v = w.w != null ? w.w : w.value;
  if (v != null) a["w:w"] = v;
  return `<w:${tag}${attrStr(a)}${extraAttr(w, ["type", "w", "value"])}/>`;
}

function marXml(tag, m) {
  if (!m) return "";
  let x = "";
  for (const side of ["top", "left", "bottom", "right", "start", "end"]) if (m[side]) x += widthXml(side, m[side]);
  return x ? `<w:${tag}>${x}</w:${tag}>` : "";
}

function valOf(v) {
  return typeof v === "object" && v !== null ? v.val ?? v.value ?? v.type : v;
}

export function ooOutlineToOoxml(lvl) {
  if (lvl == null) return null;
  const n = Number(lvl);
  if (!Number.isFinite(n)) return null;
  const z = n - 1;
  if (z < 0) return 0;
  if (z > 9) return 9;
  return z;
}

function numPrXml(numPr) {
  if (!numPr) return "";
  let x = "";
  const ilvl = valOf(numPr.ilvl);
  const numId = valOf(numPr.numId);
  if (ilvl != null) x += `<w:ilvl w:val="${esc(String(ilvl))}"/>`;
  if (numId != null) x += `<w:numId w:val="${esc(String(numId))}"/>`;
  return x || extraAttr(numPr, ["ilvl", "numId"]) ? `<w:numPr${extraAttr(numPr, ["ilvl", "numId"])}>${x}</w:numPr>` : "";
}

function tabsXml(tabs) {
  if (!tabs) return "";
  const list = Array.isArray(tabs) ? tabs : Array.isArray(tabs.tabs) ? tabs.tabs : Array.isArray(tabs.list) ? tabs.list : [];
  let x = "";
  for (const tab of list) {
    if (!tab) continue;
    const a = {};
    const val = tab.val ?? tab.value ?? tab.type;
    if (val != null) a["w:val"] = val;
    if (tab.pos != null) a["w:pos"] = tab.pos;
    if (tab.leader != null) a["w:leader"] = tab.leader;
    x += `<w:tab${attrStr(a)}${extraAttr(tab, ["val", "value", "type", "pos", "leader"])}/>`;
  }
  return x ? `<w:tabs>${x}</w:tabs>` : "";
}

function ppInnerXml(pPr, options = {}) {
  if (!pPr) return "";
  let x = "";
  if (options.includeStyleRef && pPr.pStyle != null) x += `<w:pStyle w:val="${esc(String(pPr.pStyle))}"/>`;
  for (const t of ["keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing", "suppressLineNumbers"])
    if (pPr[t] === true) x += `<w:${t}/>`;
  x += numPrXml(pPr.numPr);
  x += tabsXml(pPr.tabs);
  x += bordersXml("pBdr", pPr.pBdr);
  x += shdXml(pPr.shd);
  if (pPr.spacing) {
    const s = pPr.spacing, a = {};
    if (s.before != null) a["w:before"] = s.before;
    if (s.after != null) a["w:after"] = s.after;
    if (s.line != null) a["w:line"] = s.line;
    if (s.lineRule != null) a["w:lineRule"] = s.lineRule;
    if (s.beforeAutoSpacing != null) a["w:beforeAutospacing"] = s.beforeAutoSpacing === "off" ? 0 : 1;
    if (s.afterAutoSpacing != null) a["w:afterAutospacing"] = s.afterAutoSpacing === "off" ? 0 : 1;
    x += `<w:spacing${attrStr(a)}${extraAttr(s, ["before", "after", "line", "lineRule", "beforeAutoSpacing", "afterAutoSpacing"])}/>`;
  }
  if (pPr.ind) {
    const d = pPr.ind, a = {};
    for (const k of ["left", "right", "firstLine", "hanging", "start", "end", "firstLineChars", "hangingChars"])
      if (d[k] != null) a["w:" + k] = d[k];
    x += `<w:ind${attrStr(a)}${extraAttr(d, Object.keys(a).map((k) => k.slice(2)))}/>`;
  }
  if (pPr.jc != null) x += `<w:jc w:val="${esc(String(pPr.jc))}"/>`;
  if (pPr.textAlignment != null) x += `<w:textAlignment w:val="${esc(String(valOf(pPr.textAlignment)))}"/>`;
  if (pPr.bidi === true) x += `<w:bidi/>`;
  const olv = ooOutlineToOoxml(pPr.outlineLvl);
  if (olv != null) x += `<w:outlineLvl w:val="${olv}"/>`;
  return x;
}

export function stylePpToXml(pPr) {
  const inner = ppInnerXml(pPr);
  const ex = pPr ? extraAttr(pPr, ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing", "suppressLineNumbers", "numPr", "tabs", "pBdr", "shd", "spacing", "ind", "jc", "textAlignment", "bidi", "outlineLvl", "type", "bFromDocument"]) : "";
  return inner || ex ? `<w:pPr${ex}>${inner}</w:pPr>` : "";
}

function underlineXml(u) {
  if (u == null) return "";
  if (u === false) return `<w:u w:val="none"/>`;
  if (u === true) return `<w:u w:val="single"/>`;
  if (typeof u !== "object") return `<w:u w:val="${esc(String(u))}"/>`;
  const a = {};
  const val = u.val ?? u.value ?? u.type;
  if (val != null) a["w:val"] = val;
  if (u.color) a["w:color"] = colorToOoxml(u.color);
  return `<w:u${attrStr(a)}${u.color ? autoRgbAttr("oo:rgb", u.color) : ""}${extraAttr(u, ["val", "value", "type", "color"])}/>`;
}

function rpInnerXml(rPr) {
  if (!rPr) return "";
  let x = "";
  if (rPr.rFonts) {
    const f = rPr.rFonts, a = {};
    for (const k of ["ascii", "hAnsi", "eastAsia", "cs", "hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"])
      if (f[k] != null) a["w:" + k] = f[k];
    x += `<w:rFonts${attrStr(a)}${extraAttr(f, Object.keys(a).map((k) => k.slice(2)))}/>`;
  }
  const tgl = (name, v) => (v === true ? `<w:${name}/>` : v === false ? `<w:${name} w:val="false"/>` : "");
  for (const name of ["b", "bCs", "i", "iCs", "strike", "dstrike", "caps", "smallCaps", "vanish", "rtl", "noProof"])
    x += tgl(name, rPr[name]);
  x += underlineXml(rPr.u);
  if (rPr.color) x += `<w:color w:val="${colorToOoxml(rPr.color)}"${autoRgbAttr("oo:rgb", rPr.color)}/>`;
  if (rPr.spacing != null) x += `<w:spacing w:val="${rPr.spacing}"/>`;
  if (rPr.position != null) x += `<w:position w:val="${rPr.position}"/>`;
  if (rPr.kern != null) x += `<w:kern w:val="${rPr.kern}"/>`;
  if (rPr.sz != null) x += `<w:sz w:val="${rPr.sz}"/>`;
  if (rPr.szCs != null) x += `<w:szCs w:val="${rPr.szCs}"/>`;
  if (rPr.highlight) {
    const h = highlightToOoxml(rPr.highlight);
    if (h) x += `<w:highlight w:val="${h}"/>`;
  }
  if (rPr.vertAlign != null) x += `<w:vertAlign w:val="${esc(String(valOf(rPr.vertAlign)))}"/>`;
  x += shdXml(rPr.shd);
  if (rPr.lang) {
    const a = {};
    for (const k of ["val", "eastAsia", "bidi"]) if (rPr.lang[k] != null) a["w:" + k] = rPr.lang[k];
    x += `<w:lang${attrStr(a)}${extraAttr(rPr.lang, ["val", "eastAsia", "bidi"])}/>`;
  }
  return x;
}

export function styleRpToXml(rPr) {
  const inner = rpInnerXml(rPr);
  const ex = rPr ? extraAttr(rPr, ["rFonts", "b", "bCs", "i", "iCs", "strike", "dstrike", "caps", "smallCaps", "vanish", "rtl", "noProof", "u", "color", "spacing", "position", "kern", "sz", "szCs", "highlight", "vertAlign", "shd", "lang", "type", "bFromDocument"]) : "";
  return inner || ex ? `<w:rPr${ex}>${inner}</w:rPr>` : "";
}

function tblPrXml(t) {
  if (!t) return "";
  let x = "";
  if (t.tblStyle != null) x += `<w:tblStyle w:val="${esc(String(t.tblStyle))}"/>`;
  x += widthXml("tblW", t.tblW);
  if (t.jc != null) x += `<w:jc w:val="${esc(String(t.jc))}"/>`;
  x += bordersXml("tblBorders", t.tblBorders);
  x += shdXml(t.shd);
  x += marXml("tblCellMar", t.tblCellMar);
  if (t.tblLayout != null) x += `<w:tblLayout w:type="${esc(String(valOf(t.tblLayout)))}"/>`;
  if (t.tblLook != null && typeof t.tblLook === "object") {
    const a = {};
    for (const k of Object.keys(t.tblLook)) if (k !== "type") a["w:" + k] = t.tblLook[k] === true ? 1 : t.tblLook[k] === false ? 0 : t.tblLook[k];
    x += `<w:tblLook${attrStr(a)}/>`;
  }
  const ex = extraAttr(t, ["tblStyle", "tblW", "jc", "tblBorders", "shd", "tblCellMar", "tblLayout", "tblLook", "type"]);
  return x || ex ? `<w:tblPr${ex}>${x}</w:tblPr>` : "";
}

function trPrXml(p) {
  if (!p) return "";
  let x = "";
  if (p.tblHeader === true) x += `<w:tblHeader/>`;
  if (p.cantSplit === true) x += `<w:cantSplit/>`;
  if (p.jc != null) x += `<w:jc w:val="${esc(String(p.jc))}"/>`;
  if (p.trHeight) x += `<w:trHeight${attrStr({ "w:val": p.trHeight.value != null ? p.trHeight.value : p.trHeight.val, "w:hRule": p.trHeight.hRule })}${extraAttr(p.trHeight, ["value", "val", "hRule"])}/>`;
  const ex = extraAttr(p, ["tblHeader", "cantSplit", "jc", "trHeight", "type"]);
  return x || ex ? `<w:trPr${ex}>${x}</w:trPr>` : "";
}

function tcPrXml(p) {
  if (!p) return "";
  let x = "";
  x += widthXml("tcW", p.tcW);
  if (p.gridSpan != null) x += `<w:gridSpan w:val="${p.gridSpan}"/>`;
  if (p.vMerge != null) x += `<w:vMerge${p.vMerge && p.vMerge !== "continue" ? ` w:val="${esc(String(p.vMerge))}"` : ""}/>`;
  x += bordersXml("tcBorders", p.tcBorders);
  x += shdXml(p.shd);
  x += marXml("tcMar", p.tcMar);
  if (p.vAlign != null) x += `<w:vAlign w:val="${esc(String(valOf(p.vAlign)))}"/>`;
  if (p.noWrap === true) x += `<w:noWrap/>`;
  const ex = extraAttr(p, ["tcW", "gridSpan", "vMerge", "tcBorders", "shd", "tcMar", "vAlign", "noWrap", "type"]);
  return x || ex ? `<w:tcPr${ex}>${x}</w:tcPr>` : "";
}

function styleType(t) {
  switch (t) {
    case "paragraphStyle": return "paragraph";
    case "characterStyle": return "character";
    case "tableStyle": return "table";
    case "numberingStyle": return "numbering";
    case "character": case "table": case "numbering": case "paragraph": return t;
    default: return "paragraph";
  }
}

function tableStyleType(t) {
  switch (t) {
    case "topLeftCell": return "nwCell";
    case "topRightCell": return "neCell";
    case "bottomLeftCell": return "swCell";
    case "bottomRightCell": return "seCell";
    default: return t;
  }
}

function tableStylePrXml(type, pr) {
  if (!pr) return "";
  const wType = tableStyleType(pr.styleType ?? type);
  if (!wType) return "";
  const inner = stylePpToXml(pr.pPr) + styleRpToXml(pr.rPr) + tblPrXml(pr.tblPr) + trPrXml(pr.trPr) + tcPrXml(pr.tcPr);
  const ex = extraAttr(pr, ["styleType", "pPr", "rPr", "tblPr", "trPr", "tcPr", "type"]);
  return inner || ex ? `<w:tblStylePr w:type="${esc(String(wType))}"${ex}>${inner}</w:tblStylePr>` : "";
}

function tableStylePrListXml(tblStylePr) {
  if (!tblStylePr || typeof tblStylePr !== "object") return "";
  let x = "";
  for (const key of Object.keys(tblStylePr)) x += tableStylePrXml(key, tblStylePr[key]);
  return x;
}

export function styleToXml(styleId, s) {
  if (!s || styleId == null) return "";
  const type = styleType(s.styleType ?? s.type);
  const attrs = { "w:type": type, "w:styleId": String(styleId) };
  if (s.default === true) attrs["w:default"] = "1";
  if (s.customStyle === true) attrs["w:customStyle"] = "1";
  let inner = "";
  if (s.name != null) inner += `<w:name w:val="${esc(String(s.name))}"/>`;
  if (s.basedOn != null && s.basedOn !== "") inner += `<w:basedOn w:val="${esc(String(s.basedOn))}"/>`;
  if (s.next != null && s.next !== "") inner += `<w:next w:val="${esc(String(s.next))}"/>`;
  if (s.link != null && s.link !== "") inner += `<w:link w:val="${esc(String(s.link))}"/>`;
  if (s.hidden === true) inner += `<w:hidden/>`;
  if (s.uiPriority != null) inner += `<w:uiPriority w:val="${s.uiPriority}"/>`;
  if (s.semiHidden === true) inner += `<w:semiHidden/>`;
  if (s.unhideWhenUsed === true) inner += `<w:unhideWhenUsed/>`;
  if (s.qFormat === true) inner += `<w:qFormat/>`;
  inner += stylePpToXml(s.pPr);
  inner += styleRpToXml(s.rPr);
  inner += tblPrXml(s.tblPr);
  inner += trPrXml(s.trPr);
  inner += tcPrXml(s.tcPr);
  inner += tableStylePrListXml(s.tblStylePr);
  return `<w:style${attrStr(attrs)}>${inner}</w:style>`;
}

export function buildStylesXml(styles, docDefaults) {
  let dd = "";
  if (docDefaults && (docDefaults.paraPr || docDefaults.textPr)) {
    const rpr = styleRpToXml(docDefaults.textPr);
    const ppr = stylePpToXml(docDefaults.paraPr);
    dd = `<w:docDefaults><w:rPrDefault>${rpr}</w:rPrDefault><w:pPrDefault>${ppr}</w:pPrDefault></w:docDefaults>`;
  }
  const body = (styles || []).map((it) => styleToXml(it.styleId, it.def)).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:oo="urn:coco:oojson">` +
    dd +
    body +
    `</w:styles>`
  );
}
