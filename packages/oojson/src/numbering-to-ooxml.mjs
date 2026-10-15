// Tool 3b: ONLYOFFICE numbering definitions -> OOXML numbering.xml.
// Stable source: Api.GetDocument().ToJSON(false, true, true, true, true, true).
// The fifth argument exports the native numbering map used by paragraph pPr.numPr.
import { esc, attrStr } from "./xml.mjs";
import { extraAttr } from "./extra.mjs";
import { stylePpToXml, styleRpToXml } from "./styles-to-ooxml.mjs";

function valOf(v) {
  return typeof v === "object" && v !== null ? v.val ?? v.value ?? v.type : v;
}

function entriesOf(obj) {
  if (!obj || typeof obj !== "object") return [];
  if (Array.isArray(obj)) return obj.map((v, i) => [String(v?.abstractNumId ?? v?.numId ?? i), v]);
  return Object.keys(obj).map((k) => [k, obj[k]]);
}

function attrVal(obj, keys) {
  for (const key of keys) {
    if (obj && obj[key] != null) return obj[key];
  }
  return undefined;
}

function simpleVal(tag, value) {
  const v = valOf(value);
  return v == null ? "" : `<w:${tag} w:val="${esc(String(v))}"/>`;
}

function lvlToXml(lvl, fallbackIlvl) {
  if (!lvl || typeof lvl !== "object") return "";
  const ilvl = attrVal(lvl, ["ilvl", "level", "lvl"]);
  const attrs = { "w:ilvl": String(ilvl ?? fallbackIlvl ?? 0) };
  let inner = "";
  inner += simpleVal("start", lvl.start);
  inner += simpleVal("numFmt", lvl.numFmt);
  inner += simpleVal("lvlRestart", lvl.lvlRestart ?? lvl.restart);
  inner += simpleVal("pStyle", lvl.pStyle);
  if (lvl.isLgl === true || lvl.legal === true || lvl.isLegal === true) inner += `<w:isLgl/>`;
  inner += simpleVal("suff", lvl.suff);
  inner += simpleVal("lvlPicBulletId", lvl.lvlPicBulletId ?? lvl.picBulletId);
  if (lvl.lvlText != null) inner += `<w:lvlText w:val="${esc(String(lvl.lvlText))}"/>`;
  inner += simpleVal("lvlJc", lvl.lvlJc);
  inner += stylePpToXml(lvl.pPr);
  inner += styleRpToXml(lvl.rPr);
  const ex = extraAttr(lvl, [
    "ilvl", "level", "lvl", "start", "numFmt", "lvlRestart", "restart", "pStyle", "isLgl", "legal", "isLegal", "suff", "lvlPicBulletId", "picBulletId", "lvlText", "lvlJc", "pPr", "rPr", "type",
  ]);
  return `<w:lvl${attrStr(attrs)}${ex}>${inner}</w:lvl>`;
}

function lvlListXml(levels) {
  if (!levels) return "";
  if (Array.isArray(levels)) return levels.map((lvl, i) => lvlToXml(lvl, i)).join("");
  if (typeof levels !== "object") return "";
  return Object.keys(levels).map((key) => lvlToXml(levels[key], Number(key))).join("");
}

function abstractNumToXml(key, def) {
  if (!def || typeof def !== "object") return "";
  const id = def.abstractNumId ?? def.id ?? key;
  if (id == null || id === "") return "";
  let inner = "";
  inner += simpleVal("nsid", def.nsid);
  inner += simpleVal("multiLevelType", def.multiLevelType);
  inner += simpleVal("tmpl", def.tmpl);
  inner += simpleVal("name", def.name);
  inner += simpleVal("styleLink", def.styleLink);
  inner += simpleVal("numStyleLink", def.numStyleLink);
  inner += lvlListXml(def.lvl ?? def.levels ?? def.Ec);
  const ex = extraAttr(def, [
    "abstractNumId", "id", "nsid", "multiLevelType", "tmpl", "name", "styleLink", "numStyleLink", "lvl", "levels", "Ec", "type",
  ]);
  return `<w:abstractNum w:abstractNumId="${esc(String(id))}"${ex}>${inner}</w:abstractNum>`;
}

function lvlOverrideToXml(override, fallbackIlvl) {
  if (!override || typeof override !== "object") return "";
  const ilvl = override.ilvl ?? override.level ?? fallbackIlvl ?? 0;
  let inner = "";
  inner += simpleVal("startOverride", override.startOverride);
  inner += lvlToXml(override.lvl ?? override.levelDef, ilvl);
  const ex = extraAttr(override, ["ilvl", "level", "startOverride", "lvl", "levelDef", "type"]);
  return `<w:lvlOverride w:ilvl="${esc(String(ilvl))}"${ex}>${inner}</w:lvlOverride>`;
}

function lvlOverrideListXml(overrides) {
  if (!overrides) return "";
  if (Array.isArray(overrides)) return overrides.map((o, i) => lvlOverrideToXml(o, i)).join("");
  if (typeof overrides !== "object") return "";
  return Object.keys(overrides).map((key) => lvlOverrideToXml(overrides[key], Number(key))).join("");
}

function numToXml(key, def) {
  if (!def || typeof def !== "object") return "";
  const id = def.numId ?? def.id ?? key;
  if (id == null || id === "") return "";
  const abstractNumId = def.abstractNumId ?? def.abstractId ?? def.numStyleLink;
  let inner = "";
  if (abstractNumId != null) inner += `<w:abstractNumId w:val="${esc(String(abstractNumId))}"/>`;
  inner += lvlOverrideListXml(def.lvlOverride ?? def.overrides);
  const ex = extraAttr(def, ["numId", "id", "abstractNumId", "abstractId", "numStyleLink", "lvlOverride", "overrides", "type"]);
  return `<w:num w:numId="${esc(String(id))}"${ex}>${inner}</w:num>`;
}

export function buildNumberingXml(numbering) {
  if (!numbering || typeof numbering !== "object") return undefined;
  const abstractDefs = numbering.abstractNum ?? numbering.abstractNums;
  const numDefs = numbering.num ?? numbering.nums;
  const abstractXml = entriesOf(abstractDefs).map(([key, def]) => abstractNumToXml(key, def)).join("");
  const numXml = entriesOf(numDefs).map(([key, def]) => numToXml(key, def)).join("");
  if (!abstractXml && !numXml) return undefined;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:oo="urn:coco:oojson">` +
    abstractXml +
    numXml +
    `</w:numbering>`
  );
}