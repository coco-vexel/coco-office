// Tool 4: assemble oojson output (document.xml + package-level parts) into
// Flat OPC for SDK(@coco-vexel/coco-office) fromFlatOpc.
import { jsonToOoxml } from "./json-to-ooxml.mjs";
import { buildStylesXml } from "./styles-to-ooxml.mjs";
import { esc, attrStr } from "./xml.mjs";

const CT = {
  document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  webSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  contentTypes: "application/vnd.openxmlformats-package.content-types+xml",
};
const REL = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  fontTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  webSettings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
};

function pkgPart(name, contentType, innerXml) {
  const inner = innerXml.replace(/^<\?xml[^>]*\?>/, "");
  return (
    `<pkg:part pkg:name="${esc(name)}" pkg:contentType="${esc(contentType)}">` +
    `<pkg:xmlData>${inner}</pkg:xmlData></pkg:part>`
  );
}

function pkgBinaryPart(name, contentType, data) {
  return (
    `<pkg:part pkg:name="${esc(name)}" pkg:contentType="${esc(contentType)}" pkg:compression="store">` +
    `<pkg:binaryData>${String(data).replace(/\s+/g, "")}</pkg:binaryData></pkg:part>`
  );
}

function relXml({ id, type, target, targetMode }) {
  const mode = targetMode ? ` TargetMode="${esc(targetMode)}"` : "";
  return `<Relationship Id="${esc(id)}" Type="${esc(type)}" Target="${esc(target)}"${mode}/>`;
}

function relsPartNameForSource(source) {
  const clean = String(source || "").replace(/^\//, "");
  const slash = clean.lastIndexOf("/");
  if (slash === -1) return `/_rels/${clean}.rels`;
  return `/${clean.slice(0, slash)}/_rels/${clean.slice(slash + 1)}.rels`;
}

function relationshipTargetForPart(source, targetPart) {
  if (!targetPart.startsWith("/")) return targetPart;
  const sourceDir = String(source || "").replace(/^\//, "").split("/").slice(0, -1);
  const targetParts = targetPart.replace(/^\//, "").split("/");
  while (sourceDir.length && targetParts[0] === sourceDir[0]) {
    sourceDir.shift();
    targetParts.shift();
  }
  return "../".repeat(sourceDir.length) + targetParts.join("/");
}

function normalizeExtraPart(part, fallbackIndex) {
  if (!part || typeof part !== "object") return null;
  const name = String(part.name ?? "").trim();
  const contentType = String(part.contentType ?? "").trim();
  if (!name || !contentType) return null;
  return {
    name: name.startsWith("/") ? name : `/${name}`,
    contentType,
    xml: part.xml,
    binary: part.binary,
    rel: part.rel ? {
      source: part.rel.source ?? "/word/document.xml",
      id: part.rel.id ?? `rIdExtra${fallbackIndex}`,
      type: part.rel.type,
      target: part.rel.target,
      targetMode: part.rel.targetMode,
    } : undefined,
  };
}

function isExternalHyperlink(value) {
  return /^(https?:|mailto:|ftp:)/i.test(String(value || ""));
}

function withDocumentHyperlinkRels(node, rels, state = { index: 1 }) {
  if (Array.isArray(node)) return node.map((item) => withDocumentHyperlinkRels(item, rels, state));
  if (!node || typeof node !== "object") return node;
  const next = { ...node };
  if (next.type === "hyperlink" && isExternalHyperlink(next.value) && !next.rId && !next.id && !next["r:id"]) {
    const id = `rIdCocoHyperlink${state.index++}`;
    next.rId = id;
    rels.push(relXml({ id, type: REL.hyperlink, target: String(next.value), targetMode: "External" }));
  }
  if (Array.isArray(next.content)) next.content = withDocumentHyperlinkRels(next.content, rels, state);
  if (next.content && typeof next.content === "object" && !Array.isArray(next.content)) next.content = withDocumentHyperlinkRels(next.content, rels, state);
  return next;
}

function twoHex(n) {
  const v = Number.isFinite(Number(n)) ? Math.max(0, Math.min(255, Number(n))) : 0;
  return Math.round(v).toString(16).toUpperCase().padStart(2, "0");
}

function rgbHex(rgba, fallback) {
  if (!rgba || typeof rgba !== "object") return fallback;
  return twoHex(rgba.red) + twoHex(rgba.green) + twoHex(rgba.blue);
}

function themeColorXml(slot, value, fallback) {
  const color = value?.color || value;
  if (color?.type === "sys") {
    const lastClr = rgbHex(color.rgba, fallback.lastClr || fallback.val || "000000");
    return `<a:${slot}><a:sysClr val="${esc(color.id || fallback.sys || "windowText")}" lastClr="${esc(lastClr)}"/></a:${slot}>`;
  }
  return `<a:${slot}><a:srgbClr val="${esc(rgbHex(color?.rgba, fallback.val))}"/></a:${slot}>`;
}

function fontName(theme, group, kind, fallback) {
  return theme?.themeElements?.fontScheme?.[group]?.[kind] || fallback;
}

function fontSchemeXml(theme) {
  const scheme = theme?.themeElements?.fontScheme;
  const name = scheme?.name || "Office";
  const majorLatin = fontName(theme, "majorFont", "latin", "Cambria");
  const majorEa = fontName(theme, "majorFont", "ea", "宋体");
  const majorCs = fontName(theme, "majorFont", "cs", "Times New Roman");
  const minorLatin = fontName(theme, "minorFont", "latin", "Calibri");
  const minorEa = fontName(theme, "minorFont", "ea", "宋体");
  const minorCs = fontName(theme, "minorFont", "cs", "Arial");
  return (
    `<a:fontScheme name="${esc(name)}">` +
    `<a:majorFont><a:latin typeface="${esc(majorLatin)}"/><a:ea typeface="${esc(majorEa)}"/><a:cs typeface="${esc(majorCs)}"/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${esc(minorLatin)}"/><a:ea typeface="${esc(minorEa)}"/><a:cs typeface="${esc(minorCs)}"/></a:minorFont>` +
    `</a:fontScheme>`
  );
}

function solidFillXml(fill, fallback = "phClr") {
  const solid = fill?.solidFill ?? fill;
  const color = solid?.schemeClr ?? solid?.schemeColor ?? solid?.color;
  if (typeof color === "string") return `<a:solidFill><a:schemeClr val="${esc(color)}"/></a:solidFill>`;
  if (color?.val || color?.id) return `<a:solidFill><a:schemeClr val="${esc(color.val || color.id)}"/></a:solidFill>`;
  const srgb = solid?.srgbClr ?? solid?.rgb ?? rgbHex(color?.rgba, null);
  if (srgb) return `<a:solidFill><a:srgbClr val="${esc(String(srgb).replace(/^#/, "").toUpperCase())}"/></a:solidFill>`;
  return `<a:solidFill><a:schemeClr val="${esc(fallback)}"/></a:solidFill>`;
}

function colorChoiceXml(value, fallback = "phClr") {
  const color = value?.schemeClr ?? value?.schemeColor ?? value?.color ?? value;
  if (typeof color === "string") return `<a:schemeClr val="${esc(color)}"/>`;
  if (color?.val || color?.id) return `<a:schemeClr val="${esc(color.val || color.id)}"/>`;
  const srgb = value?.srgbClr ?? value?.rgb ?? rgbHex(color?.rgba ?? value?.rgba, null);
  if (srgb) return `<a:srgbClr val="${esc(String(srgb).replace(/^#/, "").toUpperCase())}"/>`;
  return `<a:schemeClr val="${esc(fallback)}"/>`;
}

function shadowXml(shadow) {
  if (!shadow || typeof shadow !== "object") return "";
  const attrs = {};
  for (const key of ["blurRad", "dist", "dir", "sx", "sy", "kx", "ky", "algn", "rotWithShape"]) if (shadow[key] != null) attrs[key] = shadow[key];
  return `<a:outerShdw${attrStr(attrs)}>${colorChoiceXml(shadow.solidFill ?? shadow.fill ?? shadow.color)}</a:outerShdw>`;
}

function effectStyleListXml(list) {
  const items = Array.isArray(list) ? list : list ? [list] : [];
  const styles = (items.length ? items : [{ effectLst: {} }]).map((style) => {
    const effectLst = style.effectLst ?? style.effects ?? style;
    const outer = effectLst.outerShdw ?? effectLst.outerShadow;
    const effects = shadowXml(outer);
    return `<a:effectStyle><a:effectLst>${effects}</a:effectLst></a:effectStyle>`;
  }).join("");
  return `<a:effectStyleLst>${styles}</a:effectStyleLst>`;
}

function fillStyleListXml(tag, list) {
  const items = Array.isArray(list) ? list : list ? [list] : [];
  const fills = (items.length ? items : [{ solidFill: { schemeClr: "phClr" } }]).map((fill) => solidFillXml(fill)).join("");
  return `<a:${tag}>${fills}</a:${tag}>`;
}

function lineStyleListXml(list) {
  const items = Array.isArray(list) ? list : list ? [list] : [];
  const lines = (items.length ? items : [{ w: 9525, cap: "flat", cmpd: "sng", algn: "ctr", solidFill: { schemeClr: "phClr" } }]).map((line) => {
    const attrs = { w: line.w ?? line.width ?? 9525, cap: line.cap ?? "flat", cmpd: line.cmpd ?? "sng", algn: line.algn ?? line.align ?? "ctr" };
    return `<a:ln${attrStr(attrs)}>${solidFillXml(line.solidFill ? line : line.fill ?? line)}</a:ln>`;
  }).join("");
  return `<a:lnStyleLst>${lines}</a:lnStyleLst>`;
}
function fmtSchemeXml(theme) {
  const fmt = theme?.themeElements?.fmtScheme;
  const name = fmt?.name || "Office";
  return (
    `<a:fmtScheme name="${esc(name)}">` +
    fillStyleListXml("fillStyleLst", fmt?.fillStyleLst ?? fmt?.fillStyles) +
    lineStyleListXml(fmt?.lnStyleLst ?? fmt?.lineStyleLst ?? fmt?.lineStyles) +
    effectStyleListXml(fmt?.effectStyleLst ?? fmt?.effectStyles) +
    fillStyleListXml("bgFillStyleLst", fmt?.bgFillStyleLst ?? fmt?.bgFillStyles) +
    `</a:fmtScheme>`
  );
}
export function buildThemeXml(theme) {
  const clr = theme?.themeElements?.clrScheme;
  const name = theme?.name || "Office Theme";
  const clrName = clr?.name || "Office";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${esc(name)}">` +
    `<a:themeElements>` +
    `<a:clrScheme name="${esc(clrName)}">` +
    themeColorXml("dk1", clr?.dk1, { sys: "windowText", lastClr: "000000" }) +
    themeColorXml("lt1", clr?.lt1, { sys: "window", lastClr: "FFFFFF" }) +
    themeColorXml("dk2", clr?.dk2, { val: "1F497D" }) +
    themeColorXml("lt2", clr?.lt2, { val: "EEECE1" }) +
    themeColorXml("accent1", clr?.accent1, { val: "4F81BD" }) +
    themeColorXml("accent2", clr?.accent2, { val: "C0504D" }) +
    themeColorXml("accent3", clr?.accent3, { val: "9BBB59" }) +
    themeColorXml("accent4", clr?.accent4, { val: "8064A2" }) +
    themeColorXml("accent5", clr?.accent5, { val: "4BACC6" }) +
    themeColorXml("accent6", clr?.accent6, { val: "F79646" }) +
    themeColorXml("hlink", clr?.hlink, { val: "0000FF" }) +
    themeColorXml("folHlink", clr?.folHlink, { val: "800080" }) +
    `</a:clrScheme>` +
    fontSchemeXml(theme) +
    fmtSchemeXml(theme) +
    `</a:themeElements></a:theme>`
  );
}

export function buildFontTableXml(theme) {
  const fonts = new Set([
    "Calibri",
    "Cambria",
    "Arial",
    "Times New Roman",
    "宋体",
    fontName(theme, "majorFont", "latin", null),
    fontName(theme, "majorFont", "ea", null),
    fontName(theme, "majorFont", "cs", null),
    fontName(theme, "minorFont", "latin", null),
    fontName(theme, "minorFont", "ea", null),
    fontName(theme, "minorFont", "cs", null),
  ].filter(Boolean));
  const entries = [...fonts].map((name) => {
    const charset = /[\u3400-\u9fff]/.test(name) ? "86" : "00";
    const family = /Times|Cambria|宋体/i.test(name) ? "roman" : "swiss";
    return `<w:font w:name="${esc(name)}"><w:charset w:val="${charset}"/><w:family w:val="${family}"/></w:font>`;
  }).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    entries +
    `</w:fonts>`
  );
}
export function buildSettingsXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/>` +
    `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
    `</w:settings>`
  );
}

export function buildWebSettingsXml() {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:optimizeForBrowser/>` +
    `</w:webSettings>`
  );
}

export function buildDefaultWordParts(ooJsonDoc) {
  const theme = ooJsonDoc?.theme;
  return [
    { name: "/word/theme/theme1.xml", contentType: CT.theme, xml: buildThemeXml(theme), rel: { type: REL.theme, target: "theme/theme1.xml" } },
    { name: "/word/fontTable.xml", contentType: CT.fontTable, xml: buildFontTableXml(theme), rel: { type: REL.fontTable, target: "fontTable.xml" } },
    { name: "/word/settings.xml", contentType: CT.settings, xml: buildSettingsXml(), rel: { type: REL.settings, target: "settings.xml" } },
    { name: "/word/webSettings.xml", contentType: CT.webSettings, xml: buildWebSettingsXml(), rel: { type: REL.webSettings, target: "webSettings.xml" } },
  ];
}

/**
 * oojson output -> Flat OPC package string (SDK fromFlatOpc can consume it).
 *
 * @param {object} ooJsonDoc ONLYOFFICE-JSON document object (ToJSON output)
 * @param {object} [opts]
 * @param {string} [opts.stylesXml] Built styles.xml.
 * @param {string} [opts.numberingXml] Built numbering.xml.
 * @param {boolean} [opts.includeDefaultParts=true] Add theme/fontTable/settings/webSettings.
 * @param {Array} [opts.extraParts] Extra package parts for headers/footers/comments/media and future long tails.
 * @returns {string} Flat OPC package
 */
export function jsonToFlatOpc(ooJsonDoc, opts = {}) {
  const generatedDocRels = [];
  const docForXml = withDocumentHyperlinkRels(ooJsonDoc, generatedDocRels);
  const documentXml = jsonToOoxml(docForXml);
  const { stylesXml, numberingXml } = opts;
  const includeDefaultParts = opts.includeDefaultParts !== false;
  const extraParts = [
    ...(includeDefaultParts ? buildDefaultWordParts(ooJsonDoc) : []),
    ...(opts.extraParts ?? []),
  ].map((part, i) => normalizeExtraPart(part, i + 1)).filter(Boolean);

  const overrides = [
    `<Override PartName="/word/document.xml" ContentType="${CT.document}"/>`,
  ];
  if (stylesXml) overrides.push(`<Override PartName="/word/styles.xml" ContentType="${CT.styles}"/>`);
  if (numberingXml) overrides.push(`<Override PartName="/word/numbering.xml" ContentType="${CT.numbering}"/>`);
  for (const part of extraParts) {
    overrides.push(`<Override PartName="${esc(part.name)}" ContentType="${esc(part.contentType)}"/>`);
  }
  const contentTypesXml =
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="${CT.rels}"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    overrides.join("") +
    `</Types>`;

  const rootRelsXml =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>` +
    `</Relationships>`;

  const docRels = [];
  let rid = 1;
  if (stylesXml) docRels.push(`<Relationship Id="rId${rid++}" Type="${REL.styles}" Target="styles.xml"/>`);
  if (numberingXml) docRels.push(`<Relationship Id="rId${rid++}" Type="${REL.numbering}" Target="numbering.xml"/>`);
  docRels.push(...generatedDocRels);
  const relGroups = new Map();
  for (const part of extraParts) {
    if (!part.rel || !part.rel.type) continue;
    const source = part.rel.source || "/word/document.xml";
    const target = part.rel.target ?? relationshipTargetForPart(source, part.name);
    const entry = relXml({ id: part.rel.id ?? `rId${rid}`, type: part.rel.type, target, targetMode: part.rel.targetMode });
    if (source === "/word/document.xml") {
      docRels.push(entry);
      rid++;
    } else {
      const list = relGroups.get(source) ?? [];
      list.push(entry);
      relGroups.set(source, list);
    }
  }
  const docRelsXml =
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    docRels.join("") +
    `</Relationships>`;

  const parts = [
    pkgPart("/[Content_Types].xml", CT.contentTypes, contentTypesXml),
    pkgPart("/_rels/.rels", CT.rels, rootRelsXml),
    pkgPart("/word/_rels/document.xml.rels", CT.rels, docRelsXml),
  ];
  for (const [source, rels] of relGroups.entries()) {
    const relsXml =
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      rels.join("") +
      `</Relationships>`;
    parts.push(pkgPart(relsPartNameForSource(source), CT.rels, relsXml));
  }
  if (stylesXml) parts.push(pkgPart("/word/styles.xml", CT.styles, stylesXml));
  if (numberingXml) parts.push(pkgPart("/word/numbering.xml", CT.numbering, numberingXml));
  for (const part of extraParts) {
    if (part.xml != null) parts.push(pkgPart(part.name, part.contentType, String(part.xml)));
    else if (part.binary != null) parts.push(pkgBinaryPart(part.name, part.contentType, part.binary));
  }
  parts.push(pkgPart("/word/document.xml", CT.document, documentXml));

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<?mso-application progid="Word.Document"?>` +
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    parts.join("") +
    `</pkg:package>`
  );
}

/**
 * Convenience wrapper: ONLYOFFICE-JSON document + collected style definitions -> Flat OPC package.
 * styleDefs shape is documented in styles-to-ooxml.buildStylesXml.
 */
export function jsonToFlatOpcWithStyles(ooJsonDoc, styleDefs, docDefaults, numberingXml, opts = {}) {
  const stylesXml = styleDefs && styleDefs.length ? buildStylesXml(styleDefs, docDefaults) : undefined;
  return jsonToFlatOpc(ooJsonDoc, { ...opts, stylesXml, numberingXml });
}
