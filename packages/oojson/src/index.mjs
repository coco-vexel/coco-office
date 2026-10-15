// @coco-vexel/coco-office-oojson —— ONLYOFFICE-JSON ↔ OOXML(WordprocessingML body) 双向转换器，公共入口。
//
// 管道里的位置（coco-onlyoffice 宿主，复用 coco-office 的 OOXML 能力）：
//   编辑器 ToJSON → ONLYOFFICE-JSON ──工具① jsonToOoxml──▶ OOXML → agent 改
//   agent 改完的 OOXML ──工具② ooxmlToJson──▶ ONLYOFFICE-JSON → FromJSON 整篇灌回编辑器
//
// 无损保证：json→ooxml→json 在 484 篇真实文档上裸回环(零归一)零差异。
// 原理见 ./extra.mjs 与 README.md（双层：忠实建模 + 私有 oo:x 兜底）。

// 工具①：ONLYOFFICE-JSON → OOXML(document.xml)
export { jsonToOoxml } from "./json-to-ooxml.mjs";
// 工具②：OOXML → ONLYOFFICE-JSON（+ Gaps 缺口追踪）
export { ooxmlToJson, Gaps } from "./ooxml-to-json.mjs";
// 工具③：ONLYOFFICE 样式定义集 → styles.xml（重建 styleId 定义，让 core 判标题层级）
export { buildStylesXml, styleToXml, ooOutlineToOoxml } from "./styles-to-ooxml.mjs";
export { buildNumberingXml } from "./numbering-to-ooxml.mjs";
// 工具④：document.xml [+ styles/numbering] → Flat OPC 整包（喂 SDK fromFlatOpc）
export { buildDefaultWordParts, buildFontTableXml, buildSettingsXml, buildThemeXml, buildWebSettingsXml, jsonToFlatOpc, jsonToFlatOpcWithStyles } from "./flat-opc.mjs";
export { withPackageLongTailParts } from "./long-tail-parts.mjs";
export { flatOpcToOnlyOfficeWriteback } from "./writeback.mjs";
export {
  applyPptTextMutationToOnlyOfficeContent,
  onlyOfficePptContentToTextBody,
  replaceOnlyOfficePptContentText,
  setOnlyOfficePptContentParagraphs,
  setOnlyOfficePptContentRunStyle,
  setOnlyOfficePptContentStyle,
  setOnlyOfficePptContentText,
} from "./ppt-content.mjs";
export {
  applyPptModelToOnlyOfficePptDeck,
  onlyOfficePptDeckToPptModel,
  updateOnlyOfficePptDeckTextShape,
} from "./ppt-deck.mjs";

// 进阶/复用：底层件
export { parseXml, kid, kidsOf, esc, unesc, attrStr } from "./xml.mjs";
export { colorToOoxml, colorToJson, highlightToOoxml, highlightToJson } from "./color.mjs";
export { extraAttr, readExtra, onlyKeys } from "./extra.mjs";
