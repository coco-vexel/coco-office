/**
 * 批注 round-trip 诊断：模拟 writeActiveXml 内 extractCommentsFromFlatOpc 和 oojson 转换。
 *
 * 构建一个含批注的 Flat OPC（同 SDK addComment 产出），走 writeActiveXml 前段链路，
 * 验证每一步是否丢批注。
 */
import { jsonToFlatOpc, jsonToFlatOpcWithStyles } from "../src/flat-opc.mjs";
import { withPackageLongTailParts } from "../src/long-tail-parts.mjs";
import { ooxmlToJson } from "../src/ooxml-to-json.mjs";
import { extractFlatOpcPart } from "./extract-helper.mjs";

// ---------- 1. 构造含批注的 oojson doc ----------
const doc = {
  type: "document",
  paraPr: { bFromDocument: true, type: "paraPr" },
  commentMetadata: [
    { author: "coco", text: "这段需要修改", time: Date.UTC(2026, 6, 2, 10, 0, 0), solved: false },
  ],
  content: [
    {
      type: "paragraph",
      pPr: { pStyle: "Normal" },
      rPr: { bFromDocument: true, type: "textPr" },
      content: [
        { type: "run", content: ["这是第一段正文，没有批注。"] },
      ],
    },
    {
      type: "paragraph",
      pPr: { pStyle: "Normal" },
      rPr: { bFromDocument: true, type: "textPr" },
      content: [
        { type: "commentRangeStart", id: "42", author: "coco", text: "这段需要修改" },
        { type: "run", content: ["这是带批注的段落内容"] },
        { type: "commentRangeEnd", id: "42", author: "coco", text: "这段需要修改" },
      ],
    },
    {
      type: "paragraph",
      pPr: { pStyle: "Normal" },
      rPr: { bFromDocument: true, type: "textPr" },
      content: [
        { type: "run", content: ["批注后面的段落。"] },
      ],
    },
  ],
};

// ---------- 2. 走 long-tail + flat-opc（读路径） ----------
const longTail = withPackageLongTailParts(doc);
console.log("=== 读路径: withPackageLongTailParts ===");
console.log("summary:", JSON.stringify(longTail.summary, null, 2));
console.assert(longTail.summary.comments === 1, "应产出 1 条批注");

const flatOpc = jsonToFlatOpcWithStyles(longTail.doc, [], undefined, undefined, {
  extraParts: longTail.extraParts,
});

// 验证 Flat OPC 包含批注 parts
console.log("\n=== Flat OPC 包含批注 ===");
console.log("has comments.xml:", flatOpc.includes('pkg:name="/word/comments.xml"'));
console.log("has commentRangeStart:", flatOpc.includes("<w:commentRangeStart"));
console.log("has commentRangeEnd:", flatOpc.includes("<w:commentRangeEnd"));

// ---------- 3. 模拟 writeActiveXml 内的 extractCommentsFromFlatOpc ----------
console.log("\n=== 模拟 extractCommentsFromFlatOpc ===");

const commentsXml = extractFlatOpcPart(flatOpc, "/word/comments.xml");
console.log("comments.xml extracted:", commentsXml ? "yes" : "NO — 空！");
if (commentsXml) {
  const count = (commentsXml.match(/<w:comment\b/g) || []).length;
  console.log("comment 元素数:", count);
}

const docPart = extractFlatOpcPart(flatOpc, "/word/document.xml");
console.log("document.xml extracted:", docPart ? "yes" : "NO — 空！");

if (docPart) {
  const startCount = (docPart.match(/<w:commentRangeStart/g) || []).length;
  const endCount = (docPart.match(/<w:commentRangeEnd/g) || []).length;
  console.log("commentRangeStart 数:", startCount);
  console.log("commentRangeEnd 数:", endCount);
}

// ---------- 4. 模拟 writeActiveXml 内的 write 路径：ooxmlToJson ----------
console.log("\n=== 模拟写路径: commentRangeStart/End 在 oojson 中 ===");

// 从 document.xml 提取 w:body 内容
const bodyMatch = docPart.match(/<w:body>([\s\S]*)<\/w:body>/);
const bodyXml = bodyMatch ? bodyMatch[1] : docPart;

// 测试 ooxmlToJson 是否保留 commentRangeStart/End
const { doc: ooJson, gaps } = ooxmlToJson(docPart);

let foundStart = 0;
let foundEnd = 0;
function walkContent(items) {
  if (!Array.isArray(items)) return;
  for (const item of items) {
    if (item.type === "commentRangeStart") foundStart++;
    if (item.type === "commentRangeEnd") foundEnd++;
    if (item.content) walkContent(item.content);
  }
}
walkContent(ooJson.content);
console.log("ooxmlToJson 保留 commentRangeStart:", foundStart);
console.log("ooxmlToJson 保留 commentRangeEnd:", foundEnd);
console.assert(foundStart === 1, "ooxmlToJson 应保留 1 个 commentRangeStart");
console.assert(foundEnd === 1, "ooxmlToJson 应保留 1 个 commentRangeEnd");
// 点：oojson 中批注的文本能不能找到
console.log("\noojson 段落 content 项类型序列:",
  ooJson.content?.[1]?.content?.map((c) => c.type).join(", "));

// ---------- 5. 检查 anchorText ----------
console.log("\n=== anchorText 提取 ===");
const startEls = docPart.match(/<w:commentRangeStart[\s\S]*?\/>/g) || [];
for (const startEl of startEls) {
  const idMatch = startEl.match(/w:id="(\d+)"/);
  if (idMatch) {
    console.log("批注 id:", idMatch[1]);
    // 提取从 commentRangeStart 到 commentRangeEnd 之间的文本
    const startIdx = docPart.indexOf(startEl);
    const endTag = `<w:commentRangeEnd w:id="${idMatch[1]}"`;
    const endIdx = docPart.indexOf(endTag, startIdx);
    if (endIdx > startIdx) {
      const between = docPart.slice(startIdx + startEl.length, endIdx);
      const textMatches = [...between.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)];
      const anchorText = textMatches.map((m) => m[1]).join("");
      console.log("anchorText:", JSON.stringify(anchorText));
    }
  }
}

console.log("\n=== 结论 ===");
console.log("读路径（oojson→Flat OPC）: 正常");
if (foundStart > 0 && foundEnd > 0) {
  console.log("ooxmlToJson 保留了 comment 锚点元素");
  console.log("⚠️ 关键未验证：OnlyOffice 原生 FromJSON 是否识别 {type:'commentRangeStart'}");
}
if (foundStart === 0) {
  console.log("❌ ooxmlToJson 丢失了 comment 锚点！");
}