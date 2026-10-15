/**
 * 诊2：模拟 OnlyOffice 原生 ToJSON 输出（只有 commentMetadata，body 无 commentRangeStart）。
 * 验证新的 collectComments 能否从元数据直接收集批注并插入 body 锚点。
 */
import { jsonToFlatOpcWithStyles } from "../src/flat-opc.mjs";
import { withPackageLongTailParts } from "../src/long-tail-parts.mjs";

const nativeDoc = {
  type: "document",
  paraPr: { bFromDocument: true, type: "paraPr" },
  commentMetadata: [
    {
      id: "1",
      text: "这段需要修改",
      author: "审稿人",
      quote: "需要修改的部分",
      time: Date.UTC(2026, 6, 2, 10, 0, 0),
      solved: false,
    },
  ],
  content: [
    { type: "paragraph", pPr: { pStyle: "Normal" }, rPr: { bFromDocument: true, type: "textPr" },
      content: [ { type: "run", rPr: { bFromDocument: true, type: "textPr" }, content: ["这是正文"] } ] },
    { type: "paragraph", pPr: { pStyle: "Normal" }, rPr: { bFromDocument: true, type: "textPr" },
      content: [ { type: "run", rPr: { bFromDocument: true, type: "textPr" }, content: ["前面是", "需要修改的部分", "，后面还有"] } ] },
    { type: "paragraph", pPr: { pStyle: "Normal" }, rPr: { bFromDocument: true, type: "textPr" },
      content: [ { type: "run", rPr: { bFromDocument: true, type: "textPr" }, content: ["最后一段"] } ] },
  ],
};

console.log("=== 诊2: OnlyOffice 原生 ToJSON（仅 commentMetadata）===");

const result = withPackageLongTailParts(nativeDoc);
console.log("comments:", result.summary.comments);
console.assert(result.summary.comments === 1, "应 1 条批注");

let hasStart = false, hasEnd = false;
for (const block of result.doc.content) {
  if (block.type === "paragraph" && Array.isArray(block.content)) {
    for (const item of block.content) {
      if (item.type === "commentRangeStart") hasStart = true;
      if (item.type === "commentRangeEnd") hasEnd = true;
    }
  }
}
console.log("body 插入 commentRangeStart:", hasStart);
console.log("body 插入 commentRangeEnd:", hasEnd);

const flat = jsonToFlatOpcWithStyles(result.doc, [], undefined, undefined, { extraParts: result.extraParts });
console.log("Flat OPC has comments.xml:", flat.includes('pkg:name="/word/comments.xml"'));
console.log("Flat OPC has w:commentRangeStart:", flat.includes("<w:commentRangeStart"));

const cm = flat.match(/<pkg:part\b[^>]*pkg:name="\/word\/comments\.xml"[^>]*>[\s\S]*?<\/pkg:part>/);
if (cm) {
  const a = cm[0].match(/w:author="([^"]*)"/);
  const t = cm[0].match(/<w:t[^>]*>([^<]*)<\/w:t>/);
  console.log("author:", a ? a[1] : "?");
  console.log("text:", t ? t[1] : "?");
}

const ok = result.summary.comments === 1 && hasStart && hasEnd;
console.log("\n结果:", ok ? "✅ 元数据路径正常工作" : "❌ 有问题");