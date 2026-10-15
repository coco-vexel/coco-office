/**
 * Helper: extractFlatOpcPart — 从 Flat OPC 字符串中提取指定 part 的 XML 内容。
 * 与 onlyoffice.ts 中的 extractFlatOpcPart 同逻辑，供诊断脚本使用。
 */
export function extractFlatOpcPart(flatOpc, partName) {
  const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(
    `<pkg:part\\b(?=[^>]*pkg:name="${esc(partName)}")[\\s\\S]*?<pkg:xmlData>([\\s\\S]*?)</pkg:xmlData></pkg:part>`,
  );
  const m = re.exec(flatOpc);
  return m ? m[1].trim() : null;
}