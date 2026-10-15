// 迷你 XML 解析/序列化（无依赖）。用于 WordprocessingML 片段。
// 解析结果节点：{ tag, attrs:{}, kids:[node | {text}] }；文本节点为 { text }。

export function parseXml(xml) {
  let i = 0;
  const n = xml.length;
  function node() {
    while (i < n && /\s/.test(xml[i])) i++; // 跳节点前空白（含 <?xml?> 后换行）
    if (xml[i] !== "<") return null;
    if (xml.startsWith("<?", i) || xml.startsWith("<!", i)) {
      i = xml.indexOf(">", i) + 1;
      return node();
    }
    i++; // <
    let tag = "";
    while (i < n && !/[\s/>]/.test(xml[i])) tag += xml[i++];
    const attrs = {};
    while (i < n && xml[i] !== ">" && xml[i] !== "/") {
      while (/\s/.test(xml[i])) i++;
      if (xml[i] === "/" || xml[i] === ">") break;
      let name = "";
      while (i < n && !/[\s=/>]/.test(xml[i])) name += xml[i++];
      while (/\s/.test(xml[i])) i++;
      let val = "";
      if (xml[i] === "=") {
        i++;
        while (/\s/.test(xml[i])) i++;
        const q = xml[i++];
        while (i < n && xml[i] !== q) val += xml[i++];
        i++;
      }
      attrs[name] = unesc(val);
    }
    if (xml[i] === "/") { i += 2; return { tag, attrs, kids: [] }; }
    i++; // >
    const kids = [];
    while (i < n) {
      if (xml.startsWith("</", i)) { i = xml.indexOf(">", i) + 1; break; }
      if (xml[i] === "<") { const c = node(); if (c) kids.push(c); }
      else { let t = ""; while (i < n && xml[i] !== "<") t += xml[i++]; if (t) kids.push({ text: unesc(t) }); }
    }
    return { tag, attrs, kids };
  }
  while (i < n && xml[i] !== "<") i++;
  return node();
}

export const kid = (node, tag) => (node && node.kids ? node.kids.find((k) => k.tag === tag) : undefined);
export const kidsOf = (node, tag) => (node && node.kids ? node.kids.filter((k) => k.tag === tag) : []);

export const esc = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
export const unesc = (s) =>
  String(s)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

// 把属性对象拼成 ' k="v" ...' （已转义）
export function attrStr(attrs) {
  let s = "";
  for (const k of Object.keys(attrs)) if (attrs[k] != null) s += ` ${k}="${esc(attrs[k])}"`;
  return s;
}
