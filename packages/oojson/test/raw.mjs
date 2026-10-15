// 裸回环对比（无任何 norm 归一）：量化 json→ooxml→json 的真实数据差。
import fs from "node:fs"; import path from "node:path";
import { jsonToOoxml } from "../src/json-to-ooxml.mjs";
import { ooxmlToJson, Gaps } from "../src/ooxml-to-json.mjs";
function diff(a, b, p = "", out = []) {
  if (a === b) return out;
  const ta = a === null ? "null" : Array.isArray(a) ? "array" : typeof a;
  const tb = b === null ? "null" : Array.isArray(b) ? "array" : typeof b;
  if (ta !== tb) { out.push({ path: p, kind: "changed", a, b }); return out; }
  if (ta === "array") { const n = Math.max(a.length, b.length); for (let i = 0; i < n; i++) { if (i >= a.length) out.push({ path: `${p}[${i}]`, kind: "extra", b: b[i] }); else if (i >= b.length) out.push({ path: `${p}[${i}]`, kind: "missing", a: a[i] }); else diff(a[i], b[i], `${p}[${i}]`, out); } }
  else if (ta === "object") { for (const k of Object.keys(a)) { if (!(k in b)) out.push({ path: `${p}.${k}`, kind: "missing", a: a[k] }); else diff(a[k], b[k], `${p}.${k}`, out); } for (const k of Object.keys(b)) if (!(k in a)) out.push({ path: `${p}.${k}`, kind: "extra", b: b[k] }); }
  else if (a !== b) out.push({ path: p, kind: "changed", a, b });
  return out;
}
const dir = process.argv[2] || "test/captured";
const files = fs.readdirSync(dir).filter((e) => /\.oojson\.json$/i.test(e)).map((e) => path.join(dir, e));
const agg = new Map(); let total = 0; let perfect = 0;
for (const f of files) {
  const orig = JSON.parse(fs.readFileSync(f, "utf8"));
  const { doc: back } = ooxmlToJson(jsonToOoxml(orig), new Gaps());
  const ds = diff(orig, back, "doc"); total += ds.length; if (!ds.length) perfect++; // 比整篇文档对象，非仅 .content
  for (const d of ds) { const k = d.kind + " " + d.path.replace(/\[\d+\]/g, "[]"); if (!agg.has(k)) agg.set(k, { c: 0, a: d.a, b: d.b }); agg.get(k).c++; }
}
console.log(`裸回环（无 norm）: ${perfect}/${files.length} 完全一致, 差异总数 ${total}`);
[...agg.entries()].sort((x, y) => y[1].c - x[1].c).forEach(([k, v]) => {
  console.log(`\n  ${String(v.c).padStart(4)}  ${k}`);
  if (v.a !== undefined) console.log("        a=", JSON.stringify(v.a).slice(0, 140));
  if (v.b !== undefined) console.log("        b=", JSON.stringify(v.b).slice(0, 140));
});
