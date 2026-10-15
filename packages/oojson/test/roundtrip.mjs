// 回环一致性测试（正确方向）：ONLYOFFICE-JSON → 工具① jsonToOoxml → OOXML
//   → 工具② ooxmlToJson → JSON' ，深度对比 原始.content vs JSON'.content。
//   用法: node test/roundtrip.mjs [oojson文件或目录] [最多N]
//   默认: 跑 _oo-test/*.oojson.json
import fs from "node:fs";
import path from "node:path";
import { jsonToOoxml } from "../src/json-to-ooxml.mjs";
import { ooxmlToJson, Gaps } from "../src/ooxml-to-json.mjs";

// 深度 diff：返回 [{path, kind:'missing'|'extra'|'changed', a, b}]
function diff(a, b, p = "", out = []) {
  if (a === b) return out;
  const ta = a === null ? "null" : Array.isArray(a) ? "array" : typeof a;
  const tb = b === null ? "null" : Array.isArray(b) ? "array" : typeof b;
  if (ta !== tb) { out.push({ path: p, kind: "changed", a, b }); return out; }
  if (ta === "array") {
    const n = Math.max(a.length, b.length);
    for (let i = 0; i < n; i++) {
      if (i >= a.length) out.push({ path: `${p}[${i}]`, kind: "extra", b: b[i] });
      else if (i >= b.length) out.push({ path: `${p}[${i}]`, kind: "missing", a: a[i] });
      else diff(a[i], b[i], `${p}[${i}]`, out);
    }
  } else if (ta === "object") {
    for (const k of Object.keys(a)) {
      if (!(k in b)) out.push({ path: `${p}.${k}`, kind: "missing", a: a[k] });
      else diff(a[k], b[k], `${p}.${k}`, out);
    }
    for (const k of Object.keys(b)) if (!(k in a)) out.push({ path: `${p}.${k}`, kind: "extra", b: b[k] });
  } else if (a !== b) out.push({ path: p, kind: "changed", a, b });
  return out;
}
// 把 diff 路径归一成"种类"（去掉数组下标），便于聚合
const diffKey = (d) => d.kind + " " + d.path.replace(/\[\d+\]/g, "[]");

// 语义归一：auto:true 的颜色，r/g/b 是 ONLYOFFICE 内部解析值、编辑器渲染时忽略，
// 且 OOXML "auto" 不携带 RGB（不可逆），故比对时剥掉——属语义等价非损耗。
function norm(o) {
  if (Array.isArray(o)) return o.map(norm);
  if (o && typeof o === "object") {
    if (o.auto === true && ("r" in o || "g" in o || "b" in o)) { const { r, g, b, ...rest } = o; return norm(rest); }
    const out = {};
    for (const k of Object.keys(o)) {
      if (k === "id" && o.type === "tblCell") continue; // 单元格 id 是 ONLYOFFICE 内部计数器，不可复现
      out[k] = norm(o[k]);
    }
    return out;
  }
  return o;
}

function listFiles(target, max) {
  let files = [];
  if (fs.statSync(target).isDirectory()) {
    for (const e of fs.readdirSync(target)) if (/\.oojson\.json$/i.test(e)) files.push(path.join(target, e));
  } else files = [target];
  return files.slice(0, max);
}

const target = process.argv[2] || "test/captured";
const MAX = parseInt(process.argv[3] || "9999", 10);
const files = listFiles(target, MAX);
console.log(`源 ONLYOFFICE-JSON: ${files.length} 篇\n`);

let identical = 0;
let sumJ2O = 0, sumO2J = 0, sumBytes = 0;
const diffAgg = new Map(); // diffKey → 次数
const perDoc = [];

for (const f of files) {
  const orig = JSON.parse(fs.readFileSync(f, "utf8"));
  const t0 = performance.now();
  const ooxml = jsonToOoxml(orig);
  const t1 = performance.now();
  const { doc: back } = ooxmlToJson(ooxml, new Gaps());
  const t2 = performance.now();
  sumJ2O += t1 - t0; sumO2J += t2 - t1; sumBytes += JSON.stringify(orig).length;

  const ds = diff(norm(orig.content), norm(back.content), "content");
  for (const d of ds) diffAgg.set(diffKey(d), (diffAgg.get(diffKey(d)) || 0) + 1);
  if (ds.length === 0) identical++;
  perDoc.push([path.basename(f), orig.content.length, ds.length]);
}

const n = files.length;
console.log("════════ 性能 ════════");
console.log(`工具① JSON→OOXML: 均 ${(sumJ2O / n).toFixed(2)}ms/篇`);
console.log(`工具② OOXML→JSON: 均 ${(sumO2J / n).toFixed(2)}ms/篇`);

console.log("\n════════ 一致性（content 深度对比）════════");
console.log(`完全一致: ${identical}/${n}`);
perDoc.slice(0, 12).forEach((d) => console.log(`  ${d[2] === 0 ? "✓" : "✗ " + d[2] + "处差异"}  ${d[0]} (${d[1]}段)`));

if (diffAgg.size) {
  console.log("\n════════ 差异种类（按频次）— 自修复待补 ════════");
  [...diffAgg.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40).forEach(([k, c]) => console.log(`  ${String(c).padStart(5)}  ${k}`));
}
