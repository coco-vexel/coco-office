// 诊断：对语料跑回环，按 diffKey 聚合并打印每类一个真实样本的 a/b 值。
import fs from "node:fs";
import path from "node:path";
import { jsonToOoxml } from "../src/json-to-ooxml.mjs";
import { ooxmlToJson, Gaps } from "../src/ooxml-to-json.mjs";

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
const diffKey = (d) => d.kind + " " + d.path.replace(/\[\d+\]/g, "[]");
function norm(o) {
  if (Array.isArray(o)) return o.map(norm);
  if (o && typeof o === "object") {
    if (o.auto === true && ("r" in o || "g" in o || "b" in o)) { const { r, g, b, ...rest } = o; return norm(rest); }
    const out = {};
    for (const k of Object.keys(o)) { if (k === "id" && o.type === "tblCell") continue; out[k] = norm(o[k]); }
    return out;
  }
  return o;
}

const target = process.argv[2] || "test/captured";
const want = process.argv[3]; // 只看包含此子串的 diffKey
const files = fs.statSync(target).isDirectory()
  ? fs.readdirSync(target).filter((e) => /\.oojson\.json$/i.test(e)).map((e) => path.join(target, e))
  : [target];

const sample = new Map(); // diffKey -> {count, a, b, file}
for (const f of files) {
  const orig = JSON.parse(fs.readFileSync(f, "utf8"));
  const ooxml = jsonToOoxml(orig);
  const { doc: back } = ooxmlToJson(ooxml, new Gaps());
  for (const d of diff(norm(orig.content), norm(back.content), "content")) {
    const k = diffKey(d);
    if (!sample.has(k)) sample.set(k, { count: 0, a: d.a, b: d.b, file: path.basename(f) });
    sample.get(k).count++;
  }
}
[...sample.entries()].sort((x, y) => y[1].count - x[1].count).forEach(([k, v]) => {
  if (want && !k.includes(want)) return;
  console.log(`\n■ ${v.count}×  ${k}   (${v.file})`);
  if (v.a !== undefined) console.log("  a =", JSON.stringify(v.a).slice(0, 300));
  if (v.b !== undefined) console.log("  b =", JSON.stringify(v.b).slice(0, 300));
});
