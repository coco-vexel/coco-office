// 无损兜底层：把"未被忠实建模的剩余键"原样塞进私有属性 oo:x（JSON 文本，XML 转义），
// 读回时与模型解析结果合并。模型解析出的键覆盖兜底——反映 agent 对真实 OOXML 的编辑；
// 兜底补齐一切未建模/古怪默认键（drawing、inline、false toggle、空 tcBorders{} 等）。
// 这样"json→ooxml→json 不丢数据"是构造性保证，不依赖枚举了多少元素。
import { esc } from "./xml.mjs";

// 这些是读侧无条件重建的结构/常量键，永不进兜底。
const DROP = new Set(["type", "bFromDocument", "bPresentation"]);

/** 计算 obj 去掉 structural(递归解析负责) + 常量后的剩余键，拼成 ` name="..."`；空则 ""。
 *  name 默认 oo:x；同一元素需挂多份兜底时(如段落自身 + 其 endRun)用不同 name 区分。 */
export function extraAttr(obj, structural, name = "oo:x") {
  if (!obj || typeof obj !== "object") return "";
  const skip = structural ? new Set([...structural, ...DROP]) : DROP;
  const left = {};
  let has = false;
  for (const k of Object.keys(obj)) {
    if (skip.has(k)) continue;
    left[k] = obj[k];
    has = true;
  }
  return has ? ` ${name}="${esc(JSON.stringify(left))}"` : "";
}

/** 从节点的兜底属性(默认 oo:x)还原对象（解析器已反转义）。 */
export function readExtra(node, name = "oo:x") {
  const s = node && node.attrs ? node.attrs[name] : null;
  if (!s) return {};
  try { return JSON.parse(s); } catch { return {}; }
}

/** obj 是否只含 allowed 内的键（用于判断 run 内容项能否安全建模成真元素）。 */
export function onlyKeys(obj, allowed) {
  const set = new Set(allowed);
  for (const k of Object.keys(obj)) if (!set.has(k)) return false;
  return true;
}
