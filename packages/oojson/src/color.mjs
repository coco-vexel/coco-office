// 颜色：ONLYOFFICE-JSON {auto,r,g,b} ↔ OOXML 值 "RRGGBB" | "auto"。两工具共享。
const hx = (n) => (n & 0xff).toString(16).padStart(2, "0").toUpperCase();
export const colorToOoxml = (c) => (!c || c.auto ? "auto" : hx(c.r) + hx(c.g) + hx(c.b));
export const colorToJson = (s) =>
  !s || s === "auto"
    ? { auto: true, r: 0, g: 0, b: 0 }
    : { auto: false, r: parseInt(s.slice(0, 2), 16), g: parseInt(s.slice(2, 4), 16), b: parseInt(s.slice(4, 6), 16) };

// auto 色在 OOXML 里只存 "auto"（不携带 rgb），但 ONLYOFFICE-JSON 的 auto 色带 r/g/b 渲染缓存值。
// 为字面无损：auto 且 rgb 非 0 时，把 rgb 旁挂进私有属性；读回时合并。OOXML 仍合法（w:val=auto）。
export const autoRgbAttr = (name, c) => {
  const s = c && c.auto && (c.r || c.g || c.b) ? `${c.r},${c.g},${c.b}` : null;
  return s ? ` ${name}="${s}"` : "";
};
export const mergeAutoRgb = (jc, rgbStr) => {
  if (rgbStr && jc && jc.auto) { const p = String(rgbStr).split(",").map(Number); jc.r = p[0]; jc.g = p[1]; jc.b = p[2]; }
  return jc;
};

// highlight：ONLYOFFICE-JSON 用 RGB，OOXML 用 16 个命名色。RGB↔named（精确命中，否则最近邻）。
const HL_NAMED = { "000000": "black", "0000FF": "blue", "00FFFF": "cyan", "000080": "darkBlue", "008080": "darkCyan", "808080": "darkGray", "008000": "darkGreen", "800080": "darkMagenta", "800000": "darkRed", "808000": "darkYellow", "00FF00": "green", C0C0C0: "lightGray", FF00FF: "magenta", FF0000: "red", FFFFFF: "white", FFFF00: "yellow" };
const HL_HEX = Object.fromEntries(Object.entries(HL_NAMED).map(([h, n]) => [n, h]));
export function highlightToOoxml(c) {
  if (!c || c.auto) return null;
  const hex = colorToOoxml(c);
  if (HL_NAMED[hex]) return HL_NAMED[hex];
  let best = "none", bd = Infinity;
  for (const [h, n] of Object.entries(HL_NAMED)) {
    const dr = parseInt(h.slice(0, 2), 16) - c.r, dg = parseInt(h.slice(2, 4), 16) - c.g, db = parseInt(h.slice(4, 6), 16) - c.b;
    const d = dr * dr + dg * dg + db * db;
    if (d < bd) { bd = d; best = n; }
  }
  return best;
}
export function highlightToJson(named) {
  if (!named || named === "none") return undefined;
  return HL_HEX[named] ? { auto: false, ...colorToJson(HL_HEX[named]) } : undefined;
}
