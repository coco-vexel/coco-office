// ONLYOFFICE JSON uses percentage points. OOXML numeric pct widths use 1/50 of a percent.
export function tableWidthToOoxml(type, width) {
  return type === "pct" ? Math.round(width * 50) : width;
}

export function tableWidthFromOoxml(type, width) {
  if (type !== "pct") return Number(width) || 0;
  if (width?.endsWith("%")) return Number(width.slice(0, -1)) || 0;
  return (Number(width) || 0) / 50;
}
