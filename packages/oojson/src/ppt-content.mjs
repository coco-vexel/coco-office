function clone(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function parseJson(value, label = "ONLYOFFICE PPT content JSON") {
  if (typeof value === "string") {
    try {
      return JSON.parse(value);
    } catch (error) {
      throw new Error(`${label}: invalid JSON string: ${error?.message || error}`);
    }
  }
  if (!value || typeof value !== "object") {
    throw new Error(`${label}: expected object or JSON string`);
  }
  return clone(value);
}

function normalizeHexColor(value) {
  if (typeof value !== "string") return undefined;
  const raw = value.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) return undefined;
  return `#${raw.toUpperCase()}`;
}

function ooColorToHex(value) {
  if (!value || value.auto) return undefined;
  const rgba = value.fill?.color?.color?.rgba || value.color?.rgba || value.rgba || value.fill?.color?.rgba;
  const r = Number(value.r ?? rgba?.red);
  const g = Number(value.g ?? rgba?.green);
  const b = Number(value.b ?? rgba?.blue);
  if (![r, g, b].every((n) => Number.isFinite(n))) return undefined;
  const hx = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0").toUpperCase();
  return `#${hx(r)}${hx(g)}${hx(b)}`;
}

function hexToOoColor(value) {
  const hex = normalizeHexColor(value);
  if (!hex) return undefined;
  return {
    auto: false,
    r: parseInt(hex.slice(1, 3), 16),
    g: parseInt(hex.slice(3, 5), 16),
    b: parseInt(hex.slice(5, 7), 16),
  };
}

function hexToOnlyOfficePptUniFill(value, base) {
  const color = hexToOoColor(value);
  if (!color) return undefined;
  const rgba = { red: color.r, green: color.g, blue: color.b, alpha: 255 };
  return {
    fill: {
      color: {
        rgba,
        color: {
          rgba,
          type: "srgb",
        },
        type: "uniColor",
      },
      type: "solid",
    },
    transparent: base && Object.prototype.hasOwnProperty.call(base, "transparent") ? base.transparent : null,
    type: base?.type || "fill",
  };
}

function textPrToPptStyle(rPr = {}) {
  const style = {};
  const color = ooColorToHex(rPr.uniFill || rPr.color || rPr.fill?.color);
  if (color) style.color = color;
  if (typeof rPr.b === "boolean") style.bold = rPr.b;
  if (typeof rPr.i === "boolean") style.italic = rPr.i;
  if (rPr.u != null) style.underline = rPr.u !== false && rPr.u !== "none";
  if (typeof rPr.spacing === "number") style.characterSpacing = rPr.spacing;
  if (typeof rPr.sz === "number") style.fontSize = rPr.sz / 2;
  const fonts = rPr.rFonts;
  const family = fonts?.ascii || fonts?.hAnsi || fonts?.eastAsia || fonts?.cs;
  if (family) style.fontFamily = family;
  return style;
}

function mergeStyleIntoTextPr(base = {}, style = {}) {
  const next = { ...base, type: base.type || "textPr" };
  const color = hexToOoColor(style.color);
  if (color) {
    next.color = color;
    next.uniFill = hexToOnlyOfficePptUniFill(style.color, next.uniFill);
  }
  if (typeof style.bold === "boolean") {
    next.b = style.bold;
    next.bCs = style.bold;
  }
  if (typeof style.italic === "boolean") {
    next.i = style.italic;
    next.iCs = style.italic;
  }
  if (typeof style.underline === "boolean") next.u = style.underline ? "single" : false;
  if (typeof style.characterSpacing === "number") next.spacing = style.characterSpacing;
  if (typeof style.fontSize === "number") {
    next.sz = style.fontSize * 2;
    next.szCs = style.fontSize * 2;
  }
  if (style.fontFamily) {
    next.rFonts = {
      ...(next.rFonts || {}),
      ascii: style.fontFamily,
      hAnsi: style.fontFamily,
      eastAsia: style.fontFamily,
      cs: style.fontFamily,
    };
  }
  return next;
}

function isParagraph(item) {
  return item && item.type === "paragraph" && Array.isArray(item.content);
}

function isRun(item) {
  return item && item.type === "run" && Array.isArray(item.content);
}

function isTextRunLike(item) {
  return isRun(item) || (item && item.type === "hyperlink" && Array.isArray(item.content));
}

function flattenTextContent(content) {
  if (!Array.isArray(content)) return "";
  return content
    .map((item) => {
      if (typeof item === "string") return item;
      if (item?.type === "tab") return "\t";
      if (item?.type === "break") return "\n";
      return "";
    })
    .join("");
}

function visitRunsInParagraph(paragraph, visitor) {
  for (const item of paragraph.content || []) {
    if (isRun(item)) visitor(item, paragraph);
    if (item?.type === "hyperlink" && Array.isArray(item.content)) {
      for (const child of item.content) {
        if (isRun(child)) visitor(child, paragraph);
      }
    }
  }
}

function textRunsOfContent(docContent) {
  const runs = [];
  for (const block of docContent.content || []) {
    if (!isParagraph(block)) continue;
    visitRunsInParagraph(block, (run, paragraph) => {
      const text = flattenTextContent(run.content);
      if (text) runs.push({ run, paragraph, text });
    });
  }
  return runs;
}

function makeRun(text, style = {}, template) {
  return {
    bFromDocument: true,
    footnotes: [],
    endnotes: [],
    reviewType: "common",
    ...(template ? clone(template) : {}),
    rPr: mergeStyleIntoTextPr(template?.rPr || { bFromDocument: true, type: "textPr" }, style),
    content: [text || ""],
    type: "run",
  };
}

function makeEndRun(template, style = {}) {
  return {
    bFromDocument: true,
    footnotes: [],
    endnotes: [],
    reviewType: "common",
    ...(template ? clone(template) : {}),
    rPr: mergeStyleIntoTextPr(template?.rPr || { bFromDocument: true, type: "textPr" }, style),
    content: [],
    type: "endRun",
  };
}

function makeParagraph(write = {}, template) {
  const source = template && isParagraph(template) ? clone(template) : undefined;
  const templateRun = source?.content?.find(isRun);
  const templateEndRun = source?.content?.find((item) => item?.type === "endRun");
  const runs = Array.isArray(write.runs) && write.runs.length
    ? write.runs.map((run) => makeRun(run.text || "", run.style || write.style || {}, templateRun))
    : [makeRun(write.text || "", write.style || {}, templateRun)];
  return {
    bFromDocument: true,
    ...(source || {}),
    pPr: {
      bFromDocument: true,
      type: "paraPr",
      ...(source?.pPr || {}),
    },
    rPr: mergeStyleIntoTextPr(source?.rPr || { bFromDocument: true, type: "textPr" }, write.style || {}),
    content: [...runs, makeEndRun(templateEndRun, write.style || {})],
    changes: source?.changes || [],
    type: "paragraph",
  };
}

function ensureDocContent(contentJson) {
  const docContent = parseJson(contentJson);
  if (docContent.type !== "docContent" && docContent.type !== "documentContent") {
    throw new Error(`ONLYOFFICE PPT content JSON: expected docContent/documentContent, got ${docContent.type || "unknown"}`);
  }
  if (!Array.isArray(docContent.content)) docContent.content = [];
  return docContent;
}

export function onlyOfficePptContentToTextBody(contentJson) {
  const docContent = ensureDocContent(contentJson);
  const paragraphs = [];
  const allRuns = [];
  for (const block of docContent.content || []) {
    if (!isParagraph(block)) continue;
    const runs = [];
    visitRunsInParagraph(block, (run) => {
      const text = flattenTextContent(run.content);
      if (!text) return;
      const textRun = { text, style: textPrToPptStyle(run.rPr || {}) };
      runs.push(textRun);
      allRuns.push(textRun);
    });
    paragraphs.push({
      text: runs.map((run) => run.text).join(""),
      style: block.pPr?.jc ? { alignment: block.pPr.jc } : undefined,
      runs,
    });
  }
  return {
    text: paragraphs.map((paragraph) => paragraph.text).join("\n"),
    style: allRuns[0]?.style,
    paragraphs,
  };
}

export function setOnlyOfficePptContentText(contentJson, text, options = {}) {
  const docContent = ensureDocContent(contentJson);
  const firstParagraph = docContent.content.find(isParagraph);
  docContent.content = [makeParagraph({ text: String(text ?? ""), style: options.style || {} }, firstParagraph)];
  return docContent;
}

export function setOnlyOfficePptContentParagraphs(contentJson, paragraphs) {
  const docContent = ensureDocContent(contentJson);
  const templates = docContent.content.filter(isParagraph);
  const writes = Array.isArray(paragraphs) && paragraphs.length ? paragraphs : [{ text: "" }];
  docContent.content = writes.map((paragraph, index) => makeParagraph(paragraph || {}, templates[index] || templates[0]));
  return docContent;
}

export function setOnlyOfficePptContentStyle(contentJson, style) {
  const docContent = ensureDocContent(contentJson);
  for (const block of docContent.content || []) {
    if (!isParagraph(block)) continue;
    block.rPr = mergeStyleIntoTextPr(block.rPr || { bFromDocument: true, type: "textPr" }, style || {});
    for (const item of block.content || []) {
      if (item?.type !== "endRun") continue;
      item.rPr = mergeStyleIntoTextPr(item.rPr || { bFromDocument: true, type: "textPr" }, style || {});
    }
  }
  for (const { run, paragraph } of textRunsOfContent(docContent)) {
    paragraph.rPr = mergeStyleIntoTextPr(
      paragraph.rPr || { bFromDocument: true, type: "textPr" },
      style || {},
    );
    run.rPr = mergeStyleIntoTextPr(run.rPr || { bFromDocument: true, type: "textPr" }, style || {});
  }
  return docContent;
}

export function setOnlyOfficePptContentRunStyle(contentJson, runIndex, style) {
  const docContent = ensureDocContent(contentJson);
  const runs = textRunsOfContent(docContent);
  if (!Number.isInteger(runIndex) || runIndex < 0 || runIndex >= runs.length) {
    throw new Error(`ONLYOFFICE PPT content JSON: runIndex out of range: ${runIndex}`);
  }
  const target = runs[runIndex].run;
  const paragraph = runs[runIndex].paragraph;
  paragraph.rPr = mergeStyleIntoTextPr(
    paragraph.rPr || { bFromDocument: true, type: "textPr" },
    style || {},
  );
  target.rPr = mergeStyleIntoTextPr(target.rPr || { bFromDocument: true, type: "textPr" }, style || {});
  return docContent;
}

export function replaceOnlyOfficePptContentText(contentJson, search, replacement, count = Number.POSITIVE_INFINITY) {
  const docContent = ensureDocContent(contentJson);
  if (search == null || search === "") return { content: docContent, count: 0 };
  let changed = 0;
  for (const { run } of textRunsOfContent(docContent)) {
    run.content = (run.content || []).map((item) => {
      if (typeof item !== "string" || changed >= count) return item;
      let next = item;
      while (next.includes(search) && changed < count) {
        next = next.replace(search, replacement ?? "");
        changed += 1;
      }
      return next;
    });
  }
  return { content: docContent, count: changed };
}

export function applyPptTextMutationToOnlyOfficeContent(contentJson, mutation) {
  if (!mutation || typeof mutation !== "object") {
    throw new Error("PPT mutation is required");
  }
  switch (mutation.type) {
    case "shape.text.set":
      return { content: setOnlyOfficePptContentText(contentJson, mutation.text), changedCount: 1 };
    case "shape.text.paragraphs.set":
      return {
        content: setOnlyOfficePptContentParagraphs(contentJson, mutation.paragraphs),
        changedCount: Array.isArray(mutation.paragraphs) ? mutation.paragraphs.length : 0,
      };
    case "shape.text.style.set":
      return {
        content: setOnlyOfficePptContentStyle(contentJson, mutation.style),
        changedCount: textRunsOfContent(ensureDocContent(contentJson)).length,
      };
    case "shape.text.run.style.set":
      return {
        content: setOnlyOfficePptContentRunStyle(contentJson, mutation.runIndex, mutation.style),
        changedCount: 1,
      };
    case "shape.text.replace": {
      const result = replaceOnlyOfficePptContentText(contentJson, mutation.search, mutation.replacement, mutation.count);
      return { content: result.content, changedCount: result.count };
    }
    default:
      throw new Error(`Unsupported PPT text mutation for ONLYOFFICE content bridge: ${mutation.type}`);
  }
}
