import { blocksToXml } from "./json-to-ooxml.mjs";

const CT = {
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  commentsExtended: "application/vnd.ms-word.commentsExtended+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
};

const REL = {
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  commentsExtended: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
  endnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
};

function cloneJson(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function contentBlocks(ref) {
  const content = ref?.content;
  if (Array.isArray(content?.content)) return content.content;
  if (Array.isArray(content)) return content;
  return null;
}

function wordPartXml(root, blocks) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:oo="urn:coco:oojson">` +
    blocksToXml(blocks) +
    `</w:${root}>`
  );
}

function walkBlocks(blocks, visit) {
  for (const block of blocks || []) {
    if (!block || typeof block !== "object") continue;
    if (Array.isArray(block.content)) {
      for (const item of block.content) visit(item);
    }
    const nested = block.content?.content;
    if (Array.isArray(nested)) walkBlocks(nested, visit);
  }
}

function commentMetadata(doc) {
  return Array.isArray(doc?.commentMetadata) ? doc.commentMetadata : [];
}

function commentDate(value) {
  if (value == null || value === "") return "";
  const date = typeof value === "number" ? new Date(value) : new Date(String(value));
  return Number.isNaN(date.getTime()) ? "" : date.toISOString();
}

function findCommentMetadata(doc, comment) {
  const metadata = commentMetadata(doc);
  // 先按 id 匹配（commentMetadata 的 id 字段，来自 GetAllComments().GetId()）
  if (comment.id != null && comment.id !== "") {
    const byId = metadata.find((item) => String(item?.id ?? "") === String(comment.id));
    if (byId) return byId;
  }
  // 再按 (author, text) 匹配（退化兼容）
  return metadata.find((item) =>
    String(item?.author ?? item?.authorName ?? "") === String(comment.author ?? "") &&
    String(item?.text ?? "") === String(comment.text ?? "")
  ) || null;
}

/**
 * 收集批注定义列表。
 *
 * 两种来源：
 *   ① 从 body 的 commentRangeStart/End 标记收集（SDK 产出或上一轮重建的 oojson）
 *   ② 从 doc.commentMetadata 直接收集（OnlyOffice 原生 ToJSON 仅输出元数据，无 body 标记）
 *
 * 情况②时，还会尝试把标记插入 body（searchQuoteInBody），使后续序列化生成
 * commentRangeStart/End/Reference 元素。
 */
function collectComments(doc) {
  const comments = new Map();
  walkBlocks(doc?.content, (item) => {
    if (!item || typeof item !== "object") return;
    if ((item.type !== "commentRangeStart" && item.type !== "commentRangeEnd" && item.type !== "commentReference") || item.id == null) return;
    const id = String(item.id);
    const prev = comments.get(id) || { id, author: "", text: "", date: "", solved: undefined };
    const next = {
      id,
      author: item.author ?? prev.author ?? "",
      text: item.text ?? prev.text ?? "",
      date: item.date ?? prev.date ?? "",
      solved: prev.solved,
    };
    const meta = findCommentMetadata(doc, next);
    if (meta?.time != null || meta?.date != null) next.date = commentDate(meta.time ?? meta.date);
    if (typeof meta?.solved === "boolean") next.solved = meta.solved;
    comments.set(id, next);
  });

  // 如果 body 中没有 commentRangeStart 标记，但有 commentMetadata：
  // 直接从元数据收集批注，并尝试在 body 中插入锚点标记
  if (comments.size === 0) {
    const metaList = commentMetadata(doc);
    for (const meta of metaList) {
      const id = String(meta.id ?? "");
      if (!id) continue;
      comments.set(id, {
        id,
        author: meta.author ?? meta.authorName ?? "",
        text: meta.text ?? "",
        date: commentDate(meta.time ?? meta.date),
        solved: typeof meta.solved === "boolean" ? meta.solved : undefined,
      });
    }
    // 尝试在 body 中插入锚点标记（基于 quote/anchorText 搜索）
    if (metaList.length > 0 && doc?.content) {
      insertCommentMarkersFromMetadata(doc, metaList);
    }
  }

  return [...comments.values()].sort((a, b) => Number(a.id) - Number(b.id));
}

/**
 * 从 commentMetadata 的 quote 字段，在 body 的 run 文本中搜索锚文本，
 * 插入 commentRangeStart/commentRangeEnd/commentReference 到内容数组。
 *
 * 只处理首段首次匹配，不跨段。匹配失败不报错（仅标记丢失，但批注定数和
 * 文本仍会出现在 comments.xml 中）。
 */
function insertCommentMarkersFromMetadata(doc, metaList) {
  const blocks = doc.content;
  if (!Array.isArray(blocks)) return;
  for (const meta of metaList) {
    const quote = meta.quote ?? meta.anchorText ?? "";
    if (!quote || String(quote).trim().length === 0) continue;
    const id = String(meta.id ?? "");
    if (!id) continue;

    // 在每个 paragraph 的 run 内容中搜索
    for (let bi = 0; bi < blocks.length; bi++) {
      const block = blocks[bi];
      if (block?.type !== "paragraph" || !Array.isArray(block.content)) continue;
      const content = block.content;
      // 构建连续文本
      let fullText = "";
      const runIndexes = [];
      for (let ci = 0; ci < content.length; ci++) {
        const item = content[ci];
        if (item?.type === "run" && Array.isArray(item.content)) {
          for (const t of item.content) {
            if (typeof t === "string") {
              fullText += t;
              runIndexes.push(ci);
            }
          }
        }
      }
      // 找首段首次匹配
      const idx = fullText.indexOf(quote);
      if (idx < 0) {
        // 尝试取 quote 前 40 字（部分 quote 可能过长）
        const short = quote.slice(0, 40);
        const idx2 = short.length >= 5 ? fullText.indexOf(short) : -1;
        if (idx2 < 0) continue;
      }
      // 简化策略：在段落首插入 commentRangeStart，末插入 commentRangeEnd + commentReference
      // 精确位置匹配需跨 run 拆分，留后续完善
      // 在第一个 run 前插入 start
      let firstRunIdx = -1;
      let lastRunIdx = -1;
      for (let ci = 0; ci < content.length; ci++) {
        if (content[ci]?.type === "run") {
          if (firstRunIdx < 0) firstRunIdx = ci;
          lastRunIdx = ci;
        }
      }
      if (firstRunIdx < 0) continue;
      // 插入 commentRangeStart 在首 run 前；commentRangeEnd 在末 run 后。
      // 注意：commentRangeEnd 在 json-to-ooxml.mjs 的 paraItemXml 中会自动输出
      // commentReference，不需要额外插入。
      content.splice(firstRunIdx, 0, { type: "commentRangeStart", id });
      // 插入 start 后 runs 偏移 +1
      content.splice(lastRunIdx + 2, 0, { type: "commentRangeEnd", id });
      break; // 只处理首段匹配
    }
  }
}

function commentXmlAttrs(comment) {
  const date = comment.date ? ` w:date="${escapeXml(comment.date)}"` : "";
  return `w:id="${escapeXml(comment.id)}" w:author="${escapeXml(comment.author || "")}"${date}`;
}

function commentParaId(comment) {
  const numeric = Number(comment.id);
  if (Number.isFinite(numeric)) return (numeric >>> 0).toString(16).toUpperCase().padStart(8, "0").slice(-8);
  let hash = 0;
  for (const ch of String(comment.id)) hash = ((hash * 31) + ch.charCodeAt(0)) >>> 0;
  return hash.toString(16).toUpperCase().padStart(8, "0");
}

function commentsXml(comments) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
    comments.map((comment) =>
      `<w:comment ${commentXmlAttrs(comment)}>` +
      `<w:p w14:paraId="${commentParaId(comment)}"><w:r><w:t xml:space="preserve">${escapeXml(comment.text || "")}</w:t></w:r></w:p>` +
      `</w:comment>`
    ).join("") +
    `</w:comments>`
  );
}

function commentsExtendedXml(comments) {
  const solved = comments.filter((comment) => comment.solved === true);
  if (!solved.length) return "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
    solved.map((comment) => `<w15:commentEx w15:paraId="${commentParaId(comment)}" w15:done="1"/>`).join("") +
    `</w15:commentsEx>`
  );
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function appendCommentParts(doc, extraParts, summary) {
  const comments = collectComments(doc);
  if (!comments.length) return;
  summary.comments = comments.length;
  extraParts.push({
    name: "/word/comments.xml",
    contentType: CT.comments,
    xml: commentsXml(comments),
    rel: {
      id: "rIdCocoComments",
      type: REL.comments,
      target: "comments.xml",
    },
  });
  const extendedXml = commentsExtendedXml(comments);
  if (extendedXml) {
    extraParts.push({
      name: "/word/commentsExtended.xml",
      contentType: CT.commentsExtended,
      xml: extendedXml,
      rel: {
        id: "rIdCocoCommentsExtended",
        type: REL.commentsExtended,
        target: "commentsExtended.xml",
      },
    });
  }
}
function parseDataUri(uri) {
  const m = /^data:([^;,]+);base64,([A-Za-z0-9+/=\s]+)$/.exec(String(uri || ""));
  if (!m) return null;
  const contentType = m[1];
  const ext = contentType === "image/jpeg" ? "jpg" : contentType === "image/png" ? "png" : contentType.split("/")[1] || "bin";
  return { contentType, ext, data: m[2].replace(/\s+/g, "") };
}

function collectNotes(doc, key) {
  const notes = new Map();
  walkBlocks(doc?.content, (item) => {
    if (!item || item.type !== "run" || !Array.isArray(item[key])) return;
    for (const note of item[key]) {
      if (!note || note.id == null || !Array.isArray(note.content)) continue;
      notes.set(String(note.id), { id: String(note.id), type: note.type, content: note.content });
    }
  });
  return [...notes.values()].sort((a, b) => Number(a.id) - Number(b.id));
}

function noteXmlAttrs(note) {
  const type = note.type ? ` w:type="${escapeXml(note.type)}"` : "";
  return `w:id="${escapeXml(note.id)}"${type}`;
}

function isNoteKind(note, type, id) {
  return note?.type === type || String(note?.id) === String(id);
}

function defaultNoteXml(child, type, id) {
  const marker = type === "separator" ? "separator" : "continuationSeparator";
  return `<w:${child} w:type="${type}" w:id="${id}"><w:p><w:r><w:${marker}/></w:r></w:p></w:${child}>`;
}

function noteContentCount(notes) {
  return notes.filter((note) => !isNoteKind(note, "separator", -1) && !isNoteKind(note, "continuationSeparator", 0)).length;
}

function noteEntryXml(child, note) {
  return `<w:${child} ${noteXmlAttrs(note)}>${blocksToXml(note.content)}</w:${child}>`;
}

function notesXml(root, child, notes) {
  const separator = notes.find((note) => isNoteKind(note, "separator", -1));
  const continuation = notes.find((note) => isNoteKind(note, "continuationSeparator", 0));
  const bodyNotes = notes.filter((note) => !isNoteKind(note, "separator", -1) && !isNoteKind(note, "continuationSeparator", 0));
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:oo="urn:coco:oojson">` +
    (separator ? noteEntryXml(child, separator) : defaultNoteXml(child, "separator", -1)) +
    (continuation ? noteEntryXml(child, continuation) : defaultNoteXml(child, "continuationSeparator", 0)) +
    bodyNotes.map((note) => noteEntryXml(child, note)).join("") +
    `</w:${root}>`
  );
}

function appendNoteParts(doc, extraParts, summary) {
  const footnotes = collectNotes(doc, "footnotes");
  if (footnotes.length) {
    summary.footnotes = noteContentCount(footnotes);
    extraParts.push({
      name: "/word/footnotes.xml",
      contentType: CT.footnotes,
      xml: notesXml("footnotes", "footnote", footnotes),
      rel: { id: "rIdCocoFootnotes", type: REL.footnotes, target: "footnotes.xml" },
    });
  }
  const endnotes = collectNotes(doc, "endnotes");
  if (endnotes.length) {
    summary.endnotes = noteContentCount(endnotes);
    extraParts.push({
      name: "/word/endnotes.xml",
      contentType: CT.endnotes,
      xml: notesXml("endnotes", "endnote", endnotes),
      rel: { id: "rIdCocoEndnotes", type: REL.endnotes, target: "endnotes.xml" },
    });
  }
}
function appendMediaPartsFromBlocks(blocks, extraParts, summary, sourcePart = "/word/document.xml") {
  walkBlocks(blocks, (item) => {
    if (!item || item.type !== "run" || !Array.isArray(item.content)) return;
    for (const child of item.content) {
      if (!child || child.type !== "paraDrawing") continue;
      const parsed = parseDataUri(child.graphic?.blipFill?.rasterImageId);
      if (!parsed) continue;
      const index = summary.media + 1;
      const relId = `rIdCocoImage${index}`;
      child.graphic.blipFill.rasterImageId = relId;
      summary.media = index;
      const rel = {
        id: relId,
        type: REL.image,
        target: `media/image${index}.${parsed.ext}`,
      };
      if (sourcePart !== "/word/document.xml") rel.source = sourcePart;
      extraParts.push({
        name: `/word/media/image${index}.${parsed.ext}`,
        contentType: parsed.contentType,
        binary: parsed.data,
        rel,
      });
    }
  });
}

function appendMediaParts(doc, extraParts, summary) {
  appendMediaPartsFromBlocks(doc?.content, extraParts, summary);
}
function appendHeaderFooterParts(doc, extraParts, summary) {
  const sectPr = doc?.sectPr || doc?.sectionPr;
  if (!sectPr || typeof sectPr !== "object") return;
  const groups = [
    { key: "headerReference", part: "header", root: "hdr", contentType: CT.header, relType: REL.header },
    { key: "footerReference", part: "footer", root: "ftr", contentType: CT.footer, relType: REL.footer },
  ];
  for (const group of groups) {
    const refs = sectPr[group.key];
    if (!refs || typeof refs !== "object") continue;
    for (const type of ["first", "default", "even"]) {
      const ref = refs[type];
      const blocks = contentBlocks(ref);
      if (!blocks) continue;
      const index = summary[group.part] + 1;
      summary[group.part] = index;
      const relId = `rIdCoco${group.part[0].toUpperCase()}${group.part.slice(1)}${index}`;
      const name = `/word/${group.part}${index}.xml`;
      appendMediaPartsFromBlocks(blocks, extraParts, summary, name);
      extraParts.push({
        name,
        contentType: group.contentType,
        xml: wordPartXml(group.root, blocks),
        rel: {
          id: relId,
          type: group.relType,
          target: `${group.part}${index}.xml`,
        },
      });
      refs[type] = { id: relId };
    }
  }
}

export function withPackageLongTailParts(ooJsonDoc) {
  const doc = cloneJson(ooJsonDoc);
  const extraParts = [];
  const summary = { header: 0, footer: 0, comments: 0, media: 0, footnotes: 0, endnotes: 0 };
  appendHeaderFooterParts(doc, extraParts, summary);
  appendCommentParts(doc, extraParts, summary);
  appendNoteParts(doc, extraParts, summary);
  appendMediaParts(doc, extraParts, summary);
  return { doc, extraParts, summary };
}