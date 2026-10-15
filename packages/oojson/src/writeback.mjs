import { fromFlatOpc } from '@coco-vexel/coco-office/pack';
import { ooxmlToJson } from './ooxml-to-json.mjs';
import { esc, kid, kidsOf, parseXml } from './xml.mjs';

const CT = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  commentsExtended: 'application/vnd.ms-word.commentsExtended+xml',
};

const REL = {
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
};

const GAP_DOMAIN_ORDER = ['comments', 'headersFooters', 'notes', 'bookmarks', 'media', 'drawings'];

function serializeXmlPart(part) {
  if (part?._originalBytes) {
    return new TextDecoder().decode(part._originalBytes);
  }
  if (!part?.xmlDom?.documentElement) {
    throw new Error('flatOpcToOnlyOfficeWriteback: missing XML DOM for word/document.xml');
  }
  if (typeof XMLSerializer !== 'function') {
    throw new Error(
      'flatOpcToOnlyOfficeWriteback: XMLSerializer is unavailable and part has no original bytes',
    );
  }
  return new XMLSerializer().serializeToString(part.xmlDom.documentElement);
}

function emptyGapDomain(domain) {
  return {
    domain,
    present: false,
    count: 0,
    parts: [],
    status: 'pending',
  };
}

function serializeOptionalXmlPart(part) {
  return part ? serializeXmlPart(part) : '';
}

function textContent(node) {
  if (!node) return '';
  if (typeof node.text === 'string') return node.text;
  return (node.kids || []).map(textContent).join('');
}

function parseCommentDefinitions(commentsXml) {
  if (!commentsXml) return new Map();
  const root = parseXml(commentsXml);
  const comments = new Map();
  for (const node of kidsOf(root, 'w:comment')) {
    const id = String(node.attrs?.['w:id'] ?? '');
    if (!id) continue;
    comments.set(id, {
      id,
      author: String(node.attrs?.['w:author'] ?? ''),
      date: String(node.attrs?.['w:date'] ?? ''),
      text: textContent(node),
    });
  }
  return comments;
}

function parseCommentAnchors(documentXml) {
  if (!documentXml) return new Map();
  const root = parseXml(documentXml);
  const body = kid(root, 'w:body');
  const anchors = new Map();
  const active = new Map();

  function visit(node, paragraphText = '') {
    if (!node || typeof node !== 'object') return;
    if (node.tag === 'w:p') {
      const currentParagraphText = textContent(node);
      for (const child of node.kids || []) visit(child, currentParagraphText);
      return;
    }
    if (node.tag === 'w:commentRangeStart') {
      const id = String(node.attrs?.['w:id'] ?? '');
      if (id) active.set(id, { anchorText: '', paragraphText });
      return;
    }
    if (node.tag === 'w:commentRangeEnd') {
      const id = String(node.attrs?.['w:id'] ?? '');
      if (id && active.has(id)) {
        anchors.set(id, active.get(id));
        active.delete(id);
      }
      return;
    }
    if (node.tag === 'w:t') {
      const text = textContent(node);
      for (const current of active.values()) current.anchorText += text;
      return;
    }
    for (const child of node.kids || []) visit(child, paragraphText);
  }

  visit(body || root);
  return anchors;
}

function buildCommentItems(pkg, documentXml) {
  const commentsPart = pkg.parts.get('word/comments.xml');
  if (!commentsPart) return [];
  const comments = parseCommentDefinitions(serializeOptionalXmlPart(commentsPart));
  const anchors = parseCommentAnchors(documentXml);
  return [...comments.values()].map((comment) => {
    const anchor = anchors.get(comment.id);
    return {
      ...comment,
      anchorText: anchor?.anchorText ?? '',
      paragraphText: anchor?.paragraphText ?? '',
    };
  });
}

function parseNoteDefinitions(pkg, partPath, notesXml, kind) {
  if (!notesXml) return new Map();
  const root = parseXml(inlineStoryImages(pkg, partPath, notesXml));
  const childTag = kind === 'endnote' ? 'w:endnote' : 'w:footnote';
  const notes = new Map();
  for (const node of kidsOf(root, childTag)) {
    const id = String(node.attrs?.['w:id'] ?? '');
    if (!id) continue;
    const type = String(node.attrs?.['w:type'] ?? '');
    if (type === 'separator' || type === 'continuationSeparator') continue;
    if (Number(id) < 1) continue;
    const bodyKids = (node?.kids || []).filter(
      (child) => child?.tag === 'w:p' || child?.tag === 'w:tbl',
    );
    const bodyXml = bodyKids.map((child) => serializeMiniXml(child)).join('');
    notes.set(id, {
      kind,
      id,
      type,
      text: textContent(node),
      bodyJson: ooxmlToJson(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}</w:body></w:document>`,
      ).doc,
    });
  }
  return notes;
}

function parseNoteAnchors(documentXml) {
  if (!documentXml) return new Map();
  const root = parseXml(documentXml);
  const body = kid(root, 'w:body');
  const anchors = new Map();

  function visitParagraph(paragraph) {
    let text = '';
    function visit(node) {
      if (!node || typeof node !== 'object') return;
      if (node.tag === 'w:t' || node.tag === 'w:delText') {
        text += textContent(node);
        return;
      }
      if (node.tag === 'w:footnoteReference' || node.tag === 'w:endnoteReference') {
        const id = String(node.attrs?.['w:id'] ?? '');
        const kind = node.tag === 'w:endnoteReference' ? 'endnote' : 'footnote';
        if (id) anchors.set(`${kind}:${id}`, text.trim());
        return;
      }
      for (const child of node.kids || []) visit(child);
    }
    visit(paragraph);
  }

  function visitBlocks(node) {
    if (!node || typeof node !== 'object') return;
    if (node.tag === 'w:p') {
      visitParagraph(node);
      return;
    }
    for (const child of node.kids || []) visitBlocks(child);
  }

  visitBlocks(body || root);
  return anchors;
}

function buildNoteItems(pkg, documentXml) {
  const anchors = parseNoteAnchors(documentXml);
  const definitions = [
    ...parseNoteDefinitions(
      pkg,
      'word/footnotes.xml',
      serializeOptionalXmlPart(pkg.parts.get('word/footnotes.xml')),
      'footnote',
    ).values(),
    ...parseNoteDefinitions(
      pkg,
      'word/endnotes.xml',
      serializeOptionalXmlPart(pkg.parts.get('word/endnotes.xml')),
      'endnote',
    ).values(),
  ];
  return definitions.map((note) => ({
    ...note,
    anchorText: anchors.get(`${note.kind}:${note.id}`) ?? '',
  }));
}

function parseBookmarkItems(documentXml) {
  if (!documentXml) return [];
  const root = parseXml(documentXml);
  const body = kid(root, 'w:body');
  const active = new Map();
  const items = [];

  function visit(node) {
    if (!node || typeof node !== 'object') return;
    if (node.tag === 'w:bookmarkStart') {
      const id = String(node.attrs?.['w:id'] ?? '');
      const name = String(node.attrs?.['w:name'] ?? '');
      if (id && name && !name.startsWith('_')) {
        active.set(id, { id, name, anchorText: '' });
      }
      return;
    }
    if (node.tag === 'w:bookmarkEnd') {
      const id = String(node.attrs?.['w:id'] ?? '');
      const bookmark = active.get(id);
      if (bookmark) {
        items.push({ ...bookmark, anchorText: bookmark.anchorText.trim() });
        active.delete(id);
      }
      return;
    }
    if (node.tag === 'w:t' || node.tag === 'w:delText') {
      const text = textContent(node);
      for (const bookmark of active.values()) bookmark.anchorText += text;
      return;
    }
    for (const child of node.kids || []) visit(child);
  }

  visit(body || root);
  return items.filter((item) => item.name && item.anchorText);
}

function normalizePartPath(path) {
  return String(path || '').replace(/^\/+/, '');
}

function resolveWordTarget(target) {
  const clean = String(target || '').replace(/^\/+/, '');
  if (!clean) return '';
  return clean.startsWith('word/') ? clean : `word/${clean}`;
}

function parseDocumentRelationships(pkg) {
  const relsPart = pkg.parts.get('word/_rels/document.xml.rels');
  if (!relsPart) return new Map();
  const root = parseXml(serializeOptionalXmlPart(relsPart));
  const rels = new Map();
  for (const rel of kidsOf(root, 'Relationship')) {
    const id = String(rel.attrs?.Id ?? '');
    if (!id) continue;
    rels.set(id, {
      id,
      type: String(rel.attrs?.Type ?? ''),
      target:
        String(rel.attrs?.TargetMode ?? '') === 'External'
          ? String(rel.attrs?.Target ?? '')
          : resolveWordTarget(rel.attrs?.Target),
      targetMode: String(rel.attrs?.TargetMode ?? ''),
    });
  }
  return rels;
}

function parseHeaderFooterReferences(documentXml, rels) {
  const root = parseXml(documentXml);
  const body = kid(root, 'w:body');
  const refs = [];

  function visit(node) {
    if (!node || typeof node !== 'object') return;
    if (node.tag === 'w:headerReference' || node.tag === 'w:footerReference') {
      const relId = String(node.attrs?.['r:id'] ?? '');
      const rel = rels.get(relId);
      if (rel) {
        refs.push({
          kind: rel.type === REL.footer ? 'footer' : 'header',
          type: String(node.attrs?.['w:type'] ?? 'default'),
          relId,
          partPath: rel.target,
        });
      }
      return;
    }
    for (const child of node.kids || []) visit(child);
  }

  visit(body || root);
  return refs;
}

function parseSectionProperties(documentXml) {
  const root = parseXml(documentXml);
  const body = kid(root, 'w:body');
  const sections = [];
  const bodyKids = body?.kids || [];
  let blockIndex = -1;

  for (const block of bodyKids) {
    if (!block || typeof block !== 'object') continue;
    if (block.tag === 'w:sectPr') continue;
    if (block.tag !== 'w:p' && block.tag !== 'w:tbl') continue;
    blockIndex++;
    const pPr = kid(block, 'w:pPr');
    const sectPr = kid(pPr, 'w:sectPr');
    if (sectPr) {
      sections.push({
        node: sectPr,
        boundary: 'paragraph',
        endBlockIndex: blockIndex,
      });
    }
  }

  const finalSectPr = kid(body, 'w:sectPr');
  if (finalSectPr) {
    sections.push({
      node: finalSectPr,
      boundary: 'body',
      endBlockIndex: Math.max(0, blockIndex),
    });
  }

  return sections;
}

function sectionReferenceMap(sectPr, rels, kind) {
  const tag = kind === 'footer' ? 'w:footerReference' : 'w:headerReference';
  const relType = kind === 'footer' ? REL.footer : REL.header;
  const refs = new Map();
  for (const node of kidsOf(sectPr, tag)) {
    const type = String(node.attrs?.['w:type'] ?? 'default');
    const relId = String(node.attrs?.['r:id'] ?? '');
    const rel = rels.get(relId);
    refs.set(type, {
      kind,
      type,
      relId,
      partPath: rel ? rel.target : '',
      relFound: !!rel,
      relTypeMatches: !!rel && rel.type === relType,
    });
  }
  return refs;
}

function sectionVariantState(pkg, refs, type, sectionIndex) {
  const ref = refs.get(type);
  if (!ref) {
    return sectionIndex === 0
      ? { state: 'empty', type }
      : { state: 'inherited', type, inheritedFrom: sectionIndex - 1 };
  }
  const partPath = normalizePartPath(ref.partPath);
  const part = partPath ? pkg.parts.get(partPath) : null;
  const xml = serializeOptionalXmlPart(part);
  const text = part ? textContent(parseXml(xml)) : '';
  return {
    state: text ? 'own' : 'empty',
    type,
    relId: ref.relId,
    partPath,
    text,
    bodyJson: part ? storyPartToBodyJson(pkg, partPath, xml) : undefined,
    unsupportedFields: part ? collectUnsupportedStoryFields(xml) : [],
    relFound: ref.relFound,
    relTypeMatches: ref.relTypeMatches,
    partFound: !!part,
  };
}

function collectUnsupportedStoryFields(xml) {
  if (!xml) return [];
  const root = parseXml(xml);
  const fields = [];
  let index = 0;

  function visit(node) {
    if (!node || typeof node !== 'object') return;
    if (node.tag === 'w:instrText') {
      const instruction = textContent(node).trim().replace(/\s+/g, ' ');
      const fieldType = instruction.split(/\s+/)[0]?.toUpperCase() ?? '';
      if (['DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE'].includes(fieldType)) {
        fields.push({
          index: index++,
          fieldType,
          instruction,
          unsupportedReason: 'dynamic-date-field',
        });
      }
    }
    for (const child of node.kids || []) visit(child);
  }

  visit(root);
  return fields;
}

function storyPartToBodyJson(pkg, partPath, xml) {
  if (!xml) return undefined;
  const root = parseXml(inlineStoryImages(pkg, partPath, xml));
  const bodyKids = (root?.kids || []).filter(
    (child) => child?.tag === 'w:p' || child?.tag === 'w:tbl',
  );
  const bodyXml = bodyKids.map((child) => serializeMiniXml(child)).join('');
  return ooxmlToJson(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}</w:body></w:document>`,
  ).doc;
}

function inlineStoryImages(pkg, partPath, xml) {
  const rels = parsePartRelationships(pkg, partPath);
  return String(xml).replace(/<w:drawing[\s\S]*?<\/w:drawing>/g, (drawingXml) => {
    const drawing = drawingJsonFromXml(pkg, rels, drawingXml);
    return drawing ? `<oo:ci oo:v="${esc(JSON.stringify(drawing))}"/>` : drawingXml;
  });
}

function parsePartRelationships(pkg, partPath) {
  const relsPath = relsPathForPart(partPath);
  const relsPart = pkg.parts.get(relsPath);
  if (!relsPart) return new Map();
  const root = parseXml(serializeOptionalXmlPart(relsPart));
  const rels = new Map();
  for (const rel of kidsOf(root, 'Relationship')) {
    const id = String(rel.attrs?.Id ?? '');
    if (!id) continue;
    const targetMode = String(rel.attrs?.TargetMode ?? '');
    rels.set(id, {
      id,
      type: String(rel.attrs?.Type ?? ''),
      target:
        targetMode === 'External'
          ? String(rel.attrs?.Target ?? '')
          : resolvePartTarget(partPath, rel.attrs?.Target),
      targetMode,
    });
  }
  return rels;
}

function relsPathForPart(partPath) {
  const clean = normalizePartPath(partPath);
  const slash = clean.lastIndexOf('/');
  return slash < 0
    ? `_rels/${clean}.rels`
    : `${clean.slice(0, slash)}/_rels/${clean.slice(slash + 1)}.rels`;
}

function resolvePartTarget(sourcePartPath, target) {
  let value = String(target || '').trim();
  if (!value) return '';
  if (value.startsWith('/')) return normalizePartPath(value);
  const base = normalizePartPath(sourcePartPath).split('/').slice(0, -1);
  for (const segment of value.split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') base.pop();
    else base.push(segment);
  }
  return base.join('/');
}

function drawingJsonFromXml(pkg, rels, drawingXml) {
  const analysis = analyzeDrawingXml(pkg, 'unknown', 0, rels, drawingXml);
  if (analysis.status !== 'inlined') return null;
  const root = parseXml(drawingXml);
  const extent = findNode(root, 'wp:extent') || findNode(root, 'a:ext');
  const docPr = findNode(root, 'wp:docPr');
  const cx = numberValue(extent?.attrs?.cx) || 1;
  const cy = numberValue(extent?.attrs?.cy) || 1;
  return {
    type: 'paraDrawing',
    docPr: {
      id: numberValue(docPr?.attrs?.id) || 1,
      name: String(docPr?.attrs?.name ?? 'Picture'),
      descr: docPr?.attrs?.descr == null ? undefined : String(docPr.attrs.descr),
      title: docPr?.attrs?.title == null ? undefined : String(docPr.attrs.title),
    },
    effectExtent: { l: 0, t: 0, r: 0, b: 0 },
    extent: { cx, cy },
    graphic: {
      type: 'image',
      id: '',
      extX: 0,
      extY: 0,
      nvPicPr: {
        cNvPr: {
          id: numberValue(docPr?.attrs?.id) || 1,
          name: String(docPr?.attrs?.name ?? 'Picture'),
          descr: docPr?.attrs?.descr == null ? null : String(docPr.attrs.descr),
          title: docPr?.attrs?.title == null ? null : String(docPr.attrs.title),
          hidden: null,
        },
        nvPr: { isPhoto: null, userDrawn: null },
        cNvPicPr: { picLocks: { noChangeAspect: true } },
      },
      blipFill: {
        type: 'blipFill',
        blip: [],
        srcRect: null,
        tile: null,
        stretch: { fillRect: null },
        rotWithShape: null,
        rasterImageId: analysis.rasterImageId,
      },
      spPr: {
        fill: null,
        ln: null,
        effectDag: null,
        effectLst: null,
        xfrm: {
          ext: { cx, cy },
          off: { x: 0, y: 0 },
          flipH: null,
          flipV: null,
          rot: null,
          chOffX: null,
          chOffY: null,
          chExtX: null,
          chExtY: null,
        },
        bwMode: 0,
      },
    },
    cNvGraphicFramePr: { graphicFrameLocks: { noChangeAspect: true } },
    positionH: { relativeFrom: 'column', posOffset: 0, percent: false },
    positionV: { relativeFrom: 'paragraph', posOffset: 0, percent: false },
    simplePos: { x: 0, y: 0, use: false },
    distB: 0,
    distL: 0,
    distR: 0,
    distT: 0,
    allowOverlap: true,
    behindDoc: false,
    hidden: null,
    layoutInCell: true,
    locked: false,
    relativeHeight: 0,
    wrapType: 'none',
    drawingType: 'inline',
  };
}

function analyzeDrawingXml(pkg, partPath, index, rels, drawingXml) {
  const root = parseXml(drawingXml);
  const inline = findNode(root, 'wp:inline');
  const anchor = findNode(root, 'wp:anchor');
  const graphicData = findNode(root, 'a:graphicData');
  const blip = findNode(root, 'a:blip');
  const relId = String(blip?.attrs?.['r:embed'] ?? '');
  const linkRelId = String(blip?.attrs?.['r:link'] ?? '');
  const effectiveRelId = relId || linkRelId;
  const rel = effectiveRelId ? rels.get(effectiveRelId) : null;
  const mediaPart = rel?.target ? pkg.parts.get(rel.target) : null;
  const base = {
    partPath: normalizePartPath(partPath),
    index,
    drawingType: anchor ? 'anchor' : inline ? 'inline' : 'unknown',
    relationshipId: effectiveRelId,
    relationshipType: rel?.type ?? '',
    target: rel?.target ?? '',
    targetMode: rel?.targetMode ?? '',
    contentType: mediaPart?.contentType ?? '',
    graphicDataUri: String(graphicData?.attrs?.uri ?? ''),
  };

  if (anchor) return { ...base, status: 'unsupported', unsupportedReason: 'floating-anchor' };
  if (!inline)
    return { ...base, status: 'unsupported', unsupportedReason: 'unknown-drawing-layout' };
  const graphicReason = unsupportedGraphicReason(base.graphicDataUri, drawingXml);
  if (graphicReason) return { ...base, status: 'unsupported', unsupportedReason: graphicReason };
  if (!blip) return { ...base, status: 'unsupported', unsupportedReason: 'missing-blip' };
  if (linkRelId || rel?.targetMode === 'External')
    return { ...base, status: 'unsupported', unsupportedReason: 'external-image' };
  if (!rel) return { ...base, status: 'unsupported', unsupportedReason: 'missing-relationship' };
  if (rel.type && rel.type !== REL.image)
    return { ...base, status: 'unsupported', unsupportedReason: 'non-image-relationship' };
  if (!mediaPart?.bytes)
    return { ...base, status: 'unsupported', unsupportedReason: 'missing-media-part' };
  if (!String(mediaPart.contentType || '').startsWith('image/'))
    return { ...base, status: 'unsupported', unsupportedReason: 'non-image-media' };
  if (base.graphicDataUri && !base.graphicDataUri.includes('/picture'))
    return { ...base, status: 'unsupported', unsupportedReason: 'non-picture-graphic' };

  return {
    ...base,
    status: 'inlined',
    rasterImageId: `data:${mediaPart.contentType};base64,${bytesToBase64(mediaPart.bytes)}`,
  };
}

function unsupportedGraphicReason(graphicDataUri, drawingXml) {
  const uri = String(graphicDataUri || '');
  const xml = String(drawingXml || '');
  if (!uri || uri.includes('/picture')) return '';
  if (uri.includes('/chart')) return 'chart-drawing';
  if (uri.includes('/diagram')) return 'diagram-drawing';
  if (xml.includes(':txbx') || xml.includes('<w:txbxContent')) return 'text-box-drawing';
  if (uri.includes('wordprocessingShape') || uri.includes('/shape')) return 'shape-drawing';
  return 'non-picture-graphic';
}

function findNode(node, tag) {
  if (!node || typeof node !== 'object') return null;
  if (node.tag === tag) return node;
  for (const child of node.kids || []) {
    const found = findNode(child, tag);
    if (found) return found;
  }
  return null;
}

function numberValue(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

function bytesToBase64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function serializeMiniXml(node) {
  if (!node || typeof node !== 'object') return '';
  if (typeof node.text === 'string')
    return node.text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attrs = Object.entries(node.attrs || {})
    .map(
      ([key, value]) => ` ${key}="${String(value).replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`,
    )
    .join('');
  const kids = (node.kids || []).map(serializeMiniXml).join('');
  return `<${node.tag}${attrs}>${kids}</${node.tag}>`;
}

function numberAttr(node, name) {
  if (!node?.attrs || node.attrs[name] == null) return undefined;
  const value = Number(node.attrs[name]);
  return Number.isFinite(value) ? value : undefined;
}

function parsePageSize(sectPr) {
  const pgSz = kid(sectPr, 'w:pgSz');
  if (!pgSz) return undefined;
  const width = numberAttr(pgSz, 'w:w');
  const height = numberAttr(pgSz, 'w:h');
  if (width == null && height == null) return undefined;
  return {
    width,
    height,
    orient: String(pgSz.attrs?.['w:orient'] ?? ''),
  };
}

function parsePageMargins(sectPr) {
  const pgMar = kid(sectPr, 'w:pgMar');
  if (!pgMar) return undefined;
  const margins = {
    top: numberAttr(pgMar, 'w:top'),
    right: numberAttr(pgMar, 'w:right'),
    bottom: numberAttr(pgMar, 'w:bottom'),
    left: numberAttr(pgMar, 'w:left'),
    header: numberAttr(pgMar, 'w:header'),
    footer: numberAttr(pgMar, 'w:footer'),
    gutter: numberAttr(pgMar, 'w:gutter'),
  };
  return Object.values(margins).some((value) => value != null) ? margins : undefined;
}

function parsePageNumbering(sectPr) {
  const pgNumType = kid(sectPr, 'w:pgNumType');
  if (!pgNumType) return undefined;
  const start = numberAttr(pgNumType, 'w:start');
  const numbering = {
    start,
    fmt: String(pgNumType.attrs?.['w:fmt'] ?? ''),
    chapStyle: String(pgNumType.attrs?.['w:chapStyle'] ?? ''),
    chapSep: String(pgNumType.attrs?.['w:chapSep'] ?? ''),
  };
  return Object.values(numbering).some((value) => value !== '' && value != null)
    ? numbering
    : undefined;
}

function parseColumns(sectPr) {
  const cols = kid(sectPr, 'w:cols');
  if (!cols) return undefined;
  const colNodes = kidsOf(cols, 'w:col');
  const col = colNodes.map((node) => ({
    width: numberAttr(node, 'w:w'),
    space: numberAttr(node, 'w:space'),
  }));
  const columns = {
    equalWidth: String(cols.attrs?.['w:equalWidth'] ?? '1') !== '0',
    count: numberAttr(cols, 'w:num'),
    space: numberAttr(cols, 'w:space'),
    sep: String(cols.attrs?.['w:sep'] ?? ''),
    col,
  };
  return columns.count != null || columns.space != null || col.length > 0 ? columns : undefined;
}

function buildHeaderFooterSections(pkg, documentXml) {
  const rels = parseDocumentRelationships(pkg);
  const sections = parseSectionProperties(documentXml);
  const types = ['default', 'first', 'even'];
  return sections.map((section, sectionIndex) => {
    const sectPr = section.node;
    const headerRefs = sectionReferenceMap(sectPr, rels, 'header');
    const footerRefs = sectionReferenceMap(sectPr, rels, 'footer');
    const previousEndBlockIndex = sectionIndex > 0 ? sections[sectionIndex - 1].endBlockIndex : -1;
    return {
      index: sectionIndex,
      startBlockIndex: Math.max(0, previousEndBlockIndex + 1),
      endBlockIndex: section.endBlockIndex,
      boundary: section.boundary,
      type: String(kid(sectPr, 'w:type')?.attrs?.['w:val'] ?? 'nextPage'),
      titlePage: !!kid(sectPr, 'w:titlePg'),
      pageSize: parsePageSize(sectPr),
      pageMargins: parsePageMargins(sectPr),
      pageNumbering: parsePageNumbering(sectPr),
      columns: parseColumns(sectPr),
      headers: Object.fromEntries(
        types.map((type) => [type, sectionVariantState(pkg, headerRefs, type, sectionIndex)]),
      ),
      footers: Object.fromEntries(
        types.map((type) => [type, sectionVariantState(pkg, footerRefs, type, sectionIndex)]),
      ),
    };
  });
}

function buildHeaderFooterItems(pkg, documentXml) {
  const rels = parseDocumentRelationships(pkg);
  const refs = parseHeaderFooterReferences(documentXml, rels);
  const seen = new Set();
  const items = [];
  for (const ref of refs) {
    const key = `${ref.kind}|${ref.type}|${ref.partPath}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const part = pkg.parts.get(normalizePartPath(ref.partPath));
    if (!part) continue;
    items.push({
      kind: ref.kind,
      type: ref.type,
      partPath: normalizePartPath(ref.partPath),
      text: textContent(parseXml(serializeOptionalXmlPart(part))),
    });
  }
  return items;
}

function stripCommentMarkers(value) {
  if (Array.isArray(value)) {
    return value
      .filter(
        (item) =>
          !['commentRangeStart', 'commentRangeEnd', 'commentReference'].includes(item?.type),
      )
      .map(stripCommentMarkers);
  }
  if (!value || typeof value !== 'object') return value;
  const out = {};
  for (const [key, child] of Object.entries(value)) {
    out[key] = stripCommentMarkers(child);
  }
  return out;
}

function stripBookmarkMarkers(value) {
  if (Array.isArray(value)) {
    return value
      .filter((item) => !['bookmarkStart', 'bookmarkEnd'].includes(item?.type))
      .map(stripBookmarkMarkers);
  }
  if (!value || typeof value !== 'object') return value;
  const out = {};
  for (const [key, child] of Object.entries(value)) {
    out[key] = stripBookmarkMarkers(child);
  }
  return out;
}

function addGapPart(gaps, domain, part) {
  const gap = gaps[domain];
  gap.present = true;
  gap.count += 1;
  gap.parts.push({
    path: part.path,
    contentType: part.contentType,
  });
}

function drawingXmlFragments(xml) {
  return [...String(xml || '').matchAll(/<w:drawing[\s\S]*?<\/w:drawing>/g)].map(
    (match) => match[0],
  );
}

function isStoryXmlPart(part) {
  return (
    part?.path === 'word/document.xml' ||
    part?.contentType === CT.header ||
    part?.contentType === CT.footer ||
    part?.contentType === CT.footnotes ||
    part?.contentType === CT.endnotes
  );
}

function collectDrawingItems(pkg) {
  const items = [];
  for (const part of pkg.parts.values()) {
    if (!isStoryXmlPart(part)) continue;
    const partPath = normalizePartPath(part.path);
    const rels = parsePartRelationships(pkg, partPath);
    const xml = serializeOptionalXmlPart(part);
    drawingXmlFragments(xml).forEach((drawingXml, index) => {
      items.push(analyzeDrawingXml(pkg, partPath, index, rels, drawingXml));
    });
  }
  return items;
}

function stripInternalDiagnosticFields(item) {
  const { rasterImageId, ...publicItem } = item;
  return publicItem;
}

function enrichMediaAndDrawingGaps(pkg, gaps) {
  const drawingItems = collectDrawingItems(pkg);
  if (drawingItems.length > 0) {
    gaps.drawings.present = true;
    gaps.drawings.count = drawingItems.length;
    gaps.drawings.items = drawingItems.map(stripInternalDiagnosticFields);
    gaps.drawings.inlinedCount = drawingItems.filter((item) => item.status === 'inlined').length;
    gaps.drawings.unsupportedCount = drawingItems.filter(
      (item) => item.status === 'unsupported',
    ).length;
    gaps.drawings.status = gaps.drawings.unsupportedCount > 0 ? 'pending' : 'payload-ready';
  }

  const mediaItems = [...pkg.parts.values()]
    .filter((part) => part.path.startsWith('word/media/'))
    .map((part) => {
      const partPath = normalizePartPath(part.path);
      const references = drawingItems.filter((item) => item.target === partPath);
      const unsupportedReferences = references.filter((item) => item.status === 'unsupported');
      const inlinedReferences = references.filter((item) => item.status === 'inlined');
      return {
        path: partPath,
        contentType: part.contentType,
        status:
          references.length === 0
            ? 'unreferenced'
            : unsupportedReferences.length > 0
              ? 'referenced-by-unsupported-drawing'
              : 'inlined',
        references: references.map((item) => ({
          partPath: item.partPath,
          index: item.index,
          drawingType: item.drawingType,
          status: item.status,
          unsupportedReason: item.unsupportedReason,
        })),
        inlinedReferences: inlinedReferences.length,
        unsupportedReferences: unsupportedReferences.length,
      };
    });

  if (mediaItems.length > 0) {
    gaps.media.items = mediaItems;
    gaps.media.inlinedCount = mediaItems.filter((item) => item.status === 'inlined').length;
    gaps.media.unsupportedCount = mediaItems.filter((item) => item.status !== 'inlined').length;
    gaps.media.status = gaps.media.unsupportedCount > 0 ? 'pending' : 'payload-ready';
  }
}

function gapHasUnsupported(domain, gap) {
  if (!gap.present) return false;
  if (domain === 'media' || domain === 'drawings') return Number(gap.unsupportedCount || 0) > 0;
  return true;
}

function buildUnsupportedItems(gaps) {
  const items = [];
  for (const drawing of gaps.drawings.items || []) {
    if (drawing.status !== 'unsupported') continue;
    items.push({
      domain: 'drawings',
      partPath: drawing.partPath,
      index: drawing.index,
      status: drawing.status,
      reason: drawing.unsupportedReason,
      drawingType: drawing.drawingType,
      relationshipId: drawing.relationshipId,
      target: drawing.target,
      targetMode: drawing.targetMode,
      contentType: drawing.contentType,
      graphicDataUri: drawing.graphicDataUri,
    });
  }
  for (const media of gaps.media.items || []) {
    if (media.status === 'inlined') continue;
    items.push({
      domain: 'media',
      partPath: media.path,
      status: media.status,
      reason: media.status,
      contentType: media.contentType,
      references: media.references,
    });
  }
  for (const section of gaps.headersFooters.sections || []) {
    const numbering = section.pageNumbering;
    if (numbering && (numbering.fmt || numbering.chapStyle || numbering.chapSep)) {
      items.push({
        domain: 'headersFooters',
        index: section.index,
        status: 'unsupported',
        reason: 'page-numbering-format',
        property: 'pageNumbering',
        value: {
          fmt: numbering.fmt,
          chapStyle: numbering.chapStyle,
          chapSep: numbering.chapSep,
        },
      });
    }
    for (const kind of ['headers', 'footers']) {
      for (const slot of Object.values(section[kind] || {})) {
        for (const field of slot.unsupportedFields || []) {
          items.push({
            domain: 'headersFooters',
            partPath: slot.partPath,
            index: field.index,
            status: 'unsupported',
            reason: field.unsupportedReason,
            fieldType: field.fieldType,
            instruction: field.instruction,
          });
        }
      }
    }
  }
  return items;
}

function unsupportedSummary(items) {
  if (items.length === 0) return '';
  const counts = new Map();
  for (const item of items) {
    const key = `${item.domain}:${item.reason || item.status || 'unsupported'}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return [...counts.entries()].map(([key, count]) => `${key}=${count}`).join(', ');
}

function inspectPackage(pkg) {
  const gaps = Object.fromEntries(
    GAP_DOMAIN_ORDER.map((domain) => [domain, emptyGapDomain(domain)]),
  );

  for (const part of pkg.parts.values()) {
    if (part.contentType === CT.comments) addGapPart(gaps, 'comments', part);
    if (part.contentType === CT.commentsExtended) addGapPart(gaps, 'comments', part);
    if (part.contentType === CT.header || part.contentType === CT.footer)
      addGapPart(gaps, 'headersFooters', part);
    if (part.contentType === CT.footnotes || part.contentType === CT.endnotes)
      addGapPart(gaps, 'notes', part);
    if (part.path.startsWith('word/media/')) addGapPart(gaps, 'media', part);
    if (part.path.startsWith('word/drawings/') || part.path.includes('/drawings/'))
      addGapPart(gaps, 'drawings', part);
  }
  enrichMediaAndDrawingGaps(pkg, gaps);

  const unsupportedDomains = GAP_DOMAIN_ORDER.filter((domain) =>
    gapHasUnsupported(domain, gaps[domain]),
  );
  return { gaps, unsupportedDomains };
}

/**
 * Convert a complete after-state Flat OPC package into the OnlyOffice JSON body
 * payload consumed by `Api.FromJSON`.
 *
 * This adapter is the writeback translation seam:
 *   Flat OPC -> core package model -> oojson body JSON + gap diagnostics.
 *
 * The OnlyOffice host must treat the returned data as an execution plan and
 * must not parse OOXML itself.
 */
export function flatOpcToOnlyOfficeWriteback(flatOpc, options = {}) {
  const pkg = fromFlatOpc(flatOpc);
  const documentPart = pkg.parts.get('word/document.xml');
  const documentXml = serializeXmlPart(documentPart);
  const { gaps, unsupportedDomains } = inspectPackage(pkg);
  const commentItems = buildCommentItems(pkg, documentXml);
  if (commentItems.length > 0) {
    gaps.comments.items = commentItems;
    gaps.comments.status = 'payload-ready';
  }
  const noteItems = buildNoteItems(pkg, documentXml);
  if (noteItems.length > 0) {
    gaps.notes.items = noteItems;
    gaps.notes.status = 'payload-ready';
  }
  const bookmarkItems = parseBookmarkItems(documentXml);
  if (bookmarkItems.length > 0) {
    gaps.bookmarks.present = true;
    gaps.bookmarks.count = bookmarkItems.length;
    gaps.bookmarks.items = bookmarkItems;
    gaps.bookmarks.status = 'payload-ready';
  }
  const headerFooterItems = buildHeaderFooterItems(pkg, documentXml);
  const headerFooterSections = buildHeaderFooterSections(pkg, documentXml);
  if (headerFooterItems.length > 0) {
    gaps.headersFooters.items = headerFooterItems;
    gaps.headersFooters.status = 'payload-ready';
  }
  if (headerFooterSections.length > 0) {
    gaps.headersFooters.sections = headerFooterSections;
    gaps.headersFooters.status = 'payload-ready';
  }
  const warnings = [];

  if (unsupportedDomains.length > 0) {
    warnings.push(
      `OnlyOffice writeback adapter produced bodyJson only; gap domains pending rehydration: ${unsupportedDomains.join(', ')}`,
    );
  }
  const unsupportedItems = buildUnsupportedItems(gaps);
  const unsupportedItemSummary = unsupportedSummary(unsupportedItems);
  if (unsupportedItemSummary) {
    warnings.push(`OnlyOffice writeback unsupported items: ${unsupportedItemSummary}`);
  }

  const bodyDocumentXml = inlineStoryImages(pkg, 'word/document.xml', documentXml);
  const bodyJson = stripBookmarkMarkers(
    stripCommentMarkers(
      ooxmlToJson(bodyDocumentXml, undefined, { relationships: parseDocumentRelationships(pkg) })
        .doc,
    ),
  );
  if (options.failOnUnsupported === true && unsupportedDomains.length > 0) {
    throw new Error(
      `OnlyOffice writeback has unsupported gap domains: ${unsupportedDomains.join(', ')}`,
    );
  }

  return {
    bodyJson,
    gaps,
    diagnostics: {
      warnings,
      unsupportedDomains,
      unsupportedItems,
    },
  };
}
