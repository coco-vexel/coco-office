import {
  onlyOfficePptContentToTextBody,
  setOnlyOfficePptContentParagraphs,
  setOnlyOfficePptContentStyle,
  setOnlyOfficePptContentText,
} from "./ppt-content.mjs";

function clone(value) {
  return value == null ? value : JSON.parse(JSON.stringify(value));
}

function parseJson(value, label = "ONLYOFFICE PPT JSON") {
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

function normalizeSlidesRoot(value) {
  const root = parseJson(value, "ONLYOFFICE PPT deck JSON");
  if (!Array.isArray(root.slides)) {
    throw new Error("ONLYOFFICE PPT deck JSON: expected slides[]");
  }
  return root;
}

function shapeId(shape, index) {
  return String(shape?.id ?? shape?.nvSpPr?.cNvPr?.id ?? index);
}

function shapeName(shape) {
  return shape?.nvSpPr?.cNvPr?.name;
}

function shapeGeometry(shape = {}) {
  const xfrm = shape.spPr?.xfrm || {};
  const off = xfrm.off || {};
  const ext = xfrm.ext || {};
  return {
    left: Number.isFinite(Number(off.x)) ? Number(off.x) : undefined,
    top: Number.isFinite(Number(off.y)) ? Number(off.y) : undefined,
    width: Number.isFinite(Number(ext.cx ?? shape.extX)) ? Number(ext.cx ?? shape.extX) : undefined,
    height: Number.isFinite(Number(ext.cy ?? shape.extY)) ? Number(ext.cy ?? shape.extY) : undefined,
    rotation: Number.isFinite(Number(xfrm.rot)) ? Number(xfrm.rot) : undefined,
    flipHorizontal: xfrm.flipH === true ? true : undefined,
    flipVertical: xfrm.flipV === true ? true : undefined,
  };
}

function shapeKind(shape = {}) {
  if (shape.type === "group") return "group";
  if (shape.type === "image" || shape.blipFill) return "picture";
  if (shape.graphicObject?.table || shape.table) return "table";
  if (shape.graphicObject?.chart || shape.chart) return "chart";
  if (shapeDocContent(shape)) return "textBox";
  if (shape.spPr?.custGeom || shape.spPr?.prstGeom || shape.spPr?.geometry) return "geometricShape";
  return "unknown";
}

function slideSize(root) {
  const size = root.sldSz || {};
  const width = Number(size.cx ?? size.w ?? size.width);
  const height = Number(size.cy ?? size.h ?? size.height);
  if (!Number.isFinite(width) || !Number.isFinite(height)) return undefined;
  return { width, height };
}

function textBodyFromShape(shape) {
  const docContent = shapeDocContent(shape);
  if (!docContent) return undefined;
  try {
    return onlyOfficePptContentToTextBody(docContent);
  } catch {
    return undefined;
  }
}

function shapeDocContent(shape) {
  const content = shape?.content;
  if (!content || typeof content !== "object") return undefined;
  if (content.type === "docContent" || content.type === "documentContent") return content;
  const inner = content.content;
  if (inner && typeof inner === "object" && (inner.type === "docContent" || inner.type === "documentContent")) {
    return inner;
  }
  if (inner && typeof inner === "object" && Array.isArray(inner.content)) {
    return { ...inner, type: "docContent", bPresentation: inner.bPresentation ?? true };
  }
  return undefined;
}

function setShapeDocContent(shape, docContent) {
  if (!shape?.content || !docContent) return false;
  if (shape.content.type === "docContent" || shape.content.type === "documentContent") {
    shape.content = docContent;
    return true;
  }
  shape.content = {
    ...shape.content,
    content: docContent,
  };
  return true;
}

function titleFromShapes(shapes) {
  for (const shape of shapes) {
    const text = shape.text?.text?.trim();
    if (text) return text;
  }
  return undefined;
}

function ooShapeToCoreShape(shape, index) {
  const text = textBodyFromShape(shape);
  const children = Array.isArray(shape.spTree)
    ? shape.spTree.map((child, childIndex) => ooShapeToCoreShape(child, childIndex))
    : undefined;
  return {
    id: shapeId(shape, index),
    name: shapeName(shape),
    kind: shapeKind(shape),
    geometry: shapeGeometry(shape),
    text,
    children,
    raw: {
      hostType: "onlyoffice-ppt-oojson",
      presetGeometry: shape.spPr?.prstGeom?.prst || shape.spPr?.custGeom?.preset,
    },
  };
}

export function onlyOfficePptDeckToPptModel(deckJson, options = {}) {
  const root = normalizeSlidesRoot(deckJson);
  const slides = root.slides.map((slide, slideIndex) => {
    const spTree = Array.isArray(slide?.cSld?.spTree) ? slide.cSld.spTree : [];
    const shapes = spTree.map((shape, shapeIndex) => ooShapeToCoreShape(shape, shapeIndex));
    return {
      id: String(slide?.id ?? slideIndex),
      index: slideIndex,
      name: slide?.name,
      title: titleFromShapes(shapes),
      shapes,
    };
  });
  return {
    id: options.id,
    title: options.title,
    slideSize: slideSize(root),
    slides,
  };
}

function findSlide(root, modelSlide) {
  if (Number.isInteger(modelSlide?.index) && root.slides[modelSlide.index]) {
    return root.slides[modelSlide.index];
  }
  return root.slides.find((slide) => String(slide?.id) === String(modelSlide?.id));
}

function findShape(spTree, modelShape) {
  if (!Array.isArray(spTree)) return undefined;
  const requestedId = String(modelShape?.id ?? "");
  const requestedInternalId = requestedId.includes("/shape:")
    ? requestedId.slice(requestedId.lastIndexOf("/shape:") + "/shape:".length)
    : requestedId;
  return spTree.find((shape, index) => {
    const id = shapeId(shape, index);
    return (
      id === requestedId ||
      id === requestedInternalId ||
      (shapeName(shape) && shapeName(shape) === modelShape?.name)
    );
  });
}

function applyTextBodyToShape(shape, textBody) {
  const docContent = shapeDocContent(shape);
  if (!shape || !textBody || !docContent) return false;
  let next = docContent;
  if (Array.isArray(textBody.paragraphs) && textBody.paragraphs.length) {
    next = setOnlyOfficePptContentParagraphs(next, textBody.paragraphs);
  } else if (typeof textBody.text === "string") {
    next = setOnlyOfficePptContentText(next, textBody.text, { style: textBody.style });
  }
  if (textBody.style) {
    next = setOnlyOfficePptContentStyle(next, textBody.style);
  }
  return setShapeDocContent(shape, next);
}

export function applyPptModelToOnlyOfficePptDeck(deckJson, model) {
  const root = normalizeSlidesRoot(deckJson);
  let changedCount = 0;
  const changed = [];
  for (const modelSlide of model?.slides || []) {
    const slide = findSlide(root, modelSlide);
    const spTree = slide?.cSld?.spTree;
    if (!Array.isArray(spTree)) continue;
    for (const modelShape of modelSlide.shapes || []) {
      const shape = findShape(spTree, modelShape);
      if (!shape || !modelShape.text) continue;
      if (applyTextBodyToShape(shape, modelShape.text)) {
        changedCount += 1;
        changed.push({
          slideIndex: modelSlide.index,
          shapeId: modelShape.id,
          shapeName: modelShape.name,
        });
      }
    }
  }
  return { deck: root, changedCount, changed };
}

export function updateOnlyOfficePptDeckTextShape(deckJson, target, textBody) {
  const model = onlyOfficePptDeckToPptModel(deckJson);
  const slide = model.slides[target?.slideIndex ?? 0];
  if (!slide) return { deck: normalizeSlidesRoot(deckJson), changedCount: 0, changed: [] };
  const shape = slide.shapes.find((item) => {
    const itemInternalId = item.id.includes("/shape:")
      ? item.id.slice(item.id.lastIndexOf("/shape:") + "/shape:".length)
      : item.id;
    const targetId = target?.shapeId == null ? undefined : String(target.shapeId);
    const targetInternalId = targetId?.includes("/shape:")
      ? targetId.slice(targetId.lastIndexOf("/shape:") + "/shape:".length)
      : targetId;
    return (
      item.id === targetId ||
      itemInternalId === targetInternalId ||
      (target?.shapeName && item.name === target.shapeName) ||
      (target?.text && item.text?.text === target.text)
    );
  });
  if (!shape) return { deck: normalizeSlidesRoot(deckJson), changedCount: 0, changed: [] };
  const nextText = { ...(shape.text || { text: "" }), ...(textBody || {}) };
  if (typeof textBody?.text === "string" && !Array.isArray(textBody.paragraphs)) {
    delete nextText.paragraphs;
  }
  shape.text = nextText;
  return applyPptModelToOnlyOfficePptDeck(deckJson, { slides: [slide] });
}
