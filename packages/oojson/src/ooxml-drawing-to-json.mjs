const LINE_PRESETS = new Set([
  "line", "straightConnector1",
  "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5",
  "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5",
]);

function findNode(node, tag) {
  if (!node || typeof node !== "object") return null;
  if (node.tag === tag) return node;
  for (const child of node.kids || []) {
    const found = findNode(child, tag);
    if (found) return found;
  }
  return null;
}

function nodeText(node) {
  if (!node || typeof node !== "object") return "";
  if (node.text != null) return String(node.text);
  return (node.kids || []).map(nodeText).join("");
}

function numberValue(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

function boolValue(value, fallback = false) {
  if (value == null || value === "") return fallback;
  return ["1", "true", "t", "on"].includes(String(value).toLowerCase());
}

function point(node, tag) {
  const value = findNode(node, tag);
  return { x: numberValue(value?.attrs?.x) ?? 0, y: numberValue(value?.attrs?.y) ?? 0 };
}

function extent(node, tag) {
  const value = findNode(node, tag);
  return { cx: numberValue(value?.attrs?.cx) ?? 0, cy: numberValue(value?.attrs?.cy) ?? 0 };
}

function position(container, tag, defaultRelativeFrom) {
  const value = findNode(container, tag);
  return {
    relativeFrom: String(value?.attrs?.relativeFrom ?? defaultRelativeFrom),
    posOffset: numberValue(nodeText(findNode(value, "wp:posOffset"))) ?? 0,
    percent: false,
  };
}

function solidFill(node) {
  const hex = String(findNode(node, "a:srgbClr")?.attrs?.val ?? "").trim();
  if (!/^[0-9a-f]{6}$/i.test(hex)) return null;
  const rgba = {
    red: Number.parseInt(hex.slice(0, 2), 16),
    green: Number.parseInt(hex.slice(2, 4), 16),
    blue: Number.parseInt(hex.slice(4, 6), 16),
    alpha: 255,
  };
  return {
    fill: {
      color: { rgba, color: { rgba, type: "srgb" }, mods: [], type: "uniColor" },
      type: "solid",
    },
    transparent: null,
    type: "fill",
    fillType: "solid",
  };
}

function presetGeometry(preset) {
  return {
    ahLst: { ahPolar: [], ahXY: [] },
    avLst: {},
    adjLst: {},
    cnxLst: [
      { pos: { x: "l", y: "t" }, ang: "cd4" },
      { pos: { x: "r", y: "b" }, ang: "_3cd4" },
    ],
    gdLst: [],
    pathLst: [
      {
        commands: [
          { id: "moveTo", pt: { x: "l", y: "t" } },
          { id: "lnTo", pt: { x: "r", y: "b" } },
        ],
        extrusionOk: false,
        fill: "norm",
        stroke: true,
      },
    ],
    rect: null,
    preset,
  };
}

function lineToJson(drawing) {
  const preset = String(findNode(drawing, "a:prstGeom")?.attrs?.prst ?? "");
  if (!LINE_PRESETS.has(preset)) return null;
  const anchor = findNode(drawing, "wp:anchor");
  const inline = findNode(drawing, "wp:inline");
  const container = anchor || inline;
  if (!container) return null;

  const docPr = findNode(container, "wp:docPr");
  const effect = findNode(container, "wp:effectExtent");
  const outerExtent = extent(container, "wp:extent");
  const xfrm = findNode(drawing, "a:xfrm");
  const innerExtent = extent(xfrm, "a:ext");
  const line = findNode(drawing, "a:ln");
  const simplePos = findNode(container, "wp:simplePos");
  const id = numberValue(docPr?.attrs?.id) ?? 1;
  const metadata = {
    id,
    name: String(docPr?.attrs?.name ?? `Line ${id}`),
    descr: docPr?.attrs?.descr == null ? null : String(docPr.attrs.descr),
    title: docPr?.attrs?.title == null ? null : String(docPr.attrs.title),
    hidden: null,
  };
  const shapeMetadata = {
    id: -1,
    name: "",
    descr: null,
    title: null,
    hidden: null,
  };
  const connectorLocks = findNode(drawing, "a:cxnSpLocks");

  return {
    type: "paraDrawing",
    docPr: metadata,
    effectExtent: {
      l: numberValue(effect?.attrs?.l) ?? 0, t: numberValue(effect?.attrs?.t) ?? 0,
      r: numberValue(effect?.attrs?.r) ?? 0, b: numberValue(effect?.attrs?.b) ?? 0,
    },
    extent: outerExtent,
    graphic: {
      bWordShape: true,
      extX: innerExtent.cx || outerExtent.cx,
      extY: innerExtent.cy || outerExtent.cy,
      nvCxnSpPr: {
        cNvPr: shapeMetadata,
        nvPr: { isPhoto: null, userDrawn: null },
        cNvCxnSpPr: {
          cxnSpLocks: {
            noChangeShapeType: boolValue(
              connectorLocks?.attrs?.noChangeShapeType,
              true,
            ),
          },
        },
      },
      spPr: {
        custGeom: presetGeometry(preset), effectDag: null, effectLst: null,
        ln: {
          fill: solidFill(line),
          lineJoin: findNode(line, "a:round") ? { type: "round", lim: null } : null,
          algn: null, cap: line?.attrs?.cap ?? null, cmpd: line?.attrs?.cmpd ?? null,
          w: numberValue(line?.attrs?.w) ?? 0,
          headEnd: null,
          prstDash: String(findNode(line, "a:prstDash")?.attrs?.val ?? "solid"),
          tailEnd: null, type: "stroke",
        },
        fill: {
          fill: { type: "noFill" }, transparent: null, type: "fill", fillType: "noFill",
        },
        xfrm: {
          ext: { cx: innerExtent.cx || outerExtent.cx, cy: innerExtent.cy },
          off: point(xfrm, "a:off"),
          flipH: xfrm?.attrs?.flipH == null ? null : boolValue(xfrm.attrs.flipH),
          flipV: xfrm?.attrs?.flipV == null ? null : boolValue(xfrm.attrs.flipV),
          rot: numberValue(xfrm?.attrs?.rot) ?? null,
          chOffX: null, chOffY: null, chExtX: null, chExtY: null,
        },
        bwMode: 0,
      },
      content: null,
      modelId: null,
      type: "connectShape",
      id: String(id),
    },
    cNvGraphicFramePr: { graphicFrameLocks: {} },
    positionH: position(container, "wp:positionH", "column"),
    positionV: position(container, "wp:positionV", "paragraph"),
    simplePos: {
      x: numberValue(simplePos?.attrs?.x) ?? 0,
      y: numberValue(simplePos?.attrs?.y) ?? 0,
      use: boolValue(container.attrs?.simplePos, false),
    },
    distB: numberValue(container.attrs?.distB) ?? 0,
    distL: numberValue(container.attrs?.distL) ?? 0,
    distR: numberValue(container.attrs?.distR) ?? 0,
    distT: numberValue(container.attrs?.distT) ?? 0,
    allowOverlap: boolValue(container.attrs?.allowOverlap, true),
    behindDoc: boolValue(container.attrs?.behindDoc, false),
    hidden: null,
    layoutInCell: boolValue(container.attrs?.layoutInCell, true),
    locked: boolValue(container.attrs?.locked, false),
    relativeHeight: numberValue(container.attrs?.relativeHeight) ?? 0,
    wrapType: "none",
    drawingType: anchor ? "anchor" : "inline",
  };
}

export function ooxmlDrawingToJson(node) {
  const drawing = node?.tag === "w:drawing" ? node : findNode(node, "w:drawing");
  return drawing ? lineToJson(drawing) : null;
}
