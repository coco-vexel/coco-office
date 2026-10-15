import JSZip from 'jszip';

export async function summarizePptFixture(
  Presentation,
  name,
  bytes,
  markerPrefix = 'COCO_SUMMARY',
) {
  const presentation = await Presentation.fromZip(bytes, { presentationId: `${name}:summary` });
  const model = await presentation.toModel();
  const originalEntries = await zipEntryBytes(bytes);
  const roundTripEntries = await zipEntryBytes(await presentation.toZip());
  const preservedParts = await selectedZipEntryBytes(bytes, isAdvancedPptPart);

  const editPresentation = await Presentation.fromZip(bytes, {
    presentationId: `${name}:summary-edit`,
  });
  await editPresentation.slides
    .getItem(0)
    .shapes.addTextBox(`${markerPrefix}_${name}`, { name: 'CocoSummaryTextBox' });
  const editedEntries = await zipEntryBytes(await editPresentation.toZip());

  return {
    name,
    slideCount: model.slides.length,
    shapeCount: model.slides.reduce(
      (count, slide) => count + flattenShapes(slide.shapes).length,
      0,
    ),
    shapeKinds: summarizeShapeKinds(model),
    advancedShapes: summarizeAdvancedShapes(model),
    packageParts: await packageAdvancedCounts(bytes),
    motion: await inspectSlideMotion(bytes),
    editableCandidates: editableCandidates(model),
    noEditRoundTripPreserved: entriesEqual(originalEntries, roundTripEntries),
    advancedPartsPreservedAfterSdkWrites: [...preservedParts].every(([path, originalBytes]) =>
      bytesEqual(editedEntries.get(path), originalBytes),
    ),
    advancedShapesPreservedOnEditedSlide: await advancedShapesPreservedOnEditedSlide(
      Presentation,
      name,
      bytes,
      preservedParts,
    ),
  };
}

export function summarizeShapeKinds(model) {
  const counts = {};
  for (const slide of model.slides) {
    for (const shape of flattenShapes(slide.shapes)) {
      counts[shape.kind] = (counts[shape.kind] ?? 0) + 1;
    }
  }
  return counts;
}

export function flattenShapes(shapes) {
  return shapes.flatMap((shape) => [shape, ...flattenShapes(shape.children ?? [])]);
}

export async function packageAdvancedCounts(bytes) {
  const entries = await zipEntryBytes(bytes);
  const mediaPaths = [...entries.keys()].filter((path) => path.startsWith('ppt/media/'));
  return {
    chartParts: countEntries(entries, (path) => path.startsWith('ppt/charts/')),
    mediaParts: mediaPaths.length,
    imageParts: mediaPaths.filter(isImageMediaPart).length,
    videoParts: mediaPaths.filter(isVideoMediaPart).length,
    audioParts: mediaPaths.filter(isAudioMediaPart).length,
    otherMediaParts: mediaPaths.filter(
      (path) => !isImageMediaPart(path) && !isVideoMediaPart(path) && !isAudioMediaPart(path),
    ).length,
    embeddingParts: countEntries(entries, (path) => path.startsWith('ppt/embeddings/')),
    diagramParts: countEntries(entries, (path) => path.startsWith('ppt/diagrams/')),
  };
}

export function summarizeAdvancedShapes(model) {
  const summary = {
    chartShapes: 0,
    mediaShapes: 0,
    oleObjectShapes: 0,
    smartArtShapes: 0,
    unknownShapes: 0,
  };

  for (const slide of model.slides) {
    for (const shape of slide.shapes) {
      addAdvancedShapeSummary(summary, shape);
    }
  }

  return summary;
}

export async function advancedShapesPreservedOnEditedSlide(
  Presentation,
  name,
  bytes,
  preservedParts,
) {
  const presentation = await Presentation.fromZip(bytes, {
    presentationId: `${name}:summary-advanced-slide-edit`,
  });
  const model = await presentation.toModel();
  const advancedSlide = model.slides.find((slide) => advancedShapeSnapshots(slide).length > 0);
  if (!advancedSlide) return null;

  const beforeAdvancedShapes = advancedShapeSnapshots(advancedSlide);
  const slide = presentation.slides.getItem(advancedSlide.id);
  const textCandidate = advancedSlide.shapes.find(
    (shape) => shape.text?.text.trim() && !shape.table && !isAdvancedShapeKind(shape.kind),
  );
  if (textCandidate) {
    await slide.shapes.getItem(textCandidate.id).text.setText(`COCO_ADVANCED_SAFE_TEXT_${name}`);
  } else {
    await slide.shapes.addTextBox(`COCO_ADVANCED_SAFE_TEXT_${name}`, {
      name: `CocoAdvancedSafeText_${advancedSlide.index}`,
    });
  }

  const output = await presentation.toZip();
  const editedEntries = await zipEntryBytes(output);
  const advancedPartsPreserved = [...preservedParts].every(([path, originalBytes]) =>
    bytesEqual(editedEntries.get(path), originalBytes),
  );
  if (!advancedPartsPreserved) return false;

  const reloaded = await Presentation.fromZip(output, {
    presentationId: `${name}:summary-advanced-slide-edit-roundtrip`,
  });
  const reloadedSlide = (await reloaded.toModel()).slides.find(
    (slideModel) => slideModel.id === advancedSlide.id,
  );
  return (
    Boolean(reloadedSlide) && deepEqual(advancedShapeSnapshots(reloadedSlide), beforeAdvancedShapes)
  );
}

export async function selectedZipEntryBytes(bytes, predicate) {
  const entries = await zipEntryBytes(bytes);
  return new Map([...entries].filter(([path]) => predicate(path)));
}

export async function zipEntryBytes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const entries = new Map();
  const paths = [];
  zip.forEach((relativePath, entry) => {
    if (!entry.dir) paths.push(relativePath);
  });
  for (const path of paths.sort()) {
    const entry = zip.file(path);
    if (entry) entries.set(path, await entry.async('uint8array'));
  }
  return entries;
}

export async function inspectSlideMotion(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  let transitions = 0;
  let timing = 0;
  const slidePaths = Object.keys(zip.files).filter(
    (path) => /^ppt\/slides\/slide\d+\.xml$/.test(path) && !zip.files[path].dir,
  );
  for (const path of slidePaths) {
    const xml = await zip.files[path].async('string');
    if (xml.includes('<p:transition')) transitions += 1;
    if (xml.includes('<p:timing')) timing += 1;
  }
  return { transitions, timing };
}

export function isAdvancedPptPart(path) {
  return /^ppt\/(?:media|charts|embeddings|diagrams)\//.test(path);
}

export function entriesEqual(left, right) {
  if (left.size !== right.size) return false;
  for (const [path, bytes] of left) {
    if (!bytesEqual(right.get(path), bytes)) return false;
  }
  return true;
}

export function bytesEqual(actual, expected) {
  if (!actual || actual.byteLength !== expected.byteLength) return false;
  for (let index = 0; index < expected.byteLength; index++) {
    if (actual[index] !== expected[index]) return false;
  }
  return true;
}

function editableCandidates(model) {
  return {
    text: model.slides.some((slide) =>
      flattenShapes(slide.shapes).some((shape) => shape.text?.text.trim() && !shape.table),
    ),
    table: model.slides.some((slide) => flattenShapes(slide.shapes).some((shape) => shape.table)),
    image: model.slides.some((slide) =>
      flattenShapes(slide.shapes).some(
        (shape) => shape.kind === 'picture' && Boolean(shape.image?.targetPath),
      ),
    ),
  };
}

function countEntries(entries, predicate) {
  return [...entries.keys()].filter(predicate).length;
}

function isImageMediaPart(path) {
  return /\.(?:avif|bmp|emf|gif|jpeg|jpg|png|svg|tif|tiff|webp|wmf)$/i.test(path);
}

function isVideoMediaPart(path) {
  return /\.(?:avi|m4v|mov|mp4|mpeg|mpg|wmv)$/i.test(path);
}

function isAudioMediaPart(path) {
  return /\.(?:aac|aiff|m4a|mid|midi|mp3|wav|wma)$/i.test(path);
}

function addAdvancedShapeSummary(summary, shape) {
  if (shape.kind === 'chart') summary.chartShapes++;
  if (shape.kind === 'media') summary.mediaShapes++;
  if (shape.kind === 'oleObject') summary.oleObjectShapes++;
  if (shape.kind === 'smartArt') summary.smartArtShapes++;
  if (shape.kind === 'unknown') summary.unknownShapes++;
  for (const child of shape.children ?? []) {
    addAdvancedShapeSummary(summary, child);
  }
}

function advancedShapeSnapshots(slide) {
  return slide.shapes.flatMap((shape) => advancedShapeSnapshotsFromShape(shape));
}

function advancedShapeSnapshotsFromShape(shape) {
  const own = isAdvancedShapeKind(shape.kind)
    ? [
        {
          id: shape.id,
          name: shape.name,
          kind: shape.kind,
          geometry: shape.geometry,
          chart: shape.chart,
          media: shape.media,
          oleObject: shape.oleObject,
          smartArt: shape.smartArt,
          raw: shape.raw,
        },
      ]
    : [];
  return [
    ...own,
    ...(shape.children ?? []).flatMap((child) => advancedShapeSnapshotsFromShape(child)),
  ];
}

function isAdvancedShapeKind(kind) {
  return (
    kind === 'chart' ||
    kind === 'media' ||
    kind === 'oleObject' ||
    kind === 'smartArt' ||
    kind === 'unknown'
  );
}

function deepEqual(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}
