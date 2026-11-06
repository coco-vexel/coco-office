import type { Package } from '../pack/index.js';

const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const RELATIONSHIP_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const DRAWING_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELATIONSHIP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface ExcelWorksheetLineOptions {
  name: string;
  from: { column: number; columnOffset: number; row: number; rowOffset: number };
  to: { column: number; columnOffset: number; row: number; rowOffset: number };
  width?: number;
  height?: number;
  color?: string;
  alpha?: number;
  lineWidth?: number;
  dash?: 'solid' | 'dash' | 'dot';
}

export interface ExcelWorksheetPrintSetupOptions {
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  fitToWidth?: number;
  fitToHeight?: number;
  horizontalCentered?: boolean;
  margins?: Partial<Record<'left' | 'right' | 'top' | 'bottom' | 'header' | 'footer', number>>;
  printArea?: string;
}

export function addWorksheetLinesToPackage(
  pkg: Package,
  sheetPath: string,
  lines: readonly ExcelWorksheetLineOptions[],
): void {
  if (lines.length === 0) return;
  const sheetPart = requireXmlPart(pkg, sheetPath);
  const sheetDoc = sheetPart.xmlDom!;
  if (elementChildren(sheetDoc.documentElement).some((child) => child.localName === 'drawing')) {
    throw new Error('Adding lines to a worksheet that already has drawings is not supported yet.');
  }
  const drawingPath = nextDrawingPath(pkg);
  const drawing = sheetDoc.createElementNS(sheetDoc.documentElement.namespaceURI, 'drawing');
  const relationshipPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
  const relationshipPart = pkg.parts.get(relationshipPath);
  const relationshipDoc = relationshipPart?.xmlDom ??
    sheetDoc.implementation.createDocument(RELATIONSHIP_NS, 'Relationships', null);
  const relationshipId = nextRelationshipId(relationshipDoc, 'rIdCocoDrawing');
  drawing.setAttributeNS(OFFICE_RELATIONSHIP_NS, 'r:id', relationshipId);
  sheetDoc.documentElement.appendChild(drawing);
  sheetPart._dirty = true;

  const relationship = relationshipDoc.createElementNS(RELATIONSHIP_NS, 'Relationship');
  relationship.setAttribute('Id', relationshipId);
  relationship.setAttribute('Type', DRAWING_RELATIONSHIP_TYPE);
  relationship.setAttribute('Target', `../drawings/${drawingPath.split('/').pop()}`);
  relationshipDoc.documentElement.appendChild(relationship);
  pkg.parts.set(relationshipPath, {
    path: relationshipPath,
    contentType: RELATIONSHIP_CONTENT_TYPE,
    xmlDom: relationshipDoc,
    _dirty: true,
  });
  pkg.parts.set(drawingPath, {
    path: drawingPath,
    contentType: DRAWING_CONTENT_TYPE,
    bytes: new TextEncoder().encode(serializeLines(lines)),
    _dirty: true,
  });
  ensureContentTypeOverride(pkg, drawingPath, DRAWING_CONTENT_TYPE);
}

export function setWorksheetPrintSetupInPackage(
  pkg: Package,
  sheetPath: string,
  sheetName: string,
  localSheetId: number,
  options: ExcelWorksheetPrintSetupOptions,
): void {
  const sheetPart = requireXmlPart(pkg, sheetPath);
  const sheetDoc = sheetPart.xmlDom!;
  const root = sheetDoc.documentElement;
  const namespace = root.namespaceURI;
  for (const child of elementChildren(root)) {
    if (['printOptions', 'pageMargins', 'pageSetup'].includes(child.localName)) root.removeChild(child);
  }
  let sheetPr = elementChildren(root).find((child) => child.localName === 'sheetPr');
  if (!sheetPr) {
    sheetPr = sheetDoc.createElementNS(namespace, 'sheetPr');
    root.insertBefore(sheetPr, root.firstChild);
  }
  for (const child of elementChildren(sheetPr)) {
    if (child.localName === 'pageSetUpPr') sheetPr.removeChild(child);
  }
  const pageSetUpPr = sheetDoc.createElementNS(namespace, 'pageSetUpPr');
  pageSetUpPr.setAttribute('fitToPage', '1');
  sheetPr.appendChild(pageSetUpPr);

  const printOptions = sheetDoc.createElementNS(namespace, 'printOptions');
  printOptions.setAttribute('horizontalCentered', options.horizontalCentered === false ? '0' : '1');
  const pageMargins = sheetDoc.createElementNS(namespace, 'pageMargins');
  const margins = { left: 0, right: 0, top: 0, bottom: 0, header: 0, footer: 0, ...options.margins };
  for (const [name, value] of Object.entries(margins)) pageMargins.setAttribute(name, String(value));
  const pageSetup = sheetDoc.createElementNS(namespace, 'pageSetup');
  pageSetup.setAttribute('paperSize', String(options.paperSize ?? 9));
  pageSetup.setAttribute('orientation', options.orientation ?? 'portrait');
  pageSetup.setAttribute('fitToWidth', String(options.fitToWidth ?? 1));
  pageSetup.setAttribute('fitToHeight', String(options.fitToHeight ?? 1));
  const insertionPoint = elementChildren(root).find((child) =>
    ['drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'].includes(child.localName),
  ) ?? null;
  root.insertBefore(printOptions, insertionPoint);
  root.insertBefore(pageMargins, insertionPoint);
  root.insertBefore(pageSetup, insertionPoint);
  sheetPart._dirty = true;
  if (options.printArea) setPrintArea(pkg, sheetName, localSheetId, options.printArea);
}

function setPrintArea(pkg: Package, sheetName: string, localSheetId: number, range: string): void {
  const workbookPart = requireXmlPart(pkg, 'xl/workbook.xml');
  const workbookDoc = workbookPart.xmlDom!;
  const workbookRoot = workbookDoc.documentElement;
  let definedNames = elementChildren(workbookRoot).find((child) => child.localName === 'definedNames');
  if (!definedNames) {
    definedNames = workbookDoc.createElementNS(workbookRoot.namespaceURI, 'definedNames');
    const insertionPoint = elementChildren(workbookRoot).find((child) =>
      ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'].includes(child.localName),
    ) ?? null;
    workbookRoot.insertBefore(definedNames, insertionPoint);
  }
  for (const child of elementChildren(definedNames)) {
    if (child.localName === 'definedName' && child.getAttribute('name') === '_xlnm.Print_Area' && child.getAttribute('localSheetId') === String(localSheetId)) {
      definedNames.removeChild(child);
    }
  }
  const printArea = workbookDoc.createElementNS(workbookRoot.namespaceURI, 'definedName');
  printArea.setAttribute('name', '_xlnm.Print_Area');
  printArea.setAttribute('localSheetId', String(localSheetId));
  printArea.appendChild(workbookDoc.createTextNode(`'${sheetName.replace(/'/g, "''")}'!${range}`));
  definedNames.appendChild(printArea);
  workbookPart._dirty = true;
}

function serializeLines(lines: readonly ExcelWorksheetLineOptions[]): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    lines.map((line, index) => serializeLine(line, index + 2)).join('') + '</xdr:wsDr>';
}

function serializeLine(line: ExcelWorksheetLineOptions, id: number): string {
  const color = (line.color ?? '#4F9185').replace(/^#/, '').toUpperCase();
  const dash = line.dash === 'solid' ? 'solid' : line.dash ?? 'dash';
  return '<xdr:twoCellAnchor editAs="oneCell">' + marker('from', line.from) + marker('to', line.to) +
    '<xdr:sp macro="" textlink=""><xdr:nvSpPr>' +
    `<xdr:cNvPr id="${id}" name="${xmlAttr(line.name)}"/><xdr:cNvSpPr><a:spLocks noSelect="1" noMove="1" noResize="1"/></xdr:cNvSpPr><xdr:nvPr/></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${line.width ?? 0}" cy="${line.height ?? 0}"/></a:xfrm>` +
    '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>' +
    `<a:ln w="${line.lineWidth ?? 9525}" cap="flat"><a:solidFill><a:srgbClr val="${color}"><a:alpha val="${line.alpha ?? 72000}"/></a:srgbClr></a:solidFill><a:prstDash val="${dash}"/></a:ln>` +
    '</xdr:spPr><xdr:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p/></xdr:txBody></xdr:sp><xdr:clientData fLocksWithSheet="1" fPrintsWithSheet="1"/></xdr:twoCellAnchor>';
}

function marker(kind: 'from' | 'to', value: ExcelWorksheetLineOptions['from']): string {
  return `<xdr:${kind}><xdr:col>${value.column}</xdr:col><xdr:colOff>${value.columnOffset}</xdr:colOff><xdr:row>${value.row}</xdr:row><xdr:rowOff>${value.rowOffset}</xdr:rowOff></xdr:${kind}>`;
}

function requireXmlPart(pkg: Package, path: string) {
  const part = pkg.parts.get(path);
  if (!part?.xmlDom) throw new Error(`Worksheet layout requires a package-backed workbook part: ${path}`);
  return part;
}

function elementChildren(node: Node): Element[] {
  return Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1);
}

function nextDrawingPath(pkg: Package): string {
  for (let index = 1; index < 1000; index += 1) {
    const path = `xl/drawings/drawing${index}.xml`;
    if (!pkg.parts.has(path)) return path;
  }
  throw new Error('Unable to allocate an Excel drawing part.');
}

function nextRelationshipId(document: Document, baseId: string): string {
  const used = new Set(elementChildren(document.documentElement).map((node) => node.getAttribute('Id')).filter(Boolean));
  if (!used.has(baseId)) return baseId;
  for (let index = 2; index < 1000; index += 1) {
    const candidate = `${baseId}${index}`;
    if (!used.has(candidate)) return candidate;
  }
  throw new Error('Unable to allocate an Excel drawing relationship.');
}

function ensureContentTypeOverride(pkg: Package, path: string, contentType: string): void {
  const typesPart = requireXmlPart(pkg, '[Content_Types].xml');
  const typesDoc = typesPart.xmlDom!;
  const partName = `/${path}`;
  const exists = elementChildren(typesDoc.documentElement).some((child) => child.localName === 'Override' && child.getAttribute('PartName') === partName);
  if (exists) return;
  const override = typesDoc.createElementNS(typesDoc.documentElement.namespaceURI, 'Override');
  override.setAttribute('PartName', partName);
  override.setAttribute('ContentType', contentType);
  typesDoc.documentElement.appendChild(override);
  typesPart._dirty = true;
}

function xmlAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
