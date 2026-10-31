import {
  createPackage,
  createXmlPart,
  deletePart,
  getPart,
  markDirty,
  setPart,
  type Package,
  type Part,
} from '../pack/index.js';
import { parseXmlDocument } from 'slimdom';
import { OFFICE_REL_NS, REL_NS, SS_NS } from './namespaces.js';
import type {
  ExcelAlignmentModel,
  ExcelBorderModel,
  ExcelBorderSideModel,
  ExcelCellModel,
  ExcelChartModel,
  ExcelColumnModel,
  ExcelCommentModel,
  ExcelConditionalFormatModel,
  ExcelColorModel,
  ExcelDataValidationModel,
  ExcelDrawingMarkerModel,
  ExcelDrawingModel,
  ExcelFillModel,
  ExcelFontModel,
  ExcelHyperlinkModel,
  ExcelRowModel,
  ExcelSheetModel,
  ExcelSheetViewModel,
  ExcelStylesModel,
  ExcelTableModel,
  ExcelTableAutoFilterModel,
  ExcelTableFilterColumnModel,
  ExcelTableSortStateModel,
  ExcelTableSortConditionModel,
  ExcelThreadedCommentModel,
  ExcelWorkbookModel,
} from './model.js';
import { excelColorToOoxmlArgb } from './color.js';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CONTENT_TYPES_CONTENT_TYPE = 'application/xml';
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const TC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';
const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const THREADED_COMMENTS_REL_TYPE =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const TABLE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const WORKSHEET_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const DRAWING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawing+xml';
const DRAWING_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CHART_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export function syncWorkbookModelToPackage(pkg: Package, model: ExcelWorkbookModel): void {
  syncStylesPart(pkg, model.styles);
  syncWorkbookRelationships(pkg, model);
  syncWorkbookPart(pkg, model);
  for (const sheet of model.sheets) {
    syncWorksheetRelationships(pkg, sheet);
    syncWorksheetTableRelationships(pkg, sheet);
    syncWorksheetCommentRelationships(pkg, sheet);
    syncWorksheetThreadedCommentRelationships(pkg, sheet);
    syncWorksheetDrawingRelationships(pkg, sheet);
    syncWorksheetPart(pkg, sheet);
    syncTableParts(pkg, sheet);
    syncCommentParts(pkg, sheet);
    syncThreadedCommentParts(pkg, sheet);
    syncDrawingParts(pkg, sheet);
    syncChartParts(pkg, sheet);
  }
}

function syncWorkbookRelationships(pkg: Package, model: ExcelWorkbookModel): void {
  const relsPath = 'xl/_rels/workbook.xml.rels';
  const rels = parseRelationships(getPart(pkg, relsPath));
  const usedSheetRelIds = new Set<string>();
  let changed = false;

  for (const sheet of model.sheets) {
    sheet.path ??= nextIndexedPartPath(pkg, 'xl/worksheets/sheet', '.xml');
    const target = relativeRelationshipTarget('xl/workbook.xml', sheet.path);
    const id = sheet.relationshipId || findReusableRelationshipId(rels, WORKSHEET_REL_TYPE, target);
    if (id) {
      sheet.relationshipId = id;
    } else {
      sheet.relationshipId = nextRelationshipId(rels);
      changed = true;
    }
    usedSheetRelIds.add(sheet.relationshipId);
    const nextRel: Relationship = {
      id: sheet.relationshipId,
      type: WORKSHEET_REL_TYPE,
      target,
    };
    const previous = rels.get(nextRel.id);
    if (!previous || !sameRelationship(previous, nextRel)) {
      rels.set(nextRel.id, nextRel);
      changed = true;
    }
    if (!getPart(pkg, sheet.path)) {
      setXmlPart(
        pkg,
        sheet.path,
        WORKSHEET_CONTENT_TYPE,
        `${XML_DECL}<worksheet xmlns="${SS_NS}" xmlns:r="${OFFICE_REL_NS}"/>`,
      );
      setXmlPart(
        pkg,
        relationshipPartPath(sheet.path),
        RELS_CONTENT_TYPE,
        serializeRelationships(new Map()),
      );
      changed = true;
    }
    ensureContentTypeOverride(pkg, sheet.path, WORKSHEET_CONTENT_TYPE);
  }

  for (const rel of Array.from(rels.values())) {
    if (rel.type !== WORKSHEET_REL_TYPE || usedSheetRelIds.has(rel.id)) continue;
    const targetPath = resolveRelationshipTarget('xl/workbook.xml', rel.target);
    rels.delete(rel.id);
    deletePart(pkg, targetPath);
    deletePart(pkg, relationshipPartPath(targetPath));
    removeContentTypeOverride(pkg, targetPath);
    changed = true;
  }

  if (changed) writeRelationshipsPart(pkg, relsPath, rels);
}

function syncStylesPart(pkg: Package, styles: ExcelStylesModel): void {
  const normalized = normalizeStyles(styles);
  styles.numberFormats = normalized.numberFormats;
  styles.fonts = normalized.fonts;
  styles.fills = normalized.fills;
  styles.borders = normalized.borders;
  styles.cellFormats = normalized.cellFormats;
  const xml = serializeStyles(styles);
  const part = getPart(pkg, 'xl/styles.xml');
  if (part?.xmlDom) {
    replacePartXml(part, xml);
  } else {
    setXmlPart(pkg, 'xl/styles.xml', STYLES_CONTENT_TYPE, xml);
  }
  ensureContentTypeOverride(pkg, 'xl/styles.xml', STYLES_CONTENT_TYPE);
}

export function createWorkbookPackageFromModel(model: ExcelWorkbookModel): Package {
  const pkg = createPackage();
  const standaloneModel = cloneStandaloneWorkbookModel(model);
  seedWorkbookPackage(pkg, standaloneModel);
  syncWorkbookModelToPackage(pkg, standaloneModel);
  return pkg;
}

function cloneStandaloneWorkbookModel(model: ExcelWorkbookModel): ExcelWorkbookModel {
  const clone = JSON.parse(JSON.stringify(model)) as ExcelWorkbookModel;
  for (const sheet of clone.sheets) {
    sheet.threadedComments = [];
    const creatableCharts = sheet.charts.filter((chart) => chart.type);
    sheet.charts = creatableCharts;
    sheet.drawings = sheet.drawings
      .map((drawing) => ({
        ...drawing,
        anchors: drawing.anchors.filter(
          (anchor) =>
            anchor.object?.kind === 'chart' &&
            creatableCharts.some((chart) => chart.path === anchor.object?.targetPath || chart.title === anchor.object?.name),
        ),
      }))
      .filter((drawing) => drawing.anchors.length > 0);
    sheet.images = [];
  }
  return clone;
}

function seedWorkbookPackage(pkg: Package, model: ExcelWorkbookModel): void {
  setXmlPart(
    pkg,
    '[Content_Types].xml',
    CONTENT_TYPES_CONTENT_TYPE,
    serializeBaseContentTypes(model),
  );
  setXmlPart(
    pkg,
    '_rels/.rels',
    RELS_CONTENT_TYPE,
    serializeRelationships(
      new Map([
        [
          'rId1',
          {
            id: 'rId1',
            type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
            target: 'xl/workbook.xml',
          },
        ],
      ]),
    ),
  );

  const workbookRels = new Map<string, Relationship>();
  model.sheets.forEach((sheet, index) => {
    sheet.relationshipId ||= `rId${index + 1}`;
    sheet.path ||= `xl/worksheets/sheet${index + 1}.xml`;
    workbookRels.set(sheet.relationshipId, {
      id: sheet.relationshipId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      target: `worksheets/sheet${index + 1}.xml`,
    });
    setXmlPart(
      pkg,
      sheet.path,
      WORKSHEET_CONTENT_TYPE,
      `${XML_DECL}<worksheet xmlns="${SS_NS}" xmlns:r="${OFFICE_REL_NS}"/>`,
    );
    setXmlPart(
      pkg,
      relationshipPartPath(sheet.path),
      RELS_CONTENT_TYPE,
      serializeRelationships(new Map()),
    );
  });

  setXmlPart(
    pkg,
    'xl/workbook.xml',
    WORKBOOK_CONTENT_TYPE,
    `${XML_DECL}<workbook xmlns="${SS_NS}" xmlns:r="${OFFICE_REL_NS}"><sheets/></workbook>`,
  );
  setXmlPart(
    pkg,
    'xl/_rels/workbook.xml.rels',
    RELS_CONTENT_TYPE,
    serializeRelationships(workbookRels),
  );
}

function serializeBaseContentTypes(model: ExcelWorkbookModel): string {
  const overrides: Array<[string, string]> = [
    ['/xl/workbook.xml', WORKBOOK_CONTENT_TYPE],
    ...model.sheets.map((sheet, index) => [
      `/${sheet.path ?? `xl/worksheets/sheet${index + 1}.xml`}`,
      WORKSHEET_CONTENT_TYPE,
    ]),
  ];
  return `${XML_DECL}<Types xmlns="${CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="${RELS_CONTENT_TYPE}"/><Default Extension="xml" ContentType="application/xml"/>${overrides
    .map(
      ([partName, contentType]) =>
        `<Override PartName="${xmlAttr(partName)}" ContentType="${xmlAttr(contentType)}"/>`,
    )
    .join('')}</Types>`;
}

function setXmlPart(pkg: Package, path: string, contentType: string, xml: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
    }),
  );
}

function syncWorksheetTableRelationships(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const relsPath = relationshipPartPath(sheet.path);
  const rels = parseRelationships(getPart(pkg, relsPath));
  if (!sheet.tables.length && !Array.from(rels.values()).some((rel) => rel.type === TABLE_REL_TYPE))
    return;
  const usedTableRelIds = new Set<string>();
  let changed = false;

  for (const table of sheet.tables) {
    table.id ??= nextTableId(sheet);
    table.path ??= nextIndexedPartPath(pkg, 'xl/tables/table', '.xml');
    const target = relativeRelationshipTarget(sheet.path, table.path);
    const id = table.relationshipId || findReusableRelationshipId(rels, TABLE_REL_TYPE, target);
    if (id) {
      table.relationshipId = id;
    } else {
      table.relationshipId = nextRelationshipId(rels);
      changed = true;
    }
    usedTableRelIds.add(table.relationshipId);
    const nextRel: Relationship = {
      id: table.relationshipId,
      type: TABLE_REL_TYPE,
      target,
    };
    const previous = rels.get(nextRel.id);
    if (!previous || !sameRelationship(previous, nextRel)) {
      rels.set(nextRel.id, nextRel);
      changed = true;
    }
  }

  for (const rel of Array.from(rels.values())) {
    if (rel.type === TABLE_REL_TYPE && !usedTableRelIds.has(rel.id)) {
      const targetPath = resolveRelationshipTarget(sheet.path, rel.target);
      rels.delete(rel.id);
      deletePart(pkg, targetPath);
      removeContentTypeOverride(pkg, targetPath);
      changed = true;
    }
  }
  if (changed) writeRelationshipsPart(pkg, relsPath, rels);
}

function syncWorksheetCommentRelationships(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const relsPath = relationshipPartPath(sheet.path);
  const rels = parseRelationships(getPart(pkg, relsPath));
  const commentRels = Array.from(rels.values()).filter((rel) => rel.type === COMMENTS_REL_TYPE);
  if (sheet.comments.length || !commentRels.length) return;

  for (const rel of commentRels) {
    const targetPath = resolveRelationshipTarget(sheet.path, rel.target);
    rels.delete(rel.id);
    deletePart(pkg, targetPath);
    removeContentTypeOverride(pkg, targetPath);
  }
  writeRelationshipsPart(pkg, relsPath, rels);
}

function syncWorksheetThreadedCommentRelationships(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const relsPath = relationshipPartPath(sheet.path);
  const rels = parseRelationships(getPart(pkg, relsPath));
  const commentRels = Array.from(rels.values()).filter(
    (rel) => rel.type === THREADED_COMMENTS_REL_TYPE,
  );
  if (sheet.threadedComments.length || !commentRels.length) return;

  for (const rel of commentRels) {
    const targetPath = resolveRelationshipTarget(sheet.path, rel.target);
    rels.delete(rel.id);
    deletePart(pkg, targetPath);
    removeContentTypeOverride(pkg, targetPath);
  }
  writeRelationshipsPart(pkg, relsPath, rels);
}

function syncWorksheetDrawingRelationships(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const needsDrawing = sheet.drawings.some((drawing) => drawing.anchors.length > 0) || sheet.charts.some((chart) => chart.type);
  if (!needsDrawing) return;
  const drawing = sheet.drawings[0] ?? { anchors: [] };
  if (!sheet.drawings[0]) sheet.drawings.push(drawing);

  drawing.path ??= nextIndexedPartPath(pkg, 'xl/drawings/drawing', '.xml');
  const sheetRelsPath = relationshipPartPath(sheet.path);
  const sheetRels = parseRelationships(getPart(pkg, sheetRelsPath));
  const drawingTarget = relativeRelationshipTarget(sheet.path, drawing.path);
  drawing.relationshipId ||= findReusableRelationshipId(sheetRels, DRAWING_REL_TYPE, drawingTarget);
  drawing.relationshipId ||= nextRelationshipId(sheetRels);
  sheetRels.set(drawing.relationshipId, {
    id: drawing.relationshipId,
    type: DRAWING_REL_TYPE,
    target: drawingTarget,
  });
  writeRelationshipsPart(pkg, sheetRelsPath, sheetRels);
  ensureContentTypeOverride(pkg, drawing.path, DRAWING_CONTENT_TYPE);

  const drawingRelsPath = relationshipPartPath(drawing.path);
  const drawingRels = parseRelationships(getPart(pkg, drawingRelsPath));
  const chartAnchors = drawing.anchors.filter((anchor) => anchor.object?.kind === 'chart');
  const usedChartPaths = new Set(sheet.charts.map((chart) => chart.path).filter((path): path is string => Boolean(path)));
  for (let i = 0; i < sheet.charts.length; i++) {
    const chart = sheet.charts[i];
    if (!chart?.type) continue;
    chart.path ??= nextIndexedPartPathAvoiding(pkg, 'xl/charts/chart', '.xml', usedChartPaths);
    usedChartPaths.add(chart.path);
    chart.drawingPath = drawing.path;
    const chartTarget = relativeRelationshipTarget(drawing.path, chart.path);
    chart.relationshipId ||= findReusableRelationshipId(drawingRels, CHART_REL_TYPE, chartTarget);
    chart.relationshipId ||= nextRelationshipId(drawingRels);
    drawingRels.set(chart.relationshipId, {
      id: chart.relationshipId,
      type: CHART_REL_TYPE,
      target: chartTarget,
    });
    const anchor =
      chartAnchors[i] ??
      ({
        kind: 'twoCell',
        from: chart.position?.from,
        to: chart.position?.to,
        object: { kind: 'chart' },
      } satisfies ExcelDrawingAnchorModel);
    if (!chartAnchors[i]) drawing.anchors.push(anchor);
    anchor.object ??= { kind: 'chart' };
    anchor.object.kind = 'chart';
    anchor.object.name = anchor.object.name ?? chart.title;
    anchor.object.relationshipId = chart.relationshipId;
    anchor.object.targetPath = chart.path;
    if (chart.position) {
      anchor.from = { ...chart.position.from };
      anchor.to = { ...chart.position.to };
    }
  }
  writeRelationshipsPart(pkg, drawingRelsPath, drawingRels);
}

function syncWorksheetRelationships(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const relsPath = relationshipPartPath(sheet.path);
  const relsPart = getPart(pkg, relsPath);
  const rels = parseRelationships(relsPart);
  const usedHyperlinkIds = new Set<string>();
  let changed = false;

  for (const hyperlink of sheet.hyperlinks) {
    if (!hyperlink.target) continue;
    const id = hyperlink.relationshipId || findReusableHyperlinkRelId(rels, hyperlink);
    if (id) {
      hyperlink.relationshipId = id;
    } else {
      hyperlink.relationshipId = nextRelationshipId(rels);
      changed = true;
    }
    usedHyperlinkIds.add(hyperlink.relationshipId);
    const nextRel: Relationship = {
      id: hyperlink.relationshipId,
      type: HYPERLINK_REL_TYPE,
      target: hyperlink.target,
      targetMode: hyperlink.targetMode ?? 'External',
    };
    const previous = rels.get(nextRel.id);
    if (!previous || !sameRelationship(previous, nextRel)) {
      rels.set(nextRel.id, nextRel);
      changed = true;
    }
  }

  for (const rel of Array.from(rels.values())) {
    if (rel.type === HYPERLINK_REL_TYPE && !usedHyperlinkIds.has(rel.id)) {
      rels.delete(rel.id);
      changed = true;
    }
  }
  if (!changed && relsPart) return;
  if (!relsPart && !rels.size) return;
  writeRelationshipsPart(pkg, relsPath, rels);
}

function writeRelationshipsPart(
  pkg: Package,
  relsPath: string,
  rels: Map<string, Relationship>,
): void {
  const relsPart = getPart(pkg, relsPath);
  const xml = serializeRelationships(rels);
  if (relsPart) {
    replacePartXml(relsPart, xml);
  } else {
    setPart(
      pkg,
      createXmlPart({
        path: relsPath,
        contentType: RELS_CONTENT_TYPE,
        xmlDom: parseXmlDocument(xml) as unknown as Document,
      }),
    );
  }
}

function syncWorkbookPart(pkg: Package, model: ExcelWorkbookModel): void {
  const part = getPart(pkg, 'xl/workbook.xml');
  if (!part?.xmlDom) return;
  const sheetsXml = model.sheets
    .map((sheet) => {
      const attrs = [
        `name="${xmlAttr(sheet.name)}"`,
        `sheetId="${xmlAttr(sheet.id)}"`,
        sheet.visibility ? `state="${xmlAttr(sheet.visibility)}"` : '',
        sheet.relationshipId ? `r:id="${xmlAttr(sheet.relationshipId)}"` : '',
      ].filter(Boolean);
      return `<sheet ${attrs.join(' ')}/>`;
    })
    .join('');
  const definedNamesXml = model.definedNames.length
    ? `<definedNames>${model.definedNames
        .map((name) => {
          const attrs = [
            `name="${xmlAttr(name.name)}"`,
            name.localSheetId === undefined ? '' : `localSheetId="${name.localSheetId}"`,
          ].filter(Boolean);
          return `<definedName ${attrs.join(' ')}>${xmlText(name.text)}</definedName>`;
        })
        .join('')}</definedNames>`
    : '';
  const workbookPropertiesXml = model.date1904 ? '<workbookPr date1904="1"/>' : '';
  replacePartXml(
    part,
    `${XML_DECL}<workbook xmlns="${SS_NS}" xmlns:r="${OFFICE_REL_NS}">${workbookPropertiesXml}<sheets>${sheetsXml}</sheets>${definedNamesXml}</workbook>`,
  );
}

function syncWorksheetPart(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path) return;
  const part = getPart(pkg, sheet.path);
  if (!part?.xmlDom) return;
  replacePartXml(part, serializeWorksheet(sheet));
}

function serializeWorksheet(sheet: ExcelSheetModel): string {
  return `${XML_DECL}<worksheet xmlns="${SS_NS}" xmlns:r="${OFFICE_REL_NS}">${[
    sheet.dimension ? `<dimension ref="${xmlAttr(sheet.dimension)}"/>` : '',
    serializeSheetViews(sheet.sheetViews),
    serializeColumns(sheet.columns),
    serializeSheetData(sheet),
    serializeMergedRanges(sheet),
    serializeConditionalFormats(sheet.conditionalFormats),
    serializeHyperlinks(sheet.hyperlinks),
    serializeDataValidations(sheet.dataValidations),
    serializeTableParts(sheet.tables),
    serializeDrawingRef(sheet.drawings),
  ]
    .filter(Boolean)
    .join('')}</worksheet>`;
}

function serializeStyles(styles: ExcelStylesModel): string {
  const customNumberFormats = styles.numberFormats.filter((format) => format.id >= 164);
  return `${XML_DECL}<styleSheet xmlns="${SS_NS}">${[
    customNumberFormats.length
      ? `<numFmts count="${customNumberFormats.length}">${customNumberFormats
          .map((format) => `<numFmt numFmtId="${format.id}" formatCode="${xmlAttr(format.code)}"/>`)
          .join('')}</numFmts>`
      : '',
    `<fonts count="${styles.fonts.length}">${styles.fonts.map(serializeFont).join('')}</fonts>`,
    `<fills count="${styles.fills.length}">${styles.fills.map(serializeFill).join('')}</fills>`,
    `<borders count="${styles.borders.length}">${styles.borders.map(serializeBorder).join('')}</borders>`,
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`,
    `<cellXfs count="${styles.cellFormats.length}">${styles.cellFormats.map(serializeCellFormat).join('')}</cellXfs>`,
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`,
  ]
    .filter(Boolean)
    .join('')}</styleSheet>`;
}

function serializeFont(font: ExcelFontModel): string {
  return `<font>${[
    font.size === undefined ? '' : `<sz val="${font.size}"/>`,
    font.color ? serializeColor('color', font.color) : '',
    font.name === undefined ? '' : `<name val="${xmlAttr(font.name)}"/>`,
    font.family === undefined ? '' : `<family val="${font.family}"/>`,
    font.charset === undefined ? '' : `<charset val="${font.charset}"/>`,
    font.scheme === undefined ? '' : `<scheme val="${xmlAttr(font.scheme)}"/>`,
    font.bold ? '<b/>' : '',
    font.italic ? '<i/>' : '',
    font.underline ? '<u/>' : '',
  ]
    .filter(Boolean)
    .join('')}</font>`;
}

function serializeFill(fill: ExcelFillModel): string {
  return `<fill><patternFill ${stringAttr('patternType', fill.patternType)}>${[
    fill.foregroundColor ? serializeColor('fgColor', fill.foregroundColor) : '',
    fill.backgroundColor ? serializeColor('bgColor', fill.backgroundColor) : '',
  ]
    .filter(Boolean)
    .join('')}</patternFill></fill>`;
}

function serializeBorder(border: ExcelBorderModel): string {
  return `<border>${['left', 'right', 'top', 'bottom', 'diagonal']
    .map((side) =>
      serializeBorderSide(
        side,
        border[side as keyof ExcelBorderModel] as ExcelBorderSideModel | undefined,
      ),
    )
    .join('')}</border>`;
}

function serializeBorderSide(name: string, side: ExcelBorderSideModel | undefined): string {
  if (!side) return `<${name}/>`;
  return `<${name} ${stringAttr('style', side.style)}>${side.color ? serializeColor('color', side.color) : ''}</${name}>`;
}

function serializeCellFormat(format: ExcelStylesModel['cellFormats'][number]): string {
  const attrs = [
    numberAttr('numFmtId', format.numberFormatId ?? 0),
    numberAttr('fontId', format.fontId ?? 0),
    numberAttr('fillId', format.fillId ?? 0),
    numberAttr('borderId', format.borderId ?? 0),
    numberAttr('xfId', format.xfId ?? 0),
    booleanAttr('applyNumberFormat', format.applyNumberFormat),
    booleanAttr('applyFont', format.applyFont),
    booleanAttr('applyFill', format.applyFill),
    booleanAttr('applyBorder', format.applyBorder),
    booleanAttr('applyAlignment', format.applyAlignment),
  ].filter(Boolean);
  const alignment = format.alignment ? serializeAlignment(format.alignment) : '';
  return alignment ? `<xf ${attrs.join(' ')}>${alignment}</xf>` : `<xf ${attrs.join(' ')}/>`;
}

function serializeAlignment(alignment: ExcelAlignmentModel): string {
  const attrs = [
    stringAttr('horizontal', alignment.horizontal),
    stringAttr('vertical', alignment.vertical),
    numberAttr('textRotation', alignment.textRotation),
    booleanAttr('wrapText', alignment.wrapText),
    booleanAttr('shrinkToFit', alignment.shrinkToFit),
    numberAttr('indent', alignment.indent),
  ].filter(Boolean);
  return `<alignment ${attrs.join(' ')}/>`;
}

function serializeColor(name: string, color: ExcelColorModel): string {
  return `<${name} rgb="${excelColorToOoxmlArgb(color, `Excel ${name}`)}"/>`;
}

function normalizeStyles(styles: ExcelStylesModel): ExcelStylesModel {
  return {
    numberFormats: styles.numberFormats,
    fonts: styles.fonts.length ? styles.fonts : [{}],
    fills: styles.fills.length ? styles.fills : [{}, { patternType: 'gray125' }],
    borders: styles.borders.length ? styles.borders : [{}],
    cellFormats: styles.cellFormats.length ? styles.cellFormats : [{}],
  };
}

function serializeSheetViews(views: ExcelSheetViewModel[]): string {
  if (!views.length) return '';
  return `<sheetViews>${views
    .map((view) => {
      const attrs = [
        numberAttr('workbookViewId', view.workbookViewId),
        stringAttr('view', view.view),
        booleanAttr('tabSelected', view.tabSelected),
        booleanAttr('showGridLines', view.showGridLines),
        booleanAttr('showRowColHeaders', view.showRowColHeaders),
        numberAttr('zoomScale', view.zoomScale),
        numberAttr('zoomScaleNormal', view.zoomScaleNormal),
        stringAttr('topLeftCell', view.topLeftCell),
      ].filter(Boolean);
      const pane = view.pane
        ? `<pane ${[
            numberAttr('xSplit', view.pane.xSplit),
            numberAttr('ySplit', view.pane.ySplit),
            stringAttr('topLeftCell', view.pane.topLeftCell),
            stringAttr('activePane', view.pane.activePane),
            stringAttr('state', view.pane.state),
          ]
            .filter(Boolean)
            .join(' ')}/>`
        : '';
      const selections = view.selections
        .map(
          (selection) =>
            `<selection ${[
              stringAttr('pane', selection.pane),
              stringAttr('activeCell', selection.activeCell),
              numberAttr('activeCellId', selection.activeCellId),
              stringAttr('sqref', selection.sqref),
            ]
              .filter(Boolean)
              .join(' ')}/>`,
        )
        .join('');
      return `<sheetView ${attrs.join(' ')}>${pane}${selections}</sheetView>`;
    })
    .join('')}</sheetViews>`;
}

function serializeColumns(columns: ExcelColumnModel[]): string {
  if (!columns.length) return '';
  return `<cols>${columns
    .map(
      (column) =>
        `<col ${[
          numberAttr('min', column.min),
          numberAttr('max', column.max),
          numberAttr('width', column.width),
          booleanAttr('hidden', column.hidden),
          numberAttr('style', column.styleId),
          booleanAttr('customWidth', column.customWidth),
        ]
          .filter(Boolean)
          .join(' ')}/>`,
    )
    .join('')}</cols>`;
}

function serializeSheetData(sheet: ExcelSheetModel): string {
  const rowMap = new Map<number, { row?: ExcelRowModel; cells: ExcelCellModel[] }>();
  for (const row of sheet.rows) {
    rowMap.set(row.index, { row, cells: [] });
  }
  for (const cell of sheet.cells) {
    const entry = rowMap.get(cell.row) ?? { cells: [] };
    entry.cells.push(cell);
    rowMap.set(cell.row, entry);
  }
  const rows = Array.from(rowMap.entries()).sort(([a], [b]) => a - b);
  return `<sheetData>${rows
    .map(([index, entry]) => {
      const row = entry.row;
      const attrs = [
        numberAttr('r', index),
        numberAttr('ht', row?.height),
        booleanAttr('hidden', row?.hidden),
        numberAttr('s', row?.styleId),
        booleanAttr('customHeight', row?.customHeight),
      ].filter(Boolean);
      const cells = entry.cells
        .sort((a, b) => a.column - b.column)
        .map((cell) => serializeCell(cell))
        .join('');
      return cells ? `<row ${attrs.join(' ')}>${cells}</row>` : `<row ${attrs.join(' ')}/>`;
    })
    .join('')}</sheetData>`;
}

function serializeCell(cell: ExcelCellModel): string {
  const attrs = [stringAttr('r', cell.address), numberAttr('s', cell.styleId)].filter(Boolean);
  const formula = cell.formula ? `<f>${xmlText(cell.formula)}</f>` : '';
  if (cell.kind === 'string') {
    attrs.push('t="inlineStr"');
    return `<c ${attrs.join(' ')}><is><t>${xmlText(String(cell.value ?? ''))}</t></is></c>`;
  }
  if (cell.kind === 'boolean') {
    attrs.push('t="b"');
    return `<c ${attrs.join(' ')}>${formula}<v>${cell.value ? '1' : '0'}</v></c>`;
  }
  if (cell.kind === 'error') {
    attrs.push('t="e"');
    return `<c ${attrs.join(' ')}>${formula}<v>${xmlText(String(cell.value ?? cell.rawValue ?? ''))}</v></c>`;
  }
  if (cell.kind === 'blank' && !formula) {
    return `<c ${attrs.join(' ')}/>`;
  }
  const value = cell.value === null || cell.value === undefined ? '' : String(cell.value);
  return `<c ${attrs.join(' ')}>${formula}${value ? `<v>${xmlText(value)}</v>` : ''}</c>`;
}

function serializeMergedRanges(sheet: ExcelSheetModel): string {
  if (!sheet.mergedRanges.length) return '';
  return `<mergeCells count="${sheet.mergedRanges.length}">${sheet.mergedRanges
    .map((range) => `<mergeCell ref="${xmlAttr(range.address)}"/>`)
    .join('')}</mergeCells>`;
}

function serializeConditionalFormats(formats: ExcelConditionalFormatModel[]): string {
  return formats
    .map(
      (format) =>
        `<conditionalFormatting ${stringAttr('sqref', format.sqref)}>${format.rules
          .map(
            (rule) =>
              `<cfRule ${[
                stringAttr('type', rule.type),
                numberAttr('priority', rule.priority),
                stringAttr('operator', rule.operator),
                numberAttr('dxfId', rule.dxfId),
                booleanAttr('stopIfTrue', rule.stopIfTrue),
                booleanAttr('aboveAverage', rule.aboveAverage),
                booleanAttr('percent', rule.percent),
                booleanAttr('bottom', rule.bottom),
                numberAttr('rank', rule.rank),
                stringAttr('text', rule.text),
                stringAttr('timePeriod', rule.timePeriod),
              ]
                .filter(Boolean)
                .join(
                  ' ',
                )}>${rule.formulas.map((formula) => `<formula>${xmlText(formula)}</formula>`).join('')}</cfRule>`,
          )
          .join('')}</conditionalFormatting>`,
    )
    .join('');
}

function serializeHyperlinks(hyperlinks: ExcelHyperlinkModel[]): string {
  if (!hyperlinks.length) return '';
  return `<hyperlinks>${hyperlinks
    .map(
      (hyperlink) =>
        `<hyperlink ${[
          stringAttr('ref', hyperlink.ref),
          hyperlink.relationshipId ? `r:id="${xmlAttr(hyperlink.relationshipId)}"` : '',
          stringAttr('location', hyperlink.location),
          stringAttr('display', hyperlink.display),
          stringAttr('tooltip', hyperlink.tooltip),
        ]
          .filter(Boolean)
          .join(' ')}/>`,
    )
    .join('')}</hyperlinks>`;
}

function serializeDataValidations(validations: ExcelDataValidationModel[]): string {
  if (!validations.length) return '';
  return `<dataValidations count="${validations.length}">${validations
    .map((validation) => {
      const attrs = [
        stringAttr('type', validation.type),
        stringAttr('operator', validation.operator),
        booleanAttr('allowBlank', validation.allowBlank),
        booleanAttr('showDropDown', validation.showDropDown),
        booleanAttr('showInputMessage', validation.showInputMessage),
        booleanAttr('showErrorMessage', validation.showErrorMessage),
        stringAttr('errorStyle', validation.errorStyle),
        stringAttr('sqref', validation.sqref),
        stringAttr('errorTitle', validation.errorTitle),
        stringAttr('error', validation.error),
        stringAttr('promptTitle', validation.promptTitle),
        stringAttr('prompt', validation.prompt),
      ].filter(Boolean);
      return `<dataValidation ${attrs.join(' ')}>${validation.formula1 ? `<formula1>${xmlText(validation.formula1)}</formula1>` : ''}${
        validation.formula2 ? `<formula2>${xmlText(validation.formula2)}</formula2>` : ''
      }</dataValidation>`;
    })
    .join('')}</dataValidations>`;
}

function serializeTableParts(tables: ExcelTableModel[]): string {
  const tableIds = tables
    .map((table) => table.relationshipId)
    .filter((id): id is string => Boolean(id));
  if (!tableIds.length) return '';
  return `<tableParts count="${tableIds.length}">${tableIds.map((id) => `<tablePart r:id="${xmlAttr(id)}"/>`).join('')}</tableParts>`;
}

function serializeDrawingRef(drawings: ExcelDrawingModel[]): string {
  const drawing = drawings.find((item) => item.relationshipId);
  return drawing?.relationshipId ? `<drawing r:id="${xmlAttr(drawing.relationshipId)}"/>` : '';
}

function syncTableParts(pkg: Package, sheet: ExcelSheetModel): void {
  for (const table of sheet.tables) {
    if (!table.path) continue;
    const part = getPart(pkg, table.path);
    const xml = serializeTable(table);
    if (part?.xmlDom) {
      replacePartXml(part, xml);
    } else {
      setPart(
        pkg,
        createXmlPart({
          path: table.path,
          contentType: TABLE_CONTENT_TYPE,
          xmlDom: parseXmlDocument(xml) as unknown as Document,
        }),
      );
    }
    ensureContentTypeOverride(pkg, table.path, TABLE_CONTENT_TYPE);
  }
}

function serializeTable(table: ExcelTableModel): string {
  const attrs = [
    numberAttr('id', table.id),
    stringAttr('name', table.name),
    stringAttr('displayName', table.displayName),
    stringAttr('ref', table.ref),
    numberAttr('headerRowCount', table.headerRowCount),
    numberAttr('totalsRowCount', table.totalsRowCount),
    booleanAttr('totalsRowShown', table.totalsRowShown),
  ].filter(Boolean);
  const autoFilter = serializeTableAutoFilter(table);
  const columns = `<tableColumns count="${table.columns.length}">${table.columns
    .map(
      (column) =>
        `<tableColumn ${[
          numberAttr('id', column.id),
          stringAttr('name', column.name),
          stringAttr('totalsRowLabel', column.totalsRowLabel),
          stringAttr('totalsRowFunction', column.totalsRowFunction),
        ]
          .filter(Boolean)
          .join(' ')}/>`,
    )
    .join('')}</tableColumns>`;
  const style = table.style
    ? `<tableStyleInfo ${[
        stringAttr('name', table.style.name),
        booleanAttr('showFirstColumn', table.style.showFirstColumn),
        booleanAttr('showLastColumn', table.style.showLastColumn),
        booleanAttr('showRowStripes', table.style.showRowStripes),
        booleanAttr('showColumnStripes', table.style.showColumnStripes),
      ]
        .filter(Boolean)
        .join(' ')}/>`
    : '';
  return `${XML_DECL}<table xmlns="${SS_NS}" ${attrs.join(' ')}>${autoFilter}${columns}${style}</table>`;
}

function serializeTableAutoFilter(table: ExcelTableModel): string {
  const autoFilter =
    table.autoFilter ??
    (table.autoFilterRef ? { ref: table.autoFilterRef, filterColumns: [] } : undefined);
  if (!autoFilter) return '';
  const attrs = [stringAttr('ref', autoFilter.ref ?? table.autoFilterRef)].filter(Boolean);
  return `<autoFilter ${attrs.join(' ')}>${[
    ...autoFilter.filterColumns.map(serializeTableFilterColumn),
    autoFilter.sortState ? serializeTableSortState(autoFilter.sortState) : '',
  ]
    .filter(Boolean)
    .join('')}</autoFilter>`;
}

function serializeTableFilterColumn(column: ExcelTableFilterColumnModel): string {
  const attrs = [
    numberAttr('colId', column.colId),
    booleanAttr('hiddenButton', column.hiddenButton),
    booleanAttr('showButton', column.showButton),
  ].filter(Boolean);
  const filters = column.filters?.length
    ? `<filters>${column.filters.map((filter) => `<filter val="${xmlAttr(filter)}"/>`).join('')}</filters>`
    : '';
  return `<filterColumn ${attrs.join(' ')}>${filters}</filterColumn>`;
}

function serializeTableSortState(sortState: ExcelTableSortStateModel): string {
  const attrs = [
    stringAttr('ref', sortState.ref),
    booleanAttr('caseSensitive', sortState.caseSensitive),
    stringAttr('sortMethod', sortState.sortMethod),
  ].filter(Boolean);
  return `<sortState ${attrs.join(' ')}>${sortState.conditions.map(serializeTableSortCondition).join('')}</sortState>`;
}

function serializeTableSortCondition(condition: ExcelTableSortConditionModel): string {
  const attrs = [
    stringAttr('ref', condition.ref),
    booleanAttr('descending', condition.descending),
    stringAttr('sortBy', condition.sortBy),
  ].filter(Boolean);
  return `<sortCondition ${attrs.join(' ')}/>`;
}

function syncCommentParts(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path || !sheet.comments.length) return;
  const target = ensureCommentPartTarget(pkg, sheet.path);
  const xml = serializeComments(sheet.comments);
  const part = getPart(pkg, target.path);
  if (part?.xmlDom) {
    replacePartXml(part, xml);
  } else {
    setPart(
      pkg,
      createXmlPart({
        path: target.path,
        contentType: COMMENTS_CONTENT_TYPE,
        xmlDom: parseXmlDocument(xml) as unknown as Document,
      }),
    );
  }
  ensureContentTypeOverride(pkg, target.path, COMMENTS_CONTENT_TYPE);
}

function ensureCommentPartTarget(pkg: Package, sheetPath: string): { id: string; path: string } {
  const existing = relationshipTargetsByType(pkg, sheetPath, '/comments')[0];
  if (existing) return existing;
  const relsPath = relationshipPartPath(sheetPath);
  const rels = parseRelationships(getPart(pkg, relsPath));
  const id = nextRelationshipId(rels);
  const path = nextIndexedPartPath(pkg, 'xl/comments', '.xml');
  rels.set(id, {
    id,
    type: COMMENTS_REL_TYPE,
    target: relativeRelationshipTarget(sheetPath, path),
  });
  writeRelationshipsPart(pkg, relsPath, rels);
  return { id, path };
}

function serializeComments(comments: ExcelCommentModel[]): string {
  const authors = Array.from(new Set(comments.map((comment) => comment.author ?? '')));
  return `${XML_DECL}<comments xmlns="${SS_NS}"><authors>${authors
    .map((author) => `<author>${xmlText(author)}</author>`)
    .join('')}</authors><commentList>${comments
    .map((comment) => {
      const authorId = comment.authorId ?? Math.max(0, authors.indexOf(comment.author ?? ''));
      return `<comment ${[
        stringAttr('ref', comment.ref),
        numberAttr('authorId', authorId),
        numberAttr('shapeId', comment.shapeId),
      ]
        .filter(Boolean)
        .join(' ')}><text><t>${xmlText(comment.text)}</t></text></comment>`;
    })
    .join('')}</commentList></comments>`;
}

function syncThreadedCommentParts(pkg: Package, sheet: ExcelSheetModel): void {
  if (!sheet.path || !sheet.threadedComments.length) return;
  const target = relationshipTargetsByType(pkg, sheet.path, '/threadedComment')[0];
  if (!target) return;
  const part = getPart(pkg, target.path);
  if (!part?.xmlDom) return;
  replacePartXml(part, serializeThreadedComments(sheet.threadedComments));
}

function serializeThreadedComments(comments: ExcelThreadedCommentModel[]): string {
  return `${XML_DECL}<ThreadedComments xmlns="${TC_NS}">${comments
    .map(
      (comment) =>
        `<threadedComment ${[
          stringAttr('ref', comment.ref),
          stringAttr('dT', comment.dateTime),
          stringAttr('personId', comment.personId),
          stringAttr('id', comment.id),
        ]
          .filter(Boolean)
          .join(' ')}><text>${xmlText(comment.text)}</text></threadedComment>`,
    )
    .join('')}</ThreadedComments>`;
}

function syncDrawingParts(pkg: Package, sheet: ExcelSheetModel): void {
  for (const drawing of sheet.drawings) {
    if (!drawing.path) continue;
    const part = getPart(pkg, drawing.path);
    if (!part?.xmlDom) {
      setXmlPart(pkg, drawing.path, DRAWING_CONTENT_TYPE, serializeDrawing(drawing));
      ensureContentTypeOverride(pkg, drawing.path, DRAWING_CONTENT_TYPE);
      continue;
    }
    const anchorEls = [
      ...Array.from(part.xmlDom.getElementsByTagNameNS('*', 'twoCellAnchor')),
      ...Array.from(part.xmlDom.getElementsByTagNameNS('*', 'oneCellAnchor')),
      ...Array.from(part.xmlDom.getElementsByTagNameNS('*', 'absoluteAnchor')),
    ];
    for (let i = 0; i < drawing.anchors.length; i++) {
      const anchor = drawing.anchors[i];
      const anchorEl = anchorEls[i];
      if (!anchor || !anchorEl) continue;
      updateMarker(anchorEl, 'from', anchor.from);
      updateMarker(anchorEl, 'to', anchor.to);
      updateDrawingObjectName(anchorEl, anchor.object?.name);
    }
    markDirty(part);
  }
}

function serializeDrawing(drawing: ExcelDrawingModel): string {
  return `${XML_DECL}<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="${A_NS}" xmlns:c="${CHART_NS}" xmlns:r="${OFFICE_REL_NS}">${drawing.anchors
    .map((anchor, index) => serializeDrawingAnchor(anchor, index + 2))
    .join('')}</xdr:wsDr>`;
}

function serializeDrawingAnchor(anchor: ExcelDrawingAnchorModel, id: number): string {
  const from = serializeDrawingMarker('from', anchor.from);
  const to = serializeDrawingMarker('to', anchor.to);
  if (anchor.object?.kind === 'chart' && anchor.object.relationshipId) {
    const name = anchor.object.name ?? `Chart ${id}`;
    return `<xdr:twoCellAnchor editAs="oneCell">${from}${to}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${xmlAttr(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="${xmlAttr(anchor.object.relationshipId)}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  }
  return `<xdr:twoCellAnchor editAs="oneCell">${from}${to}<xdr:clientData/></xdr:twoCellAnchor>`;
}

function serializeDrawingMarker(name: 'from' | 'to', marker: ExcelDrawingMarkerModel | undefined): string {
  const value = marker ?? {};
  return `<xdr:${name}><xdr:col>${value.column ?? 0}</xdr:col><xdr:colOff>${value.columnOffset ?? 0}</xdr:colOff><xdr:row>${value.row ?? 0}</xdr:row><xdr:rowOff>${value.rowOffset ?? 0}</xdr:rowOff></xdr:${name}>`;
}

function updateMarker(
  anchorEl: Element,
  localName: 'from' | 'to',
  marker: ExcelDrawingMarkerModel | undefined,
): void {
  if (!marker) return;
  const markerEl = anchorEl.getElementsByTagNameNS('*', localName)[0];
  if (!markerEl) return;
  setFirstLocalText(markerEl, 'col', marker.column);
  setFirstLocalText(markerEl, 'colOff', marker.columnOffset);
  setFirstLocalText(markerEl, 'row', marker.row);
  setFirstLocalText(markerEl, 'rowOff', marker.rowOffset);
}

function updateDrawingObjectName(anchorEl: Element, name: string | undefined): void {
  if (name === undefined) return;
  const nameEl = anchorEl.getElementsByTagNameNS('*', 'cNvPr')[0];
  if (nameEl) nameEl.setAttribute('name', name);
}

function syncChartParts(pkg: Package, sheet: ExcelSheetModel): void {
  for (const chart of sheet.charts) {
    syncChartPart(pkg, chart);
  }
}

function syncChartPart(pkg: Package, chart: ExcelChartModel): void {
  if (!chart.path) return;
  const part = getPart(pkg, chart.path);
  if (!part?.xmlDom) {
    if (!chart.type || !chart.categoryRange || !chart.series?.length) return;
    setXmlPart(pkg, chart.path, CHART_CONTENT_TYPE, serializeChart(chart));
    ensureContentTypeOverride(pkg, chart.path, CHART_CONTENT_TYPE);
    return;
  }
  updateChartTitle(part.xmlDom, chart.title);
  const formulaEls = Array.from(part.xmlDom.getElementsByTagNameNS('*', 'f'));
  for (let i = 0; i < chart.formulaRefs.length && i < formulaEls.length; i++) {
    formulaEls[i]!.textContent = chart.formulaRefs[i] ?? '';
  }
  markDirty(part);
}

function serializeChart(chart: ExcelChartModel): string {
  const chartBody = serializeChartBody(chart);
  return `${XML_DECL}<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${A_NS}" xmlns:r="${OFFICE_REL_NS}"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${serializeChartTitle(chart.title)}<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/><c:plotArea><c:layout/>${chartBody}</c:plotArea><c:legend><c:legendPos val="r"/><c:layout/></c:legend><c:plotVisOnly val="1"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins l="0.7" r="0.7" t="0.75" b="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

function serializeChartTitle(title: string | undefined): string {
  if (!title) return '';
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlText(title)}</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>`;
}

function serializeChartBody(chart: ExcelChartModel): string {
  if (chart.type === 'pie' || chart.type === 'doughnut') return serializePieLikeChart(chart);
  if (chart.type === 'line' || chart.type === 'lineMarkers') return serializeLineChart(chart);
  return serializeBarChart(chart);
}

function serializeBarChart(chart: ExcelChartModel): string {
  const direction = chart.type?.startsWith('bar') ? 'bar' : 'col';
  const grouping = chart.type?.endsWith('Stacked') ? 'stacked' : 'clustered';
  return `<c:barChart><c:barDir val="${direction}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serializeCartesianSeries(chart)}<c:axId val="12345678"/><c:axId val="12345679"/></c:barChart>${serializeCategoryAxis()}${serializeValueAxis()}`;
}

function serializeLineChart(chart: ExcelChartModel): string {
  const marker = chart.type === 'lineMarkers' ? '1' : '0';
  return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serializeCartesianSeries(chart, marker)}<c:axId val="12345678"/><c:axId val="12345679"/></c:lineChart>${serializeCategoryAxis()}${serializeValueAxis()}`;
}

function serializePieLikeChart(chart: ExcelChartModel): string {
  const tag = chart.type === 'doughnut' ? 'doughnutChart' : 'pieChart';
  const hole = chart.type === 'doughnut' ? '<c:holeSize val="50"/>' : '';
  return `<c:${tag}><c:varyColors val="1"/>${serializePieSeries(chart)}${hole}</c:${tag}>`;
}

function serializeCartesianSeries(chart: ExcelChartModel, marker?: string): string {
  return (chart.series ?? [])
    .map((series, index) => {
      const markerXml = marker === undefined ? '' : `<c:marker><c:symbol val="${marker === '1' ? 'auto' : 'none'}"/></c:marker>`;
      return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${serializeSeriesText(series.name)}${markerXml}<c:cat><c:strRef><c:f>${xmlText(chart.categoryRange ?? '')}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${xmlText(series.values)}</c:f></c:numRef></c:val></c:ser>`;
    })
    .join('');
}

function serializePieSeries(chart: ExcelChartModel): string {
  const series = chart.series?.[0];
  if (!series) return '';
  return `<c:ser><c:idx val="0"/><c:order val="0"/>${serializeSeriesText(series.name)}<c:cat><c:strRef><c:f>${xmlText(chart.categoryRange ?? '')}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>${xmlText(series.values)}</c:f></c:numRef></c:val></c:ser>`;
}

function serializeSeriesText(name: string | undefined): string {
  if (!name) return '';
  if (name.includes('!') || name.includes('$')) {
    return `<c:tx><c:strRef><c:f>${xmlText(name)}</c:f></c:strRef></c:tx>`;
  }
  return `<c:tx><c:v>${xmlText(name)}</c:v></c:tx>`;
}

function serializeCategoryAxis(): string {
  return '<c:catAx><c:axId val="12345678"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="12345679"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>';
}

function serializeValueAxis(): string {
  return '<c:valAx><c:axId val="12345679"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="12345678"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
}

function updateChartTitle(chartDom: Document, title: string | undefined): void {
  if (title === undefined) return;
  const titleEl = chartDom.getElementsByTagNameNS('*', 'title')[0];
  if (!titleEl) return;
  const textEl = titleEl.getElementsByTagNameNS('*', 't')[0];
  if (textEl) textEl.textContent = title;
}

function replacePartXml(part: Part, xml: string): void {
  part.xmlDom = parseXmlDocument(xml) as unknown as Document;
  markDirty(part);
}

function relationshipTargetsByType(
  pkg: Package,
  sourcePath: string,
  typeSuffix: string,
): Array<{ id: string; path: string }> {
  const rels = parseRelationships(getPart(pkg, relationshipPartPath(sourcePath)));
  const targets: Array<{ id: string; path: string }> = [];
  for (const rel of rels.values()) {
    if (rel.type.endsWith(typeSuffix)) {
      targets.push({ id: rel.id, path: resolveRelationshipTarget(sourcePath, rel.target) });
    }
  }
  return targets;
}

function findReusableHyperlinkRelId(
  rels: Map<string, Relationship>,
  hyperlink: ExcelHyperlinkModel,
): string | undefined {
  for (const rel of rels.values()) {
    if (
      rel.type === HYPERLINK_REL_TYPE &&
      rel.target === hyperlink.target &&
      (rel.targetMode ?? 'External') === (hyperlink.targetMode ?? 'External')
    ) {
      return rel.id;
    }
  }
  return undefined;
}

function findReusableRelationshipId(
  rels: Map<string, Relationship>,
  type: string,
  target: string,
  targetMode?: string,
): string | undefined {
  for (const rel of rels.values()) {
    if (
      rel.type === type &&
      rel.target === target &&
      (rel.targetMode ?? '') === (targetMode ?? '')
    ) {
      return rel.id;
    }
  }
  return undefined;
}

function nextTableId(sheet: ExcelSheetModel): number {
  let max = 0;
  for (const table of sheet.tables) {
    if (table.id && table.id > max) max = table.id;
  }
  return max + 1;
}

function nextRelationshipId(rels: Map<string, Relationship>): string {
  let next = 1;
  while (rels.has(`rId${next}`)) next++;
  return `rId${next}`;
}

function sameRelationship(left: Relationship, right: Relationship): boolean {
  return (
    left.id === right.id &&
    left.type === right.type &&
    left.target === right.target &&
    (left.targetMode ?? '') === (right.targetMode ?? '')
  );
}

function serializeRelationships(rels: Map<string, Relationship>): string {
  return `${XML_DECL}<Relationships xmlns="${REL_NS}">${Array.from(rels.values())
    .map((rel) => {
      const attrs = [
        `Id="${xmlAttr(rel.id)}"`,
        `Type="${xmlAttr(rel.type)}"`,
        `Target="${xmlAttr(rel.target)}"`,
        stringAttr('TargetMode', rel.targetMode),
      ].filter(Boolean);
      return `<Relationship ${attrs.join(' ')}/>`;
    })
    .join('')}</Relationships>`;
}

function nextIndexedPartPath(pkg: Package, prefix: string, extension: string): string {
  let next = 1;
  while (getPart(pkg, `${prefix}${next}${extension}`)) next++;
  return `${prefix}${next}${extension}`;
}

function nextIndexedPartPathAvoiding(
  pkg: Package,
  prefix: string,
  extension: string,
  used: Set<string>,
): string {
  let next = 1;
  while (getPart(pkg, `${prefix}${next}${extension}`) || used.has(`${prefix}${next}${extension}`)) {
    next++;
  }
  return `${prefix}${next}${extension}`;
}

function parseRelationships(part?: Part): Map<string, Relationship> {
  const result = new Map<string, Relationship>();
  if (!part?.xmlDom) return result;
  const rels = part.xmlDom.getElementsByTagNameNS(REL_NS, 'Relationship');
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i];
    if (!rel) continue;
    const id = rel.getAttribute('Id');
    const type = rel.getAttribute('Type');
    const target = rel.getAttribute('Target');
    if (!id || !type || !target) continue;
    result.set(id, {
      id,
      type,
      target,
      targetMode: rel.getAttribute('TargetMode') ?? undefined,
    });
  }
  return result;
}

function relationshipPartPath(path: string): string {
  return `${dirname(path)}/_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
}

function dirname(path: string): string {
  const idx = path.lastIndexOf('/');
  return idx === -1 ? '' : path.slice(0, idx);
}

function normalizePath(path: string): string {
  const parts: string[] = [];
  for (const part of path.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

function resolveRelationshipTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return normalizePath(target.slice(1));
  const base = dirname(sourcePartPath);
  return normalizePath(base ? `${base}/${target}` : target);
}

function relativeRelationshipTarget(sourcePartPath: string, targetPartPath: string): string {
  const sourceParts = dirname(sourcePartPath).split('/').filter(Boolean);
  const targetParts = targetPartPath.split('/').filter(Boolean);
  let common = 0;
  while (
    common < sourceParts.length &&
    common < targetParts.length &&
    sourceParts[common] === targetParts[common]
  ) {
    common++;
  }
  const up = sourceParts.slice(common).map(() => '..');
  const down = targetParts.slice(common);
  return [...up, ...down].join('/') || targetPartPath;
}

function ensureContentTypeOverride(pkg: Package, path: string, contentType: string): void {
  const part = getPart(pkg, '[Content_Types].xml');
  if (!part?.xmlDom?.documentElement) return;
  const partName = `/${path}`;
  const overrides = part.xmlDom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override');
  for (let i = 0; i < overrides.length; i++) {
    const override = overrides[i];
    if (!override || override.getAttribute('PartName') !== partName) continue;
    if (override.getAttribute('ContentType') !== contentType) {
      override.setAttribute('ContentType', contentType);
      markDirty(part);
    }
    return;
  }
  const override = part.xmlDom.createElementNS(CONTENT_TYPES_NS, 'Override');
  override.setAttribute('PartName', partName);
  override.setAttribute('ContentType', contentType);
  part.xmlDom.documentElement.appendChild(override);
  markDirty(part);
}

function removeContentTypeOverride(pkg: Package, path: string): void {
  const part = getPart(pkg, '[Content_Types].xml');
  if (!part?.xmlDom?.documentElement) return;
  const partName = `/${path}`;
  const overrides = Array.from(part.xmlDom.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Override'));
  for (const override of overrides) {
    if (override.getAttribute('PartName') !== partName) continue;
    override.parentNode?.removeChild(override);
    markDirty(part);
  }
}

function setFirstLocalText(parent: Element, localName: string, value: number | undefined): void {
  if (value === undefined) return;
  const el = parent.getElementsByTagNameNS('*', localName)[0];
  if (el) el.textContent = String(value);
}

function stringAttr(name: string, value: string | undefined): string {
  return value === undefined ? '' : `${name}="${xmlAttr(value)}"`;
}

function numberAttr(name: string, value: number | undefined): string {
  return value === undefined ? '' : `${name}="${value}"`;
}

function booleanAttr(name: string, value: boolean | undefined): string {
  return value === undefined ? '' : `${name}="${value ? 1 : 0}"`;
}

function xmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function xmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
