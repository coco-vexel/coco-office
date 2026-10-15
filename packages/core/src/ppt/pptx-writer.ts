import { parseXmlDocument } from 'slimdom';
import { createPackage, createXmlPart, setPart, type Package } from '../pack/index.js';
import type { PptPresentationModel, PptShapeModel, PptSlideModel } from './model.js';

const CT_REL = 'application/vnd.openxmlformats-package.relationships+xml';
const CT_PRESENTATION =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

export function createPresentationPackageFromModel(model: PptPresentationModel): Package {
  const pkg = createPackage();
  const slides = model.slides.map((slide, index) => normalizeSlide(slide, index));

  setXmlPart(pkg, '_rels/.rels', CT_REL, rootRelsXml());
  setXmlPart(pkg, 'ppt/presentation.xml', CT_PRESENTATION, presentationXml(model, slides));
  setXmlPart(pkg, 'ppt/_rels/presentation.xml.rels', CT_REL, presentationRelsXml(slides));
  for (const slide of slides) {
    setXmlPart(pkg, `ppt/slides/slide${slide.index + 1}.xml`, CT_SLIDE, slideXml(slide));
  }
  return pkg;
}

function normalizeSlide(slide: PptSlideModel, index: number): PptSlideModel {
  return {
    ...slide,
    id: slide.id || String(256 + index),
    index,
    relationshipId: slide.relationshipId || `rId${index + 1}`,
    path: `ppt/slides/slide${index + 1}.xml`,
    shapes: slide.shapes.map((shape, shapeIndex) => normalizeShape(shape, shapeIndex)),
  };
}

function normalizeShape(shape: PptShapeModel, index: number): PptShapeModel {
  return {
    ...shape,
    id: shape.id || String(index + 2),
    name: shape.name || `Shape ${index + 1}`,
    geometry: shape.geometry ?? {},
  };
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

function rootRelsXml(): string {
  return xmlDeclaration(
    `<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_REL_NS}/officeDocument" Target="ppt/presentation.xml"/>` +
      `</Relationships>`,
  );
}

function presentationRelsXml(slides: PptSlideModel[]): string {
  return xmlDeclaration(
    `<Relationships xmlns="${REL_NS}">` +
      slides
        .map(
          (slide) =>
            `<Relationship Id="${xmlAttr(slide.relationshipId ?? `rId${slide.index + 1}`)}" ` +
            `Type="${OFFICE_REL_NS}/slide" Target="slides/slide${slide.index + 1}.xml"/>`,
        )
        .join('') +
      `</Relationships>`,
  );
}

function presentationXml(model: PptPresentationModel, slides: PptSlideModel[]): string {
  const size = model.slideSize ?? { width: 9144000, height: 5143500 };
  const slideIds = slides
    .map(
      (slide) =>
        `<p:sldId id="${xmlAttr(numericSlideId(slide, slide.index))}" ` +
        `r:id="${xmlAttr(slide.relationshipId ?? `rId${slide.index + 1}`)}"/>`,
    )
    .join('');
  return xmlDeclaration(
    `<p:presentation xmlns:p="${P_NS}" xmlns:r="${R_NS}">` +
      `<p:sldSz cx="${size.width}" cy="${size.height}"/>` +
      `<p:sldIdLst>${slideIds}</p:sldIdLst>` +
      `</p:presentation>`,
  );
}

function slideXml(slide: PptSlideModel): string {
  return xmlDeclaration(
    `<p:sld xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
      `<p:cSld>` +
      `<p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
      slide.shapes.map(shapeXml).join('') +
      `</p:spTree>` +
      `</p:cSld>` +
      `</p:sld>`,
  );
}

function shapeXml(shape: PptShapeModel, index: number): string {
  if (shape.table) return tableShapeXml(shape, index);
  return textShapeXml(shape, index);
}

function textShapeXml(shape: PptShapeModel, index: number): string {
  const id = xmlAttr(shape.id || String(index + 2));
  const name = xmlAttr(shape.name || `TextBox ${index + 1}`);
  const text = xmlText(shape.text?.text ?? '');
  return (
    `<p:sp>` +
    `<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrmXml(shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    `</p:sp>`
  );
}

function tableShapeXml(shape: PptShapeModel, index: number): string {
  const id = xmlAttr(shape.id || String(index + 2));
  const name = xmlAttr(shape.name || `Table ${index + 1}`);
  const rows = shape.table?.values ?? [];
  const columnCount = rows.reduce(
    (max, row) => Math.max(max, row.length),
    shape.table?.columnCount ?? 0,
  );
  const gridCols = Array.from({ length: columnCount }, () => `<a:gridCol w="1000000"/>`).join('');
  const tr = rows
    .map(
      (row) =>
        `<a:tr h="370840">` +
        Array.from({ length: columnCount }, (_, column) => tableCellXml(row[column] ?? '')).join(
          '',
        ) +
        `</a:tr>`,
    )
    .join('');
  return (
    `<p:graphicFrame>` +
    `<p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm>${xfrmInnerXml(shape)}</p:xfrm>` +
    `<a:graphic><a:graphicData uri="${TABLE_URI}">` +
    `<a:tbl><a:tblPr/><a:tblGrid>${gridCols}</a:tblGrid>${tr}</a:tbl>` +
    `</a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
  );
}

function tableCellXml(text: string): string {
  return (
    `<a:tc>` +
    `<a:txBody><a:bodyPr/><a:p><a:r><a:t>${xmlText(text)}</a:t></a:r></a:p></a:txBody>` +
    `<a:tcPr/>` +
    `</a:tc>`
  );
}

function xfrmXml(shape: PptShapeModel): string {
  return `<a:xfrm>${xfrmInnerXml(shape)}</a:xfrm>`;
}

function xfrmInnerXml(shape: PptShapeModel): string {
  const geometry = shape.geometry ?? {};
  return (
    `<a:off x="${geometry.left ?? 0}" y="${geometry.top ?? 0}"/>` +
    `<a:ext cx="${geometry.width ?? 3000000}" cy="${geometry.height ?? 600000}"/>`
  );
}

function numericSlideId(slide: PptSlideModel, index: number): string {
  return /^\d+$/.test(slide.id) ? slide.id : String(256 + index);
}

function xmlDeclaration(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
}

function xmlText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function xmlAttr(value: string): string {
  return xmlText(value).replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
