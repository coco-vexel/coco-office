import assert from "node:assert/strict";
import { ooxmlToJson } from "../src/ooxml-to-json.mjs";

const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <mc:AlternateContent>
          <mc:Choice Requires="wps">
            <w:drawing>
              <wp:anchor distT="0" distB="0" distL="0" distR="0"
                simplePos="0" relativeHeight="251658240" behindDoc="0"
                locked="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="column"><wp:posOffset>-328930</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="paragraph"><wp:posOffset>1233170</wp:posOffset></wp:positionV>
                <wp:extent cx="6120130" cy="0"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="10" name="Line 10"/>
                <a:graphic>
                  <a:graphicData>
                    <wps:wsp>
                      <wps:spPr>
                        <a:xfrm flipV="1">
                          <a:off x="0" y="0"/>
                          <a:ext cx="6120130" cy="0"/>
                        </a:xfrm>
                        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                        <a:noFill/>
                        <a:ln w="9525">
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                          <a:prstDash val="solid"/>
                        </a:ln>
                      </wps:spPr>
                    </wps:wsp>
                  </a:graphicData>
                </a:graphic>
              </wp:anchor>
            </w:drawing>
          </mc:Choice>
        </mc:AlternateContent>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const { doc, gaps } = ooxmlToJson(xml);
const drawing = doc.content[0].content[0].content[0];

assert.equal(drawing.type, "paraDrawing");
assert.equal(drawing.drawingType, "anchor");
assert.deepEqual(drawing.positionH, {
  relativeFrom: "column",
  posOffset: -328930,
  percent: false,
});
assert.deepEqual(drawing.positionV, {
  relativeFrom: "paragraph",
  posOffset: 1233170,
  percent: false,
});
assert.deepEqual(drawing.extent, { cx: 6120130, cy: 0 });
assert.equal(drawing.graphic.type, "connectShape");
assert.equal(drawing.graphic.extX, 6120130);
assert.equal(drawing.graphic.extY, 0);
assert.equal(drawing.graphic.nvSpPr, undefined);
assert.deepEqual(drawing.graphic.nvCxnSpPr, {
  cNvPr: {
    id: -1,
    name: "",
    descr: null,
    title: null,
    hidden: null,
  },
  nvPr: { isPhoto: null, userDrawn: null },
  cNvCxnSpPr: { cxnSpLocks: { noChangeShapeType: true } },
});
assert.equal(drawing.graphic.spPr.custGeom.preset, "line");
assert.deepEqual(drawing.graphic.spPr.custGeom.pathLst[0].commands, [
  { id: "moveTo", pt: { x: "l", y: "t" } },
  { id: "lnTo", pt: { x: "r", y: "b" } },
]);
assert.equal(drawing.graphic.spPr.xfrm.flipV, true);
assert.equal(drawing.graphic.spPr.ln.w, 9525);
assert.deepEqual(drawing.graphic.spPr.ln.fill.fill.color.rgba, {
  red: 255,
  green: 0,
  blue: 0,
  alpha: 255,
});
assert.equal(gaps.counts.has("w:r>mc:AlternateContent"), false);

const paragraphTextPrXml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:jc w:val="distribute"/>
        <w:rPr>
          <w:rFonts w:eastAsia="方正小标宋简体"/>
          <w:color w:val="FF0000"/>
          <w:spacing w:val="-20"/>
          <w:sz w:val="84"/>
          <w:szCs w:val="84"/>
        </w:rPr>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="FF0000"/>
          <w:sz w:val="44"/>
          <w:szCs w:val="44"/>
        </w:rPr>
        <w:t>{发文机关标志}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
const paragraphTextPrDoc = ooxmlToJson(paragraphTextPrXml).doc;
const headingRPr = paragraphTextPrDoc.content[0].content[0].rPr;
assert.equal(headingRPr.rFonts.eastAsia, "方正小标宋简体");
assert.equal(headingRPr.sz, 84);
assert.equal(headingRPr.szCs, 84);
assert.equal(headingRPr.spacing, -20);
assert.equal(headingRPr.b, true);

console.log("OOXML drawing conversion: 2 passed");
