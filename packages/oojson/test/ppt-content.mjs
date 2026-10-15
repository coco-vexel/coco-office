import assert from "node:assert/strict";
import {
  applyPptTextMutationToOnlyOfficeContent,
  onlyOfficePptContentToTextBody,
  replaceOnlyOfficePptContentText,
  setOnlyOfficePptContentParagraphs,
  setOnlyOfficePptContentRunStyle,
  setOnlyOfficePptContentStyle,
  setOnlyOfficePptContentText,
} from "../src/index.mjs";

function fixture() {
  return {
    bPresentation: true,
    type: "docContent",
    content: [
      {
        bFromDocument: true,
        type: "paragraph",
        pPr: { bFromDocument: true, type: "paraPr", jc: "center" },
        rPr: { bFromDocument: true, type: "textPr" },
        changes: [],
        content: [
          {
            bFromDocument: true,
            type: "run",
            rPr: {
              bFromDocument: true,
              type: "textPr",
              sz: 48,
              color: { auto: false, r: 17, g: 34, b: 51 },
              rFonts: { ascii: "Aptos", hAnsi: "Aptos" },
            },
            footnotes: [],
            endnotes: [],
            reviewType: "common",
            content: ["商务风项目总结"],
          },
          {
            bFromDocument: true,
            type: "endRun",
            rPr: { bFromDocument: true, type: "textPr" },
            footnotes: [],
            endnotes: [],
            reviewType: "common",
            content: [],
          },
        ],
      },
    ],
  };
}

function firstRun(content) {
  return content.content[0].content.find((item) => item.type === "run");
}

function runColorRgba(content) {
  return firstRun(content).rPr.uniFill.fill.color.color.rgba;
}

{
  const model = onlyOfficePptContentToTextBody(fixture());
  assert.equal(model.text, "商务风项目总结");
  assert.equal(model.paragraphs.length, 1);
  assert.equal(model.paragraphs[0].style.alignment, "center");
  assert.deepEqual(model.paragraphs[0].runs[0].style, {
    color: "#112233",
    fontSize: 24,
    fontFamily: "Aptos",
  });
}

{
  const next = setOnlyOfficePptContentText(fixture(), "新标题");
  assert.equal(onlyOfficePptContentToTextBody(next).text, "新标题");
  assert.equal(next.content[0].changes.length, 0);
  assert.equal(next.content[0].content.at(-1).type, "endRun");
}

{
  const next = setOnlyOfficePptContentParagraphs(fixture(), [
    { runs: [{ text: "第一段", style: { color: "#0000ff", bold: true } }] },
    { text: "第二段", style: { italic: true } },
  ]);
  const model = onlyOfficePptContentToTextBody(next);
  assert.equal(model.text, "第一段\n第二段");
  assert.equal(firstRun(next).rPr.color.b, 255);
  assert.deepEqual(runColorRgba(next), { red: 0, green: 0, blue: 255, alpha: 255 });
  assert.equal(firstRun(next).rPr.b, true);
  assert.equal(next.content[1].content[0].rPr.i, true);
}

{
  const next = setOnlyOfficePptContentStyle(fixture(), {
    color: "#FF0000",
    fontFamily: "Arial",
    fontSize: 28,
    italic: true,
  });
  const run = firstRun(next);
  assert.deepEqual(run.rPr.color, { auto: false, r: 255, g: 0, b: 0 });
  assert.deepEqual(run.rPr.uniFill.fill.color.color.rgba, { red: 255, green: 0, blue: 0, alpha: 255 });
  assert.equal(run.rPr.uniFill.fill.color.color.type, "srgb");
  assert.equal(run.rPr.uniFill.fill.color.type, "uniColor");
  assert.equal(run.rPr.uniFill.fill.type, "solid");
  assert.equal(run.rPr.rFonts.ascii, "Arial");
  assert.equal(run.rPr.sz, 56);
  assert.equal(run.rPr.i, true);
  assert.deepEqual(next.content[0].rPr.color, { auto: false, r: 255, g: 0, b: 0 });
  assert.deepEqual(next.content[0].rPr.uniFill.fill.color.color.rgba, { red: 255, green: 0, blue: 0, alpha: 255 });
  assert.deepEqual(next.content[0].content.at(-1).rPr.color, { auto: false, r: 255, g: 0, b: 0 });
  assert.deepEqual(next.content[0].content.at(-1).rPr.uniFill.fill.color.color.rgba, {
    red: 255,
    green: 0,
    blue: 0,
    alpha: 255,
  });
}

{
  const next = setOnlyOfficePptContentRunStyle(fixture(), 0, { color: "#0000FF" });
  assert.equal(onlyOfficePptContentToTextBody(next).paragraphs[0].runs[0].style.color, "#0000FF");
  assert.deepEqual(next.content[0].rPr.color, { auto: false, r: 0, g: 0, b: 255 });
  assert.deepEqual(runColorRgba(next), { red: 0, green: 0, blue: 255, alpha: 255 });
  assert.throws(() => setOnlyOfficePptContentRunStyle(fixture(), 1, { color: "#0000FF" }), /runIndex out of range/);
}

{
  const source = fixture();
  const run = firstRun(source);
  delete run.rPr.color;
  run.rPr.uniFill = {
    fill: {
      color: {
        rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
        color: { rgba: { red: 255, green: 0, blue: 0, alpha: 255 }, type: "srgb" },
        type: "uniColor",
      },
      type: "solid",
    },
    transparent: null,
    type: "fill",
  };
  assert.equal(onlyOfficePptContentToTextBody(source).paragraphs[0].runs[0].style.color, "#FF0000");
}

{
  const source = fixture();
  const red = setOnlyOfficePptContentStyle(source, { color: "#FF0000" });
  const black = setOnlyOfficePptContentStyle(red, { color: "#000000" });
  const run = firstRun(black);
  assert.deepEqual(run.rPr.color, { auto: false, r: 0, g: 0, b: 0 });
  assert.deepEqual(run.rPr.uniFill.fill.color.rgba, { red: 0, green: 0, blue: 0, alpha: 255 });
  assert.deepEqual(run.rPr.uniFill.fill.color.color.rgba, { red: 0, green: 0, blue: 0, alpha: 255 });
  assert.equal(run.rPr.uniFill.fill.color.color.type, "srgb");
  assert.equal(run.rPr.uniFill.fill.color.mods, undefined);
  assert.equal(onlyOfficePptContentToTextBody(black).paragraphs[0].runs[0].style.color, "#000000");
}

{
  const result = replaceOnlyOfficePptContentText(fixture(), "项目", "Project", 1);
  assert.equal(result.count, 1);
  assert.equal(onlyOfficePptContentToTextBody(result.content).text, "商务风Project总结");
}

{
  const result = applyPptTextMutationToOnlyOfficeContent(fixture(), {
    type: "shape.text.run.style.set",
    slideId: "1",
    shapeId: "2",
    runIndex: 0,
    style: { color: "#0000FF" },
  });
  assert.equal(result.changedCount, 1);
  assert.equal(onlyOfficePptContentToTextBody(result.content).paragraphs[0].runs[0].style.color, "#0000FF");
}

console.log("ppt-content tests: passed");
