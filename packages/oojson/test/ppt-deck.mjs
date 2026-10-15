import assert from "node:assert/strict";
import {
  applyPptModelToOnlyOfficePptDeck,
  onlyOfficePptDeckToPptModel,
  updateOnlyOfficePptDeckTextShape,
} from "../src/index.mjs";

function contentFixture(text = "商务风项目总结") {
  return {
    bPresentation: true,
    type: "docContent",
    content: [
      {
        bFromDocument: false,
        type: "paragraph",
        pPr: { algn: "just", bFromDocument: false },
        rPr: {
          lang: "zh-CN",
          sz: 6800,
          uniFill: {
            fill: {
              color: {
                rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
                color: {
                  rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
                  id: 15,
                  type: "scheme",
                },
                mods: [
                  { name: "lumMod", val: 85000 },
                  { name: "lumOff", val: 15000 },
                ],
                type: "uniColor",
              },
              type: "solid",
            },
            transparent: null,
            type: "fill",
          },
          latin: "Source Han Sans CN Heavy",
          ea: "Source Han Sans CN Heavy",
          bFromDocument: false,
          type: "textPr",
        },
        content: [
          {
            bFromDocument: false,
            rPr: {
              lang: "zh-CN",
              sz: 6800,
              uniFill: {
                fill: {
                  color: {
                    rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
                    color: {
                      rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
                      id: 15,
                      type: "scheme",
                    },
                    mods: [
                      { name: "lumMod", val: 85000 },
                      { name: "lumOff", val: 15000 },
                    ],
                    type: "uniColor",
                  },
                  type: "solid",
                },
                transparent: null,
                type: "fill",
              },
              latin: "Source Han Sans CN Heavy",
              ea: "Source Han Sans CN Heavy",
              bFromDocument: false,
              type: "textPr",
            },
            content: [text],
            footnotes: [],
            endnotes: [],
            reviewType: "common",
            type: "run",
          },
          {
            bFromDocument: false,
            rPr: { lang: "zh-CN", sz: 6800, bFromDocument: false, type: "textPr" },
            content: [],
            footnotes: [],
            endnotes: [],
            reviewType: "common",
            type: "endRun",
          },
        ],
        changes: [],
      },
    ],
  };
}

function deckFixture() {
  return {
    type: "presentation",
    sldSz: { cx: 12192000, cy: 6858000 },
    slides: [
      {
        type: "slide",
        cSld: {
          spTree: [
            {
              id: "15328",
              type: "shape",
              nvSpPr: { cNvPr: { id: 1, name: "幻灯片图像占位符 1" } },
              spPr: {},
              content: null,
            },
            {
              id: "539",
              type: "shape",
              extX: 6369298,
              extY: 1138773,
              nvSpPr: { cNvPr: { id: 539, name: "文本框 1" } },
              spPr: {
                custGeom: { preset: "rect" },
                xfrm: {
                  off: { x: 830855, y: 2229267 },
                  ext: { cx: 6369298, cy: 1138773 },
                },
              },
              content: {
                bodyPr: {
                  textFit: { type: "autoFit" },
                  wrap: "square",
                },
                lstStyle: [null, null, null, null, null, null, null, null, null, null],
                content: contentFixture(),
              },
            },
          ],
        },
      },
    ],
  };
}

function titleRun(deck) {
  return deck.slides[0].cSld.spTree[1].content.content.content[0].content[0];
}

{
  const model = onlyOfficePptDeckToPptModel(deckFixture());
  assert.equal(model.slideSize.width, 12192000);
  assert.equal(model.slides.length, 1);
  assert.equal(model.slides[0].title, "商务风项目总结");
  assert.equal(model.slides[0].shapes[1].id, "539");
  assert.equal(model.slides[0].shapes[1].name, "文本框 1");
  assert.equal(model.slides[0].shapes[1].kind, "textBox");
  assert.equal(model.slides[0].shapes[1].text.text, "商务风项目总结");
  assert.equal(model.slides[0].shapes[1].text.style.color, "#000000");
}

{
  const source = deckFixture();
  const model = onlyOfficePptDeckToPptModel(source);
  model.slides[0].shapes[1].text = {
    ...model.slides[0].shapes[1].text,
    style: { color: "#FF0000", fontFamily: "Arial", fontSize: 28 },
  };
  const result = applyPptModelToOnlyOfficePptDeck(source, model);
  assert.equal(result.changedCount, 1);
  assert.deepEqual(titleRun(result.deck).rPr.color, { auto: false, r: 255, g: 0, b: 0 });
  assert.deepEqual(titleRun(result.deck).rPr.uniFill.fill.color.color.rgba, {
    red: 255,
    green: 0,
    blue: 0,
    alpha: 255,
  });
  assert.equal(titleRun(result.deck).rPr.uniFill.fill.color.color.type, "srgb");
  assert.equal(titleRun(result.deck).rPr.rFonts.ascii, "Arial");
  assert.equal(titleRun(result.deck).rPr.sz, 56);
}

{
  const result = updateOnlyOfficePptDeckTextShape(
    deckFixture(),
    { slideIndex: 0, shapeName: "文本框 1" },
    { text: "新标题", style: { color: "#0000FF" } },
  );
  assert.equal(result.changedCount, 1);
  assert.equal(titleRun(result.deck).content[0], "新标题");
  assert.deepEqual(titleRun(result.deck).rPr.uniFill.fill.color.color.rgba, {
    red: 0,
    green: 0,
    blue: 255,
    alpha: 255,
  });
}

console.log("ppt-deck tests: passed");
