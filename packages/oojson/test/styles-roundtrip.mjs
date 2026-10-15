// C1 单测：ONLYOFFICE-JSON(段落带数字 pStyle) + 样式定义集
//          → jsonToFlatOpcWithStyles → SDK fromFlatOpc → buildStylesIndex
//          → 断言 core 能按数字 styleId 读出 name/outlineLevel(判标题层级)。
//
// 攻头号风险(§6.1)：oojson 原本只吐 document.xml，缺 styles.xml → core 判级落空。
// 本测证明：补出 styles.xml 整包后，core 能正确判级；且对照"无 styles 薄壳"会落空。
//
// 跑：node test/styles-roundtrip.mjs
// 依赖当前 workspace 的 SDK dist；先执行 pnpm build，缺失时测试必须失败。

import { fromFlatOpc, toFlatOpc } from "@coco-vexel/coco-office/pack";
import { buildStylesIndex, getStyleById } from "@coco-vexel/coco-office/word";
import { jsonToFlatOpcWithStyles, buildStylesXml, buildNumberingXml, ooOutlineToOoxml, withPackageLongTailParts } from "../src/index.mjs";


let pass = 0, fail = 0;
const ok = (cond, msg) => { if (cond) { pass++; console.log("  ✓", msg); } else { fail++; console.log("  ✗ FAIL:", msg); } };

// ── 夹具：模拟一篇论文。段落引用 ONLYOFFICE 数字 styleId；styleDefs 是 GetStyle().ToJSON() 形态。──
// 数字 id 对应实测：701=Heading 1, 666=Heading 2, 665=Normal(正文)。
const ooJsonDoc = {
  type: "document",
  paraPr: { bFromDocument: true, type: "paraPr" },
  content: [
    { type: "paragraph", bFromDocument: true, pPr: { bFromDocument: true, type: "paraPr", pStyle: "701" },
      rPr: { bFromDocument: true, type: "textPr" }, changes: [],
      content: [{ type: "run", bFromDocument: true, rPr: {}, content: ["第一章 绪论"], footnotes: [], endnotes: [], reviewType: "common" },
                { type: "endRun", bFromDocument: true, rPr: {}, content: [], footnotes: [], endnotes: [], reviewType: "common" }] },
    { type: "paragraph", bFromDocument: true, pPr: { bFromDocument: true, type: "paraPr", pStyle: "666" },
      rPr: { bFromDocument: true, type: "textPr" }, changes: [],
      content: [{ type: "run", bFromDocument: true, rPr: {}, content: ["1.1 研究背景"], footnotes: [], endnotes: [], reviewType: "common" },
                { type: "endRun", bFromDocument: true, rPr: {}, content: [], footnotes: [], endnotes: [], reviewType: "common" }] },
    { type: "paragraph", bFromDocument: true, pPr: { bFromDocument: true, type: "paraPr", pStyle: "665" },
      rPr: { bFromDocument: true, type: "textPr" }, changes: [],
      content: [{ type: "run", bFromDocument: true, rPr: {}, content: ["这是正文。"], footnotes: [], endnotes: [], reviewType: "common" },
                { type: "endRun", bFromDocument: true, rPr: {}, content: [], footnotes: [], endnotes: [], reviewType: "common" }] },
  ],
};

// GetStyle(name).ToJSON() 实测形态：outlineLvl 是 ONLYOFFICE 1-based（Heading 1 → 1）。
const styleDefs = [
  { styleId: "701", def: { name: "heading 1", basedOn: "665", next: "665", qFormat: true,
      pPr: {
        outlineLvl: 1,
        spacing: { before: 240, after: 120, beforeAutoSpacing: "off", afterAutoSpacing: "on" },
        ind: { left: 420, hanging: 120 },
        pBdr: { bottom: { val: "single", sz: 8, color: { auto: false, r: 31, g: 78, b: 121 } } },
        type: "paraPr",
      },
      rPr: {
        b: true,
        sz: 24,
        color: { auto: false, r: 192, g: 0, b: 0 },
        rFonts: { ascii: "Aptos", hAnsiTheme: "minorHAnsi", eastAsia: "宋体" },
        u: { val: "single", color: { auto: false, r: 192, g: 0, b: 0 } },
        highlight: { auto: false, r: 255, g: 255, b: 0 },
        vertAlign: "superscript",
        spacing: 20,
        lang: { val: "en-US", eastAsia: "zh-CN" },
        type: "textPr",
      }, type: "paragraph" } },
  { styleId: "666", def: { name: "heading 2", basedOn: "665", next: "665", qFormat: true,
      pPr: { outlineLvl: 2, spacing: { before: 200, after: 100 }, type: "paraPr" },
      rPr: { b: true, sz: 21, type: "textPr" }, type: "paragraph" } },
  { styleId: "665", def: { name: "Normal", basedOn: null, qFormat: true,
      pPr: { jc: "both", type: "paraPr" }, rPr: { sz: 21, type: "textPr" }, type: "paragraph" } },
  { styleId: "710", def: { name: "Table Grid", styleType: "tableStyle", type: "style",
      tblPr: {
        tblW: { type: "pct", w: 5000 },
        tblBorders: { top: { val: "single", sz: 4, color: { auto: false, r: 0, g: 0, b: 0 } }, insideH: { val: "single", sz: 4 } },
        tblCellMar: { left: { type: "dxa", w: 108 }, right: { type: "dxa", w: 108 } },
        tblLayout: "fixed",
      },
      tblStylePr: {
        firstRow: { rPr: { b: true, color: { auto: false, r: 255, g: 255, b: 255 } }, tcPr: { shd: { fill: { auto: false, r: 31, g: 78, b: 121 } } } },
        bottomRightCell: { tcPr: { tcBorders: { bottom: { val: "double", sz: 8 } } } },
      } } },
];

console.log("=== 1) outlineLvl 偏移转换（ONLYOFFICE 1-based → OOXML 0-based）===");
ok(ooOutlineToOoxml(1) === 0, "Heading 1: outlineLvl 1 → 0");
ok(ooOutlineToOoxml(2) === 1, "Heading 2: outlineLvl 2 → 1");
ok(ooOutlineToOoxml(9) === 8, "Heading 9: outlineLvl 9 → 8");
ok(ooOutlineToOoxml(10) === 9, "10 钳到正文级 9");
ok(ooOutlineToOoxml(null) === null, "null → null（无大纲级别）");

console.log("\n=== 2) styles.xml 重建形状 ===");
const stylesXml = buildStylesXml(styleDefs);
ok(stylesXml.includes('w:styleId="701"'), "含数字 styleId 701");
ok(stylesXml.includes('<w:name w:val="heading 1"/>'), "含 name=heading 1（core 判语义角色强信号）");
ok(stylesXml.includes('<w:outlineLvl w:val="0"/>'), "Heading 1 → outlineLvl=0（已 -1）");
ok(stylesXml.includes('<w:outlineLvl w:val="1"/>'), "Heading 2 → outlineLvl=1（已 -1）");
ok(stylesXml.includes('<w:basedOn w:val="665"/>'), "basedOn 链保留（指向数字 id 665）");
ok(stylesXml.includes('<w:b/>'), "样式 rPr 保留粗体");
ok(stylesXml.includes('<w:sz w:val="24"/>'), "样式 rPr 保留字号");
ok(stylesXml.includes('<w:color w:val="C00000"/>'), "样式 rPr 保留颜色");
ok(stylesXml.includes('w:beforeAutospacing="0"'), "样式 pPr 保留 beforeAutoSpacing");
ok(stylesXml.includes('w:afterAutospacing="1"'), "样式 pPr 保留 afterAutoSpacing");
ok(stylesXml.includes('<w:pBdr>') && stylesXml.includes('<w:bottom w:val="single"'), "样式 pPr 保留段落边框");
ok(stylesXml.includes('w:hAnsiTheme="minorHAnsi"'), "样式 rPr 保留主题字体");
ok(stylesXml.includes('<w:u w:val="single"'), "样式 rPr 保留下划线");
ok(stylesXml.includes('<w:highlight w:val="yellow"/>'), "样式 rPr 保留高亮");
ok(stylesXml.includes('<w:vertAlign w:val="superscript"/>'), "样式 rPr 保留上下标");
ok(stylesXml.includes('<w:spacing w:val="20"/>'), "样式 rPr 保留字符间距");
ok(stylesXml.includes('<w:lang w:val="en-US" w:eastAsia="zh-CN"'), "样式 rPr 保留语言");
ok(stylesXml.includes('w:type="table"') && stylesXml.includes('<w:tblPr>'), "表格样式保留 tblPr");
ok(stylesXml.includes('<w:tblBorders>') && stylesXml.includes('<w:insideH'), "表格样式保留边框");
ok(stylesXml.includes('<w:tblStylePr w:type="firstRow"'), "表格条件样式保留 firstRow");
ok(stylesXml.includes('<w:tblStylePr w:type="seCell"'), "OnlyOffice bottomRightCell 映射为 OOXML seCell");
const docJsonStyleDefs = [
  { styleId: "701", def: { styleId: "701", name: "heading 1", basedOn: "665", qFormat: true, styleType: "paragraphStyle", type: "style",
      pPr: { outlineLvl: 1, type: "paraPr" },
      rPr: { b: true, sz: 24, color: { auto: false, r: 192, g: 0, b: 0 }, type: "textPr" } } },
];
const docJsonStylesXml = buildStylesXml(docJsonStyleDefs);
ok(docJsonStylesXml.includes('w:type="paragraph"'), "ApiDocument.ToJSON styles map styleType=paragraphStyle is mapped to OOXML paragraph");
ok(docJsonStylesXml.includes('<w:color w:val="C00000"/>'), "ApiDocument.ToJSON styles map keeps rPr color");
console.log("\n=== 3) Flat OPC 整包结构 ===");
const flatWith = jsonToFlatOpcWithStyles(ooJsonDoc, styleDefs);
ok(flatWith.includes("<?mso-application progid=\"Word.Document\"?>"), "Flat OPC 头");
ok(flatWith.includes('pkg:name="/word/styles.xml"'), "含 styles.xml part");
ok(flatWith.includes('pkg:name="/word/document.xml"'), "含 document.xml part");
ok(flatWith.includes('pkg:name="/[Content_Types].xml"'), "含 Content_Types");
ok(flatWith.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'), "document→styles 关系");
ok(flatWith.includes('pkg:name="/word/theme/theme1.xml"'), "含 theme.xml part");
ok(flatWith.includes('pkg:name="/word/fontTable.xml"'), "含 fontTable.xml part");
ok(flatWith.includes('pkg:name="/word/settings.xml"'), "含 settings.xml part");
ok(flatWith.includes('pkg:name="/word/webSettings.xml"'), "含 webSettings.xml part");
ok(flatWith.includes('application/vnd.openxmlformats-officedocument.theme+xml'), "Content_Types contains theme override");
ok(flatWith.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'), "Content_Types contains fontTable override");
ok(flatWith.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'), "Content_Types contains settings override");
ok(flatWith.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'), "Content_Types contains webSettings override");
ok(flatWith.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"'), "document→theme 关系");
ok(flatWith.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"'), "document→fontTable 关系");
ok(flatWith.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'), "document→settings 关系");
ok(flatWith.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"'), "document→webSettings 关系");
const customThemeDoc = {
  ...ooJsonDoc,
  theme: {
    name: "Coco Theme",
    themeElements: {
      clrScheme: {
        name: "Coco Colors",
        accent1: { color: { type: "srgb", rgba: { red: 18, green: 52, blue: 86, alpha: 255 } } },
        hlink: { color: { type: "srgb", rgba: { red: 1, green: 2, blue: 3, alpha: 255 } } },
      },
      fontScheme: {
        name: "Coco Fonts",
        majorFont: { latin: "Aptos Display", ea: "SimSun", cs: "Arial" },
        minorFont: { latin: "Aptos", ea: "Microsoft YaHei", cs: "Arial" },
      },
      fmtScheme: {
        name: "Coco Formats",
        fillStyleLst: [{ solidFill: { color: { rgba: { red: 10, green: 20, blue: 30 } } } }],
        lnStyleLst: [{ w: 12700, cap: "rnd", cmpd: "sng", algn: "ctr", solidFill: { schemeClr: "accent2" } }],
        effectStyleLst: [{ effectLst: { outerShdw: { blurRad: 40000, dist: 20000, dir: 5400000, algn: "ctr", rotWithShape: 0, color: { rgba: { red: 4, green: 5, blue: 6 } } } } }],
        bgFillStyleLst: [{ solidFill: { schemeClr: "accent1" } }],
      },
    },
  },
};
const flatWithTheme = jsonToFlatOpcWithStyles(customThemeDoc, styleDefs);
ok(flatWithTheme.includes('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Coco Theme">'), "theme.xml uses ONLYOFFICE theme name");
ok(flatWithTheme.includes('<a:clrScheme name="Coco Colors">'), "theme.xml uses ONLYOFFICE color scheme name");
ok(flatWithTheme.includes('<a:accent1><a:srgbClr val="123456"/></a:accent1>'), "theme.xml uses ONLYOFFICE accent color");
ok(flatWithTheme.includes('<a:hlink><a:srgbClr val="010203"/></a:hlink>'), "theme.xml uses ONLYOFFICE hyperlink color");
ok(flatWithTheme.includes('<a:fontScheme name="Coco Fonts">'), "theme.xml uses ONLYOFFICE font scheme name");
ok(flatWithTheme.includes('<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="SimSun"/><a:cs typeface="Arial"/></a:majorFont>'), "theme.xml uses ONLYOFFICE major fonts");
ok(flatWithTheme.includes('<w:font w:name="Microsoft YaHei">'), "fontTable includes ONLYOFFICE theme fonts");
ok(flatWithTheme.includes('<a:fmtScheme name="Coco Formats">'), "theme.xml uses ONLYOFFICE format scheme name");
ok(flatWithTheme.includes('<a:fillStyleLst><a:solidFill><a:srgbClr val="0A141E"/></a:solidFill></a:fillStyleLst>'), "theme.xml uses ONLYOFFICE fill style color");
ok(flatWithTheme.includes('<a:bgFillStyleLst><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:bgFillStyleLst>'), "theme.xml uses ONLYOFFICE background fill style");
ok(flatWithTheme.includes('<a:ln w="12700" cap="rnd" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln>'), "theme.xml uses ONLYOFFICE line style");
ok(flatWithTheme.includes('<a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="040506"/></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'), "theme.xml uses ONLYOFFICE effect style shadow");

// ── 4) 端到端：喂给 SDK core，验证 buildStylesIndex 判级 ──
const numbering = {
  type: "numbering",
  abstractNum: {
    "10": {
      abstractNumId: "10",
      multiLevelType: "multilevel",
      styleLink: "ListNumber",
      lvl: [
        {
          start: 1,
          numFmt: { val: "decimal" },
          pStyle: "701",
          suff: "tab",
          lvlText: "%1.",
          lvlJc: "left",
          pPr: { ind: { left: 720, hanging: 360 }, tabs: [{ val: "num", pos: 720 }] },
          rPr: { b: true, color: { auto: false, r: 192, g: 0, b: 0 } },
        },
        {
          start: 1,
          numFmt: "lowerLetter",
          suff: "space",
          lvlText: "%2)",
          lvlJc: "left",
          restart: 1,
          legal: true,
          picBulletId: "8",
          pPr: { ind: { left: 1440, hanging: 360 } },
        },
      ],
    },
  },
  num: {
    "20": { numId: "20", abstractNumId: "10", lvlOverride: [{ ilvl: 0, startOverride: 3 }] },
  },
};
const numberingXml = buildNumberingXml(numbering);
ok(numberingXml.includes('<w:numbering'), "numbering.xml root is emitted");
ok(numberingXml.includes('<w:abstractNum w:abstractNumId="10"'), "abstractNum id is emitted");
ok(numberingXml.includes('<w:num w:numId="20"'), "num id is emitted");
ok(numberingXml.includes('<w:abstractNumId w:val="10"/>'), "num references abstractNum");
ok(numberingXml.includes('<w:lvl w:ilvl="0"'), "level 0 is emitted");
ok(numberingXml.includes('<w:numFmt w:val="decimal"/>'), "level numFmt object is emitted");
ok(numberingXml.includes('<w:numFmt w:val="lowerLetter"/>'), "level numFmt string is emitted");
ok(numberingXml.includes('<w:pStyle w:val="701"/>'), "level pStyle is emitted");
ok(numberingXml.includes('<w:lvlText w:val="%1."/>'), "level text is emitted");
ok(numberingXml.includes('<w:lvlRestart w:val="1"/>'), "level restart alias is emitted");
ok(numberingXml.includes('<w:isLgl/>'), "legal numbering alias is emitted");
ok(numberingXml.includes('<w:lvlPicBulletId w:val="8"/>'), "picture bullet id reference is emitted");
ok(numberingXml.includes('<w:ind w:left="720" w:hanging="360"'), "level pPr is emitted");
ok(numberingXml.includes('<w:color w:val="C00000"/>'), "level rPr is emitted");
ok(numberingXml.includes('<w:lvlOverride w:ilvl="0"'), "level override is emitted");
ok(numberingXml.includes('<w:startOverride w:val="3"/>'), "startOverride is emitted");
const flatNumbered = jsonToFlatOpcWithStyles(ooJsonDoc, styleDefs, undefined, numberingXml);
ok(flatNumbered.includes('pkg:name="/word/numbering.xml"'), "Flat OPC contains numbering.xml part");
ok(flatNumbered.includes('ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"'), "Content_Types contains numbering override");
ok(flatNumbered.includes('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'), "document relationship points to numbering.xml");
const flatWithExtra = jsonToFlatOpcWithStyles(ooJsonDoc, styleDefs, undefined, undefined, {
  extraParts: [
    {
      name: "/word/header1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
      xml: "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>",
      rel: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        target: "header1.xml",
      },
    },
    {
      name: "/word/footer1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
      xml: "<w:ftr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>",
      rel: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        target: "footer1.xml",
      },
    },
    {
      name: "/word/comments.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
      xml: "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:comment w:id=\"0\" w:author=\"coco\"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:comments>",
      rel: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        target: "comments.xml",
      },
    },
    {
      name: "/word/footnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
      xml: "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:footnote w:id=\"2\"><w:p><w:r><w:t>footnote</w:t></w:r></w:p></w:footnote></w:footnotes>",
      rel: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
        target: "footnotes.xml",
      },
    },
    {
      name: "/word/endnotes.xml",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
      xml: "<w:endnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:endnote w:id=\"2\"><w:p><w:r><w:t>endnote</w:t></w:r></w:p></w:endnote></w:endnotes>",
      rel: {
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
        target: "endnotes.xml",
      },
    },
    {
      name: "/word/media/image1.png",
      contentType: "image/png",
      binary: "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lk9N1wAAAABJRU5ErkJggg==",
      rel: {
        source: "/word/header1.xml",
        id: "rIdImage1",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      },
    },
  ],
});
ok(flatWithExtra.includes('pkg:name="/word/header1.xml"'), "extraParts can carry header part");
ok(flatWithExtra.includes('pkg:name="/word/footer1.xml"'), "extraParts can carry footer part");
ok(flatWithExtra.includes('pkg:name="/word/comments.xml"'), "extraParts can carry comments part");
ok(flatWithExtra.includes('pkg:name="/word/footnotes.xml"'), "extraParts can carry footnotes part");
ok(flatWithExtra.includes('pkg:name="/word/endnotes.xml"'), "extraParts can carry endnotes part");
ok(flatWithExtra.includes('pkg:name="/word/media/image1.png"'), "extraParts can carry binary media part");
ok(flatWithExtra.includes('pkg:name="/word/_rels/header1.xml.rels"'), "extraParts can emit source-part relationship files");
ok(flatWithExtra.includes('relationships/header'), "extraParts can add document header relationship");
ok(flatWithExtra.includes('relationships/footer'), "extraParts can add document footer relationship");
ok(flatWithExtra.includes('relationships/comments'), "extraParts can add document comments relationship");
ok(flatWithExtra.includes('relationships/footnotes'), "extraParts can add document footnotes relationship");
ok(flatWithExtra.includes('relationships/endnotes'), "extraParts can add document endnotes relationship");
ok(flatWithExtra.includes('relationships/image'), "extraParts can add header image relationship");
ok(flatWithExtra.includes('Target="media/image1.png"'), "header relationship target is relative to header part");
ok(flatWithExtra.includes("<pkg:binaryData>iVBORw0KGgo"), "binary media is serialized as Flat OPC binaryData");
const sectionLongTailDoc = {
  ...ooJsonDoc,
  sectPr: {
    footnotePr: { pos: "pageBottom", numFmt: "decimal", numStart: 2, numRestart: "eachSect" },
    endnotePr: { position: "sectEnd", numFmt: { val: "lowerRoman" }, start: 3, restart: "continuous" },
    headerReference: {
      default: {
        content: {
          type: "docContent",
          content: [
            { type: "paragraph", content: [{ type: "run", content: ["real header"] }] },
            {
              type: "paragraph",
              content: [{
                type: "run",
                content: [{
                  type: "paraDrawing",
                  docPr: { id: 11, name: "Header Picture" },
                  extent: { cx: 9525, cy: 9525 },
                  graphic: { type: "image", blipFill: { rasterImageId: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lk9N1wAAAABJRU5ErkJggg==" } },
                }],
              }],
            },
          ],
        },
      },
      first: null,
      even: null,
    },
    footerReference: {
      even: {
        content: {
          type: "docContent",
          content: [
            { type: "paragraph", content: [{ type: "run", content: ["real footer"] }] },
          ],
        },
      },
      first: null,
      default: null,
    },
  },
};
const sectionLongTail = withPackageLongTailParts(sectionLongTailDoc);
const flatFromSectionLongTail = jsonToFlatOpcWithStyles(sectionLongTail.doc, styleDefs, undefined, undefined, {
  extraParts: sectionLongTail.extraParts,
});
ok(sectionLongTail.summary.header === 1, "sectPr headerReference content produces one header part");
ok(sectionLongTail.summary.footer === 1, "sectPr footerReference content produces one footer part");
ok(sectionLongTail.summary.media === 1, "sectPr header image produces one media part");
ok(flatFromSectionLongTail.includes('pkg:name="/word/header1.xml"'), "real sectPr header is emitted as header part");
ok(flatFromSectionLongTail.includes('pkg:name="/word/footer1.xml"'), "real sectPr footer is emitted as footer part");
ok(flatFromSectionLongTail.includes('<w:t xml:space="preserve">real header</w:t>'), "header part preserves header text");
ok(flatFromSectionLongTail.includes('pkg:name="/word/media/image1.png"'), "header image emits media part");
ok(flatFromSectionLongTail.includes('pkg:name="/word/_rels/header1.xml.rels"'), "header image emits header relationship part");
ok(flatFromSectionLongTail.includes('Target="media/image1.png"'), "header image relationship targets media part");
ok(flatFromSectionLongTail.includes('<wp:docPr id="11" name="Header Picture"/>'), "header image drawing preserves docPr");
ok(flatFromSectionLongTail.includes('r:embed="rIdCocoImage1"'), "header image drawing references generated image rId");
ok(flatFromSectionLongTail.includes('<w:t xml:space="preserve">real footer</w:t>'), "footer part preserves footer text");
ok(flatFromSectionLongTail.includes('<w:headerReference w:type="default" r:id="rIdCocoHeader1"/>'), "document sectPr references generated header rId");
ok(flatFromSectionLongTail.includes('<w:footerReference w:type="even" r:id="rIdCocoFooter1"/>'), "document sectPr references generated footer rId");
ok(flatFromSectionLongTail.includes('<w:footnotePr><w:pos w:val="pageBottom"/><w:numFmt w:val="decimal"/><w:numStart w:val="2"/><w:numRestart w:val="eachSect"/></w:footnotePr>'), "sectPr footnotePr emits numbering settings");
ok(flatFromSectionLongTail.includes('<w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:numRestart w:val="continuous"/></w:endnotePr>'), "sectPr endnotePr emits numbering settings");
const commentLongTailDoc = {
  type: "document",
  commentMetadata: [
    { author: "COCO", text: "review note", time: Date.UTC(2026, 6, 2, 12, 34, 56), solved: true },
  ],
  content: [
    {
      type: "paragraph",
      content: [
        { type: "commentRangeStart", id: "42", author: "COCO", text: "review note" },
        { type: "run", content: ["commented text"] },
        { type: "commentRangeEnd", id: "42", author: "COCO", text: "review note" },
      ],
    },
  ],
};
const commentLongTail = withPackageLongTailParts(commentLongTailDoc);
const flatWithComments = jsonToFlatOpcWithStyles(commentLongTail.doc, [], undefined, undefined, {
  extraParts: commentLongTail.extraParts,
});
ok(commentLongTail.summary.comments === 1, "comment range anchors produce one comments.xml entry");
ok(flatWithComments.includes('pkg:name="/word/comments.xml"'), "real comment anchors emit comments.xml part");
ok(flatWithComments.includes('pkg:name="/word/commentsExtended.xml"'), "solved comment emits commentsExtended.xml part");
ok(flatWithComments.includes('relationships/comments'), "real comment anchors add document comments relationship");
ok(flatWithComments.includes('relationships/commentsExtended'), "solved comment adds commentsExtended relationship");
ok(flatWithComments.includes('<w:comment w:id="42" w:author="COCO" w:date="2026-07-02T12:34:56.000Z">'), "comments.xml preserves comment id author and date metadata");
ok(flatWithComments.includes('<w:p w14:paraId="0000002A"><w:r><w:t xml:space="preserve">review note</w:t>'), "comments.xml emits stable comment paraId");
ok(flatWithComments.includes('<w15:commentEx w15:paraId="0000002A" w15:done="1"/>'), "commentsExtended.xml preserves solved state");
ok(flatWithComments.includes('<w:t xml:space="preserve">review note</w:t>'), "comments.xml preserves comment text");
ok(flatWithComments.includes('<w:commentRangeStart w:id="42"/>'), "document emits commentRangeStart");
ok(flatWithComments.includes('<w:commentRangeEnd w:id="42"/><w:r><w:commentReference w:id="42"/></w:r>'), "document emits commentRangeEnd and commentReference");
const imageLongTailDoc = {
  type: "document",
  content: [
    {
      type: "paragraph",
      content: [
        {
          type: "run",
          content: [
            {
              type: "paraDrawing",
              docPr: { id: 7, name: "Picture 7", descr: "sample alt", title: "sample title" },
              effectExtent: { l: 1, t: 2, r: 3, b: 4 },
              extent: { cx: 9525, cy: 9525 },
              graphic: {
                type: "image",
                blipFill: {
                  rasterImageId: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lk9N1wAAAABJRU5ErkJggg==",
                  srcRect: { l: 1000, t: 2000, r: 3000, b: 4000 },
                },
                spPr: { prstGeom: { prst: "roundRect" } },
              },
            },
            {
              type: "paraDrawing",
              drawingType: "anchor",
              docPr: { id: 8, name: "Floating Picture 8" },
              distT: 10,
              distB: 20,
              distL: 30,
              distR: 40,
              simplePos: { x: 5, y: 6 },
              relativeHeight: 251658240,
              behindDoc: false,
              layoutInCell: true,
              allowOverlap: false,
              positionH: { relativeFrom: "column", posOffset: 123 },
              positionV: { relativeFrom: "paragraph", align: "top" },
              wrapType: "through",
              wrapThrough: { edited: true },
              extent: { cx: 1000, cy: 2000 },
              graphic: {
                type: "image",
                blipFill: { rasterImageId: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lk9N1wAAAABJRU5ErkJggg==" },
              },
            },
          ],
        },
      ],
    },
  ],
};
const imageLongTail = withPackageLongTailParts(imageLongTailDoc);
const flatWithImage = jsonToFlatOpcWithStyles(imageLongTail.doc, [], undefined, undefined, {
  extraParts: imageLongTail.extraParts,
});
ok(imageLongTail.summary.media === 2, "data URI paraDrawing produces media parts");
ok(flatWithImage.includes('pkg:name="/word/media/image1.png"'), "data URI image emits media part");
ok(flatWithImage.includes('relationships/image'), "data URI image adds document image relationship");
ok(flatWithImage.includes('<pkg:binaryData>iVBORw0KGgo'), "data URI image stores binaryData");
ok(flatWithImage.includes('<w:drawing><wp:inline'), "paraDrawing emits minimal w:drawing");
ok(flatWithImage.includes('r:embed="rIdCocoImage1"'), "paraDrawing references generated image rId");
ok(flatWithImage.includes('<wp:docPr id="7" name="Picture 7" descr="sample alt" title="sample title"/>'), "paraDrawing preserves docPr alt metadata");
ok(flatWithImage.includes('<wp:effectExtent l="1" t="2" r="3" b="4"/>'), "paraDrawing preserves effect extent");
ok(flatWithImage.includes('<a:srcRect l="1000" t="2000" r="3000" b="4000"/>'), "paraDrawing preserves image crop srcRect");
ok(flatWithImage.includes('<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'), "paraDrawing preserves preset geometry");
ok(flatWithImage.includes('<wp:anchor distT="10" distB="20" distL="30" distR="40" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">'), "anchor paraDrawing emits wp:anchor positioning attributes");
ok(flatWithImage.includes('<wp:simplePos x="5" y="6"/><wp:positionH relativeFrom="column"><wp:posOffset>123</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>'), "anchor paraDrawing preserves position offsets and alignment");
ok(flatWithImage.includes('<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>'), "anchor paraDrawing preserves through wrapping");
const shapeDrawingDoc = {
  type: "document",
  drawingObjects: [
    {
      type: "paraDrawing",
      drawingType: "anchor",
      docPr: { id: 11, name: "WordArt 11" },
      distT: 1,
      distB: 2,
      distL: 3,
      distR: 4,
      positionH: { relativeFrom: "page", posOffset: 10 },
      positionV: { relativeFrom: "page", posOffset: 20 },
      wrapType: "none",
      extent: { cx: 1000, cy: 500 },
      graphic: {
        type: "shape",
        bodyPr: { fromWordArt: true },
        spPr: {
          ln: {
            w: 12700,
            fill: { fill: { color: { rgba: { red: 1, green: 2, blue: 3, alpha: 255 } }, type: "solid" }, type: "fill" },
          },
          fill: { fill: { color: { rgba: { red: 10, green: 20, blue: 30, alpha: 255 } }, type: "solid" }, type: "fill" },
        },
        content: {
          type: "docContent",
          content: [{ type: "paragraph", content: [{ type: "run", content: ["word art"] }] }],
        },
      },
    },
    {
      type: "paraDrawing",
      graphic: { type: "image", blipFill: { rasterImageId: "data:image/png;base64,abc=" } },
    },
    {
      type: "paraDrawing",
      drawingType: "anchor",
      docPr: { id: 12, name: "Line 12" },
      extent: { cx: 6120130, cy: 0 },
      graphic: {
        type: "connectShape",
        nvCxnSpPr: {
          cNvPr: { id: -1, name: "" },
          nvPr: { isPhoto: null, userDrawn: null },
          cNvCxnSpPr: {
            cxnSpLocks: { noChangeShapeType: true },
          },
        },
        spPr: {
          custGeom: { preset: "line" },
          fill: { fill: { type: "noFill" }, type: "fill" },
          ln: {
            w: 9525,
            prstDash: "solid",
            lineJoin: { type: "round" },
            fill: { fill: { color: { rgba: { red: 255, green: 0, blue: 0, alpha: 255 } }, type: "solid" }, type: "fill" },
          },
          xfrm: { off: { x: 0, y: 0 }, ext: { cx: 6120130, cy: 0 }, flipV: true },
        },
      },
    },
  ],
  content: [],
};
const flatWithTopLevelDrawing = jsonToFlatOpcWithStyles(shapeDrawingDoc, [], undefined, undefined, { includeDefaultParts: false });
ok(flatWithTopLevelDrawing.includes('<w:drawing><wp:anchor'), "top-level shape drawing emits anchored w:drawing placeholder");
ok(flatWithTopLevelDrawing.includes('<w:pict><v:shape'), "top-level shape drawing emits VML pict fallback placeholder");
ok(flatWithTopLevelDrawing.includes('<wp:positionH relativeFrom="page"><wp:posOffset>10</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>20</wp:posOffset></wp:positionV>'), "top-level shape drawing preserves anchor position");
ok(flatWithTopLevelDrawing.includes('<wp:wrapNone/>'), "top-level shape drawing preserves no-wrap layout");
ok(flatWithTopLevelDrawing.includes('style="width:0.08pt;height:0.04pt"'), "top-level shape VML fallback preserves extent style");
ok(flatWithTopLevelDrawing.includes('fillcolor="#0A141E"'), "top-level shape VML fallback preserves fill color");
ok(flatWithTopLevelDrawing.includes('strokecolor="#010203"'), "top-level shape VML fallback preserves stroke color");
ok(flatWithTopLevelDrawing.includes('strokeweight="1pt"'), "top-level shape VML fallback preserves stroke weight");
ok(flatWithTopLevelDrawing.includes('<w:t xml:space="preserve">word art</w:t>'), "top-level shape drawing preserves textbox text");
ok(flatWithTopLevelDrawing.includes('<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'), "top-level line preserves line preset geometry");
ok(flatWithTopLevelDrawing.includes('<wps:cNvCnPr><a:cxnSpLocks noChangeShapeType="1"/></wps:cNvCnPr>'), "top-level line preserves connector non-visual properties");
ok(flatWithTopLevelDrawing.includes('<a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="6120130" cy="0"/></a:xfrm>'), "top-level line preserves transform");
ok(flatWithTopLevelDrawing.includes('<a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'), "top-level line preserves stroke");
ok((flatWithTopLevelDrawing.match(/<w:drawing>/g) ?? []).length === 2, "top-level image drawing is not duplicated when media path handles it");
const bodyLineDrawing = {
  type: "paraDrawing",
  drawingType: "anchor",
  docPr: { id: 21, name: "Line 21" },
  positionH: { relativeFrom: "column", posOffset: -328930 },
  positionV: { relativeFrom: "paragraph", posOffset: 1233170 },
  extent: { cx: 6120130, cy: 0 },
  graphic: {
    type: "connectShape",
    nvCxnSpPr: {
      cNvPr: { id: -1, name: "" },
      nvPr: { isPhoto: null, userDrawn: null },
      cNvCxnSpPr: {
        cxnSpLocks: { noChangeShapeType: true },
      },
    },
    spPr: {
      custGeom: { preset: "line" },
      fill: { fill: { type: "noFill" }, type: "fill" },
      ln: {
        w: 9525,
        fill: {
          fill: {
            color: {
              rgba: { red: 0, green: 0, blue: 0, alpha: 255 },
              color: { rgba: { red: 255, green: 0, blue: 0, alpha: 255 } },
            },
            type: "solid",
          },
          type: "fill",
        },
      },
      xfrm: { off: { x: 0, y: 0 }, ext: { cx: 6120130, cy: 0 }, flipV: true },
    },
  },
};
const flatWithBodyLine = jsonToFlatOpcWithStyles(
  {
    type: "document",
    content: [
      {
        type: "paragraph",
        content: [{ type: "run", content: [bodyLineDrawing] }],
      },
    ],
    drawingObjects: [bodyLineDrawing],
  },
  [],
  undefined,
  undefined,
  { includeDefaultParts: false },
);
ok((flatWithBodyLine.match(/<w:drawing>/g) ?? []).length === 1, "body line is emitted once when drawingObjects contains the same id");
ok(!flatWithBodyLine.includes('oo:drawingObject="0"'), "body line keeps its original paragraph anchor");
ok(flatWithBodyLine.includes('<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'), "body line preserves line preset geometry");
ok(flatWithBodyLine.includes('<wps:cNvCnPr><a:cxnSpLocks noChangeShapeType="1"/></wps:cNvCnPr>'), "body line preserves connector non-visual properties");
ok(flatWithBodyLine.includes('<a:srgbClr val="FF0000"/>'), "body line prefers the nested real stroke color over fallback rgba");
const noteLongTailDoc = {
  type: "document",
  content: [
    {
      type: "paragraph",
      content: [
        {
          type: "run",
          content: [{ type: "footnoteRef", footnote: "101", customMarkFollows: true }],
          footnotes: [
            { id: "-1", type: "separator", content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "separator" }, " custom separator"] }] }] },
            { id: "0", type: "continuationSeparator", content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "continuationSeparator" }] }] }] },
            { id: "101", type: "normal", content: [{ type: "paragraph", content: [{ type: "run", content: ["footnote body"] }] }] },
          ],
        },
        {
          type: "run",
          content: [{ type: "endnoteRef", footnote: "202", customMarkFollows: true }],
          endnotes: [
            { id: "0", type: "continuationSeparator", content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "continuationSeparator" }, " end cont"] }] }] },
            { id: "202", content: [{ type: "paragraph", content: [{ type: "run", content: ["endnote body"] }] }] },
          ],
        },
      ],
    },
  ],
};
const noteLongTail = withPackageLongTailParts(noteLongTailDoc);
const flatWithNotes = jsonToFlatOpcWithStyles(noteLongTail.doc, [], undefined, undefined, {
  extraParts: noteLongTail.extraParts,
});
ok(noteLongTail.summary.footnotes === 1, "footnoteRef run produces one footnotes entry");
ok(noteLongTail.summary.endnotes === 1, "endnoteRef run produces one endnotes entry");
ok(flatWithNotes.includes('pkg:name="/word/footnotes.xml"'), "footnote body emits footnotes.xml part");
ok(flatWithNotes.includes('pkg:name="/word/endnotes.xml"'), "endnote body emits endnotes.xml part");
ok(flatWithNotes.includes('relationships/footnotes'), "footnotes relationship is emitted");
ok(flatWithNotes.includes('relationships/endnotes'), "endnotes relationship is emitted");
ok(flatWithNotes.includes('<w:footnoteReference w:id="101" w:customMarkFollows="1"/>'), "document emits footnoteReference custom mark");
ok(flatWithNotes.includes('<w:endnoteReference w:id="202" w:customMarkFollows="1"/>'), "document emits endnoteReference custom mark");
ok(flatWithNotes.includes('<w:footnote w:id="101" w:type="normal"><w:p><w:r><w:t xml:space="preserve">footnote body</w:t></w:r></w:p></w:footnote>'), "footnotes.xml preserves footnote body and type");
ok(flatWithNotes.includes('<w:endnote w:id="202"><w:p><w:r><w:t xml:space="preserve">endnote body</w:t></w:r></w:p></w:endnote>'), "endnotes.xml preserves endnote body");
ok(flatWithNotes.includes('<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/><w:t xml:space="preserve"> custom separator</w:t></w:r></w:p></w:footnote>'), "footnotes.xml preserves custom separator body");
ok((flatWithNotes.match(/<w:footnote w:[^>]*type="separator"/g) ?? []).length === 1, "custom footnote separator suppresses duplicate default separator");
ok(flatWithNotes.includes('<w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/><w:t xml:space="preserve"> end cont</w:t></w:r></w:p></w:endnote>'), "endnotes.xml preserves custom continuation separator body");
const flatWithoutDefaults = jsonToFlatOpcWithStyles(ooJsonDoc, styleDefs, undefined, undefined, { includeDefaultParts: false });
ok(!flatWithoutDefaults.includes('pkg:name="/word/theme/theme1.xml"'), "includeDefaultParts=false keeps legacy thin package shape");
{
  console.log("\n=== 4) SDK core 端到端：fromFlatOpc → buildStylesIndex 判标题层级 ===");

  // 4a) 带 styles.xml 整包：core 应判出标题层级
  const pkg = fromFlatOpc(flatWith);
  const idx = buildStylesIndex(pkg);
  ok(idx.size === 4, `读出 4 个样式定义（实际 ${idx.size}）`);

  const h1 = getStyleById(idx, "701");
  ok(h1 && h1.name === "heading 1", `styleId 701 → name=heading 1（实际 ${h1 && h1.name}）`);
  ok(h1 && h1.outlineLevel === 0, `styleId 701 → outlineLevel=0 即一级标题（实际 ${h1 && h1.outlineLevel}）`);

  const h2 = getStyleById(idx, "666");
  ok(h2 && h2.outlineLevel === 1, `styleId 666 → outlineLevel=1 即二级标题（实际 ${h2 && h2.outlineLevel}）`);

  const normal = getStyleById(idx, "665");
  ok(normal && normal.outlineLevel === undefined, `styleId 665(Normal) → 无 outlineLevel 即正文（实际 ${normal && normal.outlineLevel}）`);

  // 4b) 对照：薄壳（不挂 styles）→ buildStylesIndex 落空，证明风险真实存在
  const flatThin = jsonToFlatOpcWithStyles(ooJsonDoc, []); // 无样式
  const idxThin = buildStylesIndex(fromFlatOpc(flatThin));
  ok(idxThin.size === 0, `薄壳无 styles.xml → buildStylesIndex 空（判级落空，实际 size=${idxThin.size}）`);

  // 4c) 模拟"段落引用的 styleId"在 styles 表里能被查到（排版判级的实际调用链）
  const usedStyleIds = ["701", "666", "665"]; // 取自 document.xml 段落 pStyle
  const resolvable = usedStyleIds.every((id) => getStyleById(idx, id) !== null);
  ok(resolvable, "文档段落引用的每个数字 styleId 都能在重建 styles 表里命中");

  // 4d) Long-tail package closure: Flat OPC -> core package -> Flat OPC keeps long-tail parts/rels/binary.
  const longTailRoundTrip = toFlatOpc(fromFlatOpc(flatWithExtra));
  ok(longTailRoundTrip.includes('pkg:name="/word/header1.xml"'), "SDK round-trip keeps header part");
  ok(longTailRoundTrip.includes('pkg:name="/word/footer1.xml"'), "SDK round-trip keeps footer part");
  ok(longTailRoundTrip.includes('pkg:name="/word/comments.xml"'), "SDK round-trip keeps comments part");
  ok(longTailRoundTrip.includes('pkg:name="/word/footnotes.xml"'), "SDK round-trip keeps footnotes part");
  ok(longTailRoundTrip.includes('pkg:name="/word/endnotes.xml"'), "SDK round-trip keeps endnotes part");
  ok(longTailRoundTrip.includes('pkg:name="/word/media/image1.png"'), "SDK round-trip keeps media binary part");
  ok(longTailRoundTrip.includes('pkg:name="/word/_rels/header1.xml.rels"'), "SDK round-trip keeps header relationship part");
  ok(longTailRoundTrip.includes('relationships/image'), "SDK round-trip keeps nested image relationship");
}

console.log(`\n── 结果：${pass} 通过 / ${fail} 失败 ──`);
process.exit(fail ? 1 : 0);
