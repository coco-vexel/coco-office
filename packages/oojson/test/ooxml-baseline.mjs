// Validate against real docx OOXML:
//   original docx word/document.xml vs DocumentBuilder ToJSON -> jsonToOoxml(document.xml)
// Usage:
//   node test/ooxml-baseline.mjs [docx ...]
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { jsonToOoxml } from "../src/json-to-ooxml.mjs";
import { withPackageLongTailParts } from "../src/long-tail-parts.mjs";

const __dir = path.dirname(fileURLToPath(import.meta.url));
const DOCBUILDER_EXE = process.env.DOCBUILDER_EXE || "D:/tools/docbuilder/docbuilder.exe";
const MARKER = "@@OOJSON@@";

const DEFAULT_DOCS = [];

function flatOpcPart(name, contentType, innerXml) {
  return '<pkg:part pkg:name="' + name + '" pkg:contentType="' + contentType + '"><pkg:xmlData>' + innerXml + '</pkg:xmlData></pkg:part>';
}

async function makeGeneratedLongTailDocx() {
  const { fromFlatOpc, toZip } = await import("@coco-vexel/coco-office/pack");
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "oojson-real-longtail-"));
  const outPath = path.join(tmpDir, "comments-footnotes-endnotes.docx");
  const wordCt = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
  const relsCt = "application/vnd.openxmlformats-package.relationships+xml";
  const docXml =
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' +
    '<w:p><w:r><w:t>Footnote here</w:t><w:footnoteReference w:id="2"/></w:r><w:r><w:t>Endnote here</w:t><w:endnoteReference w:id="2"/></w:r></w:p>' +
    '<w:sectPr/></w:body></w:document>';
  const commentsXml =
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:comment w:id="0" w:author="Coco"><w:p><w:r><w:t>Comment body</w:t></w:r></w:p></w:comment>' +
    '</w:comments>';
  const footnotesXml =
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote>' +
    '</w:footnotes>';
  const endnotesXml =
    '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
    '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
    '<w:endnote w:id="2"><w:p><w:r><w:t>Endnote body</w:t></w:r></w:p></w:endnote>' +
    '</w:endnotes>';
  const contentTypesXml =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="' + relsCt + '"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="' + wordCt + 'document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="' + wordCt + 'comments+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="' + wordCt + 'footnotes+xml"/>' +
    '<Override PartName="/word/endnotes.xml" ContentType="' + wordCt + 'endnotes+xml"/>' +
    '</Types>';
  const rootRelsXml =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
  const docRelsXml =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
    '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
    '<Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>' +
    '</Relationships>';
  const flatOpc =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><?mso-application progid="Word.Document"?>' +
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    flatOpcPart('/[Content_Types].xml', 'application/vnd.openxmlformats-package.content-types+xml', contentTypesXml) +
    flatOpcPart('/_rels/.rels', relsCt, rootRelsXml) +
    flatOpcPart('/word/_rels/document.xml.rels', relsCt, docRelsXml) +
    flatOpcPart('/word/document.xml', wordCt + 'document.main+xml', docXml) +
    flatOpcPart('/word/comments.xml', wordCt + 'comments+xml', commentsXml) +
    flatOpcPart('/word/footnotes.xml', wordCt + 'footnotes+xml', footnotesXml) +
    flatOpcPart('/word/endnotes.xml', wordCt + 'endnotes+xml', endnotesXml) +
    '</pkg:package>';
  fs.writeFileSync(outPath, Buffer.from(await toZip(fromFlatOpc(flatOpc))));
  return outPath;
}
const TAGS = [
  "w:p", "w:r", "w:t", "w:tbl", "w:tr", "w:tc", "w:drawing", "w:pict",
  "w:sectPr", "w:pStyle", "w:rStyle", "w:numPr", "w:commentReference",
  "w:footnoteReference", "w:endnoteReference", "w:br", "w:tab",
];

function xmlText(xml) {
  return [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((m) => decodeXml(m[1]))
    .join("");
}

function decodeXml(s) {
  return String(s)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function countTag(xml, tag) {
  const escaped = tag.replace(":", "\\:");
  return (xml.match(new RegExp(`<${escaped}(?=[\\s>/])`, "g")) ?? []).length;
}

function fingerprint(xml) {
  const counts = Object.fromEntries(TAGS.map((tag) => [tag, countTag(xml, tag)]));
  return {
    chars: xml.length,
    textChars: xmlText(xml).length,
    counts,
  };
}

function diffCounts(a, b) {
  return TAGS
    .map((tag) => ({ tag, original: a.counts[tag], generated: b.counts[tag], delta: b.counts[tag] - a.counts[tag] }))
    .filter((row) => row.delta !== 0);
}

function firstTextDiff(a, b) {
  const n = Math.min(a.length, b.length);
  let i = 0;
  while (i < n && a[i] === b[i]) i++;
  if (i === a.length && i === b.length) return null;
  return {
    index: i,
    original: a.slice(Math.max(0, i - 30), i + 80),
    generated: b.slice(Math.max(0, i - 30), i + 80),
  };
}

function normalizeTextForReport(s) {
  return s.replace(/[\u0009\u0020\u00a0\u3000]+/g, " ");
}

function expectedLongTailFor(file) {
  if (path.basename(file) === "comments-footnotes-endnotes.docx") {
    return {
      counts: { comments: 1, footnotes: 1, endnotes: 1 },
      partText: {
        "/word/comments.xml": ["Comment body"],
        "/word/footnotes.xml": ["Footnote body"],
        "/word/endnotes.xml": ["Endnote body"],
      },
    };
  }
  return null;
}

async function originalDocumentXml(docxPath) {
  const { fromZip, getPart } = await import("@coco-vexel/coco-office/pack");
  const pkg = await fromZip(fs.readFileSync(docxPath));
  const part = getPart(pkg, "word/document.xml");
  if (!part) throw new Error(`missing word/document.xml in ${docxPath}`);
  const bytes = part._originalBytes ?? part.originalBytes;
  if (!bytes) throw new Error("core pack did not expose original bytes for word/document.xml");
  return new TextDecoder().decode(bytes);
}

function builderScript(docxPath) {
  return `builder.OpenFile(${JSON.stringify(docxPath.replace(/\\/g, "/"))});\n` +
    `var doc = Api.GetDocument();\n` +
    `var parsed = JSON.parse(doc.ToJSON(false, true, true, true, true, true));\n` +
    `try {\n` +
    `  var drawings = doc.GetAllDrawingObjects();\n` +
    `  parsed.drawingObjects = [];\n` +
    `  if (Array.isArray(drawings)) {\n` +
    `    for (var i = 0; i < drawings.length; i++) {\n` +
    `      try { parsed.drawingObjects.push(JSON.parse(drawings[i].ToJSON())); } catch (e) {}\n` +
    `    }\n` +
    `  }\n` +
    `} catch (e) { parsed.drawingObjects = []; }\n` +
    `console.log(${JSON.stringify(MARKER)} + JSON.stringify(parsed));\n` +
    `builder.CloseFile();\n`;
}

function readOojsonWithDocumentBuilder(docxPath, index) {
  if (!fs.existsSync(DOCBUILDER_EXE)) throw new Error(`Document Builder not found: ${DOCBUILDER_EXE}`);
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "oojson-ooxml-"));
  const asciiDocx = path.join(tmpDir, `source-${index}.docx`);
  const scriptPath = path.join(tmpDir, `probe-${index}.docbuilder`);
  fs.copyFileSync(docxPath, asciiDocx);
  fs.writeFileSync(scriptPath, builderScript(asciiDocx), "utf8");
  const result = spawnSync(DOCBUILDER_EXE, [scriptPath], {
    cwd: tmpDir,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(`docbuilder exited ${result.status}: ${result.stderr || result.stdout}`);
  }
  const stdout = result.stdout || "";
  const at = stdout.lastIndexOf(MARKER);
  if (at < 0) throw new Error(`docbuilder output missing ${MARKER}: ${stdout.slice(0, 500)}`);
  const jsonText = stdout.slice(at + MARKER.length).trim();
  return JSON.parse(jsonText);
}

async function checkDoc(docxPath, index) {
  const originalXml = await originalDocumentXml(docxPath);
  const oojson = readOojsonWithDocumentBuilder(docxPath, index);
  const longTail = withPackageLongTailParts(oojson);
  const generatedXml = jsonToOoxml(longTail.doc);
  const extraPartText = Object.fromEntries(
    longTail.extraParts
      .filter((part) => part.xml != null)
      .map((part) => [part.name, xmlText(String(part.xml))]),
  );
  const originalFp = fingerprint(originalXml);
  const generatedFp = fingerprint(generatedXml);
  const originalText = xmlText(originalXml);
  const generatedText = xmlText(generatedXml);
  const normalizedTextEqual = normalizeTextForReport(originalText) === normalizeTextForReport(generatedText);
  const textDiff = firstTextDiff(originalText, generatedText);
  const countDiffs = diffCounts(originalFp, generatedFp);
  return {
    file: docxPath,
    originalFp,
    generatedFp,
    textEqual: originalText === generatedText,
    normalizedTextEqual,
    textDiff,
    countDiffs,
    longTail: longTail.summary,
    extraPartText,
  };
}

const generatedDefaultDocs = process.argv.slice(2).length ? [] : [await makeGeneratedLongTailDocx()];
const docs = process.argv.slice(2).length ? process.argv.slice(2) : [...DEFAULT_DOCS, ...generatedDefaultDocs];
if (!docs.length) {
  console.error("No docx files supplied and no default fixtures found.");
  process.exit(2);
}

let failed = 0;
let exact = 0;
let accepted = 0;
console.log(`OOXML baseline docs: ${docs.length}`);
for (let i = 0; i < docs.length; i++) {
  const docx = path.resolve(docs[i]);
  try {
    const r = await checkDoc(docx, i + 1);
    const expectedLongTail = expectedLongTailFor(r.file);
    const countOk = !expectedLongTail || Object.entries(expectedLongTail.counts).every(([key, value]) => r.longTail[key] === value);
    const partTextOk = !expectedLongTail || Object.entries(expectedLongTail.partText).every(
      ([name, required]) => required.every((text) => (r.extraPartText[name] || "").includes(text)),
    );
    const longTailOk = countOk && partTextOk;
    const acceptable = (r.textEqual || r.normalizedTextEqual) && longTailOk;
    const textMark = r.textEqual && longTailOk ? "OK" : acceptable ? "WARN" : "DIFF";
    const structuralDeltas = r.countDiffs.length;
    if (r.textEqual && longTailOk) exact++;
    if (acceptable) accepted++;
    else failed++;
    console.log(`\n[${textMark}] ${path.basename(docx)}`);
    console.log(`  original:  xml=${r.originalFp.chars} text=${r.originalFp.textChars}`);
    console.log(`  generated: xml=${r.generatedFp.chars} text=${r.generatedFp.textChars}`);
    if (!r.textEqual) {
      console.log(`  normalized whitespace text: ${r.normalizedTextEqual ? "OK" : "DIFF"}`);
    }
    console.log(`  long-tail parts: header=${r.longTail.header} footer=${r.longTail.footer} comments=${r.longTail.comments} media=${r.longTail.media} footnotes=${r.longTail.footnotes} endnotes=${r.longTail.endnotes}`);
    if (expectedLongTail && !longTailOk) {
      console.log(`  expected long-tail: ${JSON.stringify(expectedLongTail.counts)}`);
      console.log(`  extra part text: ${JSON.stringify(r.extraPartText)}`);
    }
    console.log(`  tag deltas: ${structuralDeltas}`);
    for (const d of r.countDiffs.slice(0, 16)) {
      console.log(`    ${d.tag}: original=${d.original} generated=${d.generated} delta=${d.delta}`);
    }
    if (r.textDiff) {
      console.log(`  first text diff @${r.textDiff.index}`);
      console.log(`    original : ${JSON.stringify(r.textDiff.original)}`);
      console.log(`    generated: ${JSON.stringify(r.textDiff.generated)}`);
    }
  } catch (e) {
    failed++;
    console.log(`\n[ERROR] ${docx}`);
    console.log(`  ${(e && e.stack) || e}`);
  }
}

console.log(`\nResult: ${exact}/${docs.length} exact text-identical; ${accepted}/${docs.length} accepted against whitespace-normalized original`);
process.exit(failed ? 1 : 0);
