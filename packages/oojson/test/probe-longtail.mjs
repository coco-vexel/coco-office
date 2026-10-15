// Probe ONLYOFFICE DocumentBuilder APIs that may provide package-level long-tail sources.
// Usage:
//   node test/probe-longtail.mjs [docx ...]
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dir, "../../../../..");
const DOCBUILDER_EXE = process.env.DOCBUILDER_EXE || "D:/tools/docbuilder/docbuilder.exe";
const MARKER = "@@LONGTAIL_PROBE@@";
const SKELETON_DOCX_DIR = path.join(repoRoot, "coco-office/examples/skeleton-cli/docx");
const DEFAULT_DOCS = [
  path.join(repoRoot, "_oo-test/oo-test.docx"),
  path.join(SKELETON_DOCX_DIR, "\u4f1a\u8bae\u7eaa\u8981.docx"),
  path.join(SKELETON_DOCX_DIR, "\u3010\u6837\u5f0f\u3011\u8bba\u6587\u6392\u7248_\u526f\u672c.docx"),
].filter((file) => fs.existsSync(file));

function builderScript(docxPath) {
  return `
function clip(s, n) {
  s = String(s == null ? "" : s);
  return s.length > n ? s.slice(0, n) + "..." : s;
}
function topKeys(x) {
  return x && typeof x === "object" && !Array.isArray(x) ? Object.keys(x).slice(0, 40) : [];
}
function collectText(x, out) {
  if (typeof x === "string") out.push(x);
  else if (Array.isArray(x)) for (var i = 0; i < x.length; i++) collectText(x[i], out);
  else if (x && typeof x === "object") for (var k in x) collectText(x[k], out);
}
function parseMaybeJson(s) {
  if (typeof s !== "string") return s;
  if (!s.length || (s[0] !== "{" && s[0] !== "[")) return s;
  try { return JSON.parse(s); } catch (e) { return s; }
}
function toJsonValue(obj, args) {
  if (!obj || typeof obj.ToJSON !== "function") return { ok: false, reason: "no ToJSON" };
  try {
    var raw = obj.ToJSON.apply(obj, args || []);
    var parsed = parseMaybeJson(raw);
    var text = [];
    collectText(parsed, text);
    return {
      ok: true,
      rawLength: String(raw == null ? "" : raw).length,
      parsedType: Array.isArray(parsed) ? "array" : typeof parsed,
      topKeys: topKeys(parsed),
      textLength: text.join("").length,
      textSample: clip(text.join(""), 600),
      sample: typeof parsed === "object" ? parsed : clip(parsed, 600)
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
function commentInfo(obj) {
  if (!obj || typeof obj !== "object") return null;
  var names = ["GetId", "GetText", "GetAuthorName", "GetQuoteText", "GetTime", "IsSolved"];
  var out = {};
  var any = false;
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    try {
      if (typeof obj[name] === "function") {
        out[name.replace(/^Get/, "").replace(/^Is/, "is")] = obj[name]();
        any = true;
      }
    } catch (e) {
      out[name] = { error: String(e) };
      any = true;
    }
  }
  return any ? out : null;
}
function summarizeValue(value) {
  var entry = {
    valueType: Array.isArray(value) ? "array" : typeof value,
    hasToJSON: !!(value && typeof value.ToJSON === "function")
  };
  if (Array.isArray(value)) {
    entry.length = value.length;
    entry.items = [];
    for (var i = 0; i < Math.min(value.length, 5); i++) {
      entry.items.push({
        index: i,
        valueType: typeof value[i],
        hasToJSON: !!(value[i] && typeof value[i].ToJSON === "function"),
        json: toJsonValue(value[i], []),
        jsonFull: toJsonValue(value[i], [false, true, true, true, true, true]),
        comment: commentInfo(value[i])
      });
    }
  } else if (value && typeof value === "object") {
    entry.keys = topKeys(value);
    entry.json = toJsonValue(value, []);
    entry.jsonFull = toJsonValue(value, [false, true, true, true, true, true]);
  } else {
    entry.value = clip(value, 600);
  }
  return entry;
}
function callDocMethod(doc, name) {
  if (!doc || typeof doc[name] !== "function") return { exists: false };
  try {
    return Object.assign({ exists: true }, summarizeValue(doc[name]()));
  } catch (e) {
    return { exists: true, error: String(e) };
  }
}

builder.OpenFile(${JSON.stringify(docxPath.replace(/\\/g, "/"))});
var doc = Api.GetDocument();
var parsed = JSON.parse(doc.ToJSON(false, true, true, true, true, true));
var docText = [];
collectText(parsed, docText);
var report = {
  document: {
    topKeys: topKeys(parsed),
    textLength: docText.join("").length,
    hasSectPr: !!parsed.sectPr,
    hasSectionPr: !!parsed.sectionPr,
    sectPrKeys: topKeys(parsed.sectPr || parsed.sectionPr),
    hasNumbering: !!parsed.numbering,
    hasStyles: !!parsed.styles,
    hasTheme: !!parsed.theme
  },
  methods: {}
};
var names = [
  "GetSections",
  "GetAllComments",
  "GetAllImages",
  "GetAllDrawingObjects",
  "GetFootnotesFirstParagraphs",
  "GetEndNotesFirstParagraphs"
];
for (var n = 0; n < names.length; n++) report.methods[names[n]] = callDocMethod(doc, names[n]);
console.log(${JSON.stringify(MARKER)} + JSON.stringify(report));
builder.CloseFile();
`;
}

function runProbe(docxPath, index) {
  if (!fs.existsSync(DOCBUILDER_EXE)) throw new Error(`Document Builder not found: ${DOCBUILDER_EXE}`);
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "oojson-longtail-"));
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
  if (result.status !== 0) throw new Error(`docbuilder exited ${result.status}: ${result.stderr || result.stdout}`);
  const stdout = result.stdout || "";
  const at = stdout.lastIndexOf(MARKER);
  if (at < 0) throw new Error(`docbuilder output missing ${MARKER}: ${stdout.slice(0, 500)}`);
  return JSON.parse(stdout.slice(at + MARKER.length).trim());
}

function compactMethod(entry) {
  if (!entry.exists) return "missing";
  if (entry.error) return `error: ${entry.error}`;
  if (entry.valueType === "array") return `array(${entry.length})`;
  if (entry.hasToJSON) return "object(ToJSON)";
  return entry.valueType;
}

const docs = process.argv.slice(2).filter((arg) => arg !== "--json").length ? process.argv.slice(2).filter((arg) => arg !== "--json") : DEFAULT_DOCS;
if (process.argv.includes("--json")) process.env.PROBE_JSON = "1";
if (!docs.length) {
  console.error("No docx files supplied and no default fixtures found.");
  process.exit(2);
}

let failed = 0;
for (let i = 0; i < docs.length; i++) {
  const docx = path.resolve(docs[i]);
  try {
    const report = runProbe(docx, i + 1);
    console.log(`\n# ${path.basename(docx)}`);
    console.log(`document: text=${report.document.textLength}, sectPr=${report.document.hasSectPr || report.document.hasSectionPr}, numbering=${report.document.hasNumbering}, styles=${report.document.hasStyles}, theme=${report.document.hasTheme}`);
    for (const [name, entry] of Object.entries(report.methods)) {
      console.log(`${name}: ${compactMethod(entry)}`);
      if (entry.items?.length) {
        for (const item of entry.items.slice(0, 3)) {
          const json = item.json?.ok ? item.json : item.jsonFull;
          const text = json?.textSample ? ` text=${JSON.stringify(json.textSample.slice(0, 120))}` : "";
          const keys = json?.topKeys?.length ? ` keys=${json.topKeys.join(",")}` : "";
          const comment = item.comment ? ` comment=${JSON.stringify(item.comment).slice(0, 180)}` : "";
          console.log(`  [${item.index}]${keys}${text}${comment}`);
        }
      }
    }
    if (process.env.PROBE_JSON === "1") console.log(`${MARKER}${JSON.stringify({ file: docx, report })}`);
  } catch (e) {
    failed++;
    console.log(`\n[ERROR] ${docx}`);
    console.log(`  ${(e && e.stack) || e}`);
  }
}
process.exit(failed ? 1 : 0);