// Probe ONLYOFFICE DocumentBuilder comment APIs.
// Usage:
//   node test/probe-comment-api.mjs [docx]
//   node test/probe-comment-api.mjs --try-add [docx]
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dir, "../../../../..");
const DOCBUILDER_EXE = process.env.DOCBUILDER_EXE || "D:/tools/docbuilder/docbuilder.exe";
const MARKER = "@@COMMENT_API_PROBE@@";
const DEFAULT_DOCX = path.join(repoRoot, "_oo-test/oo-test.docx");

function builderScript(docxPath, tryAdd) {
  return `
function clip(value, max) {
  var text = String(value == null ? "" : value);
  return text.length > max ? text.slice(0, max) + "..." : text;
}
function methodNames(value, pattern) {
  var names = [];
  var seen = {};
  var current = value;
  while (current) {
    try {
      var own = Object.getOwnPropertyNames(current);
      for (var i = 0; i < own.length; i++) {
        var name = own[i];
        if (seen[name]) continue;
        seen[name] = true;
        var member;
        try { member = value[name]; } catch (e) { continue; }
        if (typeof member === "function" && (!pattern || pattern.test(name))) names.push(name);
      }
    } catch (e) {}
    current = Object.getPrototypeOf(current);
  }
  names.sort();
  return names;
}
function parseJson(value) {
  if (typeof value !== "string") return value;
  try { return JSON.parse(value); } catch (e) { return null; }
}
function commentInfo(comment) {
  var out = {};
  var names = ["GetId", "GetText", "GetQuoteText", "GetAuthorName", "GetTime", "IsSolved"];
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (comment && typeof comment[name] === "function") {
      try { out[name] = comment[name](); } catch (e) { out[name] = { error: String(e) }; }
    }
  }
  out.methods = methodNames(comment, /comment|text|quote|author|time|solv|delete|reply/i).slice(0, 80);
  return out;
}
function call(fn) {
  try { return { ok: true, value: fn() }; } catch (e) { return { ok: false, error: String(e) }; }
}

builder.OpenFile(${JSON.stringify(docxPath.replace(/\\/g, "/"))});
var doc = Api.GetDocument();
var parsed = parseJson(doc.ToJSON(false, true, true, true, true, true)) || {};
var report = {
  docMethods: methodNames(doc, /comment|range|select|move|element|paragraph/i).slice(0, 120),
  apiMethods: methodNames(Api, /comment|range|paragraph|document/i).slice(0, 120),
  jsonKeys: parsed && typeof parsed === "object" ? Object.keys(parsed).slice(0, 80) : [],
  commentMetadataCount: Array.isArray(parsed.commentMetadata) ? parsed.commentMetadata.length : 0,
  comments: [],
  tryAdd: null
};

var before = doc.GetAllComments && doc.GetAllComments();
if (Array.isArray(before)) {
  for (var i = 0; i < before.length; i++) report.comments.push(commentInfo(before[i]));
}

if (${tryAdd ? "true" : "false"}) {
  report.tryAdd = call(function () {
    var target = null;
    if (typeof doc.GetElement === "function") target = doc.GetElement(0);
    if (!target && typeof doc.GetContent === "function") {
      var content = doc.GetContent();
      if (content && typeof content.GetElement === "function") target = content.GetElement(0);
    }
    if (!target) return { added: false, reason: "no paragraph target found" };
    if (typeof Api.AddComment === "function") {
      Api.AddComment(target, "Coco comment API probe", "Coco");
      return { added: true, via: "Api.AddComment(target, text, author)" };
    }
    if (typeof target.AddComment === "function") {
      target.AddComment("Coco comment API probe", "Coco");
      return { added: true, via: "target.AddComment(text, author)" };
    }
    return {
      added: false,
      reason: "no AddComment API found",
      targetMethods: methodNames(target, /comment|range|select|text/i).slice(0, 80)
    };
  });
}

console.log(${JSON.stringify(MARKER)} + JSON.stringify(report));
builder.CloseFile();
`;
}

function runProbe(docxPath, tryAdd) {
  if (!fs.existsSync(DOCBUILDER_EXE)) throw new Error(`Document Builder not found: ${DOCBUILDER_EXE}`);
  if (!fs.existsSync(docxPath)) throw new Error(`docx not found: ${docxPath}`);
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "oojson-comment-api-"));
  const asciiDocx = path.join(tmpDir, "source.docx");
  const scriptPath = path.join(tmpDir, "probe.docbuilder");
  fs.copyFileSync(docxPath, asciiDocx);
  fs.writeFileSync(scriptPath, builderScript(asciiDocx, tryAdd), "utf8");
  const result = spawnSync(DOCBUILDER_EXE, [scriptPath], {
    cwd: tmpDir,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`docbuilder exited ${result.status}: ${result.stderr || result.stdout}`);
  const stdout = result.stdout || "";
  const at = stdout.lastIndexOf(MARKER);
  if (at < 0) throw new Error(`docbuilder output missing ${MARKER}: ${stdout.slice(0, 800)}`);
  return JSON.parse(stdout.slice(at + MARKER.length).trim());
}

function clip(value, max) {
  const text = String(value == null ? "" : value);
  return text.length > max ? `${text.slice(0, max)}...` : text;
}

const args = process.argv.slice(2);
const tryAdd = args.includes("--try-add");
const docxArg = args.find((arg) => arg !== "--try-add");
const docxPath = path.resolve(docxArg || DEFAULT_DOCX);

const report = runProbe(docxPath, tryAdd);
console.log(`# ${path.basename(docxPath)}`);
console.log(`comments: ${report.comments.length}`);
console.log(`commentMetadata: ${report.commentMetadataCount}`);
console.log(`doc methods: ${report.docMethods.join(", ") || "(none)"}`);
console.log(`Api methods: ${report.apiMethods.join(", ") || "(none)"}`);
for (let i = 0; i < report.comments.length; i++) {
  const comment = report.comments[i];
  console.log(`  [${i}] text=${JSON.stringify(clip(comment.GetText, 120))} quote=${JSON.stringify(clip(comment.GetQuoteText, 120))}`);
}
if (report.tryAdd) console.log(`tryAdd: ${JSON.stringify(report.tryAdd)}`);
