# 探针：OnlyOffice 长尾 part 相关 API 形态。
# 用法: python test/probe-longtail.py [docx路径]
# 目标：确认哪些对象能 ToJSON，哪些只能作为 document.xml 里的引用/兜底项出现。
import sys, os, json

DOCBUILDER_DIR = "D:/tools/docbuilder"
os.chdir(DOCBUILDER_DIR)
sys.path.insert(0, DOCBUILDER_DIR)
import docbuilder

src = sys.argv[1] if len(sys.argv) > 1 else "test/fixtures/longtail.docx"
V = docbuilder.CDocBuilderValue

METHODS = [
    ("ToJSON_full", "ToJSON", [V(False), V(True), V(True), V(True), V(True), V(True)]),
    ("GetSections", "GetSections", []),
    ("GetAllComments", "GetAllComments", []),
    ("GetAllImages", "GetAllImages", []),
    ("GetAllDrawingObjects", "GetAllDrawingObjects", []),
    ("GetFootnotesFirstParagraphs", "GetFootnotesFirstParagraphs", []),
    ("GetEndNotesFirstParagraphs", "GetEndNotesFirstParagraphs", []),
]


def value_to_py(v, depth=0):
    if v is None:
        return None
    try:
        s = v.ToString()
    except Exception:
        s = str(v)
    if s and isinstance(s, str) and s[:1] in "[{":
        try:
            return json.loads(s)
        except Exception:
            pass
    return s


def try_to_json(obj):
    if obj is None:
        return None
    try:
        j = obj.Call("ToJSON")
        return value_to_py(j)
    except Exception as e:
        return {"__toJSON_error": str(e)}


def sample_array(value, limit=3):
    out = {"raw": value_to_py(value)}
    try:
        count = value.GetLength()
        out["length"] = int(count)
        items = []
        for i in range(min(int(count), limit)):
            item = value.Get(i)
            items.append({
                "string": value_to_py(item),
                "json": try_to_json(item),
            })
        out["items"] = items
        return out
    except Exception:
        return out

b = docbuilder.CDocBuilder()
try:
    b.OpenFile(src.replace("\\", "/"), "")
    ctx = b.GetContext()
    api = ctx.GetGlobal().GetProperty("Api")
    doc = api.Call("GetDocument")
    report = {"source": src, "methods": {}}

    for label, method, args in METHODS:
        try:
            value = doc.Call(method, *args)
            entry = sample_array(value)
            if label == "ToJSON_full" and isinstance(entry.get("raw"), dict):
                raw = entry["raw"]
                entry = {
                    "topKeys": list(raw.keys()),
                    "lenChars": len(json.dumps(raw, ensure_ascii=False)),
                    "themeType": type(raw.get("theme")).__name__ if "theme" in raw else None,
                    "sectionPrType": type(raw.get("sectionPr")).__name__ if "sectionPr" in raw else None,
                    "numberingType": type(raw.get("numbering")).__name__ if "numbering" in raw else None,
                    "stylesType": type(raw.get("styles")).__name__ if "styles" in raw else None,
                    "sectionPrSample": raw.get("sectionPr"),
                }
            report["methods"][label] = entry
        except Exception as e:
            report["methods"][label] = {"error": str(e)}

    print(json.dumps(report, ensure_ascii=False, indent=2)[:20000])
finally:
    try:
        b.CloseFile()
    except Exception:
        pass
