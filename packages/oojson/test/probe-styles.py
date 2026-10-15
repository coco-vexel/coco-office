# 探针：实测 ONLYOFFICE ToJSON 各参数组合是否带 styles 段。
#   用法: python test/probe-styles.py [docx路径]
# 对同一篇 docx 跑 GetDocument().ToJSON(...) 多种参数，打印顶层键 + styles 段样本。
import sys, os, json

DOCBUILDER_DIR = "D:/tools/docbuilder"
os.chdir(DOCBUILDER_DIR)
sys.path.insert(0, DOCBUILDER_DIR)
import docbuilder

src = sys.argv[1] if len(sys.argv) > 1 else "test/fixtures/styles.docx"
V = docbuilder.CDocBuilderValue

# 待测参数组合：(label, [args])
combos = [
    ("ToJSON()", []),
    ("ToJSON(false,true)", [V(False), V(True)]),
    ("ToJSON(true,true)", [V(True), V(True)]),
    ("ToJSON(true,true,true,true)", [V(True), V(True), V(True), V(True)]),
]

for label, args in combos:
    b = docbuilder.CDocBuilder()
    try:
        b.OpenFile(src.replace("\\", "/"), "")
        ctx = b.GetContext()
        api = ctx.GetGlobal().GetProperty("Api")
        doc = api.Call("GetDocument")
        jv = doc.Call("ToJSON", *args)
        s = jv.ToString()
        b.CloseFile()
        print(f"\n===== {label} =====")
        if not (s and s.startswith("{")):
            print("  (非 JSON 返回)", repr(s[:80]))
            continue
        j = json.loads(s)
        print("  top keys:", list(j.keys()))
        print("  len(chars):", len(s))
        if "styles" in j:
            st = j["styles"]
            print("  styles type:", type(st).__name__)
            # styles 可能是 {style:[...]} 或 数组
            arr = st.get("style") if isinstance(st, dict) else st
            if isinstance(st, dict):
                print("  styles dict keys:", list(st.keys()))
            if isinstance(arr, list) and arr:
                print("  style count:", len(arr))
                # 找一个看起来是标题的
                def is_heading(x):
                    nm = json.dumps(x, ensure_ascii=False).lower()
                    return "heading" in nm or "标题" in nm or "outline" in nm
                head = next((x for x in arr if is_heading(x)), arr[0])
                print("  sample style keys:", list(head.keys()) if isinstance(head, dict) else head)
                print("  sample style:", json.dumps(head, ensure_ascii=False)[:1200])
        else:
            print("  NO styles key")
    except Exception as e:
        print(f"\n===== {label} ===== EXCEPTION:", e)
        try: b.CloseFile()
        except Exception: pass
