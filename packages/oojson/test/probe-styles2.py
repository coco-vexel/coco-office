# 探针2：在 docbuilder 里用 JS 直接挖 styles 表。
# 试三条路：①GetStyle 取已知标题样式看其属性 ②Document.ToJSON 看是否漏 ③直接读底层 GetStyles().ToJSON
import sys, os, json
DOCBUILDER_DIR = "D:/tools/docbuilder"
os.chdir(DOCBUILDER_DIR)
sys.path.insert(0, DOCBUILDER_DIR)
import docbuilder

src = sys.argv[1] if len(sys.argv) > 1 else "test/fixtures/styles.docx"

# 用 OpenFile + 注入 JS 脚本（builder.RunTextW 不行；用 Context eval 通过 GetGlobal 调）
# docbuilder python 没有直接 eval，改用对象模型逐步调。
b = docbuilder.CDocBuilder()
b.OpenFile(src.replace("\\", "/"), "")
ctx = b.GetContext()
g = ctx.GetGlobal()
api = g.GetProperty("Api")
doc = api.Call("GetDocument")

# 路1：取一个段落的 pStyle，再 GetStyle 看能否拿到定义
# 先列已应用样式：遍历段落找 pStyle
out = {}

# 尝试 doc.GetStyle("Heading 1") —— ONLYOFFICE 用显示名
for nm in ["Heading 1", "heading 1", "标题 1", "Title", "Normal"]:
    try:
        st = doc.Call("GetStyle", docbuilder.CDocBuilderValue(nm))
        # st 可能为 null
        sv = st.ToString() if st else "null"
        # 试 st.ToJSON()
        sj = None
        try:
            jv = st.Call("ToJSON")
            sj = jv.ToString()
        except Exception as e:
            sj = f"(ToJSON err: {e})"
        out[nm] = {"obj": sv[:60], "ToJSON": (sj[:600] if sj else None)}
    except Exception as e:
        out[nm] = f"GetStyle err: {e}"

print(json.dumps(out, ensure_ascii=False, indent=1))
b.CloseFile()
