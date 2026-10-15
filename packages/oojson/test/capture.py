# 无头批量产 ONLYOFFICE-JSON：Python 驱动 ONLYOFFICE Document Builder。
#   用法: python test/capture.py <语料目录> [最多N] [输出目录]
# 对每篇 docx：OpenFile → Api.GetDocument().ToJSON(false,true) → 写 <名>.oojson.json。
# 走对象模型取返回值（不靠 console），不存文档=不沾水印。
import sys, os

DOCBUILDER_DIR = "D:/tools/docbuilder"
os.chdir(DOCBUILDER_DIR)
sys.path.insert(0, DOCBUILDER_DIR)
import docbuilder

corpus = sys.argv[1] if len(sys.argv) > 1 else "test/fixtures/source"
maxn = int(sys.argv[2]) if len(sys.argv) > 2 else 99999
out = sys.argv[3] if len(sys.argv) > 3 else "test/captured"
os.makedirs(out, exist_ok=True)

files = []
for root, dirs, fs in os.walk(corpus):
    dirs[:] = [d for d in dirs if d != "node_modules"]
    for f in fs:
        if f.lower().endswith(".docx") and not f.startswith("~$"):
            files.append(os.path.join(root, f))
files = files[:maxn]
print(f"语料: {corpus}\n篇数: {len(files)}\n")

b = docbuilder.CDocBuilder()
ok = err = 0
seen = {}
import time
t0 = time.time()
for f in files:
    base = os.path.splitext(os.path.basename(f))[0]
    seen[base] = seen.get(base, 0) + 1
    name = base if seen[base] == 1 else f"{base}__{seen[base]}"
    try:
        b.OpenFile(f.replace("\\", "/"), "")
        ctx = b.GetContext()
        g = ctx.GetGlobal()
        api = g.GetProperty("Api")
        doc = api.Call("GetDocument")
        jv = doc.Call("ToJSON", docbuilder.CDocBuilderValue(False), docbuilder.CDocBuilderValue(True))
        s = jv.ToString()
        b.CloseFile()
        if s and s.startswith("{"):
            open(os.path.join(out, name + ".oojson.json"), "w", encoding="utf-8").write(s)
            ok += 1
        else:
            err += 1
    except Exception as e:
        err += 1
        try: b.CloseFile()
        except Exception: pass
    sys.stdout.write(f"\r{ok+err}/{len(files)}  ok={ok} err={err}   ")
    sys.stdout.flush()
print(f"\n\n完成: ok={ok} err={err}  用时 {time.time()-t0:.0f}s  → {out}")
