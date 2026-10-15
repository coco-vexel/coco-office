# 功能与运行环境核对

本仓库发布三个 npm 库：Core、Editor、OOJSON。根工作区不发包。Node 22/24 是维护者安装、构建和测试环境；浏览器部署使用静态产物，不要求用户电脑安装 Node。

| 库     | 负责的功能                                                                                                             | 运行入口与限制                                                                                                 | 必跑演示/验证                                                              |
| ------ | ---------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| Core   | OOXML ZIP/Flat OPC 读写；Word 段落/范围及变更；Excel 工作簿、单元格/范围与公式子集；PPT 幻灯片、文本、形状、表格及导出 | Node 或浏览器；ESM/CJS 和类型声明；inline 默认示例，Worker 需宿主配置；不调用 Office 宿主 API                  | `pnpm example` 的三格式生成→修改→另存→重读断言；包消费者 Node Worker 检查  |
| Editor | Word DOM 编辑；Canvas2D 表格；固定画布 PPT 编辑，选择、命令及导出                                                      | 浏览器 DOM 或 Electron renderer；ESM/CJS 和类型声明；文件对话框、存储、IPC 由接入方提供；高级布局/对象支持有限 | `pnpm demo` 首页，逐个打开 Word/Excel/PPT 样例，改动后导出副本并重新打开   |
| OOJSON | ONLYOFFICE JSON 与 Word XML/Flat OPC 转换，写回正文/诊断；PPT 内容/模型适配                                            | Node ESM、浏览器 bundler 或 IIFE；暂未提供 CJS/TypeScript 声明；不含 ONLYOFFICE 服务，不是 Excel JSON 转换库   | `node examples/oojson-roundtrip.mjs`；Demo `/oojson.html`；`test:portable` |

## 公共开发和演示环境

安装 Node.js 22 或 24，并确保 Corepack 可用；使用 packageManager 锁定的 pnpm 10.12.4。初次安装需访问公共 npm registry。基础检查不需要 Python、Office/WPS/ONLYOFFICE、私有仓库、密钥或真实客户文档。

```bash
corepack pnpm install --frozen-lockfile
corepack pnpm build
corepack pnpm example
corepack pnpm demo
```

打开 Vite 输出的本机 HTTP 地址，不能直接双击 HTML。首页链接进入 OOJSON 演示页。浏览器需要 DOM、Canvas2D、文件选择和 Blob 下载能力；当前验证使用 Chromium，不承诺所有浏览器已覆盖。

静态构建：`corepack pnpm build:demo`，产物在 `packages/editor/demo-dist/`，包含 `index.html` 和 `oojson.html`。这两个页面必须一起部署。文档在浏览器本地处理；示例不包含上传服务。

## 必须保留的自动检查

```bash
corepack pnpm test:ci
corepack pnpm --filter @coco-vexel/coco-office test:types:excel
corepack pnpm example
corepack pnpm build:demo
corepack pnpm licenses:packages
corepack pnpm verify:packages
```

包消费者验证会安装三个实际 tarball，检查导出文件、许可、workspace 依赖替换、Core ESM/CJS/Node Worker、OOJSON 转换、两页 Demo 构建和 Core/Editor TypeScript 入口。OOJSON 暂无类型声明，文档明确这一限制。

## 可选真实宿主测试环境

| 检查                | 额外环境                                                                                   | 不应混同的结论                   |
| ------------------- | ------------------------------------------------------------------------------------------ | -------------------------------- |
| OOJSON `test:ooxml` | ONLYOFFICE DocumentBuilder，设置 `DOCBUILDER_EXE`；不可并发运行该宿主                      | 转换器单测不等于实际宿主写回成功 |
| PPT/Word 参考渲染   | Windows、Microsoft Office 或 WPS、相关 Python/COM 工具，见 `packages/editor/dev/README.md` | 导出重读不等于像素一致           |
| 大语料回归          | 维护者自备授权语料，设置 `COCO_PPTX_FIXTURE_DIR` / `COCO_DOCX_FIXTURE_DIR` 等              | 没有语料时明确报告跳过           |

Node 22/24 × Ubuntu/Windows 已配置 CI。未实际执行的矩阵或宿主环境不能写成已验证。
