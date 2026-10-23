# Coco Office

面向浏览器和 Node.js 的 Office 文档 SDK，以及实验性的 Word、Excel、PPT
浏览器编辑器。文档在本地读取、修改和导出，不要求安装 Office。

项目已经开源并发布 `v1.0.2`。Core、Editor、OOJSON 及三份自制 Excel
测试样例采用 Apache-2.0，具体范围见 [LICENSING.md](LICENSING.md)。

## 本地运行

使用 Node.js 22 或 24、Corepack 和仓库锁定的 pnpm：

```bash
corepack pnpm install --frozen-lockfile
corepack pnpm build
corepack pnpm example
corepack pnpm demo
```

`example` 会生成三种格式的纯合成样例，修改后导出到 `examples/output/`，
重新读取并断言修改成功。`demo` 提供本地网页，可打开自有文档或生成样例，
编辑后点击 Export copy 下载副本。文档内容不上传；切换文档前请先导出当前修改。

给外部开发者离线试用时，运行 `corepack pnpm bundle:sdk` 会在
`.release-check/sdk/` 生成包含 Core 安装包、公开文档和 Word/Excel/PPT
可运行示例的 ZIP；同一文件会附在对应的 GitHub Release 中。

## 边界

- `core`：OOXML / Flat OPC 真值层和 Word、Excel、PPT 公共 API。
- `editor`：复用 Core 的浏览器交互层，不建立第二套文档内核。
- `oojson`：ONLYOFFICE 格式转换，采用 Apache-2.0；纳入 npm 发布目标，提供独立转换 Demo。
- Agent、具体 Office 宿主调用和商业产品不属于本仓库。

编辑器和高级格式支持仍在演进。能保留不等于能编辑，结构往返通过也不等于
与 Office 像素一致。具体能力见 [Word](docs/core/getting-started.md)、
[Excel](docs/core/excel/README.md)、[PPT](docs/core/ppt/README.md)。

验证命令、包说明和 Demo 构建见 [英文 README](README.md)。
贡献流程和 PR 要求见 [CONTRIBUTING.md](CONTRIBUTING.md)。问题请使用仓库的
结构化 Issue 表单；安全漏洞请按照 [SECURITY.md](SECURITY.md) 私下报告，
不要在公开 Issue 中附加客户文档、凭证或未披露漏洞。社区参与须遵守
[CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)。

三个库的 `1.0.2` 版本均已发布到 npm。详见[功能与运行环境核对](docs/runtime-and-demos.md)、[OOJSON 说明](packages/oojson/README.md)和[npm 发布流程](docs/npm-release.md)。Demo 首页包含 OOJSON 转换页入口。
