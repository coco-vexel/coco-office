# OnlyOffice oojson TODO

## 长尾 part 当前状态

本轮确认 `GetFileToDownload("docx")` 不作为主路，读链路继续走 `Api.GetDocument().ToJSON(...)` + oojson 重建。当前已经接入：

- `ToJSON(false, true, true, true, true, true)`：正文、theme、sectPr、numbering、styles。
- 默认 Word 包 part：`theme/theme1.xml`、`fontTable.xml`、`settings.xml`、`webSettings.xml`；theme/fontTable 优先使用 `doc.theme.themeElements.clrScheme/fontScheme` 的真实主题色与字体，`fmtScheme` 会保留 name、简单 solid fill/bgFill 与 lineStyle。
- `extraParts`：支持额外 XML/二进制 part、document relationship、source-part relationship。
- 真实页眉页脚：从 `sectPr.headerReference/footerReference` 的 `docContent` 生成 `/word/headerN.xml`、`/word/footerN.xml`，并把正文 `sectPr` 改成 `r:id` 引用；页眉页脚内的 data URI 图片也会生成 `/word/media/*` 与对应 header/footer source-part relationship。
- 真实批注：从正文 `commentRangeStart/commentRangeEnd/commentReference` 锚点生成 `comments.xml`，并输出 `<w:commentRangeStart>` / `<w:commentRangeEnd>` / `<w:commentReference>`；宿主侧通过 `GetAllComments()` 补充 `w:date`，并用 `commentsExtended.xml` 保留 solved/done 状态。
- 媒体图片：正文和页眉页脚 `paraDrawing` 中的 `data:image/...;base64,...` 会提取为 `/word/media/imageN.*`，输出 DrawingML `<w:drawing>` 并用 `r:embed` 关联 media part；若 JSON 提供 `docPr.descr/title`、`blipFill.srcRect`、`spPr.prstGeom`，会保留 alt metadata、effectExtent、裁剪与 preset geometry。`【样式】论文排版_副本.docx` 基线已显示 `media=1` 且 `w:drawing` tag delta 收敛。
- drawing/WordArt：通过 `GetAllDrawingObjects().ToJSON()` 补充可见文本，并为非 image 的顶层 drawing 生成最小 `w:drawing` + `w:pict` textbox 占位，VML fallback 会保留 extent 宽高 style；默认基线里 `会议纪要.docx` 的 `w:drawing/w:pict` delta 已收敛。复杂布局/几何/VML 样式仍不是完整结构保真。
- 基础脚注/尾注：正文 `footnoteRef/endnoteRef` 输出 `<w:footnoteReference>` / `<w:endnoteReference>`，run 上的 `footnotes/endnotes` body 生成 `footnotes.xml` / `endnotes.xml`，保留 note `type` 为 `w:type`，保留引用上的 `customMarkFollows`，并映射 `sectPr.footnotePr/endnotePr` 的编号格式、起始值、重启策略与位置。最小 DocumentBuilder fixture 已验证 `footnotes=1 endnotes=1`。

## 仍需继续补的范围

- 媒体结构保真：普通图片已接 DrawingML + media part，并保留 alt metadata、effectExtent、crop `srcRect`、preset geometry；shape/WordArt 已接最小 drawing/pict 占位与 VML extent style；复杂 anchor/wrap/effect、VML 几何与样式仍需继续补字段映射。
- 脚注/尾注复杂样例：基础正文引用、body part、note `w:type`、引用 `customMarkFollows`、分节 notePr 编号设置已接；还需要找更多真实业务样例验证多脚注、separator/continuationSeparator、跨栏等复杂形态。
- 批注完整性：当前从正文锚点与 `GetAllComments()` 合并 `id/author/text/date/solved`，生成 `comments.xml` 与 solved 用的 `commentsExtended.xml`；reply/thread 等扩展状态还需要继续找真实样例补相应扩展 part。
- numbering 长尾：`isLgl/legal/isLegal`、`lvlRestart/restart`、`lvlPicBulletId/picBulletId` 已映射；真正图片项目符号的 media part/relationship 与更复杂 restart 形态继续根据真机日志追加字段映射。
- 主题/字体复杂格式：`clrScheme/fontScheme` 已接入 `doc.theme`；`fmtScheme` 的 name、简单 solid fill/bgFill 与 lineStyle 已接，复杂 effect/fill 变体仍使用安全默认值，后续需要用更多真实主题样例补字段映射。

## 验证入口

- API 形态探针：`npm run probe:longtail -- [docx...]`；需要完整 JSON 时加 `-- --json` 或设置 `PROBE_JSON=1`。
- 本地包级回归：`node test/styles-roundtrip.mjs`。
- 真实 OOXML 基线：`npm run test:ooxml -- [docx...]`，对比原始 `word/document.xml` 与 `DocumentBuilder -> oojson -> ooxml`。默认会跑 3 份真实样例 + 1 份运行时生成的批注/脚注/尾注 docx，并断言 `comments=1 footnotes=1 endnotes=1`；DocumentBuilder 不是并发安全入口，多个 `test:ooxml` 不要并行跑。
- 宿主类型检查：`npm.cmd run typecheck -w coco-onlyoffice`。