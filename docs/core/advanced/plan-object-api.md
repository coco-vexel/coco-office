# Live Object API

> Superseded note: current public API details live in `docs/core/domains/*.md`.
> This file is a design note and must not override the domain API docs.

本文定义 core 新的对象化操作模型。`batch` 不再作为 public API；业务和 agent 应直接使用 live object。

## 原则

- `doc.paragraphs.list()` / `doc.paragraphs.get()` 返回可操作对象，不返回裸 DTO。
- 对象方法是立即提交语义：`await paragraph.replaceWith(text)` 完成后，文档状态已经更新。
- 读取方法总是读取当前已提交文档状态。
- 底层仍可复用 `applyPlan`、事务回滚、索引重建、domain cache dirty 等机制，但这些不暴露给调用方。
- 多步复杂流程后续由 workflow/skill 编排；core 对象方法保持单个明确操作。

## Paragraph

```ts
const result = await doc.paragraphs.list({
  pagination: { mode: 'all' },
});

for (const paragraph of result.paragraphs) {
  const text = await paragraph.getText();
  if (text?.includes('错别字')) {
    await paragraph.findText('错别字', { scope: 'all' }).setColor('FF0000');
    await paragraph.findText('错别字', { scope: 'all' }).setBold(true);
  }
}
```

### `doc.paragraphs.list(opts?)`

返回：

```ts
interface ParagraphList {
  paragraphs: Paragraph[];
  pagination: {
    mode: 'all' | 'byCount' | 'byChar';
    pageIndex: number;
    totalPages: number;
    totalParagraphs: number;
    hasNext: boolean;
  };
}
```

### `doc.paragraphs.get(paraId)`

返回 `Paragraph | null`。

### `Paragraph`

| API | 说明 |
|---|---|
| `paragraph.paraId` / `paragraph.id` | 段落 ID |
| `paragraph.story` | 段落所在 story |
| `paragraph.getInfo()` | 读取当前 `ParagraphInfo` 快照 |
| `paragraph.getText()` | 读取当前段落文本 |
| `paragraph.replaceWith(text)` | 替换整段文本并立即提交 |
| `paragraph.setAlignment(value)` | 设置段落水平对齐并立即提交 |
| `paragraph.setOutlineLevel(level)` | 设置段落大纲级别并立即提交，取值 0-8 |
| `paragraph.setPageBreakBefore(value)` | 设置或移除段前分页并立即提交 |
| `paragraph.setIndent(value)` | 设置段落缩进并立即提交，单位 twips |
| `paragraph.setSpacing(value)` | 设置段落间距并立即提交，单位 twips |
| `paragraph.insertBefore(text)` | 在当前段落前插入新段落并立即提交 |
| `paragraph.insertAfter(text)` | 在当前段落后插入新段落并立即提交 |
| `paragraph.delete(opts?)` | 删除段落并立即提交 |
| `paragraph.addComment(text, opts?)` | 给整段添加批注；`opts.author` 可指定批注作者 |
| `paragraph.createBookmark(name)` | 给整段创建书签 |
| `paragraph.findText(search, opts?)` | 获取段内文字片段对象 |

写方法可传 `opts.trackChanges`；批注/修订类操作还可传 `opts.author` 和 `opts.date`：

```ts
await paragraph.replaceWith('新的段落内容', { trackChanges: true });
await paragraph.findText('旧词').replaceWith('新词', { trackChanges: true });
await paragraph.addComment('修改建议。', { author: 'halvee' });
```

### `TextRange`

| API | 说明 |
|---|---|
| `text.replaceWith(replacement)` | 替换段内匹配文字 |
| `text.delete()` | 删除段内匹配文字 |
| `text.setBold(value)` | 设置匹配文字加粗 |
| `text.setColor(color)` | 设置匹配文字颜色，使用 `RRGGBB` |
| `text.setItalic(value)` | 设置匹配文字斜体 |
| `text.setUnderline(value)` | 设置匹配文字单下划线 |
| `text.setHighlight(color)` | 设置匹配文字高亮色，例如 `yellow` |
| `text.setFontSize(points)` | 设置匹配文字字号，单位为 pt |
| `text.setFontFamily(name)` | 设置匹配文字字体 |
| `text.addComment(comment, opts?)` | 给匹配文字添加批注；`opts.author` 可指定批注作者 |
| `text.createBookmark(name)` | 给匹配文字创建书签 |

`paragraph.findText(search, { scope })` 的 `scope` 默认为 `'first'`，可传 `'all'`。

## Bookmark

```ts
const bookmark = await doc.bookmarks.get('review_target');
if (bookmark) {
  const text = await bookmark.getText();
  const paragraph = await bookmark.getParagraph();
  await bookmark.delete();
}
```

| API | 说明 |
|---|---|
| `doc.bookmarks.list(opts?)` | 返回可操作的 `Bookmark[]` |
| `doc.bookmarks.get(name)` | 返回 `Bookmark | null` |
| `bookmark.name` / `bookmark.id` | 书签名称 |
| `bookmark.rangeKind` | `paragraph` / `text` / `complex` |
| `bookmark.getInfo()` | 读取当前 `BookmarkInfo` |
| `bookmark.getText()` | 读取书签覆盖文本，复杂书签可能返回 `null` |
| `bookmark.getParagraph()` | 读取书签所在段落，复杂书签可能返回 `null` |
| `bookmark.delete()` | 删除该书签并立即提交 |

## Comment

```ts
const comment = await doc.comments.get(7);
if (comment) {
  const bodyText = await comment.getText();
  const anchorText = await comment.getAnchorText();
  const paragraph = await comment.getParagraph();
  await comment.delete();
}
```

| API | 说明 |
|---|---|
| `doc.comments.list()` | 返回可操作的 `Comment[]` |
| `doc.comments.get(id)` | 返回 `Comment | null` |
| `comment.id` | 批注 ID |
| `comment.text` | 批注正文快照 |
| `comment.anchorText` | 被批注正文片段快照 |
| `comment.rangeKind` | `paragraph` / `text` / `complex` |
| `comment.getInfo()` | 读取当前 `CommentInfo` |
| `comment.getText()` | 读取当前批注正文 |
| `comment.getAnchorText()` | 读取当前被批注正文片段，复杂锚点可能返回 `null` |
| `comment.getParagraph()` | 读取批注锚点所在段落，复杂锚点可能返回 `null` |
| `comment.delete()` | 删除批注正文和锚点并立即提交 |

## Revision

```ts
const revision = await doc.revisions.get('7', { scope: 'all' });
if (revision) {
  const text = await revision.getText();
  const paragraph = await revision.getParagraph();
}
```

| API | 说明 |
|---|---|
| `doc.revisions.list(opts?)` | 返回可读取的 `Revision[]` |
| `doc.revisions.get(revId, opts?)` | 返回 `Revision | null` |
| `doc.revisions.accept(opts?)` | 接受全部支持的修订 |
| `doc.revisions.reject(opts?)` | 拒绝全部支持的修订 |
| `revision.revId` / `revision.id` | 修订 ID |
| `revision.kind` | `insert` / `delete` |
| `revision.text` | 修订涉及文本快照 |
| `revision.getInfo()` | 读取当前 `RevisionInfo` |
| `revision.getText()` | 读取当前修订涉及文本 |
| `revision.getParagraph()` | 读取修订所在段落，无法定位到单段时返回 `null` |
| `revision.accept()` | 接受这一条修订 |
| `revision.reject()` | 拒绝这一条修订 |

单条修订使用 `revision.accept()` / `revision.reject()`。全部修订使用
`doc.revisions.accept()` / `doc.revisions.reject()`，或 WordPlan 的
`acceptRevisions({ scope: 'all' })` / `rejectRevisions({ scope: 'all' })`。
如果按 `doc.revisions.list()` 过滤后逐条处理，必须按 `revId` 去重。

## Table

当前表格已提供 live object：

```ts
const tables = await doc.tables.list({ includeCellTexts: true });
```

```ts
const table = await doc.tables.get(tableId);
await table?.cell(0, 0).replaceWith('new text');
await table?.cell(0, 0).insertParagraphAtStart('cell start');
await table?.cell(0, 0).insertParagraphAtEnd('cell end');
await table?.cell(0, 0).findText('target', { scope: 'all' }).setColor('FF0000');
await table?.cell(0, 0).setShading('FFF2CC');
await table?.cell(0, 0).setVerticalAlignment('center');
await table?.cell(0, 0).setWidth(2600);
await table?.row(2).delete();
await table?.column(1).setWidth(2400);
const blocks = await table?.cell(0, 0).getBlocks();
const cells = await table?.getCells();
// Table extends TableInfo; tableId/rowCount/cells/columnWidths remain readable.
```

`cell.getBlocks()` 保留单元格直接内容顺序，可返回段落、drawing、嵌套表格。
`cell.findText()` 搜索单元格直接段落，并返回支持 `replaceWith`、`delete`、
`setBold`、`setColor`、`setItalic`、`setUnderline`、`setHighlight`、`setFontSize`、`setFontFamily`、`addComment`、`createBookmark` 的文字范围。

## 和 `plan` 的关系

`doc.plan()` 是底层能力，可以继续保留给高级场景和内部实现。面向业务和小模型的主要入口是 live object。

```ts
const paragraph = await doc.paragraphs.get(paraId);
await paragraph?.replaceWith('新的段落内容');
```

而不是：

```ts
await doc.plan().replaceParagraph(paraId, '新的段落内容').apply();
```

## Cache 边界

live object 不直接持有 OOXML 节点引用。它持有稳定 ID，并在每次读取/写入时通过 core 内核访问当前文档状态。

domain cache 是 live object 的底层优化：

- 对象读取可以复用 domain cache。
- 对象写入成功后由内核统一 dirty 或 patch cache。
- 对象 API 不暴露 cache 细节。
