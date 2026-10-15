# Cookbook

## 加载 docx 并列出前 25 段

```ts
import { WordDocument } from '@coco-vexel/coco-office/word';

const doc = await WordDocument.fromZip(bytes, { docId: 'file:demo.docx' });

const { paragraphs, pagination } = await doc.paragraphs.list({
  textFormat: 'preview',
});

console.log(paragraphs, pagination);
```

## 新建空白文档并写入第一段

```ts
import { WordDocument } from '@coco-vexel/coco-office/word';

const doc = await WordDocument.createBlank({ docId: 'new-doc' });
const { paragraphs } = await doc.paragraphs.list({
  pagination: { mode: 'all' },
});

await paragraphs[0]?.replaceWith('第一段内容');

const out = await doc.toZip();
await doc.dispose();
```

`createBlank` 创建的是 core 内存文档。把导出的 docx 打开到具体编辑器中，应由宿主集成层处理。

## 全文查找包含某个词的段落

```ts
const { paragraphs } = await doc.paragraphs.list({
  scope: 'all',
  filter: { contains: '合同编号' },
  pagination: { mode: 'all' },
});
```

`contains` 是 O(N) 子串匹配，大文档中应配合 scope、分页或上层索引使用。

## 替换段内第一个匹配文本

```ts
const paragraph = await doc.paragraphs.get(paraId);
await paragraph?.findText('甲方', { scope: 'first' }).replaceWith('委托方');
```

## 保留图片替换整段文字

```ts
const paragraph = await doc.paragraphs.get(paraId);
await paragraph?.replaceWith('新的段落文字', { keepInlineGraphics: true });
```

## 插入一个继承样式的段落

```ts
const anchor = await doc.paragraphs.get(anchorParaId);
const result = await anchor?.insertAfter('补充说明', {
  style: 'inherit',
  runFormat: 'inherit',
});

const newParaId = result?.insertedParagraphs[0]?.newParaId;
```

## 加批注

```ts
const paragraph = await doc.paragraphs.get(paraId);
const result = await paragraph?.findText('付款期限', { scope: 'first' }).addComment(
  '请确认是否与合同首页一致',
  { author: 'Reviewer' },
);

console.log(result?.createdComments);
```

## 列出表格并修改单元格文本

```ts
const tables = await doc.tables.list();
const table = tables[0];

await table?.cell(0, 1).replaceWith('新值');
```

## 设置表格行高和列宽

```ts
const table = await doc.tables.get(tableId);
await table?.setColumnWidths([2000, 3000, 3000]);
await table?.setRowHeights([{ val: 420, hRule: 'atLeast' }, null, null]);
```

## 操作单元格内部文本

```ts
const table = await doc.tables.get(tableId);
const cell = table?.cell(0, 0);
await cell?.insertParagraphAtEnd('补充说明');
await cell?.findText('风险').setColor('FF0000');
```

## 追加正文段落

```ts
await doc.appendParagraph('补充说明');
```

## 在段落后插入图片

```ts
const anchor = await doc.paragraphs.get(anchorParaId);
await anchor?.insertImageAfter({
  bytes: imageBytes,
  contentType: 'image/png',
  widthEmu: 914400,
  heightEmu: 914400,
});
```

## 在段落后插入简单表格

```ts
const anchor = await doc.paragraphs.get(anchorParaId);
await anchor?.insertTableAfter({
  rows: [
    ['项目', '负责人'],
    ['预算', '财务'],
  ],
  columnWidths: [2200, 2200],
});
```

## 删除图片

```ts
const pictures = await doc.drawings.listPictures();
await pictures[0]?.delete();
```

## 低层能力

`doc.plan()` 只作为 advanced/low-level API 记录在 [advanced/plan.md](./advanced/plan.md)。
普通文档操作和 codeAgent 生成代码应优先使用 live object APIs。

## 导出并释放

```ts
const out = await doc.toZip();
await doc.dispose();
```
