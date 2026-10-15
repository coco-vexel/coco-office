# Getting Started

本页展示最小可运行路径：加载文档、读取段落、修改段落、导出。

```ts
import { WordDocument } from '@coco-vexel/coco-office/word';

const doc = await WordDocument.fromZip(bytes, {
  docId: 'file:demo.docx',
});

const { paragraphs } = await doc.paragraphs.list({
  textFormat: 'preview',
});

const first = paragraphs[0];
if (first) {
  await first.findText('旧词').replaceWith('新词');
  await first.insertAfter('新增段落', { style: 'inherit', runFormat: 'inherit' });
}

const outBytes = await doc.toZip();
await doc.dispose();
```

也可以从 Flat OPC XML 加载：

```ts
const doc = await WordDocument.fromFlatOpc(xml, {
  docId: 'wps:active-doc',
});
```

也可以直接创建一个内存里的空白文档，再写入内容并导出 docx：

```ts
const doc = await WordDocument.createBlank({
  docId: 'new-doc',
});

const { paragraphs } = await doc.paragraphs.list({
  pagination: { mode: 'all' },
});

await paragraphs[0]?.replaceWith('第一段内容');
await doc.appendParagraph('第二段内容');
const outBytes = await doc.toZip();
await doc.dispose();
```

`createBlank` 只创建 core 内存文档，不负责在 WPS、Microsoft Word 或 OnlyOffice 中打开新窗口；宿主打开文档仍由 host 层负责。

如果需要插入图片或简单表格，先拿到一个段落作为锚点：

```ts
await paragraphs[0]?.insertImageAfter({
  bytes: imageBytes,
  contentType: 'image/png',
  widthEmu: 914400,
  heightEmu: 914400,
});

await paragraphs[0]?.insertTableAfter({
  rows: [
    ['项目', '金额'],
    ['预算', '1000'],
  ],
});
```

`docId` 必填。core 不解析它，但会用它做错误定位、快照命名空间和诊断日志。
