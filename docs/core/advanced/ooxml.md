# OOXML

core 是 OOXML-native，但普通调用方不应直接拼接任意 XML。高层 API 应优先于低层 XML 操作。

## 输入输出

```ts
WordDocument.fromZip(bytes, opts);
WordDocument.fromFlatOpc(xml, opts);
doc.toZip();
doc.toFlatOpc();
```

## 受控 OOXML 暴露点

`tables.getVerbatim(tableId)` 返回表格 OOXML 子树和相关图片 payload，适合高保真搬表：

```ts
const source = await sourceDoc.tables.getVerbatim(tableId);
await targetDoc.plan().insertTableFrom({
  anchorParaId,
  tableXml: source.tableXml,
  images: source.images,
}).apply();
```

不要随意手改 `tableXml`，除非非常清楚 WordprocessingML 表格结构。

## XQuery Update

`applyXQueryUpdate` 是低层工具，不是普通文档操作首选。它可能绕过 high-level API 的锚点检查、索引维护和事务语义。

适合 core 内部实现、测试、受控迁移脚本。不适合直接暴露给 agent 或普通业务代码。

## Internal API

`WordDocument._getPackage()` 和 `_getIndices()` 是 internal 入口，只在 inline 模式下可用。业务集成不要依赖它们。

## 性能约束

大文档中避免无谓全局 XPath。新增能力应优先使用加载期索引、`getElementsByTagNameNS` 枚举结构、JS filter 和 dirty part 序列化。
