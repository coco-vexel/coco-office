# Host Integration

core 是 host-agnostic SDK。宿主层负责文档 I/O、选择区、文件系统、UI 和回写，core 负责 OOXML 文档操作。

## 集成流程

```ts
const input = await host.readDocument();
const doc = await WordDocument.fromZip(input, { docId });

await doc.plan().replaceParagraph(paraId, '新内容').apply();

const output = await doc.toZip();
await host.writeDocument(output);
await doc.dispose();
```

## 宿主边界

不要把这些逻辑写进 core：宿主专有 API 调用、本地文件路径读取、网络下载、UI 确认。

这些应由 host/app/agent 层完成。

## 选择区映射

core 当前不直接理解宿主 selection。宿主需要把 selection 转为 core 标识：`paraId`、`{ paraId, text }`、`tableId`、`drawingId` 或其他明确 API 参数。

## 回写

core 导出：

```ts
await doc.toZip();
await doc.toFlatOpc();
```

具体回写方式由宿主决定。

## 错误处理

宿主应记录 `WordError.code`、`message`、`docId`、`details`。

常见处理：

- `*_NOT_FOUND`: 刷新索引或提示目标已变化。
- `ANCHOR_CONFLICT`: 询问用户是否允许丢弃锚点。
- `UNSUPPORTED_OOXML_FEATURE`: 降级或交给人工处理。
- `WORKER_CREATION_FAILED`: 改 inline 或修正 worker 配置。
