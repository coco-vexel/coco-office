# Errors

core 使用 `WordError` 表达可预期错误。

```ts
class WordError extends Error {
  readonly code: WordErrorCode;
  readonly docId: string;
  readonly details?: Record<string, unknown>;
}
```

## 处理模式

```ts
try {
  await doc.plan().replaceParagraph(paraId, '新内容').apply();
} catch (err) {
  if (err instanceof WordError) {
    console.error(err.code, err.docId, err.details);
  }
}
```

## 常见类别

| 类别 | 示例 |
| --- | --- |
| 加载/格式 | `PACKAGE_CORRUPT`、`PACKAGE_UNSUPPORTED_FORMAT` |
| 定位失败 | `PARA_NOT_FOUND`、`TABLE_NOT_FOUND`、`DRAWING_NOT_FOUND` |
| 参数/锚点 | `INVALID_ARGUMENT`、`ANCHOR_CONFLICT` |
| 能力缺口 | `UNSUPPORTED_OOXML_FEATURE` |
| worker | `WORKER_CREATION_FAILED`、`WORKER_RELOADED` |
| plan 状态 | `PLAN_ALREADY_APPLIED` |

完整错误码以 `packages/core/src/word/word-error.ts` 为准。

## 建议处理

- `*_NOT_FOUND`: 目标可能已变化，重新读取上下文。
- `ANCHOR_CONFLICT`: 默认不要丢弃锚点，先让调用方或用户确认。
- `INVALID_ARGUMENT`: 调用方 bug，修正参数。
- `UNSUPPORTED_OOXML_FEATURE`: 降级或交给人工。
- `WORKER_CREATION_FAILED`: 检查 worker factory/bundle，或改用 inline。
