# Runtime

core 支持 inline 和 worker 两种运行模式。public API 都是 async，因此调用方不需要因为运行模式改变调用方式。

## 模式

```ts
type RuntimeMode = 'auto' | 'worker' | 'inline';
```

| 模式 | 说明 |
| --- | --- |
| `inline` | kernel 跑在当前线程 |
| `worker` | kernel 跑在 worker，主线程通过 RPC 调用 |
| `auto` | 按环境和配置选择 |

## 配置默认模式

```ts
WordDocument.setDefaultMode('auto');

const doc = await WordDocument.fromZip(bytes, {
  docId,
  mode: 'inline',
});
```

## 诊断实际模式

```ts
WordDocument.detectMode();
WordDocument.detectModeDecision();
```

加载时记录：

```ts
await WordDocument.fromZip(bytes, {
  docId,
  onModeResolved(decision, docId) {
    console.log(docId, decision.mode, decision.source);
  },
});
```

## Worker 配置

浏览器或宿主 WebView 中，core 不会自动知道 worker bundle 地址。需要 worker 时由宿主提供：

```ts
WordDocument.setWorkerFactory(() => new Worker(...));
```

Node 环境可以配置 bundle URL：

```ts
WordDocument.setWorkerBundleUrl(new URL('./kernel-worker.mjs', import.meta.url));
```

## 选择建议

UI 宿主优先使用 worker，避免大文档解析阻塞主线程。CLI、服务端、测试可以使用 inline。不确定时使用 `auto`，并通过 `onModeResolved` 记录实际结果。
