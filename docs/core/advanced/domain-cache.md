# Domain Cache

本文定义 core 的懒加载 domain cache 机制。它是文档内核的基础能力，用于降低读取和后续对象 API 的重复扫描成本。

重要边界：domain cache 不决定对象 API 的形态。当前 public API 采用 live object；缓存层只负责“按需构建派生数据，并在写入后保持它与当前内存 openxml/package 一致”。

## 目标

- 初始化 `WordDocument` 时不全量扫描文档。
- 第一次读取某个功能域时才扫描并构建该 domain 的缓存。
- 后续读取同一 domain 时直接复用缓存。
- 写操作成功后，缓存要么被精确 patch，要么被标记 dirty。
- 不确定能否精确 patch 时，一律 dirty，下次读取时按需重建。
- 真实内容始终是当前内存里的 openxml/package，不是用户最初传入的文件。

## 非目标

- 不把缓存设计成新的 public object API。
- 不要求所有 domain 在初始化时建立完整索引。
- 不缓存宿主文件路径或原始用户文件。
- 不以性能为由牺牲一致性；缓存可慢，不可错。

## Domain 拆分

缓存按功能域独立存在：

```ts
interface WordDomainCache {
  paragraphs: DomainCache<ParagraphCache>;
  tables: DomainCache<TableCache>;
  bookmarks: DomainCache<BookmarkCache>;
  comments: DomainCache<CommentCache>;
  drawings: DomainCache<DrawingCache>;
  summary: DomainCache<DocumentSummaryCache>;
}
```

每个 domain cache 有独立状态：

```ts
type CacheStatus = 'empty' | 'ready' | 'dirty';

interface DomainCache<T> {
  status: CacheStatus;
  builtAtRevision: number;
  dirtyReason?: string;
  value?: T;
}
```

读取流程：

```ts
getOrBuild(domain):
  if domain.status !== 'ready':
    domain.value = scanCurrentOpenXmlForDomain()
    domain.status = 'ready'
    domain.builtAtRevision = contentRevision
  return domain.value
```

## 真实内容和版本

core 内部真实内容是当前内存中的 OOXML package / DOM。

```ts
class Kernel {
  private contentRevision: number;
  private cache: WordDomainCache;
}
```

每次写操作成功修改 openxml/package 后：

- `contentRevision += 1`
- 根据已执行 op 更新或失效相关 cache

写操作失败并回滚时：

- openxml/package 回到写前状态
- cache 不变

外部宿主绕过 core 修改内容时，宿主必须通知 core：

```ts
kernel.invalidateAll('host content changed');
```

如果宿主无法保证通知，core 可以在宿主集成层选择读前保守 invalidation，性能退化但正确性不退化。

## 缓存粒度

第一阶段建议只接入最常用且成本敏感的 domain：

- `paragraphs`
- `tables`
- `summary`

第二阶段再接入：

- `bookmarks`
- `comments`
- `drawings`

示例缓存结构：

```ts
interface ParagraphCache {
  byId: Map<string, ParagraphInfo>;
  order: string[];
  totalParagraphs: number;
  totalChars: number;
}

interface TableCache {
  byId: Map<string, TableInfo>;
  order: string[];
}

interface DocumentSummaryCache {
  paragraphCount: number;
  tableCount: number;
  charCount: number;
  previewText: string;
}
```

## 当前实现状态

当前已落地第一步基础设施：

- `Kernel` 持有 `WordDomainCache`。
- 现有 `load()` 阶段的结构索引 `buildIndices` 暂时保留；本阶段只把段落 DTO/list 这类 domain 派生数据改为懒加载缓存。
- `paragraphs.list({ pagination: { mode: 'all' } })` 在无 `scope/filter/includeFormat/structuralWhitespace/includeTableInfo` 等扩展参数时按需构建并复用缓存。
- `paragraphs.get(paraId)` 按需缓存单段结果。
- `tables.list()` 和 `tables.list({ includeCellTexts: true })` 在默认 `scope='body'`、`includeNested=false` 场景下按需构建并复用缓存；其它表格查询 profile 暂时回退到纯扫描。
- `tables.get(tableId)` 按需缓存基础 `TableInfo`，不会因为此前缓存过 `includeCellTexts` 列表而额外返回 `cells/cellGrid`。
- 写操作成功后推进 `revision`；影响段落文本或结构的 op 将 `paragraphs` 标记为 dirty，下次读取从当前内存 OOXML 重新扫描。
- 影响表格结构或表格内容的 op 会将 `tables` 标记为 dirty，下次读取从当前内存 OOXML 重新扫描。
- `formatTextInParagraph` 这类只改样式、不影响当前段落文本 DTO 的 op 不使段落文本缓存 dirty。
- 写操作失败并回滚时缓存不变。
- 提供 inline-only 的 `_debugCacheSnapshot()` / `_debugInvalidateCache()` 供 core 单元测试使用。

当前实现故意没有做精确 patch。原因是第一步优先保证一致性和边界清晰；后续可以在单元测试覆盖充分后，把 `replaceParagraph`、`replaceTextInParagraph` 等简单 op 从 dirty 优化为 patch。

## 写后更新策略

每个 op 对 cache 的影响分为三种：

```ts
type CacheEffect =
  | { kind: 'none'; domain?: CacheDomain }
  | { kind: 'patch'; domain: CacheDomain; patch: unknown }
  | { kind: 'invalidate'; domain: CacheDomain; reason: string };
```

原则：

- 简单文本 op 可以 patch。
- 复杂结构 op 优先 invalidate。
- 只改样式且不影响当前缓存内容时可以 none。
- 任何不确定的影响都必须 invalidate。

第一阶段 patch 范围建议非常小：

| op | cache strategy |
|---|---|
| `replaceTextInParagraph` | patch `paragraphs`; dirty `summary` |
| `formatTextInParagraph` | none |
| `replaceParagraph` | patch `paragraphs`; dirty `summary` |

其他 op 第一阶段可以全部按影响域 dirty。

## Op 影响矩阵

| op | paragraphs | tables | bookmarks | comments | drawings | summary |
|---|---|---|---|---|---|---|
| `replaceTextInParagraph` | patch | none | none | none | none | dirty |
| `formatTextInParagraph` | none | none | none | none | none | none |
| `replaceParagraph` | patch | none | dirty if anchors may change | dirty if anchors may change | dirty if drawings may change | dirty |
| `insertParagraphBefore/After` | dirty | none | none | none | none | dirty |
| `insertParagraphAtBodyStart` | dirty | none | none | none | none | dirty |
| `deleteParagraph` | dirty | dirty if structure may change | dirty | dirty | dirty | dirty |
| `createBookmark` | none | none | dirty | none | none | none |
| `deleteBookmark` | none | none | dirty | none | none | none |
| `addComment` | none | none | none | dirty | none | none |
| `deleteComment` | none | none | none | dirty | none | none |
| `addTableComment` | none | none | none | dirty | none | none |
| `deleteTable` | dirty | dirty | dirty | dirty | dirty | dirty |
| `replaceTableCellText` | dirty | dirty | dirty if anchors may change | dirty if anchors may change | none | dirty |
| `replaceTableCellWithSdt` | dirty | dirty | dirty if anchors may change | dirty if anchors may change | none | dirty |
| `insertTableRow` | dirty | dirty | dirty | dirty | none | dirty |
| `deleteTableRow` | dirty | dirty | dirty | dirty | none | dirty |
| `cloneTableAfter` | dirty | dirty | none | none | none | dirty |
| `setTableColumnWidths` | none | dirty | none | none | none | none |
| `setTableRowHeights` | none | dirty | none | none | none | none |
| `replaceDrawingWithPlaceholder` | none | none | none | none | dirty | dirty |
| `deleteDrawing` | dirty if containing run/paragraph changes | dirty if inside table | none | none | dirty | dirty |
| `replaceShapeInnerText` | dirty | none | none | none | dirty | dirty |
| `insertImageAfter` | dirty | none | none | none | dirty | dirty |
| `insertTableFrom` | dirty | dirty | none | none | dirty if imported images exist | dirty |
| `setUpdateFields` | none | none | none | none | none | none |

这个矩阵可以先保守实现。后续发现某个 op 的 patch 足够安全，再从 dirty 优化为 patch。

## Domain 依赖

`summary` 是派生摘要，依赖多个 domain。任何影响文本、结构数量、表格数量、drawing 数量的操作都应 dirty summary。

```ts
const CACHE_DEPENDENCIES = {
  paragraphs: ['summary'],
  tables: ['summary'],
  drawings: ['summary'],
  bookmarks: [],
  comments: [],
};
```

依赖传播只传播 dirty，不传播 patch。patch 后如果依赖域不能精确同步，则依赖域 dirty。

## 和对象 API 的关系

domain cache 是对象 API 的底座。当前 public API 采用 live object：

```ts
paragraph.getText()
table.cell(row, col).getText()
```

这些读取应落到对应 domain cache。

live object 写入时：

```ts
paragraph.replaceWith(text)
table.cell(row, col).replaceWith(text)
table.cell(row, col).insertParagraphAtEnd(text)
```

这些写入应先修改真实 openxml/package，成功后通过相同的 cache update policy patch 或 dirty cache。

因此缓存层只提供：

- 按需读取派生数据
- 写后同步/失效
- debug/testing hooks

它不暴露 OOXML 节点引用，也不要求 live object 长期持有真实 DOM。live object 持有稳定 ID，读取和写入时再进入 kernel。

## Debug 和测试接口

建议提供仅测试/调试使用的能力：

```ts
doc._debugCacheSnapshot()
doc._debugInvalidateCache(domainOrAll, reason)
```

不要把 debug API 放进普通 agent 文档，也不要让 CodePlan 使用。

## 一致性测试

每个接入缓存的写 op 都应有一致性测试：

```ts
await doc.paragraphs.list(); // build cache
await writeOperation();
const cached = await doc.paragraphs.list();

doc._debugInvalidateCache('paragraphs', 'test rescan');
const rescanned = await doc.paragraphs.list();

expect(cached).toEqual(rescanned);
```

还需要测试：

- domain 懒加载：读取 paragraphs 不应构建 tables cache。
- 写失败：openxml 回滚后 cache 不变。
- 外部 invalidation：dirty 后下一次读取会重建。
- 多 op apply：一次 apply 后按 op 集合合并更新 cache。

## 落地阶段

第一阶段：

- 新增 `WordDomainCache` 基础设施。
- 接入 `paragraphs.list/get`。
- 段落 live object 读取复用 paragraph cache。
- 影响段落文本或结构的 op dirty paragraphs。
- 补 cache consistency tests。

第二阶段：

- 接入 `tables.list/get`。
- 表格相关 op 后 dirty tables / paragraphs。
- 补 table cache consistency tests。

第三阶段：

- 继续把 bookmark / comment / drawing 改为 live object。
- cache 作为这些对象的底层读优化，不反向牵引 public API。

第四阶段：

- 接入 bookmarks / comments / drawings。
- 完善外部宿主 invalidation contract。
