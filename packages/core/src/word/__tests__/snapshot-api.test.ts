import { afterEach, describe, expect, test, vi } from 'vitest';
import JSZip from 'jszip';
import { InMemorySnapshotStorage } from '../../shared/index.js';
import { WordDocument } from '../word-document.js';
import { W14_NS, W_NS } from '../namespaces.js';

const DOC_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docXml(paragraphs: string[]): string {
  const body = paragraphs.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>${body}</w:body>
</w:document>`;
}

async function makeZip(paragraphs: string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml(paragraphs));
  return zip.generateAsync({ type: 'uint8array' });
}

function collectBodyTexts(doc: WordDocument): string[] {
  // 直接读 _getPackage(inline 模式)看 <w:t> 文本,验证 rollback / take 语义
  const pkg = doc._getPackage();
  const docPart = pkg.parts.get('word/document.xml')!;
  const ts = docPart.xmlDom!.getElementsByTagNameNS(W_NS, 't');
  const out: string[] = [];
  for (let i = 0; i < ts.length; i++) out.push(ts[i]?.textContent ?? '');
  return out;
}

afterEach(() => {
  WordDocument._resetDefaults();
});

describe('WordDocument.snapshots 挂载', () => {
  test('装载后 snapshots 就位,类型满足 SnapshotApi', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['hello']), { docId: 'd' });
    expect(doc.snapshots).toBeDefined();
    expect(typeof doc.snapshots.take).toBe('function');
    expect(typeof doc.snapshots.restore).toBe('function');
    expect(typeof doc.snapshots.rollback).toBe('function');
    expect(typeof doc.snapshots.list).toBe('function');
    expect(typeof doc.snapshots.delete).toBe('function');
    expect(typeof doc.snapshots.clear).toBe('function');
    expect(typeof doc.snapshots.recordChange).toBe('function');
    expect(typeof doc.snapshots.changes).toBe('function');
    await doc.dispose();
  });

  test('未传 snapshotStorage 时默认用 InMemorySnapshotStorage', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    expect(doc.snapshotStorage).toBeInstanceOf(InMemorySnapshotStorage);
    await doc.dispose();
  });

  test('显式传 snapshotStorage 优先于默认', async () => {
    const custom = new InMemorySnapshotStorage();
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd',
      snapshotStorage: custom,
    });
    expect(doc.snapshotStorage).toBe(custom);
    await doc.dispose();
  });

  test('setDefaultSnapshotStorage 作为全局默认(LoadOptions 未显式传时生效)', async () => {
    const custom = new InMemorySnapshotStorage();
    WordDocument.setDefaultSnapshotStorage(custom);
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    expect(doc.snapshotStorage).toBe(custom);
    await doc.dispose();
  });
});

describe('snapshots.take / rollback', () => {
  test('take 返回自增 id(snap-0 / snap-1 / ...)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    const id0 = await doc.snapshots.take();
    const id1 = await doc.snapshots.take();
    expect(id0).toBe('snap-0');
    expect(id1).toBe('snap-1');
    await doc.dispose();
  });

  test('take 后 list 能看到 meta(不含 parts 字节)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    await doc.snapshots.take('初始状态');
    await doc.snapshots.take('第二次');
    const infos = await doc.snapshots.list();
    expect(infos.map((i) => i.id)).toEqual(['snap-0', 'snap-1']);
    expect(infos[0]?.reason).toBe('初始状态');
    expect(infos[1]?.reason).toBe('第二次');
    await doc.dispose();
  });

  test('rollback:状态恢复到 take 时刻', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['a', 'b']), { docId: 'd' });
    expect(collectBodyTexts(doc)).toEqual(['a', 'b']);

    const snapId = await doc.snapshots.take('before change');

    // 模拟一次修改:直接改 DOM 文本
    const pkg = doc._getPackage();
    const t0 = pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 't')[0]!;
    t0.textContent = 'MODIFIED';
    expect(collectBodyTexts(doc)).toEqual(['MODIFIED', 'b']);

    await doc.snapshots.rollback(snapId);
    expect(collectBodyTexts(doc)).toEqual(['a', 'b']);
    await doc.dispose();
  });

  test('rollback 消费语义:回滚后该快照被删除(list 不再含它)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    const snap0 = await doc.snapshots.take('s0');
    const snap1 = await doc.snapshots.take('s1');
    await doc.snapshots.rollback(snap0);
    const remaining = (await doc.snapshots.list()).map((i) => i.id);
    expect(remaining).toContain(snap1);
    expect(remaining).not.toContain(snap0);
    await doc.dispose();
  });

  test('restore 只恢复状态,不消费快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['before']), { docId: 'd' });
    const id = await doc.snapshots.take('before change');
    const t0 = doc
      ._getPackage()
      .parts.get('word/document.xml')!
      .xmlDom!.getElementsByTagNameNS(W_NS, 't')[0]!;
    t0.textContent = 'after';

    await doc.snapshots.restore(id);

    expect(collectBodyTexts(doc)).toEqual(['before']);
    expect((await doc.snapshots.list()).map((item) => item.id)).toContain(id);
    await doc.dispose();
  });

  test('重新装载同 docId + 共享 adapter 时从历史最大 id 继续分配', async () => {
    const shared = new InMemorySnapshotStorage();
    const first = await WordDocument.fromZip(await makeZip(['first']), {
      docId: 'shared-doc',
      snapshotStorage: shared,
    });
    expect(await first.snapshots.take()).toBe('snap-0');
    expect(await first.snapshots.take()).toBe('snap-1');
    await first.dispose();

    const reopened = await WordDocument.fromZip(await makeZip(['reopened']), {
      docId: 'shared-doc',
      snapshotStorage: shared,
    });
    expect(await reopened.snapshots.take()).toBe('snap-2');
    expect((await reopened.snapshots.list()).map((item) => item.id)).toEqual([
      'snap-0',
      'snap-1',
      'snap-2',
    ]);
    await reopened.dispose();
  });

  test('首次装载无法读取 adapter 历史时 fail-fast，避免从 snap-0 覆盖', async () => {
    class ListFailingStorage extends InMemorySnapshotStorage {
      override async list(): Promise<never> {
        throw new Error('list unavailable');
      }
    }
    await expect(
      WordDocument.fromZip(await makeZip(['x']), {
        docId: 'list-failure',
        snapshotStorage: new ListFailingStorage(),
      }),
    ).rejects.toThrow('list unavailable');
  });

  test('rollback 不存在的 id 抛错', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    await expect(doc.snapshots.rollback('snap-999')).rejects.toThrow();
    await doc.dispose();
  });

  test('rollback 后可继续 take 新快照(id 继续自增)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    const s0 = await doc.snapshots.take();
    await doc.snapshots.rollback(s0);
    const s1 = await doc.snapshots.take();
    expect(s1).toBe('snap-1');
    await doc.dispose();
  });

  test('多 doc 使用同一 adapter:docId 命名空间隔离', async () => {
    const shared = new InMemorySnapshotStorage();
    const docA = await WordDocument.fromZip(await makeZip(['a']), {
      docId: 'doc-A',
      snapshotStorage: shared,
    });
    const docB = await WordDocument.fromZip(await makeZip(['b']), {
      docId: 'doc-B',
      snapshotStorage: shared,
    });
    await docA.snapshots.take();
    await docB.snapshots.take();
    await docB.snapshots.take();
    expect((await docA.snapshots.list()).length).toBe(1);
    expect((await docB.snapshots.list()).length).toBe(2);
    await docA.dispose();
    await docB.dispose();
  });

  test('同 dedupeKey + 相同整篇内容复用已有 id，不新增快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['same']), { docId: 'dedupe-same' });
    const first = await doc.snapshots.take('first', undefined, { dedupeKey: 'user:s1' });
    doc.snapshots.recordChange({ toolName: 'net-zero' });
    const second = await doc.snapshots.take('second', undefined, { dedupeKey: 'user:s1' });
    expect(second).toBe(first);
    expect(doc.snapshots.changes()).toEqual([]);
    const list = await doc.snapshots.list();
    expect(list).toHaveLength(1);
    expect(list[0]?.contentHash).toMatch(/^[0-9a-f]{64}$/);
    expect(list[0]?.dedupeKey).toBe('user:s1');
    await doc.dispose();
  });

  test('相同内容但 dedupeKey 不同仍保留两个逻辑快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['same']), { docId: 'dedupe-owner' });
    const internal = await doc.snapshots.take('session-start', undefined, {
      dedupeKey: 'agent:s1',
    });
    const user = await doc.snapshots.take('turn-start', undefined, { dedupeKey: 'user:s1' });
    expect(user).not.toBe(internal);
    expect(await doc.snapshots.list()).toHaveLength(2);
    await doc.dispose();
  });

  test('同 dedupeKey 但文档内容变化时创建新快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['before']), { docId: 'dedupe-change' });
    const first = await doc.snapshots.take('before', undefined, { dedupeKey: 'user:s1' });
    const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
    const [paragraph] = paragraphs;
    await doc.plan().replaceParagraph(paragraph!.paraId, 'after').apply();
    const second = await doc.snapshots.take('after', undefined, { dedupeKey: 'user:s1' });
    expect(second).not.toBe(first);
    const list = await doc.snapshots.list();
    expect(list).toHaveLength(2);
    expect(list[0]?.contentHash).not.toBe(list[1]?.contentHash);
    await doc.dispose();
  });

  test('未提供 dedupeKey 时保持原语义，相同内容仍创建新快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['same']), { docId: 'dedupe-off' });
    expect(await doc.snapshots.take()).toBe('snap-0');
    expect(await doc.snapshots.take()).toBe('snap-1');
    await doc.dispose();
  });

  test('save 已成功但 FIFO delete 失败时 take 仍返回 id 并保留快照', async () => {
    class DeleteFailingStorage extends InMemorySnapshotStorage {
      failDelete = false;
      override async delete(docId: string, id: string): Promise<void> {
        if (this.failDelete) throw new Error('delete failed');
        return super.delete(docId, id);
      }
    }
    const storage = new DeleteFailingStorage();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'eviction-failure',
      snapshotStorage: storage,
      maxSnapshots: 1,
    });
    expect(await doc.snapshots.take()).toBe('snap-0');
    storage.failDelete = true;

    await expect(doc.snapshots.take()).resolves.toBe('snap-1');
    expect((await doc.snapshots.list()).map((item) => item.id)).toEqual(['snap-0', 'snap-1']);
    expect(warn).toHaveBeenCalledWith(
      expect.stringContaining('eviction failed after saving'),
      expect.any(Error),
    );
    warn.mockRestore();
    await doc.dispose();
  });
});

describe('snapshots — dirty 优化(引用共享)', () => {
  test('take:未改过的 part 的 bytes 与 _originalBytes 是同一 Uint8Array 引用(零 serialize)', async () => {
    const adapter = new InMemorySnapshotStorage();
    const doc = await WordDocument.fromZip(await makeZip(['hello']), {
      docId: 'd',
      snapshotStorage: adapter,
    });
    const id = await doc.snapshots.take();
    const payload = await adapter.load('d', id);

    // `_rels/.rels` 装载时不会被 ensureParaIds 触及,保持 _dirty=false
    // 此 part 的 snapshot bytes 应是 _originalBytes 的同一 Uint8Array 引用(零 serialize)
    const relsPart = doc._getPackage().parts.get('_rels/.rels');
    const relsInSnapshot = payload.parts.find((p) => p.path === '_rels/.rels')?.bytes;
    expect(relsPart?._originalBytes).toBeDefined();
    expect(relsInSnapshot).toBe(relsPart?._originalBytes);

    await doc.dispose();
  });

  test('take:ensureParaIds 改过的 document.xml → _dirty=true,bytes 是新 serialize 的结果', async () => {
    const adapter = new InMemorySnapshotStorage();
    const doc = await WordDocument.fromZip(await makeZip(['hello']), {
      docId: 'd',
      snapshotStorage: adapter,
    });

    // 装载时 ensureParaIds 给没有 paraId 的段落补了 w14:paraId → 标脏 document.xml
    const docPart = doc._getPackage().parts.get('word/document.xml')!;
    expect(docPart._dirty).toBe(true);

    const id = await doc.snapshots.take();
    const payload = await adapter.load('d', id);
    const bytes = payload.parts.find((p) => p.path === 'word/document.xml')?.bytes;

    expect(bytes).not.toBe(docPart._originalBytes); // 新字节
    expect(bytes).toBeInstanceOf(Uint8Array);
    // 新字节里应含 w14:paraId 属性(证明是从 DOM serialize 的,不是原始字节)
    const text = new TextDecoder().decode(bytes);
    expect(text).toMatch(/paraId/);
    await doc.dispose();
  });
});

describe('snapshots.recordChange / changes', () => {
  test('recordChange 同步累加,changes 返回快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    doc.snapshots.recordChange({ toolName: 'replace_paragraphs', summary: { paraId: 'A' } });
    doc.snapshots.recordChange({ toolName: 'add_comment' });
    const arr = doc.snapshots.changes();
    expect(arr).toHaveLength(2);
    expect(arr[0]?.toolName).toBe('replace_paragraphs');
    expect(arr[0]?.summary).toEqual({ paraId: 'A' });
    expect(arr[1]?.toolName).toBe('add_comment');
    expect(arr[1]?.summary).toBeUndefined();
    // 每条带 timestamp
    expect(typeof arr[0]?.timestamp).toBe('number');
    await doc.dispose();
  });

  test('take 时吸入当前 changelog,之后 changes() 为空', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    doc.snapshots.recordChange({ toolName: 't1' });
    doc.snapshots.recordChange({ toolName: 't2' });
    expect(doc.snapshots.changes()).toHaveLength(2);

    const id = await doc.snapshots.take();
    expect(doc.snapshots.changes()).toHaveLength(0);

    const infos = await doc.snapshots.list();
    const info = infos.find((i) => i.id === id)!;
    expect(info.changesCaptured.map((c) => c.toolName)).toEqual(['t1', 't2']);
    await doc.dispose();
  });

  test('changes 返回 shallow copy,修改不影响内部', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    doc.snapshots.recordChange({ toolName: 'x' });
    const arr = doc.snapshots.changes();
    arr.length = 0;
    expect(doc.snapshots.changes()).toHaveLength(1);
    await doc.dispose();
  });

  test('不调 recordChange 时 take 的 changesCaptured 为空数组', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    const id = await doc.snapshots.take();
    const info = (await doc.snapshots.list()).find((i) => i.id === id)!;
    expect(info.changesCaptured).toEqual([]);
    await doc.dispose();
  });
});

describe('snapshots FIFO 上限(#22 §3.2)', () => {
  test('默认 10 张:连续 take 11 次 → list 只剩 10 条,snap-0 被淘汰', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd-fifo' });
    for (let i = 0; i < 11; i++) {
      await doc.snapshots.take();
    }
    const ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toHaveLength(10);
    expect(ids).not.toContain('snap-0');
    expect(ids).toContain('snap-1');
    expect(ids).toContain('snap-10');
    await doc.dispose();
  });

  test('默认 10:再 take 一次 → snap-1 被淘汰(顺次)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd-fifo2' });
    for (let i = 0; i < 12; i++) {
      await doc.snapshots.take();
    }
    const ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toHaveLength(10);
    expect(ids).not.toContain('snap-0');
    expect(ids).not.toContain('snap-1');
    expect(ids).toContain('snap-2');
    expect(ids).toContain('snap-11');
    await doc.dispose();
  });

  test('自定义 maxSnapshots=3:超出后只保留最新 3 条', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-m3',
      maxSnapshots: 3,
    });
    for (let i = 0; i < 5; i++) {
      await doc.snapshots.take();
    }
    const ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toEqual(['snap-2', 'snap-3', 'snap-4']);
    await doc.dispose();
  });

  test('maxSnapshots=1:每次 take 都会淘汰上一次的', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-m1',
      maxSnapshots: 1,
    });
    await doc.snapshots.take();
    expect((await doc.snapshots.list()).map((i) => i.id)).toEqual(['snap-0']);
    await doc.snapshots.take();
    expect((await doc.snapshots.list()).map((i) => i.id)).toEqual(['snap-1']);
    await doc.snapshots.take();
    expect((await doc.snapshots.list()).map((i) => i.id)).toEqual(['snap-2']);
    await doc.dispose();
  });

  test('maxSnapshots 非法(0 / 负数 / 非整数)→ 装载期抛 WordError', async () => {
    const buf = await makeZip(['x']);
    await expect(WordDocument.fromZip(buf, { docId: 'd-bad', maxSnapshots: 0 })).rejects.toThrow(
      /maxSnapshots/,
    );
    await expect(WordDocument.fromZip(buf, { docId: 'd-bad', maxSnapshots: -1 })).rejects.toThrow(
      /maxSnapshots/,
    );
    await expect(WordDocument.fromZip(buf, { docId: 'd-bad', maxSnapshots: 1.5 })).rejects.toThrow(
      /maxSnapshots/,
    );
  });
});

describe('snapshots.rollback onProgress(#22 §6)', () => {
  test('inline 模式:rollback onProgress 收到 parse / build-indices / done 三个 phase', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['a', 'b']), { docId: 'd-prog' });
    const id = await doc.snapshots.take();

    const calls: Array<{ phase: string; pct?: number }> = [];
    await doc.snapshots.rollback(id, {
      onProgress: (phase, pct) => {
        calls.push({ phase, pct });
      },
    });

    const phases = calls.map((c) => c.phase);
    expect(phases).toContain('parse');
    expect(phases).toContain('build-indices');
    expect(phases).toContain('done');

    // 所有 pct 非负
    for (const c of calls) {
      if (c.pct !== undefined) {
        expect(c.pct).toBeGreaterThanOrEqual(0);
        expect(c.pct).toBeLessThanOrEqual(100);
      }
    }

    // done 最后
    expect(phases[phases.length - 1]).toBe('done');
    // done 时 pct=100
    expect(calls[calls.length - 1]?.pct).toBe(100);

    await doc.dispose();
  });

  test('不传 onProgress 时 rollback 仍然工作', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['a']), { docId: 'd-prog2' });
    const id = await doc.snapshots.take();
    await expect(doc.snapshots.rollback(id)).resolves.toBeUndefined();
    await doc.dispose();
  });
});

describe('snapshots pinned FIFO 保护(#108,docs/03.2 §3.3)', () => {
  test('take 不传 options → 向后兼容,FIFO 照常淘汰', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-pin-compat',
      maxSnapshots: 3,
    });
    for (let i = 0; i < 4; i++) {
      await doc.snapshots.take();
    }
    const ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toEqual(['snap-1', 'snap-2', 'snap-3']); // snap-0 被淘汰
    await doc.dispose();
  });

  test('take(reason, undefined, { pin: true }) → SnapshotInfo.pinned=true', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-pin-flag',
    });
    const id = await doc.snapshots.take('session-start', undefined, { pin: true });
    const infos = await doc.snapshots.list();
    const info = infos.find((i) => i.id === id)!;
    expect(info.pinned).toBe(true);
    await doc.dispose();
  });

  test('take(reason) 默认不 pinned:SnapshotInfo.pinned 为 undefined / 不为 true', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd-pin-default' });
    const id = await doc.snapshots.take('regular');
    const info = (await doc.snapshots.list()).find((i) => i.id === id)!;
    expect(info.pinned).not.toBe(true);
    await doc.dispose();
  });

  test('FIFO 跳过 pinned:take 1 pinned + 5 unpinned,max=5 → 淘汰最老非 pinned', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-pin-fifo',
      maxSnapshots: 5,
    });
    // snap-0 pinned + snap-1..snap-4 unpinned(总 5 条,还没超)
    await doc.snapshots.take('session-start', undefined, { pin: true }); // snap-0 pinned
    for (let i = 0; i < 4; i++) {
      await doc.snapshots.take(); // snap-1..snap-4
    }
    let ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toEqual(['snap-0', 'snap-1', 'snap-2', 'snap-3', 'snap-4']);

    // 再 take 一次 → 超限,需淘汰;snap-0 pinned 跳过,最老非 pinned 是 snap-1
    await doc.snapshots.take();
    ids = (await doc.snapshots.list()).map((i) => i.id);
    expect(ids).toHaveLength(5);
    expect(ids).toContain('snap-0'); // pinned 保留
    expect(ids).not.toContain('snap-1'); // 最老非 pinned 被淘汰
    expect(ids).toContain('snap-5'); // 新 take
    await doc.dispose();
  });

  test('多次 take 超限 pinned 始终保留,总是淘汰最老非 pinned', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), {
      docId: 'd-pin-multi',
      maxSnapshots: 3,
    });
    await doc.snapshots.take('pinned', undefined, { pin: true }); // snap-0 pinned
    await doc.snapshots.take(); // snap-1
    await doc.snapshots.take(); // snap-2
    await doc.snapshots.take(); // snap-3 → 淘汰 snap-1
    await doc.snapshots.take(); // snap-4 → 淘汰 snap-2
    await doc.snapshots.take(); // snap-5 → 淘汰 snap-3

    const infos = await doc.snapshots.list();
    const ids = infos.map((i) => i.id);
    expect(ids).toHaveLength(3);
    expect(ids).toContain('snap-0'); // pinned 保留
    expect(ids).toContain('snap-4');
    expect(ids).toContain('snap-5');
    const pinnedInfo = infos.find((i) => i.id === 'snap-0')!;
    expect(pinnedInfo.pinned).toBe(true);
    await doc.dispose();
  });

  test('所有快照都是 pinned 且超限:console.warn 提示,不抛错,take 仍成功', async () => {
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const doc = await WordDocument.fromZip(await makeZip(['x']), {
        docId: 'd-pin-all',
        maxSnapshots: 2,
      });
      await doc.snapshots.take('a', undefined, { pin: true });
      await doc.snapshots.take('b', undefined, { pin: true });
      // 第 3 次也 pinned,超限但无可淘汰 → warn + 放行
      const id = await doc.snapshots.take('c', undefined, { pin: true });
      expect(id).toBe('snap-2');
      const ids = (await doc.snapshots.list()).map((i) => i.id);
      expect(ids).toEqual(['snap-0', 'snap-1', 'snap-2']); // 全保留
      expect(warnSpy).toHaveBeenCalled();
      const warnMsg = String(warnSpy.mock.calls[0]?.[0] ?? '');
      expect(warnMsg).toMatch(/pinned/);
      await doc.dispose();
    } finally {
      warnSpy.mockRestore();
    }
  });

  test('pinned 快照可正常 rollback(消费语义;pin 不阻止手动删除)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['a', 'b']), { docId: 'd-pin-roll' });
    const id = await doc.snapshots.take('session-start', undefined, { pin: true });

    // 修改 DOM
    const pkg = doc._getPackage();
    const t0 = pkg.parts.get('word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 't')[0]!;
    t0.textContent = 'MODIFIED';

    await doc.snapshots.rollback(id);
    expect(collectBodyTexts(doc)).toEqual(['a', 'b']);
    // 消费语义:pinned 也被 rollback 删除(rollback 是显式操作,不受 FIFO 保护影响)
    expect((await doc.snapshots.list()).find((i) => i.id === id)).toBeUndefined();
    await doc.dispose();
  });

  test('pinned 快照可 delete(显式删除不受保护)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd-pin-del' });
    const id = await doc.snapshots.take('pinned', undefined, { pin: true });
    await doc.snapshots.delete(id);
    expect((await doc.snapshots.list()).find((i) => i.id === id)).toBeUndefined();
    await doc.dispose();
  });

  test('options.pin=false 等价于不传 → 普通快照', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd-pin-false' });
    const id = await doc.snapshots.take('r', undefined, { pin: false });
    const info = (await doc.snapshots.list()).find((i) => i.id === id)!;
    expect(info.pinned).not.toBe(true);
    await doc.dispose();
  });
});

describe('snapshots.delete / clear', () => {
  test('delete 按 id 删除快照(幂等)', async () => {
    const doc = await WordDocument.fromZip(await makeZip(['x']), { docId: 'd' });
    const id = await doc.snapshots.take();
    await doc.snapshots.delete(id);
    await doc.snapshots.delete(id); // 再次删除不抛
    expect((await doc.snapshots.list()).find((i) => i.id === id)).toBeUndefined();
    await doc.dispose();
  });

  test('clear 清空该 doc 的所有快照,但不影响其它 doc', async () => {
    const shared = new InMemorySnapshotStorage();
    const a = await WordDocument.fromZip(await makeZip(['a']), {
      docId: 'doc-A',
      snapshotStorage: shared,
    });
    const b = await WordDocument.fromZip(await makeZip(['b']), {
      docId: 'doc-B',
      snapshotStorage: shared,
    });
    await a.snapshots.take();
    await a.snapshots.take();
    await b.snapshots.take();

    await a.snapshots.clear();
    expect(await a.snapshots.list()).toEqual([]);
    expect((await b.snapshots.list()).length).toBe(1);
    await a.dispose();
    await b.dispose();
  });
});
