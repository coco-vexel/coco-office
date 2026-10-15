import { describe, expect, test } from 'vitest';
import { InMemorySnapshotStorage, type SnapshotPayload } from '../index.js';

function makePayload(id: string, extra?: Partial<SnapshotPayload['info']>): SnapshotPayload {
  return {
    info: {
      id,
      timestamp: Date.now(),
      reason: `take-${id}`,
      changesCaptured: [],
      ...extra,
    },
    parts: [
      {
        path: 'word/document.xml',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        bytes: new Uint8Array([0x3c, 0x72, 0x2f, 0x3e]),
      },
    ],
  };
}

describe('InMemorySnapshotStorage — CRUD', () => {
  test('save + load 返回同一 payload', async () => {
    const storage = new InMemorySnapshotStorage();
    const payload = makePayload('snap-0');
    await storage.save('doc-A', 'snap-0', payload);
    const back = await storage.load('doc-A', 'snap-0');
    expect(back).toBe(payload);
  });

  test('load 不存在抛错', async () => {
    const storage = new InMemorySnapshotStorage();
    await expect(storage.load('doc-A', 'snap-0')).rejects.toThrow(/snapshot not found/);
  });

  test('load 存在但 docId 空时抛错', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await expect(storage.load('doc-B', 'snap-0')).rejects.toThrow(/snapshot not found/);
  });

  test('delete 幂等:不存在时不抛错', async () => {
    const storage = new InMemorySnapshotStorage();
    await expect(storage.delete('doc-A', 'snap-0')).resolves.toBeUndefined();
  });

  test('delete 存在时移除,load 后再 load 抛错', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await storage.delete('doc-A', 'snap-0');
    await expect(storage.load('doc-A', 'snap-0')).rejects.toThrow(/snapshot not found/);
  });
});

describe('InMemorySnapshotStorage — list', () => {
  test('空 docId 返空数组', async () => {
    const storage = new InMemorySnapshotStorage();
    expect(await storage.list('unknown')).toEqual([]);
  });

  test('只返回 info,不含 parts', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0', { reason: 'r0' }));
    const infos = await storage.list('doc-A');
    expect(infos).toHaveLength(1);
    expect(infos[0]?.id).toBe('snap-0');
    expect(infos[0]?.reason).toBe('r0');
    // @ts-expect-error parts 不应在 info 上
    expect(infos[0]?.parts).toBeUndefined();
  });

  test('按 snap-<N> 数字自增序排序(不是插入顺序)', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-2', makePayload('snap-2'));
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await storage.save('doc-A', 'snap-10', makePayload('snap-10'));
    await storage.save('doc-A', 'snap-1', makePayload('snap-1'));
    const infos = await storage.list('doc-A');
    expect(infos.map((i) => i.id)).toEqual(['snap-0', 'snap-1', 'snap-2', 'snap-10']);
  });

  test('非 snap-<N> 模式的 id 走字典序回退', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'beta', makePayload('beta'));
    await storage.save('doc-A', 'alpha', makePayload('alpha'));
    const infos = await storage.list('doc-A');
    expect(infos.map((i) => i.id)).toEqual(['alpha', 'beta']);
  });
});

describe('InMemorySnapshotStorage — clear', () => {
  test('clear 只清指定 docId 的快照', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await storage.save('doc-A', 'snap-1', makePayload('snap-1'));
    await storage.save('doc-B', 'snap-0', makePayload('snap-0'));

    await storage.clear('doc-A');

    expect(await storage.list('doc-A')).toEqual([]);
    expect((await storage.list('doc-B')).map((i) => i.id)).toEqual(['snap-0']);
  });

  test('clear 不存在的 docId 不抛错', async () => {
    const storage = new InMemorySnapshotStorage();
    await expect(storage.clear('never-exists')).resolves.toBeUndefined();
  });
});

describe('InMemorySnapshotStorage — docId 命名空间隔离', () => {
  test('不同 docId 存同一 snap id 互不影响', async () => {
    const storage = new InMemorySnapshotStorage();
    const payloadA = makePayload('snap-0', { reason: 'doc-A-origin' });
    const payloadB = makePayload('snap-0', { reason: 'doc-B-origin' });
    await storage.save('doc-A', 'snap-0', payloadA);
    await storage.save('doc-B', 'snap-0', payloadB);

    const backA = await storage.load('doc-A', 'snap-0');
    const backB = await storage.load('doc-B', 'snap-0');
    expect(backA.info.reason).toBe('doc-A-origin');
    expect(backB.info.reason).toBe('doc-B-origin');
  });

  test('list 按 docId 隔离', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await storage.save('doc-A', 'snap-1', makePayload('snap-1'));
    await storage.save('doc-B', 'snap-0', makePayload('snap-0'));

    expect((await storage.list('doc-A')).map((i) => i.id)).toEqual(['snap-0', 'snap-1']);
    expect((await storage.list('doc-B')).map((i) => i.id)).toEqual(['snap-0']);
  });
});

describe('InMemorySnapshotStorage — 引用语义', () => {
  test('save 后 load 返回的是同一对象引用(不拷贝)', async () => {
    const storage = new InMemorySnapshotStorage();
    const payload = makePayload('snap-0');
    await storage.save('doc-A', 'snap-0', payload);
    const back = await storage.load('doc-A', 'snap-0');
    expect(back).toBe(payload);
    expect(back.parts[0]?.bytes).toBe(payload.parts[0]?.bytes);
  });
});

describe('InMemorySnapshotStorage — pinned 支持(#108)', () => {
  test('save 不传 meta.pinned → list 里 pinned 字段 undefined(向后兼容)', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    const infos = await storage.list('doc-A');
    expect(infos).toHaveLength(1);
    expect(infos[0]?.pinned).toBeUndefined();
  });

  test('save 传 meta.pinned=true → list 里 pinned=true', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'), { pinned: true });
    const infos = await storage.list('doc-A');
    expect(infos[0]?.pinned).toBe(true);
  });

  test('save 传 meta.pinned=false → list 里 pinned 字段仍 undefined(假值不携带)', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'), { pinned: false });
    const infos = await storage.list('doc-A');
    expect(infos[0]?.pinned).toBeUndefined();
  });

  test('混合 pinned / 非 pinned:list 每条各携带自己的 pinned 状态', async () => {
    const storage = new InMemorySnapshotStorage();
    await storage.save('doc-A', 'snap-0', makePayload('snap-0'));
    await storage.save('doc-A', 'snap-1', makePayload('snap-1'), { pinned: true });
    await storage.save('doc-A', 'snap-2', makePayload('snap-2'));
    const infos = await storage.list('doc-A');
    const byId: Record<string, boolean | undefined> = {};
    for (const i of infos) byId[i.id] = i.pinned;
    expect(byId['snap-0']).toBeUndefined();
    expect(byId['snap-1']).toBe(true);
    expect(byId['snap-2']).toBeUndefined();
  });

  test('pinned save 不改变 load 返回引用(load 仍返原 payload)', async () => {
    const storage = new InMemorySnapshotStorage();
    const payload = makePayload('snap-0');
    await storage.save('doc-A', 'snap-0', payload, { pinned: true });
    const back = await storage.load('doc-A', 'snap-0');
    expect(back).toBe(payload);
  });
});
