import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import {
  createBinaryPart,
  createPackage,
  createXmlPart,
  getPart,
  hasPart,
  setPart,
} from '../../pack/index.js';
import { runTransaction } from '../transactions.js';
import { W_NS } from '../namespaces.js';

function xmlPart(path: string, xml: string) {
  return createXmlPart({
    path,
    contentType: 'application/xml',
    xmlDom: parseXmlDocument(xml) as unknown as Document,
    originalBytes: new TextEncoder().encode(xml),
  });
}

function serialize(part: { xmlDom?: Document }): string {
  return serializeToWellFormedString(part.xmlDom as never);
}

describe('runTransaction — 成功路径', () => {
  test('work 成功:改动保留,work 返回值被传出', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );

    const result = await runTransaction(pkg, ['word/document.xml'], () => {
      const part = getPart(pkg, 'word/document.xml')!;
      // 手工在 body 下加一个段落
      const body = part.xmlDom!.getElementsByTagNameNS(W_NS, 'body')[0]!;
      const p = part.xmlDom!.createElementNS(W_NS, 'w:p');
      body.appendChild(p);
      part._dirty = true;
      return 42;
    });

    expect(result).toBe(42);
    const doc = getPart(pkg, 'word/document.xml')!.xmlDom!;
    const ps = doc.getElementsByTagNameNS(W_NS, 'p');
    expect(ps.length).toBe(1);
    expect(getPart(pkg, 'word/document.xml')!._dirty).toBe(true);
  });

  test('work 是 async:await 的变更同样提交', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );

    await runTransaction(pkg, ['word/document.xml'], async () => {
      await Promise.resolve();
      const part = getPart(pkg, 'word/document.xml')!;
      const body = part.xmlDom!.getElementsByTagNameNS(W_NS, 'body')[0]!;
      body.appendChild(part.xmlDom!.createElementNS(W_NS, 'w:p'));
      part._dirty = true;
    });

    expect(
      getPart(pkg, 'word/document.xml')!.xmlDom!.getElementsByTagNameNS(W_NS, 'p').length,
    ).toBe(1);
  });
});

describe('runTransaction — 回滚:已存在的 part', () => {
  test('work 抛错 → xmlDom 恢复原状', async () => {
    const pkg = createPackage();
    const xml = `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>orig</w:t></w:r></w:p></w:body></w:document>`;
    setPart(pkg, xmlPart('word/document.xml', xml));
    const beforeSerialize = serialize(getPart(pkg, 'word/document.xml')!);

    await expect(
      runTransaction(pkg, ['word/document.xml'], () => {
        const part = getPart(pkg, 'word/document.xml')!;
        const t = part.xmlDom!.getElementsByTagNameNS(W_NS, 't')[0]!;
        t.textContent = 'MODIFIED';
        part._dirty = true;
        throw new Error('boom');
      }),
    ).rejects.toThrow(/boom/);

    const after = getPart(pkg, 'word/document.xml')!;
    expect(serialize(after)).toBe(beforeSerialize);
    expect(after._dirty).toBe(false);
  });

  test('_dirty 在 work 前为 true 时,回滚后仍是 true', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );
    const part = getPart(pkg, 'word/document.xml')!;
    part._dirty = true; // 预设为 dirty

    await expect(
      runTransaction(pkg, ['word/document.xml'], () => {
        part._dirty = false;
        throw new Error('x');
      }),
    ).rejects.toThrow();

    expect(getPart(pkg, 'word/document.xml')!._dirty).toBe(true);
  });

  test('二进制 part:bytes 在回滚后恢复', async () => {
    const pkg = createPackage();
    const orig = new Uint8Array([1, 2, 3, 4]);
    setPart(
      pkg,
      createBinaryPart({ path: 'word/media/a.png', contentType: 'image/png', bytes: orig }),
    );

    await expect(
      runTransaction(pkg, ['word/media/a.png'], () => {
        const part = getPart(pkg, 'word/media/a.png')!;
        part.bytes = new Uint8Array([9, 9, 9]);
        part._dirty = true;
        throw new Error('x');
      }),
    ).rejects.toThrow();

    const back = getPart(pkg, 'word/media/a.png')!;
    expect(Array.from(back.bytes!)).toEqual([1, 2, 3, 4]);
    expect(back._dirty).toBe(false);
  });

  test('多 part 事务:任一 work 抛错 → 全部回滚', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );
    setPart(pkg, xmlPart('word/styles.xml', `<w:styles xmlns:w="${W_NS}"/>`));

    await expect(
      runTransaction(pkg, ['word/document.xml', 'word/styles.xml'], () => {
        // 先改 document.xml
        const doc = getPart(pkg, 'word/document.xml')!;
        const body = doc.xmlDom!.getElementsByTagNameNS(W_NS, 'body')[0]!;
        body.appendChild(doc.xmlDom!.createElementNS(W_NS, 'w:p'));
        doc._dirty = true;
        // 再改 styles.xml
        const styles = getPart(pkg, 'word/styles.xml')!;
        styles.xmlDom!.documentElement!.setAttribute('x', 'y');
        styles.xmlDom!.documentElement!.appendChild(
          styles.xmlDom!.createElementNS(W_NS, 'w:style'),
        );
        styles._dirty = true;
        // 抛错
        throw new Error('rollback please');
      }),
    ).rejects.toThrow();

    const doc = getPart(pkg, 'word/document.xml')!;
    const styles = getPart(pkg, 'word/styles.xml')!;
    expect(doc.xmlDom!.getElementsByTagNameNS(W_NS, 'p').length).toBe(0);
    expect(styles.xmlDom!.documentElement!.getAttribute('x')).toBeNull();
    expect(doc._dirty).toBe(false);
    expect(styles._dirty).toBe(false);
  });
});

describe('runTransaction — 回滚:work 新增的 part', () => {
  test('work 新建 part 后抛错 → part 被删除', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);

    await expect(
      runTransaction(pkg, ['word/document.xml', 'word/comments.xml'], () => {
        setPart(pkg, xmlPart('word/comments.xml', `<w:comments xmlns:w="${W_NS}"/>`));
        throw new Error('rollback new part');
      }),
    ).rejects.toThrow();

    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
  });

  test('work 新建 part 后成功 → part 保留', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );

    await runTransaction(pkg, ['word/document.xml', 'word/comments.xml'], () => {
      setPart(pkg, xmlPart('word/comments.xml', `<w:comments xmlns:w="${W_NS}"/>`));
    });

    expect(hasPart(pkg, 'word/comments.xml')).toBe(true);
  });
});

describe('runTransaction — 未列入的 path 不受事务保护', () => {
  test('work 改了未列入 affectedPartPaths 的 part,抛错后该 part 不回滚', async () => {
    const pkg = createPackage();
    setPart(
      pkg,
      xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"><w:body/></w:document>`),
    );
    setPart(pkg, xmlPart('word/styles.xml', `<w:styles xmlns:w="${W_NS}"/>`));

    await expect(
      runTransaction(pkg, ['word/document.xml'], () => {
        const styles = getPart(pkg, 'word/styles.xml')!;
        // styles 未列入 affected,work 抛错后不会回滚
        styles.xmlDom!.documentElement!.setAttribute('leaked', 'yes');
        styles._dirty = true;
        throw new Error('x');
      }),
    ).rejects.toThrow();

    // styles.xml 未回滚(预期;文档约定调用方必须列全 affected paths)
    expect(getPart(pkg, 'word/styles.xml')!.xmlDom!.documentElement!.getAttribute('leaked')).toBe(
      'yes',
    );
  });
});

describe('runTransaction — 边界', () => {
  test('affectedPartPaths 含不存在的 path:被视为"事务中可能新建",work 成功就保留不管有无创建', async () => {
    const pkg = createPackage();
    // 事务声明 affected 含 word/comments.xml,但 work 没新建
    await runTransaction(pkg, ['word/comments.xml'], () => {});
    expect(hasPart(pkg, 'word/comments.xml')).toBe(false);
  });

  test('work 抛错时 affectedPartPaths 含不存在的 path 不会误删已有 part', async () => {
    const pkg = createPackage();
    setPart(pkg, xmlPart('word/document.xml', `<w:document xmlns:w="${W_NS}"/>`));

    await expect(
      runTransaction(pkg, ['word/comments.xml'], () => {
        throw new Error('x');
      }),
    ).rejects.toThrow();

    // document.xml 不在 affected,不受影响
    expect(hasPart(pkg, 'word/document.xml')).toBe(true);
  });

  test('空 affectedPartPaths + work 成功 → 返回值透传', async () => {
    const pkg = createPackage();
    const out = await runTransaction(pkg, [], () => 'ok');
    expect(out).toBe('ok');
  });
});
