import { describe, expect, test } from 'vitest';
import { parseXmlDocument, serializeToWellFormedString } from 'slimdom';
import { W_NS } from '../namespaces.js';
import { applyXQueryUpdate } from '../xquery-update.js';

function wNs(prefix: string): string | null {
  if (prefix === 'w') return W_NS;
  return null;
}

function parse(xml: string): Document {
  return parseXmlDocument(xml) as unknown as Document;
}

function serialize(node: Node): string {
  return serializeToWellFormedString(node as never);
}

describe('applyXQueryUpdate — 基础', () => {
  test('replace value of:改 <w:t> 文本', async () => {
    const dom = parse(`<w:p xmlns:w="${W_NS}"><w:r><w:t>old text</w:t></w:r></w:p>`);
    const p = dom.documentElement!;
    const res = await applyXQueryUpdate(p, 'replace value of node ./w:r/w:t with "new text"', {
      namespaceResolver: wNs,
    });
    expect(res.mutations).toBe(1);
    const t = dom.getElementsByTagNameNS(W_NS, 't')[0];
    expect(t?.textContent).toBe('new text');
  });

  test('insert node into:添加新 <w:t>', async () => {
    const dom = parse(`<w:r xmlns:w="${W_NS}"></w:r>`);
    const r = dom.documentElement!;
    const res = await applyXQueryUpdate(
      r,
      'insert node <w:t xmlns:w="' + W_NS + '">added</w:t> into .',
      { namespaceResolver: wNs },
    );
    expect(res.mutations).toBe(1);
    const ts = dom.getElementsByTagNameNS(W_NS, 't');
    expect(ts.length).toBe(1);
    expect(ts[0]?.textContent).toBe('added');
  });

  test('delete node:移除 <w:r>', async () => {
    const dom = parse(
      `<w:p xmlns:w="${W_NS}"><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>`,
    );
    const res = await applyXQueryUpdate(dom.documentElement!, 'delete node ./w:r[1]', {
      namespaceResolver: wNs,
    });
    expect(res.mutations).toBe(1);
    const rs = dom.getElementsByTagNameNS(W_NS, 'r');
    expect(rs.length).toBe(1);
    expect(rs[0]?.textContent).toBe('b');
  });

  test('replace node:整体替换一个 <w:r>', async () => {
    const dom = parse(`<w:p xmlns:w="${W_NS}"><w:r><w:t>old</w:t></w:r></w:p>`);
    const res = await applyXQueryUpdate(
      dom.documentElement!,
      `replace node ./w:r with <w:r xmlns:w="${W_NS}"><w:t>replaced</w:t></w:r>`,
      { namespaceResolver: wNs },
    );
    expect(res.mutations).toBe(1);
    expect(dom.getElementsByTagNameNS(W_NS, 't')[0]?.textContent).toBe('replaced');
  });

  test('无效 XQuery 表达式抛错', async () => {
    const dom = parse(`<w:p xmlns:w="${W_NS}"><w:r/></w:p>`);
    await expect(
      applyXQueryUpdate(dom.documentElement!, 'not valid xquery !!!', { namespaceResolver: wNs }),
    ).rejects.toThrow();
  });
});

describe('applyXQueryUpdate — 变量绑定', () => {
  test('$variable 从 options.variables 读取', async () => {
    const dom = parse(`<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>`);
    const res = await applyXQueryUpdate(
      dom.documentElement!,
      'replace value of node ./w:r/w:t with $text',
      { namespaceResolver: wNs, variables: { text: 'from-var' } },
    );
    expect(res.mutations).toBe(1);
    expect(dom.getElementsByTagNameNS(W_NS, 't')[0]?.textContent).toBe('from-var');
  });
});

describe('applyXQueryUpdate — mutations 计数', () => {
  test('多个节点被修改时 mutations 反映条目数', async () => {
    const dom = parse(
      `<w:p xmlns:w="${W_NS}"><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>`,
    );
    const res = await applyXQueryUpdate(
      dom.documentElement!,
      'for $t in ./w:r/w:t return replace value of node $t with "same"',
      { namespaceResolver: wNs },
    );
    expect(res.mutations).toBe(2);
    const ts = dom.getElementsByTagNameNS(W_NS, 't');
    expect(ts[0]?.textContent).toBe('same');
    expect(ts[1]?.textContent).toBe('same');
  });

  test('空 PUL(无匹配 target)→ mutations = 0,不改变 DOM', async () => {
    const dom = parse(`<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>`);
    const before = serialize(dom);
    const res = await applyXQueryUpdate(
      dom.documentElement!,
      'for $t in ./w:nonexistent return replace value of node $t with "y"',
      { namespaceResolver: wNs },
    );
    expect(res.mutations).toBe(0);
    expect(serialize(dom)).toBe(before);
  });
});
