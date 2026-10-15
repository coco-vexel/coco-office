import { describe, expect, test } from 'vitest';
import { A_NS, MC_NS, PKG_NS, R_NS, W14_NS, W_NS, XML_NS } from '../namespaces.js';

describe('OOXML 命名空间 URI 常量', () => {
  test('值与 docs/01 附录 A 一致', () => {
    expect(W_NS).toBe('http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    expect(R_NS).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    expect(A_NS).toBe('http://schemas.openxmlformats.org/drawingml/2006/main');
    expect(MC_NS).toBe('http://schemas.openxmlformats.org/markup-compatibility/2006');
    expect(W14_NS).toBe('http://schemas.microsoft.com/office/word/2010/wordml');
    expect(PKG_NS).toBe('http://schemas.microsoft.com/office/2006/xmlPackage');
    expect(XML_NS).toBe('http://www.w3.org/XML/1998/namespace');
  });

  test('全部为字符串且非空', () => {
    for (const ns of [W_NS, R_NS, A_NS, MC_NS, W14_NS, PKG_NS, XML_NS]) {
      expect(typeof ns).toBe('string');
      expect(ns.length).toBeGreaterThan(0);
    }
  });
});
