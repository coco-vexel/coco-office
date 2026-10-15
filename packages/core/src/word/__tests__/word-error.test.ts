import { describe, expect, test } from 'vitest';
import { WordError, type WordErrorCode } from '../word-error.js';

describe('WordError', () => {
  test('基本构造:code / message / docId / details', () => {
    const err = new WordError('PARA_NOT_FOUND', 'paragraph not found', 'doc-xyz', {
      paraId: 'AAAA1111',
    });
    expect(err.code).toBe('PARA_NOT_FOUND');
    expect(err.message).toBe('paragraph not found');
    expect(err.docId).toBe('doc-xyz');
    expect(err.details).toEqual({ paraId: 'AAAA1111' });
    expect(err.name).toBe('WordError');
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(WordError);
  });

  test('details 默认为空对象', () => {
    const err = new WordError('PLAN_EMPTY', 'empty', 'd');
    expect(err.details).toEqual({});
  });

  test('可用 instanceof 捕获', () => {
    try {
      throw new WordError('BOOKMARK_NOT_FOUND', 'x', 'd');
    } catch (e) {
      expect(e instanceof WordError).toBe(true);
      expect(e instanceof Error).toBe(true);
    }
  });

  test('所有 WordErrorCode 值都是字符串(类型级 smoke)', () => {
    const codes: WordErrorCode[] = [
      'PACKAGE_CORRUPT',
      'PACKAGE_UNSUPPORTED_FORMAT',
      'BOOKMARK_NAME_DUPLICATE_ON_LOAD',
      'BOOKMARK_NAME_EMPTY_ON_LOAD',
      'WORKER_CREATION_FAILED',
      'WORKER_RELOADED',
      'PARA_NOT_FOUND',
      'BOOKMARK_NOT_FOUND',
      'COMMENT_NOT_FOUND',
      'REVISION_NOT_FOUND',
      'TEXT_NOT_FOUND_IN_PARA',
      'BOOKMARK_NAME_EXISTS',
      'BOOKMARK_NAME_INVALID',
      'ANCHOR_CONFLICT',
      'UNSUPPORTED_OOXML_FEATURE',
      'PLAN_ALREADY_APPLIED',
      'PLAN_EMPTY',
    ];
    for (const c of codes) expect(typeof c).toBe('string');
    // 至少应该覆盖 docs/03.1 §4.1 列出的 17 个(含 WORKER_RELOADED 来自 docs/03 §4.7)
    expect(codes.length).toBeGreaterThanOrEqual(17);
  });
});
