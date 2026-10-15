import { describe, expect, test } from 'vitest';
import { parseXmlDocument } from 'slimdom';
import {
  createBinaryPart,
  createPackage,
  createXmlPart,
  setPart,
  type Package,
} from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import { getStoryParagraphs, listStories } from '../stories.js';

const CT = {
  body: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  macroBody: 'application/vnd.ms-word.document.macroEnabled.main+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
};

function putXml(pkg: Package, path: string, xml: string, ct: string): void {
  setPart(
    pkg,
    createXmlPart({
      path,
      contentType: ct,
      xmlDom: parseXmlDocument(xml) as unknown as Document,
      originalBytes: new TextEncoder().encode(xml),
    }),
  );
}

const DOC_XML = `<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p><w:r><w:t>body-p1</w:t></w:r></w:p>
    <w:p><w:r><w:t>body-p2</w:t></w:r></w:p>
  </w:body>
</w:document>`;

const HDR_XML = `<w:hdr xmlns:w="${W_NS}">
  <w:p><w:r><w:t>header-p</w:t></w:r></w:p>
</w:hdr>`;

const FTR_XML = `<w:ftr xmlns:w="${W_NS}">
  <w:p><w:r><w:t>footer-p</w:t></w:r></w:p>
</w:ftr>`;

const FOOTNOTES_XML = `<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:r><w:t>fn-1</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>fn-2a</w:t></w:r></w:p><w:p><w:r><w:t>fn-2b</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

const COMMENTS_XML = `<w:comments xmlns:w="${W_NS}">
  <w:comment w:id="5" w:author="A"><w:p><w:r><w:t>c5</w:t></w:r></w:p></w:comment>
  <w:comment w:id="6" w:author="B"><w:p><w:r><w:t>c6-p1</w:t></w:r></w:p><w:p><w:r><w:t>c6-p2</w:t></w:r></w:p></w:comment>
</w:comments>`;

describe('listStories — 基础', () => {
  test('空包返空数组', () => {
    expect(listStories(createPackage())).toEqual([]);
  });

  test('只有 body:返回一个 body story', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    const stories = listStories(pkg);
    expect(stories).toHaveLength(1);
    expect(stories[0]?.id).toBe('body');
    expect(stories[0]?.kind).toBe('body');
    expect(stories[0]?.root.localName).toBe('body');
    expect(stories[0]?.partPath).toBe('word/document.xml');
  });

  test('body + 多 header + 多 footer 正确枚举', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    putXml(pkg, 'word/header1.xml', HDR_XML, CT.header);
    putXml(pkg, 'word/header2.xml', HDR_XML, CT.header);
    putXml(pkg, 'word/footer1.xml', FTR_XML, CT.footer);
    const stories = listStories(pkg);
    expect(stories.map((s) => s.id).sort()).toEqual([
      'body',
      'footer:word/footer1.xml',
      'header:word/header1.xml',
      'header:word/header2.xml',
    ]);
  });

  test('footnotes:每个 <w:footnote> 成独立 story,id = footnote:<w:id>', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/footnotes.xml', FOOTNOTES_XML, CT.footnotes);
    const stories = listStories(pkg);
    const fnIds = stories.filter((s) => s.kind === 'footnote').map((s) => s.id);
    expect(fnIds).toEqual(['footnote:-1', 'footnote:1', 'footnote:2']);
  });

  test('comments:每个 <w:comment> 成独立 story,id = comment:<w:id>', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/comments.xml', COMMENTS_XML, CT.comments);
    const stories = listStories(pkg);
    expect(stories).toHaveLength(2);
    expect(stories.map((s) => s.id).sort()).toEqual(['comment:5', 'comment:6']);
    expect(stories.every((s) => s.kind === 'comment')).toBe(true);
  });

  test('全 6 类 stories 同时存在:数量正确', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    putXml(pkg, 'word/header1.xml', HDR_XML, CT.header);
    putXml(pkg, 'word/footer1.xml', FTR_XML, CT.footer);
    putXml(pkg, 'word/footnotes.xml', FOOTNOTES_XML, CT.footnotes);
    putXml(
      pkg,
      'word/endnotes.xml',
      FOOTNOTES_XML.replace(/footnote/g, 'endnote').replace(/footnotes/g, 'endnotes'),
      CT.endnotes,
    );
    putXml(pkg, 'word/comments.xml', COMMENTS_XML, CT.comments);

    const stories = listStories(pkg);
    const byKind = new Map<string, number>();
    for (const s of stories) byKind.set(s.kind, (byKind.get(s.kind) ?? 0) + 1);
    expect(byKind.get('body')).toBe(1);
    expect(byKind.get('header')).toBe(1);
    expect(byKind.get('footer')).toBe(1);
    expect(byKind.get('footnote')).toBe(3);
    expect(byKind.get('endnote')).toBe(3);
    expect(byKind.get('comment')).toBe(2);
  });

  test('非 Word content part(styles / 图片)不被枚举', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    putXml(pkg, 'word/styles.xml', `<w:styles xmlns:w="${W_NS}"/>`, CT.styles);
    setPart(
      pkg,
      createBinaryPart({
        path: 'word/media/img.png',
        contentType: 'image/png',
        bytes: new Uint8Array([1]),
      }),
    );
    const stories = listStories(pkg);
    expect(stories).toHaveLength(1);
    expect(stories[0]?.kind).toBe('body');
  });
});

describe('listStories macro-enabled main documents', () => {
  test('macroEnabled main document is treated as body story', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.macroBody);
    const stories = listStories(pkg);
    expect(stories).toHaveLength(1);
    expect(stories[0]?.id).toBe('body');
    expect(stories[0]?.kind).toBe('body');
    expect(getStoryParagraphs(stories[0]!)).toHaveLength(2);
  });
});

describe('listStories — kinds 过滤', () => {
  test('只取 header', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    putXml(pkg, 'word/header1.xml', HDR_XML, CT.header);
    putXml(pkg, 'word/footer1.xml', FTR_XML, CT.footer);
    const stories = listStories(pkg, { kinds: ['header'] });
    expect(stories).toHaveLength(1);
    expect(stories[0]?.kind).toBe('header');
  });

  test('只取 body + footnote', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    putXml(pkg, 'word/header1.xml', HDR_XML, CT.header);
    putXml(pkg, 'word/footnotes.xml', FOOTNOTES_XML, CT.footnotes);
    const stories = listStories(pkg, { kinds: ['body', 'footnote'] });
    expect(stories.map((s) => s.kind).sort()).toEqual(['body', 'footnote', 'footnote', 'footnote']);
  });

  test('kinds=[] 返空数组', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    expect(listStories(pkg, { kinds: [] })).toEqual([]);
  });
});

describe('getStoryParagraphs', () => {
  test('body story 返回 <w:body> 下所有 <w:p>', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    const [body] = listStories(pkg);
    const ps = getStoryParagraphs(body!);
    expect(ps).toHaveLength(2);
    expect(ps[0]?.localName).toBe('p');
  });

  test('footnote story 只返回该 <w:footnote> 内的段落(不跨 sub-story)', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/footnotes.xml', FOOTNOTES_XML, CT.footnotes);
    const stories = listStories(pkg);
    const fn2 = stories.find((s) => s.id === 'footnote:2')!;
    const ps = getStoryParagraphs(fn2);
    expect(ps).toHaveLength(2); // fn-2a 和 fn-2b,不会把 fn-1 的段落也拿来
  });

  test('comment story 段落单独隔离', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/comments.xml', COMMENTS_XML, CT.comments);
    const stories = listStories(pkg);
    const c5 = stories.find((s) => s.id === 'comment:5')!;
    const c6 = stories.find((s) => s.id === 'comment:6')!;
    expect(getStoryParagraphs(c5)).toHaveLength(1);
    expect(getStoryParagraphs(c6)).toHaveLength(2);
  });
});

describe('listStories — 性能基线', () => {
  test('body + 5 header + 5 footer + 100 footnotes + 100 endnotes + 50 comments < 100ms', () => {
    const pkg = createPackage();
    putXml(pkg, 'word/document.xml', DOC_XML, CT.body);
    for (let i = 1; i <= 5; i++) {
      putXml(pkg, `word/header${i}.xml`, HDR_XML, CT.header);
      putXml(pkg, `word/footer${i}.xml`, FTR_XML, CT.footer);
    }
    const fnBody = Array.from(
      { length: 100 },
      (_, i) => `<w:footnote w:id="${i}"><w:p><w:r><w:t>fn${i}</w:t></w:r></w:p></w:footnote>`,
    ).join('');
    putXml(
      pkg,
      'word/footnotes.xml',
      `<w:footnotes xmlns:w="${W_NS}">${fnBody}</w:footnotes>`,
      CT.footnotes,
    );
    const enBody = Array.from(
      { length: 100 },
      (_, i) => `<w:endnote w:id="${i}"><w:p><w:r><w:t>en${i}</w:t></w:r></w:p></w:endnote>`,
    ).join('');
    putXml(
      pkg,
      'word/endnotes.xml',
      `<w:endnotes xmlns:w="${W_NS}">${enBody}</w:endnotes>`,
      CT.endnotes,
    );
    const commBody = Array.from(
      { length: 50 },
      (_, i) => `<w:comment w:id="${i}"><w:p><w:r><w:t>c${i}</w:t></w:r></w:p></w:comment>`,
    ).join('');
    putXml(
      pkg,
      'word/comments.xml',
      `<w:comments xmlns:w="${W_NS}">${commBody}</w:comments>`,
      CT.comments,
    );

    const start = performance.now();
    const stories = listStories(pkg);
    const elapsed = performance.now() - start;

    expect(stories).toHaveLength(1 + 5 + 5 + 100 + 100 + 50);
    expect(elapsed).toBeLessThan(100);
  });
});
