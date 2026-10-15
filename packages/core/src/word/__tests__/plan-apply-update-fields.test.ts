/**
 * plan.setUpdateFields 端到端测试（task-core-update-fields）。
 *
 * 测试矩阵（spec §4）：
 *   1  settings 存在、无 updateFields → setUpdateFields(true)
 *      → settings.xml 出现 <w:updateFields w:val="true"/>，且只有一个
 *   2  settings 已有 updateFields="false" → setUpdateFields(true)
 *      → 改为 "true"，不新增重复元素
 *   3  setUpdateFields(false) → 写 "false"
 *   4  round-trip zip → set(true) → toZip → fromZip → 读回 updateFields == "true"
 *   4b FlatOpc round-trip
 *   5  幂等：连续 set(true) 两次 → 仍只一个 updateFields 元素
 *   6  （边界）settings 不存在 → 抛 UNSUPPORTED_OOXML_FEATURE
 *   7  不回归：现有 settings 其他子元素（zoom/defaultTabStop）不被破坏
 *
 * inline 模式。
 */

import { afterAll, describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { WordDocument } from '../word-document.js';
import { WordError } from '../word-error.js';
import { W14_NS, W_NS } from '../namespaces.js';

// ─────────── ContentType 常量（测试用） ──────────────────────────────────────

const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

// ─────────── fixture 构造 ─────────────────────────────────────────────────────

const CT_XML_WITH_SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="${SETTINGS_CT}"/>
</Types>`;

const CT_XML_NO_SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const MAIN_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:body>
    <w:p w14:paraId="AAAA0001"><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>`;
}

/** settings.xml 无 updateFields（典型干净状态） */
function settingsXmlClean(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:defaultTabStop w:val="720"/>
  <w:zoom w:percent="100"/>
</w:settings>`;
}

/** settings.xml 已有 updateFields="false" */
function settingsXmlWithFalse(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:defaultTabStop w:val="720"/>
  <w:updateFields w:val="false"/>
  <w:zoom w:percent="100"/>
</w:settings>`;
}

/** settings.xml 已有 updateFields="true" */
function settingsXmlWithTrue(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:defaultTabStop w:val="720"/>
  <w:updateFields w:val="true"/>
</w:settings>`;
}

/**
 * 构造含 settings.xml 的 fixture zip。
 * settingsXml 为 undefined 时不写 settings.xml（用于测试边界场景 6）。
 */
async function makeFixture(settingsXml?: string): Promise<Uint8Array> {
  const zip = new JSZip();
  if (settingsXml !== undefined) {
    zip.file('[Content_Types].xml', CT_XML_WITH_SETTINGS);
    zip.file('word/settings.xml', settingsXml);
  } else {
    zip.file('[Content_Types].xml', CT_XML_NO_SETTINGS);
  }
  zip.file('_rels/.rels', MAIN_RELS);
  zip.file('word/document.xml', docXml());
  return zip.generateAsync({ type: 'uint8array' });
}

/**
 * 从 zip 字节读取 settings.xml 内容（字符串），文件不存在返回 null。
 */
async function readSettingsXml(zipBytes: Uint8Array): Promise<string | null> {
  const zip = await JSZip.loadAsync(zipBytes);
  const file = zip.file('word/settings.xml');
  if (!file) return null;
  return file.async('string');
}

afterAll(() => {
  WordDocument._resetDefaults();
});

// ─────────── 用例 1: settings 存在无 updateFields → set(true) ────────────────

describe('setUpdateFields 基本写入', () => {
  test('1 settings 存在无 updateFields → setUpdateFields(true) → 出现 true，只一个', async () => {
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-1', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();

    // updateFields 出现，且 val="true"
    expect(settingsXml).toContain('updateFields');
    expect(settingsXml).toContain('"true"');

    // 确认只有一个 updateFields 元素
    const matches = [...(settingsXml ?? '').matchAll(/updateFields/g)];
    expect(matches.length).toBe(1);
  });

  // ─────────── 用例 2: settings 已有 updateFields="false" → 改为 true ──────────

  test('2 settings 已有 updateFields="false" → setUpdateFields(true) → 改为 true，无重复', async () => {
    const buf = await makeFixture(settingsXmlWithFalse());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-2', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();

    // 改为 true
    expect(settingsXml).toContain('"true"');
    // 不应再有 false（改掉了）
    expect(settingsXml).not.toContain('"false"');

    // 仍只有一个 updateFields 元素（不重复插入）
    const matches = [...(settingsXml ?? '').matchAll(/updateFields/g)];
    expect(matches.length).toBe(1);
  });

  // ─────────── 用例 3: setUpdateFields(false) → 写 false ───────────────────────

  test('3 setUpdateFields(false) → 写 false', async () => {
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-3', mode: 'inline' });

    await doc.plan().setUpdateFields(false).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();
    expect(settingsXml).toContain('updateFields');
    expect(settingsXml).toContain('"false"');
  });

  // ─────────── 用例 5: 幂等 ────────────────────────────────────────────────────

  test('5 幂等：连续 set(true) 两次 → 仍只一个 updateFields 元素', async () => {
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-5', mode: 'inline' });

    // 连续两次 apply（两个独立 plan）
    await doc.plan().setUpdateFields(true).apply();
    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();

    const matches = [...(settingsXml ?? '').matchAll(/updateFields/g)];
    expect(matches.length).toBe(1);
    expect(settingsXml).toContain('"true"');
  });
});

// ─────────── round-trip ──────────────────────────────────────────────────────

describe('setUpdateFields round-trip', () => {
  test('4 set(true) → toZip → fromZip → settings part 中读回 updateFields == "true"', async () => {
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-4', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    // fromZip 重加载并验证 settings.xml 中 updateFields 依然为 true
    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();
    expect(settingsXml).toContain('updateFields');
    expect(settingsXml).toContain('"true"');

    // 同时用 WordDocument 重加载验证整包合法（无抛错）
    const doc2 = await WordDocument.fromZip(zipBytes, { docId: 'uf-4-rt', mode: 'inline' });
    await doc2.dispose();
  });

  test('4b FlatOpc round-trip → set(true) → toFlatOpc → fromFlatOpc → updateFields == "true"', async () => {
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-4b', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const flatOpc = await doc.toFlatOpc();
    await doc.dispose();

    // FlatOpc 是 XML 字符串，直接检查包含 updateFields w:val="true"
    expect(flatOpc).toContain('updateFields');
    expect(flatOpc).toContain('"true"');

    // fromFlatOpc 重加载验证整包合法
    const doc2 = await WordDocument.fromFlatOpc(flatOpc, { docId: 'uf-4b-rt', mode: 'inline' });
    await doc2.dispose();
  });
});

// ─────────── 边界 ────────────────────────────────────────────────────────────

describe('setUpdateFields 边界', () => {
  test('6 settings 不存在 → 抛 UNSUPPORTED_OOXML_FEATURE', async () => {
    // makeFixture 不传 settingsXml → 不包含 settings.xml
    const buf = await makeFixture(undefined);
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-6', mode: 'inline' });

    await expect(doc.plan().setUpdateFields(true).apply()).rejects.toSatisfy(
      (e: unknown) => e instanceof WordError && e.code === 'UNSUPPORTED_OOXML_FEATURE',
    );

    await doc.dispose();
  });
});

// ─────────── 不回归 ──────────────────────────────────────────────────────────

describe('setUpdateFields 不回归', () => {
  test('7 现有 settings 子元素（zoom/defaultTabStop）不被破坏', async () => {
    // settingsXmlClean 包含 defaultTabStop 和 zoom
    const buf = await makeFixture(settingsXmlClean());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-7', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();

    // 原有元素依然存在
    expect(settingsXml).toContain('defaultTabStop');
    expect(settingsXml).toContain('zoom');

    // updateFields 也写入了
    expect(settingsXml).toContain('updateFields');
    expect(settingsXml).toContain('"true"');
  });

  test('7b 已有 updateFields="true" 时 set(true) 不破坏其他元素', async () => {
    // settingsXmlWithTrue 包含 defaultTabStop 和 updateFields=true
    const buf = await makeFixture(settingsXmlWithTrue());
    const doc = await WordDocument.fromZip(buf, { docId: 'uf-7b', mode: 'inline' });

    await doc.plan().setUpdateFields(true).apply();

    const zipBytes = await doc.toZip();
    await doc.dispose();

    const settingsXml = await readSettingsXml(zipBytes);
    expect(settingsXml).not.toBeNull();

    // defaultTabStop 还在
    expect(settingsXml).toContain('defaultTabStop');
    // updateFields 仍为 true
    expect(settingsXml).toContain('updateFields');
    expect(settingsXml).toContain('"true"');

    // 只有一个 updateFields
    const matches = [...(settingsXml ?? '').matchAll(/updateFields/g)];
    expect(matches.length).toBe(1);
  });
});
