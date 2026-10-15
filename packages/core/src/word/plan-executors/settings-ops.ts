/**
 * Settings executor —— task-core-update-fields。
 *
 * 职责:
 *   - `executeSetUpdateFields`：在 settings.xml 根 `<w:settings>` 写入/更新
 *     `<w:updateFields w:val="true|false"/>`，并把 settings part 标脏，
 *     确保 `toZip()` / `toFlatOpc()` 正确持久化。
 *
 * 设计要点:
 *   - 按 contentType 定位 settings part（不硬编码路径），镜像 style-queries.ts 的 findStylesPart。
 *   - settings part 不存在（非 B 场景边缘）：抛 UNSUPPORTED_OOXML_FEATURE，留后续 follow-up。
 *   - 幂等：已有 `<w:updateFields>` 则只改 `w:val`，不重复插入。
 *   - CT_Settings schema 顺序：updateFields 较靠后（在 defaultTabStop/zoom 类之后、
 *     footnotePr/endnotePr/compat/rsids 之前）。实现上追加到 `<w:settings>` 末尾；
 *     Word 对 CT_Settings 子元素顺序宽容，追加安全。若已有则原地改值不移位。
 *   - namespace：W_NS + setAttributeNS 写 `w:val`（与其他 ops 统一）。
 */
import { listParts, markDirty } from '../../pack/index.js';
import { W_NS } from '../namespaces.js';
import type { Op } from '../plan.js';
import { WordError } from '../word-error.js';
import type { ExecutorContext } from './executor-context.js';

// ─────────── ContentType 常量 ───────────────────────────────────────────────

const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

// ─────────── 内部工具 ────────────────────────────────────────────────────────

/**
 * 按 contentType 定位 settings part（不硬编码路径）。
 * 镜像 style-queries.ts 的 findStylesPart 模式。
 */
function findSettingsPart(ctx: ExecutorContext): { xmlDom: Document; _dirty: boolean } | undefined {
  for (const part of listParts(ctx.pkg)) {
    if (part.contentType === SETTINGS_CONTENT_TYPE && part.xmlDom) {
      return part as { xmlDom: Document; _dirty: boolean };
    }
  }
  return undefined;
}

// ─────────── executeSetUpdateFields ──────────────────────────────────────────

export function executeSetUpdateFields(
  ctx: ExecutorContext,
  op: Extract<Op, { type: 'setUpdateFields' }>,
): void {
  // ── Step 1: 定位 settings part ─────────────────────────────────────────────
  const settingsPart = findSettingsPart(ctx);
  if (!settingsPart) {
    // B 场景中 settings.xml 必然存在；此分支仅用于健壮性兜底。
    // 新建 settings part 需注册 ContentTypes Override + document.xml rels，
    // 基础设施成本较高，留后续 follow-up。
    throw new WordError(
      'UNSUPPORTED_OOXML_FEATURE',
      'setUpdateFields: settings.xml part not found in package. ' +
        'Creating a new settings part is not yet implemented (follow-up task).',
      ctx.docId,
      { opType: 'setUpdateFields' },
    );
  }

  const xmlDom = settingsPart.xmlDom;

  // ── Step 2: 找根 <w:settings> ─────────────────────────────────────────────
  const settingsRoot = xmlDom.documentElement;
  if (!settingsRoot) {
    throw new WordError(
      'PACKAGE_CORRUPT',
      'setUpdateFields: settings.xml has no documentElement',
      ctx.docId,
      { opType: 'setUpdateFields' },
    );
  }

  // ── Step 3: 幂等 —— 查已有 <w:updateFields> ──────────────────────────────
  const existing = findFirstChildW(settingsRoot, 'updateFields');
  const valStr = op.enabled ? 'true' : 'false';

  if (existing) {
    // 已有：只改 w:val，不移位，不重复插入（幂等核心）
    existing.setAttributeNS(W_NS, 'w:val', valStr);
  } else {
    // 不存在：新建并追加到 <w:settings> 末尾
    // CT_Settings 顺序：updateFields 在较靠后位置，追加到末尾符合规范且 Word 宽容
    const newEl = xmlDom.createElementNS(W_NS, 'w:updateFields');
    newEl.setAttributeNS(W_NS, 'w:val', valStr);
    settingsRoot.appendChild(newEl);
  }

  // ── Step 4: 标脏 ──────────────────────────────────────────────────────────
  // settings part 不是 story part，直接通过 pkg 找到 Part 对象调 markDirty。
  for (const part of listParts(ctx.pkg)) {
    if (part.contentType === SETTINGS_CONTENT_TYPE && part.xmlDom) {
      markDirty(part);
      break;
    }
  }
}

// ─────────── 内部工具 ────────────────────────────────────────────────────────

/** 找 el 的第一个 W_NS 直属子元素（localName 匹配）。 */
function findFirstChildW(el: Element, localName: string): Element | null {
  const children = el.childNodes;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (
      c &&
      c.nodeType === 1 &&
      (c as Element).namespaceURI === W_NS &&
      (c as Element).localName === localName
    ) {
      return c as Element;
    }
  }
  return null;
}
