/**
 * Package / Part 数据模型 —— CocoOffice 的 OOXML 统一内存表示。
 *
 * 设计背景:docs/01 §3.2(Package 抽象)、docs/03 §13.4(dirty 追踪)。
 *
 * 一个 Part 要么是 XML 类(`xmlDom` 有值、`bytes` 无值),
 * 要么是二进制类(`bytes` 有值、`xmlDom` 无值)。
 * 这种"择一"约束由 `createXmlPart` / `createBinaryPart` 工厂函数在构造期守住。
 */

/** OOXML Package 内的单一 Part:路径 + MIME contentType + payload + dirty 元信息 */
export interface Part {
  /** Part 路径,相对包根,**无前导斜杠**。例:'word/document.xml' */
  readonly path: string;
  /**
   * MIME content type。例:'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
   * 二进制 part 如图片:'image/png'
   */
  readonly contentType: string;
  /**
   * XML 类 part 的 DOM(slimdom Document 兼容 lib.dom Document)。
   * 与 `bytes` 互斥 —— 仅择一赋值。
   */
  xmlDom?: Document;
  /**
   * 二进制 part 的字节(如图片、字体)。
   * 与 `xmlDom` 互斥 —— 仅择一赋值。
   */
  bytes?: Uint8Array;
  /**
   * 装载时保留的原始字节。`_dirty === false` 时,`toZip` / `toFlatOpc` 直接复用此字节,零 serialize。
   * 二进制 part 初始化时通常 === `bytes`。
   * 装载后首次被改动(markDirty 置 true)前为只读快照。
   */
  _originalBytes?: Uint8Array;
  /**
   * 修改标记。
   * - `false`:自装载以来未变更,输出时可复用 `_originalBytes`
   * - `true`:已被改过,输出时需要重新序列化
   *
   * 由 Plan.apply 或外部工具通过 `markDirty(part)` 置 true;读操作不应改变它。
   */
  _dirty: boolean;
}

/**
 * 归一化后的 Package:路径到 Part 的 Map。
 *
 * - 外部 API(zip / Flat OPC)经 adapter 归一到 Package,内部只面对这个形态
 * - 同一 Package 内 path 唯一,由 Map 天然保证
 */
export interface Package {
  parts: Map<string, Part>;
}
