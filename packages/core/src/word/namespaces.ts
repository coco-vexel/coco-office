/**
 * OOXML 常用命名空间 URI 常量。
 *
 * 所有 XPath / DOM 查询必须按 URI 比较(不依赖前缀)—— serializer 可能改前缀,URI 不变。
 * 来源:docs/01 附录 A。
 */

/** Word 正文命名空间 */
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** 关系(.rels 文件 / r:id 引用) */
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** DrawingML(Word / Excel / PPT 共用) */
export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** DrawingML wordprocessing 容器(wp:inline / wp:anchor / wp:extent / wp:wrap*) */
export const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

/** DrawingML picture(pic:pic) */
export const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/** DrawingML chart(c:chart 引用) */
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** DrawingML diagram(SmartArt;dgm:relIds) */
export const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** Word 2010+ wordprocessing shape(wps:wsp / wps:txbx) */
export const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

/** Word 2010+ wordprocessing group(wpg:wgp) */
export const WPG_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

/** Word 2010+ wordprocessing canvas(wpc:wpc) */
export const WPC_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';

/** Markup Compatibility(Choice / Fallback) */
export const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** VML(旧版形状降级；v:line / v:rect / v:shape，多见于 mc:Fallback) */
export const VML_NS = 'urn:schemas-microsoft-com:vml';

/** Word 2010+ 扩展(w14:paraId / w14:textId) */
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** Flat OPC 包格式 */
export const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';

/** xml:space 等 xml 标准命名空间 */
export const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/** OMML 数学公式命名空间（Word「插入→公式」生成的 m:oMath / m:oMathPara） */
export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
