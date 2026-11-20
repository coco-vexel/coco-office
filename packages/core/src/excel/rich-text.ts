import { Document as XmlDocument, serializeToWellFormedString } from 'slimdom';
import { excelColorToOoxmlArgb } from './color.js';
import type { ExcelCellModel, ExcelRichTextRun } from './model.js';
import { SS_NS } from './namespaces.js';
import { parseFontProperties, type ThemeColors } from './styles-reader.js';

export interface ExcelStringContent {
  text: string;
  richText?: ExcelRichTextRun[];
}

export function readExcelStringContent(element: Element, themeColors: ThemeColors): ExcelStringContent {
  const runs: ExcelRichTextRun[] = [];
  let rich = false;
  for (const child of Array.from(element.children)) {
    if (child.namespaceURI !== SS_NS) continue;
    if (child.localName === 't') {
      runs.push({ text: child.textContent ?? '' });
    } else if (child.localName === 'r') {
      const text = child.getElementsByTagNameNS(SS_NS, 't')[0]?.textContent ?? '';
      const properties = child.getElementsByTagNameNS(SS_NS, 'rPr')[0];
      runs.push(properties ? { text, font: parseFontProperties(properties, themeColors, 'rFont') } : { text });
      rich = true;
    }
  }
  const text = runs.map((run) => run.text).join('');
  return rich ? { text, richText: runs } : { text };
}

export function cloneExcelRichText(runs: readonly ExcelRichTextRun[]): ExcelRichTextRun[] {
  return runs.map((run) => {
    if (!run.font) return { text: run.text };
    const font = { ...run.font };
    if (font.color) font.color = { ...font.color };
    return { text: run.text, font };
  });
}

export function cloneExcelCell(cell: ExcelCellModel): ExcelCellModel {
  return cell.richText ? { ...cell, richText: cloneExcelRichText(cell.richText) } : { ...cell };
}

export function excelCellRichText(cell: ExcelCellModel): ExcelRichTextRun[] | undefined {
  const runs = cell.richText;
  return cell.kind === 'string' && runs?.length && runs.map((run) => run.text).join('') === cell.value
    ? runs : undefined;
}

export function serializeExcelRichText(runs: readonly ExcelRichTextRun[]): string {
  const doc = new XmlDocument();
  const inline = createExcelRichTextElement(doc as unknown as Document, runs);
  return serializeToWellFormedString(inline as unknown as Parameters<typeof serializeToWellFormedString>[0]);
}

export function createExcelRichTextElement(
  document: Document,
  runs: readonly ExcelRichTextRun[],
): Element {
  const inline = document.createElementNS(SS_NS, 'is');
  for (const { text, font } of runs) {
    const run = document.createElementNS(SS_NS, 'r');
    if (font) {
      const properties = document.createElementNS(SS_NS, 'rPr');
      for (const [name, value] of [
        ['rFont', font.name], ['charset', font.charset], ['family', font.family],
        ['b', font.bold], ['i', font.italic], ['u', font.underline], ['sz', font.size],
        ['scheme', font.scheme],
      ] as const) {
        if (value === undefined) continue;
        const child = document.createElementNS(SS_NS, name);
        if (value !== true) child.setAttribute('val', value === false ? (name === 'u' ? 'none' : '0') : String(value));
        properties.appendChild(child);
      }
      if (font.color) {
        const color = document.createElementNS(SS_NS, 'color');
        color.setAttribute('rgb', excelColorToOoxmlArgb(font.color));
        properties.appendChild(color);
      }
      run.appendChild(properties);
    }
    const value = document.createElementNS(SS_NS, 't');
    if (/^\s|\s$/.test(text)) {
      value.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    }
    value.textContent = text;
    run.appendChild(value);
    inline.appendChild(run);
  }
  return inline;
}
