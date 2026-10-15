// Synthetic examples: no external document, font, image or Office installation.
import { WordDocument } from '@coco-vexel/coco-office/word';
import { ExcelWorkbook } from '@coco-vexel/coco-office/excel';
import { Presentation } from '@coco-vexel/coco-office/ppt';

export async function createSample(kind) {
  if (kind === 'docx') {
    const doc = await WordDocument.createBlank({ docId: 'sample-word', mode: 'inline' });
    try {
      const { paragraphs } = await doc.paragraphs.list({ pagination: { mode: 'all' } });
      await paragraphs[0].replaceWith('Coco Office sample');
      await doc.appendParagraph('Open this document, edit this sentence, and export a copy.');
      return await doc.toZip();
    } finally {
      await doc.dispose();
    }
  }
  if (kind === 'xlsx') {
    const workbook = await ExcelWorkbook.fromModel(
      {
        sharedStrings: [],
        definedNames: [],
        styles: { numberFormats: [], fonts: [], fills: [], borders: [], cellFormats: [] },
        sheets: [
          {
            id: '1',
            name: 'Sample',
            dimension: 'A1:B3',
            sheetViews: [],
            columns: [],
            rows: [],
            mergedRanges: [],
            hyperlinks: [],
            tables: [],
            dataValidations: [],
            conditionalFormats: [],
            comments: [],
            threadedComments: [],
            drawings: [],
            charts: [],
            images: [],
            cells: [],
          },
        ],
      },
      { workbookId: 'sample-excel', mode: 'inline' },
    );
    try {
      await workbook.worksheets
        .getItem('Sample')
        .getRange('A1:B3')
        .setValues([
          ['Item', 'Quantity'],
          ['Apples', 12],
          ['Oranges', 8],
        ]);
      return await workbook.toZip();
    } finally {
      await workbook.dispose();
    }
  }
  if (kind === 'pptx') {
    const presentation = await Presentation.fromModel({
      id: 'sample-ppt',
      slideSize: { width: 9144000, height: 5143500 },
      slides: [
        {
          id: '256',
          index: 0,
          shapes: [
            {
              id: '2',
              name: 'Title',
              kind: 'textBox',
              geometry: { left: 600000, top: 600000, width: 7500000, height: 1200000 },
              text: { text: 'Coco Office sample' },
            },
          ],
        },
      ],
    });
    try {
      return await presentation.toZip();
    } finally {
      await presentation.dispose();
    }
  }
  throw new Error(`Unsupported sample: ${kind}`);
}
