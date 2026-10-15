import type {
  PptShapeGeometry,
  PptShapeKind,
  PptShapeStyle,
  PptHyperlinkWriteOptions,
  PptCommentAddOptions,
  PptChartData,
  PptImageStyle,
  PptTableCellStyle,
  PptTextBodyProperties,
  PptTextParagraphWrite,
  PptTextStyle,
} from './model.js';

export type PptMutation =
  | {
      type: 'shape.add';
      slideId: string;
      shapeId: string;
      kind: PptShapeKind;
      imageRef?: string;
    }
  | {
      type: 'shape.delete';
      slideId: string;
      shapeId: string;
    }
  | {
      type: 'shape.name.set';
      slideId: string;
      shapeId: string;
      name: string;
    }
  | {
      type: 'shape.zOrder.set';
      slideId: string;
      shapeId: string;
      position: 'front' | 'back';
    }
  | {
      type: 'shape.geometry.set';
      slideId: string;
      shapeId: string;
      geometry: Partial<PptShapeGeometry>;
    }
  | {
      type: 'shape.style.set';
      slideId: string;
      shapeId: string;
      style: Partial<PptShapeStyle>;
    }
  | {
      type: 'shape.hyperlink.set';
      slideId: string;
      shapeId: string;
      hyperlink: PptHyperlinkWriteOptions;
    }
  | {
      type: 'shape.hyperlink.clear';
      slideId: string;
      shapeId: string;
    }
  | {
      type: 'shape.text.set';
      slideId: string;
      shapeId: string;
      text: string;
    }
  | {
      type: 'shape.text.style.set';
      slideId: string;
      shapeId: string;
      style: Partial<PptTextStyle>;
    }
  | {
      type: 'shape.text.body.set';
      slideId: string;
      shapeId: string;
      body: Partial<PptTextBodyProperties>;
    }
  | {
      type: 'shape.text.paragraphs.set';
      slideId: string;
      shapeId: string;
      paragraphs: PptTextParagraphWrite[];
    }
  | {
      type: 'shape.text.run.style.set';
      slideId: string;
      shapeId: string;
      runIndex: number;
      style: Partial<PptTextStyle>;
    }
  | {
      type: 'shape.text.run.hyperlink.set';
      slideId: string;
      shapeId: string;
      runIndex: number;
      hyperlink: PptHyperlinkWriteOptions;
    }
  | {
      type: 'shape.text.run.hyperlink.clear';
      slideId: string;
      shapeId: string;
      runIndex: number;
    }
  | {
      type: 'shape.text.replace';
      slideId: string;
      shapeId: string;
      search: string;
      replacement: string;
      count: number;
    }
  | {
      type: 'slide.notes.set';
      slideId: string;
      text: string;
    }
  | {
      type: 'slide.notes.clear';
      slideId: string;
    }
  | {
      type: 'slide.comment.add';
      slideId: string;
      comment: PptCommentAddOptions;
    }
  | {
      type: 'slide.comment.delete';
      slideId: string;
      commentId: string;
    }
  | {
      type: 'slide.comments.clear';
      slideId: string;
    }
  | {
      type: 'shape.chart.data.set';
      slideId: string;
      shapeId: string;
      data: PptChartData;
    }
  | {
      type: 'shape.image.replace';
      slideId: string;
      shapeId: string;
      imageRef: string;
    }
  | {
      type: 'shape.image.style.set';
      slideId: string;
      shapeId: string;
      style: Partial<PptImageStyle>;
    }
  | {
      type: 'table.cell.text.set';
      slideId: string;
      shapeId: string;
      row: number;
      column: number;
      text: string;
    }
  | {
      type: 'table.cell.style.set';
      slideId: string;
      shapeId: string;
      row: number;
      column: number;
      style: Partial<PptTableCellStyle>;
    }
  | {
      type: 'table.cell.text.paragraphs.set';
      slideId: string;
      shapeId: string;
      row: number;
      column: number;
      paragraphs: PptTextParagraphWrite[];
    }
  | {
      type: 'table.cell.text.run.style.set';
      slideId: string;
      shapeId: string;
      row: number;
      column: number;
      runIndex: number;
      style: Partial<PptTextStyle>;
    }
  | {
      type: 'table.values.set';
      slideId: string;
      shapeId: string;
      values: string[][];
    }
  | {
      type: 'table.row.add';
      slideId: string;
      shapeId: string;
      index: number;
      values: string[];
    }
  | {
      type: 'table.row.delete';
      slideId: string;
      shapeId: string;
      index: number;
    }
  | {
      type: 'table.column.add';
      slideId: string;
      shapeId: string;
      index: number;
      values: string[];
    }
  | {
      type: 'table.column.delete';
      slideId: string;
      shapeId: string;
      index: number;
    };

export class PptMutationLog {
  private readonly items: PptMutation[] = [];

  record(mutation: PptMutation): void {
    this.items.push(cloneMutation(mutation));
  }

  list(): PptMutation[] {
    return this.items.map(cloneMutation);
  }

  clear(): void {
    this.items.length = 0;
  }
}

function cloneMutation(mutation: PptMutation): PptMutation {
  return JSON.parse(JSON.stringify(mutation)) as PptMutation;
}
