import type {
  Presentation,
  PresentationThemeCollection,
  Shape,
  ShapeChartApi,
  ShapeCollection,
  ShapeGroupApi,
  ShapeImageApi,
  ShapeTextApi,
  Slide,
  SlideCollection,
  SlideCommentsApi,
  SlideNotesApi,
  TableApi,
  TableCell,
} from './runtime.js';

/** Receivers exposed inside Presentation CodePlan scripts. */
export interface PptScriptApiReceiverMap {
  presentation: Presentation;
  'presentation.slides': SlideCollection;
  'presentation.themes': PresentationThemeCollection;
  slide: Slide;
  'slide.shapes': ShapeCollection;
  'slide.notes': SlideNotesApi;
  'slide.comments': SlideCommentsApi;
  shape: Shape;
  'shape.text': ShapeTextApi;
  'shape.image': ShapeImageApi;
  'shape.table': TableApi;
  'shape.chart': ShapeChartApi;
  'shape.group': ShapeGroupApi;
  cell: TableCell;
}

export type PptScriptApiReceiver = keyof PptScriptApiReceiverMap;

type MethodName<T> = {
  [K in keyof T]-?: T[K] extends (...args: never[]) => unknown ? K : never;
}[keyof T] &
  string;

/** Type-safe receiver-qualified method path for the asynchronous PPT script runtime. */
export type PptScriptApiPath = {
  [R in PptScriptApiReceiver]: `${R}.${MethodName<PptScriptApiReceiverMap[R]>}`;
}[PptScriptApiReceiver];
