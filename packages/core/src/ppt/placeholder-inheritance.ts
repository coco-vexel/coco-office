import type {
  PptPlaceholderModel,
  PptShapeGeometry,
  PptShapeModel,
  PptShapeStyle,
  PptSlideLayoutModel,
  PptTextBodyProperties,
  PptTextStyle,
} from './model.js';

export interface PptEffectivePlaceholderProperties {
  placeholder?: PptPlaceholderModel;
  geometry: PptShapeGeometry;
  style?: PptShapeStyle;
  textStyle?: PptTextStyle;
  textBody?: PptTextBodyProperties;
}

export function resolvePptEffectivePlaceholderProperties(
  shape: PptShapeModel,
  layout?: PptSlideLayoutModel,
): PptEffectivePlaceholderProperties {
  if (!shape.placeholder) {
    const textStyle = optionalMergedObject(layout?.master?.textStyles?.other, shape.text?.style);
    return {
      geometry: { ...shape.geometry },
      ...(shape.style ? { style: { ...shape.style } } : {}),
      ...(textStyle ? { textStyle } : {}),
      ...(shape.text?.body ? { textBody: { ...shape.text.body } } : {}),
    };
  }

  const layoutShape = findMatchingPptPlaceholder(layout?.shapes, shape);
  const masterShape = findMatchingPptPlaceholder(layout?.master?.shapes, layoutShape ?? shape);
  const placeholder = mergeDefinedObjects(
    masterShape?.placeholder,
    layoutShape?.placeholder,
    shape.placeholder,
  );
  const masterTextStyle = layout?.master?.textStyles?.[masterTextStyleKind(placeholder?.type)];

  return {
    placeholder,
    geometry: mergeDefinedObjects(masterShape?.geometry, layoutShape?.geometry, shape.geometry),
    style: optionalMergedObject(masterShape?.style, layoutShape?.style, shape.style),
    textStyle: optionalMergedObject(
      masterTextStyle,
      masterShape?.text?.style,
      layoutShape?.text?.style,
      shape.text?.style,
    ),
    textBody: optionalMergedObject(
      masterShape?.text?.body,
      layoutShape?.text?.body,
      shape.text?.body,
    ),
  };
}

export function findMatchingPptPlaceholder(
  candidates: readonly PptShapeModel[] | undefined,
  shape: PptShapeModel,
): PptShapeModel | undefined {
  if (!shape.placeholder || !candidates?.length) return undefined;
  const index = shape.placeholder.index;
  if (index !== undefined) {
    const indexed = candidates.find(
      (candidate) =>
        candidate.placeholder?.index !== undefined && candidate.placeholder.index === index,
    );
    if (indexed) return indexed;
  }
  const type = shape.placeholder.type;
  if (type !== undefined) {
    return candidates.find((candidate) => candidate.placeholder?.type === type);
  }
  return undefined;
}

function masterTextStyleKind(type: string | undefined): 'title' | 'body' | 'other' {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === 'body' || type === 'subTitle' || type === 'obj') return 'body';
  return 'other';
}

function optionalMergedObject<T extends object>(...values: Array<T | undefined>): T | undefined {
  const merged = mergeDefinedObjects(...values);
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function mergeDefinedObjects<T extends object>(...values: Array<T | undefined>): T {
  return Object.assign({}, ...values.filter((value): value is T => value !== undefined)) as T;
}
