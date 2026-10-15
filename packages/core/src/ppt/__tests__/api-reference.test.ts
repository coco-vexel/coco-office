import { readdirSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = dirname(fileURLToPath(import.meta.url));
const runtimePath = resolve(here, '../runtime.ts');
const docsRoot = resolve(here, '../../../../../docs/core/ppt');
const apiReferencePath = resolve(here, '../../../../../docs/core/ppt/api-reference.md');
const styleCapabilitiesPath = resolve(here, '../../../../../docs/core/ppt/style-capabilities.md');
const sdkApiPath = resolve(here, '../../../../../docs/core/ppt-sdk-api.md');

const runtimeSource = readFileSync(runtimePath, 'utf8');
const apiReference = readFileSync(apiReferencePath, 'utf8');
const styleCapabilities = readFileSync(styleCapabilitiesPath, 'utf8');
const sdkApi = readFileSync(sdkApiPath, 'utf8');

describe('PPT SDK API reference', () => {
  it('lists every public runtime class', () => {
    expect([...extractExportedClasses(runtimeSource)].sort()).toEqual([
      'Presentation',
      'PresentationThemeCollection',
      'Shape',
      'ShapeChartApi',
      'ShapeCollection',
      'ShapeGroupApi',
      'ShapeImageApi',
      'ShapeTextApi',
      'Slide',
      'SlideCollection',
      'SlideCommentsApi',
      'SlideNotesApi',
      'TableApi',
      'TableCell',
    ]);
  });

  it('keeps the method index aligned with public runtime methods', () => {
    const documentedMethodNames = methodNames(apiReference);
    const missing = [...extractPublicRuntimeMethods(runtimeSource)].filter(
      (entry) => !documentedMethodNames.has(entry.method),
    );

    expect(missing).toEqual([]);
  });

  it('documents the core-only style capability boundary', () => {
    for (const expected of [
      'setTextStyle',
      'setRunTextStyle',
      'setFillColor',
      'setLineStyle',
      'cell.setStyle',
      'setCrop',
      'setFillRect',
      'setOpacity',
      'PptImageStyle',
      'fillRect',
      'opacity',
      'lineOpacity',
      'lineCap',
      'lineDash',
      'shadow',
      'fontFamily',
      'underline',
      'strike',
      'baseline',
      'highlight',
      'lineHeight',
      'characterSpacing',
      'bulletColor',
      'bulletSize',
      'spaceBefore',
      'anchor',
      'wrap',
      'insetLeft',
      'read-write',
      'read-only',
      'planned',
      'out-of-scope',
    ]) {
      expect(styleCapabilities).toContain(expected);
    }
  });

  it('documents high-risk SDK fields that are easy to drift from implementation', () => {
    const docsByName = new Map(sdkOperationDocs());

    expect(apiReference).toContain('presentation.exportFonts()');
    expect(apiReference).toContain('package media bytes');
    expect(apiReference).toContain('`smartArt`, or `unknown`');
    expect(apiReference).toContain("PptShapeModel['raw']");

    expect(docsByName.get('domains/presentation.md')).toContain('Export Fonts');
    expect(docsByName.get('domains/presentation.md')).toContain('audio, or video parts');
    expect(docsByName.get('domains/images.md')).toContain('non-picture media parts');
    expect(docsByName.get('domains/shapes.md')).toContain('line dash');
    expect(docsByName.get('domains/shapes.md')).toContain('flipHorizontal');
    expect(docsByName.get('domains/shapes.md')).toContain("kind: 'unknown'");

    for (const expected of [
      'fonts?: PptFontModel[]',
      'flipHorizontal?: boolean',
      'flipVertical?: boolean',
      'lineCap?: string',
      'lineDash?: string',
      'presetAdjustments?: Record<string, number>',
      'customGeometry?: PptCustomGeometryModel',
      'groupTransform?: PptGroupTransformModel',
      'data?: PptChartData',
      'pointColors?: string[]',
      'tickLabelPosition?: string',
      'columnWidths?: number[]',
      'rowHeights?: number[]',
      'interface PptImageStyle',
      'interface PptTextBodyProperties',
    ]) {
      expect(docsByName.get('model.md')).toContain(expected);
    }
  });

  it('keeps the public docs aligned with current PPT API names', () => {
    const publicDocs = sdkOperationDocs()
      .map(([, markdown]) => markdown)
      .join('\n');
    for (const expected of [
      'shapes.addImage',
      'shapes.addShape',
      'image.bytes',
      'image.setFillRect',
      'cell.setStyle',
    ]) {
      expect(publicDocs).toContain(expected);
    }
    for (const stale of [
      'shapes.addPicture',
      'shapes.addGeometricShape',
      'PptPictureAddOptions',
      'image.getBytes',
      'PptImageBytes',
      'cell.fill',
      'PptFillFormat',
    ]) {
      expect(publicDocs).not.toContain(stale);
    }
  });

  it('keeps the compatibility API entry focused on SDK docs', () => {
    expect(sdkApi).toContain('public API index');
    expect(sdkApi).toContain('domain retrieval map');
    for (const [name, markdown] of sdkOperationDocs()) {
      expect(markdown, name).not.toContain('CodePlan');
    }
  });
});

function sdkOperationDocs(): Array<[string, string]> {
  return [
    ['ppt-sdk-api.md', sdkApi],
    ...markdownFiles(docsRoot)
      .filter((name) => !['host-integration.md', 'sdk-plan.md'].includes(name))
      .map((name) => [name, readFileSync(resolve(docsRoot, name), 'utf8')] as [string, string]),
  ];
}

function markdownFiles(dir: string, prefix = ''): string[] {
  const entries = readdirSync(dir, { withFileTypes: true });
  return entries.flatMap((entry) => {
    const name = prefix ? `${prefix}/${entry.name}` : entry.name;
    const path = resolve(dir, entry.name);
    if (entry.isDirectory()) return markdownFiles(path, name);
    return entry.isFile() && entry.name.endsWith('.md') ? [name] : [];
  });
}

interface RuntimeMethodEntry {
  className: string;
  method: string;
}

function extractExportedClasses(source: string): Set<string> {
  return new Set([...source.matchAll(/^export class (\w+)/gm)].map((match) => match[1]!));
}

function extractPublicRuntimeMethods(source: string): RuntimeMethodEntry[] {
  const entries: RuntimeMethodEntry[] = [];
  const classMatches = [...source.matchAll(/^export class (\w+)/gm)];

  for (let index = 0; index < classMatches.length; index++) {
    const match = classMatches[index]!;
    const className = match[1]!;
    const start = match.index ?? 0;
    const end = classMatches[index + 1]?.index ?? source.length;
    const body = source.slice(start, end);
    const methodNames = [
      ...body.matchAll(/^\s*(?:static\s+)?async\s+(\w+)\(/gm),
      ...body.matchAll(/^\s*(getItem|getItemAt|getCell)\(/gm),
    ].map((methodMatch) => methodMatch[1]!);

    for (const method of methodNames) {
      if (method === 'executeCode') continue;
      entries.push({ className, method });
    }
  }

  return entries;
}

function methodNames(markdown: string): Set<string> {
  return new Set(
    [...markdown.matchAll(/\.?([A-Za-z]\w*)\s*\(/g)]
      .map((match) => match[1]!)
      .filter((name) => name !== 'Promise'),
  );
}
