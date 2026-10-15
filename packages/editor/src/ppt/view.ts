import { Presentation, type PptPresentationModel } from '@coco-vexel/coco-office/ppt';
import { buildPptPresentationScene, type PptPresentationScene } from './scene.js';

export interface PptEditorImageResource {
  url: string;
  contentType: string;
}

export interface PptEditorImagePreviewRequest {
  path: string;
  contentType: string;
  bytes: Uint8Array;
}

export interface PptEditorFontResource {
  url: string;
  contentType: string;
  family: string;
  weight?: number;
  style?: 'normal' | 'italic';
}

export interface PptEditorView {
  presentation: Presentation;
  model: PptPresentationModel;
  scene: PptPresentationScene;
  images: Record<string, PptEditorImageResource>;
  fonts: PptEditorFontResource[];
}

export interface PptEditorViewOptions {
  presentationId?: string;
  resolveImagePreview?: (
    image: PptEditorImagePreviewRequest,
  ) => Promise<PptEditorImageResource | undefined> | PptEditorImageResource | undefined;
}

export async function buildPptEditorViewFromPptx(
  bytes: Uint8Array | ArrayBuffer,
  options: PptEditorViewOptions = {},
): Promise<PptEditorView> {
  const presentation = await Presentation.fromZip(bytes, {
    presentationId: options.presentationId,
    mode: 'inline',
  });
  return buildPptEditorViewFromPresentation(presentation, options);
}

export async function buildPptEditorViewFromModel(
  model: PptPresentationModel,
  options: PptEditorViewOptions = {},
): Promise<PptEditorView> {
  const presentation = await Presentation.fromModel(model, {
    presentationId: options.presentationId ?? model.id,
    mode: 'inline',
  });
  return buildPptEditorViewFromPresentation(presentation, options);
}

export async function buildPptEditorViewFromPresentation(
  presentation: Presentation,
  options: PptEditorViewOptions = {},
): Promise<PptEditorView> {
  const [model, images, fonts] = await Promise.all([
    presentation.toModel(),
    buildPptEditorImageResources(presentation, options),
    buildPptEditorFontResources(presentation),
  ]);
  return {
    presentation,
    model,
    scene: buildPptPresentationScene(model),
    images,
    fonts,
  };
}

async function buildPptEditorImageResources(
  presentation: Presentation,
  options: PptEditorViewOptions,
): Promise<Record<string, PptEditorImageResource>> {
  if (typeof Blob === 'undefined' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
    return {};
  }
  const images = await presentation.exportImages();
  const entries = await Promise.all(
    images.map(async (image) => {
      const preview = await options.resolveImagePreview?.({
        path: image.path,
        contentType: image.contentType,
        bytes: image.bytes,
      });
      return [
        image.path,
        preview ?? {
          url: URL.createObjectURL(new Blob([arrayBufferFromBytes(image.bytes)], { type: image.contentType })),
          contentType: image.contentType,
        },
      ] as const;
    }),
  );
  return Object.fromEntries(entries);
}

async function buildPptEditorFontResources(presentation: Presentation): Promise<PptEditorFontResource[]> {
  if (typeof Blob === 'undefined' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
    return [];
  }
  const fonts = await presentation.exportFonts();
  return fonts.map((font) => ({
    url: URL.createObjectURL(new Blob([arrayBufferFromBytes(font.bytes)], { type: font.contentType })),
    contentType: font.contentType,
    family: font.family,
    ...(font.weight ? { weight: font.weight } : {}),
    ...(font.style ? { style: font.style } : {}),
  }));
}

function arrayBufferFromBytes(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}
