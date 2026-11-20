import {
  readExcelEmbeddedObjectPreviews,
  type ExcelDrawingAnchorModel,
  type ExcelDrawingObjectModel,
} from '@coco-vexel/coco-office/excel';
import { getPart, type Package } from '@coco-vexel/coco-office/pack';
import type { SpreadsheetRect } from './types.js';

const SVG_NS = 'http://www.w3.org/2000/svg';
const BROWSER_IMAGE_TYPES = new Set([
  'image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp', 'image/svg+xml',
]);

export class SpreadsheetDrawingPreviews {
  private readonly urls = new Map<string, string>();
  private readonly embedded = new Map<string, ExcelDrawingAnchorModel[]>();

  constructor(private readonly pkg: Package) {}

  embeddedOnSheet(path: string): ExcelDrawingAnchorModel[] {
    const cached = this.embedded.get(path);
    if (cached) return cached;
    const previews = readExcelEmbeddedObjectPreviews(this.pkg, path);
    this.embedded.set(path, previews);
    return previews;
  }

  render(object: ExcelDrawingObjectModel, rect: SpreadsheetRect, embedded: boolean): HTMLElement | SVGElement | null {
    if (object.hidden) return null;
    if (object.kind === 'shape') return renderShape(object, rect);
    if (object.kind !== 'image' || object.targetMode === 'External') return null;
    const url = object.targetPath && this.imageUrl(object.targetPath);
    if (!url) return embedded ? embeddedIcon(object.name) : null;
    const image = document.createElement('img');
    image.alt = object.name ?? '';
    image.draggable = false;
    if (embedded) {
      image.addEventListener('error', () => {
        const fallback = embeddedIcon(object.name);
        fallback.style.cssText += image.style.cssText;
        image.replaceWith(fallback);
      }, { once: true });
    }
    image.src = url;
    return image;
  }

  dispose(): void {
    for (const url of this.urls.values()) URL.revokeObjectURL(url);
    this.urls.clear();
    this.embedded.clear();
  }

  private imageUrl(path: string): string | undefined {
    const cached = this.urls.get(path);
    if (cached) return cached;
    const part = getPart(this.pkg, path);
    const type = part?.contentType.toLowerCase();
    if (!part?.bytes || !type || !BROWSER_IMAGE_TYPES.has(type)) return undefined;
    const url = URL.createObjectURL(new Blob([new Uint8Array(part.bytes)], { type }));
    this.urls.set(path, url);
    return url;
  }
}

function renderShape(object: ExcelDrawingObjectModel, { width, height }: SpreadsheetRect): SVGElement | null {
  let geometry: SVGElement;
  switch (object.presetGeometry) {
    case 'rightArrow':
    case 'leftArrow': {
      const shaft = width - Math.min(width, height) / 2;
      geometry = document.createElementNS(SVG_NS, 'path');
      geometry.setAttribute('d', `M 0 ${height / 4} H ${shaft} V 0 L ${width} ${height / 2} L ${shaft} ${height} V ${height * 0.75} H 0 Z`);
      if (object.presetGeometry === 'leftArrow') geometry.setAttribute('transform', `translate(${width} 0) scale(-1 1)`);
      break;
    }
    case 'rect':
      geometry = document.createElementNS(SVG_NS, 'rect');
      geometry.setAttribute('width', String(width));
      geometry.setAttribute('height', String(height));
      break;
    case 'ellipse':
      geometry = document.createElementNS(SVG_NS, 'ellipse');
      geometry.setAttribute('cx', String(width / 2));
      geometry.setAttribute('cy', String(height / 2));
      geometry.setAttribute('rx', String(width / 2));
      geometry.setAttribute('ry', String(height / 2));
      break;
    default:
      return null;
  }
  geometry.setAttribute('fill', object.fillColor ?? '#4472C4');
  geometry.setAttribute('stroke', object.lineColor ?? '#2F528F');
  geometry.setAttribute('stroke-width', String(Math.max(0, (object.lineWidth ?? 12_700) / 9_525)));
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
  svg.setAttribute('aria-hidden', 'true');
  svg.appendChild(geometry);
  return svg;
}

function embeddedIcon(name: string | undefined): HTMLElement {
  const icon = document.createElement('div');
  Object.assign(icon.style, {
    display: 'flex', alignItems: 'center', justifyContent: 'center',
    flexDirection: 'column', gap: '3px', font: '10px Arial, sans-serif', color: '#1f2937',
  });
  const symbol = document.createElement('span');
  Object.assign(symbol.style, {
    background: '#107c41', color: '#fff', borderRadius: '2px', padding: '2px 5px', font: 'bold 12px Arial, sans-serif',
  });
  symbol.textContent = /\.xlsx?$/i.test(name ?? '') ? 'X' : '↗';
  const label = document.createElement('span');
  Object.assign(label.style, {
    maxWidth: '100%', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
  });
  label.textContent = name ?? '嵌入文件';
  icon.append(symbol, label);
  return icon;
}
