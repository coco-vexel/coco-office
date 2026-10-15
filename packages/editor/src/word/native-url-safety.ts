const RASTER_IMAGE_DATA_URL = /^data:(image\/(?:png|jpeg|gif|webp));base64,([a-z0-9+/]*={0,2})$/i;

const SAFE_HYPERLINK_PROTOCOLS = new Set(['http:', 'https:', 'mailto:', 'tel:']);

export function safeEmbeddedRasterImageUrl(value: string): string | null {
  const trimmed = value.trim();
  return RASTER_IMAGE_DATA_URL.test(trimmed) ? trimmed : null;
}

export function embeddedRasterImageBlob(value: string): Blob {
  const safeUrl = safeEmbeddedRasterImageUrl(value);
  const match = safeUrl ? RASTER_IMAGE_DATA_URL.exec(safeUrl) : null;
  if (!match) throw new Error('Only embedded PNG, JPEG, GIF, or WebP images are supported');
  const binary = atob(match[2]!);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  return new Blob([bytes], { type: match[1]!.toLowerCase() });
}

export function safeRenderedHyperlinkHref(
  target: string | undefined,
  anchor: string | undefined,
): string {
  if (target) {
    try {
      const parsed = new URL(target.trim());
      if (SAFE_HYPERLINK_PROTOCOLS.has(parsed.protocol)) return parsed.href;
    } catch {
      // Invalid and relative targets are not navigable from rendered document content.
    }
  }
  if (anchor) return `#${encodeURIComponent(anchor)}`;
  return '#';
}
