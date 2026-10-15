import type { ExcelColorModel } from './model.js';

const RGB_PATTERN = /^#[0-9A-Fa-f]{6}$/;

export function normalizeExcelColor(
  color: ExcelColorModel | undefined,
  context = 'Excel color',
): ExcelColorModel | undefined {
  if (color === undefined) return undefined;
  if (!color || typeof color !== 'object' || !RGB_PATTERN.test(color.rgb)) {
    throw new Error(`${context}: color must use { rgb: "#RRGGBB" }`);
  }
  if (Object.keys(color).some((key) => key !== 'rgb')) {
    throw new Error(`${context}: only the rgb field is supported`);
  }
  return { rgb: color.rgb.toUpperCase() };
}

export function ooxmlArgbToExcelColor(value: string | undefined): ExcelColorModel | undefined {
  if (!value) return undefined;
  const normalized = value.trim().replace(/^#/, '');
  if (/^[0-9A-Fa-f]{8}$/.test(normalized)) {
    return { rgb: `#${normalized.slice(2).toUpperCase()}` };
  }
  if (/^[0-9A-Fa-f]{6}$/.test(normalized)) {
    return { rgb: `#${normalized.toUpperCase()}` };
  }
  return undefined;
}

export function excelColorToOoxmlArgb(color: ExcelColorModel, context = 'Excel color'): string {
  const normalized = normalizeExcelColor(color, context)!;
  return `FF${normalized.rgb.slice(1)}`;
}

export function applyExcelTint(color: ExcelColorModel, tint: number | undefined): ExcelColorModel {
  if (tint === undefined || tint === 0 || !Number.isFinite(tint)) return color;
  const ratio = Math.max(-1, Math.min(1, tint));
  const channels = [1, 3, 5].map((offset) =>
    Number.parseInt(color.rgb.slice(offset, offset + 2), 16),
  );
  const adjusted = channels.map((channel) =>
    Math.round(ratio < 0 ? channel * (1 + ratio) : channel + (255 - channel) * ratio),
  );
  return {
    rgb: `#${adjusted
      .map((channel) => channel.toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()}`,
  };
}
