import { defineConfig } from 'vite';
import { spawnSync } from 'node:child_process';
import { createReadStream, statSync } from 'node:fs';
import { mkdtemp, open, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import type { IncomingMessage } from 'node:http';
import { tmpdir } from 'node:os';
import { basename, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = fileURLToPath(new URL('.', import.meta.url));
const fixtureDir = process.env.COCO_DOCX_FIXTURE_DIR || resolve(packageDir, '../../examples/output');
const additionalDocxFixturePaths = process.env.COCO_EXTRA_DOCX_FIXTURE ? [process.env.COCO_EXTRA_DOCX_FIXTURE] : [];
const spreadsheetFixtureDir = resolve(packageDir, '../core/test-fixtures/excel');
const pptFixtureDir = process.env.COCO_PPTX_FIXTURE_DIR || resolve(packageDir, '../../examples/output');
let pptxFixtureCache: Promise<Array<{ path: string; name: string; size: number }>> | null = null;

async function findLargestDocx(dir: string): Promise<{ path: string; size: number } | null> {
  const files = await listDocxFixtures(dir);
  return files[0] ?? null;
}

async function listDocxFixtures(dir: string): Promise<Array<{ path: string; name: string; size: number }>> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: Array<{ path: string; name: string; size: number }> = [];
  for (const entry of entries) {
    const fullPath = join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await listDocxFixtures(fullPath)));
      continue;
    }
    if (!entry.isFile()) continue;
    if (entry.name.startsWith('~$')) continue;
    const ext = extname(entry.name).toLowerCase();
    if (ext !== '.docx' && ext !== '.docm' && ext !== '.dotx') continue;
    const size = statSync(fullPath).size;
    if (size < 4 || !(await hasZipHeader(fullPath))) continue;
    files.push({ path: fullPath, name: entry.name, size });
  }
  return files.sort((a, b) => b.size - a.size || a.name.localeCompare(b.name));
}

async function listEditorDocxFixtures(): Promise<Array<{ path: string; name: string; size: number }>> {
  const files = await listDocxFixtures(fixtureDir);
  for (const path of additionalDocxFixturePaths) {
    try {
      const size = statSync(path).size;
      if (size >= 4 && await hasZipHeader(path)) files.push({ path, name: basename(path), size });
    } catch {
      // Ignore optional local fixtures that are not present on this machine.
    }
  }
  return files;
}

async function listXlsxFixtures(dir: string): Promise<Array<{ path: string; name: string; size: number }>> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: Array<{ path: string; name: string; size: number }> = [];
  for (const entry of entries) {
    const fullPath = join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await listXlsxFixtures(fullPath)));
      continue;
    }
    if (!entry.isFile()) continue;
    if (entry.name.startsWith('~$')) continue;
    const ext = extname(entry.name).toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xlsm') continue;
    const size = statSync(fullPath).size;
    if (size < 4 || !(await hasZipHeader(fullPath))) continue;
    files.push({ path: fullPath, name: entry.name, size });
  }
  return files.sort((a, b) => b.size - a.size || a.name.localeCompare(b.name));
}

async function listPptxFixtures(dir: string): Promise<Array<{ path: string; name: string; size: number }>> {
  const entries = await readdir(dir, { withFileTypes: true });
  const files: Array<{ path: string; name: string; size: number }> = [];
  for (const entry of entries) {
    const fullPath = join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await listPptxFixtures(fullPath)));
      continue;
    }
    if (!entry.isFile()) continue;
    if (entry.name.startsWith('~$')) continue;
    const ext = extname(entry.name).toLowerCase();
    if (ext !== '.pptx' && ext !== '.pptm') continue;
    const size = statSync(fullPath).size;
    if (size < 4 || !(await hasZipHeaderWithin(fullPath, 1_000))) continue;
    files.push({ path: fullPath, name: entry.name, size });
  }
  return files.sort((a, b) => b.size - a.size || a.name.localeCompare(b.name));
}

function listCachedPptxFixtures(): Promise<Array<{ path: string; name: string; size: number }>> {
  pptxFixtureCache ??= listPptxFixtures(pptFixtureDir);
  return pptxFixtureCache;
}

async function hasZipHeader(path: string): Promise<boolean> {
  const file = await open(path, 'r');
  try {
    const header = Buffer.alloc(4);
    const { bytesRead } = await file.read(header, 0, header.length, 0);
    return bytesRead === 4 && header[0] === 0x50 && header[1] === 0x4b;
  } finally {
    await file.close();
  }
}

function hasZipHeaderWithin(path: string, timeoutMs: number): boolean {
  const result = spawnSync(
    process.execPath,
    [
      '-e',
      [
        "const fs = require('node:fs');",
        "let fd;",
        "try {",
        "  fd = fs.openSync(process.argv[1], 'r');",
        "  const header = Buffer.alloc(4);",
        "  const read = fs.readSync(fd, header, 0, 4, 0);",
        "  process.stdout.write(read === 4 && header[0] === 0x50 && header[1] === 0x4b ? '1' : '0');",
        "} catch {",
        "  process.stdout.write('0');",
        "} finally {",
        "  if (fd !== undefined) { try { fs.closeSync(fd); } catch {} }",
        "}",
      ].join('\n'),
      path,
    ],
    { encoding: 'utf8', timeout: timeoutMs, windowsHide: true },
  );
  return result.status === 0 && result.stdout.trim() === '1';
}

function pipeFixtureFile(path: string, res: import('node:http').ServerResponse): void {
  const stream = createReadStream(path);
  stream.on('error', (error) => {
    if (!res.headersSent) {
      res.statusCode = 500;
      res.end(error instanceof Error ? error.stack || error.message : String(error));
      return;
    }
    res.destroy(error instanceof Error ? error : new Error(String(error)));
  });
  stream.pipe(res);
}

function isMetafileContentType(contentType: string): boolean {
  const normalized = contentType.toLowerCase();
  return (
    normalized.includes('emf') ||
    normalized.includes('wmf') ||
    normalized === 'image/x-emf' ||
    normalized === 'image/x-wmf' ||
    normalized === 'image/emf' ||
    normalized === 'image/wmf'
  );
}

function metafileExtension(contentType: string, mediaPath: string): '.emf' | '.wmf' | null {
  const pathExt = extname(mediaPath).toLowerCase();
  if (pathExt === '.emf' || contentType.toLowerCase().includes('emf')) return '.emf';
  if (pathExt === '.wmf' || contentType.toLowerCase().includes('wmf')) return '.wmf';
  return null;
}

function readRequestBody(req: IncomingMessage, maxBytes = 20_000_000): Promise<Buffer> {
  return new Promise((resolveBody, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on('data', (chunk: Buffer) => {
      size += chunk.byteLength;
      if (size > maxBytes) {
        reject(new Error(`Request body exceeds ${maxBytes} bytes`));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on('end', () => resolveBody(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

async function convertMetafilePreview(bytes: Buffer, ext: '.emf' | '.wmf'): Promise<Buffer> {
  const dir = await mkdtemp(join(tmpdir(), 'coco-ppt-preview-'));
  try {
    const inputPath = join(dir, `input${ext}`);
    const outputPath = join(dir, 'preview.png');
    await writeFile(inputPath, bytes);
    const encodedInput = Buffer.from(inputPath, 'utf8').toString('base64');
    const encodedOutput = Buffer.from(outputPath, 'utf8').toString('base64');
    const script = [
      '$ErrorActionPreference = "Stop"',
      'Add-Type -AssemblyName System.Drawing',
      `$inputPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${encodedInput}'))`,
      `$outputPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${encodedOutput}'))`,
      '$image = [System.Drawing.Image]::FromFile($inputPath)',
      'try {',
      '  $maxSize = 1600',
      '  $width = [Math]::Max(1, [int]$image.Width)',
      '  $height = [Math]::Max(1, [int]$image.Height)',
      '  $scale = [Math]::Min(1.0, $maxSize / [Math]::Max($width, $height))',
      '  $outWidth = [Math]::Max(1, [int][Math]::Round($width * $scale))',
      '  $outHeight = [Math]::Max(1, [int][Math]::Round($height * $scale))',
      '  $bitmap = New-Object System.Drawing.Bitmap $outWidth, $outHeight',
      '  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)',
      '  try {',
      '    $graphics.Clear([System.Drawing.Color]::Transparent)',
      '    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic',
      '    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality',
      '    $graphics.DrawImage($image, 0, 0, $outWidth, $outHeight)',
      '    $bitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)',
      '  } finally {',
      '    $graphics.Dispose()',
      '    $bitmap.Dispose()',
      '  }',
      '} finally {',
      '  $image.Dispose()',
      '}',
    ].join('\n');
    const result = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script], {
      encoding: 'utf8',
      timeout: 30_000,
      windowsHide: true,
    });
    if (result.status !== 0) {
      throw new Error(result.stderr.trim() || result.stdout.trim() || 'PowerShell metafile conversion failed');
    }
    return await readFile(outputPath);
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

export default defineConfig({
  root: 'dev',
  resolve: {
    alias: {
      '@coco-vexel/coco-office/pack': resolve(packageDir, '../core/src/pack/index.ts'),
      '@coco-vexel/coco-office/excel': resolve(packageDir, '../core/src/excel/index.ts'),
      '@coco-vexel/coco-office/ppt': resolve(packageDir, '../core/src/ppt/index.ts'),
      '@coco-vexel/coco-office/word-view': resolve(packageDir, '../core/src/word/view.ts'),
    },
  },
  server: {
    port: 5177,
    host: '127.0.0.1',
    watch:
      process.env.COCO_PPT_SMOKE === '1'
        ? {
            ignored: [
              '**/src/generated/editor-styles.ts',
              '**/src/styles/editor.css',
              '**/src/word/native-renderer.ts',
            ],
          }
        : undefined,
  },
  plugins: [
    {
      name: 'coco-editor-fixture',
      configureServer(server) {
        server.middlewares.use('/__coco_fixture__/files', async (_req, res) => {
          try {
            const files = await listEditorDocxFixtures();
            res.statusCode = 200;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.end(JSON.stringify(files.map((file, index) => ({ index, ...file }))));
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/docx', async (req, res) => {
          try {
            const files = await listEditorDocxFixtures();
            const url = new URL(req.url ?? '', 'http://127.0.0.1');
            const index = Number(url.searchParams.get('index') ?? '0');
            const fixture = files[index];
            if (!fixture) {
              res.statusCode = 404;
              res.end('No docx fixture found for index');
              return;
            }
            res.statusCode = 200;
            res.setHeader(
              'Content-Type',
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            );
            res.setHeader('Content-Length', String(fixture.size));
            res.setHeader('X-Coco-Fixture-Path', encodeURIComponent(fixture.path));
            pipeFixtureFile(fixture.path, res);
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/xlsx-files', async (_req, res) => {
          try {
            const files = await listXlsxFixtures(spreadsheetFixtureDir);
            res.statusCode = 200;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.end(JSON.stringify(files.map((file, index) => ({ index, ...file }))));
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/xlsx', async (req, res) => {
          try {
            const files = await listXlsxFixtures(spreadsheetFixtureDir);
            const url = new URL(req.url ?? '', 'http://127.0.0.1');
            const index = Number(url.searchParams.get('index') ?? '0');
            const fixture = files[index];
            if (!fixture) {
              res.statusCode = 404;
              res.end('No xlsx fixture found for index');
              return;
            }
            res.statusCode = 200;
            res.setHeader(
              'Content-Type',
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            );
            res.setHeader('Content-Length', String(fixture.size));
            res.setHeader('X-Coco-Fixture-Path', encodeURIComponent(fixture.path));
            pipeFixtureFile(fixture.path, res);
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/pptx-files', async (_req, res) => {
          try {
            const files = await listCachedPptxFixtures();
            res.statusCode = 200;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.end(JSON.stringify(files.map((file, index) => ({ index, ...file }))));
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/pptx', async (req, res) => {
          try {
            const files = await listCachedPptxFixtures();
            const url = new URL(req.url ?? '', 'http://127.0.0.1');
            const index = Number(url.searchParams.get('index') ?? '0');
            const fixture = files[index];
            if (!fixture) {
              res.statusCode = 404;
              res.end('No pptx fixture found for index');
              return;
            }
            res.statusCode = 200;
            res.setHeader(
              'Content-Type',
              'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            );
            res.setHeader('Content-Length', String(fixture.size));
            res.setHeader('X-Coco-Fixture-Path', encodeURIComponent(fixture.path));
            pipeFixtureFile(fixture.path, res);
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/ppt-image-preview', async (req, res) => {
          try {
            if (req.method !== 'POST') {
              res.statusCode = 405;
              res.setHeader('Allow', 'POST');
              res.end('Method not allowed');
              return;
            }
            const contentType = String(req.headers['x-coco-image-content-type'] ?? '');
            const mediaPath = decodeURIComponent(String(req.headers['x-coco-image-path'] ?? ''));
            if (!isMetafileContentType(contentType)) {
              res.statusCode = 415;
              res.end('Only EMF/WMF preview conversion is supported');
              return;
            }
            const ext = metafileExtension(contentType, mediaPath);
            if (!ext) {
              res.statusCode = 415;
              res.end('Unsupported metafile extension');
              return;
            }
            const body = await readRequestBody(req);
            if (!body.byteLength) {
              res.statusCode = 400;
              res.end('Empty image body');
              return;
            }
            const png = await convertMetafilePreview(body, ext);
            res.statusCode = 200;
            res.setHeader('Content-Type', 'image/png');
            res.setHeader('Content-Length', String(png.byteLength));
            res.end(png);
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
        server.middlewares.use('/__coco_fixture__/largest-docx', async (_req, res) => {
          try {
            const fixture = await findLargestDocx(fixtureDir);
            if (!fixture) {
              res.statusCode = 404;
              res.end('No docx fixture found');
              return;
            }
            res.statusCode = 200;
            res.setHeader(
              'Content-Type',
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            );
            res.setHeader('Content-Length', String(fixture.size));
            res.setHeader('X-Coco-Fixture-Path', encodeURIComponent(fixture.path));
            pipeFixtureFile(fixture.path, res);
          } catch (err) {
            res.statusCode = 500;
            res.end(err instanceof Error ? err.stack || err.message : String(err));
          }
        });
      },
    },
  ],
});
