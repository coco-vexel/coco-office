import { spawn } from 'node:child_process';
import http from 'node:http';
import { access, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const regressionSuites = {
  'word-native-p0': {
    description: 'Native Word P0 content preservation baseline',
    fixtures: [0, 1, 2, 7, 8, 12, 19, 25, 46, 73, 74],
    expectedReviewNotes: new Map([
      [2, 'long CJK document page drift accepted when coverage and overflow pass'],
      [7, '200+ page document page drift accepted when coverage and overflow pass'],
      [12, 'table-heavy product document page drift accepted when coverage and overflow pass'],
      [25, 'revision sample page drift accepted when coverage and overflow pass'],
    ]),
  },
  'word-native-p1-extra': {
    description: 'Native Word extra rendering closure samples',
    fixtures: [37, 42, 48, 94, 115, 117, 119, 147],
    expectedReviewNotes: new Map([
      [37, 'generated red-head document visual parity is reviewed manually'],
      [42, 'formal generated document visual parity is reviewed manually'],
      [48, 'vMerge table visual parity is reviewed manually'],
      [94, 'page-specific header/footer visual parity is reviewed manually'],
      [115, 'header/footer unicode visual parity is reviewed manually'],
      [117, 'complex numbering visual parity is reviewed manually'],
      [119, 'different first-page header/footer visual parity is reviewed manually'],
      [147, 'comment reply visual parity is reviewed manually'],
    ]),
  },
};

function parseArgs(argv) {
  const parsed = {
    all: false,
    baseUrl: undefined,
    fixtureSpec: undefined,
    limit: 5,
    outputDir: resolve(packageDir, 'dev/native-regression-output'),
    port: 5177,
    screenshotPages: 3,
    screenshots: true,
    startServer: false,
    suite: undefined,
    wordReferenceDir: undefined,
  };
  for (const arg of argv) {
    if (arg === '--') {
      continue;
    } else if (arg === '--all') {
      parsed.all = true;
      parsed.limit = undefined;
    } else if (arg === '--start-server') {
      parsed.startServer = true;
    } else if (arg.startsWith('--base-url=')) {
      parsed.baseUrl = arg.slice('--base-url='.length).replace(/\/$/, '');
    } else if (arg.startsWith('--fixtures=')) {
      parsed.fixtureSpec = arg.slice('--fixtures='.length);
    } else if (arg.startsWith('--limit=')) {
      parsed.limit = Number(arg.slice('--limit='.length));
    } else if (arg === '--no-screenshots') {
      parsed.screenshots = false;
    } else if (arg.startsWith('--output-dir=')) {
      parsed.outputDir = resolve(packageDir, arg.slice('--output-dir='.length));
    } else if (arg.startsWith('--port=')) {
      parsed.port = Number(arg.slice('--port='.length));
    } else if (arg.startsWith('--screenshot-pages=')) {
      parsed.screenshotPages = Number(arg.slice('--screenshot-pages='.length));
    } else if (arg.startsWith('--suite=')) {
      parsed.suite = arg.slice('--suite='.length);
    } else if (arg.startsWith('--word-reference-dir=')) {
      parsed.wordReferenceDir = resolve(packageDir, arg.slice('--word-reference-dir='.length));
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!Number.isInteger(parsed.port) || parsed.port <= 0) throw new Error('Invalid --port value.');
  if (parsed.limit !== undefined && (!Number.isInteger(parsed.limit) || parsed.limit <= 0)) {
    throw new Error('Invalid --limit value.');
  }
  if (!Number.isInteger(parsed.screenshotPages) || parsed.screenshotPages < 0) {
    throw new Error('Invalid --screenshot-pages value.');
  }
  if (parsed.suite !== undefined && !(parsed.suite in regressionSuites)) {
    throw new Error(`Unknown --suite value: ${parsed.suite}. Available suites: ${Object.keys(regressionSuites).join(', ')}`);
  }
  if (parsed.suite !== undefined && parsed.fixtureSpec !== undefined) {
    throw new Error('Use either --suite or --fixtures, not both.');
  }
  parsed.wordReferenceDir ??= resolve(parsed.outputDir, 'word-reference');
  return parsed;
}

async function startVite(port) {
  const viteEntry = resolve(packageDir, 'node_modules/vite/bin/vite.js');
  const child = spawn(process.execPath, [viteEntry, '--host', '127.0.0.1', '--port', String(port), '--strictPort'], {
    cwd: packageDir,
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  child.stdout.on('data', (chunk) => process.stdout.write(`[vite] ${chunk}`));
  child.stderr.on('data', (chunk) => process.stderr.write(`[vite] ${chunk}`));
  await waitForHttp(`http://127.0.0.1:${port}/__coco_fixture__/files`, 20_000);
  return child;
}

async function startChrome() {
  const port = await findFreePort(9222);
  const userDataDir = join(tmpdir(), `coco-word-regression-${Date.now()}-${Math.random().toString(16).slice(2)}`);
  await mkdir(userDataDir, { recursive: true });
  const executable = chromeExecutable();
  const child = spawn(
    executable,
    [
      '--headless=new',
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${userDataDir}`,
      '--disable-gpu',
      '--no-first-run',
      '--no-default-browser-check',
      'about:blank',
    ],
    { stdio: 'ignore' },
  );
  await waitForHttp(`http://127.0.0.1:${port}/json/version`, 20_000);
  return { process: child, port, userDataDir };
}

function chromeExecutable() {
  if (process.env.CHROME_PATH) return process.env.CHROME_PATH;
  if (process.platform === 'win32') return 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
  if (process.platform === 'darwin') return '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  return 'google-chrome';
}

async function runFixture(chromePort, appBaseUrl, fixture, options) {
  const pageInfo = await createCdpPage(chromePort);
  const client = await CdpClient.connect(pageInfo.webSocketDebuggerUrl);
  try {
    await client.send('Page.enable');
    await client.send('Runtime.enable');
    await client.send('Emulation.setDeviceMetricsOverride', {
      width: 1440,
      height: 1200,
      deviceScaleFactor: 1,
      mobile: false,
    });
    const url = `${appBaseUrl}/?engine=native&fixture=${fixture.index}&bust=regression-${Date.now()}`;
    await client.send('Page.navigate', { url });
    await waitForPageReady(client);
    const value = await evaluate(client, acceptanceExpression());
    const screenshots = options.screenshots
      ? await capturePageScreenshots(client, fixture, value.pageLayouts, options)
      : [];
    const wordReference = await findWordReferenceMetadata(fixture, options);
    return {
      fixture,
      url,
      status: value.status,
      rootClassName: value.rootClassName,
      metrics: value.metrics,
      pageLayouts: value.pageLayouts,
      pageFieldSamples: value.pageFieldSamples,
      pageStorySamples: value.pageStorySamples,
      pageAnchorSamples: value.pageAnchorSamples,
      tableLayoutSamples: value.tableLayoutSamples,
      pageTextSamples: value.pageTextSamples,
      pageBlockSamples: value.pageBlockSamples,
      risks: value.risks,
      consoleErrors: value.consoleErrors,
      screenshots,
      wordReferences: wordReference.screenshots,
      wordReferencePageCount: wordReference.pageCount,
    };
  } finally {
    await client.send('Page.close').catch(() => undefined);
    client.close();
  }
}

async function findWordReferenceMetadata(fixture, options) {
  const fixtureDir = join(options.wordReferenceDir, sanitizePathSegment(fixture.name));
  const files = [];
  for (let page = 1; page <= options.screenshotPages; page++) {
    const file = await firstExistingPath([
      join(fixtureDir, `page-${page}.png`),
      join(fixtureDir, `page-${String(page).padStart(2, '0')}.png`),
      join(fixtureDir, `page-${String(page).padStart(3, '0')}.png`),
    ]);
    files.push(file);
  }
  const pdfPath = await firstExistingPath([join(fixtureDir, 'reference.pdf')]);
  const pdfPageCount = pdfPath ? await countPdfPages(pdfPath) : null;
  const pngPageCount = await countWordReferencePngPages(fixtureDir);
  return {
    screenshots: files,
    pageCount: maxNullable(pdfPageCount, pngPageCount),
  };
}

function maxNullable(...values) {
  const numeric = values.filter((value) => typeof value === 'number');
  return numeric.length > 0 ? Math.max(...numeric) : null;
}

async function countWordReferencePngPages(directory) {
  try {
    const entries = await readdir(directory);
    const pages = entries
      .map((entry) => entry.match(/^page-(\d+)\.png$/i)?.[1])
      .filter(Boolean)
      .map((page) => Number(page))
      .filter((page) => Number.isInteger(page) && page > 0);
    return pages.length > 0 ? Math.max(...pages) : null;
  } catch {
    return null;
  }
}

async function firstExistingPath(paths) {
  for (const file of paths) {
    try {
      await access(file);
      return file;
    } catch {
      // Try the next name pattern.
    }
  }
  return null;
}

async function countPdfPages(file) {
  try {
    const content = await readFile(file, 'latin1');
    return (content.match(/\/Type\s*\/Page\b/g) ?? []).length;
  } catch {
    return null;
  }
}

async function createCdpPage(chromePort) {
  let lastError;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      return await new Promise((resolvePage, rejectPage) => {
        const request = http.request(
          `http://127.0.0.1:${chromePort}/json/new?about:blank`,
          { method: 'PUT', timeout: 5_000 },
          (response) => {
            let body = '';
            response.setEncoding('utf8');
            response.on('data', (chunk) => { body += chunk; });
            response.on('end', () => {
              if (response.statusCode !== 200) {
                rejectPage(new Error(`Failed to create Chrome tab: ${response.statusCode}`));
                return;
              }
              try { resolvePage(JSON.parse(body)); } catch (error) { rejectPage(error); }
            });
          },
        );
        request.on('timeout', () => request.destroy(new Error('Chrome tab creation timed out')));
        request.on('error', rejectPage);
        request.end();
      });
    } catch (error) {
      lastError = error;
      await sleep(250 * (attempt + 1));
    }
  }
  throw new Error(`Failed to create Chrome tab after retries: ${lastError?.message ?? 'unknown error'}`);
}

async function waitForPageReady(client) {
  const start = Date.now();
  while (Date.now() - start < 60_000) {
    const ready = await evaluate(
      client,
      `(() => { const host = document.querySelector('.coco-word-editor-host'); const root = host?.shadowRoot ?? document; return document.readyState === 'complete' && !!root.querySelector('coco-word-native-document') && !!document.querySelector('.coco-word-acceptance-status'); })()`,
    ).catch(() => false);
    if (ready) {
      await sleep(750);
      return;
    }
    await sleep(250);
  }
  throw new Error('Timed out waiting for native editor acceptance panel.');
}

async function evaluate(client, expression) {
  const response = await client.send('Runtime.evaluate', {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (response.exceptionDetails) {
    throw new Error(response.exceptionDetails.text ?? 'Runtime.evaluate failed');
  }
  return response.result.value;
}

function acceptanceExpression() {
  return `(() => {
    const editorHost = document.querySelector('.coco-word-editor-host');
    const editorRoot = editorHost?.shadowRoot ?? document;
    const metrics = Array.from(document.querySelectorAll('.coco-word-acceptance-metric')).map((el) => ({
      label: el.querySelector('span')?.textContent?.trim() ?? '',
      value: el.querySelector('strong')?.textContent?.trim() ?? '',
      tone: el.classList.contains('is-warn') ? 'warn' : el.classList.contains('is-good') ? 'good' : undefined,
    }));
    return {
      status: document.querySelector('.coco-word-acceptance-status')?.textContent?.trim() ?? 'CHECK',
      metrics,
      rootClassName: editorRoot.querySelector('coco-word-native-document')?.className ?? '',
      pageLayouts: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).map((page, index) => {
        const rect = page.getBoundingClientRect();
        const body = page.querySelector('coco-word-native-page-body');
        const bodyRect = body?.getBoundingClientRect();
        return {
          index: index + 1,
          width: Math.round(rect.width),
          height: Math.round(rect.height),
          bodyWidth: Math.round(bodyRect?.width ?? 0),
          bodyHeight: Math.round(bodyRect?.height ?? 0),
          x: Math.round(rect.left + scrollX),
          y: Math.round(rect.top + scrollY),
          paragraphs: page.querySelectorAll('coco-word-native-p').length,
          tables: page.querySelectorAll('coco-word-native-table').length,
          images: page.querySelectorAll('img').length,
          drawings: page.querySelectorAll('coco-word-native-drawing').length,
          comments: page.querySelectorAll('.coco-word-native-comment-bubble').length,
          overflow: body && bodyRect ? body.scrollHeight > bodyRect.height + 24 : false,
        };
      }),
      pageFieldSamples: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).slice(0, 12).map((page, index) => ({
        index: index + 1,
        pageNumber: page.getAttribute('data-page-number') ?? '',
        displayPageNumber: page.getAttribute('data-display-page-number') ?? '',
        sectionIndex: page.getAttribute('data-section-index') ?? '',
        sectionPageNumber: page.getAttribute('data-section-page-number') ?? '',
        fields: Array.from(page.querySelectorAll('[data-field-kind]')).map((field) => ({
          kind: field.getAttribute('data-field-kind') ?? '',
          text: field.textContent?.trim() ?? '',
          pageNumber: field.getAttribute('data-page-number') ?? '',
          sectionPageNumber: field.getAttribute('data-section-page-number') ?? '',
          totalPages: field.getAttribute('data-total-pages') ?? '',
          sectionPages: field.getAttribute('data-section-pages') ?? '',
        })),
      })),
      pageStorySamples: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).slice(0, 12).map((page, index) => ({
        index: index + 1,
        pageNumber: page.getAttribute('data-page-number') ?? '',
        sectionIndex: page.getAttribute('data-section-index') ?? '',
        sectionPageNumber: page.getAttribute('data-section-page-number') ?? '',
        headers: Array.from(page.querySelectorAll(':scope > coco-word-native-story[data-story-kind="header"]')).map((story) => ({
          partName: story.getAttribute('data-part-name') ?? '',
          text: story.textContent?.replace(/\\s+/g, ' ').trim() ?? '',
          images: story.querySelectorAll('img').length,
          drawings: story.querySelectorAll('coco-word-native-drawing').length,
          borderedParagraphs: Array.from(story.querySelectorAll('coco-word-native-p')).filter((paragraph) =>
            (paragraph.getAttribute('style') ?? '').includes('border-'),
          ).length,
        })),
        footers: Array.from(page.querySelectorAll(':scope > coco-word-native-story[data-story-kind="footer"]')).map((story) => ({
          partName: story.getAttribute('data-part-name') ?? '',
          text: story.textContent?.replace(/\\s+/g, ' ').trim() ?? '',
          images: story.querySelectorAll('img').length,
          drawings: story.querySelectorAll('coco-word-native-drawing').length,
          borderedParagraphs: Array.from(story.querySelectorAll('coco-word-native-p')).filter((paragraph) =>
            (paragraph.getAttribute('style') ?? '').includes('border-'),
          ).length,
        })),
      })),
      pageAnchorSamples: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).slice(0, 12).map((page, index) => ({
        index: index + 1,
        pageNumber: page.getAttribute('data-page-number') ?? '',
        anchors: Array.from(page.querySelectorAll('.is-page-layer-anchor, .is-paragraph-relative-anchor, [data-anchoring-mode="anchor"]')).map((anchor) => ({
          tag: anchor.tagName.toLowerCase(),
          drawingId: anchor.getAttribute('data-drawing-id') ?? '',
          wrap: anchor.getAttribute('data-wrap') ?? '',
          sourceStoryKind: anchor.getAttribute('data-source-story-kind') ?? '',
          sourcePartName: anchor.getAttribute('data-source-part-name') ?? '',
          pageLayer: anchor.getAttribute('data-page-layer-anchor') ?? '',
          pageLayerKind: anchor.getAttribute('data-page-layer') ?? '',
          zIndex: window.getComputedStyle(anchor).zIndex,
          parentClassName: anchor.parentElement?.className ?? '',
          horizontalRelativeFrom: anchor.getAttribute('data-position-horizontal-relative-from') ?? '',
          verticalRelativeFrom: anchor.getAttribute('data-position-vertical-relative-from') ?? '',
        })),
      })),
      tableLayoutSamples: Array.from(editorRoot.querySelectorAll('coco-word-native-table')).map((wrapper, index) => {
        const table = wrapper.querySelector('table');
        const page = wrapper.closest('coco-word-native-page');
        const body = wrapper.closest('coco-word-native-page-body');
        const wrapperRect = wrapper.getBoundingClientRect();
        const tableRect = table?.getBoundingClientRect();
        const columns = table
          ? Array.from(table.querySelectorAll('col')).map((col) => window.getComputedStyle(col).width)
          : [];
        const rows = table ? Array.from(table.querySelectorAll('tbody > tr')) : [];
        return {
          index: index + 1,
          pageIndex: page ? Array.from(editorRoot.querySelectorAll('coco-word-native-page')).indexOf(page) + 1 : 0,
          bodyClientHeight: body?.clientHeight ?? 0,
          bodyScrollHeight: body?.scrollHeight ?? 0,
          wrapperClientWidth: wrapper.clientWidth,
          wrapperScrollWidth: wrapper.scrollWidth,
          wrapperRectWidth: Math.round(wrapperRect.width),
          wrapperRectHeight: Math.round(wrapperRect.height),
          tableRectWidth: Math.round(tableRect?.width ?? 0),
          tableRectHeight: Math.round(tableRect?.height ?? 0),
          tableStyleWidth: table?.style.width ?? '',
          overflowX: wrapper.scrollWidth > wrapper.clientWidth + 1,
          columns: columns.slice(0, 16),
          rows: rows.slice(0, 24).map((row, rowIndex) => {
            const rowRect = row.getBoundingClientRect();
            const firstCell = row.querySelector('td');
            const firstParagraph = firstCell?.querySelector('coco-word-native-p');
            const cellStyle = firstCell ? window.getComputedStyle(firstCell) : null;
            const paragraphStyle = firstParagraph ? window.getComputedStyle(firstParagraph) : null;
            return {
              index: rowIndex + 1,
              height: Math.round(rowRect.height),
              fragment: row.getAttribute('data-row-fragment') ?? '',
              rule: row.getAttribute('data-height-rule') ?? '',
              cssHeight: row.style.height,
              cellPadding:
                cellStyle === null
                  ? ''
                  : [cellStyle.paddingTop, cellStyle.paddingRight, cellStyle.paddingBottom, cellStyle.paddingLeft].join(' '),
              cellFontSize: cellStyle?.fontSize ?? '',
              cellLineHeight: cellStyle?.lineHeight ?? '',
              paragraphFontSize: paragraphStyle?.fontSize ?? '',
              paragraphLineHeight: paragraphStyle?.lineHeight ?? '',
              text: row.textContent?.replace(/\s+/g, ' ').trim().slice(0, 100) ?? '',
            };
          }),
          text: wrapper.textContent?.replace(/\s+/g, ' ').trim().slice(0, 120) ?? '',
        };
      }),
      pageTextSamples: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).map((page, index) => {
        const body = page.querySelector('coco-word-native-page-body');
        const text = (body?.textContent ?? '').replace(/\s+/g, ' ').trim();
        return {
          index: index + 1,
          pageNumber: page.getAttribute('data-page-number') ?? '',
          displayPageNumber: page.getAttribute('data-display-page-number') ?? '',
          sectionIndex: page.getAttribute('data-section-index') ?? '',
          sectionPageNumber: page.getAttribute('data-section-page-number') ?? '',
          length: text.length,
          head: text.slice(0, 160),
          tail: text.slice(Math.max(0, text.length - 160)),
        };
      }),
      pageBlockSamples: Array.from(editorRoot.querySelectorAll('coco-word-native-page')).map((page, index) => {
        const body = page.querySelector('coco-word-native-page-body');
        const blocks = body ? Array.from(body.children).filter((child) => child instanceof HTMLElement) : [];
        let cumulative = 0;
        return {
          index: index + 1,
          pageNumber: page.getAttribute('data-page-number') ?? '',
          bodyScrollHeight: body?.scrollHeight ?? 0,
          bodyClientHeight: body?.clientHeight ?? 0,
          blocks: blocks.map((block, blockIndex) => {
            const rect = block.getBoundingClientRect();
            const style = window.getComputedStyle(block);
            const marginTop = Number.parseFloat(style.marginTop) || 0;
            const marginBottom = Number.parseFloat(style.marginBottom) || 0;
            cumulative += rect.height;
            return {
              index: blockIndex + 1,
              tag: block.tagName.toLowerCase(),
              height: Math.round(rect.height),
              marginTop: Math.round(marginTop),
              marginBottom: Math.round(marginBottom),
              outerHeight: Math.round(rect.height + marginTop + marginBottom),
              cumulative: Math.round(cumulative),
              flowEmpty: block.classList.contains('is-flow-empty'),
              fontSize: style.fontSize,
              lineHeight: style.lineHeight,
              fontFamily: style.fontFamily,
              textIndent: style.textIndent,
              textAlign: style.textAlign,
              snapToGrid: block.getAttribute('data-snap-to-grid') ?? '',
              keepNext: block.getAttribute('data-keep-next') ?? '',
              text: block.textContent?.replace(/\\s+/g, ' ').trim().slice(0, 120) ?? '',
              images: block.querySelectorAll('img').length,
              tables: block.querySelectorAll('table').length,
            };
          }),
        };
      }),
      risks: Array.from(document.querySelectorAll('.coco-word-acceptance-section + ul li')).map((li) => li.textContent?.trim()).filter(Boolean),
      consoleErrors: performance.getEntriesByType('resource')
        .filter((entry) => entry.name.includes('/__coco_fixture__/') && entry.responseStatus >= 400)
        .map((entry) => entry.name),
    };
  })()`;
}

async function capturePageScreenshots(client, fixture, pageLayouts, options) {
  const pages = pageLayouts.slice(0, options.screenshotPages);
  if (pages.length === 0) return [];
  await enableScreenshotMode(client);
  const fixtureDir = join(options.outputDir, `${String(fixture.index).padStart(3, '0')}-${sanitizePathSegment(fixture.name)}`);
  await mkdir(fixtureDir, { recursive: true });
  const files = [];
  for (const page of pages) {
    const file = join(fixtureDir, `page-${String(page.index).padStart(2, '0')}.png`);
    const rect = await scrollPageIntoViewport(client, page.index);
    const screenshot = await client.send('Page.captureScreenshot', {
      format: 'png',
      fromSurface: true,
      captureBeyondViewport: false,
      clip: {
        x: Math.max(0, rect.left),
        y: Math.max(0, rect.top),
        width: Math.max(1, rect.width),
        height: Math.max(1, rect.height),
        scale: 1,
      },
    });
    await writeFile(file, Buffer.from(screenshot.data, 'base64'));
    files.push(file);
  }
  return files;
}

async function enableScreenshotMode(client) {
  await evaluate(
    client,
    `(() => {
      if (document.getElementById('coco-native-regression-screenshot-style')) return true;
      const style = document.createElement('style');
      style.id = 'coco-native-regression-screenshot-style';
      style.textContent = [
        '.coco-word-devbar, .coco-word-toolbar, .coco-word-inspector { display: none !important; }',
        '.coco-word-shell { display: block !important; }',
        '.coco-word-document-scroll { height: auto !important; min-height: 0 !important; overflow: visible !important; padding: 0 !important; background: #ffffff !important; }',
        '.coco-word-pages { align-items: center !important; padding: 0 !important; }',
        'body { margin: 0 !important; background: #ffffff !important; }',
      ].join('\\n');
      document.head.appendChild(style);
      return true;
    })()`,
  );
  await sleep(150);
}

async function scrollPageIntoViewport(client, pageIndex) {
  const rect = await evaluate(
    client,
    `(() => {
      const host = document.querySelector('.coco-word-editor-host');
      const root = host?.shadowRoot ?? document;
      const page = root.querySelectorAll('coco-word-native-page')[${pageIndex - 1}];
      if (!page) throw new Error('Native page ${pageIndex} was not found.');
      page.scrollIntoView({ block: 'start', inline: 'center' });
      const rect = page.getBoundingClientRect();
      return {
        left: Math.round(rect.left),
        top: Math.round(rect.top),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      };
    })()`,
  );
  await sleep(150);
  return rect;
}

function selectFixtures(fixtures, options) {
  if (options.suite) {
    const suite = regressionSuites[options.suite];
    const byIndex = new Map(fixtures.map((fixture) => [fixture.index, fixture]));
    const missing = suite.fixtures.filter((index) => !byIndex.has(index));
    if (missing.length > 0) {
      throw new Error(`Suite ${options.suite} references missing fixture index(es): ${missing.join(', ')}`);
    }
    return suite.fixtures.map((index) => byIndex.get(index));
  }
  if (options.fixtureSpec) {
    const wanted = expandFixtureSpec(options.fixtureSpec);
    return fixtures.filter(
      (fixture) =>
        wanted.indices.has(fixture.index) ||
        wanted.names.some((name) => fixture.name.toLowerCase().includes(name)),
    );
  }
  if (options.all) return fixtures;
  return fixtures.slice(0, options.limit);
}

function expandFixtureSpec(spec) {
  const selected = { indices: new Set(), names: [] };
  for (const part of spec.split(',')) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    if (/^\d+-\d+$/.test(trimmed)) {
      const [start, end] = trimmed.split('-').map((value) => Number(value));
      if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) {
        throw new Error(`Invalid fixture range: ${trimmed}`);
      }
      for (let index = start; index <= end; index++) selected.indices.add(index);
    } else {
      const index = Number(trimmed);
      if (Number.isInteger(index)) selected.indices.add(index);
      else selected.names.push(trimmed.toLowerCase());
    }
  }
  return selected;
}

function printResults(results) {
  for (const result of results) {
    const keyMetrics = ['Paragraphs', 'Tables', 'Table Pagination', 'Images', 'Anchored Images', 'Comments', 'Style Probe']
      .map((label) => {
        const metric = result.metrics.find((item) => item.label === label);
        return metric ? `${label}: ${metric.value}` : null;
      })
      .filter(Boolean)
      .join(' | ');
    console.log(`${result.status} fixture ${result.fixture.index}: ${result.fixture.name}`);
    console.log(`  ${keyMetrics}`);
    console.log(`  Pages: ${summarizePages(result.pageLayouts)}`);
    if (result.screenshots.length > 0) {
      console.log(`  Screenshots: ${result.screenshots.join(', ')}`);
    }
    const referenceCount = result.wordReferences.filter(Boolean).length;
    if (referenceCount > 0) {
      console.log(`  Word references: ${referenceCount}/${result.wordReferences.length}`);
    }
    if (result.wordReferencePageCount != null) {
      console.log(`  Word pages: ${result.wordReferencePageCount}`);
    }
    if (result.risks.length > 0) {
      for (const risk of result.risks) console.log(`  risk: ${risk}`);
    }
  }
  printAcceptanceMatrix(buildRegressionSummary(results));
}

function summarizePages(pageLayouts) {
  if (pageLayouts.length === 0) return 'none';
  const overflow = pageLayouts.filter((page) => page.overflow).map((page) => page.index);
  const sample = pageLayouts
    .slice(0, 6)
    .map((page) => `p${page.index}:${page.width}x${page.height}, para ${page.paragraphs}, table ${page.tables}, img ${page.images}`)
    .join('; ');
  const suffix = pageLayouts.length > 6 ? `; +${pageLayouts.length - 6} pages` : '';
  const overflowText = overflow.length > 0 ? ` | overflow pages ${overflow.join(',')}` : '';
  return `${pageLayouts.length} pages | ${sample}${suffix}${overflowText}`;
}

function buildRegressionSummary(results) {
  const rows = results.map(summarizeRegressionResult);
  return {
    totals: {
      fixtures: rows.length,
      acceptancePass: rows.filter((row) => row.acceptanceStatus === 'PASS').length,
      acceptanceReview: rows.filter((row) => row.acceptanceStatus === 'REVIEW').length,
      acceptanceFail: rows.filter((row) => row.acceptanceStatus === 'FAIL').length,
      contentPass: rows.filter((row) => row.contentStatus === 'PASS').length,
      layoutPass: rows.filter((row) => row.layoutStatus === 'PASS').length,
      check: rows.filter((row) => row.contentStatus !== 'PASS' || row.layoutStatus !== 'PASS').length,
    },
    rows,
  };
}

function summarizeRegressionResult(result) {
  const coverageIssues = findCoverageIssues(result);
  const overflowPages = result.pageLayouts.filter((page) => page.overflow).map((page) => page.index);
  const wordReferenceTotal =
    result.wordReferencePageCount == null
      ? result.wordReferences.length
      : Math.min(result.wordReferences.length, result.wordReferencePageCount);
  const wordReferenceCount = result.wordReferences.slice(0, wordReferenceTotal).filter(Boolean).length;
  const missingWordReferences = wordReferenceTotal - wordReferenceCount;
  const layoutIssues = [];
  if (result.wordReferencePageCount != null && result.pageLayouts.length !== result.wordReferencePageCount) {
    layoutIssues.push(`page count ${result.pageLayouts.length}/${result.wordReferencePageCount}`);
  }
  if (overflowPages.length > 0) layoutIssues.push(`overflow pages ${overflowPages.join(',')}`);
  if (result.risks.length > 0) layoutIssues.push(`${result.risks.length} risk(s)`);
  if (missingWordReferences > 0 && wordReferenceTotal > 0) {
    layoutIssues.push(`missing ${missingWordReferences} Word reference(s)`);
  }
  const consoleIssues = result.consoleErrors.map((error) => `console: ${String(error).slice(0, 80)}`);
  const contentStatus =
    result.status === 'PASS' && coverageIssues.length === 0 && consoleIssues.length === 0 ? 'PASS' : 'CHECK';
  const layoutStatus = layoutIssues.length === 0 ? 'PASS' : 'CHECK';
  const acceptanceStatus =
    contentStatus !== 'PASS' || overflowPages.length > 0
      ? 'FAIL'
      : layoutStatus === 'PASS'
        ? 'PASS'
        : 'REVIEW';
  const suiteReviewNote = result.suiteExpectedReviewNote;
  return {
    fixtureIndex: result.fixture.index,
    fixtureName: result.fixture.name,
    status: result.status,
    acceptanceStatus,
    contentStatus,
    layoutStatus,
    pageCount: result.pageLayouts.length,
    wordPageCount: result.wordReferencePageCount,
    overflowCount: overflowPages.length,
    wordReferenceCount,
    wordReferenceTotal,
    expectedReview: suiteReviewNote,
    notes: [...coverageIssues, ...consoleIssues, ...layoutIssues, suiteReviewNote ? `expected review: ${suiteReviewNote}` : null]
      .filter(Boolean)
      .slice(0, 6),
  };
}

function findCoverageIssues(result) {
  const issues = [];
  for (const label of [
    'Paragraphs',
    'Tables',
    'Images',
    'Anchored Images',
    'Drawings',
    'TextBoxes',
    'Numbered',
    'Page Breaks',
    'Comments',
    'Revisions',
  ]) {
    const metric = result.metrics.find((item) => item.label === label);
    if (!metric) continue;
    const issue = coverageIssueForMetric(metric);
    if (issue) issues.push(issue);
  }
  return issues;
}

function coverageIssueForMetric(metric) {
  const numbers = Array.from(metric.value.matchAll(/\d+/g), (match) => Number(match[0]));
  if (numbers.length < 2) return null;
  if (metric.label === 'Comments') {
    return numbers.every((value) => value === numbers[0]) ? null : `${metric.label} mismatch ${metric.value}`;
  }
  return numbers[0] === numbers[1] ? null : `${metric.label} mismatch ${metric.value}`;
}

function printAcceptanceMatrix(summary) {
  console.log('');
  console.log('Acceptance matrix');
  const rows = summary.rows.map((row) => ({
    Fixture: `${row.fixtureIndex}: ${row.fixtureName}`,
    Acceptance: row.acceptanceStatus,
    Content: row.contentStatus,
    Layout: row.layoutStatus,
    Pages: row.wordPageCount == null ? String(row.pageCount) : `${row.pageCount}/${row.wordPageCount}`,
    Overflow: String(row.overflowCount),
    'Word refs': `${row.wordReferenceCount}/${row.wordReferenceTotal}`,
    Notes: row.notes.length > 0 ? row.notes.join('; ') : '-',
  }));
  printTable(rows);
  console.log(
    `Summary: ${summary.totals.fixtures} fixture(s), acceptance ${summary.totals.acceptancePass} PASS / ${summary.totals.acceptanceReview} REVIEW / ${summary.totals.acceptanceFail} FAIL, content ${summary.totals.contentPass}/${summary.totals.fixtures} PASS, layout ${summary.totals.layoutPass}/${summary.totals.fixtures} PASS.`,
  );
}

function printTable(rows) {
  if (rows.length === 0) return;
  const headers = Object.keys(rows[0]);
  const widths = headers.map((header) =>
    Math.max(header.length, ...rows.map((row) => String(row[header]).length)),
  );
  const formatRow = (row) =>
    headers.map((header, index) => String(row[header]).padEnd(widths[index])).join(' | ');
  console.log(formatRow(Object.fromEntries(headers.map((header) => [header, header]))));
  console.log(widths.map((width) => '-'.repeat(width)).join('-|-'));
  for (const row of rows) console.log(formatRow(row));
}

function sanitizePathSegment(value) {
  return value.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 80) || 'fixture';
}

async function readJson(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Failed to fetch ${url}: ${response.status}`);
  return response.json();
}

async function waitForHttp(url, timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Keep polling until the server is ready.
    }
    await sleep(250);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

async function findFreePort(startPort) {
  for (let port = startPort; port < startPort + 100; port++) {
    try {
      await fetch(`http://127.0.0.1:${port}/json/version`, { signal: AbortSignal.timeout(100) });
    } catch {
      return port;
    }
  }
  throw new Error('Could not find a free Chrome debugging port.');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

class CdpClient {
  static connect(url) {
    return new Promise((resolve, reject) => {
      const socket = new WebSocket(url);
      const client = new CdpClient(socket);
      socket.addEventListener('open', () => resolve(client), { once: true });
      socket.addEventListener('error', reject, { once: true });
    });
  }

  constructor(socket) {
    this.socket = socket;
    this.nextId = 1;
    this.pending = new Map();
    socket.addEventListener('message', (event) => this.onMessage(event));
    socket.addEventListener('close', () => {
      for (const { reject } of this.pending.values()) reject(new Error('CDP socket closed.'));
      this.pending.clear();
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    this.socket.send(JSON.stringify({ id, method, params }));
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
    });
  }

  close() {
    this.socket.close();
  }

  onMessage(event) {
    const message = JSON.parse(event.data);
    if (!message.id) return;
    const pending = this.pending.get(message.id);
    if (!pending) return;
    this.pending.delete(message.id);
    if (message.error) pending.reject(new Error(message.error.message));
    else pending.resolve(message.result);
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const baseUrl = args.baseUrl ?? `http://127.0.0.1:${args.port}`;
  let viteProcess = null;
  let chromeProcess = null;
  let chromeUserDataDir = null;

  try {
    if (args.startServer) {
      viteProcess = await startVite(args.port);
    } else {
      await waitForHttp(`${baseUrl}/__coco_fixture__/files`, 1_500);
    }

    const fixtures = await readJson(`${baseUrl}/__coco_fixture__/files`);
    const selectedFixtures = selectFixtures(fixtures, args);
    if (selectedFixtures.length === 0) {
      throw new Error('No fixtures selected.');
    }

    const results = [];
    for (const fixture of selectedFixtures) {
      const chrome = await startChrome();
      try {
        const result = await runFixture(chrome.port, baseUrl, fixture, args);
        if (args.suite) {
          result.suiteName = args.suite;
          result.suiteExpectedReviewNote = regressionSuites[args.suite].expectedReviewNotes.get(fixture.index);
        }
        results.push(result);
      } finally {
        chrome.process.kill();
        await rm(chrome.userDataDir, { recursive: true, force: true }).catch(() => {});
      }
    }

    printResults(results);
    await writeRegressionReport(args.outputDir, results);
    const failed = buildRegressionSummary(results).rows.filter((row) => row.acceptanceStatus === 'FAIL');
    if (failed.length > 0) process.exitCode = 1;
  } finally {
    if (chromeProcess) chromeProcess.kill();
    if (viteProcess) viteProcess.kill();
    if (chromeUserDataDir) {
      await rm(chromeUserDataDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

async function writeRegressionReport(outputDir, results) {
  await mkdir(outputDir, { recursive: true });
  const reportPath = join(outputDir, 'latest-report.json');
  const htmlPath = join(outputDir, 'latest-report.html');
  const summary = buildRegressionSummary(results);
  await writeFile(
    reportPath,
    JSON.stringify(
      {
        generatedAt: new Date().toISOString(),
        summary,
        results,
      },
      null,
      2,
    ),
  );
  await writeFile(htmlPath, renderHtmlReport(results, summary), 'utf-8');
  console.log(`Report: ${reportPath}`);
  console.log(`HTML report: ${htmlPath}`);
}

function renderHtmlReport(results, summary) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Coco Native Word Regression</title>
  <style>
    body { margin: 0; font-family: Arial, "Microsoft YaHei", sans-serif; background: #eef1f5; color: #1f2937; }
    header { position: sticky; top: 0; z-index: 2; padding: 16px 24px; background: #ffffff; border-bottom: 1px solid #d8dee8; }
    h1 { margin: 0 0 6px; font-size: 20px; }
    .note { margin: 0; color: #64748b; font-size: 13px; }
    main { padding: 20px 24px 40px; }
    section { margin: 0 0 24px; padding: 16px; background: #ffffff; border: 1px solid #d8dee8; border-radius: 6px; }
    h2 { margin: 0 0 12px; font-size: 16px; }
    .status { display: inline-flex; align-items: center; min-width: 58px; justify-content: center; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 12px; }
    .status.pass { color: #14532d; background: #dcfce7; }
    .status.check { color: #7c2d12; background: #ffedd5; }
    .status.fail { color: #7f1d1d; background: #fee2e2; }
    .matrix-wrap { margin: 0 0 24px; overflow-x: auto; background: #ffffff; border: 1px solid #d8dee8; border-radius: 6px; }
    .matrix { width: 100%; border-collapse: collapse; font-size: 13px; }
    .matrix th, .matrix td { padding: 9px 10px; border-bottom: 1px solid #e2e8f0; text-align: left; vertical-align: top; }
    .matrix th { color: #475569; background: #f8fafc; font-size: 12px; }
    .matrix tr:last-child td { border-bottom: 0; }
    .metrics { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 8px; margin: 12px 0; }
    .metric { padding: 8px 10px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; }
    .metric span { display: block; color: #64748b; font-size: 12px; }
    .metric strong { display: block; margin-top: 3px; font-size: 13px; }
    .risks { margin: 10px 0; color: #9a3412; }
    .pages { display: grid; gap: 16px; }
    .page { border-top: 1px solid #e2e8f0; padding-top: 14px; }
    .page-title { margin: 0 0 8px; color: #475569; font-size: 13px; }
    .compare { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; align-items: start; }
    figure { margin: 0; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; overflow: hidden; }
    figcaption { padding: 7px 9px; color: #475569; background: #f1f5f9; border-bottom: 1px solid #e2e8f0; font-size: 12px; }
    img { display: block; width: 100%; height: auto; background: #ffffff; }
    .missing { padding: 80px 12px; text-align: center; color: #64748b; font-size: 13px; }
    @media (max-width: 960px) { .compare { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
  <header>
    <h1>Coco Native Word Regression</h1>
    <p class="note">Acceptance fails on content loss, console errors, or page overflow. Page count drift and missing Word screenshots require review but are not pixel-perfect assertions.</p>
  </header>
  <main>
    ${renderHtmlMatrix(summary)}
    ${results.map(renderHtmlFixture).join('\n')}
  </main>
</body>
</html>`;
}

function renderHtmlMatrix(summary) {
  return `<div class="matrix-wrap">
    <table class="matrix">
      <thead>
        <tr>
          <th>Fixture</th>
          <th>Acceptance</th>
          <th>Content</th>
          <th>Layout</th>
          <th>Pages</th>
          <th>Overflow</th>
          <th>Word refs</th>
          <th>Notes</th>
        </tr>
      </thead>
      <tbody>
        ${summary.rows.map(renderHtmlMatrixRow).join('\n')}
      </tbody>
    </table>
  </div>`;
}

function renderHtmlMatrixRow(row) {
  return `<tr>
    <td>${escapeHtml(`${row.fixtureIndex}: ${row.fixtureName}`)}</td>
    <td><span class="status ${statusClassName(row.acceptanceStatus)}">${escapeHtml(row.acceptanceStatus)}</span></td>
    <td><span class="status ${statusClassName(row.contentStatus)}">${escapeHtml(row.contentStatus)}</span></td>
    <td><span class="status ${statusClassName(row.layoutStatus)}">${escapeHtml(row.layoutStatus)}</span></td>
    <td>${escapeHtml(row.wordPageCount == null ? row.pageCount : `${row.pageCount}/${row.wordPageCount}`)}</td>
    <td>${escapeHtml(row.overflowCount)}</td>
    <td>${escapeHtml(`${row.wordReferenceCount}/${row.wordReferenceTotal}`)}</td>
    <td>${escapeHtml(row.notes.length > 0 ? row.notes.join('; ') : '-')}</td>
  </tr>`;
}

function statusClassName(status) {
  if (status === 'PASS') return 'pass';
  if (status === 'FAIL') return 'fail';
  return 'check';
}

function renderHtmlFixture(result) {
  const statusClass = statusClassName(result.status);
  const metrics = result.metrics
    .filter((metric) => ['Paragraphs', 'Tables', 'Images', 'Anchored Images', 'Comments', 'Style Probe'].includes(metric.label))
    .map((metric) => `<div class="metric"><span>${escapeHtml(metric.label)}</span><strong>${escapeHtml(metric.value)}</strong></div>`)
    .join('');
  const risks =
    result.risks.length > 0
      ? `<ul class="risks">${result.risks.map((risk) => `<li>${escapeHtml(risk)}</li>`).join('')}</ul>`
      : '';
  return `<section>
    <h2><span class="status ${statusClass}">${escapeHtml(result.status)}</span> fixture ${result.fixture.index}: ${escapeHtml(result.fixture.name)}</h2>
    <div class="metrics">${metrics}</div>
    <div class="metric"><span>Page Summary</span><strong>${escapeHtml(summarizePages(result.pageLayouts))}</strong></div>
    ${risks}
    <div class="pages">
      ${renderHtmlPages(result)}
    </div>
  </section>`;
}

function renderHtmlPages(result) {
  const count = Math.max(result.screenshots.length, result.wordReferences.length);
  if (count === 0) return '<div class="missing">No screenshots captured.</div>';
  return Array.from({ length: count }, (_, index) => {
    const native = result.screenshots[index];
    const word = result.wordReferences[index];
    return `<div class="page">
      <p class="page-title">Page ${index + 1}</p>
      <div class="compare">
        ${renderFigure('Microsoft Word reference', word)}
        ${renderFigure('Coco native browser render', native)}
      </div>
    </div>`;
  }).join('\n');
}

function renderFigure(label, file) {
  if (!file) {
    return `<figure><figcaption>${escapeHtml(label)}</figcaption><div class="missing">No image available</div></figure>`;
  }
  return `<figure><figcaption>${escapeHtml(label)}</figcaption><img src="${escapeHtml(pathToFileUrl(file))}" alt="${escapeHtml(label)}"></figure>`;
}

function pathToFileUrl(path) {
  return pathToFileURL(path).href;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

await main();
