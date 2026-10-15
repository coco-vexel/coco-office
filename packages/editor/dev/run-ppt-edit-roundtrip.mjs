import { spawn } from 'node:child_process';
import http from 'node:http';
import { mkdir, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function parseArgs(argv) {
  const parsed = {
    baseUrl: undefined,
    caseName: 'default',
    compareSlides: [],
    fixtureIndex: 0,
    fixtureIndexes: undefined,
    officeVerify: true,
    outputDir: resolve(packageDir, 'dev/native-regression-output/ppt-edit-roundtrip'),
    port: 5177,
    startServer: false,
  };
  for (const arg of argv) {
    if (arg === '--') {
      continue;
    } else if (arg === '--start-server') {
      parsed.startServer = true;
    } else if (arg.startsWith('--base-url=')) {
      parsed.baseUrl = arg.slice('--base-url='.length).replace(/\/$/, '');
    } else if (arg.startsWith('--case=')) {
      parsed.caseName = arg.slice('--case='.length);
    } else if (arg.startsWith('--compare-slides=')) {
      parsed.compareSlides = parseSlideIndexes(arg.slice('--compare-slides='.length));
    } else if (arg.startsWith('--fixture=')) {
      parsed.fixtureIndex = Number(arg.slice('--fixture='.length));
    } else if (arg.startsWith('--fixtures=')) {
      parsed.fixtureIndexes = parseFixtureIndexes(arg.slice('--fixtures='.length));
    } else if (arg === '--no-office-verify') {
      parsed.officeVerify = false;
    } else if (arg.startsWith('--output-dir=')) {
      parsed.outputDir = resolve(packageDir, arg.slice('--output-dir='.length));
    } else if (arg.startsWith('--port=')) {
      parsed.port = Number(arg.slice('--port='.length));
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!Number.isInteger(parsed.fixtureIndex) || parsed.fixtureIndex < 0) {
    throw new Error('Invalid --fixture value.');
  }
  if (
    Array.isArray(parsed.fixtureIndexes) &&
    parsed.fixtureIndexes.some((index) => !Number.isInteger(index) || index < 0)
  ) {
    throw new Error('Invalid --fixtures value.');
  }
  if (
    !['default', 'group-child', 'group-parent-drag', 'group-table-cell', 'group-text'].includes(
      parsed.caseName,
    )
  ) {
    throw new Error(`Invalid --case value: ${parsed.caseName}`);
  }
  if (!Number.isInteger(parsed.port) || parsed.port <= 0) throw new Error('Invalid --port value.');
  return parsed;
}

function parseSlideIndexes(raw) {
  return String(raw ?? '')
    .split(',')
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const slideNumber = Number(part);
      if (!Number.isInteger(slideNumber) || slideNumber <= 0) {
        throw new Error(`Invalid --compare-slides value: ${part}`);
      }
      return slideNumber - 1;
    });
}

function parseFixtureIndexes(raw) {
  if (raw === 'all') return 'all';
  const indexes = [];
  for (const part of raw.split(',')) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const range = trimmed.match(/^(\d+)-(\d+)$/);
    if (range) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      if (end < start) throw new Error(`Invalid --fixtures range: ${trimmed}`);
      for (let index = start; index <= end; index += 1) indexes.push(index);
      continue;
    }
    indexes.push(Number(trimmed));
  }
  return [...new Set(indexes)];
}

async function resolveFixtureIndexes(baseUrl, args) {
  if (args.fixtureIndexes === 'all') {
    const response = await fetch(`${baseUrl}/__coco_fixture__/pptx-files`);
    if (!response.ok) throw new Error(`Failed to list ppt fixtures: ${response.status}`);
    const fixtures = await response.json();
    if (!Array.isArray(fixtures)) throw new Error('PPT fixtures response is not an array.');
    return fixtures.map((_, index) => index);
  }
  if (Array.isArray(args.fixtureIndexes)) return args.fixtureIndexes;
  return [args.fixtureIndex];
}

function isSkippedFixtureError(message) {
  return /No (fully editable text shape|movable parent group|editable flattened group child|editable grouped table cell) found/.test(
    message,
  );
}

async function startVite(port) {
  const viteEntry = resolve(packageDir, 'node_modules/vite/bin/vite.js');
  const child = spawn(
    process.execPath,
    [viteEntry, '--host', '127.0.0.1', '--port', String(port), '--strictPort'],
    {
      cwd: packageDir,
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );
  child.stdout.on('data', (chunk) => process.stdout.write(`[vite] ${chunk}`));
  child.stderr.on('data', (chunk) => process.stderr.write(`[vite] ${chunk}`));
  await waitForHttp(`http://127.0.0.1:${port}/__coco_fixture__/pptx-files`, 20_000);
  return child;
}

async function startChrome() {
  const port = await findFreePort(9222);
  const userDataDir = join(
    tmpdir(),
    `coco-ppt-edit-roundtrip-${Date.now()}-${Math.random().toString(16).slice(2)}`,
  );
  await mkdir(userDataDir, { recursive: true });
  const child = spawn(
    chromeExecutable(),
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
  if (process.platform === 'win32')
    return 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
  if (process.platform === 'darwin')
    return '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  return 'google-chrome';
}

async function runRoundtrip(chromePort, baseUrl, options) {
  const pageInfo = await createCdpPage(chromePort);
  const client = await CdpClient.connect(pageInfo.webSocketDebuggerUrl);
  try {
    await client.send('Page.enable');
    await client.send('Runtime.enable');
    await client.send('Emulation.setDeviceMetricsOverride', {
      width: 1440,
      height: 900,
      deviceScaleFactor: 1,
      mobile: false,
    });
    const url = `${baseUrl}/?surface=ppt&pptFixture=${options.fixtureIndex}&bust=ppt-edit-roundtrip-${Date.now()}`;
    await client.send('Page.navigate', { url });
    await waitForPptReady(client);

    await evaluate(client, selectTargetsExpression(options.caseName));
    await evaluate(client, applyRoundtripEditsExpression(options.caseName));
    const result = await evaluate(client, collectRoundtripResultExpression(options.caseName));

    const runId = new Date().toISOString().replace(/[:.]/g, '-');
    await mkdir(options.outputDir, { recursive: true });
    const downloadDir = join(options.outputDir, `downloads-${runId}`);
    await mkdir(downloadDir, { recursive: true });
    await allowDownloads(client, downloadDir);
    const screenshotPath = join(
      options.outputDir,
      `fixture-${options.fixtureIndex}-browser-edited-${runId}.png`,
    );
    const reloadScreenshotPath = join(
      options.outputDir,
      `fixture-${options.fixtureIndex}-export-reloaded-${runId}.png`,
    );
    const summaryPath = join(options.outputDir, `summary-${runId}.json`);
    await evaluate(client, `document.querySelector('#pptExportButton').click()`);
    const pptxPath = await waitForDownloadedPptx(downloadDir, 60_000);
    const screenshot = await client.send('Page.captureScreenshot', {
      format: 'png',
      fromSurface: true,
    });
    await writeFile(screenshotPath, Buffer.from(screenshot.data, 'base64'));
    const targetedBrowserScreenshots = await captureTargetSlideScreenshots(
      client,
      options,
      runId,
      'browser-edited',
    );
    const exportedBytes = await readFile(pptxPath);
    await evaluate(
      client,
      reloadExportedPptxExpression(exportedBytes.toString('base64'), options.caseName),
    );
    const reloadedAfterExport = await evaluate(
      client,
      collectReloadedRoundtripResultExpression(options.caseName),
    );
    const reloadScreenshot = await client.send('Page.captureScreenshot', {
      format: 'png',
      fromSurface: true,
    });
    await writeFile(reloadScreenshotPath, Buffer.from(reloadScreenshot.data, 'base64'));
    const targetedReloadScreenshots = await captureTargetSlideScreenshots(
      client,
      options,
      runId,
      'export-reloaded',
    );
    const visualReloadComparison = await compareRoundtripScreenshots(
      screenshotPath,
      reloadScreenshotPath,
    );
    const officeVerification = options.officeVerify
      ? await verifyPptxWithPowerPoint(pptxPath, {
          outputDir: options.outputDir,
          runId,
          slides: targetedReloadScreenshots.map((item) => ({
            slideIndex: item.slideIndex,
            width: item.canvasRect?.width,
            height: item.canvasRect?.height,
          })),
        })
      : { skipped: true };
    const referenceVisualComparisons = await compareReferenceSlideScreenshots(
      targetedReloadScreenshots,
      officeVerification?.slideImages ?? [],
    );
    await writeFile(
      summaryPath,
      JSON.stringify(
        {
          ...result,
          reloadedAfterExport,
          officeVerification,
          url,
          pptxPath,
          screenshotPath,
          reloadScreenshotPath,
          visualReloadComparison,
          targetedBrowserScreenshots,
          targetedReloadScreenshots,
          referenceVisualComparisons,
          summaryPath,
        },
        null,
      ),
    );
    return {
      ...result,
      reloadedAfterExport,
      officeVerification,
      url,
      pptxPath,
      screenshotPath,
      reloadScreenshotPath,
      visualReloadComparison,
      targetedBrowserScreenshots,
      targetedReloadScreenshots,
      referenceVisualComparisons,
      summaryPath,
    };
  } finally {
    await client.send('Page.close').catch(() => undefined);
    client.close();
  }
}

async function waitForPptReady(client) {
  const start = Date.now();
  while (Date.now() - start < 60_000) {
    const ready = await evaluate(
      client,
      `Boolean(window.__cocoPptDev && window.__cocoPptDev.session().scene().slides.length > 0 && document.querySelector('#pptApplyTextButton'))`,
    ).catch(() => false);
    if (ready) {
      await sleep(500);
      return;
    }
    await sleep(250);
  }
  throw new Error('Timed out waiting for PPT editor dev page.');
}

async function captureTargetSlideScreenshots(client, options, runId, phase) {
  const screenshots = [];
  for (const slideIndex of options.compareSlides ?? []) {
    const selected = await evaluate(
      client,
      `(${selectSlideForScreenshotExpression.toString()})(${JSON.stringify(slideIndex)})`,
    );
    const rect = selected.canvasRect;
    const path = join(
      options.outputDir,
      `fixture-${options.fixtureIndex}-slide-${String(slideIndex + 1).padStart(3, '0')}-${phase}-${runId}.png`,
    );
    const shot = await client.send('Page.captureScreenshot', {
      format: 'png',
      fromSurface: true,
      clip: {
        x: rect.x,
        y: rect.y,
        width: rect.width,
        height: rect.height,
        scale: 1,
      },
    });
    await writeFile(path, Buffer.from(shot.data, 'base64'));
    screenshots.push({
      slideIndex,
      slideNumber: slideIndex + 1,
      path,
      canvasRect: rect,
      selected,
    });
  }
  return screenshots;
}

async function compareReferenceSlideScreenshots(browserScreenshots, slideImages) {
  const comparisons = [];
  for (const browserShot of browserScreenshots) {
    const reference = slideImages.find((item) => item.slideIndex === browserShot.slideIndex);
    if (!reference?.path) {
      comparisons.push({
        slideIndex: browserShot.slideIndex,
        slideNumber: browserShot.slideNumber,
        status: 'REVIEW',
        pass: true,
        skipped: true,
        warnings: ['PowerPoint reference slide image was not exported.'],
        failures: [],
        browserPath: browserShot.path,
        referencePath: reference?.path ?? null,
      });
      continue;
    }
    const comparison = await compareRoundtripScreenshots(browserShot.path, reference.path);
    comparisons.push({
      slideIndex: browserShot.slideIndex,
      slideNumber: browserShot.slideNumber,
      ...comparison,
      browserPath: browserShot.path,
      referencePath: reference.path,
    });
  }
  return comparisons;
}

function selectSlideForScreenshotExpression(slideIndex) {
  return (async () => {
    const editor = window.__cocoPptDev;
    const slide = editor.session().scene().slides[slideIndex];
    if (!slide) throw new Error(`PPT compare slide not found: ${slideIndex + 1}`);
    await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
    const canvas = document.querySelector('#pptEditorHost')?.shadowRoot?.querySelector('.coco-ppt-canvas');
    const start = Date.now();
    while (Date.now() - start < 5_000) {
      const selection = editor.selection();
      const renderedNodes = canvas?.querySelectorAll('.coco-ppt-node').length ?? 0;
      if (selection?.slideId === slide.id && renderedNodes === slide.nodes.length) break;
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
    const rect = canvas?.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) {
      throw new Error(`PPT compare slide canvas has empty bounds: ${slideIndex + 1}`);
    }
    return {
      slideId: slide.id,
      slideIndex,
      slideNumber: slideIndex + 1,
      sceneNodes: slide.nodes.length,
      domNodes: canvas?.querySelectorAll('.coco-ppt-node').length ?? 0,
      canvasRect: {
        x: Math.round(rect.x),
        y: Math.round(rect.y),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      },
    };
  })();
}

function selectTargetsExpression(caseName) {
  if (caseName === 'group-child') return selectGroupChildTargetsExpression();
  if (caseName === 'group-parent-drag') return selectGroupParentDragTargetsExpression();
  if (caseName === 'group-table-cell') return selectGroupTableCellTargetsExpression();
  if (caseName === 'group-text') return selectGroupTextTargetsExpression();
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    const isStableTextTarget = (node) =>
      !node.groupContext &&
      node.capabilities?.editText &&
      node.capabilities?.editTextStyle &&
      node.capabilities?.editShapeStyle &&
      node.capabilities?.editGeometry;
    const textTargetScore = (slide, node) => {
      if (!isStableTextTarget(node)) return -Infinity;
      const text = String(node.text ?? '').trim();
      if (!text || text.length > 140) return -Infinity;
      if (/^[0-9a-f]{48,}$/i.test(text.replace(/\\s+/g, ''))) return -Infinity;
      const bounds = node.bounds;
      const size = slide.size;
      if (!bounds || !size || bounds.width <= 0 || bounds.height <= 0) return -Infinity;
      const visibleWidth = Math.min(bounds.left + bounds.width, size.width) - Math.max(bounds.left, 0);
      const visibleHeight = Math.min(bounds.top + bounds.height, size.height) - Math.max(bounds.top, 0);
      if (visibleWidth <= 0 || visibleHeight <= 0) return -Infinity;
      const visibleRatio = (visibleWidth * visibleHeight) / (bounds.width * bounds.height);
      if (visibleRatio < 0.7) return -Infinity;
      const areaRatio = (bounds.width * bounds.height) / (size.width * size.height);
      if (areaRatio < 0.0005 || areaRatio > 0.55) return -Infinity;
      let score = 0;
      score += Math.min(text.length, 60);
      if (/[\\u4e00-\\u9fff]{2,}|[A-Za-z]{3,}/.test(text)) score += 30;
      if (bounds.left >= 0 && bounds.top >= 0 && bounds.left + bounds.width <= size.width && bounds.top + bounds.height <= size.height) {
        score += 20;
      }
      score -= Math.abs(areaRatio - 0.04) * 100;
      return score;
    };
    const textCandidates = scene.slides.flatMap((slide) =>
      slide.nodes.map((node) => ({ slide, node, score: textTargetScore(slide, node) }))
    ).filter((candidate) => Number.isFinite(candidate.score))
      .sort((a, b) => b.score - a.score);
    const textSlide = textCandidates[0]?.slide ?? scene.slides[0];
    const textNode = textCandidates[0]?.node;
    if (!textNode) throw new Error('No fully editable text shape found');
    const firstEditableTableCell = (node) => {
      if (
        node.kind !== 'table' ||
        !node.capabilities?.editTableCells ||
        !node.table?.rowCount ||
        !node.table?.columnCount
      ) return null;
      for (let row = 0; row < node.table.rowCount; row += 1) {
        for (let column = 0; column < node.table.columnCount; column += 1) {
          const text = node.table.values[row]?.[column] ?? '';
          if (String(text).trim()) return { row, column, text };
        }
      }
      return null;
    };
    const tableSlide = scene.slides.find((slide) => slide.nodes.some((node) =>
      Boolean(firstEditableTableCell(node))
    ));
    const tableNode = tableSlide?.nodes.find((node) =>
      Boolean(firstEditableTableCell(node))
    );
    const tableCell = tableNode ? firstEditableTableCell(tableNode) : null;
    window.__cocoPptRoundtripTargets = {
      fixture: {
        summary: document.querySelector('#pptSummary')?.textContent ?? '',
      },
      text: {
        slideId: textSlide.id,
        shapeId: textNode.id,
        beforeText: textNode.text,
        beforeBounds: textNode.bounds,
      },
      table: tableNode ? {
        slideId: tableSlide.id,
        shapeId: tableNode.id,
        row: tableCell.row,
        column: tableCell.column,
        beforeText: tableCell.text,
      } : null,
    };
    return window.__cocoPptRoundtripTargets;
  })()`;
}

function selectGroupTableCellTargetsExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    for (const slide of scene.slides) {
      const node = slide.nodes.find((candidate) =>
        candidate.groupContext &&
        candidate.kind === 'table' &&
        candidate.capabilities?.editTableCells &&
        candidate.table?.rowCount > 0 &&
        candidate.table?.columnCount > 0
      );
      if (!node) continue;
      let cell = null;
      for (let row = 0; row < node.table.rowCount && !cell; row += 1) {
        for (let column = 0; column < node.table.columnCount; column += 1) {
          const text = node.table.values[row]?.[column] ?? '';
          if (String(text).trim()) {
            cell = { row, column, text };
            break;
          }
        }
      }
      cell ??= { row: 0, column: 0, text: node.table.values[0]?.[0] ?? '' };
      window.__cocoPptRoundtripTargets = {
        fixture: {
          summary: document.querySelector('#pptSummary')?.textContent ?? '',
        },
        groupTableCell: {
          slideId: slide.id,
          shapeId: node.id,
          parentId: node.groupContext.parentId,
          parentName: node.groupContext.parentName,
          beforeBounds: node.bounds,
          beforeParentBounds: node.groupContext.parentBounds,
          row: cell.row,
          column: cell.column,
          beforeText: cell.text,
          capabilities: node.capabilities,
        },
      };
      return window.__cocoPptRoundtripTargets;
    }
    throw new Error('No editable grouped table cell found');
  })()`;
}

function selectGroupTextTargetsExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    for (const slide of scene.slides) {
      const node = slide.nodes.find((candidate) =>
        candidate.groupContext &&
        candidate.capabilities?.editText &&
        candidate.capabilities?.editTextStyle &&
        String(candidate.text ?? '').trim()
      );
      if (!node) continue;
      window.__cocoPptRoundtripTargets = {
        fixture: {
          summary: document.querySelector('#pptSummary')?.textContent ?? '',
        },
        groupText: {
          slideId: slide.id,
          shapeId: node.id,
          parentId: node.groupContext.parentId,
          parentName: node.groupContext.parentName,
          beforeText: node.text,
          beforeBounds: node.bounds,
          beforeParentBounds: node.groupContext.parentBounds,
          capabilities: node.capabilities,
        },
      };
      return window.__cocoPptRoundtripTargets;
    }
    throw new Error('No editable grouped text found');
  })()`;
}

function selectGroupParentDragTargetsExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    const hasMoveRoom = (slide, bounds) =>
      bounds.left >= 0 &&
      bounds.top >= 0 &&
      bounds.width > 0 &&
      bounds.height > 0 &&
      bounds.left + bounds.width <= slide.size.width - 420000 &&
      bounds.top + bounds.height <= slide.size.height - 360000;
    for (const slide of scene.slides) {
      const node = slide.nodes.find((candidate) =>
        candidate.groupContext &&
        !candidate.groupContext.parentId.includes(':') &&
        candidate.capabilities?.editGeometry &&
        hasMoveRoom(slide, candidate.groupContext.parentBounds)
      );
      if (!node) continue;
      const parent = slide.groups.find((candidate) => candidate.id === node.groupContext.parentId);
      window.__cocoPptRoundtripTargets = {
        fixture: {
          summary: document.querySelector('#pptSummary')?.textContent ?? '',
        },
        groupParent: {
          slideId: slide.id,
          shapeId: node.groupContext.parentId,
          shapeName: node.groupContext.parentName,
          childShapeId: node.id,
          childKind: node.kind,
          beforeBounds: parent?.bounds ?? node.groupContext.parentBounds,
          beforeChildBounds: node.bounds,
          capabilities: parent?.capabilities ?? {
            select: true,
            editText: false,
            editTextStyle: false,
            editTableCells: false,
            editGeometry: true,
            editShapeStyle: false,
            reason: 'Parent group',
          },
        },
      };
      return window.__cocoPptRoundtripTargets;
    }
    throw new Error('No movable parent group found for a flattened child');
  })()`;
}

function selectGroupChildTargetsExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    for (const slide of scene.slides) {
      const node = slide.nodes.find((candidate) =>
        candidate.groupContext &&
        candidate.capabilities?.editGeometry &&
        !candidate.id.startsWith('layout:') &&
        !candidate.id.startsWith('master:')
      );
      if (!node) continue;
      window.__cocoPptRoundtripTargets = {
        fixture: {
          summary: document.querySelector('#pptSummary')?.textContent ?? '',
        },
        groupChild: {
          slideId: slide.id,
          shapeId: node.id,
          nodeKind: node.kind,
          parentId: node.groupContext.parentId,
          parentName: node.groupContext.parentName,
          beforeBounds: node.bounds,
          beforeParentBounds: node.groupContext.parentBounds,
          capabilities: node.capabilities,
        },
      };
      return window.__cocoPptRoundtripTargets;
    }
    throw new Error('No editable flattened group child found');
  })()`;
}

function applyRoundtripEditsExpression(caseName) {
  if (caseName === 'group-child') return applyGroupChildRoundtripEditsExpression();
  if (caseName === 'group-parent-drag') return applyGroupParentDragRoundtripEditsExpression();
  if (caseName === 'group-table-cell') return applyGroupTableCellRoundtripEditsExpression();
  if (caseName === 'group-text') return applyGroupTextRoundtripEditsExpression();
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const textInput = document.querySelector('#pptTextInput');
    const fontSizeInput = document.querySelector('#pptFontSizeInput');
    const textColorInput = document.querySelector('#pptTextColorInput');
    const fillColorInput = document.querySelector('#pptFillColorInput');
    const lineColorInput = document.querySelector('#pptLineColorInput');
    const lineWidthInput = document.querySelector('#pptLineWidthInput');
    const leftInput = document.querySelector('#pptLeftInput');
    const topInput = document.querySelector('#pptTopInput');
    const widthInput = document.querySelector('#pptWidthInput');
    const heightInput = document.querySelector('#pptHeightInput');
    const tableRowInput = document.querySelector('#pptTableRowInput');
    const tableColumnInput = document.querySelector('#pptTableColumnInput');
    const tableCellTextInput = document.querySelector('#pptTableCellTextInput');
    const setInput = (input, value) => {
      input.value = String(value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    };
    const waitApplied = async (step) => {
      targets.currentStep = step;
      const start = Date.now();
      while (Date.now() - start < 30_000) {
        const status = document.querySelector('#pptEditStatus')?.textContent ?? '';
        if (status === 'Applied' || status === '已应用到内存') return;
        if (status && status !== '正在应用...' && status !== 'Applying...') {
          throw new Error('Edit failed during ' + step + ' for shape ' + targets.text.shapeId + ': ' + status);
        }
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error(
        'Timed out waiting for Applied status during ' +
          step +
          ' for shape ' +
          targets.text.shapeId +
          '; last status: ' +
          (document.querySelector('#pptEditStatus')?.textContent ?? '')
      );
    };

    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: targets.text.slideId,
      shapeIdOrName: targets.text.shapeId,
    });
    setInput(textInput, 'P2.3 TEXT ROUNDTRIP');
    document.querySelector('#pptApplyTextButton').click();
    await waitApplied('text');
    setInput(fontSizeInput, '30');
    setInput(textColorInput, '#336699');
    setInput(fillColorInput, '#ffcc00');
    setInput(lineColorInput, '#2255aa');
    setInput(lineWidthInput, '19050');
    document.querySelector('#pptApplyStyleButton').click();
    await waitApplied('style');
    setInput(leftInput, Math.round(targets.text.beforeBounds.left + 100000));
    setInput(topInput, Math.round(targets.text.beforeBounds.top + 100000));
    setInput(widthInput, Math.round(Math.max(1, targets.text.beforeBounds.width + 120000)));
    setInput(heightInput, Math.round(Math.max(1, targets.text.beforeBounds.height + 80000)));
    targets.expected = {
      text: 'P2.3 TEXT ROUNDTRIP',
      fontSize: 30,
      textColor: '#336699',
      fillColor: '#ffcc00',
      lineColor: '#2255aa',
      lineWidth: 19050,
      bounds: {
        left: Math.round(targets.text.beforeBounds.left + 100000),
        top: Math.round(targets.text.beforeBounds.top + 100000),
        width: Math.round(Math.max(1, targets.text.beforeBounds.width + 120000)),
        height: Math.round(Math.max(1, targets.text.beforeBounds.height + 80000)),
      },
      tableText: targets.table ? 'P2.3 TABLE ROUNDTRIP' : null,
    };
    document.querySelector('#pptApplyGeometryButton').click();
    await waitApplied('geometry');

    if (targets.table) {
      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: targets.table.slideId,
        shapeIdOrName: targets.table.shapeId,
      });
      setInput(tableRowInput, targets.table.row);
      setInput(tableColumnInput, targets.table.column);
      setInput(tableCellTextInput, 'P2.3 TABLE ROUNDTRIP');
      document.querySelector('#pptApplyTableCellButton').click();
      await waitApplied('table');
      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: targets.text.slideId,
        shapeIdOrName: targets.text.shapeId,
      });
    }

    document.querySelector('#pptVerifyExportButton').click();
    const start = Date.now();
    while (Date.now() - start < 30_000) {
      const status = document.querySelector('#pptEditStatus')?.textContent ?? '';
      if (/^(Verified \\d+ slides|导出校验通过：\\d+ 页)$/.test(status)) return true;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    throw new Error('Timed out waiting for verifyExport status');
  })()`;
}

function applyGroupTableCellRoundtripEditsExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupTableCell;
    const expectedText = 'P2.4 GROUP TABLE ROUNDTRIP';
    const waitFor = async (predicate, label, timeoutMs = 30_000) => {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (predicate()) return;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error('Timed out waiting for ' + label);
    };
    const findNode = () => {
      const slide = editor.session().scene().slides.find((item) => item.id === target.slideId);
      return slide?.nodes.find((item) => item.id === target.shapeId);
    };
    const cellElement = () =>
      document
        .querySelector('#pptEditorHost')
        ?.shadowRoot?.querySelector(
          '.coco-ppt-node[data-shape-id="' +
            CSS.escape(target.shapeId) +
            '"] td[data-row="' +
            target.row +
            '"][data-column="' +
            target.column +
            '"]',
        );
    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    await waitFor(() => {
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const parent = root?.querySelector('.coco-ppt-parent-group-outline');
      return parent?.dataset.shapeId === target.parentId &&
        Boolean(parent.querySelector('.coco-ppt-parent-group-handle'));
    }, 'parent group affordance');
    const cellBefore = cellElement();
    if (!cellBefore) throw new Error('Grouped table cell DOM is unavailable');
    cellBefore.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, composed: true }));
    await waitFor(
      () =>
        document.querySelector('#pptTableRowInput')?.value === String(target.row) &&
        document.querySelector('#pptTableColumnInput')?.value === String(target.column) &&
        document.querySelector('#pptTableCellTextInput')?.value === (cellBefore.textContent ?? ''),
      'grouped table cell toolbar sync',
    );
    const tableButton = document.querySelector('#pptApplyTableCellButton');
    if (!tableButton || tableButton.disabled) throw new Error('Grouped table cell toolbar is disabled');
    const cell = cellElement();
    if (!cell) throw new Error('Grouped table cell DOM disappeared before edit');
    cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, composed: true }));
    await waitFor(() => cellElement()?.contentEditable === 'true', 'grouped table cell direct editing');
    const editingCell = cellElement();
    editingCell.focus();
    editingCell.textContent = expectedText;
    editingCell.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    targets.expected = {
      tableText: expectedText,
      parentId: target.parentId,
    };
    await waitFor(() => {
      const node = findNode();
      return (node?.table?.values[target.row]?.[target.column] ?? '') === expectedText;
    }, 'grouped table cell edit');
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \\d+ slides|导出校验通过：\\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'verifyExport status',
      60_000,
    );
    return true;
  })()`;
}

function applyGroupTextRoundtripEditsExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupText;
    const expectedText = 'P2.4 GROUP TEXT ROUNDTRIP';
    const waitFor = async (predicate, label, timeoutMs = 30_000) => {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (predicate()) return;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error('Timed out waiting for ' + label);
    };
    const findNode = () => {
      const slide = editor.session().scene().slides.find((item) => item.id === target.slideId);
      return slide?.nodes.find((item) => item.id === target.shapeId);
    };
    const nodeElement = () =>
      document
        .querySelector('#pptEditorHost')
        ?.shadowRoot?.querySelector('.coco-ppt-node[data-shape-id="' + CSS.escape(target.shapeId) + '"]');

    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    await waitFor(() => {
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const parent = root?.querySelector('.coco-ppt-parent-group-outline');
      return parent?.dataset.shapeId === target.parentId &&
        Boolean(parent.querySelector('.coco-ppt-parent-group-handle'));
    }, 'parent group affordance');
    const textInput = document.querySelector('#pptTextInput');
    const textButton = document.querySelector('#pptApplyTextButton');
    const fontSizeInput = document.querySelector('#pptFontSizeInput');
    const textColorInput = document.querySelector('#pptTextColorInput');
    if (!textInput || !textButton || textButton.disabled || textInput.disabled) {
      throw new Error('Group text toolbar is disabled or unavailable');
    }
    if (fontSizeInput?.disabled || textColorInput?.disabled) {
      throw new Error('Group text style toolbar is disabled or unavailable');
    }
    const inspectorText = document.querySelector('#pptInspector')?.textContent ?? '';
    if (!inspectorText.includes('父组') || !inspectorText.includes('文字：可用')) {
      throw new Error('Group text inspector did not expose parent group and text capability state');
    }
    const node = nodeElement();
    if (!node) throw new Error('Group text DOM is unavailable');
    if (node.dataset.draggable !== 'true' || node.dataset.resizable !== 'true') {
      throw new Error('Group text canvas geometry affordance is unavailable');
    }
    node.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true, composed: true }));
    await waitFor(() => nodeElement()?.contentEditable === 'true' || Boolean(nodeElement()?.querySelector('[contenteditable="true"]')), 'group text direct editing');
    const editingTarget =
      nodeElement()?.querySelector('[contenteditable="true"]') ??
      nodeElement();
    editingTarget.focus();
    editingTarget.textContent = expectedText;
    editingTarget.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    targets.expected = {
      text: expectedText,
      parentId: target.parentId,
    };
    await waitFor(() => findNode()?.text === expectedText, 'group text direct edit');
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \\d+ slides|导出校验通过：\\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'verifyExport status',
      60_000,
    );
    return true;
  })()`;
}

function applyGroupParentDragRoundtripEditsExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupParent;
    const waitFor = async (predicate, label, timeoutMs = 30_000) => {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (predicate()) return;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error('Timed out waiting for ' + label);
    };
    const sceneSlide = () => editor.session().scene().slides.find((item) => item.id === target.slideId);
    const findGroup = () => sceneSlide()?.groups.find((item) => item.id === target.shapeId);
    const findChild = () => sceneSlide()?.nodes.find((item) => item.id === target.childShapeId);
    const mouse = (type, x, y) => new MouseEvent(type, {
      bubbles: true,
      cancelable: true,
      clientX: x,
      clientY: y,
      button: 0,
      altKey: true,
    });

    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.childShapeId,
    });
    await waitFor(() => {
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const parent = root?.querySelector('.coco-ppt-parent-group-outline');
      return parent?.dataset.shapeId === target.shapeId &&
        Boolean(parent.querySelector('.coco-ppt-parent-group-handle')) &&
        parent.querySelectorAll('.coco-ppt-parent-group-edge').length === 4;
    }, 'parent group affordance before drag');
    const root = document.querySelector('#pptEditorHost').shadowRoot;
    const parentOutline = root.querySelector('.coco-ppt-parent-group-outline');
    const dragTarget = parentOutline.querySelector('.coco-ppt-parent-group-edge[data-edge="right"]');
    const rect = dragTarget.getBoundingClientRect();
    const outlineRect = parentOutline.getBoundingClientRect();
    const scale = outlineRect.width / target.beforeBounds.width;
    const dxPx = 28;
    const dyPx = 22;
    const startX = rect.left + rect.width / 2;
    const startY = rect.top + rect.height / 2;
    targets.dragIntent = {
      deltaPx: { x: dxPx, y: dyPx },
      estimatedDelta: { x: dxPx / scale, y: dyPx / scale },
    };
    dragTarget.dispatchEvent(mouse('mousedown', startX, startY));
    window.dispatchEvent(mouse('mousemove', startX + dxPx, startY + dyPx));
    window.dispatchEvent(mouse('mouseup', startX + dxPx, startY + dyPx));
    await waitFor(() => {
      const group = findGroup();
      const child = findChild();
      if (!group || !child) return false;
      const groupMoved = Math.hypot(group.bounds.left - target.beforeBounds.left, group.bounds.top - target.beforeBounds.top);
      const childMoved = Math.hypot(child.bounds.left - target.beforeChildBounds.left, child.bounds.top - target.beforeChildBounds.top);
      return groupMoved > 1000 && childMoved > 1000;
    }, 'parent group drag commit');
    const groupAfterDrag = findGroup();
    const childAfterDrag = findChild();
    targets.expected = {
      parentBounds: groupAfterDrag.bounds,
      childBounds: childAfterDrag.bounds,
    };
    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \\d+ slides|导出校验通过：\\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'verifyExport status',
      60_000,
    );
    return true;
  })()`;
}

function applyGroupChildRoundtripEditsExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupChild;
    const waitFor = async (predicate, label, timeoutMs = 30_000) => {
      const start = Date.now();
      while (Date.now() - start < timeoutMs) {
        if (predicate()) return;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error('Timed out waiting for ' + label);
    };
    const findNode = () => {
      const scene = editor.session().scene();
      const slide = scene.slides.find((item) => item.id === target.slideId);
      return slide?.nodes.find((item) => item.id === target.shapeId);
    };
    const setInput = (input, value) => {
      input.value = String(value);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    };
    const waitApplied = async (step) => {
      targets.currentStep = step;
      const start = Date.now();
      while (Date.now() - start < 30_000) {
        const status = document.querySelector('#pptEditStatus')?.textContent ?? '';
        if (status === 'Applied' || status === '已应用到内存') return;
        if (status && status !== '正在应用...' && status !== 'Applying...') {
          throw new Error('Group child edit failed during ' + step + ': ' + status);
        }
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      throw new Error(
        'Timed out waiting for Applied status during ' +
          step +
          '; last status: ' +
          (document.querySelector('#pptEditStatus')?.textContent ?? ''),
      );
    };

    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    await waitFor(() => {
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const parent = root?.querySelector('.coco-ppt-parent-group-outline');
      return parent?.dataset.shapeId === target.parentId &&
        Boolean(parent.querySelector('.coco-ppt-parent-group-handle')) &&
        parent.querySelectorAll('.coco-ppt-parent-group-edge').length === 4;
    }, 'parent group affordance');
    const leftInput = document.querySelector('#pptLeftInput');
    const topInput = document.querySelector('#pptTopInput');
    const widthInput = document.querySelector('#pptWidthInput');
    const heightInput = document.querySelector('#pptHeightInput');
    const geometryButton = document.querySelector('#pptApplyGeometryButton');
    if (
      !leftInput ||
      !topInput ||
      !widthInput ||
      !heightInput ||
      !geometryButton ||
      leftInput.disabled ||
      topInput.disabled ||
      widthInput.disabled ||
      heightInput.disabled ||
      geometryButton.disabled
    ) {
      throw new Error('Group child geometry toolbar is disabled or unavailable');
    }
    const inspectorText = document.querySelector('#pptInspector')?.textContent ?? '';
    if (!inspectorText.includes('父组') || !inspectorText.includes('位置尺寸：可用')) {
      throw new Error('Group child inspector did not expose parent group and geometry capability state');
    }
    const childElement = document
      .querySelector('#pptEditorHost')
      ?.shadowRoot?.querySelector(
        '.coco-ppt-node[data-shape-id="' + CSS.escape(target.shapeId) + '"]',
      );
    if (childElement?.dataset.draggable !== 'true' || childElement?.dataset.resizable !== 'true') {
      throw new Error('Group child canvas geometry affordance is unavailable');
    }

    const expectedBounds = {
      left: Math.round(target.beforeBounds.left + 90000),
      top: Math.round(target.beforeBounds.top + 70000),
      width: Math.round(Math.max(1, target.beforeBounds.width + 60000)),
      height: Math.round(Math.max(1, target.beforeBounds.height + 50000)),
      rotation: target.beforeBounds.rotation,
      flipHorizontal: target.beforeBounds.flipHorizontal,
      flipVertical: target.beforeBounds.flipVertical,
    };
    targets.expected = {
      bounds: expectedBounds,
      parentId: target.parentId,
    };
    setInput(leftInput, expectedBounds.left);
    setInput(topInput, expectedBounds.top);
    setInput(widthInput, expectedBounds.width);
    setInput(heightInput, expectedBounds.height);
    geometryButton.click();
    await waitApplied('group child geometry toolbar');
    try {
      await waitFor(() => {
      const node = findNode();
      if (!node) return false;
      return ['left', 'top', 'width', 'height'].every((key) =>
        Math.abs((node.bounds?.[key] ?? NaN) - expectedBounds[key]) <= 1
      );
      }, 'group child geometry edit');
    } catch (error) {
      const node = findNode();
      throw new Error(JSON.stringify({
        reason: error instanceof Error ? error.message : String(error),
        status: document.querySelector('#pptEditStatus')?.textContent ?? '',
        shapeId: target.shapeId,
        expectedBounds,
        actualBounds: node?.bounds ?? null,
        groupContext: node?.groupContext ?? null,
      }));
    }
    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \\d+ slides|导出校验通过：\\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'verifyExport status',
      60_000,
    );
    return true;
  })()`;
}

function reloadExportedPptxExpression(base64, caseName) {
  const targetKey =
    caseName === 'group-parent-drag'
      ? 'groupParent'
      : caseName === 'group-child'
        ? 'groupChild'
        : caseName === 'group-table-cell'
          ? 'groupTableCell'
          : caseName === 'group-text'
            ? 'groupText'
          : 'text';
  return `(async () => {
    const editor = window.__cocoPptDev;
    if (!editor) throw new Error('PPT editor dev API is not available');
    const binary = atob(${JSON.stringify(base64)});
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    await editor.loadPresentation(bytes);
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets?.${targetKey};
    if (target) {
      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: target.slideId,
        shapeIdOrName: target.shapeId,
      }).catch(() => undefined);
    }
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    return true;
  })()`;
}

function collectReloadedRoundtripResultExpression(caseName) {
  if (caseName === 'group-parent-drag')
    return collectReloadedGroupParentDragRoundtripResultExpression();
  if (caseName === 'group-child') return collectReloadedGroupChildRoundtripResultExpression();
  if (caseName === 'group-table-cell')
    return collectReloadedGroupTableCellRoundtripResultExpression();
  if (caseName === 'group-text') return collectReloadedGroupTextRoundtripResultExpression();
  return `(() => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const expected = targets?.expected ?? {};
    const scene = editor.session().scene();
    const textSlide = scene.slides.find((slide) => slide.id === targets.text.slideId);
    const textNode = textSlide?.nodes.find((node) => node.id === targets.text.shapeId);
    const tableSlide = targets.table ? scene.slides.find((slide) => slide.id === targets.table.slideId) : null;
    const tableNode = targets.table ? tableSlide?.nodes.find((node) => node.id === targets.table.shapeId) : null;
    const host = document.querySelector('#pptEditorHost');
    const root = host?.shadowRoot;
    const domNode = root?.querySelector('[data-shape-id="' + targets.text.shapeId + '"]');
    const textElement = domNode?.querySelector('.coco-ppt-rich-text span') ?? domNode;
    const domStyle = domNode ? getComputedStyle(domNode) : null;
    const textStyle = textElement ? getComputedStyle(textElement) : null;
    const failures = [];
    const normalizeColor = (value) => String(value ?? '').trim().toLowerCase();
    const closeEnough = (actual, wanted) => Number.isFinite(actual) && Number.isFinite(wanted) && Math.abs(actual - wanted) <= 1;

    if (!textNode) {
      failures.push('reloaded text shape missing');
    } else {
      if (textNode.text !== expected.text) failures.push('reloaded text mismatch');
      if (!closeEnough(textNode.textStyle?.fontSize ?? NaN, expected.fontSize)) failures.push('reloaded font size mismatch');
      if (normalizeColor(textNode.textStyle?.color) !== normalizeColor(expected.textColor)) failures.push('reloaded text color mismatch');
      if (normalizeColor(textNode.fillColor) !== normalizeColor(expected.fillColor)) failures.push('reloaded fill color mismatch');
      if (normalizeColor(textNode.lineColor) !== normalizeColor(expected.lineColor)) failures.push('reloaded line color mismatch');
      if (!closeEnough(textNode.lineWidth ?? NaN, expected.lineWidth)) failures.push('reloaded line width mismatch');
      for (const key of ['left', 'top', 'width', 'height']) {
        if (!closeEnough(textNode.bounds?.[key] ?? NaN, expected.bounds?.[key] ?? NaN)) {
          failures.push('reloaded bounds.' + key + ' mismatch');
        }
      }
    }

    if (targets.table && (tableNode?.table?.values[targets.table.row]?.[targets.table.column] ?? '') !== expected.tableText) {
      failures.push('reloaded table cell mismatch');
    }
    if (expected.text && !(root?.textContent ?? '').includes(expected.text)) {
      failures.push('reloaded DOM text missing');
    }

    return {
      failures,
      pass: failures.length === 0,
      slideCount: scene.slides.length,
      textAfterReload: textNode ? {
        text: textNode.text,
        bounds: textNode.bounds,
        textStyle: textNode.textStyle,
        fillColor: textNode.fillColor,
        lineColor: textNode.lineColor,
        lineWidth: textNode.lineWidth,
      } : null,
      tableAfterReload: targets.table ? {
        text: tableNode?.table?.values[targets.table.row]?.[targets.table.column] ?? '',
      } : null,
      domAfterReload: domStyle ? {
        background: domStyle.backgroundColor,
        borderColor: domStyle.borderColor,
        borderWidth: domStyle.borderWidth,
        color: textStyle?.color ?? '',
        fontSize: textStyle?.fontSize ?? '',
      } : null,
    };
  })()`;
}

function collectReloadedGroupTableCellRoundtripResultExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupTableCell;
    const expected = targets.expected ?? {};
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const actualText = node?.table?.values[target.row]?.[target.column] ?? '';
    const failures = [];
    if (!node) failures.push('reloaded grouped table missing');
    else {
      if (!node.capabilities?.editTableCells) failures.push('reloaded grouped table cell capability missing');
      if (actualText !== expected.tableText) failures.push('reloaded grouped table cell mismatch');
      if (node.groupContext?.parentId !== expected.parentId) failures.push('reloaded grouped table parent mismatch');
    }
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('reloaded parent group outline missing');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('reloaded parent group handle missing');
    return {
      failures,
      pass: failures.length === 0,
      slideCount: scene.slides.length,
      groupTableCellAfterReload: node ? {
        id: node.id,
        kind: node.kind,
        row: target.row,
        column: target.column,
        text: actualText,
        table: node.table,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordanceAfterReload: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
      } : null,
    };
  })()`;
}

function collectReloadedGroupTextRoundtripResultExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupText;
    const expected = targets.expected ?? {};
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    if (!node) failures.push('reloaded grouped text missing');
    else {
      if (!node.capabilities?.editText) failures.push('reloaded grouped text capability missing');
      if (node.text !== expected.text) failures.push('reloaded grouped text mismatch');
      if (node.groupContext?.parentId !== expected.parentId) failures.push('reloaded grouped text parent mismatch');
    }
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('reloaded parent group outline missing');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('reloaded parent group handle missing');
    return {
      failures,
      pass: failures.length === 0,
      slideCount: scene.slides.length,
      groupTextAfterReload: node ? {
        id: node.id,
        kind: node.kind,
        text: node.text,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordanceAfterReload: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
      } : null,
    };
  })()`;
}

function collectReloadedGroupParentDragRoundtripResultExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupParent;
    const expected = targets.expected ?? {};
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const group = slide?.groups.find((item) => item.id === target.shapeId);
    const child = slide?.nodes.find((item) => item.id === target.childShapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    const closeEnough = (actual, wanted) => Number.isFinite(actual) && Number.isFinite(wanted) && Math.abs(actual - wanted) <= 2;
    const compareGeometry = (label, actual, wanted, fields) => {
      for (const key of fields) {
        const actualValue = actual?.[key];
        const wantedValue = wanted?.[key];
        if (wantedValue === undefined) continue;
        if (typeof wantedValue === 'boolean') {
          if (actualValue !== wantedValue) failures.push(label + '.' + key + ' mismatch');
        } else if (!closeEnough(Number(actualValue), Number(wantedValue))) {
          failures.push(label + '.' + key + ' mismatch');
        }
      }
    };
    if (!group) {
      failures.push('reloaded parent group missing');
    } else {
      compareGeometry(
        'reloaded parent group bounds',
        group.bounds,
        expected.parentBounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
    }
    if (!child) {
      failures.push('reloaded child missing after parent drag');
    } else {
      compareGeometry(
        'reloaded child bounds after parent drag',
        child.bounds,
        expected.childBounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
      if (child.groupContext?.parentId !== target.shapeId) failures.push('reloaded child parent mismatch');
    }
    if (parentOutline?.dataset.shapeId !== target.shapeId) failures.push('reloaded parent group outline missing');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('reloaded parent group handle missing');
    if ((parentOutline?.querySelectorAll('.coco-ppt-parent-group-edge').length ?? 0) !== 4)
      failures.push('reloaded parent group edge drag handles missing');
    return {
      failures,
      pass: failures.length === 0,
      slideCount: scene.slides.length,
      groupParentAfterReload: group ? {
        id: group.id,
        bounds: group.bounds,
        capabilities: group.capabilities,
      } : null,
      childAfterParentReload: child ? {
        id: child.id,
        kind: child.kind,
        bounds: child.bounds,
        groupContext: child.groupContext,
      } : null,
      parentAffordanceAfterReload: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        selected: parentOutline.dataset.selected,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
        edgeCount: parentOutline.querySelectorAll('.coco-ppt-parent-group-edge').length,
      } : null,
    };
  })()`;
}

function collectReloadedGroupChildRoundtripResultExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupChild;
    const expected = targets.expected ?? {};
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const domNode = root?.querySelector('[data-shape-id="' + target.shapeId + '"]');
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    const closeEnough = (actual, wanted) => Number.isFinite(actual) && Number.isFinite(wanted) && Math.abs(actual - wanted) <= 1;
    const compareGeometry = (label, actual, wanted, fields) => {
      for (const key of fields) {
        const actualValue = actual?.[key];
        const wantedValue = wanted?.[key];
        if (wantedValue === undefined) continue;
        if (typeof wantedValue === 'boolean') {
          if (actualValue !== wantedValue) failures.push(label + '.' + key + ' mismatch');
        } else if (!closeEnough(Number(actualValue), Number(wantedValue))) {
          failures.push(label + '.' + key + ' mismatch');
        }
      }
    };
    if (!node) {
      failures.push('reloaded group child missing');
    } else {
      compareGeometry(
        'reloaded group child bounds',
        node.bounds,
        expected.bounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
      if (node.groupContext?.parentId !== expected.parentId) failures.push('reloaded group child parent mismatch');
    }
    if (!domNode) failures.push('reloaded group child DOM missing');
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('reloaded parent group outline missing');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('reloaded parent group handle missing');
    if ((parentOutline?.querySelectorAll('.coco-ppt-parent-group-edge').length ?? 0) !== 4)
      failures.push('reloaded parent group edge drag handles missing after group child edit');
    return {
      failures,
      pass: failures.length === 0,
      slideCount: scene.slides.length,
      groupChildAfterReload: node ? {
        id: node.id,
        kind: node.kind,
        bounds: node.bounds,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordanceAfterReload: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
        edgeCount: parentOutline.querySelectorAll('.coco-ppt-parent-group-edge').length,
      } : null,
    };
  })()`;
}

function collectRoundtripResultExpression(caseName) {
  if (caseName === 'group-parent-drag') return collectGroupParentDragRoundtripResultExpression();
  if (caseName === 'group-child') return collectGroupChildRoundtripResultExpression();
  if (caseName === 'group-table-cell') return collectGroupTableCellRoundtripResultExpression();
  if (caseName === 'group-text') return collectGroupTextRoundtripResultExpression();
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const exported = await editor.toPptx();
    const verified = await editor.session().verifyExport();
    const scene = editor.session().scene();
    const textSlide = scene.slides.find((slide) => slide.id === targets.text.slideId);
    const textNode = textSlide?.nodes.find((node) => node.id === targets.text.shapeId);
    const tableSlide = targets.table ? scene.slides.find((slide) => slide.id === targets.table.slideId) : null;
    const tableNode = targets.table ? tableSlide?.nodes.find((node) => node.id === targets.table.shapeId) : null;
    const host = document.querySelector('#pptEditorHost');
    const domNode = host?.shadowRoot?.querySelector('[data-shape-id="' + targets.text.shapeId + '"]');
    const textElement = domNode?.querySelector('.coco-ppt-rich-text span') ?? domNode;
    const domStyle = domNode ? getComputedStyle(domNode) : null;
    const textStyle = textElement ? getComputedStyle(textElement) : null;
    return {
      ...targets,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      verifiedSlides: verified.slides.length,
      exportedBytes: exported.length,
      textAfter: textNode ? {
        text: textNode.text,
        bounds: textNode.bounds,
        textStyle: textNode.textStyle,
        fillColor: textNode.fillColor,
        lineColor: textNode.lineColor,
        lineWidth: textNode.lineWidth,
      } : null,
      tableAfter: targets.table ? {
        text: tableNode?.table?.values[targets.table.row]?.[targets.table.column] ?? '',
      } : null,
      dom: domStyle ? {
        background: domStyle.backgroundColor,
        borderColor: domStyle.borderColor,
        borderWidth: domStyle.borderWidth,
        color: textStyle?.color ?? '',
        fontSize: textStyle?.fontSize ?? '',
      } : null,
    };
  })()`;
}

function collectGroupTextRoundtripResultExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupText;
    const expected = targets.expected ?? {};
    const exported = await editor.toPptx();
    const verified = await editor.session().verifyExport();
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    if (!node) failures.push('grouped text missing after edit');
    else {
      if (!node.capabilities?.editText) failures.push('grouped text capability missing after edit');
      if (node.text !== expected.text) failures.push('grouped text mismatch after edit');
      if (node.groupContext?.parentId !== expected.parentId) failures.push('grouped text parent mismatch after edit');
    }
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('parent group outline missing after grouped text edit');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('parent group handle missing after grouped text edit');
    return {
      ...targets,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      verifiedSlides: verified.slides.length,
      exportedBytes: exported.length,
      editFailures: failures,
      groupTextAfter: node ? {
        id: node.id,
        kind: node.kind,
        text: node.text,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordance: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
      } : null,
    };
  })()`;
}

function collectGroupTableCellRoundtripResultExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupTableCell;
    const expected = targets.expected ?? {};
    const exported = await editor.toPptx();
    const verified = await editor.session().verifyExport();
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const actualText = node?.table?.values[target.row]?.[target.column] ?? '';
    const failures = [];
    if (!node) failures.push('grouped table missing after edit');
    else {
      if (!node.capabilities?.editTableCells) failures.push('grouped table cell capability missing after edit');
      if (actualText !== expected.tableText) failures.push('grouped table cell mismatch after edit');
      if (node.groupContext?.parentId !== expected.parentId) failures.push('grouped table parent mismatch after edit');
    }
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('parent group outline missing after grouped table edit');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('parent group handle missing after grouped table edit');
    return {
      ...targets,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      verifiedSlides: verified.slides.length,
      exportedBytes: exported.length,
      editFailures: failures,
      groupTableCellAfter: node ? {
        id: node.id,
        kind: node.kind,
        row: target.row,
        column: target.column,
        text: actualText,
        table: node.table,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordance: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
      } : null,
    };
  })()`;
}

function collectGroupParentDragRoundtripResultExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupParent;
    const expected = targets.expected ?? {};
    const exported = await editor.toPptx();
    const verified = await editor.session().verifyExport();
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const group = slide?.groups.find((item) => item.id === target.shapeId);
    const child = slide?.nodes.find((item) => item.id === target.childShapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    const closeEnough = (actual, wanted) => Number.isFinite(actual) && Number.isFinite(wanted) && Math.abs(actual - wanted) <= 2;
    const compareGeometry = (label, actual, wanted, fields) => {
      for (const key of fields) {
        const actualValue = actual?.[key];
        const wantedValue = wanted?.[key];
        if (wantedValue === undefined) continue;
        if (typeof wantedValue === 'boolean') {
          if (actualValue !== wantedValue) failures.push(label + '.' + key + ' mismatch');
        } else if (!closeEnough(Number(actualValue), Number(wantedValue))) {
          failures.push(label + '.' + key + ' mismatch');
        }
      }
    };
    if (!group) {
      failures.push('parent group missing after drag');
    } else {
      compareGeometry(
        'parent group bounds after drag',
        group.bounds,
        expected.parentBounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
    }
    if (!child) {
      failures.push('child missing after parent drag');
    } else {
      compareGeometry(
        'child bounds after parent drag',
        child.bounds,
        expected.childBounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
      if (child.groupContext?.parentId !== target.shapeId) failures.push('child parent mismatch after drag');
    }
    if (parentOutline?.dataset.shapeId !== target.shapeId) failures.push('parent group outline missing after drag');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('parent group handle missing after drag');
    if ((parentOutline?.querySelectorAll('.coco-ppt-parent-group-edge').length ?? 0) !== 4)
      failures.push('parent group edge drag handles missing after drag');
    return {
      ...targets,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      verifiedSlides: verified.slides.length,
      exportedBytes: exported.length,
      editFailures: failures,
      groupParentAfter: group ? {
        id: group.id,
        bounds: group.bounds,
        capabilities: group.capabilities,
      } : null,
      childAfterParentDrag: child ? {
        id: child.id,
        kind: child.kind,
        bounds: child.bounds,
        groupContext: child.groupContext,
      } : null,
      parentAffordance: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        selected: parentOutline.dataset.selected,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
        edgeCount: parentOutline.querySelectorAll('.coco-ppt-parent-group-edge').length,
      } : null,
    };
  })()`;
}

function collectGroupChildRoundtripResultExpression() {
  return `(async () => {
    const editor = window.__cocoPptDev;
    const targets = window.__cocoPptRoundtripTargets;
    const target = targets.groupChild;
    const expected = targets.expected ?? {};
    const exported = await editor.toPptx();
    const verified = await editor.session().verifyExport();
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const node = slide?.nodes.find((item) => item.id === target.shapeId);
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const domNode = root?.querySelector('[data-shape-id="' + target.shapeId + '"]');
    const parentOutline = root?.querySelector('.coco-ppt-parent-group-outline');
    const failures = [];
    const closeEnough = (actual, wanted) => Number.isFinite(actual) && Number.isFinite(wanted) && Math.abs(actual - wanted) <= 1;
    const compareGeometry = (label, actual, wanted, fields) => {
      for (const key of fields) {
        const actualValue = actual?.[key];
        const wantedValue = wanted?.[key];
        if (wantedValue === undefined) continue;
        if (typeof wantedValue === 'boolean') {
          if (actualValue !== wantedValue) failures.push(label + '.' + key + ' mismatch');
        } else if (!closeEnough(Number(actualValue), Number(wantedValue))) {
          failures.push(label + '.' + key + ' mismatch');
        }
      }
    };
    if (!node) {
      failures.push('group child missing after edit');
    } else {
      compareGeometry(
        'group child bounds after edit',
        node.bounds,
        expected.bounds,
        ['left', 'top', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical'],
      );
      if (node.groupContext?.parentId !== expected.parentId) failures.push('group child parent mismatch after edit');
    }
    if (!domNode) failures.push('group child DOM missing after edit');
    if (parentOutline?.dataset.shapeId !== expected.parentId) failures.push('parent group outline missing after edit');
    if (!parentOutline?.querySelector('.coco-ppt-parent-group-handle')) failures.push('parent group handle missing after edit');
    if ((parentOutline?.querySelectorAll('.coco-ppt-parent-group-edge').length ?? 0) !== 4)
      failures.push('parent group edge drag handles missing after group child edit');
    return {
      ...targets,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      verifiedSlides: verified.slides.length,
      exportedBytes: exported.length,
      editFailures: failures,
      groupChildAfter: node ? {
        id: node.id,
        kind: node.kind,
        bounds: node.bounds,
        groupContext: node.groupContext,
        capabilities: node.capabilities,
      } : null,
      parentAffordance: parentOutline ? {
        shapeId: parentOutline.dataset.shapeId,
        hasHandle: Boolean(parentOutline.querySelector('.coco-ppt-parent-group-handle')),
        edgeCount: parentOutline.querySelectorAll('.coco-ppt-parent-group-edge').length,
      } : null,
    };
  })()`;
}

async function allowDownloads(client, downloadDir) {
  try {
    await client.send('Browser.setDownloadBehavior', {
      behavior: 'allow',
      downloadPath: downloadDir,
    });
  } catch {
    await client.send('Page.setDownloadBehavior', {
      behavior: 'allow',
      downloadPath: downloadDir,
    });
  }
}

async function waitForDownloadedPptx(downloadDir, timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const entries = await readdir(downloadDir).catch(() => []);
    const pending = entries.some(
      (entry) => entry.endsWith('.crdownload') || entry.endsWith('.tmp'),
    );
    const pptx = entries.find((entry) => entry.toLowerCase().endsWith('.pptx'));
    if (pptx && !pending) return join(downloadDir, pptx);
    await sleep(250);
  }
  throw new Error(`Timed out waiting for exported PPTX download in ${downloadDir}`);
}

async function verifyPptxWithPowerPoint(pptxPath, options = {}) {
  const tempDir = await mkdtemp(join(tmpdir(), 'coco-ppt-office-verify-'));
  const scriptPath = join(tempDir, 'verify.ps1');
  const resultPath = join(tempDir, 'result.json');
  const referenceSlides = Array.isArray(options.slides) ? options.slides : [];
  const referenceOutputDir =
    referenceSlides.length > 0
      ? join(options.outputDir ?? dirname(pptxPath), `powerpoint-reference-${options.runId ?? Date.now()}`)
      : '';
  if (referenceOutputDir) await mkdir(referenceOutputDir, { recursive: true });
  const releaseLock = await acquireOfficeAutomationLock();
  const script = `
$ErrorActionPreference = 'Stop'
$ppt = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(pptxPath, 'utf8').toString('base64')}'))
$resultPath = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(resultPath, 'utf8').toString('base64')}'))
$referenceOutputDir = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(referenceOutputDir, 'utf8').toString('base64')}'))
$referenceSlides = '${Buffer.from(JSON.stringify(referenceSlides), 'utf8').toString('base64')}'
$app = $null
$presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  try { $app.Visible = $false } catch {}
  $presentation = $app.Presentations.Open($ppt, $true, $false, $false)
  $count = [int]$presentation.Slides.Count
  if ($count -le 0) { throw "Opened presentation has no slides." }
  $slideImages = @()
  if (-not [string]::IsNullOrWhiteSpace($referenceOutputDir)) {
    New-Item -ItemType Directory -Force -Path $referenceOutputDir | Out-Null
    $slideSpecsJson = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($referenceSlides))
    $slideSpecs = ConvertFrom-Json -InputObject $slideSpecsJson
    foreach ($spec in $slideSpecs) {
      $slideNumber = ([int]$spec.slideIndex) + 1
      if ($slideNumber -lt 1 -or $slideNumber -gt $count) { continue }
      $width = [int]$spec.width
      $height = [int]$spec.height
      if ($width -le 0) { $width = 960 }
      if ($height -le 0) { $height = 540 }
      $imagePath = Join-Path $referenceOutputDir ("slide-{0:D3}.png" -f $slideNumber)
      $presentation.Slides.Item($slideNumber).Export($imagePath, "PNG", $width, $height)
      $slideImages += @{
        slideIndex = [int]$spec.slideIndex
        slideNumber = $slideNumber
        path = $imagePath
        width = $width
        height = $height
      }
    }
  }
  $json = @{
    ok = $true
    referenceApp = "powerpoint"
    referenceAppName = [string]$app.Name
    referenceAppVersion = [string]$app.Version
    referenceAppBuild = [string]$app.Build
    slideCount = $count
    slideImages = $slideImages
  } | ConvertTo-Json -Depth 4
  Write-Output $json
  $json | Set-Content -LiteralPath $resultPath -Encoding UTF8
} catch {
  $errorMessage = $_.Exception.Message
  if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = $_.ToString() }
  $json = @{
    ok = $false
    referenceApp = "powerpoint"
    error = $errorMessage
    category = [string]$_.CategoryInfo
    fullyQualifiedErrorId = [string]$_.FullyQualifiedErrorId
    hresult = [string]$_.Exception.HResult
    scriptStackTrace = [string]$_.ScriptStackTrace
  } | ConvertTo-Json -Depth 4
  Write-Output $json
  $json | Set-Content -LiteralPath $resultPath -Encoding UTF8
  exit 1
} finally {
  if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
  if ($app -ne $null) {
    try { $app.Quit() } catch {}
    [Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
`;
  try {
    await writeFile(scriptPath, script, 'utf8');
    const run = await spawnCapture(
      'powershell',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath],
      {
        cwd: packageDir,
        timeoutMs: 120_000,
      },
    );
    const content = await readFile(resultPath, 'utf8').catch(() => '');
    let parsed = null;
    try {
      parsed = parsePowerShellJson(content || run.stdout);
    } catch {
      parsed = null;
    }
    if (run.code !== 0 || !parsed?.ok) {
      const details = [
        parsed?.error,
        run.stderr.trim(),
        run.stdout.trim(),
        !content ? 'No result JSON was written by PowerPoint verification process.' : '',
        `Process exit code: ${run.code}`,
      ]
        .filter(Boolean)
        .join('\n');
      throw new Error(details || 'PowerPoint verification failed');
    }
    return parsed;
  } finally {
    await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    await releaseLock();
  }
}

async function acquireOfficeAutomationLock() {
  const lockDir = join(tmpdir(), 'coco-ppt-reference-export.lock');
  const startedAt = Date.now();
  while (Date.now() - startedAt < 180_000) {
    try {
      await mkdir(lockDir);
      await writeFile(
        join(lockDir, 'owner.txt'),
        `ppt-edit-roundtrip ${process.pid} ${new Date().toISOString()}`,
        'utf8',
      );
      return async () => {
        await rm(lockDir, { recursive: true, force: true }).catch(() => {});
      };
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error;
      const ageMs =
        Date.now() -
        (await stat(lockDir)
          .then((value) => value.mtimeMs)
          .catch(() => Date.now()));
      if (ageMs > 10 * 60_000) {
        await rm(lockDir, { recursive: true, force: true }).catch(() => {});
        continue;
      }
      await sleep(500);
    }
  }
  throw new Error(`Timed out waiting for Office automation lock: ${lockDir}`);
}

function parsePowerShellJson(stdout) {
  const text = stdout.replace(/^\uFEFF/, '').trim();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    const start = text.indexOf('{');
    const end = text.lastIndexOf('}');
    if (start >= 0 && end > start) return JSON.parse(text.slice(start, end + 1));
    throw new Error('PowerShell output did not contain JSON.');
  }
}

async function compareRoundtripScreenshots(editedPath, reloadedPath) {
  const tempDir = await mkdtemp(join(tmpdir(), 'coco-ppt-reload-visual-'));
  const scriptPath = join(tempDir, 'compare.py');
  const resultPath = join(tempDir, 'result.json');
  const script = `
import json
import sys
from pathlib import Path
from PIL import Image, ImageChops, ImageStat

edited_path = Path(sys.argv[1])
reloaded_path = Path(sys.argv[2])
result_path = Path(sys.argv[3])
failures = []
warnings = []

def image_stats(path):
    image = Image.open(path).convert("RGB")
    stat = ImageStat.Stat(image)
    extrema = image.getextrema()
    channel_ranges = [high - low for low, high in extrema]
    return image, {
        "width": image.width,
        "height": image.height,
        "mean": [round(value, 3) for value in stat.mean],
        "channelRanges": channel_ranges,
        "lowInformation": max(channel_ranges) < 8,
    }

edited, edited_stats = image_stats(edited_path)
reloaded, reloaded_stats = image_stats(reloaded_path)
if edited.size != reloaded.size:
    warnings.append(f"size mismatch {edited.size}/{reloaded.size}; resized for comparison")
    reloaded = reloaded.resize(edited.size)

for label, stats in (("edited", edited_stats), ("reloaded", reloaded_stats)):
    if stats["width"] <= 0 or stats["height"] <= 0:
        failures.append(f"{label} screenshot has invalid dimensions")
    if stats["lowInformation"]:
        failures.append(f"{label} screenshot has very low image information")

diff = ImageChops.difference(edited, reloaded)
diff_stat = ImageStat.Stat(diff)
mean_diff = sum(diff_stat.mean) / 3
luma = diff.convert("L")
hist = luma.histogram()
pixel_count = max(1, edited.width * edited.height)
severe_ratio = sum(hist[40:]) / pixel_count
visible_ratio = sum(hist[8:]) / pixel_count

if mean_diff > 70:
    failures.append(f"reload visual mean difference is extremely high: {mean_diff:.2f}")
elif mean_diff > 32:
    warnings.append(f"reload visual mean difference is high: {mean_diff:.2f}")
if severe_ratio > 0.42:
    failures.append(f"reload severe pixel ratio is extremely high: {severe_ratio:.3f}")
elif severe_ratio > 0.18:
    warnings.append(f"reload severe pixel ratio is high: {severe_ratio:.3f}")

result = {
    "status": "FAIL" if failures else ("REVIEW" if warnings else "PASS"),
    "pass": not failures,
    "edited": edited_stats,
    "reloaded": reloaded_stats,
    "meanDiff": round(mean_diff, 4),
    "visibleRatio": round(visible_ratio, 6),
    "severeRatio": round(severe_ratio, 6),
    "warnings": warnings,
    "failures": failures,
}
result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(result, ensure_ascii=False))
`;
  try {
    await writeFile(scriptPath, script, 'utf8');
    const run = await spawnCapture('python', [scriptPath, editedPath, reloadedPath, resultPath], {
      cwd: packageDir,
      timeoutMs: 60_000,
    });
    const content = await readFile(resultPath, 'utf8').catch(() => '');
    const parsed = content ? JSON.parse(content) : JSON.parse(run.stdout.trim());
    if (run.code !== 0) {
      parsed.failures = [
        ...(parsed.failures ?? []),
        run.stderr.trim() || run.stdout.trim() || `visual compare exited ${run.code}`,
      ];
      parsed.status = 'FAIL';
      parsed.pass = false;
    }
    return parsed;
  } catch (error) {
    return {
      status: 'REVIEW',
      pass: true,
      skipped: true,
      warnings: [
        `visual reload comparison skipped: ${error instanceof Error ? error.message : String(error)}`,
      ],
      failures: [],
    };
  } finally {
    await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function createCdpPage(chromePort) {
  return new Promise((resolvePage, rejectPage) => {
    const request = http.request(
      `http://127.0.0.1:${chromePort}/json/new?about:blank`,
      { method: 'PUT', timeout: 5_000 },
      (response) => {
        let body = '';
        response.setEncoding('utf8');
        response.on('data', (chunk) => {
          body += chunk;
        });
        response.on('end', () => {
          if (response.statusCode !== 200) {
            rejectPage(new Error(`Failed to create Chrome tab: ${response.statusCode}`));
            return;
          }
          try {
            resolvePage(JSON.parse(body));
          } catch (error) {
            rejectPage(error);
          }
        });
      },
    );
    request.on('timeout', () => request.destroy(new Error('Chrome tab creation timed out')));
    request.on('error', rejectPage);
    request.end();
  });
}

async function evaluate(client, expression) {
  const response = await client.send('Runtime.evaluate', {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (response.exceptionDetails) {
    const detail =
      response.exceptionDetails.exception?.description ?? response.exceptionDetails.text;
    throw new Error(detail ?? 'Runtime.evaluate failed');
  }
  return response.result.value;
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

async function spawnCapture(command, args, options = {}) {
  return new Promise((resolveSpawn) => {
    const child = spawn(command, args, { cwd: options.cwd, stdio: ['ignore', 'pipe', 'pipe'] });
    let stdout = '';
    let stderr = '';
    const timer = options.timeoutMs
      ? setTimeout(() => {
          child.kill();
          stderr += `\nTimed out after ${options.timeoutMs}ms`;
        }, options.timeoutMs)
      : null;
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('close', (code) => {
      if (timer) clearTimeout(timer);
      resolveSpawn({ code, stdout, stderr });
    });
  });
}

class CdpClient {
  static connect(url) {
    return new Promise((resolveClient, rejectClient) => {
      const socket = new WebSocket(url);
      const client = new CdpClient(socket);
      socket.addEventListener('open', () => resolveClient(client), { once: true });
      socket.addEventListener('error', rejectClient, { once: true });
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
    return new Promise((resolveSend, rejectSend) => {
      this.pending.set(id, { resolve: resolveSend, reject: rejectSend });
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
  let chrome = null;

  try {
    if (args.startServer) {
      viteProcess = await startVite(args.port);
    } else {
      await waitForHttp(`${baseUrl}/__coco_fixture__/pptx-files`, 1_500);
    }

    chrome = await startChrome();
    const fixtureIndexes = await resolveFixtureIndexes(baseUrl, args);
    const batchResults = [];
    for (const fixtureIndex of fixtureIndexes) {
      try {
        const result = await runRoundtrip(chrome.port, baseUrl, { ...args, fixtureIndex });
        batchResults.push({ fixtureIndex, ok: true, result });
        console.log(JSON.stringify(result, null, 2));
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const skipped = isSkippedFixtureError(message);
        batchResults.push({ fixtureIndex, ok: skipped, skipped, error: message });
        console.error(
          JSON.stringify({ fixtureIndex, ok: skipped, skipped, error: message }, null, 2),
        );
      }
    }
    const batchRunId = new Date().toISOString().replace(/[:.]/g, '-');
    const batchSummaryPath = join(args.outputDir, `batch-summary-${batchRunId}.json`);
    await mkdir(args.outputDir, { recursive: true });
    await writeFile(
      batchSummaryPath,
      JSON.stringify({ fixtures: fixtureIndexes, results: batchResults }, null, 2),
    );
    console.log(
      JSON.stringify(
        {
          batchSummaryPath,
          total: batchResults.length,
          skipped: batchResults.filter((item) => item.skipped).length,
          failed: batchResults.filter((item) => !item.ok).length,
        },
        null,
        2,
      ),
    );
    const failures = batchResults.filter(
      (item) =>
        (!item.ok && !item.skipped) ||
        item.result?.reloadedAfterExport?.pass === false ||
        item.result?.visualReloadComparison?.pass === false ||
        item.result?.referenceVisualComparisons?.some((comparison) => comparison.pass === false) ||
        item.result?.officeVerification?.ok === false ||
        (item.result?.editFailures?.length ?? 0) > 0,
    );
    if (failures.length > 0) {
      throw new Error(
        `Export reload check failed for fixtures: ${failures.map((item) => item.fixtureIndex).join(', ')}`,
      );
    }
  } finally {
    if (chrome) {
      chrome.process.kill();
      await rm(chrome.userDataDir, { recursive: true, force: true }).catch(() => {});
    }
    if (viteProcess) viteProcess.kill();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
