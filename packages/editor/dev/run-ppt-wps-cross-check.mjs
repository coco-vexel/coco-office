import { spawn } from 'node:child_process';
import http from 'node:http';
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function parseArgs(argv) {
  const parsed = {
    allFixtures: false,
    baseUrl: undefined,
    contactSheet: true,
    contactSheetPerFixture: 2,
    fixtures: undefined,
    gradientAngleMode: 'current',
    height: 720,
    hideShapes: [],
    outputDir: resolve(packageDir, 'dev/native-regression-output/ppt-reference-cross-check'),
    port: 5177,
    python: process.env.PYTHON ?? 'python',
    reference: 'powerpoint',
    startServer: false,
    width: 1280,
  };
  for (const arg of argv) {
    if (arg === '--') continue;
    if (arg === '--all-fixtures') parsed.allFixtures = true;
    else if (arg === '--no-contact-sheet') parsed.contactSheet = false;
    else if (arg === '--start-server') parsed.startServer = true;
    else if (arg.startsWith('--base-url='))
      parsed.baseUrl = arg.slice('--base-url='.length).replace(/\/$/, '');
    else if (arg.startsWith('--fixtures=')) parsed.fixtures = arg.slice('--fixtures='.length);
    else if (arg.startsWith('--gradient-angle-mode=')) {
      parsed.gradientAngleMode = arg.slice('--gradient-angle-mode='.length);
    } else if (arg.startsWith('--height=')) parsed.height = Number(arg.slice('--height='.length));
    else if (arg.startsWith('--contact-sheet-per-fixture=')) {
      parsed.contactSheetPerFixture = Number(arg.slice('--contact-sheet-per-fixture='.length));
    } else if (arg.startsWith('--hide-shapes=')) {
      parsed.hideShapes = arg
        .slice('--hide-shapes='.length)
        .split(',')
        .map((value) => value.trim())
        .filter(Boolean);
    } else if (arg.startsWith('--output-dir='))
      parsed.outputDir = resolve(packageDir, arg.slice('--output-dir='.length));
    else if (arg.startsWith('--port=')) parsed.port = Number(arg.slice('--port='.length));
    else if (arg.startsWith('--python=')) parsed.python = arg.slice('--python='.length);
    else if (arg.startsWith('--reference=')) {
      const reference = arg.slice('--reference='.length);
      parsed.reference = reference === 'office' ? 'powerpoint' : reference;
    }
    else if (arg.startsWith('--width=')) parsed.width = Number(arg.slice('--width='.length));
    else throw new Error(`Unknown argument: ${arg}`);
  }
  if (!Number.isInteger(parsed.port) || parsed.port <= 0) throw new Error('Invalid --port value.');
  if (!Number.isInteger(parsed.width) || parsed.width <= 0)
    throw new Error('Invalid --width value.');
  if (!Number.isInteger(parsed.height) || parsed.height <= 0)
    throw new Error('Invalid --height value.');
  if (!Number.isInteger(parsed.contactSheetPerFixture) || parsed.contactSheetPerFixture <= 0) {
    throw new Error('Invalid --contact-sheet-per-fixture value.');
  }
  if (!['wps', 'powerpoint', 'auto'].includes(parsed.reference)) {
    throw new Error('Invalid --reference value. Expected wps, powerpoint/office, or auto.');
  }
  if (!['current', 'plus90', 'minus90', 'office-css'].includes(parsed.gradientAngleMode)) {
    throw new Error(
      'Invalid --gradient-angle-mode value. Expected current, plus90, minus90, or office-css.',
    );
  }
  return parsed;
}

function logStep(message) {
  process.stderr.write(`[ppt-reference-cross] ${message}\n`);
}

function delay(ms) {
  return new Promise((resolveDelay) => setTimeout(resolveDelay, ms));
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
  try {
    await waitForHttp(`http://127.0.0.1:${port}/__coco_fixture__/pptx-files`, 60_000);
  } catch (error) {
    child.kill();
    throw error;
  }
  return child;
}

async function startChrome() {
  const port = await findFreePort(9222);
  const userDataDir = join(
    tmpdir(),
    `coco-ppt-reference-cross-${Date.now()}-${Math.random().toString(16).slice(2)}`,
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

async function fetchFixtures(baseUrl) {
  const response = await fetch(`${baseUrl}/__coco_fixture__/pptx-files`);
  if (!response.ok) throw new Error(`Failed to load PPT fixtures: ${response.status}`);
  return response.json();
}

function selectFixtures(fixtures, options) {
  if (options.allFixtures) return fixtures;
  if (!options.fixtures) return fixtures.slice(0, 1);
  const wanted = expandFixtureSpec(options.fixtures);
  return fixtures.filter(
    (fixture) =>
      wanted.indices.has(fixture.index) ||
      wanted.names.some((name) => fixture.name.toLowerCase().includes(name)),
  );
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

async function getFixturePath(baseUrl, fixtureIndex) {
  const response = await fetch(`${baseUrl}/__coco_fixture__/pptx?index=${fixtureIndex}`);
  if (!response.ok)
    throw new Error(`Failed to load PPT fixture ${fixtureIndex}: ${response.status}`);
  await response.arrayBuffer();
  return decodeURIComponent(response.headers.get('X-Coco-Fixture-Path') ?? '');
}

async function exportWithWps(fixturePath, outputDir, options) {
  return exportWithPresentationApp('KWPP.Application', 'wps', fixturePath, outputDir, options);
}

async function exportWithPowerPoint(fixturePath, outputDir, options) {
  return exportWithPresentationApp(
    'PowerPoint.Application',
    'powerpoint',
    fixturePath,
    outputDir,
    options,
  );
}

async function exportWithReferenceApp(fixturePath, outputDir, options) {
  if (options.reference === 'wps') return exportWithWps(fixturePath, outputDir, options);
  if (options.reference === 'powerpoint')
    return exportWithPowerPoint(fixturePath, outputDir, options);
  try {
    return await exportWithWps(fixturePath, outputDir, options);
  } catch (error) {
    logStep(`WPS export failed; falling back to PowerPoint: ${formatError(error).split('\n')[0]}`);
    const fallback = await exportWithPowerPoint(fixturePath, outputDir, options);
    return {
      ...fallback,
      fallbackFrom: 'wps',
      fallbackReason: formatError(error),
    };
  }
}

async function exportWithPresentationApp(progId, referenceApp, fixturePath, outputDir, options) {
  await mkdir(outputDir, { recursive: true });
  const tempDir = await mkdtemp(join(tmpdir(), 'coco-ppt-wps-export-'));
  const scriptPath = join(tempDir, 'export.ps1');
  const resultPath = join(tempDir, 'result.json');
  const script = `
$ErrorActionPreference = 'Stop'
$ppt = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(fixturePath, 'utf8').toString('base64')}'))
$out = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(outputDir, 'utf8').toString('base64')}'))
$progId = '${progId}'
$referenceApp = '${referenceApp}'
$width = ${options.width}
$height = ${options.height}
$resultPath = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${Buffer.from(resultPath, 'utf8').toString('base64')}'))
$app = $null
$presentation = $null
try {
  $app = New-Object -ComObject $progId
  try { $app.Visible = $false } catch {}
  $appName = ''
  $appVersion = ''
  $appBuild = ''
  try { $appName = [string]$app.Name } catch {}
  try { $appVersion = [string]$app.Version } catch {}
  try { $appBuild = [string]$app.Build } catch {}
  $presentation = $app.Presentations.Open($ppt, $true, $false, $false)
  if ([string]::IsNullOrWhiteSpace($appName)) { try { $appName = [string]$app.Name } catch {} }
  if ([string]::IsNullOrWhiteSpace($appVersion)) { try { $appVersion = [string]$app.Version } catch {} }
  if ([string]::IsNullOrWhiteSpace($appBuild)) { try { $appBuild = [string]$app.Build } catch {} }
  $count = [int]$presentation.Slides.Count
  if ($count -le 0) { throw "Opened presentation has no slides." }
  $files = @()
  for ($i = 1; $i -le $count; $i++) {
    $file = Join-Path $out ("slide-{0:D2}.png" -f $i)
    $presentation.Slides.Item($i).Export($file, 'PNG', $width, $height)
    $files += $file
  }
  $json = @{ ok = $true; referenceApp = $referenceApp; appName = $appName; appVersion = $appVersion; appBuild = $appBuild; slideCount = $count; files = $files } | ConvertTo-Json -Depth 4
  Write-Output $json
  $json | Set-Content -LiteralPath $resultPath -Encoding UTF8
} catch {
  $errorMessage = $_.Exception.Message
  if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = $_.ToString() }
  $json = @{
    ok = $false
    referenceApp = $referenceApp
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
  await writeFile(scriptPath, script, 'utf8');
  const releaseLock = await acquirePresentationExportLock(referenceApp);
  let run;
  try {
    run = await spawnCapture(
      'powershell',
      ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', scriptPath],
      {
        cwd: packageDir,
        timeoutMs: 180_000,
      },
    );
  } finally {
    await releaseLock();
  }
  const content = await readFile(resultPath, 'utf8').catch(() => '');
  await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  let parsed;
  try {
    parsed = content ? JSON.parse(stripJsonBom(content)) : parsePowerShellJson(run.stdout);
  } catch {
    parsed = null;
  }
  if (!parsed?.ok) {
    const files = (await readdir(outputDir).catch(() => []))
      .filter((entry) => /^slide-\d+\.png$/i.test(entry))
      .sort()
      .map((entry) => join(outputDir, entry));
    if (files.length > 0) {
      return {
        ok: true,
        referenceApp,
        appName: parsed?.appName,
        appVersion: parsed?.appVersion,
        appBuild: parsed?.appBuild,
        slideCount: parsed?.slideCount ?? files.length,
        files,
        fallbackFromFiles: true,
        fallbackReason:
          parsed?.error ||
          run.stderr.trim() ||
          run.stdout.trim() ||
          `No ok result JSON was written by ${referenceApp} export process.`,
      };
    }
    const details = [
      parsed?.error,
      run.stderr.trim(),
      run.stdout.trim(),
      !content ? `No result JSON was written by ${referenceApp} export process.` : '',
      `Process exit code: ${run.code}`,
    ]
      .filter(Boolean)
      .join('\n');
    throw new Error(details || `${referenceApp} export failed`);
  }
  return parsed;
}

async function acquirePresentationExportLock(referenceApp) {
  const lockDir = join(tmpdir(), 'coco-ppt-reference-export.lock');
  const start = Date.now();
  while (true) {
    try {
      await mkdir(lockDir);
      await writeFile(
        join(lockDir, 'owner.txt'),
        JSON.stringify(
          {
            pid: process.pid,
            referenceApp,
            acquiredAt: new Date().toISOString(),
          },
          null,
          2,
        ),
        'utf8',
      );
      return async () => {
        await rm(lockDir, { recursive: true, force: true }).catch(() => {});
      };
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error;
      await clearStalePresentationExportLock(lockDir);
      if (Date.now() - start > 180_000) {
        throw new Error(`Timed out waiting for ${referenceApp} export lock.`);
      }
      await delay(750);
    }
  }
}

async function clearStalePresentationExportLock(lockDir) {
  const info = await stat(lockDir).catch(() => null);
  if (!info || Date.now() - info.mtimeMs < 10 * 60_000) return;
  await rm(lockDir, { recursive: true, force: true }).catch(() => {});
}

function parsePowerShellJson(stdout) {
  const text = stripJsonBom(stdout).trim();
  if (!text) return null;
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end <= start) return null;
  return JSON.parse(text.slice(start, end + 1));
}

function stripJsonBom(value) {
  return value.replace(/^\uFEFF/, '');
}

async function captureBrowserSlides(
  chromePort,
  baseUrl,
  fixture,
  outputDir,
  expectedSlideCount,
  options,
) {
  await mkdir(outputDir, { recursive: true });
  const pageInfo = await createCdpPage(chromePort);
  const client = await CdpClient.connect(pageInfo.webSocketDebuggerUrl);
  try {
    await client.send('Page.enable');
    await client.send('Runtime.enable');
    await client.send('Emulation.setDeviceMetricsOverride', {
      width: Math.max(1440, options.width + 640),
      height: Math.max(900, options.height + 260),
      deviceScaleFactor: 1,
      mobile: false,
    });
    const url = `${baseUrl}/?surface=ppt&pptFixture=${fixture.index}&bust=cross-${Date.now()}`;
    await client.send('Page.navigate', { url });
    await waitForPptReady(client);
    await waitForBrowserVisualSettle(client);
    const pageState = await evaluate(client, pptStateExpression());
    const limit = Number.isFinite(expectedSlideCount) ? expectedSlideCount : pageState.slideCount;
    const slideCount = Math.min(limit, pageState.slideCount);
    const files = [];
    const diagnostics = [];
    for (let slideIndex = 0; slideIndex < slideCount; slideIndex++) {
      const selected = await evaluate(
        client,
        selectSlideExpression(slideIndex, {
          gradientAngleMode: options.gradientAngleMode,
          hideShapeIds: options.hideShapes,
          targetWidth: options.width,
          targetHeight: options.height,
        }),
      );
      await waitForBrowserVisualSettle(client);
      const clip = {
        x: selected.canvasRect.x,
        y: selected.canvasRect.y,
        width: selected.canvasRect.width,
        height: selected.canvasRect.height,
        scale: 1,
      };
      const screenshot = await client.send('Page.captureScreenshot', {
        format: 'png',
        fromSurface: true,
        clip,
      });
      const file = join(outputDir, `slide-${String(slideIndex + 1).padStart(2, '0')}.png`);
      await writeFile(file, Buffer.from(screenshot.data, 'base64'));
      files.push(file);
      diagnostics.push(selected);
    }
    return { files, diagnostics, pageState, slideCount };
  } finally {
    await client.send('Page.close').catch(() => undefined);
    client.close();
  }
}

function pptStateExpression() {
  return `(() => {
    const editor = window.__cocoPptDev;
    const scene = editor?.session().scene();
    return {
      slideCount: scene?.slides.length ?? 0,
      nodeCount: scene?.slides.reduce((sum, slide) => sum + slide.nodes.length, 0) ?? 0,
      name: document.querySelector('#pptSummary')?.textContent?.trim() ?? '',
    };
  })()`;
}

function selectSlideExpression(slideIndex, options = {}) {
  const gradientAngleMode = JSON.stringify(options.gradientAngleMode ?? 'current');
  const hideShapeIds = options.hideShapeIds ?? [];
  const hiddenShapeIdsJson = JSON.stringify(hideShapeIds);
  const targetWidth = Number.isFinite(options.targetWidth) ? Number(options.targetWidth) : 0;
  const targetHeight = Number.isFinite(options.targetHeight) ? Number(options.targetHeight) : 0;
  return `(async () => {
    const editor = window.__cocoPptDev;
    const slide = editor.session().scene().slides[${slideIndex}];
    if (!slide) throw new Error('PPT slide not found: ${slideIndex}');
    if (${targetWidth} > 0 && ${targetHeight} > 0 && typeof editor.computeScale === 'function') {
      editor.computeScale = (candidateSlide) =>
        Math.min(${targetWidth} / candidateSlide.size.width, ${targetHeight} / candidateSlide.size.height);
    }
    await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
    if (typeof editor.render === 'function') editor.render();
    const root = document.querySelector('#pptEditorHost')?.shadowRoot;
    const canvas = root?.querySelector('.coco-ppt-canvas');
    const start = Date.now();
    while (Date.now() - start < 5000) {
      const selected = editor.selection();
      const rendered = canvas?.querySelectorAll('.coco-ppt-node').length ?? 0;
      if (selected?.slideId === slide.id && rendered === slide.nodes.length) break;
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
    const hiddenShapeIds = new Set(${hiddenShapeIdsJson});
    const gradientAngleMode = ${gradientAngleMode};
    let captureStyle = root.querySelector('#coco-ppt-cross-check-capture-style');
    if (!captureStyle) {
      captureStyle = document.createElement('style');
      captureStyle.id = 'coco-ppt-cross-check-capture-style';
      captureStyle.textContent = [
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-node,',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-text-layer,',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-table td { outline: 0 !important; }',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-resize-handle,',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-snap-guides,',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-parent-group-outline,',
        '.coco-ppt-canvas[data-cross-check-capture="true"] .coco-ppt-parent-group-handle,',
        '.coco-ppt-context-menu { display: none !important; }',
      ].join('\\n');
      root.append(captureStyle);
    }
    canvas.dataset.crossCheckCapture = 'true';
    delete canvas.dataset.interactionMode;
    for (const node of Array.from(canvas.querySelectorAll('.coco-ppt-node'))) {
      delete node.dataset.selected;
      delete node.dataset.dragging;
      delete node.dataset.editing;
      delete node.dataset.snapReference;
      if (hiddenShapeIds.has(node.dataset.shapeId)) node.style.visibility = 'hidden';
    }
    if (gradientAngleMode !== 'current') {
      applyGradientAngleMode(canvas, slide, gradientAngleMode);
    }
    const rect = canvas.getBoundingClientRect();
    const text = (canvas.textContent ?? '').replace(/\\s+/g, ' ').trim();
    const sceneText = slide.nodes
      .map((node) => node.text ?? '')
      .join(' ')
      .replace(/\\s+/g, ' ')
      .trim();
    const internalNames = Array.from(canvas.querySelectorAll('.coco-ppt-node')).filter((node) =>
      /^图片\\s*\\d+$|^矩形\\s*\\d+$|^直线\\s*\\d+$|^椭圆\\s*\\d+$/u.test((node.textContent ?? '').trim())
    ).map((node) => (node.textContent ?? '').trim()).slice(0, 12);
    const hatchNodes = Array.from(canvas.querySelectorAll('.coco-ppt-node')).filter((node) => {
      const style = getComputedStyle(node);
      return style.backgroundImage.includes('repeating-linear-gradient');
    }).length;
    const unsupportedImages = Array.from(canvas.querySelectorAll('.coco-ppt-node[data-unsupported-image]'))
      .map((node) => {
        const bounds = node.getBoundingClientRect();
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        return {
          shapeId: node.dataset.shapeId ?? '',
          name: sceneNode?.name ?? '',
          kind: node.dataset.kind ?? '',
          imagePath: node.dataset.imagePath ?? sceneNode?.imagePath ?? '',
          contentType: node.dataset.imageContentType ?? '',
          reason: node.dataset.unsupportedImage,
          rect: {
            x: Math.round(bounds.x - rect.x),
            y: Math.round(bounds.y - rect.y),
            width: Math.round(bounds.width),
            height: Math.round(bounds.height),
          },
        };
      })
      .slice(0, 25);
    const gradientNodes = Array.from(canvas.querySelectorAll('.coco-ppt-node'))
      .map((node) => {
        const style = getComputedStyle(node);
        const bounds = node.getBoundingClientRect();
        const backgroundImage = style.backgroundImage || '';
        const svgGradientCount = node.querySelectorAll('linearGradient').length;
        const hasGradient =
          backgroundImage.includes('linear-gradient') ||
          backgroundImage.includes('radial-gradient') ||
          svgGradientCount > 0;
        if (!hasGradient) return null;
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        return {
          shapeId: node.dataset.shapeId ?? '',
          name: sceneNode?.name ?? '',
          kind: node.dataset.kind ?? '',
          presetGeometry: sceneNode?.presetGeometry ?? '',
          rect: {
            x: Math.round(bounds.x - rect.x),
            y: Math.round(bounds.y - rect.y),
            width: Math.round(bounds.width),
            height: Math.round(bounds.height),
          },
          opacity: style.opacity,
          backgroundImage: backgroundImage.slice(0, 240),
          svgGradientCount,
          fillGradient: sceneNode?.fillGradient
            ? {
                colors: sceneNode.fillGradient.colors,
                positions: sceneNode.fillGradient.positions,
                opacities: sceneNode.fillGradient.opacities,
                angle: sceneNode.fillGradient.angle,
                flip: sceneNode.fillGradient.flip,
                rotWithShape: sceneNode.fillGradient.rotWithShape,
                scaled: sceneNode.fillGradient.scaled,
                tileRect: sceneNode.fillGradient.tileRect,
              }
            : undefined,
          lineGradient: sceneNode?.lineGradient
            ? {
                colors: sceneNode.lineGradient.colors,
                positions: sceneNode.lineGradient.positions,
                opacities: sceneNode.lineGradient.opacities,
                angle: sceneNode.lineGradient.angle,
                flip: sceneNode.lineGradient.flip,
                rotWithShape: sceneNode.lineGradient.rotWithShape,
                scaled: sceneNode.lineGradient.scaled,
                tileRect: sceneNode.lineGradient.tileRect,
              }
            : undefined,
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.rect.width * b.rect.height - a.rect.width * a.rect.height)
      .slice(0, 25);
    const shapeNodes = Array.from(canvas.querySelectorAll('.coco-ppt-node'))
      .map((node) => {
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        if (!sceneNode) return null;
        const bounds = node.getBoundingClientRect();
        const relativeRect = {
          x: Math.round(bounds.x - rect.x),
          y: Math.round(bounds.y - rect.y),
          width: Math.round(bounds.width),
          height: Math.round(bounds.height),
        };
        const area = relativeRect.width * relativeRect.height;
        const isRightTop = relativeRect.x > rect.width * 0.65 && relativeRect.y < rect.height * 0.35;
        const svg = node.querySelector('svg');
        const pathCount = svg?.querySelectorAll('path').length ?? 0;
        const shapeCount = svg?.querySelectorAll('path, rect, ellipse, circle, polygon, polyline, line').length ?? 0;
        if (!isRightTop && area < 10000 && pathCount === 0) return null;
        return {
          shapeId: node.dataset.shapeId ?? '',
          name: sceneNode.name ?? '',
          kind: sceneNode.kind,
          presetGeometry: sceneNode.presetGeometry ?? '',
          rect: relativeRect,
          rotation: sceneNode.bounds.rotation,
          flipHorizontal: sceneNode.bounds.flipHorizontal,
          flipVertical: sceneNode.bounds.flipVertical,
          fillColor: sceneNode.fillColor,
          fillGradient: sceneNode.fillGradient,
          lineColor: sceneNode.lineColor,
          lineOpacity: sceneNode.lineOpacity,
          lineWidth: sceneNode.lineWidth,
          lineDash: sceneNode.lineDash,
          opacity: getComputedStyle(node).opacity,
          transform: getComputedStyle(node).transform,
          transformOrigin: getComputedStyle(node).transformOrigin,
          presetAdjustments: sceneNode.presetAdjustments,
          customCommandKinds: sceneNode.customGeometry?.paths
            ?.flatMap((path) => path.commands.map((command) => command.kind))
            .slice(0, 24),
          svgPathCount: pathCount,
          svgShapeCount: shapeCount,
          svgViewBox: svg?.getAttribute('viewBox') ?? '',
        };
      })
      .filter(Boolean)
      .sort((a, b) => {
        const rightTopScore = Number(b.rect.x > rect.width * 0.65 && b.rect.y < rect.height * 0.35) - Number(a.rect.x > rect.width * 0.65 && a.rect.y < rect.height * 0.35);
        if (rightTopScore) return rightTopScore;
        return b.rect.width * b.rect.height - a.rect.width * a.rect.height;
      })
      .slice(0, 40);
    const allTextNodes = Array.from(canvas.querySelectorAll('.coco-ppt-node'))
      .map((node) => {
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        const textValue = (sceneNode?.text ?? '').replace(/\\s+/g, ' ').trim();
        if (!textValue) return null;
        const style = getComputedStyle(node);
        const bounds = node.getBoundingClientRect();
        const textElement =
          node.querySelector('.coco-ppt-text-layer, .coco-ppt-rich-text') ?? node;
        const childBounds = textElement?.getBoundingClientRect();
        return {
          shapeId: node.dataset.shapeId ?? '',
          name: sceneNode?.name ?? '',
          kind: node.dataset.kind ?? '',
          text: textValue.slice(0, 120),
          rect: {
            x: Math.round(bounds.x - rect.x),
            y: Math.round(bounds.y - rect.y),
            width: Math.round(bounds.width),
            height: Math.round(bounds.height),
          },
          childRect: childBounds
            ? {
                x: Math.round(childBounds.x - rect.x),
                y: Math.round(childBounds.y - rect.y),
                width: Math.round(childBounds.width),
                height: Math.round(childBounds.height),
              }
            : undefined,
          firstTextRect: firstTextFragmentRect(textElement, rect),
          clientWidth: node.clientWidth,
          scrollWidth: node.scrollWidth,
          clientHeight: node.clientHeight,
          scrollHeight: node.scrollHeight,
          overflow: style.overflow,
          whiteSpace: style.whiteSpace,
          fontSize: style.fontSize,
          fontFamily: style.fontFamily,
          lineHeight: style.lineHeight,
          letterSpacing: style.letterSpacing,
          transform: style.transform,
          textBody: sceneNode?.textBody
            ? {
                wrap: sceneNode.textBody.wrap,
                autoFit: sceneNode.textBody.autoFit,
                vertical: sceneNode.textBody.vertical,
                insets: sceneNode.textBody.insets,
              }
            : undefined,
          fillColor: sceneNode?.fillColor,
          fillGradient: sceneNode?.fillGradient,
          lineColor: sceneNode?.lineColor,
          lineOpacity: sceneNode?.lineOpacity,
          lineWidth: sceneNode?.lineWidth,
          lineDash: sceneNode?.lineDash,
          presetGeometry: sceneNode?.presetGeometry,
          customGeometry: Boolean(sceneNode?.customGeometry),
          textStyle: sceneNode?.textStyle
            ? {
                fontSize: sceneNode.textStyle.fontSize,
                fontFamily: sceneNode.textStyle.fontFamily,
                lineHeight: sceneNode.textStyle.lineHeight,
                characterSpacing: sceneNode.textStyle.characterSpacing,
                alignment: sceneNode.textStyle.alignment,
                bold: sceneNode.textStyle.bold,
              }
            : undefined,
          textElement: textElement?.className ? String(textElement.className) : textElement?.tagName,
        };
      })
      .filter(Boolean);
    const textLayoutRisks = allTextNodes
      .filter((node) => {
        const overflowHidden = node.overflow === 'hidden' || node.overflow === 'clip';
        const widthOverflow = Math.max(0, node.scrollWidth - node.clientWidth);
        const heightOverflow = Math.max(0, node.scrollHeight - node.clientHeight);
        const longText = node.text.length >= 6;
        const widthRisk = widthOverflow > Math.max(6, node.clientWidth * 0.05);
        const heightRisk = heightOverflow > Math.max(6, node.clientHeight * 0.15);
        const severeShortTextRisk =
          !longText &&
          (widthOverflow > Math.max(14, node.clientWidth * 0.25) ||
            heightOverflow > Math.max(14, node.clientHeight * 0.25));
        return overflowHidden && ((longText && (widthRisk || heightRisk)) || severeShortTextRisk);
      })
      .sort((a, b) => {
        const aOverflow = Math.max(a.scrollWidth - a.clientWidth, a.scrollHeight - a.clientHeight);
        const bOverflow = Math.max(b.scrollWidth - b.clientWidth, b.scrollHeight - b.clientHeight);
        return bOverflow - aOverflow;
      })
      .slice(0, 25);
    const textNodes = allTextNodes
      .sort((a, b) => b.rect.width * b.rect.height - a.rect.width * a.rect.height)
      .slice(0, 25);
    const tableNodes = Array.from(canvas.querySelectorAll('.coco-ppt-node'))
      .map((node) => {
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        const table = sceneNode?.table;
        const tableElement = node.querySelector('.coco-ppt-table');
        if (!table || !tableElement) return null;
        const bounds = node.getBoundingClientRect();
        const tableBounds = tableElement.getBoundingClientRect();
        const sampledCells = Array.from(tableElement.querySelectorAll('td'))
          .slice(0, 16)
          .map((cell) => {
            const cellBounds = cell.getBoundingClientRect();
            const cellStyle = getComputedStyle(cell);
            const row = Number(cell.getAttribute('data-row') ?? -1);
            const column = Number(cell.getAttribute('data-column') ?? -1);
            const cellModelStyle = table.cellStyles?.[row]?.[column];
            return {
              row,
              column,
              text: (cell.textContent ?? '').replace(/\\s+/g, ' ').trim().slice(0, 80),
              rect: {
                x: Math.round(cellBounds.x - rect.x),
                y: Math.round(cellBounds.y - rect.y),
                width: Math.round(cellBounds.width),
                height: Math.round(cellBounds.height),
              },
              textAlign: cellStyle.textAlign,
              verticalAlign: cellStyle.verticalAlign,
              padding: cellStyle.padding,
              fontSize: cellStyle.fontSize,
              fontFamily: cellStyle.fontFamily,
              borderLeft: cellStyle.borderLeft,
              borderTop: cellStyle.borderTop,
              modelTextStyle: cellModelStyle?.textStyle,
              modelBody: cellModelStyle?.body,
              modelBorders: cellModelStyle?.borders,
            };
          });
        return {
          shapeId: node.dataset.shapeId ?? '',
          name: sceneNode.name ?? '',
          kind: sceneNode.kind,
          rect: {
            x: Math.round(bounds.x - rect.x),
            y: Math.round(bounds.y - rect.y),
            width: Math.round(bounds.width),
            height: Math.round(bounds.height),
          },
          tableRect: {
            x: Math.round(tableBounds.x - rect.x),
            y: Math.round(tableBounds.y - rect.y),
            width: Math.round(tableBounds.width),
            height: Math.round(tableBounds.height),
          },
          rows: table.rowCount,
          columns: table.columnCount,
          columnWidths: table.columnWidths,
          rowHeights: table.rowHeights,
          sampledCells,
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.rect.width * b.rect.height - a.rect.width * a.rect.height)
      .slice(0, 12);
    const referencedFonts = Array.from(new Set(slide.nodes.flatMap((node) => {
      const families = [];
      if (node.textStyle?.fontFamily) families.push(node.textStyle.fontFamily);
      for (const paragraph of node.textParagraphs ?? []) {
        if (paragraph.style?.bulletFontFamily) families.push(paragraph.style.bulletFontFamily);
        for (const run of paragraph.runs ?? []) {
          if (run.style?.fontFamily) families.push(run.style.fontFamily);
        }
      }
      for (const row of node.table?.cellStyles ?? []) {
        for (const cell of row) {
          if (cell?.textStyle?.fontFamily) families.push(cell.textStyle.fontFamily);
        }
      }
      return families;
    }))).filter(Boolean).sort();
    const fontChecks = referencedFonts.map((family) => ({
      family,
      loaded: Boolean(document.fonts?.check?.(\`16px "\${String(family).replace(/\\\\/g, '\\\\\\\\').replace(/"/g, '\\\\"')}"\`)),
      stylesheet: Boolean(document.getElementById(\`coco-ppt-google-font-\${String(family).replace(/[^a-zA-Z0-9_-]/g, '_')}\`)),
    }));
    return {
      slideIndex: ${slideIndex},
      sceneNodes: slide.nodes.length,
      domNodes: canvas?.querySelectorAll('.coco-ppt-node').length ?? 0,
      sceneTextLength: sceneText.length,
      textLength: text.length,
      internalNames,
      hatchNodes,
      unsupportedImages,
      gradientNodes,
      shapeNodes,
      textNodes,
      textLayoutRisks,
      tableNodes,
      fontChecks,
      canvasRect: {
        x: Math.round(rect.x),
        y: Math.round(rect.y),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      },
    };

    function firstTextFragmentRect(root, canvasRect) {
      if (!root) return undefined;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let textNode = walker.nextNode();
      while (textNode && !(textNode.textContent ?? '').trim()) textNode = walker.nextNode();
      if (!textNode) return undefined;
      const range = document.createRange();
      range.selectNodeContents(textNode);
      const bounds = range.getBoundingClientRect();
      range.detach();
      if (!bounds.width && !bounds.height) return undefined;
      return {
        x: Math.round(bounds.x - canvasRect.x),
        y: Math.round(bounds.y - canvasRect.y),
        width: Math.round(bounds.width),
        height: Math.round(bounds.height),
      };
    }

    function applyGradientAngleMode(canvas, slide, mode) {
      for (const node of Array.from(canvas.querySelectorAll('.coco-ppt-node'))) {
        const sceneNode = slide.nodes.find((item) => item.id === node.dataset.shapeId);
        if (!sceneNode?.fillGradient?.colors?.length || sceneNode.fillGradient.colors.length < 2) continue;
        const background = getComputedStyle(node).backgroundImage || '';
        if (!background.includes('linear-gradient')) continue;
        node.style.background = cssGradient(sceneNode.fillGradient, mode);
      }
    }

    function cssGradient(gradient, mode) {
      return \`linear-gradient(\${cssAngle(gradient.angle, mode)}deg, \${gradientStops(gradient).join(', ')})\`;
    }

    function cssAngle(angle, mode) {
      const value = angle === undefined ? 160 : angle;
      if (mode === 'plus90') return normalizeAngle(value + 90);
      if (mode === 'minus90') return normalizeAngle(value - 90);
      if (mode === 'office-css') return normalizeAngle(450 - value);
      return value === 0 ? 90 : value;
    }

    function normalizeAngle(value) {
      return ((value % 360) + 360) % 360;
    }

    function gradientStops(gradient) {
      return gradient.colors
        .map((color, index) => {
          const parsed = parseRgba(color);
          return {
            color: parsed?.color ?? color,
            opacity: Math.max(0, Math.min(1, gradient.opacities?.[index] ?? parsed?.opacity ?? 1)),
            position: gradient.positions?.[index],
            index,
          };
        })
        .sort((a, b) => {
          if (a.position === undefined && b.position === undefined) return a.index - b.index;
          if (a.position === undefined) return 1;
          if (b.position === undefined) return -1;
          return a.position - b.position || a.index - b.index;
        })
        .map((stop) => {
          const color = stop.opacity < 1 ? withAlpha(stop.color, stop.opacity) : stop.color;
          return stop.position === undefined ? color : \`\${color} \${Math.max(0, Math.min(100, stop.position))}%\`;
        });
    }

    function parseRgba(color) {
      const match = /^rgba?\\(\\s*(\\d{1,3})\\s*,\\s*(\\d{1,3})\\s*,\\s*(\\d{1,3})(?:\\s*,\\s*([0-9.]+)\\s*)?\\)$/i.exec(color);
      if (!match) return undefined;
      const r = Math.max(0, Math.min(255, Number(match[1])));
      const g = Math.max(0, Math.min(255, Number(match[2])));
      const b = Math.max(0, Math.min(255, Number(match[3])));
      return {
        color: \`#\${hex(r)}\${hex(g)}\${hex(b)}\`,
        opacity: match[4] === undefined ? 1 : Math.max(0, Math.min(1, Number(match[4]))),
      };
    }

    function withAlpha(color, opacity) {
      const match = /^#([0-9a-fA-F]{6})$/.exec(color);
      if (!match) return color;
      const hexValue = match[1];
      return \`rgba(\${parseInt(hexValue.slice(0, 2), 16)}, \${parseInt(hexValue.slice(2, 4), 16)}, \${parseInt(hexValue.slice(4, 6), 16)}, \${opacity})\`;
    }

    function hex(value) {
      return Math.round(value).toString(16).padStart(2, '0').toUpperCase();
    }
  })()`;
}

async function diffSlides(python, referenceDir, webDir, diffDir, slideCount) {
  await mkdir(diffDir, { recursive: true });
  const tempDir = await mkdtemp(join(tmpdir(), 'coco-ppt-diff-'));
  const scriptPath = join(tempDir, 'diff.py');
  const resultPath = join(tempDir, 'result.json');
  await writeFile(scriptPath, diffPythonScript(), 'utf8');
  const args = [scriptPath, referenceDir, webDir, diffDir, resultPath, String(slideCount)];
  const run = await spawnCapture(python, args, { cwd: packageDir, timeoutMs: 180_000 });
  const content = await readFile(resultPath, 'utf8').catch(() => '');
  await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  if (run.code !== 0)
    throw new Error(run.stderr.trim() || run.stdout.trim() || 'diff script failed');
  return JSON.parse(content);
}

function diffPythonScript() {
  return `
import json
import sys
from pathlib import Path
from PIL import Image, ImageChops, ImageStat

def diff_hotspots(diff_luma, threshold=40):
    width, height = diff_luma.size
    pixels = diff_luma.load()
    min_x, min_y = width, height
    max_x, max_y = -1, -1
    count = 0
    bands = {
        "top": 0,
        "middle": 0,
        "bottom": 0,
        "left": 0,
        "center": 0,
        "right": 0,
    }
    for y in range(height):
        row_band = "top" if y < height / 3 else "bottom" if y >= height * 2 / 3 else "middle"
        for x in range(width):
            if pixels[x, y] < threshold:
                continue
            col_band = "left" if x < width / 3 else "right" if x >= width * 2 / 3 else "center"
            count += 1
            bands[row_band] += 1
            bands[col_band] += 1
            if x < min_x:
                min_x = x
            if y < min_y:
                min_y = y
            if x > max_x:
                max_x = x
            if y > max_y:
                max_y = y
    if count == 0:
        return {
            "pixelCount": 0,
            "bbox": None,
            "bands": {key: 0 for key in bands},
        }
    return {
        "pixelCount": count,
        "bbox": {
            "x": min_x,
            "y": min_y,
            "width": max_x - min_x + 1,
            "height": max_y - min_y + 1,
        },
        "bands": {key: round(value / count, 6) for key, value in bands.items()},
        "components": diff_components(diff_luma, threshold),
    }

def diff_components(diff_luma, threshold=40, max_components=12, min_pixels=32):
    width, height = diff_luma.size
    pixels = diff_luma.load()
    visited = bytearray(width * height)
    components = []
    for y in range(height):
        for x in range(width):
            start_index = y * width + x
            if visited[start_index] or pixels[x, y] < threshold:
                continue
            visited[start_index] = 1
            stack = [(x, y)]
            count = 0
            min_x = max_x = x
            min_y = max_y = y
            while stack:
                px, py = stack.pop()
                count += 1
                if px < min_x:
                    min_x = px
                if px > max_x:
                    max_x = px
                if py < min_y:
                    min_y = py
                if py > max_y:
                    max_y = py
                for nx, ny in ((px - 1, py), (px + 1, py), (px, py - 1), (px, py + 1)):
                    if nx < 0 or ny < 0 or nx >= width or ny >= height:
                        continue
                    index = ny * width + nx
                    if visited[index] or pixels[nx, ny] < threshold:
                        continue
                    visited[index] = 1
                    stack.append((nx, ny))
            if count >= min_pixels:
                components.append({
                    "pixelCount": count,
                    "bbox": {
                        "x": min_x,
                        "y": min_y,
                        "width": max_x - min_x + 1,
                        "height": max_y - min_y + 1,
                    },
                })
    components.sort(key=lambda item: item["pixelCount"], reverse=True)
    return components[:max_components]

reference_dir = Path(sys.argv[1])
web_dir = Path(sys.argv[2])
diff_dir = Path(sys.argv[3])
result_path = Path(sys.argv[4])
slide_count = int(sys.argv[5])
rows = []
for idx in range(1, slide_count + 1):
    name = f"slide-{idx:02d}.png"
    reference_path = reference_dir / name
    web_path = web_dir / name
    diff_path = diff_dir / name
    if not reference_path.exists() or not web_path.exists():
        rows.append({"slide": idx, "status": "MISSING", "reference": str(reference_path), "wps": str(reference_path), "web": str(web_path)})
        continue
    reference_image = Image.open(reference_path).convert("RGB")
    web_image = Image.open(web_path).convert("RGB")
    reference_size = reference_image.size
    web_size = web_image.size
    compare_size = web_size
    reference = reference_image.resize(compare_size)
    web = web_image
    diff = ImageChops.difference(reference, web)
    stat = ImageStat.Stat(diff)
    mean = sum(stat.mean) / 3.0
    rms = sum((v ** 2 for v in stat.rms)) ** 0.5 / (3 ** 0.5)
    diff_luma = diff.convert("L")
    hist = diff_luma.histogram()
    pixels = reference.width * reference.height
    severe = sum(hist[40:]) / pixels
    hotspots = diff_hotspots(diff_luma)
    amplified = diff.point(lambda p: min(255, p * 4))
    amplified.save(diff_path)
    status = "REVIEW" if mean > 12 or severe > 0.055 else "OK"
    rows.append({
        "slide": idx,
        "status": status,
        "mean": round(mean, 6),
        "rms": round(rms, 6),
        "severeRatio": round(severe, 6),
        "hotspots": hotspots,
        "reference": str(reference_path),
        "wps": str(reference_path),
        "web": str(web_path),
        "diff": str(diff_path),
        "referenceSize": list(reference_size),
        "wpsSize": list(reference_size),
        "webSize": list(web_size),
        "compareSize": list(compare_size),
    })
review = [row for row in rows if row.get("status") == "REVIEW"]
summary = {
    "slideCount": slide_count,
    "reviewSlides": [row["slide"] for row in review],
    "top": sorted(rows, key=lambda row: row.get("mean", -1), reverse=True)[:8],
    "rows": rows,
}
result_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
`;
}

async function runFixture(chromePort, baseUrl, fixture, runDir, options) {
  const fixtureDir = join(
    runDir,
    `${String(fixture.index).padStart(2, '0')}-${sanitizePathSegment(fixture.name)}`,
  );
  const referenceDir = join(fixtureDir, 'reference');
  const webDir = join(fixtureDir, 'web');
  const diffDir = join(fixtureDir, 'diff');
  const fixturePath = await getFixturePath(baseUrl, fixture.index);
  logStep(`fixture ${fixture.index}: ${fixture.name}`);
  let reference;
  try {
    reference = await exportWithReferenceApp(fixturePath, referenceDir, options);
  } catch (error) {
    const browser = await captureBrowserSlides(
      chromePort,
      baseUrl,
      fixture,
      webDir,
      Number.POSITIVE_INFINITY,
      options,
    );
    const diagnosticsPath = await writeFixtureDiagnostics(fixtureDir, {
      fixtureIndex: fixture.index,
      fixtureName: fixture.name,
      fixturePath,
      diagnostics: browser.diagnostics,
      reviewSlides: [],
      top: [],
      rows: [],
    });
    return {
      fixtureIndex: fixture.index,
      fixtureName: fixture.name,
      fixturePath,
      slideCount: browser.slideCount,
      referenceUnavailable: true,
      referenceError: formatError(error) || 'Reference export failed',
      browserSlideCount: browser.slideCount,
      nodeCount: browser.pageState.nodeCount,
      reviewSlides: [],
      top: [],
      diagnostics: browser.diagnostics,
      diagnosticsPath,
      referenceDir,
      outputDir: fixtureDir,
    };
  }
  const browser = await captureBrowserSlides(
    chromePort,
    baseUrl,
    fixture,
    webDir,
    reference.slideCount,
    options,
  );
  const comparedSlideCount = Math.min(reference.slideCount, browser.slideCount);
  const diff = await diffSlides(options.python, referenceDir, webDir, diffDir, comparedSlideCount);
  const diagnosticsPath = await writeFixtureDiagnostics(fixtureDir, {
    fixtureIndex: fixture.index,
    fixtureName: fixture.name,
    fixturePath,
    referenceApp: reference.referenceApp ?? 'wps',
    referenceAppName: reference.appName,
    referenceAppVersion: reference.appVersion,
    referenceAppBuild: reference.appBuild,
    diagnostics: browser.diagnostics,
    reviewSlides: diff.reviewSlides,
    top: diff.top,
    rows: diff.rows,
  });
  return {
    fixtureIndex: fixture.index,
    fixtureName: fixture.name,
    fixturePath,
    slideCount: comparedSlideCount,
    referenceApp: reference.referenceApp ?? 'wps',
    referenceAppName: reference.appName,
    referenceAppVersion: reference.appVersion,
    referenceAppBuild: reference.appBuild,
    referenceFallbackFromFiles: reference.fallbackFromFiles,
    referenceFallbackFrom: reference.fallbackFrom,
    referenceFallbackReason: reference.fallbackReason,
    referenceSlideCount: reference.slideCount,
    wpsSlideCount: reference.slideCount,
    browserSlideCount: browser.slideCount,
    nodeCount: browser.pageState.nodeCount,
    reviewSlides: diff.reviewSlides,
    top: diff.top,
    diagnostics: browser.diagnostics,
    diagnosticsPath,
    referenceDir,
    outputDir: fixtureDir,
  };
}

async function writeFixtureDiagnostics(fixtureDir, details) {
  const reviewSlides = new Set(details.reviewSlides ?? []);
  const rowsBySlide = new Map((details.rows ?? []).map((row) => [row.slide, row]));
  const diagnosticsBySlide = Object.fromEntries(
    (details.diagnostics ?? [])
      .map((diagnostic, index) => [
        String(index + 1),
        enrichDiagnosticWithDiffHotspotCandidates(diagnostic, rowsBySlide.get(index + 1)),
      ])
      .filter(([slide]) => reviewSlides.size === 0 || reviewSlides.has(Number(slide))),
  );
  const diffRowsBySlide = Object.fromEntries(
    (details.rows ?? [])
      .filter((row) => reviewSlides.size === 0 || reviewSlides.has(row.slide))
      .map((row) => [String(row.slide), row]),
  );
  const outputPath = join(fixtureDir, 'review-diagnostics.json');
  await writeFile(
    outputPath,
    JSON.stringify(
      {
        fixtureIndex: details.fixtureIndex,
        fixtureName: details.fixtureName,
        fixturePath: details.fixturePath,
        referenceApp: details.referenceApp,
        referenceAppName: details.referenceAppName,
        referenceAppVersion: details.referenceAppVersion,
        referenceAppBuild: details.referenceAppBuild,
        reviewSlides: details.reviewSlides ?? [],
        top: details.top ?? [],
        diffRowsBySlide,
        diagnosticsBySlide,
      },
      null,
      2,
    ),
    'utf8',
  );
  return outputPath;
}

function enrichDiagnosticWithDiffHotspotCandidates(diagnostic, diffRow) {
  const bbox = diffRow?.hotspots?.bbox;
  if (!bbox) return diagnostic;
  const shapeNodes = Array.isArray(diagnostic?.shapeNodes) ? diagnostic.shapeNodes : [];
  const candidates = diffHotspotCandidatesForRect(shapeNodes, bbox);
  const components = Array.isArray(diffRow?.hotspots?.components)
    ? diffRow.hotspots.components.map((component) => ({
        ...component,
        candidates: diffHotspotCandidatesForRect(shapeNodes, component.bbox).slice(0, 5),
      }))
    : [];
  return {
    ...diagnostic,
    diffHotspot: bbox,
    diffHotspotCandidates: candidates,
    diffHotspotComponents: components,
  };
}

function diffHotspotCandidatesForRect(shapeNodes, bbox) {
  const slideArea = shapeNodes.reduce((max, shape) => Math.max(max, rectAreaValue(shape.rect)), 0);
  return shapeNodes
    .map((shape) => {
      const overlap = rectOverlapArea(shape.rect, bbox);
      if (overlap <= 0) return null;
      const rectArea = rectAreaValue(shape.rect);
      const hotspotArea = rectAreaValue(bbox);
      const backdrop = isBackdropShape(shape, rectArea, slideArea);
      return {
        shapeId: shape.shapeId,
        name: shape.name,
        kind: shape.kind,
        presetGeometry: shape.presetGeometry,
        backdrop,
        rect: shape.rect,
        overlap,
        overlapRatioOfShape: rectArea > 0 ? Number((overlap / rectArea).toFixed(4)) : 0,
        overlapRatioOfHotspot: hotspotArea > 0 ? Number((overlap / hotspotArea).toFixed(4)) : 0,
        fillColor: shape.fillColor,
        lineColor: shape.lineColor,
        lineWidth: shape.lineWidth,
        lineDash: shape.lineDash,
        rotation: shape.rotation,
        flipHorizontal: shape.flipHorizontal,
        flipVertical: shape.flipVertical,
        customCommandKinds: shape.customCommandKinds,
        presetAdjustments: shape.presetAdjustments,
      };
    })
    .filter(Boolean)
    .sort((a, b) => Number(a.backdrop) - Number(b.backdrop) || b.overlap - a.overlap)
    .slice(0, 12);
}

function isBackdropShape(shape, rectArea, slideArea) {
  if (!slideArea || rectArea < slideArea * 0.55) return false;
  return shape.kind === 'picture' || String(shape.shapeId ?? '').startsWith('layout:');
}

function rectAreaValue(rect) {
  if (!rect) return 0;
  return Math.max(0, Number(rect.width) || 0) * Math.max(0, Number(rect.height) || 0);
}

function rectOverlapArea(a, b) {
  if (!a || !b) return 0;
  const left = Math.max(Number(a.x) || 0, Number(b.x) || 0);
  const top = Math.max(Number(a.y) || 0, Number(b.y) || 0);
  const right = Math.min((Number(a.x) || 0) + (Number(a.width) || 0), (Number(b.x) || 0) + (Number(b.width) || 0));
  const bottom = Math.min((Number(a.y) || 0) + (Number(a.height) || 0), (Number(b.y) || 0) + (Number(b.height) || 0));
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function summarizeAll(results, runDir) {
  const reviewFixtures = results.filter((result) => (result.reviewSlides?.length ?? 0) > 0);
  const failedFixtures = results.filter((result) => result.error);
  const referenceUnavailable = results.filter((result) => result.referenceUnavailable);
  const unsupportedImages = results.flatMap((result) => collectUnsupportedImages(result));
  const structuralIssues = results.flatMap((result) => collectStructuralIssues(result));
  const textLayoutRisks = results.flatMap((result) => collectTextLayoutRisks(result));
  return {
    generatedAt: new Date().toISOString(),
    outputDir: runDir,
    totals: {
      fixtures: results.length,
      failedFixtures: failedFixtures.length,
      referenceUnavailable: referenceUnavailable.length,
      unsupportedImages: unsupportedImages.length,
      structuralIssues: structuralIssues.length,
      textLayoutRisks: textLayoutRisks.length,
      comparedSlides: results.reduce((sum, result) => sum + (result.slideCount ?? 0), 0),
      reviewFixtures: reviewFixtures.length,
      reviewSlides: results.reduce((sum, result) => sum + (result.reviewSlides?.length ?? 0), 0),
    },
    review: reviewFixtures.map((result) => ({
      fixtureIndex: result.fixtureIndex,
      fixtureName: result.fixtureName,
      referenceApp: result.referenceApp,
      referenceAppName: referenceAppDisplayName(result),
      referenceAppVersion: result.referenceAppVersion,
      referenceAppBuild: result.referenceAppBuild,
      referenceFallbackFromFiles: result.referenceFallbackFromFiles,
      reviewSlides: result.reviewSlides,
      topMean: result.top?.[0]?.mean ?? null,
      outputDir: result.outputDir,
    })),
    failures: failedFixtures.map((result) => ({
      fixtureIndex: result.fixtureIndex,
      fixtureName: result.fixtureName,
      error: result.error,
    })),
    referenceUnavailable: referenceUnavailable.map((result) => ({
      fixtureIndex: result.fixtureIndex,
      fixtureName: result.fixtureName,
      browserSlideCount: result.browserSlideCount,
      nodeCount: result.nodeCount,
      outputDir: result.outputDir,
      error: result.referenceError,
    })),
    unsupportedImages,
    structural: {
      status: structuralIssues.length > 0 ? 'REVIEW' : 'PASS',
      issues: structuralIssues,
    },
    layout: {
      status: textLayoutRisks.length > 0 ? 'REVIEW' : 'PASS',
      textRisks: textLayoutRisks,
    },
    warnings: [],
    results,
  };
}

function referenceAppDisplayName(result) {
  if (result.referenceApp === 'wps') return 'WPS Presentation';
  if (result.referenceApp === 'powerpoint') return result.referenceAppName || 'Microsoft PowerPoint';
  return result.referenceAppName || result.referenceApp || 'reference';
}

function collectStructuralIssues(result) {
  if (result.error || result.referenceUnavailable) return [];
  const issues = [];
  if (
    Number.isFinite(result.referenceSlideCount) &&
    Number.isFinite(result.browserSlideCount) &&
    result.referenceSlideCount !== result.browserSlideCount
  ) {
    issues.push({
      fixtureIndex: result.fixtureIndex,
      fixtureName: result.fixtureName,
      kind: 'slide-count-mismatch',
      referenceSlideCount: result.referenceSlideCount,
      browserSlideCount: result.browserSlideCount,
    });
  }
  const diagnostics = Array.isArray(result.diagnostics) ? result.diagnostics : [];
  for (const [index, diagnostic] of diagnostics.entries()) {
    if (!diagnostic) continue;
    const slide = index + 1;
    if (
      Number.isFinite(diagnostic.sceneNodes) &&
      Number.isFinite(diagnostic.domNodes) &&
      diagnostic.sceneNodes !== diagnostic.domNodes
    ) {
      issues.push({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        slide,
        kind: 'node-count-mismatch',
        sceneNodes: diagnostic.sceneNodes,
        domNodes: diagnostic.domNodes,
      });
    }
    if (Array.isArray(diagnostic.unsupportedImages) && diagnostic.unsupportedImages.length > 0) {
      issues.push({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        slide,
        kind: 'unsupported-images',
        count: diagnostic.unsupportedImages.length,
        images: diagnostic.unsupportedImages.slice(0, 5),
      });
    }
    const sceneTextLength = Number(diagnostic.sceneTextLength) || 0;
    const domTextLength = Number(diagnostic.textLength) || 0;
    if (sceneTextLength >= 8 && domTextLength < sceneTextLength * 0.5) {
      issues.push({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        slide,
        kind: 'text-coverage-low',
        sceneTextLength,
        domTextLength,
      });
    }
    const canvasRect = diagnostic.canvasRect;
    if (
      !canvasRect ||
      (Number(canvasRect.width) || 0) <= 0 ||
      (Number(canvasRect.height) || 0) <= 0
    ) {
      issues.push({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        slide,
        kind: 'empty-canvas',
      });
    }
  }
  return issues;
}

function collectUnsupportedImages(result) {
  const diagnostics = result.diagnostics;
  if (!diagnostics) return [];
  return Object.entries(diagnostics).flatMap(([slideIndex, diagnostic]) => {
    const unsupported = diagnostic?.unsupportedImages;
    if (!Array.isArray(unsupported)) return [];
    return unsupported.map((image) => ({
      fixtureIndex: result.fixtureIndex,
      fixtureName: result.fixtureName,
      slide: Number(slideIndex) + 1,
      ...image,
    }));
  });
}

function collectTextLayoutRisks(result) {
  if (result.error || result.referenceUnavailable) return [];
  const diagnostics = Array.isArray(result.diagnostics) ? result.diagnostics : [];
  const risks = [];
  for (const [index, diagnostic] of diagnostics.entries()) {
    const textLayoutRisks = Array.isArray(diagnostic?.textLayoutRisks)
      ? diagnostic.textLayoutRisks
      : [];
    for (const risk of textLayoutRisks) {
      risks.push({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        slide: index + 1,
        shapeId: risk.shapeId,
        name: risk.name,
        kind: 'text-overflow-risk',
        text: risk.text,
        rect: risk.rect,
        clientWidth: risk.clientWidth,
        scrollWidth: risk.scrollWidth,
        clientHeight: risk.clientHeight,
        scrollHeight: risk.scrollHeight,
        overflow: risk.overflow,
        whiteSpace: risk.whiteSpace,
        textBody: risk.textBody,
        textStyle: risk.textStyle,
      });
    }
  }
  return risks;
}

async function writeReviewContactSheet(python, summary, runDir, options) {
  const reviewEntries = summary.results.flatMap((result) => {
    const reviewSlides = new Set(result.reviewSlides ?? []);
    return (result.top ?? [])
      .filter((row) => reviewSlides.has(row.slide) && row.reference && row.web)
      .slice(0, options.contactSheetPerFixture)
      .map((row) => ({
        fixtureIndex: result.fixtureIndex,
        fixtureName: result.fixtureName,
        referenceApp: result.referenceApp,
        slide: row.slide,
        mean: row.mean,
        rms: row.rms,
        severeRatio: row.severeRatio,
        reference: row.reference,
        web: row.web,
      }));
  });
  if (reviewEntries.length === 0) return null;
  const tempDir = await mkdtemp(join(tmpdir(), 'coco-ppt-contact-sheet-'));
  const scriptPath = join(tempDir, 'contact-sheet.py');
  const inputPath = join(tempDir, 'input.json');
  const outputPath = join(runDir, 'review-contact-sheet.jpg');
  await writeFile(scriptPath, contactSheetPythonScript(), 'utf8');
  await writeFile(
    inputPath,
    JSON.stringify({ outputPath, entries: reviewEntries }, null, 2),
    'utf8',
  );
  const run = await spawnCapture(python, [scriptPath, inputPath], {
    cwd: packageDir,
    timeoutMs: 180_000,
  });
  await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  if (run.code !== 0)
    throw new Error(run.stderr.trim() || run.stdout.trim() || 'contact sheet failed');
  return outputPath;
}

async function writeLatestCrossCheckArtifacts(summaryPath, summary, outputDir) {
  await mkdir(outputDir, { recursive: true });
  const latestSummaryPath = join(outputDir, 'latest-summary.json');
  await copyFile(summaryPath, latestSummaryPath);
  let latestContactSheetPath = null;
  if (summary.contactSheetPath) {
    latestContactSheetPath = join(outputDir, 'latest-review-contact-sheet.jpg');
    await copyFile(summary.contactSheetPath, latestContactSheetPath);
  }
  return { latestSummaryPath, latestContactSheetPath };
}

function contactSheetPythonScript() {
  return `
import json
import sys
from pathlib import Path
from PIL import Image, ImageDraw

data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
entries = data["entries"]
output_path = Path(data["outputPath"])
thumb_w, thumb_h = 320, 180
label_h = 46
cols = 2
cell_w = thumb_w * 2 + 16
cell_h = thumb_h + label_h
rows = max(1, (len(entries) + cols - 1) // cols)
sheet = Image.new("RGB", (cols * cell_w, rows * cell_h), "white")
draw = ImageDraw.Draw(sheet)

for index, entry in enumerate(entries):
    row = index // cols
    col = index % cols
    x = col * cell_w
    y = row * cell_h
    reference = Image.open(entry["reference"]).convert("RGB").resize((thumb_w, thumb_h))
    web = Image.open(entry["web"]).convert("RGB").resize((thumb_w, thumb_h))
    sheet.paste(reference, (x, y + label_h))
    sheet.paste(web, (x + thumb_w + 8, y + label_h))
    label = f"F{entry['fixtureIndex']} S{entry['slide']} {entry.get('referenceApp') or 'reference'} mean={entry.get('mean', 0):.2f} rms={entry.get('rms', 0):.2f} severe={entry.get('severeRatio', 0):.3f}"
    draw.text((x + 4, y + 4), label, fill=(0, 0, 0))
    draw.text((x + 4, y + 25), "Reference", fill=(30, 80, 180))
    draw.text((x + thumb_w + 12, y + 25), "Browser", fill=(180, 60, 30))

output_path.parent.mkdir(parents=True, exist_ok=True)
sheet.save(output_path, quality=92)
`;
}

function sanitizePathSegment(value) {
  return (
    basename(String(value))
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80) || 'fixture'
  );
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

async function waitForPptReady(client) {
  const start = Date.now();
  while (Date.now() - start < 120_000) {
    const ready = await evaluate(
      client,
      `Boolean(window.__cocoPptDev && window.__cocoPptDev.session().scene().slides.length > 0 && document.querySelector('#pptEditorHost')?.shadowRoot?.querySelector('.coco-ppt-canvas'))`,
    ).catch(() => false);
    if (ready) {
      await sleep(700);
      return;
    }
    await sleep(250);
  }
  throw new Error('Timed out waiting for PPT editor dev page.');
}

async function waitForBrowserVisualSettle(client) {
  await evaluate(
    client,
    `(async () => {
      if (document.fonts?.ready) {
        await Promise.race([
          document.fonts.ready,
          new Promise((resolve) => setTimeout(resolve, 5000)),
        ]);
      }
      if (window.__cocoPptFontRenderReady) {
        await Promise.race([
          window.__cocoPptFontRenderReady,
          new Promise((resolve) => setTimeout(resolve, 5000)),
        ]);
      }
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const canvas = root?.querySelector('.coco-ppt-canvas');
      const rect = canvas?.getBoundingClientRect();
      return Boolean(rect && rect.width > 0 && rect.height > 0);
    })()`,
  ).catch(() => undefined);
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
      // Keep polling.
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

function formatError(error) {
  if (error instanceof Error) {
    return error.stack || error.message || error.name || 'Unknown Error';
  }
  const text = String(error);
  return text && text !== '[object Object]' ? text : JSON.stringify(error);
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
  const options = parseArgs(process.argv.slice(2));
  const baseUrl = options.baseUrl ?? `http://127.0.0.1:${options.port}`;
  let viteProcess = null;
  let chrome = null;
  const runId = new Date().toISOString().replace(/[:.]/g, '-');
  const runDir = join(options.outputDir, runId);
  try {
    if (options.startServer) {
      viteProcess = await startVite(options.port);
    } else {
      await waitForHttp(`${baseUrl}/__coco_fixture__/pptx-files`, 1500);
    }
    const fixtures = selectFixtures(await fetchFixtures(baseUrl), options);
    if (fixtures.length === 0) throw new Error('No PPT fixtures selected.');
    await mkdir(runDir, { recursive: true });
    chrome = await startChrome();
    const results = [];
    for (const fixture of fixtures) {
      try {
        results.push(await runFixture(chrome.port, baseUrl, fixture, runDir, options));
      } catch (error) {
        results.push({
          fixtureIndex: fixture.index,
          fixtureName: fixture.name,
          error: formatError(error) || 'Unknown failure',
        });
      }
    }
    const summary = summarizeAll(results, runDir);
    const summaryPath = join(runDir, 'summary.json');
    await writeFile(summaryPath, JSON.stringify(summary, null, 2), 'utf8');
    if (options.contactSheet) {
      try {
        summary.contactSheetPath = await writeReviewContactSheet(
          options.python,
          summary,
          runDir,
          options,
        );
      } catch (error) {
        summary.contactSheetPath = null;
        summary.contactSheetError = formatError(error) || 'contact sheet failed';
        summary.warnings.push(summary.contactSheetError);
      }
    }
    await writeFile(summaryPath, JSON.stringify(summary, null, 2), 'utf8');
    const latestArtifacts = await writeLatestCrossCheckArtifacts(
      summaryPath,
      summary,
      options.outputDir,
    );
    console.log(
      JSON.stringify(
        {
          summaryPath,
          ...latestArtifacts,
          contactSheetPath: summary.contactSheetPath,
          totals: summary.totals,
          structural: summary.structural,
          layout: summary.layout,
          unsupportedImages: summary.unsupportedImages,
          review: summary.review,
          failures: summary.failures,
          referenceUnavailable: summary.referenceUnavailable,
        },
        null,
        2,
      ),
    );
    if (summary.totals.failedFixtures > 0) process.exitCode = 1;
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
