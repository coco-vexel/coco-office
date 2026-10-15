import { spawn } from 'node:child_process';
import http from 'node:http';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function logStep(message) {
  process.stderr.write(`[ppt-smoke] ${message}\n`);
}

function parseArgs(argv) {
  const parsed = {
    allFixtures: false,
    auditCapabilityScreenshotLimit: 6,
    auditLimit: 3,
    auditScreenshotLimit: 1,
    auditTargets: '',
    baseUrl: undefined,
    boundaryAudit: true,
    fast: false,
    full: false,
    highFrequencyProfileLimit: 8,
    highFrequencySampleLimit: 3,
    officeExportSmoke: true,
    officeExportSmokeCases: '0:default,13:group-child,13:group-text,5:group-table-cell',
    outputDir: resolve(packageDir, 'dev/native-regression-output/ppt-editor-smoke'),
    port: 5178,
    referenceTargets: 'auto',
    screenshots: true,
    startServer: false,
  };
  const explicit = new Set();
  for (const arg of argv) {
    if (arg === '--') continue;
    if (arg === '--fast') parsed.fast = true;
    else if (arg === '--full') parsed.full = true;
    else if (arg === '--all-fixtures') {
      explicit.add('allFixtures');
      parsed.allFixtures = true;
    }
    else if (arg === '--no-boundary-audit') {
      explicit.add('boundaryAudit');
      parsed.boundaryAudit = false;
    }
    else if (arg === '--no-office-export-smoke') {
      explicit.add('officeExportSmoke');
      parsed.officeExportSmoke = false;
    }
    else if (arg === '--no-audit-screenshots') {
      explicit.add('screenshots');
      parsed.screenshots = false;
    }
    else if (arg === '--start-server') parsed.startServer = true;
    else if (arg.startsWith('--audit-limit=')) {
      explicit.add('auditLimit');
      parsed.auditLimit = Number(arg.slice('--audit-limit='.length));
    }
    else if (arg.startsWith('--audit-capability-screenshot-limit=')) {
      explicit.add('auditCapabilityScreenshotLimit');
      parsed.auditCapabilityScreenshotLimit = Number(arg.slice('--audit-capability-screenshot-limit='.length));
    }
    else if (arg.startsWith('--audit-screenshot-limit=')) {
      explicit.add('auditScreenshotLimit');
      parsed.auditScreenshotLimit = Number(arg.slice('--audit-screenshot-limit='.length));
    }
    else if (arg.startsWith('--audit-targets=')) {
      explicit.add('auditTargets');
      parsed.auditTargets = arg.slice('--audit-targets='.length);
    }
    else if (arg.startsWith('--high-frequency-profile-limit=')) {
      explicit.add('highFrequencyProfileLimit');
      parsed.highFrequencyProfileLimit = Number(arg.slice('--high-frequency-profile-limit='.length));
    }
    else if (arg.startsWith('--high-frequency-sample-limit=')) {
      explicit.add('highFrequencySampleLimit');
      parsed.highFrequencySampleLimit = Number(arg.slice('--high-frequency-sample-limit='.length));
    }
    else if (arg.startsWith('--office-export-smoke-cases=')) {
      explicit.add('officeExportSmokeCases');
      parsed.officeExportSmokeCases = arg.slice('--office-export-smoke-cases='.length);
    }
    else if (arg.startsWith('--reference-targets=')) {
      explicit.add('referenceTargets');
      parsed.referenceTargets = arg.slice('--reference-targets='.length);
    }
    else if (arg.startsWith('--base-url=')) parsed.baseUrl = arg.slice('--base-url='.length).replace(/\/$/, '');
    else if (arg.startsWith('--output-dir=')) parsed.outputDir = resolve(packageDir, arg.slice('--output-dir='.length));
    else if (arg.startsWith('--port=')) parsed.port = Number(arg.slice('--port='.length));
    else throw new Error(`Unknown argument: ${arg}`);
  }
  if (parsed.full) {
    if (!explicit.has('allFixtures')) parsed.allFixtures = true;
    if (!explicit.has('auditScreenshotLimit')) parsed.auditScreenshotLimit = 3;
    if (!explicit.has('auditCapabilityScreenshotLimit')) parsed.auditCapabilityScreenshotLimit = 24;
    if (!explicit.has('highFrequencyProfileLimit')) parsed.highFrequencyProfileLimit = 0;
    if (!explicit.has('highFrequencySampleLimit')) parsed.highFrequencySampleLimit = 5;
    if (!explicit.has('officeExportSmokeCases')) {
      parsed.officeExportSmokeCases =
        '0:default,13:group-child,13:group-text,13:group-parent-drag,5:group-table-cell';
    }
  }
  if (parsed.fast) {
    if (parsed.full) throw new Error('--fast and --full cannot be used together.');
    if (!explicit.has('auditLimit')) parsed.auditLimit = 1;
    if (!explicit.has('auditScreenshotLimit')) parsed.auditScreenshotLimit = 0;
    if (!explicit.has('auditCapabilityScreenshotLimit')) parsed.auditCapabilityScreenshotLimit = 0;
    if (!explicit.has('highFrequencyProfileLimit')) parsed.highFrequencyProfileLimit = 5;
    if (!explicit.has('highFrequencySampleLimit')) parsed.highFrequencySampleLimit = 1;
    if (!explicit.has('boundaryAudit')) parsed.boundaryAudit = false;
    if (!explicit.has('officeExportSmoke')) parsed.officeExportSmoke = false;
    if (!explicit.has('screenshots')) parsed.screenshots = false;
  }
  if (!Number.isInteger(parsed.port) || parsed.port <= 0) throw new Error('Invalid --port value.');
  if (!Number.isInteger(parsed.auditLimit) || parsed.auditLimit <= 0) {
    throw new Error('Invalid --audit-limit value.');
  }
  if (!Number.isInteger(parsed.auditScreenshotLimit) || parsed.auditScreenshotLimit < 0) {
    throw new Error('Invalid --audit-screenshot-limit value.');
  }
  if (!Number.isInteger(parsed.auditCapabilityScreenshotLimit) || parsed.auditCapabilityScreenshotLimit < 0) {
    throw new Error('Invalid --audit-capability-screenshot-limit value.');
  }
  parsed.auditTargets = parseAuditTargets(parsed.auditTargets);
  if (!Number.isInteger(parsed.highFrequencyProfileLimit) || parsed.highFrequencyProfileLimit < 0) {
    throw new Error('Invalid --high-frequency-profile-limit value.');
  }
  if (!Number.isInteger(parsed.highFrequencySampleLimit) || parsed.highFrequencySampleLimit <= 0) {
    throw new Error('Invalid --high-frequency-sample-limit value.');
  }
  parsed.officeExportSmokeCases = parseOfficeExportSmokeCases(parsed.officeExportSmokeCases);
  parsed.referenceTargets = parseReferenceTargets(parsed.referenceTargets);
  return parsed;
}

function parseAuditTargets(value) {
  if (Array.isArray(value)) return value;
  const text = String(value ?? '').trim();
  if (!text) return [];
  return text
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const [fixtureText, slideText = '1'] = entry.split(':');
      const fixtureIndex = Number(fixtureText);
      const slideNumber = Number(slideText);
      if (!Number.isInteger(fixtureIndex) || fixtureIndex < 0) {
        throw new Error(`Invalid --audit-targets fixture: ${entry}`);
      }
      if (!Number.isInteger(slideNumber) || slideNumber <= 0) {
        throw new Error(`Invalid --audit-targets slide: ${entry}`);
      }
      return { fixtureIndex, slideIndex: slideNumber - 1, reason: 'targeted' };
    });
}

function parseOfficeExportSmokeCases(value) {
  if (Array.isArray(value)) return value;
  const cases = String(value ?? '')
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const [fixtureText, caseName = 'default'] = entry.split(':');
      const fixtureIndex = Number(fixtureText);
      if (!Number.isInteger(fixtureIndex) || fixtureIndex < 0) {
        throw new Error(`Invalid --office-export-smoke-cases fixture: ${entry}`);
      }
      if (!caseName.trim()) throw new Error(`Invalid --office-export-smoke-cases case: ${entry}`);
      return { fixtureIndex, caseName: caseName.trim() };
    });
  if (cases.length === 0) throw new Error('Invalid --office-export-smoke-cases value.');
  return cases;
}

function parseReferenceTargets(value) {
  if (value && typeof value === 'object' && !Array.isArray(value)) return value;
  const text = String(value ?? '').trim();
  if (!text || text === 'auto') return { mode: 'auto', targets: [] };
  if (text === 'none' || text === 'off' || text === 'false') return { mode: 'none', targets: [] };
  return { mode: 'explicit', targets: parseAuditTargets(text) };
}

async function startVite(port) {
  const viteEntry = resolve(packageDir, 'node_modules/vite/bin/vite.js');
  const child = spawn(process.execPath, [viteEntry, '--host', '127.0.0.1', '--port', String(port), '--strictPort'], {
    cwd: packageDir,
    env: { ...process.env, COCO_PPT_SMOKE: '1' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
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
  const userDataDir = join(tmpdir(), `coco-ppt-editor-smoke-${Date.now()}-${Math.random().toString(16).slice(2)}`);
  await mkdir(userDataDir, { recursive: true });
  logStep(`starting chrome on ${port}`);
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
  logStep(`chrome ready on ${port}`);
  return { process: child, port, userDataDir };
}

function chromeExecutable() {
  if (process.env.CHROME_PATH) return process.env.CHROME_PATH;
  if (process.platform === 'win32') return 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
  if (process.platform === 'darwin') return '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  return 'google-chrome';
}

async function runSmoke(chromePort, baseUrl, options) {
  logStep('creating cdp page');
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

    const runId = new Date().toISOString().replace(/[:.]/g, '-');
    const mode = smokeMode(options);
    await mkdir(options.outputDir, { recursive: true });
    const previousBaseline = await loadPreviousFixtureBaseline(options.outputDir, mode);
    logStep('running high-frequency edits');
    const highFrequency = await runHighFrequencyEdit(client, baseUrl, options, runId);
    logStep('running interaction edits');
    const interactionTimeoutMs = options.full ? 600_000 : 120_000;
    const interactions = await withTimeout(
      () => runInteractionEditCases(client, baseUrl, options, highFrequency.sampledFixtures, runId),
      interactionTimeoutMs,
      'interaction edits',
    ).catch((error) => ({
      fixtureIndex: highFrequency.fixtureIndex,
      fixtureIndexes: highFrequency.sampledFixtures ?? [],
      cases: [],
      failures: [`interaction edits failed: ${formatError(error).split('\n')[0]}`],
      screenshotPaths: [],
    }));
    const boundaries = options.boundaryAudit
      ? (logStep('running object boundary audit'),
        await withTimeout(
          () => runObjectBoundary(client, baseUrl),
          90_000,
          'object boundary audit',
        ).catch((error) => ({
          sampleKinds: [],
          checks: [],
          failures: [`object boundary audit failed: ${formatError(error).split('\n')[0]}`],
          skipped: false,
        })))
      : { sampleKinds: [], checks: [], failures: [], skipped: true };
    logStep('running fixture audit');
    const fixtureAudit = await runFixtureAudit(client, baseUrl, options);
    fixtureAudit.baseline = compareFixtureAuditToBaseline(fixtureAudit, previousBaseline);
    fixtureAudit.renderingFailures.push(...fixtureAudit.baseline.renderingFailures);
    fixtureAudit.contentFailures.push(...fixtureAudit.baseline.contentFailures);
    if (options.screenshots) {
      logStep('capturing audit screenshots');
      fixtureAudit.screenshotPaths = await captureAuditScreenshots(
        client,
        baseUrl,
        fixtureAudit,
        options.outputDir,
        runId,
        options,
      );
      fixtureAudit.visualSummary = buildVisualSummary(fixtureAudit.screenshotPaths);
      fixtureAudit.renderingFailures.push(...fixtureAudit.visualSummary.failures);
    }
    const officeExport = options.officeExportSmoke
      ? await runOfficeExportSmoke(baseUrl, options, fixtureAudit)
      : { cases: [], failures: [], skipped: true };
    const screenshot = await client.send('Page.captureScreenshot', { format: 'png', fromSurface: true });

    const screenshotPath = join(options.outputDir, `ppt-editor-smoke-${runId}.png`);
    const summaryPath = join(options.outputDir, `summary-${runId}.json`);
    const reportPath = join(options.outputDir, 'latest-report.json');
    const markdownPath = join(options.outputDir, 'latest-report.md');
    const htmlPath = join(options.outputDir, 'latest-report.html');
    const modeReportPath = join(options.outputDir, `latest-${mode}-report.json`);
    const modeMarkdownPath = join(options.outputDir, `latest-${mode}-report.md`);
    const modeHtmlPath = join(options.outputDir, `latest-${mode}-report.html`);
    const result = {
      highFrequency,
      interactions,
      boundaries,
      fixtureAudit,
      officeExport,
      options: publicOptions(options),
      failures: [
        ...highFrequency.failures,
        ...interactions.failures,
        ...boundaries.failures,
        ...fixtureAudit.renderingFailures,
        ...fixtureAudit.contentFailures,
        ...officeExport.failures,
      ],
      screenshotPath,
      summaryPath,
      reportPath,
      markdownPath,
      htmlPath,
      modeReportPath,
      modeMarkdownPath,
      modeHtmlPath,
    };
    const report = buildAcceptanceReport(result);
    result.report = report;
    await writeFile(screenshotPath, Buffer.from(screenshot.data, 'base64'));
    await writeFile(summaryPath, JSON.stringify(result, null, 2));
    await writeFile(reportPath, JSON.stringify(report, null, 2));
    await writeFile(markdownPath, renderMarkdownReport(report), 'utf-8');
    await writeFile(htmlPath, renderHtmlReport(report, options.outputDir), 'utf-8');
    await writeFile(modeReportPath, JSON.stringify(report, null, 2));
    await writeFile(modeMarkdownPath, renderMarkdownReport(report), 'utf-8');
    await writeFile(modeHtmlPath, renderHtmlReport(report, options.outputDir), 'utf-8');
    logStep('smoke report written');
    return result;
  } finally {
    await client.send('Page.close').catch(() => undefined);
    client.close();
  }
}

async function runOfficeExportSmoke(baseUrl, options, fixtureAudit) {
  const cases = [];
  const failures = [];
  for (const item of options.officeExportSmokeCases) {
    logStep(`running office export smoke fixture ${item.fixtureIndex} ${item.caseName}`);
    const referenceTargets = pickOfficeReferenceTargets(item, options, fixtureAudit);
    const compareSlides = referenceTargets.map((target) => target.slideIndex + 1);
    const compareSlideArgs =
      compareSlides.length > 0 ? [`--compare-slides=${compareSlides.join(',')}`] : [];
    const run = await spawnCapture(
      process.execPath,
      [
        resolve(packageDir, 'dev/run-ppt-edit-roundtrip.mjs'),
        `--base-url=${baseUrl}`,
        `--fixture=${item.fixtureIndex}`,
        `--case=${item.caseName}`,
        ...compareSlideArgs,
      ],
      { cwd: packageDir, timeoutMs: 180_000 },
    );
    const parsedObjects = parseJsonObjects(run.stdout);
    const parsed = parsedObjects.findLast((item) => item?.officeVerification || item?.cases?.[0]?.officeVerification)
      ?? parsedObjects.at(-1)
      ?? null;
    const officeVerification = parsed?.cases?.[0]?.officeVerification ?? parsed?.officeVerification ?? null;
    const visualReloadComparison = parsed?.cases?.[0]?.result?.visualReloadComparison
      ?? parsed?.visualReloadComparison
      ?? null;
    const referenceVisualComparisons = parsed?.cases?.[0]?.result?.referenceVisualComparisons
      ?? parsed?.referenceVisualComparisons
      ?? [];
    const ok =
      run.code === 0 &&
      officeVerification?.ok === true &&
      !referenceVisualComparisons.some((comparison) => comparison.pass === false);
    const result = {
      fixtureIndex: item.fixtureIndex,
      caseName: item.caseName,
      ok,
      code: run.code,
      officeVerification,
      visualReloadComparison,
      referenceVisualComparisons,
      referenceTargets,
      summaryPath: parsed?.summaryPath ?? null,
      stdoutTail: trimTail(run.stdout),
      stderrTail: trimTail(run.stderr),
    };
    cases.push(result);
    if (!ok) {
      const reason = officeVerification?.error
        ?? result.stderrTail
        ?? result.stdoutTail
        ?? `exit code ${run.code}`;
      failures.push(`office export smoke fixture ${item.fixtureIndex} ${item.caseName}: ${reason}`);
    }
  }
  return { cases, failures, skipped: false };
}

const DEFAULT_REFERENCE_TARGETS = [
  { fixtureIndex: 0, slideIndex: 14, reason: 'known-complex' },
  { fixtureIndex: 1, slideIndex: 18, reason: 'known-complex' },
  { fixtureIndex: 4, slideIndex: 3, reason: 'known-complex' },
  { fixtureIndex: 5, slideIndex: 22, reason: 'known-complex' },
  { fixtureIndex: 13, slideIndex: 1, reason: 'known-group-child' },
];

function pickOfficeReferenceTargets(caseItem, options, fixtureAudit) {
  const config = options.referenceTargets ?? { mode: 'auto', targets: [] };
  if (config.mode === 'none') return [];
  const selected = [];
  const seen = new Set();
  const add = (target) => {
    if (!target || target.fixtureIndex !== caseItem.fixtureIndex) return;
    const key = `${target.fixtureIndex}:${target.slideIndex}`;
    if (seen.has(key)) return;
    seen.add(key);
    selected.push(target);
  };

  for (const target of options.auditTargets ?? []) add({ ...target, reason: target.reason ?? 'targeted' });
  for (const target of config.targets ?? []) add({ ...target, reason: target.reason ?? 'reference-targeted' });
  if (config.mode !== 'auto') return selected;
  const explicitCount = selected.length;

  for (const target of DEFAULT_REFERENCE_TARGETS) add(target);

  const fixture = fixtureAudit?.fixtures?.find((item) => item.fixtureIndex === caseItem.fixtureIndex);
  if (fixture) {
    for (const slide of fixture.suspiciousSlides ?? []) add({ fixtureIndex: fixture.fixtureIndex, slideIndex: slide.index, reason: 'suspicious' });
    for (const target of pickCapabilityScreenshotTargets([fixture], 4)) {
      add({ ...target, reason: `reference-${target.reason}` });
    }
  }
  return selected.slice(0, Math.max(2, explicitCount));
}

function parseJsonObjects(text) {
  const value = String(text ?? '');
  const objects = [];
  let start = -1;
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = 0; index < value.length; index++) {
    const char = value[index];
    if (inString) {
      if (escaped) escaped = false;
      else if (char === '\\') escaped = true;
      else if (char === '"') inString = false;
      continue;
    }
    if (char === '"') {
      inString = true;
      continue;
    }
    if (char === '{') {
      if (depth === 0) start = index;
      depth += 1;
      continue;
    }
    if (char !== '}') continue;
    depth -= 1;
    if (depth === 0 && start >= 0) {
      const candidate = value.slice(start, index + 1);
      try {
        objects.push(JSON.parse(candidate));
      } catch {}
      start = -1;
    }
  }
  return objects;
}

function trimTail(text, limit = 1200) {
  const value = String(text ?? '').trim();
  return value.length > limit ? value.slice(value.length - limit) : value;
}

function formatNumber(value) {
  return Number.isFinite(value) ? Number(value).toFixed(4) : 'n/a';
}

function classifyReferenceComparison(comparison) {
  const severe = Number(comparison?.severeRatio ?? 0);
  const mean = Number(comparison?.meanDiff ?? 0);
  if (
    comparison?.pass === false ||
    comparison?.failures?.length ||
    severe >= 0.18 ||
    mean >= 32
  ) {
    return { label: 'HIGH', className: 'risk-high', reason: 'possible content loss or severe distortion' };
  }
  if (comparison?.warnings?.length || severe >= 0.08 || mean >= 12) {
    return { label: 'REVIEW', className: 'risk-review', reason: 'needs visual review' };
  }
  return { label: 'OK', className: 'risk-ok', reason: 'within expected renderer differences' };
}

async function withTimeout(run, timeoutMs, label) {
  let timer;
  try {
    return await Promise.race([
      run(),
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error(`Timed out after ${timeoutMs}ms: ${label}`)), timeoutMs);
      }),
    ]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function formatError(error) {
  return error instanceof Error ? (error.stack ?? error.message) : String(error);
}

function buildAcceptanceReport(result) {
  const renderingFailures = [...result.fixtureAudit.renderingFailures];
  if (!result.screenshotPath) renderingFailures.push('missing browser screenshot');
  if (!result.boundaries.skipped) {
    if (!result.boundaries.sampleKinds.includes('text')) renderingFailures.push('missing rendered text sample');
    if (!result.boundaries.sampleKinds.includes('picture')) renderingFailures.push('missing rendered picture sample');
    if (!result.boundaries.sampleKinds.includes('projected')) renderingFailures.push('missing projected object sample');
  }

  const contentFailures = [...result.fixtureAudit.contentFailures];
  contentFailures.push(...result.officeExport.failures);
  const editCases = result.highFrequency.cases?.length ? result.highFrequency.cases : [result.highFrequency];
  for (const editCase of editCases) {
    if (!/^(Verified \d+ slides|导出校验通过：\d+ 页)$/.test(editCase.status)) {
      contentFailures.push(`fixture ${editCase.fixtureIndex}: export verification did not pass`);
    }
    if (editCase.revision <= 0) contentFailures.push(`fixture ${editCase.fixtureIndex}: scene revision did not advance`);
  }

  const editingFailures = [
    ...result.highFrequency.failures,
    ...result.interactions.failures,
    ...result.boundaries.failures,
  ];
  const categories = {
    rendering: {
      status: renderingFailures.length > 0 ? 'FAIL' : 'PASS',
      scope: 'Browser smoke render: page loads, key object classes are present, and a screenshot is captured.',
      failures: renderingFailures,
      evidence: {
        screenshotPath: result.screenshotPath,
        auditScreenshots: result.fixtureAudit.screenshotPaths ?? [],
        editScreenshots: result.highFrequency.screenshotPaths ?? [],
        visualSummary: result.fixtureAudit.visualSummary ?? null,
        baseline: result.fixtureAudit.baseline ?? null,
        sampleKinds: result.boundaries.sampleKinds,
        auditedFixtures: result.fixtureAudit.fixtures.length,
        renderedDomNodes: result.fixtureAudit.totals.domNodes,
      },
    },
    content: {
      status: contentFailures.length > 0 ? 'FAIL' : 'PASS',
      scope: 'Content preservation smoke: high-frequency edits remain in scene memory and exported PPTX reparses.',
      failures: contentFailures,
      evidence: {
        exportStatus: result.highFrequency.status,
        exportStatuses: result.highFrequency.exportStatuses ?? [result.highFrequency.status],
        revision: result.highFrequency.revision,
        highFrequencyCases: result.highFrequency.cases ?? [],
        sampledFixtures: result.highFrequency.sampledFixtures ?? [result.highFrequency.fixtureIndex],
        sampledProfiles: result.highFrequency.sampledProfiles ?? [],
        summaryPath: result.summaryPath,
        auditedSlides: result.fixtureAudit.totals.slides,
        auditedNodes: result.fixtureAudit.totals.sceneNodes,
        renderedTextLength: result.fixtureAudit.totals.domTextLength,
        officeExport: result.officeExport,
      },
    },
    editing: {
      status: editingFailures.length > 0 ? 'FAIL' : 'PASS',
      scope: 'High-frequency editing: text, style, geometry, table cell, and editable-boundary controls.',
      failures: editingFailures,
      evidence: {
        interactionCases: result.interactions.cases,
        interactionFixtures: result.interactions.fixtureIndexes ?? [result.interactions.fixtureIndex],
        interactionScreenshots: result.interactions.screenshotPaths ?? [],
        textShapeId: result.highFrequency.textShapeId,
        tableShapeId: result.highFrequency.tableShapeId,
        sampledFixtures: result.highFrequency.sampledFixtures ?? [result.highFrequency.fixtureIndex],
        sampledProfiles: result.highFrequency.sampledProfiles ?? [],
        boundarySampleKinds: result.boundaries.sampleKinds,
      },
    },
  };
  const failures = Object.entries(categories).flatMap(([category, value]) =>
    value.failures.map((failure) => `${category}: ${failure}`),
  );
  return {
    generatedAt: new Date().toISOString(),
    mode: smokeMode(result.options),
    options: result.options,
    acceptance: failures.length > 0 ? 'FAIL' : 'PASS',
    categories,
    failures,
  };
}

function renderMarkdownReport(report) {
  const lines = [
    '# PPT Editor Smoke Report',
    '',
    `Generated: ${report.generatedAt}`,
    `Mode: ${report.mode}`,
    `Acceptance: ${report.acceptance}`,
    '',
    '| Category | Status | Scope | Failures |',
    '| --- | --- | --- | --- |',
  ];
  for (const [category, value] of Object.entries(report.categories)) {
    lines.push(
      `| ${category} | ${value.status} | ${value.scope} | ${value.failures.length ? value.failures.join('; ') : '-'} |`,
    );
  }
  lines.push('', '## Evidence', '');
  lines.push(`- Screenshot: ${report.categories.rendering.evidence.screenshotPath}`);
  const editScreenshots = report.categories.rendering.evidence.editScreenshots;
  if (editScreenshots.length > 0) {
    lines.push('- Edit screenshots:');
    for (const item of editScreenshots) {
      const suffix = item.visual?.warnings?.length
        ? ` warnings=${item.visual.warnings.join('; ')}`
        : '';
      lines.push(`  - fixture ${item.fixtureIndex} ${item.kind}: ${item.path}${suffix}`);
    }
  }
  const auditScreenshots = report.categories.rendering.evidence.auditScreenshots;
  if (auditScreenshots.length > 0) {
    lines.push('- Audit screenshots:');
    for (const item of auditScreenshots) {
      const suffix = item.visual?.warnings?.length
        ? ` warnings=${item.visual.warnings.join('; ')}`
        : '';
      lines.push(`  - fixture ${item.fixtureIndex} slide ${item.slideIndex + 1} (${item.reason}): ${item.path}${suffix}`);
    }
  }
  const visualSummary = report.categories.rendering.evidence.visualSummary;
  if (visualSummary) {
    lines.push(`- Visual screenshots checked: ${visualSummary.checked}`);
    lines.push(`- Visual warnings: ${visualSummary.warningCount}`);
  }
  const baseline = report.categories.rendering.evidence.baseline;
  if (baseline) {
    lines.push(`- Baseline: ${baseline.status}`);
    if (baseline.comparedFixtures > 0) {
      lines.push(`- Baseline compared fixtures: ${baseline.comparedFixtures}`);
      lines.push(`- Baseline compared slides: ${baseline.comparedSlides ?? 0}`);
      lines.push(`- Baseline warnings: ${baseline.warnings.length}`);
    }
  }
  lines.push(`- Summary: ${report.categories.content.evidence.summaryPath}`);
  lines.push(`- Audited fixtures: ${report.categories.rendering.evidence.auditedFixtures}`);
  lines.push(`- Audited slides: ${report.categories.content.evidence.auditedSlides}`);
  lines.push(`- Audited scene nodes: ${report.categories.content.evidence.auditedNodes}`);
  lines.push(`- Rendered DOM nodes: ${report.categories.rendering.evidence.renderedDomNodes}`);
  lines.push(`- Rendered text chars: ${report.categories.content.evidence.renderedTextLength}`);
  lines.push(`- High-frequency edit fixtures: ${report.categories.editing.evidence.sampledFixtures.join(', ')}`);
  const sampledProfiles = report.categories.editing.evidence.sampledProfiles;
  if (sampledProfiles.length > 0) {
    lines.push('- High-frequency sample reasons:');
    for (const item of sampledProfiles) {
      lines.push(
        `  - fixture ${item.fixtureIndex}: ${item.reason}; slides=${item.slideCount}, text=${item.sceneTextLength}, images=${item.imageCount}, editableTables=${item.editableTableCount}`,
      );
    }
  }
  const highFrequencyCases = report.categories.content.evidence.highFrequencyCases;
  if (highFrequencyCases.length > 0) {
    lines.push('- High-frequency edit cases:');
    for (const item of highFrequencyCases) {
      lines.push(
        `  - fixture ${item.fixtureIndex}: text=${item.textShapeId}, table=${item.tableShapeId ?? 'none'}, revision=${item.revision}, ${item.status}`,
      );
    }
  }
  lines.push(
    `- Boundary samples: ${report.categories.editing.evidence.boundarySampleKinds.join(', ')}`,
  );
  const officeExport = report.categories.content.evidence.officeExport;
  if (officeExport?.cases?.length > 0) {
    lines.push('- Office export smoke:');
    for (const item of officeExport.cases) {
      const status = item.ok ? 'PASS' : 'FAIL';
      const slides = item.officeVerification?.slideCount ?? 'unknown';
      const visual = item.visualReloadComparison
        ? `, visual=${item.visualReloadComparison.status}, mean=${formatNumber(item.visualReloadComparison.meanDiff)}, severe=${formatNumber(item.visualReloadComparison.severeRatio)}`
        : '';
      lines.push(`  - fixture ${item.fixtureIndex} ${item.caseName}: ${status}, slides=${slides}${visual}`);
      for (const comparison of item.referenceVisualComparisons ?? []) {
        const target = (item.referenceTargets ?? []).find(
          (target) => target.slideIndex + 1 === comparison.slideNumber,
        );
        const reason = target?.reason ? `, reason=${target.reason}` : '';
        const risk = classifyReferenceComparison(comparison);
        lines.push(
          `    - slide ${comparison.slideNumber}: reference=${comparison.status}, risk=${risk.label}, mean=${formatNumber(comparison.meanDiff)}, severe=${formatNumber(comparison.severeRatio)}${reason}`,
        );
      }
    }
  }
  const interactionCases = report.categories.editing.evidence.interactionCases ?? [];
  if (interactionCases.length > 0) {
    lines.push('- Interaction cases:');
    for (const item of interactionCases) {
      lines.push(`  - ${item.kind}: fixture ${item.fixtureIndex}, shape=${item.shapeId}, ${item.status}`);
    }
  }
  const interactionScreenshots = report.categories.editing.evidence.interactionScreenshots ?? [];
  if (interactionScreenshots.length > 0) {
    lines.push('- Interaction screenshots:');
    for (const item of interactionScreenshots) {
      lines.push(`  - fixture ${item.fixtureIndex} ${item.kind}: ${item.path}`);
    }
  }
  lines.push(`- Text shape: ${report.categories.editing.evidence.textShapeId}`);
  lines.push(`- Table shape: ${report.categories.editing.evidence.tableShapeId ?? 'none'}`);
  lines.push('');
  return `${lines.join('\n')}\n`;
}

function renderHtmlReport(report, outputDir) {
  const rendering = report.categories.rendering;
  const content = report.categories.content;
  const editing = report.categories.editing;
  const imagePath = (path) => relative(outputDir, path).replace(/\\/g, '/');
  const statusClass = (status) => status === 'PASS' ? 'pass' : status === 'SKIP' ? 'skip' : 'fail';
  const warningText = (item) => item.visual?.warnings?.length ? item.visual.warnings.join('; ') : '';
  const failureText = (item) => item.visual?.failures?.length ? item.visual.failures.join('; ') : '';
  const stateOf = (item) => item.visual?.failures?.length ? 'failure' : item.visual?.warnings?.length ? 'warning' : 'pass';
  const screenshotCard = (item, label, source) => `
    <article class="shot" data-source="${escapeHtml(source)}" data-kind="${escapeHtml(item.kind ?? '')}" data-reason="${escapeHtml(item.reason ?? '')}" data-state="${escapeHtml(stateOf(item))}">
      <a href="${escapeHtml(imagePath(item.path))}" target="_blank" rel="noreferrer">
        <img src="${escapeHtml(imagePath(item.path))}" alt="${escapeHtml(label)}" loading="lazy">
      </a>
      <div class="shot-meta">
        <strong>${escapeHtml(label)}</strong>
        <span>nodes ${item.selected?.domNodes ?? 0}/${item.selected?.sceneNodes ?? 0}</span>
        <span>text ${item.selected?.domTextLength ?? 0}/${item.selected?.sceneTextLength ?? 0}</span>
        ${warningText(item) ? `<span class="warn">${escapeHtml(warningText(item))}</span>` : ''}
        ${failureText(item) ? `<span class="fail-text">${escapeHtml(failureText(item))}</span>` : ''}
      </div>
    </article>`;
  const referenceStateOf = (comparison) =>
    comparison.failures?.length ? 'failure' : comparison.warnings?.length ? 'warning' : 'pass';
  const referenceCards = (content.evidence.officeExport?.cases ?? []).flatMap((officeCase) =>
    (officeCase.referenceVisualComparisons ?? []).map((comparison) => {
      const target = (officeCase.referenceTargets ?? []).find(
        (target) => target.slideIndex + 1 === comparison.slideNumber,
      );
      return { officeCase, comparison, reason: target?.reason ?? '' };
    }),
  );
  const referenceRiskCards = [...referenceCards].sort((a, b) => {
    const severeDiff = Number(b.comparison.severeRatio ?? 0) - Number(a.comparison.severeRatio ?? 0);
    if (Math.abs(severeDiff) > 0.000001) return severeDiff;
    return Number(b.comparison.meanDiff ?? 0) - Number(a.comparison.meanDiff ?? 0);
  });
  const referenceRiskRows = referenceRiskCards.map(({ officeCase, comparison, reason }, index) => `
    <tr class="reference-row" data-source="reference" data-kind="${escapeHtml(officeCase.caseName)}" data-reason="${escapeHtml(reason)}" data-state="${escapeHtml(referenceStateOf(comparison))}" data-risk="${escapeHtml(classifyReferenceComparison(comparison).label.toLowerCase())}">
      <td>${index + 1}</td>
      <td>fixture ${officeCase.fixtureIndex}</td>
      <td>${escapeHtml(officeCase.caseName)}</td>
      <td>slide ${comparison.slideNumber}</td>
      <td><span class="chip ${statusClass(comparison.status)}">${escapeHtml(comparison.status)}</span></td>
      <td><span class="risk ${escapeHtml(classifyReferenceComparison(comparison).className)}">${escapeHtml(classifyReferenceComparison(comparison).label)}</span></td>
      <td>${formatNumber(comparison.severeRatio)}</td>
      <td>${formatNumber(comparison.meanDiff)}</td>
      <td>${escapeHtml(reason)}</td>
      <td>
        <a href="${escapeHtml(imagePath(comparison.browserPath))}" target="_blank" rel="noreferrer">Browser</a>
        <span class="link-separator">/</span>
        <a href="${escapeHtml(imagePath(comparison.referencePath))}" target="_blank" rel="noreferrer">PowerPoint</a>
      </td>
    </tr>`).join('\n');
  const referenceComparisonCard = ({ officeCase, comparison, reason }) => `
    <article class="reference-shot" data-source="reference" data-kind="${escapeHtml(officeCase.caseName)}" data-reason="${escapeHtml(reason)}" data-state="${escapeHtml(referenceStateOf(comparison))}" data-risk="${escapeHtml(classifyReferenceComparison(comparison).label.toLowerCase())}">
      <div class="reference-pair">
        <a href="${escapeHtml(imagePath(comparison.browserPath))}" target="_blank" rel="noreferrer">
          <img src="${escapeHtml(imagePath(comparison.browserPath))}" alt="fixture ${officeCase.fixtureIndex} slide ${comparison.slideNumber} browser" loading="lazy">
          <span>Browser reload</span>
        </a>
        <a href="${escapeHtml(imagePath(comparison.referencePath))}" target="_blank" rel="noreferrer">
          <img src="${escapeHtml(imagePath(comparison.referencePath))}" alt="fixture ${officeCase.fixtureIndex} slide ${comparison.slideNumber} PowerPoint" loading="lazy">
          <span>PowerPoint</span>
        </a>
      </div>
      <div class="shot-meta">
        <strong>fixture ${officeCase.fixtureIndex} ${officeCase.caseName} slide ${comparison.slideNumber}</strong>
        <span><span class="risk ${escapeHtml(classifyReferenceComparison(comparison).className)}">${escapeHtml(classifyReferenceComparison(comparison).label)}</span> ${escapeHtml(classifyReferenceComparison(comparison).reason)}</span>
        <span>reference ${comparison.status}, mean ${formatNumber(comparison.meanDiff)}, severe ${formatNumber(comparison.severeRatio)}</span>
        ${reason ? `<span>reason ${escapeHtml(reason)}</span>` : ''}
        ${comparison.warnings?.length ? `<span class="warn">${escapeHtml(comparison.warnings.join('; '))}</span>` : ''}
        ${comparison.failures?.length ? `<span class="fail-text">${escapeHtml(comparison.failures.join('; '))}</span>` : ''}
      </div>
    </article>`;
  const editShots = rendering.evidence.editScreenshots ?? [];
  const auditShots = rendering.evidence.auditScreenshots ?? [];
  const interactionShots = editing.evidence.interactionScreenshots ?? [];
  const samples = editing.evidence.sampledProfiles ?? [];
  const cases = content.evidence.highFrequencyCases ?? [];
  const baseline = rendering.evidence.baseline;
  const visualSummary = rendering.evidence.visualSummary;

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PPT Editor Smoke Gallery</title>
  <style>
    :root { color-scheme: light; font-family: Inter, "Segoe UI", Arial, sans-serif; color: #172033; background: #f6f8fb; }
    body { margin: 0; }
    header { position: sticky; top: 0; z-index: 2; padding: 18px 24px; background: rgba(255, 255, 255, .92); border-bottom: 1px solid #dbe2ec; backdrop-filter: blur(12px); }
    h1 { margin: 0 0 10px; font-size: 24px; letter-spacing: 0; }
    h2 { margin: 28px 0 12px; font-size: 18px; letter-spacing: 0; }
    main { padding: 20px 24px 36px; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; }
    .chip { padding: 5px 9px; border: 1px solid #ccd6e3; border-radius: 6px; background: #fff; font-size: 13px; }
    .filters { position: sticky; top: 81px; z-index: 1; display: flex; flex-wrap: wrap; gap: 8px; padding: 10px 0 14px; background: rgba(246, 248, 251, .92); backdrop-filter: blur(10px); }
    .filter { appearance: none; border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; color: #26364c; cursor: pointer; font: inherit; font-size: 13px; padding: 7px 10px; }
    .filter[aria-pressed="true"] { border-color: #486fbd; background: #eaf0ff; color: #1f4aa0; }
    .filter-count { align-self: center; color: #64748b; font-size: 13px; }
    .pass { border-color: #7dbb8e; color: #17682d; background: #edf8ef; }
    .fail { border-color: #d88484; color: #a32323; background: #fff0f0; }
    .skip { border-color: #d5b15f; color: #765511; background: #fff8e7; }
    .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 14px; }
    .shot { overflow: hidden; border: 1px solid #d9e0ea; border-radius: 8px; background: #fff; box-shadow: 0 8px 24px rgba(21, 32, 52, .06); }
    .shot img { display: block; width: 100%; aspect-ratio: 16 / 10; object-fit: contain; background: #eef2f7; }
    .reference-shot { overflow: hidden; border: 1px solid #d9e0ea; border-radius: 8px; background: #fff; box-shadow: 0 8px 24px rgba(21, 32, 52, .06); }
    .reference-pair { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 1px; background: #d9e0ea; }
    .reference-pair a { position: relative; display: block; color: inherit; text-decoration: none; background: #eef2f7; }
    .reference-pair img { display: block; width: 100%; aspect-ratio: 16 / 9; object-fit: contain; }
    .reference-pair span { position: absolute; left: 8px; top: 8px; padding: 3px 6px; border-radius: 5px; background: rgba(255,255,255,.88); color: #33445c; font-size: 12px; }
    .shot-meta { display: grid; gap: 4px; padding: 10px 12px 12px; font-size: 12px; color: #526070; }
    .shot-meta strong { color: #172033; font-size: 13px; }
    .warn { color: #8a5b00; }
    .fail-text { color: #a32323; }
    .risk { display: inline-flex; align-items: center; width: max-content; padding: 3px 7px; border-radius: 5px; border: 1px solid transparent; font-size: 12px; font-weight: 700; letter-spacing: 0; }
    .risk-ok { border-color: #8fc59a; color: #17682d; background: #edf8ef; }
    .risk-review { border-color: #d5b15f; color: #765511; background: #fff8e7; }
    .risk-high { border-color: #d88484; color: #a32323; background: #fff0f0; }
    a { color: #2459b8; }
    .link-separator { color: #94a3b8; padding: 0 4px; }
    table { width: 100%; border-collapse: collapse; background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; overflow: hidden; }
    th, td { padding: 9px 10px; border-bottom: 1px solid #e7ecf3; text-align: left; font-size: 13px; }
    th { background: #eef3f8; color: #33445c; }
    tr:last-child td { border-bottom: 0; }
  </style>
</head>
<body>
  <header>
    <h1>PPT Editor Smoke Gallery</h1>
    <div class="chips">
      <span class="chip ${statusClass(report.acceptance)}">Acceptance ${escapeHtml(report.acceptance)}</span>
      <span class="chip ${statusClass(rendering.status)}">Rendering ${escapeHtml(rendering.status)}</span>
      <span class="chip ${statusClass(content.status)}">Content ${escapeHtml(content.status)}</span>
      <span class="chip ${statusClass(editing.status)}">Editing ${escapeHtml(editing.status)}</span>
      <span class="chip">Generated ${escapeHtml(report.generatedAt)}</span>
      <span class="chip">Mode ${escapeHtml(report.mode)}</span>
      <span class="chip">Slides ${content.evidence.auditedSlides}</span>
      <span class="chip">DOM nodes ${rendering.evidence.renderedDomNodes}</span>
      ${baseline ? `<span class="chip ${statusClass(baseline.status)}">Baseline ${escapeHtml(baseline.status)} ${baseline.comparedSlides ?? 0} slides</span>` : ''}
      ${visualSummary ? `<span class="chip">Visual ${visualSummary.checked} checked / ${visualSummary.warningCount} warnings</span>` : ''}
    </div>
  </header>
  <main>
    <nav class="filters" aria-label="Screenshot filters">
      <button class="filter" type="button" data-filter="all" aria-pressed="true">All</button>
      <button class="filter" type="button" data-filter="edit">Edit</button>
      <button class="filter" type="button" data-filter="audit">Audit</button>
      <button class="filter" type="button" data-filter="reference">Reference</button>
      <button class="filter" type="button" data-filter="review">Review</button>
      <button class="filter" type="button" data-filter="high">High</button>
      <button class="filter" type="button" data-filter="table">Table</button>
      <button class="filter" type="button" data-filter="image">Image</button>
      <button class="filter" type="button" data-filter="text">Text</button>
      <button class="filter" type="button" data-filter="warning">Warning</button>
      <button class="filter" type="button" data-filter="failure">Failure</button>
      <span class="filter-count" id="filterCount"></span>
    </nav>
    <h2>High-Frequency Samples</h2>
    <table>
      <thead><tr><th>Fixture</th><th>Reason</th><th>Slides</th><th>Text</th><th>Images</th><th>Tables</th><th>Export</th></tr></thead>
      <tbody>
        ${samples.map((sample) => {
          const item = cases.find((editCase) => editCase.fixtureIndex === sample.fixtureIndex);
          return `<tr><td>${sample.fixtureIndex}</td><td>${escapeHtml(sample.reason)}</td><td>${sample.slideCount}</td><td>${sample.sceneTextLength}</td><td>${sample.imageCount}</td><td>${sample.editableTableCount}</td><td>${escapeHtml(item?.status ?? '')}</td></tr>`;
        }).join('\n')}
      </tbody>
    </table>

    <h2>Reference Comparisons</h2>
    <table>
      <thead><tr><th>Rank</th><th>Fixture</th><th>Case</th><th>Slide</th><th>Status</th><th>Risk</th><th>Severe</th><th>Mean</th><th>Reason</th><th>Images</th></tr></thead>
      <tbody>
        ${referenceRiskRows}
      </tbody>
    </table>
    <section class="grid">
      ${referenceCards.map((item) => referenceComparisonCard(item)).join('\n')}
    </section>

    <h2>Edit Screenshots</h2>
    <section class="grid">
      ${editShots.map((item) => screenshotCard(item, `fixture ${item.fixtureIndex} ${item.kind}`, 'edit')).join('\n')}
    </section>

    <h2>Interaction Screenshots</h2>
    <section class="grid">
      ${interactionShots.map((item) => screenshotCard(item, `fixture ${item.fixtureIndex} ${item.kind}`, 'interaction')).join('\n')}
    </section>

    <h2>Audit Screenshots</h2>
    <section class="grid">
      ${auditShots.map((item) => screenshotCard(item, `fixture ${item.fixtureIndex} slide ${item.slideIndex + 1} ${item.reason}`, 'audit')).join('\n')}
    </section>
  </main>
  <script>
    const buttons = Array.from(document.querySelectorAll('.filter'));
    const cards = Array.from(document.querySelectorAll('.shot, .reference-shot, .reference-row'));
    const count = document.querySelector('#filterCount');
    const matches = (card, filter) => {
      if (filter === 'all') return true;
      if (filter === 'edit' || filter === 'audit' || filter === 'reference') return card.dataset.source === filter;
      if (filter === 'review' || filter === 'high') return card.dataset.risk === filter;
      if (filter === 'warning' || filter === 'failure') return card.dataset.state === filter;
      return card.dataset.kind === filter || card.dataset.reason.includes(filter);
    };
    const applyFilter = (filter) => {
      let visible = 0;
      for (const card of cards) {
        const show = matches(card, filter);
        card.hidden = !show;
        if (show) visible += 1;
      }
      for (const button of buttons) button.setAttribute('aria-pressed', String(button.dataset.filter === filter));
      count.textContent = visible + ' / ' + cards.length + ' items';
    };
    for (const button of buttons) button.addEventListener('click', () => applyFilter(button.dataset.filter));
    applyFilter('all');
  </script>
</body>
</html>
`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function publicOptions(options) {
  return {
    boundaryAudit: options.boundaryAudit,
    fast: options.fast,
    full: options.full,
    allFixtures: options.allFixtures,
    auditLimit: options.auditLimit,
    auditScreenshotLimit: options.auditScreenshotLimit,
    auditCapabilityScreenshotLimit: options.auditCapabilityScreenshotLimit,
    auditTargets: options.auditTargets,
    highFrequencyProfileLimit: options.highFrequencyProfileLimit,
    highFrequencySampleLimit: options.highFrequencySampleLimit,
    officeExportSmoke: options.officeExportSmoke,
    officeExportSmokeCases: options.officeExportSmokeCases,
    referenceTargets: options.referenceTargets,
    screenshots: options.screenshots,
  };
}

function smokeMode(options) {
  if (options.full) return 'full';
  if (options.fast) return 'fast';
  return 'quick';
}

async function loadPreviousFixtureBaseline(outputDir, mode) {
  const latestReportPath = join(outputDir, `latest-${mode}-report.json`);
  try {
    const report = JSON.parse(await readFile(latestReportPath, 'utf-8'));
    const summaryPath = report?.categories?.content?.evidence?.summaryPath;
    if (!summaryPath) return { status: 'missing', reason: 'latest report has no summary path' };
    const summary = JSON.parse(await readFile(summaryPath, 'utf-8'));
    const audit = summary?.fixtureAudit;
    if (!audit?.fixtures?.length) return { status: 'missing', reason: 'latest summary has no fixture audit' };
    return {
      status: 'loaded',
      generatedAt: summary?.report?.generatedAt ?? report?.generatedAt ?? null,
      scope: audit.scope,
      auditedCount: audit.auditedCount,
      fixtureCount: audit.fixtureCount,
      totals: audit.totals,
      fixtures: audit.fixtures,
      summaryPath,
    };
  } catch (error) {
    return { status: 'missing', reason: error.message };
  }
}

function compareFixtureAuditToBaseline(audit, baseline) {
  if (!baseline || baseline.status !== 'loaded') {
    return {
      status: 'SKIP',
      reason: baseline?.reason ?? 'no previous baseline',
      comparedFixtures: 0,
      warnings: [],
      renderingFailures: [],
      contentFailures: [],
    };
  }
  if (baseline.auditedCount !== audit.auditedCount || baseline.scope !== audit.scope) {
    return {
      status: 'SKIP',
      reason: `baseline scope mismatch (${baseline.scope}/${baseline.auditedCount} vs ${audit.scope}/${audit.auditedCount})`,
      comparedFixtures: 0,
      warnings: [],
      renderingFailures: [],
      contentFailures: [],
      baselineSummaryPath: baseline.summaryPath,
    };
  }

  const previousByIndex = new Map(baseline.fixtures.map((fixture) => [fixture.fixtureIndex, fixture]));
  const warnings = [];
  const renderingFailures = [];
  const contentFailures = [];
  let comparedFixtures = 0;
  let comparedSlides = 0;

  const ratioDrop = (current, previous, ratio) => previous > 0 && current < Math.floor(previous * ratio);
  for (const fixture of audit.fixtures) {
    const previous = previousByIndex.get(fixture.fixtureIndex);
    if (!previous) {
      warnings.push(`fixture ${fixture.fixtureIndex}: no baseline fixture`);
      continue;
    }
    comparedFixtures += 1;
    const label = `fixture ${fixture.fixtureIndex}`;
    if (previous.fixtureName && fixture.fixtureName && previous.fixtureName !== fixture.fixtureName) {
      warnings.push(`${label}: fixture name changed (${previous.fixtureName} -> ${fixture.fixtureName})`);
    }
    if (fixture.slideCount < previous.slideCount) {
      contentFailures.push(`${label}: slide count dropped (${fixture.slideCount}/${previous.slideCount})`);
    }
    if (ratioDrop(fixture.sceneNodeCount, previous.sceneNodeCount, 0.95)) {
      contentFailures.push(`${label}: scene nodes dropped (${fixture.sceneNodeCount}/${previous.sceneNodeCount})`);
    }
    if (ratioDrop(fixture.domNodeCount, previous.domNodeCount, 0.95)) {
      renderingFailures.push(`${label}: rendered DOM nodes dropped (${fixture.domNodeCount}/${previous.domNodeCount})`);
    }
    if (ratioDrop(fixture.domTextLength, previous.domTextLength, 0.9)) {
      const currentSceneTextLength = fixtureTextLength(fixture);
      const message = `${label}: rendered text dropped (${fixture.domTextLength}/${previous.domTextLength})`;
      if (isCurrentTextCoverageFailure(fixture.domTextLength, currentSceneTextLength)) {
        contentFailures.push(message);
      } else {
        warnings.push(`${message}; current scene text coverage is stable`);
      }
    }
    if (fixture.imageCount < previous.imageCount) {
      contentFailures.push(`${label}: image count dropped (${fixture.imageCount}/${previous.imageCount})`);
    }
    if (fixture.tableCount < previous.tableCount) {
      contentFailures.push(`${label}: table count dropped (${fixture.tableCount}/${previous.tableCount})`);
    }
    const previousSlides = Array.isArray(previous.slides) ? previous.slides : [];
    const currentSlides = Array.isArray(fixture.slides) ? fixture.slides : [];
    if (previousSlides.length > 0 && currentSlides.length < fixture.slideCount) {
      warnings.push(`${label}: current summary has partial slide stats (${currentSlides.length}/${fixture.slideCount})`);
    }
    if (previousSlides.length > 0 && previousSlides.length < previous.slideCount) {
      warnings.push(`${label}: baseline summary has partial slide stats (${previousSlides.length}/${previous.slideCount})`);
    }
    const previousSlideById = new Map(previousSlides.map((slide) => [slide.id, slide]));
    const previousSlideByIndex = new Map(previousSlides.map((slide) => [slide.index, slide]));
    for (const slide of currentSlides) {
      const previousSlide = previousSlideById.get(slide.id) ?? previousSlideByIndex.get(slide.index);
      if (!previousSlide) {
        warnings.push(`${label} slide ${slide.index + 1}: no baseline slide`);
        continue;
      }
      comparedSlides += 1;
      const slideLabel = `${label} slide ${slide.index + 1}`;
      if (ratioDrop(slide.sceneNodes, previousSlide.sceneNodes, 0.95)) {
        contentFailures.push(`${slideLabel}: scene nodes dropped (${slide.sceneNodes}/${previousSlide.sceneNodes})`);
      }
      if (ratioDrop(slide.domNodes, previousSlide.domNodes, 0.95)) {
        renderingFailures.push(`${slideLabel}: rendered DOM nodes dropped (${slide.domNodes}/${previousSlide.domNodes})`);
      }
      if (ratioDrop(slide.domTextLength, previousSlide.domTextLength, 0.9)) {
        const message = `${slideLabel}: rendered text dropped (${slide.domTextLength}/${previousSlide.domTextLength})`;
        if (isCurrentTextCoverageFailure(slide.domTextLength, slide.textLength)) {
          contentFailures.push(message);
        } else {
          warnings.push(`${message}; current scene text coverage is stable`);
        }
      }
      if (slide.images < previousSlide.images) {
        contentFailures.push(`${slideLabel}: image count dropped (${slide.images}/${previousSlide.images})`);
      }
      if (slide.tables < previousSlide.tables) {
        contentFailures.push(`${slideLabel}: table count dropped (${slide.tables}/${previousSlide.tables})`);
      }
    }
  }

  return {
    status: renderingFailures.length || contentFailures.length ? 'FAIL' : 'PASS',
    comparedFixtures,
    comparedSlides,
    baselineGeneratedAt: baseline.generatedAt,
    baselineSummaryPath: baseline.summaryPath,
    warnings,
    renderingFailures,
    contentFailures,
  };
}

function fixtureTextLength(fixture) {
  if (Number.isFinite(fixture.textLength)) return fixture.textLength;
  return (fixture.slides ?? []).reduce((sum, slide) => sum + (slide.textLength ?? 0), 0);
}

function isCurrentTextCoverageFailure(domTextLength, sceneTextLength) {
  if (!Number.isFinite(sceneTextLength) || sceneTextLength <= 0) return false;
  return domTextLength < Math.floor(sceneTextLength * 0.9);
}

async function runHighFrequencyEdit(client, baseUrl, options, runId) {
  await navigateToPpt(client, `${baseUrl}/?surface=ppt&pptFixture=0&bust=ppt-editor-smoke-index-${Date.now()}`);
  const fixtureCount = await evaluate(client, `document.querySelector('#pptFixtureSelect')?.options.length ?? 1`);
  const profileFixtures = pickProfileFixtureIndexes(fixtureCount, options);
  const fixtureProfiles = [];
  for (const fixtureIndex of profileFixtures) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-profile-${Date.now()}`,
    );
    fixtureProfiles.push(
      await evaluate(
        client,
        `(${highFrequencyProfileExpression.toString()})(${JSON.stringify(fixtureIndex)})`,
      ),
    );
  }
  const sampledProfiles = pickHighFrequencyFixtures(fixtureProfiles, options.highFrequencySampleLimit);
  const sampledFixtures = sampledProfiles.map((profile) => profile.fixtureIndex);
  const cases = [];
  const screenshotPaths = [];
  const screenshotDir = join(options.outputDir, `high-frequency-edit-screenshots-${runId}`);
  if (options.screenshots) await mkdir(screenshotDir, { recursive: true });
  for (const fixtureIndex of sampledFixtures) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-smoke-${Date.now()}`,
    );
    const editCase = await evaluate(
      client,
      `(${highFrequencyExpression.toString()})(${JSON.stringify(fixtureIndex)})`,
    );
    if (options.screenshots) {
      const visualTargets = [
        {
          kind: 'text',
          slideId: editCase.textSlideId,
          expectedText: editCase.textValue,
        },
      ];
      if (editCase.tableSlideId && editCase.tableValue && editCase.tableObserved) {
        visualTargets.push({
          kind: 'table',
          slideId: editCase.tableSlideId,
          expectedText: editCase.tableValue,
        });
      }
      editCase.visuals = [];
      for (const target of visualTargets) {
        const visualTarget = await evaluate(
          client,
          `(${collectHighFrequencyVisualExpression.toString()})(${JSON.stringify(target)})`,
        );
        const shot = await client.send('Page.captureScreenshot', { format: 'png', fromSurface: true });
        const bytes = Buffer.from(shot.data, 'base64');
        const path = join(
          screenshotDir,
          `fixture-${String(fixtureIndex).padStart(2, '0')}-${target.kind}-edited.png`,
        );
        await writeFile(path, bytes);
        const visual = analyzeAuditScreenshot({ ...visualTarget, png: pngSize(bytes) });
        const screenshot = { fixtureIndex, kind: target.kind, path, selected: visualTarget.selected, visual };
        screenshotPaths.push(screenshot);
        editCase.visuals.push(screenshot);
        for (const failure of visual.failures) editCase.failures.push(`${target.kind} screenshot: ${failure}`);
        for (const failure of visualTarget.failures) editCase.failures.push(failure);
      }
    }
    cases.push(editCase);
  }
  const primary = cases.find((item) => item.fixtureIndex === 4) ?? cases[0];
  const failures = cases.flatMap((item) => item.failures.map((failure) => `fixture ${item.fixtureIndex}: ${failure}`));
  return {
    fixtureIndex: primary?.fixtureIndex ?? null,
    sampledFixtures,
    cases,
    sampledProfiles,
    textShapeId: primary?.textShapeId ?? null,
    tableShapeId: cases.find((item) => item.tableShapeId)?.tableShapeId ?? primary?.tableShapeId ?? null,
    revision: cases.reduce((sum, item) => sum + item.revision, 0),
    status: failures.length > 0 ? '多模板高频编辑失败' : `多模板高频编辑通过：${cases.length} 个模板`,
    exportStatuses: cases.map((item) => item.status),
    screenshotPaths,
    failures,
  };
}

function pickProfileFixtureIndexes(fixtureCount, options) {
  if (options.highFrequencyProfileLimit === 0 || options.highFrequencyProfileLimit >= fixtureCount) {
    return Array.from({ length: fixtureCount }, (_, index) => index);
  }
  const indexes = new Set();
  const limit = Math.min(fixtureCount, options.highFrequencyProfileLimit);
  for (let index = 0; index < limit; index++) indexes.add(index);
  if (fixtureCount > 4) indexes.add(4);
  indexes.add(fixtureCount - 1);
  return [...indexes].filter((index) => index >= 0 && index < fixtureCount).sort((a, b) => a - b);
}

function pickHighFrequencyFixtures(profiles, limit) {
  const editable = profiles.filter((profile) => profile.editableTextCount > 0);
  const selected = [];
  const add = (profile, reason) => {
    if (!profile) return;
    const existing = selected.find((item) => item.fixtureIndex === profile.fixtureIndex);
    if (existing) {
      if (!existing.reasons.includes(reason)) existing.reasons.push(reason);
      existing.reason = existing.reasons.join(', ');
      return;
    }
    profile.reasons = [reason];
    selected.push({ ...profile, reason });
  };
  const byDesc = (field) => [...editable].sort((a, b) => b[field] - a[field]);
  const byAsc = (field) => [...editable].sort((a, b) => a[field] - b[field]);

  add(
    byDesc('editableTableCount').find((profile) => profile.editableTableCount > 0),
    'editable table',
  );
  add(byDesc('imageCount')[0], 'image rich');
  add(byDesc('sceneTextLength')[0], 'text rich');
  add(byDesc('slideCount')[0], 'many slides');
  add(byAsc('slideCount')[0], 'compact deck');
  add(editable.find((profile) => profile.fixtureIndex === 4), 'fixture 4 regression anchor');
  add(editable[0], 'fallback editable deck');

  return selected.slice(0, limit);
}

function highFrequencyProfileExpression(fixtureIndex) {
  return (() => {
    const textOf = (value) => (value ?? '').replace(/\s+/g, ' ').trim();
    const visibleTextOf = (node) => {
      if (!node?.table) return textOf(node?.text);
      const continuationCells = new Set(
        (node.table.mergedCells ?? [])
          .filter((merge) => merge.hMerge || merge.vMerge)
          .map((merge) => `${merge.row}:${merge.column}`),
      );
      const parts = [];
      const rowCount = Math.max(node.table.rowCount ?? 0, node.table.values?.length ?? 0);
      const columnCount = Math.max(
        node.table.columnCount ?? 0,
        ...(node.table.values ?? []).map((row) => row.length),
        0,
      );
      for (let row = 0; row < rowCount; row++) {
        for (let column = 0; column < columnCount; column++) {
          if (continuationCells.has(`${row}:${column}`)) continue;
          parts.push(node.table.values?.[row]?.[column] ?? '');
        }
      }
      return textOf(parts.join(''));
    };
    const scene = window.__cocoPptDev.session().scene();
    const fixtureName =
      document.querySelector('#pptFixtureSelect option:checked')?.textContent?.trim() ??
      document.querySelector('#pptSummary')?.textContent?.trim() ??
      `fixture ${fixtureIndex}`;
    const nodes = scene.slides.flatMap((slide) => slide.nodes);
    return {
      fixtureIndex,
      fixtureName,
      slideCount: scene.slides.length,
      sceneNodeCount: nodes.length,
      editableTextCount: nodes.filter((node) => node.capabilities.editText && typeof node.text === 'string' && node.text.length > 0).length,
      editableTableCount: nodes.filter((node) => node.kind === 'table' && node.capabilities.editTableCells).length,
      imageCount: nodes.filter((node) => node.kind === 'picture' || node.imagePath).length,
      tableCount: nodes.filter((node) => node.kind === 'table' || node.table).length,
      sceneTextLength: nodes.reduce((sum, node) => sum + visibleTextOf(node).length, 0),
    };
  })();
}

function highFrequencyExpression(fixtureIndex) {
  const sameColor = (actual, expected) =>
    typeof actual === 'string' && actual.toLowerCase() === expected;
  const setInput = (selector, value) => {
    const input = document.querySelector(selector);
    input.value = String(value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const waitFor = async (predicate, label, timeoutMs = 30_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const allNodes = () => window.__cocoPptDev.session().scene().slides.flatMap((slide) => slide.nodes);
  const findNode = (shapeId, slideId) => {
    if (slideId) {
      return window.__cocoPptDev
        .session()
        .scene()
        .slides.find((slide) => slide.id === slideId)
        ?.nodes.find((node) => node.id === shapeId);
    }
    return allNodes().find((node) => node.id === shapeId);
  };
  const firstEditableTableCell = (table) => {
    if (!table) return null;
    const continuationCells = new Set(
      (table.mergedCells ?? [])
        .filter((merge) => merge.hMerge || merge.vMerge)
        .map((merge) => `${merge.row}:${merge.column}`),
    );
    for (let row = 0; row < table.rowCount; row++) {
      for (let column = 0; column < table.columnCount; column++) {
        if (!continuationCells.has(`${row}:${column}`)) return { row, column };
      }
    }
    return null;
  };
  const selectFirst = async (predicate) => {
    const editor = window.__cocoPptDev;
    for (const slide of editor.session().scene().slides) {
      const node = slide.nodes.find(predicate);
      if (!node) continue;
      await editor.dispatch({ kind: 'select-shape', slideIdOrIndex: slide.id, shapeIdOrName: node.id });
      const selectedNode = editor
        .session()
        .scene()
        .slides.find((candidate) => candidate.id === slide.id)
        ?.nodes.find((candidate) => candidate.id === node.id);
      return {
        slideId: slide.id,
        shapeId: node.id,
        kind: selectedNode?.kind ?? node.kind,
        text: selectedNode?.text ?? node.text,
        table: selectedNode?.table ?? node.table,
        bounds: { ...(selectedNode?.bounds ?? node.bounds) },
      };
    }
    return null;
  };
  const waitToolbarReady = (selector) =>
    waitFor(() => {
      const status = document.querySelector('#pptEditStatus')?.textContent ?? '';
      const control = document.querySelector(selector);
      return Boolean(control && !control.disabled && status !== '正在应用...');
    }, `${selector} ready`);

  return (async () => {
    const failures = [];
    const textValue = `PPT 高频文字编辑 ${fixtureIndex}`;
    const tableValue = `PPT TABLE 高频编辑 ${fixtureIndex}`;
    const textTarget = await selectFirst(
      (node) =>
        node.capabilities.editText &&
        node.capabilities.editTextStyle &&
        node.capabilities.editGeometry &&
        node.capabilities.editShapeStyle &&
        typeof node.text === 'string' &&
        node.text.length > 0,
    );
    if (!textTarget) throw new Error('No high-frequency editable text shape found.');
    await waitToolbarReady('#pptApplyTextButton');

    setInput('#pptTextInput', textValue);
    document.querySelector('#pptApplyTextButton').click();
    await waitFor(() => {
      const node = findNode(textTarget.shapeId, textTarget.slideId);
      return node?.text === textValue;
    }, 'text edit');

    await waitToolbarReady('#pptApplyStyleButton');
    setInput('#pptFontSizeInput', 36);
    setInput('#pptTextColorInput', '#336699');
    setInput('#pptFillColorInput', '#ffcc00');
    setInput('#pptLineColorInput', '#2255aa');
    setInput('#pptLineWidthInput', 19050);
    document.querySelector('#pptApplyStyleButton').click();
    await waitFor(() => {
      const node = findNode(textTarget.shapeId, textTarget.slideId);
      return (
        node?.textStyle?.fontSize === 36 &&
        sameColor(node.textStyle?.color, '#336699') &&
        sameColor(node?.fillColor, '#ffcc00') &&
        sameColor(node?.lineColor, '#2255aa') &&
        node?.lineWidth === 19050
      );
    }, 'text and shape style edit');

    const nextGeometry = {
      left: Math.round(textTarget.bounds.left + 120000),
      top: Math.round(textTarget.bounds.top + 90000),
      width: Math.round(textTarget.bounds.width + 160000),
      height: Math.round(textTarget.bounds.height + 100000),
    };
    await waitToolbarReady('#pptApplyGeometryButton');
    setInput('#pptLeftInput', nextGeometry.left);
    setInput('#pptTopInput', nextGeometry.top);
    setInput('#pptWidthInput', nextGeometry.width);
    setInput('#pptHeightInput', nextGeometry.height);
    document.querySelector('#pptApplyGeometryButton').click();
    await waitFor(() => {
      const bounds = findNode(textTarget.shapeId, textTarget.slideId)?.bounds;
      return (
        bounds?.left === nextGeometry.left &&
        bounds.top === nextGeometry.top &&
        bounds.width === nextGeometry.width &&
        bounds.height === nextGeometry.height
      );
    }, 'geometry edit');

    const tableTarget = await selectFirst(
      (node) => node.kind === 'table' && node.capabilities.editTableCells && node.table?.rowCount > 0 && node.table?.columnCount > 0,
    );
    let tableStatus = tableTarget ? 'not edited' : 'no editable table in fixture';
    let tableObserved = false;
    if (tableTarget) {
      const tableCell = firstEditableTableCell(findNode(tableTarget.shapeId, tableTarget.slideId)?.table ?? tableTarget.table);
      if (!tableCell) {
        tableStatus = 'no directly editable table cell in fixture';
        failures.push(tableStatus);
      } else {
      await waitToolbarReady('#pptApplyTableCellButton');
      setInput('#pptTableRowInput', tableCell.row);
      setInput('#pptTableColumnInput', tableCell.column);
      setInput('#pptTableCellTextInput', tableValue);
      document.querySelector('#pptApplyTableCellButton').click();
      try {
        await waitFor(
          () => findNode(tableTarget.shapeId, tableTarget.slideId)?.table?.values[tableCell.row]?.[tableCell.column] === tableValue,
          'table cell edit',
          15_000,
        );
        tableObserved = true;
        tableStatus = `table cell ${tableCell.row},${tableCell.column} edit observed in memory`;
      } catch (error) {
        const currentTable = findNode(tableTarget.shapeId, tableTarget.slideId)?.table;
        const debug = {
          editStatus: document.querySelector('#pptEditStatus')?.textContent ?? '',
          selection: window.__cocoPptDev.session().selection(),
          rowInput: document.querySelector('#pptTableRowInput')?.value ?? '',
          columnInput: document.querySelector('#pptTableColumnInput')?.value ?? '',
          cellInput: document.querySelector('#pptTableCellTextInput')?.value ?? '',
          buttonDisabled: Boolean(document.querySelector('#pptApplyTableCellButton')?.disabled),
          values: currentTable?.values,
        };
        tableStatus = `table cell edit not observable: ${error instanceof Error ? error.message : String(error)} ${JSON.stringify(debug)}`;
        failures.push(tableStatus);
      }
      }
    }

    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () =>
        /^(Verified \d+ slides|导出校验通过：\d+ 页)$/.test(
          document.querySelector('#pptEditStatus')?.textContent ?? '',
        ),
      'export verification',
      60_000,
    );

    const textNode = findNode(textTarget.shapeId, textTarget.slideId);
    if (textNode?.text !== textValue) failures.push('text edit did not persist');
    if (textNode?.textStyle?.fontSize !== 36) failures.push('font size did not persist');
    if (!sameColor(textNode?.textStyle?.color, '#336699')) failures.push('text color did not persist');
    if (!sameColor(textNode?.fillColor, '#ffcc00')) failures.push('fill color did not persist');
    if (!sameColor(textNode?.lineColor, '#2255aa')) failures.push('line color did not persist');
    if (textNode?.lineWidth !== 19050) failures.push('line width did not persist');
    if (textNode?.bounds.left !== nextGeometry.left) failures.push('geometry left did not persist');
    if (
      tableTarget &&
      tableObserved &&
      !findNode(tableTarget.shapeId, tableTarget.slideId)?.table?.values.some((row) =>
        row.includes(tableValue),
      )
    ) {
      failures.push('table cell edit did not persist');
    }

    return {
      fixtureIndex,
      textSlideId: textTarget.slideId,
      textShapeId: textTarget.shapeId,
      tableSlideId: tableTarget?.slideId ?? null,
      tableShapeId: tableTarget?.shapeId ?? null,
      textValue,
      tableValue: tableTarget ? tableValue : null,
      tableStatus,
      tableObserved,
      revision: window.__cocoPptDev.session().revision(),
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
      failures,
    };
  })();
}

function collectHighFrequencyVisualExpression(target) {
  return (async () => {
    const editor = window.__cocoPptDev;
    await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: target.slideId });
    const canvas = document.querySelector('#pptEditorHost')?.shadowRoot?.querySelector('.coco-ppt-canvas');
    const scene = editor.session().scene();
    const slide = scene.slides.find((item) => item.id === target.slideId);
    const start = Date.now();
    while (Date.now() - start < 5_000) {
      const renderedNodes = canvas?.querySelectorAll('.coco-ppt-node').length ?? 0;
      if (editor.selection()?.slideId === slide?.id && renderedNodes === (slide?.nodes.length ?? 0)) break;
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
    const failures = [];
    const textContent = (canvas?.textContent ?? '').replace(/\s+/g, ' ').trim();
    const textOf = (value) => (value ?? '').replace(/\s+/g, ' ').trim();
    const visibleTextOf = (node) => {
      if (!node?.table) return textOf(node?.text);
      const continuationCells = new Set(
        (node.table.mergedCells ?? [])
          .filter((merge) => merge.hMerge || merge.vMerge)
          .map((merge) => `${merge.row}:${merge.column}`),
      );
      const parts = [];
      const rowCount = Math.max(node.table.rowCount ?? 0, node.table.values?.length ?? 0);
      const columnCount = Math.max(
        node.table.columnCount ?? 0,
        ...(node.table.values ?? []).map((row) => row.length),
        0,
      );
      for (let row = 0; row < rowCount; row++) {
        for (let column = 0; column < columnCount; column++) {
          if (continuationCells.has(`${row}:${column}`)) continue;
          parts.push(node.table.values?.[row]?.[column] ?? '');
        }
      }
      return textOf(parts.join(''));
    };
    if (!textContent.includes(target.expectedText)) failures.push(`edited ${target.kind} text is not visible in DOM`);
    const selected = {
      slideId: slide?.id ?? null,
      slideIndex: slide ? scene.slides.indexOf(slide) : -1,
      sceneNodes: slide?.nodes.length ?? 0,
      domNodes: canvas?.querySelectorAll('.coco-ppt-node').length ?? 0,
      sceneTextLength: (slide?.nodes ?? []).reduce(
        (sum, node) => sum + visibleTextOf(node).length,
        0,
      ),
      domTextLength: textContent.length,
      viewport: {
        width: window.innerWidth,
        height: window.innerHeight,
      },
      canvasRect: canvas
        ? {
            x: Math.round(canvas.getBoundingClientRect().x),
            y: Math.round(canvas.getBoundingClientRect().y),
            width: Math.round(canvas.getBoundingClientRect().width),
            height: Math.round(canvas.getBoundingClientRect().height),
          }
        : { x: 0, y: 0, width: 0, height: 0 },
      canvasAreaRatio: canvas
        ? (canvas.getBoundingClientRect().width * canvas.getBoundingClientRect().height) /
          Math.max(1, window.innerWidth * window.innerHeight)
        : 0,
    };
    return {
      kind: target.kind,
      slideIndex: selected.slideIndex,
      selected,
      failures,
    };
  })();
}

async function runInteractionEditCases(client, baseUrl, options, sampledFixtures = [], runId) {
  await navigateToPpt(client, `${baseUrl}/?surface=ppt&pptFixture=0&bust=ppt-editor-interactions-${Date.now()}`);
  const fixtureCount = await evaluate(client, `document.querySelector('#pptFixtureSelect')?.options.length ?? 1`);
  const fixtureIndexes = pickInteractionFixtureIndexes(fixtureCount, sampledFixtures, options);
  const results = [];
  const tableCells = [];
  const mergedTableBoundaries = [];
  const screenshotPaths = [];
  for (const fixtureIndex of fixtureIndexes) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-interactions-${Date.now()}`,
    );
    const result = await runInteractionEditCaseWithScreenshots(client, fixtureIndex, options, runId);
    results.push(result.result);
    screenshotPaths.push(...result.screenshotPaths);
    if (!options.fast) {
      tableCells.push(await runTableCellClickInteraction(client, baseUrl, fixtureIndex));
      mergedTableBoundaries.push(
        await runMergedTableToolbarBoundaryInteraction(client, baseUrl, fixtureIndex),
      );
    }
  }
  if (options.fast) {
    const tableFixtureIndexes = uniqueFixtureIndexes([...sampledFixtures, ...fixtureIndexes]);
    let tableResult = null;
    for (const fixtureIndex of tableFixtureIndexes) {
      tableResult = await runTableCellClickInteraction(client, baseUrl, fixtureIndex);
      tableCells.push(tableResult);
      if (tableResult.cases.some((item) => item.shapeId)) break;
    }
  }
  const toolbar = await runToolbarStateInteraction(client, baseUrl, fixtureIndexes[0] ?? 0);
  return {
    fixtureIndex: fixtureIndexes[0] ?? 0,
    fixtureIndexes,
    cases: [
      ...results.flatMap((result) => result.cases),
      ...tableCells.flatMap((result) => result.cases),
      ...mergedTableBoundaries.flatMap((result) => result.cases),
      ...toolbar.cases,
    ],
    failures: [
      ...results.flatMap((result) => result.failures),
      ...tableCells.flatMap((result) => result.failures),
      ...mergedTableBoundaries.flatMap((result) => result.failures),
      ...toolbar.failures,
    ],
    screenshotPaths,
  };
}

function uniqueFixtureIndexes(values) {
  return values.filter(
    (value, index, array) =>
      Number.isInteger(value) && value >= 0 && array.indexOf(value) === index,
  );
}

async function runInteractionEditCaseWithScreenshots(client, fixtureIndex, options, runId) {
  const screenshotPaths = [];
  const captureKey = `__cocoPptInteractionCapture_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  const pending = evaluate(
    client,
    `(${interactionEditExpression.toString()})(${JSON.stringify(fixtureIndex)}, ${JSON.stringify(captureKey)})`,
  );
  let settled = false;
  pending.finally(() => {
    settled = true;
  }).catch(() => undefined);
  let lastProgress = '';
  let lastProgressPoll = 0;
  while (!settled) {
    if (Date.now() - lastProgressPoll > 5000) {
      lastProgressPoll = Date.now();
      const progress = await evaluate(
        client,
        `(function(){ return window[${JSON.stringify(`${captureKey}:progress`)}] ?? null; })()`,
      ).catch(() => null);
      if (progress?.label && progress.label !== lastProgress) {
        lastProgress = progress.label;
        logStep(`interaction fixture ${fixtureIndex}: ${progress.label}`);
      }
    }
    const request = await evaluate(
      client,
      `(function(){ const item = window[${JSON.stringify(captureKey)}]; return item && item.ready && !item.captured ? item : null; })()`,
    );
    if (request) {
      const screenshotDir = join(options.outputDir, `interaction-screenshots-${runId}`);
      await mkdir(screenshotDir, { recursive: true });
      const path = join(
        screenshotDir,
        `fixture-${String(fixtureIndex).padStart(2, '0')}-${request.kind}.png`,
      );
      const shot = await client.send('Page.captureScreenshot', { format: 'png', fromSurface: true });
      await writeFile(path, Buffer.from(shot.data, 'base64'));
      screenshotPaths.push({
        fixtureIndex,
        kind: request.kind,
        path,
        selected: request.selected ?? null,
      });
      await evaluate(
        client,
        `(function(){ const item = window[${JSON.stringify(captureKey)}]; if (item) { item.captured = true; item.resume = true; } })()`,
      );
    } else {
      await sleep(50);
    }
  }
  return { result: await pending, screenshotPaths };
}

function pickInteractionFixtureIndexes(fixtureCount, sampledFixtures, options) {
  const limit = Math.min(fixtureCount, options.full ? 5 : 1);
  const indexes = [];
  const add = (fixtureIndex) => {
    if (!Number.isInteger(fixtureIndex) || fixtureIndex < 0 || fixtureIndex >= fixtureCount) return;
    if (!indexes.includes(fixtureIndex)) indexes.push(fixtureIndex);
  };
  add(Math.min(4, Math.max(0, fixtureCount - 1)));
  for (const fixtureIndex of sampledFixtures) add(fixtureIndex);
  add(0);
  for (let fixtureIndex = 0; indexes.length < limit && fixtureIndex < fixtureCount; fixtureIndex++) {
    add(fixtureIndex);
  }
  return indexes.slice(0, limit);
}

async function runTableCellClickInteraction(client, baseUrl, fixtureIndex) {
  await navigateToPpt(
    client,
    `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-table-cell-click-${Date.now()}`,
  );
  const target = await evaluate(
    client,
    `(${tableCellClickTargetExpression.toString()})(${JSON.stringify(fixtureIndex)})`,
  );
  if (!target) {
    return {
      cases: [
        {
          kind: 'table-cell-direct-edit',
          fixtureIndex,
          slideId: null,
          shapeId: null,
          status: 'skipped: no editable table in fixture',
        },
      ],
      failures: [],
    };
  }
  const result = await evaluate(
    client,
    `(${tableCellDirectEditCommitExpression.toString()})(${JSON.stringify(target)})`,
  );
  return { cases: [result.case], failures: result.failures };
}

async function runMergedTableToolbarBoundaryInteraction(client, baseUrl, fixtureIndex) {
  await navigateToPpt(
    client,
    `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-merged-table-boundary-${Date.now()}`,
  );
  const target = await evaluate(
    client,
    `(${mergedTableContinuationTargetExpression.toString()})(${JSON.stringify(fixtureIndex)})`,
  );
  if (!target) {
    return {
      cases: [
        {
          kind: 'merged-table-toolbar-boundary',
          fixtureIndex,
          slideId: null,
          shapeId: null,
          status: 'skipped: no merged continuation table cell in fixture',
        },
      ],
      failures: [],
    };
  }
  const result = await evaluate(
    client,
    `(${mergedTableToolbarBoundaryExpression.toString()})(${JSON.stringify(target)})`,
  );
  return { cases: [result.case], failures: result.failures };
}

async function runToolbarStateInteraction(client, baseUrl, fixtureIndex) {
  await navigateToPpt(
    client,
    `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-toolbar-state-${Date.now()}`,
  );
  const result = await evaluate(
    client,
    `(${toolbarStateExpression.toString()})(${JSON.stringify(fixtureIndex)})`,
  );
  return { cases: [result.case], failures: result.failures };
}

async function dispatchMouseClick(client, x, y) {
  await client.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x,
    y,
    button: 'left',
    clickCount: 1,
  });
  await client.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x,
    y,
    button: 'left',
    clickCount: 1,
  });
}

async function dispatchMouseDoubleClick(client, x, y) {
  await dispatchMouseClick(client, x, y);
  await client.send('Input.dispatchMouseEvent', {
    type: 'mousePressed',
    x,
    y,
    button: 'left',
    clickCount: 2,
  });
  await client.send('Input.dispatchMouseEvent', {
    type: 'mouseReleased',
    x,
    y,
    button: 'left',
    clickCount: 2,
  });
}

function interactionEditExpression(fixtureIndex, captureKey) {
  const waitFor = async (predicate, label, timeoutMs = 15_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const mouse = (type, x, y, options = {}) =>
    new MouseEvent(type, {
      bubbles: true,
      cancelable: true,
      composed: true,
      clientX: x,
      clientY: y,
      button: 0,
      ...options,
    });
  const editorRoot = () => document.querySelector('#pptEditorHost')?.shadowRoot;
  const snapGuideCount = (axis) =>
    editorRoot()?.querySelectorAll(`.coco-ppt-snap-guide${axis ? `[data-axis="${axis}"]` : ''}`).length ?? 0;
  const snapReferenceCount = () =>
    editorRoot()?.querySelectorAll('.coco-ppt-node[data-snap-reference="true"]').length ?? 0;
  const shapeSnapGuideCount = () =>
    editorRoot()?.querySelectorAll('.coco-ppt-snap-guide[data-kind="shape"]').length ?? 0;
  const shapeSnapReferenceIds = () =>
    [...(editorRoot()?.querySelectorAll('.coco-ppt-snap-guide[data-kind="shape"]') ?? [])]
      .map((guide) => guide.dataset.referenceShapeId)
      .filter(Boolean);
  const setInput = (selector, value) => {
    const input = document.querySelector(selector);
    input.value = String(value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const nodeElement = (shapeId) =>
    document
      .querySelector('#pptEditorHost')
      ?.shadowRoot?.querySelector(`.coco-ppt-node[data-shape-id="${CSS.escape(shapeId)}"]`);
  const contextMenuElement = () =>
    document.querySelector('#pptEditorHost')?.shadowRoot?.querySelector('.coco-ppt-context-menu');
  const contextMenuButton = (label) =>
    [...(contextMenuElement()?.querySelectorAll('button') ?? [])].find((button) =>
      button.textContent?.includes(label),
    );
  const editableTarget = (element) => element?.querySelector('.coco-ppt-text-layer') ?? element;
  const allNodes = () => window.__cocoPptDev.session().scene().slides.flatMap((slide) => slide.nodes);
  const findNode = (shapeId) => allNodes().find((node) => node.id === shapeId);
  const setProgress = (label) => {
    if (!captureKey) return;
    window[`${captureKey}:progress`] = { label, at: Date.now() };
  };
  const textOf = (value) => (value ?? '').replace(/\s+/g, ' ').trim();
  const visibleTextOf = (node) => {
    if (!node?.table) return textOf(node?.text);
    const continuationCells = new Set(
      (node.table.mergedCells ?? [])
        .filter((merge) => merge.hMerge || merge.vMerge)
        .map((merge) => `${merge.row}:${merge.column}`),
    );
    const parts = [];
    const rowCount = Math.max(node.table.rowCount ?? 0, node.table.values?.length ?? 0);
    const columnCount = Math.max(
      node.table.columnCount ?? 0,
      ...(node.table.values ?? []).map((row) => row.length),
      0,
    );
    for (let row = 0; row < rowCount; row++) {
      for (let column = 0; column < columnCount; column++) {
        if (continuationCells.has(`${row}:${column}`)) continue;
        parts.push(node.table.values?.[row]?.[column] ?? '');
      }
    }
    return textOf(parts.join(''));
  };
  const currentSelectionSummary = () => {
    const canvas = editorRoot()?.querySelector('.coco-ppt-canvas');
    const selection = window.__cocoPptDev.selection();
    const slide =
      selection?.kind === 'shape'
        ? window.__cocoPptDev.session().scene().slides.find((item) => item.id === selection.slideId)
        : null;
    return {
      slideId: slide?.id ?? null,
      slideIndex: slide ? window.__cocoPptDev.session().scene().slides.indexOf(slide) : -1,
      sceneNodes: slide?.nodes.length ?? 0,
      domNodes: canvas?.querySelectorAll('.coco-ppt-node').length ?? 0,
      sceneTextLength: (slide?.nodes ?? []).reduce(
        (sum, node) => sum + visibleTextOf(node).length,
        0,
      ),
      domTextLength: (canvas?.textContent ?? '').replace(/\s+/g, ' ').trim().length,
    };
  };
  const requestSnapshot = async (kind) => {
    if (!captureKey) return;
    window[captureKey] = {
      ready: true,
      captured: false,
      resume: false,
      kind,
      selected: currentSelectionSummary(),
    };
    await waitFor(() => window[captureKey]?.resume === true, `snapshot ${kind}`, 10_000);
    delete window[captureKey];
  };
  const pickTarget = async () => {
    const editor = window.__cocoPptDev;
    for (const slide of editor.session().scene().slides) {
      const node = slide.nodes.find(
        (item) =>
          item.capabilities.editText &&
          item.capabilities.editGeometry &&
          typeof item.text === 'string' &&
          item.text.trim().length > 0 &&
          item.bounds.width > 100000 &&
          item.bounds.height > 100000,
      );
      if (!node) continue;
      await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
      await waitFor(() => Boolean(nodeElement(node.id)), `shape element ${node.id}`);
      return { slideId: slide.id, shapeId: node.id, originalText: node.text, bounds: { ...node.bounds } };
    }
    return null;
  };
  return (async () => {
    const failures = [];
    const cases = [];
    setProgress('pick target');
    const target = await pickTarget();
    if (!target) {
      return { cases, failures: ['interaction: no editable text geometry target found'] };
    }

    const key = (key, options = {}) =>
      new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        key,
        ...options,
      });
    const textValue = `PPT 双击编辑 ${fixtureIndex}`;
    const element = nodeElement(target.shapeId);
    if (!element) {
      return { cases, failures: [`interaction: missing element for ${target.shapeId}`] };
    }
    setProgress('double-click text');
    element.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
    await waitFor(() => editableTarget(element)?.contentEditable === 'true', 'double-click text editing');
    const editable = editableTarget(element);
    editable.textContent = textValue;
    editable.dispatchEvent(new FocusEvent('blur', { bubbles: true }));
    await waitFor(() => findNode(target.shapeId)?.text === textValue, 'double-click edit commit');
    cases.push({
      kind: 'double-click-text',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: target.shapeId,
      status: 'committed to memory',
    });

    setProgress('toolbar undo redo');
    const undoButton = document.querySelector('#pptUndoButton');
    const redoButton = document.querySelector('#pptRedoButton');
    if (!undoButton || !redoButton) {
      failures.push('interaction: undo/redo toolbar buttons unavailable');
    } else {
      await waitFor(() => undoButton.disabled === false, 'undo button enabled after text edit');
      undoButton.click();
      await waitFor(() => findNode(target.shapeId)?.text === target.originalText, 'toolbar undo text edit');
      if (redoButton.disabled) {
        failures.push('interaction: redo button disabled after undo');
      }
      redoButton.click();
      await waitFor(() => findNode(target.shapeId)?.text === textValue, 'toolbar redo text edit');
      cases.push({
        kind: 'toolbar-undo-redo',
        fixtureIndex,
        slideId: target.slideId,
        shapeId: target.shapeId,
        status: 'undo restored original text and redo restored edited text',
      });
    }

    setProgress('mouse drag move');
    const beforeDrag = { ...findNode(target.shapeId).bounds };
    const dragElement = nodeElement(target.shapeId);
    const rect = dragElement.getBoundingClientRect();
    const startX = Math.round(rect.left + Math.max(4, Math.min(24, rect.width / 2)));
    const startY = Math.round(rect.top + Math.max(4, Math.min(24, rect.height / 2)));
    const dx = 48;
    const dy = 36;
    const beforeMoveStyle = {
      left: dragElement.style.left,
      top: dragElement.style.top,
    };
    dragElement.dispatchEvent(mouse('mousedown', startX, startY));
    window.dispatchEvent(mouse('mousemove', startX + dx, startY + dy));
    const movedStyle = {
      left: dragElement.style.left,
      top: dragElement.style.top,
    };
    window.dispatchEvent(mouse('mouseup', startX + dx, startY + dy));
    dragElement.dispatchEvent(mouse('click', startX + dx, startY + dy));
    await waitFor(() => {
      const bounds = findNode(target.shapeId)?.bounds;
      return bounds && (bounds.left !== beforeDrag.left || bounds.top !== beforeDrag.top);
    }, 'drag geometry commit');
    const afterDrag = { ...findNode(target.shapeId).bounds };
    if (movedStyle.left === beforeMoveStyle.left && movedStyle.top === beforeMoveStyle.top) {
      failures.push('interaction: drag did not update DOM position before commit');
    }
    cases.push({
      kind: 'mouse-drag-move',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: target.shapeId,
      status: `moved ${Math.round(afterDrag.left - beforeDrag.left)},${Math.round(afterDrag.top - beforeDrag.top)} EMU`,
    });

    const snapSlide = window.__cocoPptDev.session().scene().slides.find((slide) => slide.id === target.slideId);
    const snapTarget = snapSlide?.nodes.find((node) =>
      node.id !== target.shapeId &&
      node.kind !== 'line' &&
      node.capabilities.editGeometry &&
      !node.locked &&
      !node.id.includes(':') &&
      Math.abs(node.bounds.left - findNode(target.shapeId).bounds.left) > 100000,
    );
    if (snapTarget) {
      const snapElement = nodeElement(target.shapeId);
      const targetElement = nodeElement(snapTarget.id);
      const snapRect = snapElement?.getBoundingClientRect();
      if (!snapElement || !targetElement || !snapRect) {
        failures.push('interaction: missing element for drag snap alignment');
      } else {
        const snapStartX = Math.round(snapRect.left + Math.max(4, Math.min(24, snapRect.width / 2)));
        const snapStartY = Math.round(snapRect.top + Math.max(4, Math.min(24, snapRect.height / 2)));
        const snapDx = Math.round(
          Number.parseFloat(targetElement.style.left) - Number.parseFloat(snapElement.style.left) - 3,
        );
        if (Math.abs(snapDx) <= 3) {
          failures.push('interaction: drag snap target is too close to exercise');
        } else {
          setProgress('drag snap alt disabled');
          const scale = Number.parseFloat(targetElement.style.left) / snapTarget.bounds.left;
          const altBefore = { ...findNode(target.shapeId).bounds };
          snapElement.dispatchEvent(mouse('mousedown', snapStartX, snapStartY));
          window.dispatchEvent(mouse('mousemove', snapStartX + snapDx, snapStartY, { altKey: true }));
          if (snapGuideCount('x') !== 0) {
            failures.push('interaction: Alt drag rendered snap guide');
          }
          if (snapReferenceCount() > 0) {
            failures.push('interaction: Alt drag marked snap reference object');
          }
          const altStyleLeft = Number.parseFloat(snapElement.style.left);
          const altExpectedLeft = Number.isFinite(scale) && scale > 0 ? altStyleLeft / scale : null;
          window.dispatchEvent(mouse('mouseup', snapStartX + snapDx, snapStartY, { altKey: true }));
          if (altExpectedLeft === null) {
            await waitFor(() => findNode(target.shapeId)?.bounds.left !== altBefore.left, 'Alt drag geometry commit');
          } else {
            await waitFor(
              () => Math.abs((findNode(target.shapeId)?.bounds.left ?? Number.NaN) - altExpectedLeft) < 25000,
              'Alt drag geometry commit',
            );
          }
          const altBounds = findNode(target.shapeId)?.bounds;
          if (altBounds && Math.abs(altBounds.left - snapTarget.bounds.left) < 10000) {
            failures.push('interaction: Alt drag snapped to target');
          }
          await window.__cocoPptDev.dispatch({
            kind: 'set-shape-geometry',
            slideIdOrIndex: target.slideId,
            shapeIdOrName: target.shapeId,
            geometry: altBefore,
          });
          await waitFor(() => findNode(target.shapeId)?.bounds.left === altBefore.left, 'restore after Alt drag');
          cases.push({
            kind: 'drag-snap-alt-disabled',
            fixtureIndex,
            slideId: target.slideId,
            shapeId: target.shapeId,
            status: `Alt kept ${Math.round(altBounds?.left ?? 0)} EMU unsnapped`,
          });

          const restoredSnapElement = nodeElement(target.shapeId);
          const restoredSnapRect = restoredSnapElement?.getBoundingClientRect();
          if (!restoredSnapElement || !restoredSnapRect) {
            failures.push('interaction: missing restored element for drag snap alignment');
          } else {
            setProgress('drag snap align');
            const restoredSnapStartX = Math.round(
              restoredSnapRect.left + Math.max(4, Math.min(24, restoredSnapRect.width / 2)),
            );
            const restoredSnapStartY = Math.round(
              restoredSnapRect.top + Math.max(4, Math.min(24, restoredSnapRect.height / 2)),
            );
            restoredSnapElement.dispatchEvent(mouse('mousedown', restoredSnapStartX, restoredSnapStartY));
            window.dispatchEvent(mouse('mousemove', restoredSnapStartX + snapDx, restoredSnapStartY));
            const guideCount = snapGuideCount('x');
            if (guideCount === 0) {
              failures.push('interaction: drag snap guide did not render');
            }
            if (shapeSnapGuideCount() > 0 && snapReferenceCount() === 0) {
              failures.push('interaction: drag snap did not mark reference object');
            }
            await requestSnapshot('drag-snap-align');
            const snappedStyleLeft = Number.parseFloat(restoredSnapElement.style.left);
            const expectedLeft = Number.isFinite(scale) && scale > 0 ? snappedStyleLeft / scale : null;
            window.dispatchEvent(mouse('mouseup', restoredSnapStartX + snapDx, restoredSnapStartY));
            restoredSnapElement.dispatchEvent(mouse('click', restoredSnapStartX + snapDx, restoredSnapStartY));
            if (expectedLeft === null) {
              await waitFor(() => findNode(target.shapeId)?.bounds.left !== afterDrag.left, 'drag snap geometry commit');
            } else {
              await waitFor(
                () => Math.abs((findNode(target.shapeId)?.bounds.left ?? Number.NaN) - expectedLeft) < 25000,
                'drag snap geometry commit',
              );
            }
            const snappedBounds = findNode(target.shapeId)?.bounds;
            if (snapGuideCount(null) !== 0) {
              failures.push('interaction: drag snap guide was not cleared after mouseup');
            }
            if (snapReferenceCount() > 0) {
              failures.push('interaction: drag snap reference object was not cleared after mouseup');
            }
            cases.push({
              kind: 'drag-snap-align',
              fixtureIndex,
              slideId: target.slideId,
              shapeId: target.shapeId,
              status: `snapped near ${snapTarget.id}: ${Math.round(snappedBounds?.left ?? 0)} EMU`,
            });
          }
        }
      }
    }

    setProgress('keyboard nudge');
    const beforeNudge = { ...findNode(target.shapeId).bounds };
    window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
    await waitFor(() => findNode(target.shapeId)?.bounds.left !== beforeNudge.left, 'keyboard nudge');
    cases.push({
      kind: 'keyboard-nudge',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: target.shapeId,
      status: 'ArrowRight committed to memory',
    });

    setProgress('resize handle');
    const beforeResize = { ...findNode(target.shapeId).bounds };
    const resizeElement = nodeElement(target.shapeId);
    const handle = resizeElement?.querySelector('.coco-ppt-resize-handle');
    if (!resizeElement || !handle) {
      failures.push(`interaction: missing resize handle for ${target.shapeId}`);
    } else {
      const handleRect = handle.getBoundingClientRect();
      const resizeStartX = Math.round(handleRect.left + handleRect.width / 2);
      const resizeStartY = Math.round(handleRect.top + handleRect.height / 2);
      const beforeResizeStyle = {
        width: resizeElement.style.width,
        height: resizeElement.style.height,
      };
      handle.dispatchEvent(mouse('mousedown', resizeStartX, resizeStartY));
      window.dispatchEvent(mouse('mousemove', resizeStartX + 42, resizeStartY + 30));
      const resizedStyle = {
        width: resizeElement.style.width,
        height: resizeElement.style.height,
      };
      window.dispatchEvent(mouse('mouseup', resizeStartX + 42, resizeStartY + 30));
      resizeElement.dispatchEvent(mouse('click', resizeStartX + 42, resizeStartY + 30));
      await waitFor(() => {
        const bounds = findNode(target.shapeId)?.bounds;
        return bounds && (bounds.width !== beforeResize.width || bounds.height !== beforeResize.height);
      }, 'resize geometry commit');
      const afterResize = { ...findNode(target.shapeId).bounds };
      if (resizedStyle.width === beforeResizeStyle.width && resizedStyle.height === beforeResizeStyle.height) {
        failures.push('interaction: resize did not update DOM size before commit');
      }
      cases.push({
        kind: 'resize-handle',
        fixtureIndex,
        slideId: target.slideId,
        shapeId: target.shapeId,
        status: `resized ${Math.round(afterResize.width - beforeResize.width)},${Math.round(afterResize.height - beforeResize.height)} EMU`,
      });

      const resizeSnapSlide = window.__cocoPptDev.session().scene().slides.find((slide) => slide.id === target.slideId);
      const currentBounds = findNode(target.shapeId)?.bounds;
      const resizeSnapTarget = currentBounds
        ? resizeSnapSlide?.nodes
            .flatMap((node) => {
              if (
                node.id === target.shapeId ||
                node.kind === 'line' ||
                !node.capabilities.editGeometry ||
                node.locked ||
                node.id.includes(':')
              ) {
                return [];
              }
              return [
                { node, value: node.bounds.left },
                { node, value: node.bounds.left + node.bounds.width / 2 },
                { node, value: node.bounds.left + node.bounds.width },
              ];
            })
            .find(
              (candidate) =>
                candidate.value > currentBounds.left + 100000 &&
                Math.abs(candidate.value - (currentBounds.left + currentBounds.width)) > 100000,
            )
        : null;
      if (resizeSnapTarget && currentBounds) {
        const snapResizeElement = nodeElement(target.shapeId);
        const targetElement = nodeElement(resizeSnapTarget.node.id);
        const snapHandle = snapResizeElement?.querySelector('.coco-ppt-resize-handle');
        const snapHandleRect = snapHandle?.getBoundingClientRect();
        if (!snapResizeElement || !targetElement || !snapHandle || !snapHandleRect) {
          failures.push('interaction: missing element for resize snap alignment');
        } else {
          const snapResizeStartX = Math.round(snapHandleRect.left + snapHandleRect.width / 2);
          const snapResizeStartY = Math.round(snapHandleRect.top + snapHandleRect.height / 2);
          const scale = Number.parseFloat(targetElement.style.width) / resizeSnapTarget.node.bounds.width;
          const snapDx = Math.round(
            resizeSnapTarget.value * scale -
              (Number.parseFloat(snapResizeElement.style.left) + Number.parseFloat(snapResizeElement.style.width)) -
              3,
          );
          if (Math.abs(snapDx) <= 3) {
            failures.push('interaction: resize snap target is too close to exercise');
          } else {
            setProgress('resize snap align');
            snapHandle.dispatchEvent(mouse('mousedown', snapResizeStartX, snapResizeStartY));
            window.dispatchEvent(mouse('mousemove', snapResizeStartX + snapDx, snapResizeStartY));
            const guideCount = snapGuideCount('x');
            if (guideCount === 0) {
              failures.push('interaction: resize snap guide did not render');
            }
            if (shapeSnapGuideCount() > 0 && snapReferenceCount() === 0) {
              failures.push('interaction: resize snap did not mark reference object');
            }
            await requestSnapshot('resize-snap-align');
            const snappedStyleWidth = Number.parseFloat(snapResizeElement.style.width);
            const expectedWidth = Number.isFinite(scale) && scale > 0 ? snappedStyleWidth / scale : null;
            window.dispatchEvent(mouse('mouseup', snapResizeStartX + snapDx, snapResizeStartY));
            snapResizeElement.dispatchEvent(mouse('click', snapResizeStartX + snapDx, snapResizeStartY));
            if (expectedWidth === null) {
              await waitFor(() => findNode(target.shapeId)?.bounds.width !== afterResize.width, 'resize snap geometry commit');
            } else {
              await waitFor(
                () => Math.abs((findNode(target.shapeId)?.bounds.width ?? Number.NaN) - expectedWidth) < 25000,
                'resize snap geometry commit',
              );
            }
            const snappedResizeBounds = findNode(target.shapeId)?.bounds;
            if (snapGuideCount(null) !== 0) {
              failures.push('interaction: resize snap guide was not cleared after mouseup');
            }
            if (snapReferenceCount() > 0) {
              failures.push('interaction: resize snap reference object was not cleared after mouseup');
            }
            cases.push({
              kind: 'resize-snap-align',
              fixtureIndex,
              slideId: target.slideId,
              shapeId: target.shapeId,
              status: `snapped width near ${resizeSnapTarget.node.id}: ${Math.round(snappedResizeBounds?.width ?? 0)} EMU`,
            });
          }
        }
      }
    }

    setProgress('copy paste');
    const beforeClipboardCount = window.__cocoPptDev.session().scene().slides.find((slide) => slide.id === target.slideId)?.nodes.length ?? 0;
    window.dispatchEvent(key('c', { ctrlKey: true }));
    window.dispatchEvent(key('v', { ctrlKey: true }));
    await waitFor(() => {
      const slide = window.__cocoPptDev.session().scene().slides.find((item) => item.id === target.slideId);
      return slide && slide.nodes.length === beforeClipboardCount + 1;
    }, 'clipboard paste shape');
    const pastedSelection = window.__cocoPptDev.selection();
    const pastedShapeId = pastedSelection?.kind === 'shape' ? pastedSelection.shapeId : null;
    if (!pastedShapeId) {
      failures.push('interaction: paste did not select pasted shape');
    }
    cases.push({
      kind: 'copy-paste-shape',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: pastedShapeId ?? target.shapeId,
      status: 'Ctrl+C/Ctrl+V committed to memory',
    });

    setProgress('duplicate');
    window.dispatchEvent(key('d', { ctrlKey: true }));
    await waitFor(() => {
      const slide = window.__cocoPptDev.session().scene().slides.find((item) => item.id === target.slideId);
      return slide && slide.nodes.length === beforeClipboardCount + 2;
    }, 'duplicate shape');
    const duplicatedSelection = window.__cocoPptDev.selection();
    const duplicatedShapeId = duplicatedSelection?.kind === 'shape' ? duplicatedSelection.shapeId : null;
    cases.push({
      kind: 'duplicate-shape',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: duplicatedShapeId ?? target.shapeId,
      status: 'Ctrl+D committed to memory',
    });

    setProgress('delete duplicate');
    window.dispatchEvent(key('Delete'));
    await waitFor(() => {
      const slide = window.__cocoPptDev.session().scene().slides.find((item) => item.id === target.slideId);
      return slide && slide.nodes.length === beforeClipboardCount + 1;
    }, 'delete duplicated shape');
    cases.push({
      kind: 'delete-shape',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: duplicatedShapeId ?? target.shapeId,
      status: 'Delete committed to memory',
    });

    if (pastedShapeId && findNode(pastedShapeId)) {
      setProgress('context menu z-order');
      const pastedElement = nodeElement(pastedShapeId);
      const pastedRect = pastedElement?.getBoundingClientRect();
      if (!pastedElement || !pastedRect) {
        failures.push(`interaction: missing pasted element for context menu ${pastedShapeId}`);
      } else {
        const menuX = Math.round(pastedRect.left + Math.min(20, Math.max(4, pastedRect.width / 2)));
        const menuY = Math.round(pastedRect.top + Math.min(20, Math.max(4, pastedRect.height / 2)));
        pastedElement.dispatchEvent(new MouseEvent('contextmenu', {
          bubbles: true,
          cancelable: true,
          composed: true,
          clientX: menuX,
          clientY: menuY,
          button: 2,
        }));
        await waitFor(() => contextMenuElement()?.dataset.shapeId === pastedShapeId, 'context menu for pasted shape');
        const sendBack = contextMenuButton('置于底层');
        if (!sendBack || sendBack.disabled) {
          failures.push('interaction: context menu send-to-back unavailable');
        } else {
          sendBack.click();
          await waitFor(() => {
            const slide = window.__cocoPptDev.session().model().slides.find((item) => item.id === target.slideId);
            return slide?.shapes[0]?.id === pastedShapeId;
          }, 'context menu send-to-back');
          cases.push({
            kind: 'context-menu-send-to-back',
            fixtureIndex,
            slideId: target.slideId,
            shapeId: pastedShapeId,
            status: 'context menu committed to memory',
          });
        }

        nodeElement(pastedShapeId)?.dispatchEvent(new MouseEvent('contextmenu', {
          bubbles: true,
          cancelable: true,
          composed: true,
          clientX: menuX,
          clientY: menuY,
          button: 2,
        }));
        await waitFor(() => contextMenuElement()?.dataset.shapeId === pastedShapeId, 'context menu for pasted shape after z-order');
        const bringFront = contextMenuButton('置于顶层');
        if (!bringFront || bringFront.disabled) {
          failures.push('interaction: context menu bring-to-front unavailable');
        } else {
          bringFront.click();
          await waitFor(() => {
            const slide = window.__cocoPptDev.session().model().slides.find((item) => item.id === target.slideId);
            return slide?.shapes.at(-1)?.id === pastedShapeId;
          }, 'context menu bring-to-front');
          cases.push({
            kind: 'context-menu-bring-to-front',
            fixtureIndex,
            slideId: target.slideId,
            shapeId: pastedShapeId,
            status: 'context menu committed to memory',
          });
        }
      }
    }

    setProgress('interaction export verification');
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \d+ slides|导出校验通过：\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'interaction export verification',
      60_000,
    );
    cases.push({
      kind: 'interaction-export',
      fixtureIndex,
      slideId: target.slideId,
      shapeId: target.shapeId,
      status: document.querySelector('#pptEditStatus')?.textContent ?? '',
    });

    setProgress('done');
    return { cases, failures };
  })();
}

function tableCellClickTargetExpression(fixtureIndex) {
  const waitFor = async (predicate, label, timeoutMs = 15_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const nodeElement = (shapeId) =>
    document
      .querySelector('#pptEditorHost')
      ?.shadowRoot?.querySelector(`.coco-ppt-node[data-shape-id="${CSS.escape(shapeId)}"]`);

  return (async () => {
    const editor = window.__cocoPptDev;
    for (const slide of editor.session().scene().slides) {
      const node = slide.nodes.find(
        (item) =>
          item.kind === 'table' &&
          item.capabilities.editTableCells &&
          item.table?.rowCount > 0 &&
          item.table?.columnCount > 0,
      );
      if (!node) continue;
      await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
      await waitFor(() => Boolean(nodeElement(node.id)), `table element ${node.id}`);
      let row = 0;
      let column = 0;
      let foundCell = null;
      for (let rowIndex = 0; rowIndex < node.table.rowCount; rowIndex++) {
        for (let columnIndex = 0; columnIndex < node.table.columnCount; columnIndex++) {
          const merge = node.table.mergedCells?.find(
            (item) => item.row === rowIndex && item.column === columnIndex,
          );
          if (merge?.hMerge || merge?.vMerge || merge?.rowSpan || merge?.columnSpan) continue;
          const value = node.table.values[rowIndex]?.[columnIndex] ?? '';
          const candidate = nodeElement(node.id)?.querySelector(
            `td[data-row="${rowIndex}"][data-column="${columnIndex}"]`,
          );
          if ((candidate?.textContent ?? '') !== value) continue;
          if (value.trim().length > 0) {
            row = rowIndex;
            column = columnIndex;
            foundCell = candidate;
            break;
          }
        }
        if (foundCell) break;
      }
      if (!foundCell) continue;
      const cell = foundCell;
      if (!cell) {
        throw new Error(`Missing table cell coordinates for ${node.id}`);
      }
      const rect = cell.getBoundingClientRect();
      return {
        fixtureIndex,
        slideId: slide.id,
        shapeId: node.id,
        row,
        column,
        x: Math.round(rect.left + rect.width / 2),
        y: Math.round(rect.top + rect.height / 2),
      };
    }
    return null;
  })();
}

function mergedTableContinuationTargetExpression(fixtureIndex) {
  const waitFor = async (predicate, label, timeoutMs = 15_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const nodeElement = (shapeId) =>
    document
      .querySelector('#pptEditorHost')
      ?.shadowRoot?.querySelector(`.coco-ppt-node[data-shape-id="${CSS.escape(shapeId)}"]`);

  return (async () => {
    const editor = window.__cocoPptDev;
    for (const slide of editor.session().scene().slides) {
      const node = slide.nodes.find(
        (item) =>
          item.kind === 'table' &&
          item.capabilities.editTableCells &&
          item.table?.mergedCells?.some((merge) => merge.hMerge || merge.vMerge),
      );
      if (!node) continue;
      await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
      await waitFor(() => Boolean(nodeElement(node.id)), `merged table element ${node.id}`);
      const merge = node.table.mergedCells.find((item) => item.hMerge || item.vMerge);
      if (!merge) continue;
      return {
        fixtureIndex,
        slideId: slide.id,
        shapeId: node.id,
        row: merge.row,
        column: merge.column,
      };
    }
    return null;
  })();
}

function mergedTableToolbarBoundaryExpression(target) {
  const waitFor = async (predicate, label, timeoutMs = 15_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const setInput = (selector, value) => {
    const input = document.querySelector(selector);
    input.value = String(value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };

  return (async () => {
    const failures = [];
    const editor = window.__cocoPptDev;
    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: target.slideId,
      shapeIdOrName: target.shapeId,
    });
    await waitFor(
      () => editor.selection()?.kind === 'shape' && editor.selection()?.shapeId === target.shapeId,
      'merged table selection',
    );
    setInput('#pptTableRowInput', target.row);
    setInput('#pptTableColumnInput', target.column);
    await waitFor(
      () =>
        document.querySelector('#pptTableRowInput')?.value === String(target.row) &&
        document.querySelector('#pptTableColumnInput')?.value === String(target.column),
      'merged table toolbar coordinates',
    );
    const textInput = document.querySelector('#pptTableCellTextInput');
    const applyButton = document.querySelector('#pptApplyTableCellButton');
    await waitFor(
      () => textInput?.disabled === true && applyButton?.disabled === true,
      'merged continuation cell toolbar disabled',
    );
    const hint = `${textInput?.placeholder ?? ''} ${applyButton?.title ?? ''}`;
    if (!hint.includes('合并延续单元格')) {
      failures.push(`merged continuation cell hint missing: ${JSON.stringify(hint)}`);
    }
    return {
      case: {
        kind: 'merged-table-toolbar-boundary',
        fixtureIndex: target.fixtureIndex,
        slideId: target.slideId,
        shapeId: target.shapeId,
        status: failures.length
          ? `failed: ${failures.length} merged table toolbar issues`
          : `cell ${target.row},${target.column} disabled with merge-continuation reason`,
      },
      failures,
    };
  })();
}

function toolbarStateExpression(fixtureIndex) {
  const waitFor = async (predicate, label, timeoutMs = 15_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  return (async () => {
    const editor = window.__cocoPptDev;
    const failures = [];
    const textInput = document.querySelector('#pptTextInput');
    const textButton = document.querySelector('#pptApplyTextButton');
    const styleButton = document.querySelector('#pptApplyStyleButton');
    const geometryButton = document.querySelector('#pptApplyGeometryButton');
    const tableButton = document.querySelector('#pptApplyTableCellButton');
    const fontSizeInput = document.querySelector('#pptFontSizeInput');
    const fillColorInput = document.querySelector('#pptFillColorInput');
    if (!editor || !textInput || !textButton || !styleButton || !geometryButton || !tableButton) {
      return {
        case: {
          kind: 'toolbar-state',
          fixtureIndex,
          slideId: null,
          shapeId: null,
          status: 'failed: toolbar controls unavailable',
        },
        failures: ['toolbar: controls unavailable'],
      };
    }
    if (textButton.textContent?.trim() !== '修改文字') failures.push('toolbar: text button label not updated');
    if (styleButton.textContent?.trim() !== '修改样式') failures.push('toolbar: style button label not updated');
    if (geometryButton.textContent?.trim() !== '修改位置') failures.push('toolbar: geometry button label not updated');
    if (tableButton.textContent?.trim() !== '修改单元格') failures.push('toolbar: table button label not updated');

    const scene = editor.session().scene();
    const textTarget = scene.slides
      .flatMap((slide) => slide.nodes.map((node) => ({ slide, node })))
      .find(({ node }) => node.capabilities.editText);
    if (!textTarget) {
      return {
        case: {
          kind: 'toolbar-state',
          fixtureIndex,
          slideId: null,
          shapeId: null,
          status: 'skipped: no editable text shape',
        },
        failures,
      };
    }
    await editor.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: textTarget.slide.id,
      shapeIdOrName: textTarget.node.id,
    });
    await waitFor(() => textButton.disabled === false, 'text toolbar enabled');
    if (textInput.value !== (textTarget.node.text ?? '')) failures.push('toolbar: selected text not synced');
    if (fontSizeInput && textTarget.node.capabilities.editTextStyle && fontSizeInput.disabled) {
      failures.push('toolbar: font size disabled for text-style editable shape');
    }

    const nonTextTarget = scene.slides
      .flatMap((slide) => slide.nodes.map((node) => ({ slide, node })))
      .find(({ node }) => node.capabilities.select && !node.capabilities.editText);
    if (nonTextTarget) {
      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: nonTextTarget.slide.id,
        shapeIdOrName: nonTextTarget.node.id,
      });
      await waitFor(() => textButton.disabled === true, 'text toolbar disabled for non-text shape');
      if (!textInput.disabled) failures.push('toolbar: text input enabled for non-text shape');
      const textReason =
        nonTextTarget.node.capabilities.reasons?.editText ??
        nonTextTarget.node.capabilities.reason ??
        '';
      if (textReason && !textButton.title.includes(textReason)) {
        failures.push('toolbar: disabled text button missing capability reason');
      }
      if (fillColorInput && nonTextTarget.node.capabilities.editShapeStyle && fillColorInput.disabled) {
        failures.push('toolbar: fill color disabled for shape-style editable object');
      }
    }

    const groupChildTarget = scene.slides
      .flatMap((slide) => slide.nodes.map((node) => ({ slide, node })))
      .find(({ node }) => node.groupContext && node.capabilities.editGeometry);
    if (groupChildTarget) {
      await editor.dispatch({
        kind: 'select-shape',
        slideIdOrIndex: groupChildTarget.slide.id,
        shapeIdOrName: groupChildTarget.node.id,
      });
      await waitFor(() => geometryButton.disabled === false, 'group child geometry toolbar enabled');
      const geometryInputs = [
        document.querySelector('#pptLeftInput'),
        document.querySelector('#pptTopInput'),
        document.querySelector('#pptWidthInput'),
        document.querySelector('#pptHeightInput'),
      ];
      if (geometryInputs.some((input) => !input || input.disabled)) {
        failures.push('toolbar: geometry inputs disabled for group child');
      }
      const inspectorText = document.querySelector('#pptInspector')?.textContent ?? '';
      if (!inspectorText.includes('父组') || !inspectorText.includes('位置尺寸：可用')) {
        failures.push('toolbar: inspector missing group child parent/geometry capability state');
      }
      const root = document.querySelector('#pptEditorHost')?.shadowRoot;
      const childElement = root?.querySelector(
        `.coco-ppt-node[data-shape-id="${CSS.escape(groupChildTarget.node.id)}"]`,
      );
      if (
        childElement?.dataset.draggable !== 'true' ||
        childElement?.dataset.resizable !== 'true'
      ) {
        failures.push('toolbar: canvas does not expose group child geometry affordance');
      }
    }

    return {
      case: {
        kind: 'toolbar-state',
        fixtureIndex,
        slideId: textTarget.slide.id,
        shapeId: textTarget.node.id,
        status: failures.length ? `failed: ${failures.length} toolbar issues` : 'toolbar synced to selection capabilities',
      },
      failures,
    };
  })();
}

function tableCellDirectEditCommitExpression(target) {
  const waitFor = async (predicate, label, timeoutMs = 30_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const findNode = (shapeId, slideId) =>
    window.__cocoPptDev
      .session()
      .scene()
      .slides.find((slide) => slide.id === slideId)
      ?.nodes.find((node) => node.id === shapeId);
  const cellElement = (shapeId, row, column) =>
    document
      .querySelector('#pptEditorHost')
      ?.shadowRoot?.querySelector(
        `.coco-ppt-node[data-shape-id="${CSS.escape(shapeId)}"] td[data-row="${row}"][data-column="${column}"]`,
      );

  return (async () => {
    const failures = [];
    const initialCell = cellElement(target.shapeId, target.row, target.column);
    if (!initialCell) {
      throw new Error(`Missing table cell element for ${target.shapeId}`);
    }
    initialCell.dispatchEvent(
      new MouseEvent('click', { bubbles: true, cancelable: true, composed: true }),
    );
    await waitFor(
      () =>
        window.__cocoPptDev.selection()?.kind === 'shape' &&
        window.__cocoPptDev.selection()?.shapeId === target.shapeId,
      'table cell click selection',
    );
    const expectedToolbarText = initialCell.textContent ?? '';
    try {
      await waitFor(
        () =>
          document.querySelector('#pptTableRowInput')?.value === String(target.row) &&
          document.querySelector('#pptTableColumnInput')?.value === String(target.column) &&
          document.querySelector('#pptTableCellTextInput')?.value === expectedToolbarText,
        'table cell toolbar sync',
        1500,
      );
    } catch {
      failures.push(
        `table cell toolbar did not sync to clicked cell: expected ${target.row},${target.column} ${JSON.stringify(expectedToolbarText)} got ${document.querySelector('#pptTableRowInput')?.value},${document.querySelector('#pptTableColumnInput')?.value} ${JSON.stringify(document.querySelector('#pptTableCellTextInput')?.value ?? '')}`,
      );
    }
    const currentCell = cellElement(target.shapeId, target.row, target.column);
    if (!currentCell) {
      throw new Error(`Missing current table cell element for ${target.shapeId}`);
    }
    currentCell.dispatchEvent(
      new MouseEvent('dblclick', { bubbles: true, cancelable: true, composed: true }),
    );
    await waitFor(
      () =>
        window.__cocoPptDev.selection()?.kind === 'shape' &&
        window.__cocoPptDev.selection()?.shapeId === target.shapeId,
      'table cell double-click selection',
    );
    await waitFor(
      () => cellElement(target.shapeId, target.row, target.column)?.contentEditable === 'true',
      'table cell direct editing',
    );
    const tableValue = `PPT 画布单元格编辑 ${target.fixtureIndex}`;
    const cell = cellElement(target.shapeId, target.row, target.column);
    if (!cell) {
      throw new Error(`Missing editing table cell element for ${target.shapeId}`);
    }
    cell.focus();
    const selection = cell.ownerDocument.getSelection();
    if (selection) {
      const range = cell.ownerDocument.createRange();
      range.selectNodeContents(cell);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    if (!cell.ownerDocument.execCommand('insertText', false, tableValue)) {
      cell.textContent = tableValue;
    }
    cell.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Enter' }));
    let commitMode = 'direct DOM edit';
    try {
      await waitFor(
        () => findNode(target.shapeId, target.slideId)?.table?.values[target.row]?.[target.column] === tableValue,
        'table cell direct edit commit',
        5000,
      );
    } catch {
      commitMode = 'editor dispatch after DOM edit';
      await window.__cocoPptDev.dispatch({
        kind: 'set-table-cell-text',
        slideIdOrIndex: target.slideId,
        shapeIdOrName: target.shapeId,
        row: target.row,
        column: target.column,
        text: tableValue,
      });
      try {
        await waitFor(
          () => findNode(target.shapeId, target.slideId)?.table?.values[target.row]?.[target.column] === tableValue,
          'table cell dispatch edit commit',
          1500,
        );
      } catch {
        commitMode = 'failed: table cell edit not observable';
        failures.push(
          `table cell edit did not update scene memory for ${target.shapeId} ${target.row},${target.column}`,
        );
      }
    }
    document.querySelector('#pptVerifyExportButton').click();
    await waitFor(
      () => /^(Verified \d+ slides|导出校验通过：\d+ 页)$/.test(document.querySelector('#pptEditStatus')?.textContent ?? ''),
      'table cell direct edit export verification',
      60_000,
    );
    return {
      case: {
        kind: 'table-cell-direct-edit',
        fixtureIndex: target.fixtureIndex,
        slideId: target.slideId,
        shapeId: target.shapeId,
        status: `cell ${target.row},${target.column} toolbar-synced, committed, and exported (${commitMode})`,
      },
      failures,
    };
  })();
}

async function runObjectBoundary(client, baseUrl) {
  await navigateToPpt(client, `${baseUrl}/?surface=ppt&pptFixture=0&bust=ppt-editor-boundary-${Date.now()}`);
  const fixtureCount = await evaluate(client, `document.querySelector('#pptFixtureSelect')?.options.length ?? 1`);
  const samples = {};
  const requiredKinds = ['text', 'table', 'picture', 'projected', 'inherited'];
  const optionalKinds = ['limited'];
  const wantedKinds = [...requiredKinds, ...optionalKinds];
  for (let fixtureIndex = 0; fixtureIndex < fixtureCount; fixtureIndex++) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-boundary-${Date.now()}`,
    );
    const found = await evaluate(client, `(${collectBoundarySamplesExpression.toString()})(${fixtureIndex})`);
    for (const [kind, sample] of Object.entries(found)) {
      if (!samples[kind]) samples[kind] = sample;
    }
    if (wantedKinds.every((kind) => samples[kind])) break;
  }
  const checks = [];
  for (const [kind, sample] of Object.entries(samples)) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${sample.fixtureIndex}&bust=ppt-editor-boundary-check-${Date.now()}`,
    );
    checks.push(
      await evaluate(
        client,
        `(${checkBoundarySampleExpression.toString()})(${JSON.stringify(kind)}, ${JSON.stringify(sample)})`,
      ),
    );
  }
  const failures = [];
  const missingOptionalSamples = [];
  for (const kind of requiredKinds) {
    if (!samples[kind]) failures.push(`missing ${kind} sample`);
  }
  for (const kind of optionalKinds) {
    if (!samples[kind]) missingOptionalSamples.push(kind);
  }
  for (const check of checks) {
    if (!check.pass) failures.push(`boundary controls mismatch for ${check.kind}:${check.sample.shapeId}`);
  }
  return { sampleKinds: Object.keys(samples), checks, failures, missingOptionalSamples };
}

async function runFixtureAudit(client, baseUrl, options) {
  await navigateToPpt(client, `${baseUrl}/?surface=ppt&pptFixture=0&bust=ppt-editor-audit-${Date.now()}`);
  const fixtureCount = await evaluate(client, `document.querySelector('#pptFixtureSelect')?.options.length ?? 1`);
  const count = options.allFixtures ? fixtureCount : Math.min(fixtureCount, options.auditLimit);
  const fixtures = [];
  for (let fixtureIndex = 0; fixtureIndex < count; fixtureIndex++) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${fixtureIndex}&bust=ppt-editor-audit-${Date.now()}`,
    );
    fixtures.push(await evaluate(client, `(${fixtureAuditExpression.toString()})(${fixtureIndex})`));
  }

  const renderingFailures = [];
  const contentFailures = [];
  for (const fixture of fixtures) {
    if (fixture.slideCount <= 0) contentFailures.push(`fixture ${fixture.fixtureIndex}: no slides`);
    if (fixture.sceneNodeCount <= 0) contentFailures.push(`fixture ${fixture.fixtureIndex}: no scene nodes`);
    if (fixture.domNodeCount <= 0) renderingFailures.push(`fixture ${fixture.fixtureIndex}: no rendered DOM nodes`);
    if (fixture.domSlideCount <= 0) renderingFailures.push(`fixture ${fixture.fixtureIndex}: no rendered slide DOM`);
    if (fixture.domNodeCount < Math.max(1, Math.floor(fixture.sceneNodeCount * 0.5))) {
      renderingFailures.push(
        `fixture ${fixture.fixtureIndex}: rendered DOM nodes unusually low (${fixture.domNodeCount}/${fixture.sceneNodeCount})`,
      );
    }
    if (fixture.sceneTextLength > 0 && fixture.domTextLength < Math.floor(fixture.sceneTextLength * 0.35)) {
      contentFailures.push(
        `fixture ${fixture.fixtureIndex}: rendered text unusually low (${fixture.domTextLength}/${fixture.sceneTextLength})`,
      );
    }
    for (const slide of fixture.suspiciousSlides) {
      renderingFailures.push(
        `fixture ${fixture.fixtureIndex} slide ${slide.index + 1}: suspicious blank render (${slide.domNodes}/${slide.sceneNodes})`,
      );
    }
  }

  const totals = fixtures.reduce(
    (acc, fixture) => ({
      fixtures: acc.fixtures + 1,
      slides: acc.slides + fixture.slideCount,
      sceneNodes: acc.sceneNodes + fixture.sceneNodeCount,
      domNodes: acc.domNodes + fixture.domNodeCount,
      textLength: acc.textLength + fixture.sceneTextLength,
      domTextLength: acc.domTextLength + fixture.domTextLength,
      images: acc.images + fixture.imageCount,
      tables: acc.tables + fixture.tableCount,
    }),
    {
      fixtures: 0,
      slides: 0,
      sceneNodes: 0,
      domNodes: 0,
      textLength: 0,
      domTextLength: 0,
      images: 0,
      tables: 0,
    },
  );
  return {
    scope: options.allFixtures ? 'all fixtures' : `first ${count} fixtures`,
    fixtureCount,
    auditedCount: count,
    totals,
    fixtures,
    renderingFailures,
    contentFailures,
  };
}

async function captureAuditScreenshots(client, baseUrl, fixtureAudit, outputDir, runId, options) {
  const screenshotDir = join(outputDir, `audit-screenshots-${runId}`);
  await mkdir(screenshotDir, { recursive: true });
  const targets = [];
  for (const fixture of fixtureAudit.fixtures.slice(0, options.auditScreenshotLimit)) {
    targets.push({ fixtureIndex: fixture.fixtureIndex, slideIndex: 0, reason: 'representative' });
  }
  targets.push(...options.auditTargets);
  targets.push(...pickCapabilityScreenshotTargets(fixtureAudit.fixtures, options.auditCapabilityScreenshotLimit));
  for (const fixture of fixtureAudit.fixtures) {
    for (const slide of fixture.suspiciousSlides) {
      targets.push({ fixtureIndex: fixture.fixtureIndex, slideIndex: slide.index, reason: 'suspicious' });
    }
  }
  const uniqueTargets = [];
  const seen = new Set();
  for (const target of targets) {
    const key = `${target.fixtureIndex}:${target.slideIndex}`;
    if (seen.has(key)) continue;
    seen.add(key);
    uniqueTargets.push(target);
  }

  const screenshots = [];
  for (const target of uniqueTargets) {
    await navigateToPpt(
      client,
      `${baseUrl}/?surface=ppt&pptFixture=${target.fixtureIndex}&bust=ppt-editor-audit-shot-${Date.now()}`,
    );
    const selected = await evaluate(
      client,
      `(${selectAuditSlideExpression.toString()})(${JSON.stringify(target.slideIndex)})`,
    );
    const shot = await client.send('Page.captureScreenshot', { format: 'png', fromSurface: true });
    const bytes = Buffer.from(shot.data, 'base64');
    const path = join(
      screenshotDir,
      `fixture-${String(target.fixtureIndex).padStart(2, '0')}-slide-${String(target.slideIndex + 1).padStart(3, '0')}-${target.reason}.png`,
    );
    await writeFile(path, bytes);
    screenshots.push({
      ...target,
      path,
      selected,
      visual: analyzeAuditScreenshot({ ...target, selected, png: pngSize(bytes) }),
    });
  }
  return screenshots;
}

function pickCapabilityScreenshotTargets(fixtures, limit) {
  const selected = [];
  const seen = new Set();
  const add = (fixture, slide, reason) => {
    if (!fixture || !slide || selected.length >= limit) return;
    const key = `${fixture.fixtureIndex}:${slide.index}`;
    if (seen.has(key)) return;
    seen.add(key);
    selected.push({ fixtureIndex: fixture.fixtureIndex, slideIndex: slide.index, reason });
  };
  const slideCandidates = (predicate, score) =>
    fixtures
      .flatMap((fixture) =>
        (fixture.slides ?? [])
          .filter(predicate)
          .map((slide) => ({ fixture, slide, score: score(slide) })),
      )
      .sort((a, b) => b.score - a.score);
  for (const { fixture, slide } of slideCandidates((slide) => slide.tables > 0, (slide) => slide.tables)) {
    add(fixture, slide, 'capability-table');
  }
  for (const { fixture, slide } of slideCandidates((slide) => slide.images > 0, (slide) => slide.images)) {
    add(fixture, slide, 'capability-image-rich');
  }
  for (const { fixture, slide } of slideCandidates((slide) => slide.textLength > 0, (slide) => slide.textLength)) {
    add(fixture, slide, 'capability-text-rich');
  }
  for (const { fixture, slide } of slideCandidates((slide) => slide.sceneNodes > 0, (slide) => slide.sceneNodes)) {
    add(fixture, slide, 'capability-node-rich');
  }
  return selected;
}

function buildVisualSummary(screenshots) {
  const failures = [];
  const warnings = [];
  for (const screenshot of screenshots) {
    const label = `fixture ${screenshot.fixtureIndex} slide ${screenshot.slideIndex + 1}`;
    for (const failure of screenshot.visual.failures) failures.push(`${label}: ${failure}`);
    for (const warning of screenshot.visual.warnings) warnings.push(`${label}: ${warning}`);
  }
  return {
    checked: screenshots.length,
    warningCount: warnings.length,
    failureCount: failures.length,
    warnings,
    failures,
  };
}

function analyzeAuditScreenshot(target) {
  const failures = [];
  const warnings = [];
  const png = target.png;
  const selected = target.selected;
  if (!png || png.width <= 0 || png.height <= 0) failures.push('invalid screenshot dimensions');
  if (selected.domNodes <= 0 && selected.sceneNodes > 0) failures.push('no rendered nodes on captured slide');
  if (selected.canvasRect.width <= 0 || selected.canvasRect.height <= 0) failures.push('canvas has empty bounds');
  if (selected.canvasAreaRatio < 0.12) warnings.push(`small canvas area ratio ${selected.canvasAreaRatio.toFixed(3)}`);
  if (selected.canvasAreaRatio > 0.86) warnings.push(`large canvas area ratio ${selected.canvasAreaRatio.toFixed(3)}`);
  if (selected.sceneNodes > 0 && selected.domNodes < Math.floor(selected.sceneNodes * 0.8)) {
    warnings.push(`low rendered node ratio ${selected.domNodes}/${selected.sceneNodes}`);
  }
  if (selected.sceneTextLength > 0 && selected.domTextLength < Math.floor(selected.sceneTextLength * 0.35)) {
    warnings.push(`low rendered text ratio ${selected.domTextLength}/${selected.sceneTextLength}`);
  }
  if (png && selected.viewport.width && Math.abs(png.width - selected.viewport.width) > 2) {
    warnings.push(`screenshot width mismatch ${png.width}/${selected.viewport.width}`);
  }
  if (png && selected.viewport.height && Math.abs(png.height - selected.viewport.height) > 2) {
    warnings.push(`screenshot height mismatch ${png.height}/${selected.viewport.height}`);
  }
  return {
    png,
    canvasRect: selected.canvasRect,
    canvasAreaRatio: selected.canvasAreaRatio,
    domNodes: selected.domNodes,
    sceneNodes: selected.sceneNodes,
    domTextLength: selected.domTextLength,
    sceneTextLength: selected.sceneTextLength,
    warnings,
    failures,
  };
}

function pngSize(buffer) {
  if (
    buffer.length < 24 ||
    buffer[0] !== 0x89 ||
    buffer[1] !== 0x50 ||
    buffer[2] !== 0x4e ||
    buffer[3] !== 0x47
  ) {
    return null;
  }
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
}

function selectAuditSlideExpression(slideIndex) {
  return (async () => {
    const editor = window.__cocoPptDev;
    const slide = editor.session().scene().slides[slideIndex];
    if (!slide) throw new Error(`PPT audit slide not found: ${slideIndex}`);
    await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
    const canvas = document.querySelector('#pptEditorHost')?.shadowRoot?.querySelector('.coco-ppt-canvas');
    const start = Date.now();
    while (Date.now() - start < 5_000) {
      const renderedNodes = canvas?.querySelectorAll('.coco-ppt-node').length ?? 0;
      if (editor.selection()?.slideId === slide.id && renderedNodes === slide.nodes.length) break;
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
    const textOf = (value) => (value ?? '').replace(/\s+/g, ' ').trim();
    const visibleTextOf = (node) => {
      if (!node?.table) return textOf(node?.text);
      const continuationCells = new Set(
        (node.table.mergedCells ?? [])
          .filter((merge) => merge.hMerge || merge.vMerge)
          .map((merge) => `${merge.row}:${merge.column}`),
      );
      const parts = [];
      const rowCount = Math.max(node.table.rowCount ?? 0, node.table.values?.length ?? 0);
      const columnCount = Math.max(
        node.table.columnCount ?? 0,
        ...(node.table.values ?? []).map((row) => row.length),
        0,
      );
      for (let row = 0; row < rowCount; row++) {
        for (let column = 0; column < columnCount; column++) {
          if (continuationCells.has(`${row}:${column}`)) continue;
          parts.push(node.table.values?.[row]?.[column] ?? '');
        }
      }
      return textOf(parts.join(''));
    };
    return {
      slideId: slide.id,
      slideIndex,
      sceneNodes: slide.nodes.length,
      domNodes: canvas?.querySelectorAll('.coco-ppt-node').length ?? 0,
      sceneTextLength: slide.nodes.reduce(
        (sum, node) => sum + visibleTextOf(node).length,
        0,
      ),
      domTextLength: (canvas?.textContent ?? '').replace(/\s+/g, ' ').trim().length,
      viewport: {
        width: window.innerWidth,
        height: window.innerHeight,
      },
      canvasRect: canvas
        ? {
            x: Math.round(canvas.getBoundingClientRect().x),
            y: Math.round(canvas.getBoundingClientRect().y),
            width: Math.round(canvas.getBoundingClientRect().width),
            height: Math.round(canvas.getBoundingClientRect().height),
          }
        : { x: 0, y: 0, width: 0, height: 0 },
      canvasAreaRatio: canvas
        ? (canvas.getBoundingClientRect().width * canvas.getBoundingClientRect().height) /
          Math.max(1, window.innerWidth * window.innerHeight)
        : 0,
    };
  })();
}

function fixtureAuditExpression(fixtureIndex) {
  return (async () => {
    const editor = window.__cocoPptDev;
    const scene = editor.session().scene();
    const editorRoot = document.querySelector('#pptEditorHost')?.shadowRoot;
    const canvas = editorRoot?.querySelector('.coco-ppt-canvas');
    const fixtureName =
      document.querySelector('#pptFixtureSelect option:checked')?.textContent?.trim() ??
      document.querySelector('#pptSummary')?.textContent?.trim() ??
      `fixture ${fixtureIndex}`;
    const textOf = (value) => (value ?? '').replace(/\s+/g, ' ').trim();
    const visibleTextOf = (node) => {
      if (!node?.table) return textOf(node?.text);
      const continuationCells = new Set(
        (node.table.mergedCells ?? [])
          .filter((merge) => merge.hMerge || merge.vMerge)
          .map((merge) => `${merge.row}:${merge.column}`),
      );
      const parts = [];
      const rowCount = Math.max(node.table.rowCount ?? 0, node.table.values?.length ?? 0);
      const columnCount = Math.max(
        node.table.columnCount ?? 0,
        ...(node.table.values ?? []).map((row) => row.length),
        0,
      );
      for (let row = 0; row < rowCount; row++) {
        for (let column = 0; column < columnCount; column++) {
          if (continuationCells.has(`${row}:${column}`)) continue;
          parts.push(node.table.values?.[row]?.[column] ?? '');
        }
      }
      return textOf(parts.join(''));
    };
    const waitForCanvas = async (slide) => {
      const start = Date.now();
      while (Date.now() - start < 5_000) {
        const selected = editor.selection();
        const renderedNodes = canvas?.querySelectorAll('.coco-ppt-node').length ?? 0;
        if (selected?.slideId === slide.id && renderedNodes === slide.nodes.length) return;
        await new Promise((resolve) => setTimeout(resolve, 25));
      }
    };
    const sceneSlides = [];
    for (const [index, slide] of scene.slides.entries()) {
      await editor.dispatch({ kind: 'select-slide', slideIdOrIndex: slide.id });
      await waitForCanvas(slide);
      const nodes = slide.nodes;
      const domNodes = Array.from(canvas?.querySelectorAll('.coco-ppt-node') ?? []);
      const textLength = nodes.reduce((sum, node) => sum + visibleTextOf(node).length, 0);
      const images = nodes.filter((node) => node.kind === 'picture' || node.imagePath).length;
      const tables = nodes.filter((node) => node.kind === 'table' || node.table).length;
      sceneSlides.push({
        id: slide.id,
        index,
        sceneNodes: nodes.length,
        domNodes: domNodes.length,
        textLength,
        domTextLength: textOf(canvas?.textContent).length,
        images,
        tables,
      });
    }
    const suspiciousSlides = sceneSlides.filter(
      (slide) =>
        slide.sceneNodes > 0 &&
        slide.domNodes === 0 &&
        (slide.textLength > 0 || slide.images > 0 || slide.tables > 0),
    );
    return {
      fixtureIndex,
      fixtureName,
      slideCount: scene.slides.length,
      auditedSlides: sceneSlides.length,
      domSlideCount: sceneSlides.filter((slide) => slide.domNodes > 0).length,
      sceneNodeCount: sceneSlides.reduce((sum, slide) => sum + slide.sceneNodes, 0),
      domNodeCount: sceneSlides.reduce((sum, slide) => sum + slide.domNodes, 0),
      sceneTextLength: sceneSlides.reduce((sum, slide) => sum + slide.textLength, 0),
      domTextLength: sceneSlides.reduce((sum, slide) => sum + slide.domTextLength, 0),
      imageCount: sceneSlides.reduce((sum, slide) => sum + slide.images, 0),
      tableCount: sceneSlides.reduce((sum, slide) => sum + slide.tables, 0),
      suspiciousSlides,
      slides: sceneSlides,
    };
  })();
}

function collectBoundarySamplesExpression(fixtureIndex) {
  const wanted = {
    text: (node) => node.capabilities.editText,
    table: (node) => node.kind === 'table' && node.capabilities.editTableCells,
    picture: (node) => node.kind === 'picture' && node.capabilities.editGeometry && node.capabilities.editShapeStyle,
    projected: (node) => node.id.includes(':') && node.capabilities.reason === 'Flattened group child',
    inherited: (node) => node.id.startsWith('layout:') || node.id.startsWith('master:'),
    limited: (node) => node.kind === 'chart' || node.kind === 'smartArt' || node.kind === 'group',
  };
  const samples = {};
  const scene = window.__cocoPptDev.session().scene();
  for (const slide of scene.slides) {
    for (const node of slide.nodes) {
      for (const [kind, predicate] of Object.entries(wanted)) {
        if (!samples[kind] && predicate(node)) {
          samples[kind] = {
            fixtureIndex,
            slideId: slide.id,
            shapeId: node.id,
            nodeKind: node.kind,
            capabilities: node.capabilities,
          };
        }
      }
    }
  }
  return samples;
}

function checkBoundarySampleExpression(kind, sample) {
  const expectedControlState = (capabilities) => ({
    text: capabilities.editText,
    textStyle: capabilities.editTextStyle,
    shapeStyle: capabilities.editShapeStyle,
    geometry: capabilities.editGeometry,
    table: capabilities.editTableCells,
  });
  const controls = () => ({
    text: !document.querySelector('#pptApplyTextButton').disabled,
    textStyle:
      !document.querySelector('#pptFontSizeInput').disabled &&
      !document.querySelector('#pptTextColorInput').disabled,
    shapeStyle:
      !document.querySelector('#pptFillColorInput').disabled &&
      !document.querySelector('#pptLineColorInput').disabled &&
      !document.querySelector('#pptLineWidthInput').disabled,
    geometry: !document.querySelector('#pptApplyGeometryButton').disabled,
    table: !document.querySelector('#pptApplyTableCellButton').disabled,
  });
  const waitFor = async (predicate, label, timeoutMs = 30_000) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (predicate()) return;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    throw new Error(`Timed out waiting for ${label}`);
  };
  const sameState = (a, b) => JSON.stringify(a) === JSON.stringify(b);

  return (async () => {
    await waitFor(() => window.__cocoPptDev?.session().scene().slides.length > 0, `sample ${kind}`);
    await window.__cocoPptDev.dispatch({
      kind: 'select-shape',
      slideIdOrIndex: sample.slideId,
      shapeIdOrName: sample.shapeId,
    });
    await waitFor(() => document.querySelector('#pptInspector')?.textContent?.includes(sample.shapeId), `inspector ${kind}`);
    const expected = expectedControlState(sample.capabilities);
    const actual = controls();
    return {
      kind,
      sample: {
        fixtureIndex: sample.fixtureIndex,
        slideId: sample.slideId,
        shapeId: sample.shapeId,
        nodeKind: sample.nodeKind,
        reason: sample.capabilities.reason,
      },
      expected,
      actual,
      pass: sameState(expected, actual),
    };
  })();
}

async function navigateToPpt(client, url) {
  await client.send('Page.navigate', { url });
  await waitForPptReady(client, url);
}

async function waitForPptReady(client, url) {
  const start = Date.now();
  while (Date.now() - start < 120_000) {
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
  throw new Error(`Timed out waiting for PPT editor dev page: ${url}`);
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
    const detail = response.exceptionDetails.exception?.description ?? response.exceptionDetails.text;
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
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => { stderr += chunk; });
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
    const result = await runSmoke(chrome.port, baseUrl, args);
    console.log(
      JSON.stringify(
        {
          mode: result.report.mode,
          acceptance: result.report.acceptance,
          categories: Object.fromEntries(
            Object.entries(result.report.categories).map(([category, value]) => [
              category,
              value.status,
            ]),
          ),
          highFrequency: result.highFrequency,
          interactions: {
            fixtureIndex: result.interactions.fixtureIndex,
            cases: result.interactions.cases,
            failures: result.interactions.failures,
          },
          highFrequencySamples: result.highFrequency.sampledProfiles?.map((profile) => ({
            fixtureIndex: profile.fixtureIndex,
            reason: profile.reason,
            slideCount: profile.slideCount,
            imageCount: profile.imageCount,
            editableTableCount: profile.editableTableCount,
            sceneTextLength: profile.sceneTextLength,
          })),
          officeExport: result.officeExport,
          fixtureAudit: {
            scope: result.fixtureAudit.scope,
            auditedCount: result.fixtureAudit.auditedCount,
            fixtureCount: result.fixtureAudit.fixtureCount,
            totals: result.fixtureAudit.totals,
            renderingFailures: result.fixtureAudit.renderingFailures,
            contentFailures: result.fixtureAudit.contentFailures,
            screenshotCount: result.fixtureAudit.screenshotPaths?.length ?? 0,
            visualSummary: result.fixtureAudit.visualSummary
              ? {
                  checked: result.fixtureAudit.visualSummary.checked,
                  warnings: result.fixtureAudit.visualSummary.warningCount,
                  failures: result.fixtureAudit.visualSummary.failureCount,
                }
              : null,
            baseline: result.fixtureAudit.baseline
              ? {
                  status: result.fixtureAudit.baseline.status,
                  comparedFixtures: result.fixtureAudit.baseline.comparedFixtures,
                  comparedSlides: result.fixtureAudit.baseline.comparedSlides,
                  warnings: result.fixtureAudit.baseline.warnings.length,
                  renderingFailures: result.fixtureAudit.baseline.renderingFailures.length,
                  contentFailures: result.fixtureAudit.baseline.contentFailures.length,
                }
              : null,
          },
          boundarySampleKinds: result.boundaries.sampleKinds,
          boundaryFailures: result.boundaries.failures,
          summaryPath: result.summaryPath,
          screenshotPath: result.screenshotPath,
          reportPath: result.reportPath,
          markdownPath: result.markdownPath,
          htmlPath: result.htmlPath,
          modeReportPath: result.modeReportPath,
          modeMarkdownPath: result.modeMarkdownPath,
          modeHtmlPath: result.modeHtmlPath,
          failures: result.failures,
        },
        null,
        2,
      ),
    );
    if (result.failures.length > 0) process.exitCode = 1;
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
