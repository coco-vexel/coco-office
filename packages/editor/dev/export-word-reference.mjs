import { spawn } from 'node:child_process';
import { access, mkdir } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const defaultOutputDir = resolve(packageDir, 'dev/native-regression-output/word-reference');

const options = parseArgs(process.argv.slice(2));
await mkdir(options.outputDir, { recursive: true });

const fixtureName = basename(options.input);
const fixtureDir = join(options.outputDir, sanitizePathSegment(fixtureName));
await mkdir(fixtureDir, { recursive: true });

const pdfPath = join(fixtureDir, 'reference.pdf');
await exportDocxToPdfWithWord(options.input, pdfPath);
await renderPdfPages(pdfPath, join(fixtureDir, 'page'), options.pages);

console.log(`Word reference PDF: ${pdfPath}`);
console.log(`Word reference PNGs: ${fixtureDir}`);

function parseArgs(argv) {
  const parsed = {
    input: undefined,
    outputDir: defaultOutputDir,
    pages: 3,
  };
  for (const arg of argv) {
    if (arg === '--') {
      continue;
    } else if (arg.startsWith('--input=')) {
      parsed.input = resolve(arg.slice('--input='.length));
    } else if (arg.startsWith('--output-dir=')) {
      parsed.outputDir = resolve(packageDir, arg.slice('--output-dir='.length));
    } else if (arg.startsWith('--pages=')) {
      parsed.pages = Number(arg.slice('--pages='.length));
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  if (!parsed.input) throw new Error('Missing --input=<docx path>.');
  if (!Number.isInteger(parsed.pages) || parsed.pages <= 0) throw new Error('Invalid --pages value.');
  return parsed;
}

async function exportDocxToPdfWithWord(inputPath, outputPath) {
  const script = `
$ErrorActionPreference = 'Stop'
$inputPath = ${psString(inputPath)}
$outputPath = ${psString(outputPath)}
$word = New-Object -ComObject Word.Application
$initialDocumentCount = $word.Documents.Count
$word.Visible = $false
$word.DisplayAlerts = 0
$document = $null
try {
  try {
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
  } catch {
    $document = $word.Documents.Open($inputPath, $false, $true, $false, '', '', $false, '', '', 0, '', $false, $true)
  }
  $document.ExportAsFixedFormat($outputPath, 17)
} finally {
  if ($document -ne $null) {
    $document.Close($false)
  }
  if ($initialDocumentCount -eq 0) {
    $word.Quit()
  }
}
`;
  await run('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script]);
}

async function renderPdfPages(pdfPath, outputPrefix, pages) {
  const pdftoppm = await findPdftoppm();
  await run(pdftoppm, [
    '-png',
    '-f',
    '1',
    '-l',
    String(pages),
    '-r',
    '120',
    pdfPath,
    outputPrefix,
  ]);
}

async function findPdftoppm() {
  const candidates = [
    process.env.PDFTOPPM_PATH,
    process.env.USERPROFILE
      ? join(
          process.env.USERPROFILE,
          '.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe',
        )
      : undefined,
    'pdftoppm.cmd',
  ].filter(Boolean);
  for (const candidate of candidates) {
    if (candidate === 'pdftoppm.cmd') return candidate;
    try {
      await access(candidate);
      return candidate;
    } catch {
      // Try the next candidate.
    }
  }
  return 'pdftoppm.cmd';
}

function psString(value) {
  return `'${value.replace(/'/g, "''")}'`;
}

function sanitizePathSegment(value) {
  return value.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 80) || 'fixture';
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: packageDir,
      stdio: ['ignore', 'pipe', 'pipe'],
      ...options,
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });
    child.on('error', reject);
    child.on('exit', (code) => {
      if (code === 0) {
        resolve(stdout);
      } else {
        reject(new Error(`${command} exited with ${code}\n${stderr || stdout}`));
      }
    });
  });
}
