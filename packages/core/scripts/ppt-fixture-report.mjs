#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { summarizePptFixture } from './ppt-fixture-summary.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const defaultFixtureDir = process.env.COCO_PPTX_FIXTURE_DIR || '';
const args = process.argv.slice(2).filter((arg) => arg !== '--');
const json = args.includes('--json');
const fixtureDir = args.find((arg) => !arg.startsWith('--')) ?? defaultFixtureDir;

async function main() {
  const distPath = resolve(here, '../dist/ppt.js');
  if (!existsSync(distPath)) {
    throw new Error(
      `Missing ${distPath}. Run "corepack pnpm --filter @coco-vexel/coco-office build" first.`,
    );
  }
  if (!existsSync(fixtureDir)) {
    throw new Error(`PPT fixture directory not found: ${fixtureDir}`);
  }

  const { Presentation } = await import(pathToFileURL(distPath).href);
  const files = readdirSync(fixtureDir)
    .filter((name) => name.toLowerCase().endsWith('.pptx') && !name.startsWith('~$'))
    .sort((left, right) => left.localeCompare(right));

  const summaries = [];
  for (const name of files) {
    try {
      summaries.push(
        await summarizePptFixture(
          Presentation,
          name,
          readFileSync(join(fixtureDir, name)),
          'COCO_REPORT',
        ),
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`Failed to summarize PPT fixture "${name}": ${message}`);
    }
  }

  if (json) {
    console.log(
      JSON.stringify({ fixtureDir, generatedAt: new Date().toISOString(), summaries }, null, 2),
    );
    return;
  }

  printMarkdownReport(fixtureDir, summaries);
}

function printMarkdownReport(fixtureDir, summaries) {
  const totals = summaries.reduce(
    (next, summary) => ({
      slides: next.slides + summary.slideCount,
      shapes: next.shapes + summary.shapeCount,
      charts: next.charts + summary.packageParts.chartParts,
      media: next.media + summary.packageParts.mediaParts,
      imageMedia: next.imageMedia + summary.packageParts.imageParts,
      videoMedia: next.videoMedia + summary.packageParts.videoParts,
      audioMedia: next.audioMedia + summary.packageParts.audioParts,
      otherMedia: next.otherMedia + summary.packageParts.otherMediaParts,
      embeddings: next.embeddings + summary.packageParts.embeddingParts,
      diagrams: next.diagrams + summary.packageParts.diagramParts,
      chartShapes: next.chartShapes + summary.advancedShapes.chartShapes,
      mediaShapes: next.mediaShapes + summary.advancedShapes.mediaShapes,
      oleObjectShapes: next.oleObjectShapes + summary.advancedShapes.oleObjectShapes,
      smartArtShapes: next.smartArtShapes + summary.advancedShapes.smartArtShapes,
      unknownShapes: next.unknownShapes + summary.advancedShapes.unknownShapes,
    }),
    {
      slides: 0,
      shapes: 0,
      charts: 0,
      media: 0,
      imageMedia: 0,
      videoMedia: 0,
      audioMedia: 0,
      otherMedia: 0,
      embeddings: 0,
      diagrams: 0,
      chartShapes: 0,
      mediaShapes: 0,
      oleObjectShapes: 0,
      smartArtShapes: 0,
      unknownShapes: 0,
    },
  );

  console.log(`# PPT Fixture Report`);
  console.log();
  console.log(`Fixture directory: \`${fixtureDir}\``);
  console.log();
  console.log(
    `Totals: ${summaries.length} files, ${totals.slides} slides, ${totals.shapes} shapes, ${totals.charts} chart parts, ${totals.media} media parts, ${totals.embeddings} embedding parts, ${totals.diagrams} diagram parts.`,
  );
  console.log();
  console.log(
    '| File | Slides | Shapes | Shape kinds | Advanced parts | Motion | Editable | Preservation |',
  );
  console.log('| --- | ---: | ---: | --- | --- | --- | --- | --- |');
  for (const summary of summaries) {
    console.log(
      `| ${escapeCell(summary.name)} | ${summary.slideCount} | ${summary.shapeCount} | ${escapeCell(formatCounts(summary.shapeKinds))} | ${escapeCell(formatAdvancedParts(summary.packageParts))} | ${escapeCell(formatMotion(summary.motion))} | ${escapeCell(formatEditable(summary.editableCandidates))} | ${escapeCell(formatPreservation(summary))} |`,
    );
  }
  console.log();
  printAdvancedCoverage(summaries, totals);
}

function formatCounts(counts) {
  const entries = Object.entries(counts).sort(([left], [right]) => left.localeCompare(right));
  return entries.length ? entries.map(([key, value]) => `${key}:${value}`).join(', ') : '-';
}

function formatAdvancedParts(parts) {
  return `charts:${parts.chartParts}, media:${parts.mediaParts} (${formatMediaBreakdown(parts)}), embeddings:${parts.embeddingParts}, diagrams:${parts.diagramParts}`;
}

function formatMediaBreakdown(parts) {
  return `images:${parts.imageParts}, video:${parts.videoParts}, audio:${parts.audioParts}, other:${parts.otherMediaParts}`;
}

function formatMotion(motion) {
  return `transitions:${motion.transitions}, timing:${motion.timing}`;
}

function formatEditable(editable) {
  return (
    [editable.text ? 'text' : '', editable.table ? 'table' : '', editable.image ? 'image' : '']
      .filter(Boolean)
      .join(', ') || '-'
  );
}

function formatPreservation(summary) {
  return [
    summary.noEditRoundTripPreserved ? 'roundtrip:ok' : 'roundtrip:changed',
    summary.advancedPartsPreservedAfterSdkWrites ? 'advanced:ok' : 'advanced:changed',
    formatAdvancedSlidePreservation(summary.advancedShapesPreservedOnEditedSlide),
  ].join(', ');
}

function printAdvancedCoverage(summaries, totals) {
  const advancedSlideFiles = summaries
    .filter((summary) => summary.advancedShapesPreservedOnEditedSlide === true)
    .map((summary) => summary.name);
  const missingShapeFixtures = [
    totals.chartShapes > 0 ? '' : 'chart',
    totals.mediaShapes > 0 ? '' : 'mediaObject',
    totals.oleObjectShapes > 0 ? '' : 'oleObject',
    totals.smartArtShapes > 0 ? '' : 'smartArt',
    totals.unknownShapes > 0 ? '' : 'unknown',
  ].filter(Boolean);

  console.log(`## Advanced Coverage`);
  console.log();
  console.log(
    `Advanced shape models: chart:${totals.chartShapes}, media:${totals.mediaShapes}, oleObject:${totals.oleObjectShapes}, smartArt:${totals.smartArtShapes}, unknown:${totals.unknownShapes}.`,
  );
  console.log(
    `Advanced package parts: charts:${totals.charts}, media:${totals.media}, embeddings:${totals.embeddings}, diagrams:${totals.diagrams}.`,
  );
  console.log(
    `Media package breakdown: images:${totals.imageMedia}, video:${totals.videoMedia}, audio:${totals.audioMedia}, other:${totals.otherMedia}.`,
  );
  console.log(
    `Edited-slide preservation exercised: ${advancedSlideFiles.length ? advancedSlideFiles.join(', ') : 'none'}.`,
  );
  console.log(
    `Fixture gaps: ${missingShapeFixtures.length ? missingShapeFixtures.join(', ') : 'none'}.`,
  );
}

function formatAdvancedSlidePreservation(value) {
  if (value === null) return 'advanced-slide:n/a';
  return value ? 'advanced-slide:ok' : 'advanced-slide:changed';
}

function escapeCell(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
});
