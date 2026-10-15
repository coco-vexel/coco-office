import { fromZip, toZip } from '@coco-vexel/coco-office/pack';
import { buildIndices } from '@coco-vexel/coco-office/word-view';
import {
  buildWordEditorViewFromPackage,
  renderNativeWordEditor,
  applyNativeWordChangesToPackage,
} from '@coco-vexel/coco-office-editor/word';
import { renderSpreadsheetEditor } from '@coco-vexel/coco-office-editor/spreadsheet';
import { buildPptEditorViewFromPptx, renderNativePptEditor } from '@coco-vexel/coco-office-editor/ppt';
import { createSample } from '../../../examples/samples.mjs';

const host = document.querySelector('#editor');
const status = document.querySelector('#status');
const save = document.querySelector('#save');
const fileInput = document.querySelector('#file');
const controls = [...document.querySelectorAll('button, input')];
let current;
let busy = false;

async function run(action) {
  if (busy) return;
  busy = true;
  controls.forEach((control) => {
    control.disabled = true;
  });
  try {
    await action();
  } catch (error) {
    status.textContent = `Unable to complete: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    busy = false;
    controls.forEach((control) => {
      control.disabled = false;
    });
    save.disabled = !current;
  }
}

async function open(bytes, name) {
  const kind = name.split('.').pop().toLowerCase();
  if (!['docx', 'xlsx', 'pptx'].includes(kind))
    throw new Error('Choose a DOCX, XLSX or PPTX file.');
  status.textContent = `Opening ${name}…`;
  await current?.destroy();
  current = undefined;
  // A fresh host avoids reattaching a ShadowRoot when changing document types.
  const mount = document.createElement('div');
  mount.style.cssText = 'height:100%;min-height:500px';
  host.replaceChildren(mount);
  if (kind === 'docx') {
    const pkg = await fromZip(bytes);
    const editor = renderNativeWordEditor(
      mount,
      buildWordEditorViewFromPackage(pkg, buildIndices(pkg)),
    );
    current = {
      name,
      destroy: () => editor.destroy(),
      export: async () => {
        // Replay onto a fresh package so repeated exports never duplicate edits.
        const copy = await fromZip(bytes);
        const result = applyNativeWordChangesToPackage(
          copy,
          editor.getChanges(),
          buildIndices(copy),
        );
        if (result.plan.unsupported.length || result.skipped.length || result.unresolved.length) {
          throw new Error(
            'Some Word changes cannot be exported yet. Keep the document open and review your edits.',
          );
        }
        return toZip(copy);
      },
    };
  } else if (kind === 'xlsx') {
    const editor = await renderSpreadsheetEditor(mount, bytes, {
      workbookId: name,
      showSheetTabs: true,
      styleIsolation: 'shadow',
    });
    current = { name, destroy: () => editor.destroy(), export: () => editor.toZip() };
  } else {
    const view = await buildPptEditorViewFromPptx(bytes, { presentationId: name });
    const editor = renderNativePptEditor(mount, view);
    current = { name, destroy: () => editor.destroy(), export: () => editor.toPptx() };
  }
  status.textContent = `${name} — edit the document, then export a copy.`;
}

fileInput.addEventListener('change', () =>
  run(async () => {
    const file = fileInput.files?.[0];
    if (file) await open(new Uint8Array(await file.arrayBuffer()), file.name);
    fileInput.value = '';
  }),
);
document.querySelectorAll('[data-sample]').forEach((button) => {
  button.addEventListener('click', () =>
    run(async () => {
      const kind = button.dataset.sample;
      await open(await createSample(kind), `sample.${kind}`);
    }),
  );
});
save.addEventListener('click', () =>
  run(async () => {
    if (!current) return;
    const bytes = await current.export();
    const url = URL.createObjectURL(new Blob([bytes]));
    const link = document.createElement('a');
    link.href = url;
    link.download = current.name.replace(/\.(docx|xlsx|pptx)$/i, '.edited.$1');
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    status.textContent = `Exported ${link.download}. The original file is unchanged.`;
  }),
);
