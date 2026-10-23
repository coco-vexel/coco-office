import { EditorView as ProseMirrorEditorView } from 'prosemirror-view';
import {
  collectSpreadsheetFormulaReferences,
  formatCellAddress,
  parseCellAddress,
  parseRangeAddress,
  splitSheetQualifiedAddress,
  summarizeExcelTransactionGroup,
} from '@coco-vexel/coco-office/excel';
import { fromZip, toZip } from '@coco-vexel/coco-office/pack';
import { buildIndices } from '@coco-vexel/coco-office/word-view';
import {
  SpreadsheetCanvasRenderer,
  SpreadsheetEditorRuntime,
  renderSpreadsheetEditor,
  SpreadsheetLayout,
  SpreadsheetViewportManager,
  applyWheelDelta,
  computeScrollbarState,
  createColumnRangeSelection,
  createColumnSelection,
  createRowRangeSelection,
  createRowSelection,
  createSelection,
  extendSelection,
  expandSelectionForMergedRanges,
  applySpreadsheetAgentAction,
  freezeStateFromSheet,
  getRenderCell,
  moveSelection,
  normalizeSelectionRange,
  parseTsv,
  previewSpreadsheetAgentAction,
  buildSpreadsheetClipboardPayload,
  hitSpreadsheetResizeHandle as hitSpreadsheetResizeHandleAtPoint,
  scrollFromThumbPosition,
  selectionAfterColumnDeletion,
  selectionAfterRowDeletion,
  supportsSpreadsheetAutoFill,
  startSpreadsheetResizeDrag,
  summarizeSpreadsheetSelectionFormat,
  findSpreadsheetMatches,
  nextSpreadsheetMatch,
  spreadsheetFilterValues,
  updateSpreadsheetResizeDrag,
  type SpreadsheetAgentAction,
  type SpreadsheetEditor,
  type SpreadsheetAgentActionPreview,
  type SpreadsheetAgentActionApplyResult,
  type SpreadsheetClipboardPayload,
  type SpreadsheetPasteMode,
  type SpreadsheetRange,
  type SpreadsheetResizeDragState,
  type SpreadsheetResizeHandle,
  type SpreadsheetSelectionState,
  type SpreadsheetSheetRenderModel,
  type SpreadsheetViewport,
  type SpreadsheetMixedValue,
  type SpreadsheetFindMatch,
} from '../src/spreadsheet/index.js';
import { summarizeEditorCoverage, summarizeNativeCoverage } from '../src/word/coverage.js';
import { buildWordEditorViewFromPackage } from '../src/word/extract.js';
import {
  createWordEditingPlugin,
  getWordChangeSet,
  getWordSelectionInfo,
  setColor,
  setFontSize,
  setParagraphAlign,
  toggleBold,
} from '../src/word/editing.js';
import {
  applyNativeWordChangesToPackage,
  renderNativeWordEditor,
  summarizeNativeTablePagination,
  type NativeWordEditor,
  type NativeWordChangeApplyResult,
} from '../src/word/native.js';
import {
  buildPptEditorViewFromPptx,
  renderNativePptEditor,
  type NativePptEditor,
  type PptEditorImagePreviewRequest,
  type PptEditorImageResource,
  type PptSceneNode,
} from '../src/ppt/index.js';
import { createWordEditorState } from '../src/word/prosemirror.js';
import type { WordEditorLineSpacingState } from '../src/word/runtime.js';
import '../src/styles/editor.css';
import editorStyles from '../src/styles/editor.css?inline';
import './dev.css';
import 'prosemirror-view/style/prosemirror.css';

interface FixtureFile {
  index: number;
  name: string;
  path: string;
  size: number;
}

async function main(): Promise<void> {
  const app = document.querySelector<HTMLDivElement>('#app');
  if (!app) return;
  if (selectedSurface() === 'spreadsheet-public') {
    await renderSpreadsheetPublicDev(app);
    return;
  }
  if (selectedSurface() === 'spreadsheet') {
    await renderSpreadsheetDev(app);
    return;
  }
  if (selectedSurface() === 'ppt') {
    await renderPptDev(app);
    return;
  }
  app.className = 'coco-word-editor';
  app.innerHTML = `
    <div class="coco-word-shell">
      <main class="coco-word-document-scroll">
        <div class="coco-word-devbar">
          <select id="fixtureSelect" aria-label="Fixture document"></select>
          <select id="engineSelect" aria-label="Editor engine">
            <option value="prosemirror">ProseMirror</option>
            <option value="native">Native POC</option>
          </select>
          <span id="fixtureSummary">Loading fixture...</span>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="exportDocxButton" title="Export edited DOCX">Export DOCX</button>
        </div>
        <div class="coco-word-toolbar" aria-label="Word editing toolbar">
          <button type="button" id="wordUndoButton" title="Undo">↶</button>
          <button type="button" id="wordRedoButton" title="Redo">↷</button>
          <button type="button" id="trackRevisionsButton" title="Track changes">修订</button>
          <button type="button" id="acceptRevisionButton" title="Accept selected revisions">接受</button>
          <button type="button" id="rejectRevisionButton" title="Reject selected revisions">拒绝</button>
          <button type="button" id="acceptAllRevisionsButton" title="Accept all revisions">全收</button>
          <button type="button" id="rejectAllRevisionsButton" title="Reject all revisions">全拒</button>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="boldButton" title="Bold selected text"><strong>B</strong></button>
          <button type="button" id="italicButton" title="Italic selected text"><em>I</em></button>
          <button type="button" id="underlineButton" title="Underline selected text"><u>U</u></button>
          <button type="button" id="strikeButton" title="Strike selected text"><s>S</s></button>
          <label title="Font family">
            <span>字体</span>
            <select id="fontFamilySelect">
              <option value="宋体">宋体</option>
              <option value="等线">等线</option>
              <option value="微软雅黑">微软雅黑</option>
              <option value="Arial">Arial</option>
              <option value="Times New Roman">Times New Roman</option>
            </select>
          </label>
          <label title="Font size">
            <span>字号</span>
            <select id="fontSizeSelect">
              <option value="10">10</option>
              <option value="12" selected>12</option>
              <option value="14">14</option>
              <option value="16">16</option>
              <option value="18">18</option>
              <option value="24">24</option>
            </select>
          </label>
          <label title="Text color">
            <span>颜色</span>
            <input id="fontColorInput" type="color" value="#C00000" />
          </label>
          <button type="button" id="applyHighlightButton" title="Apply current highlight color">
            <span class="coco-word-highlight-icon">H</span>
          </button>
          <label title="Choose highlight color">
            <input id="highlightColorInput" type="color" value="#FFFF00" aria-label="Highlight color" />
          </label>
          <button type="button" id="clearHighlightButton" title="Remove highlight">H×</button>
          <button type="button" id="clearFormatButton" title="Clear text formatting">Tx</button>
          <input id="hyperlinkInput" class="coco-word-hyperlink-input" type="text" placeholder="https:// 或 #锚点" aria-label="Hyperlink target" />
          <button type="button" id="setHyperlinkButton" title="Set hyperlink">Link</button>
          <button type="button" id="removeHyperlinkButton" title="Remove hyperlink">Link×</button>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" data-align="left" title="Align left">左</button>
          <button type="button" data-align="center" title="Align center">中</button>
          <button type="button" data-align="right" title="Align right">右</button>
          <button type="button" data-align="both" title="Justify">齐</button>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="bulletListButton" title="Toggle bullet list">•</button>
          <button type="button" id="numberedListButton" title="Toggle numbered list">1.</button>
          <button type="button" id="decreaseListLevelButton" title="Decrease list level">⇤</button>
          <button type="button" id="increaseListLevelButton" title="Increase list level">⇥</button>
          <button type="button" id="removeListButton" title="Remove list formatting">×</button>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="decreaseIndentButton" title="Decrease left indent">←</button>
          <button type="button" id="increaseIndentButton" title="Increase left indent">→</button>
          <label title="First line or hanging indent">
            <span>首行</span>
            <select id="firstLineIndentSelect">
              <option value="0">无</option>
              <option value="420">2 字符</option>
              <option value="-420">悬挂 2 字符</option>
            </select>
          </label>
          <label title="Line spacing">
            <span>行距</span>
            <select id="lineSpacingSelect">
              <option value="240">1.0</option>
              <option value="276">1.15</option>
              <option value="360">1.5</option>
              <option value="480">2.0</option>
            </select>
          </label>
          <label title="Paragraph spacing before">
            <span>段前</span>
            <select id="spacingBeforeSelect">
              <option value="0">0</option>
              <option value="120">6 pt</option>
              <option value="240">12 pt</option>
              <option value="360">18 pt</option>
            </select>
          </label>
          <label title="Paragraph spacing after">
            <span>段后</span>
            <select id="spacingAfterSelect">
              <option value="0">0</option>
              <option value="120">6 pt</option>
              <option value="240">12 pt</option>
              <option value="360">18 pt</option>
            </select>
          </label>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="insertPageBreakButton" title="Insert page break (Ctrl+Enter)">分页</button>
          <button type="button" id="insertImageButton" title="Insert image">Image</button>
          <input id="imageFileInput" type="file" accept="image/*" hidden />
          <label title="Image width">
            <span>W</span>
            <input id="imageWidthInput" class="coco-word-image-size-input" type="number" min="24" step="1" aria-label="Image width" />
          </label>
          <label title="Image height">
            <span>H</span>
            <input id="imageHeightInput" class="coco-word-image-size-input" type="number" min="24" step="1" aria-label="Image height" />
          </label>
          <button type="button" id="deleteImageButton" title="Delete selected image">×</button>
          <span class="coco-word-toolbar-divider"></span>
          <label title="Table rows"><span>行</span><input id="newTableRowsInput" class="coco-word-table-size-input" type="number" min="1" max="20" value="2" /></label>
          <label title="Table columns"><span>列</span><input id="newTableColumnsInput" class="coco-word-table-size-input" type="number" min="1" max="20" value="2" /></label>
          <button type="button" id="insertTableButton" title="Insert table">表格+</button>
          <button type="button" id="deleteTableButton" title="Delete current table">表格×</button>
          <span class="coco-word-toolbar-divider"></span>
          <button type="button" id="insertTableRowBeforeButton" title="Insert row above">行↑</button>
          <button type="button" id="insertTableRowAfterButton" title="Insert row below">行↓</button>
          <button type="button" id="deleteTableRowButton" title="Delete current row">删行</button>
          <button type="button" id="insertTableColumnBeforeButton" title="Insert column left">列←</button>
          <button type="button" id="insertTableColumnAfterButton" title="Insert column right">列→</button>
          <button type="button" id="deleteTableColumnButton" title="Delete current column">删列</button>
          <button type="button" id="mergeTableCellsButton" title="Merge selected cells">合并</button>
          <button type="button" id="splitTableCellButton" title="Split selected cell">拆分</button>
          <label title="Cell background color">
            <span>底色</span>
            <input id="tableCellBackgroundInput" type="color" value="#FFF2CC" aria-label="Cell background color" />
          </label>
          <button type="button" id="clearTableCellBackgroundButton" title="Clear cell background">底×</button>
          <label title="Cell vertical alignment">
            <span>垂直</span>
            <select id="tableCellVerticalAlignmentSelect">
              <option value="top">上</option>
              <option value="center">中</option>
              <option value="bottom">下</option>
            </select>
          </label>
          <label title="Cell border color">
            <span>边框</span>
            <input id="tableCellBorderColorInput" type="color" value="#000000" aria-label="Cell border color" />
          </label>
          <button type="button" id="applyTableCellBordersButton" title="Apply cell borders">□</button>
          <button type="button" id="clearTableCellBordersButton" title="Remove cell borders">□×</button>
        </div>
        <div class="coco-word-searchbar" id="wordSearchBar" hidden>
          <input id="wordSearchInput" type="search" placeholder="查找" aria-label="Find text" />
          <span id="wordSearchCount" aria-live="polite">0 / 0</span>
          <button type="button" id="wordSearchPreviousButton" title="Previous match">↑</button>
          <button type="button" id="wordSearchNextButton" title="Next match">↓</button>
          <label title="Match case"><input id="wordSearchMatchCase" type="checkbox" />Aa</label>
          <label title="Whole word"><input id="wordSearchWholeWord" type="checkbox" />词</label>
          <span class="coco-word-toolbar-divider"></span>
          <input id="wordReplaceInput" type="text" placeholder="替换为" aria-label="Replacement text" />
          <button type="button" id="wordReplaceCurrentButton">替换</button>
          <button type="button" id="wordReplaceAllButton">全部替换</button>
          <button type="button" id="wordSearchCloseButton" title="Close find">×</button>
        </div>
        <section class="coco-word-page" id="editor"></section>
      </main>
      <aside class="coco-word-inspector">
        <h2>Validation</h2>
        <div id="acceptancePanel" class="coco-word-acceptance">Loading...</div>
        <h2>EditorViewModel</h2>
        <pre id="inspector">Loading...</pre>
      </aside>
    </div>
  `;

  const fixtures = await loadFixtures();
  const selectedIndex = selectedFixtureIndex(fixtures);
  const engine = selectedEngine();
  renderFixtureSelect(fixtures, selectedIndex);
  renderEngineSelect(engine);

  const response = await fetch(`/__coco_fixture__/docx?index=${selectedIndex}`);
  if (!response.ok) throw new Error(`Failed to load fixture: ${response.status}`);
  const fixturePath = decodeURIComponent(response.headers.get('X-Coco-Fixture-Path') ?? '');
  const bytes = new Uint8Array(await response.arrayBuffer());
  const pkg = await fromZip(bytes);
  const indices = buildIndices(pkg);
  const wordView = buildWordEditorViewFromPackage(pkg, indices);
  let renderInspector = () => {};
  const state = createWordEditorState(wordView, {
    plugins: [
      createWordEditingPlugin({
        onChangeSet: () => renderInspector(),
      }),
    ],
  });
  const editorRoot = document.querySelector<HTMLDivElement>('#editor');
  const inspector = document.querySelector<HTMLPreElement>('#inspector');
  const acceptancePanel = document.querySelector<HTMLDivElement>('#acceptancePanel');
  const fixtureSummary = document.querySelector<HTMLSpanElement>('#fixtureSummary');
  if (!editorRoot || !inspector) return;

  const previewWidth = Number(new URLSearchParams(location.search).get('previewWidth'));
  if (Number.isFinite(previewWidth) && previewWidth > 0) {
    editorRoot.style.width = `${previewWidth}px`;
    editorRoot.style.marginInline = 'auto';
  }

  if (engine === 'native') {
    editorRoot.classList.remove('coco-word-page');
    editorRoot.classList.add('coco-word-pages');
    let nativeEditor: NativeWordEditor | null = null;
    let appliedChangeCount = 0;
    let lastSaveResult: NativeWordChangeApplyResult | null = null;
    let lastSavedDocxSummary: unknown = null;
    nativeEditor = renderNativeWordEditor(editorRoot, wordView, {
      readOnly: selectedReadOnly(),
      onChange: () => renderInspector(),
      styleText: editorStyles,
    });
    wireNativeToolbar(nativeEditor);
    document
      .querySelector<HTMLButtonElement>('#exportDocxButton')
      ?.addEventListener('click', async () => {
        if (!nativeEditor) return;
        const changes = nativeEditor.getChanges();
        const pendingChanges = changes.slice(appliedChangeCount);
        lastSaveResult = applyNativeWordChangesToPackage(pkg, pendingChanges, buildIndices(pkg));
        if (
          lastSaveResult.plan.unsupported.length === 0 &&
          lastSaveResult.skipped.length === 0 &&
          lastSaveResult.unresolved.length === 0
        ) {
          appliedChangeCount = changes.length;
        }
        const out = await toZip(pkg);
        lastSavedDocxSummary = await summarizeSavedDocx(out, fixturePath);
        downloadBytes(out, exportedFileName(fixtures[selectedIndex]?.name ?? 'document.docx'));
        renderInspector();
      });
    if (fixtureSummary) {
      fixtureSummary.textContent = `${fixtures[selectedIndex]?.name ?? 'fixture'} | native POC | ${wordView.paragraphs.length} paragraphs | ${wordView.tables.length} tables`;
    }
    renderInspector = () => {
      const nativeDocument = nativeEditor?.getDocument();
      renderAcceptancePanel(
        acceptancePanel,
        buildAcceptanceReport(wordView, nativeDocument, collectRenderedDomSnapshot(editorRoot)),
      );
      inspector.textContent = JSON.stringify(
        {
          engine: 'native',
          document: summarizeWordView(fixturePath, bytes.byteLength, wordView),
          coverage: summarizeNativeCoverage(wordView, nativeDocument),
          nativeDocument: summarizeNativeDocument(nativeDocument),
          changes: nativeEditor?.getChanges(),
          history: {
            canUndo: nativeEditor?.canUndo() ?? false,
            canRedo: nativeEditor?.canRedo() ?? false,
          },
          save: {
            appliedChangeCount,
            pendingChangeCount: Math.max(
              0,
              (nativeEditor?.getChanges().length ?? 0) - appliedChangeCount,
            ),
            lastSaveResult,
            lastSavedDocxSummary,
          },
        },
        null,
        2,
      );
    };
    renderInspector();
    return;
  }

  applyEditorPageLayout(editorRoot, wordView.pageLayout);
  const editorView = new ProseMirrorEditorView(editorRoot, {
    state,
    dispatchTransaction(transaction) {
      const nextState = editorView.state.apply(transaction);
      editorView.updateState(nextState);
      renderInspector();
    },
    handleDOMEvents: {
      mouseup() {
        requestAnimationFrame(renderInspector);
        return false;
      },
      keyup() {
        requestAnimationFrame(renderInspector);
        return false;
      },
    },
  });
  document.querySelector<HTMLButtonElement>('#exportDocxButton')?.setAttribute('disabled', 'true');
  if (fixtureSummary) {
    fixtureSummary.textContent = `${fixtures[selectedIndex]?.name ?? 'fixture'} · ${wordView.paragraphs.length} paragraphs · ${wordView.tables.length} tables`;
  }

  wireToolbar(editorView);
  renderInspector = () => {
    renderAcceptancePanel(
      acceptancePanel,
      buildAcceptanceReport(wordView, undefined, collectRenderedDomSnapshot(editorRoot)),
    );
    inspector.textContent = JSON.stringify(
      {
        document: summarizeWordView(fixturePath, bytes.byteLength, wordView),
        selection: getWordSelectionInfo(editorView.state),
        changes: getWordChangeSet(editorView.state),
      },
      null,
      2,
    );
  };
  renderInspector();
}

async function renderPptDev(app: HTMLDivElement): Promise<void> {
  app.className = 'coco-ppt-dev';
  app.innerHTML = `
    <style>
      .coco-ppt-dev { height: 100vh; display: grid; grid-template-rows: auto 1fr; background: #f6f7f9; color: #17202a; font-family: Arial, sans-serif; overflow: hidden; }
      .coco-ppt-devbar { display: flex; align-items: center; gap: 8px; min-height: 42px; padding: 6px 10px; border-bottom: 1px solid #d1d5db; background: #ffffff; flex-wrap: wrap; }
      .coco-ppt-devbar select, .coco-ppt-devbar button, .coco-ppt-devbar input { height: 28px; border: 1px solid #cbd5e1; border-radius: 4px; background: #ffffff; color: #111827; }
      .coco-ppt-devbar input { min-width: 220px; padding: 0 8px; }
      .coco-ppt-devbar input[type="color"] { width: 34px; min-width: 34px; padding: 2px; }
      .coco-ppt-devbar input[type="number"] { width: 74px; min-width: 74px; }
      .coco-ppt-devbar input.coco-ppt-cell-text { width: 160px; min-width: 160px; }
      .coco-ppt-devbar label { display: inline-flex; align-items: center; gap: 4px; color: #475569; font-size: 12px; }
      .coco-ppt-devbar button:disabled, .coco-ppt-devbar input:disabled { opacity: 0.45; cursor: not-allowed; }
      .coco-ppt-summary { color: #475569; font-size: 12px; white-space: nowrap; }
      .coco-ppt-edit-status { color: #2563eb; font-size: 12px; white-space: nowrap; }
      .coco-ppt-shell { min-height: 0; display: grid; grid-template-columns: minmax(0, 1fr) 360px; }
      .coco-ppt-host { min-height: 0; }
      .coco-ppt-inspector { min-width: 0; overflow: auto; border-left: 1px solid #d1d5db; background: #ffffff; padding: 12px; }
      .coco-ppt-inspector h2 { margin: 0 0 8px; font-size: 14px; }
      .coco-ppt-inspector pre { margin: 0; white-space: pre-wrap; word-break: break-word; font-size: 12px; line-height: 1.45; }
      .coco-ppt-object-panel { display: flex; flex-direction: column; gap: 12px; font-size: 12px; color: #334155; }
      .coco-ppt-object-title { font-size: 14px; font-weight: 700; color: #111827; line-height: 1.35; }
      .coco-ppt-object-meta { display: grid; grid-template-columns: 72px minmax(0, 1fr); gap: 6px 8px; }
      .coco-ppt-object-meta dt { margin: 0; color: #64748b; }
      .coco-ppt-object-meta dd { margin: 0; min-width: 0; word-break: break-word; }
      .coco-ppt-capability-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px; }
      .coco-ppt-capability { border: 1px solid #dbe3ee; border-radius: 4px; padding: 6px 8px; background: #f8fafc; color: #64748b; line-height: 1.35; }
      .coco-ppt-capability[data-enabled="true"] { border-color: #93c5fd; background: #eff6ff; color: #1d4ed8; font-weight: 700; }
      .coco-ppt-capability-reason { display: block; margin-top: 3px; color: #7c2d12; font-weight: 400; word-break: break-word; }
      .coco-ppt-reason { border-left: 3px solid #f59e0b; background: #fffbeb; color: #92400e; padding: 8px 10px; line-height: 1.5; }
      .coco-ppt-hit-hint { border-left: 3px solid #2563eb; background: #eff6ff; color: #1e3a8a; padding: 8px 10px; line-height: 1.5; }
      .coco-ppt-text-preview { max-height: 96px; overflow: auto; padding: 8px; border: 1px solid #e2e8f0; border-radius: 4px; background: #f8fafc; white-space: pre-wrap; word-break: break-word; }
      .coco-ppt-json summary { cursor: pointer; color: #475569; margin-bottom: 6px; }
    </style>
    <div class="coco-ppt-devbar">
      <select id="pptFixtureSelect" aria-label="PPT fixture"></select>
      <button type="button" id="pptUndoButton" title="撤销">撤销</button>
      <button type="button" id="pptRedoButton" title="重做">重做</button>
      <input id="pptTextInput" type="text" placeholder="选中文本对象后编辑" aria-label="Selected shape text" />
      <button type="button" id="pptApplyTextButton" title="只修改选中对象的文字内容">修改文字</button>
      <label title="Font size">字号 <input id="pptFontSizeInput" type="number" min="1" max="200" step="1" /></label>
      <label title="Text color">文字 <input id="pptTextColorInput" type="color" /></label>
      <label title="Fill color">填充 <input id="pptFillColorInput" type="color" /></label>
      <label title="Line color">线条 <input id="pptLineColorInput" type="color" /></label>
      <label title="Line width">线宽 <input id="pptLineWidthInput" type="number" min="0" step="9525" /></label>
      <button type="button" id="pptApplyStyleButton" title="修改选中对象支持的文字/形状样式">修改样式</button>
      <label title="Left">X <input id="pptLeftInput" type="number" step="10000" /></label>
      <label title="Top">Y <input id="pptTopInput" type="number" step="10000" /></label>
      <label title="Width">W <input id="pptWidthInput" type="number" min="1" step="10000" /></label>
      <label title="Height">H <input id="pptHeightInput" type="number" min="1" step="10000" /></label>
      <button type="button" id="pptApplyGeometryButton" title="修改选中对象的位置和大小">修改位置</button>
      <label title="Table row index">行 <input id="pptTableRowInput" type="number" min="0" step="1" /></label>
      <label title="Table column index">列 <input id="pptTableColumnInput" type="number" min="0" step="1" /></label>
      <input id="pptTableCellTextInput" class="coco-ppt-cell-text" type="text" placeholder="table cell text" aria-label="Table cell text" />
      <button type="button" id="pptApplyTableCellButton" title="修改选中表格的指定单元格">修改单元格</button>
      <button type="button" id="pptExportButton">导出 PPTX</button>
      <button type="button" id="pptVerifyExportButton">校验导出</button>
      <span id="pptSummary" class="coco-ppt-summary">Loading...</span>
      <span id="pptEditStatus" class="coco-ppt-edit-status"></span>
    </div>
    <div class="coco-ppt-shell">
      <div id="pptEditorHost" class="coco-ppt-host"></div>
      <aside class="coco-ppt-inspector">
        <h2>对象边界</h2>
        <div id="pptInspector" class="coco-ppt-object-panel">Loading...</div>
      </aside>
    </div>
  `;

  const fixtures = await loadPptFixtures();
  const selectedIndex = selectedPptFixtureIndex(fixtures);
  const fixtureSelect = app.querySelector<HTMLSelectElement>('#pptFixtureSelect');
  const summary = app.querySelector<HTMLSpanElement>('#pptSummary');
  const host = app.querySelector<HTMLDivElement>('#pptEditorHost');
  const inspector = app.querySelector<HTMLElement>('#pptInspector');
  const textInput = app.querySelector<HTMLInputElement>('#pptTextInput');
  const applyTextButton = app.querySelector<HTMLButtonElement>('#pptApplyTextButton');
  const fontSizeInput = app.querySelector<HTMLInputElement>('#pptFontSizeInput');
  const textColorInput = app.querySelector<HTMLInputElement>('#pptTextColorInput');
  const fillColorInput = app.querySelector<HTMLInputElement>('#pptFillColorInput');
  const lineColorInput = app.querySelector<HTMLInputElement>('#pptLineColorInput');
  const lineWidthInput = app.querySelector<HTMLInputElement>('#pptLineWidthInput');
  const applyStyleButton = app.querySelector<HTMLButtonElement>('#pptApplyStyleButton');
  const leftInput = app.querySelector<HTMLInputElement>('#pptLeftInput');
  const topInput = app.querySelector<HTMLInputElement>('#pptTopInput');
  const widthInput = app.querySelector<HTMLInputElement>('#pptWidthInput');
  const heightInput = app.querySelector<HTMLInputElement>('#pptHeightInput');
  const applyGeometryButton = app.querySelector<HTMLButtonElement>('#pptApplyGeometryButton');
  const tableRowInput = app.querySelector<HTMLInputElement>('#pptTableRowInput');
  const tableColumnInput = app.querySelector<HTMLInputElement>('#pptTableColumnInput');
  const tableCellTextInput = app.querySelector<HTMLInputElement>('#pptTableCellTextInput');
  const applyTableCellButton = app.querySelector<HTMLButtonElement>('#pptApplyTableCellButton');
  const undoButton = app.querySelector<HTMLButtonElement>('#pptUndoButton');
  const redoButton = app.querySelector<HTMLButtonElement>('#pptRedoButton');
  const exportButton = app.querySelector<HTMLButtonElement>('#pptExportButton');
  const verifyExportButton = app.querySelector<HTMLButtonElement>('#pptVerifyExportButton');
  const editStatus = app.querySelector<HTMLSpanElement>('#pptEditStatus');
  if (
    !fixtureSelect ||
    !summary ||
    !host ||
    !inspector ||
    !textInput ||
    !applyTextButton ||
    !fontSizeInput ||
    !textColorInput ||
    !fillColorInput ||
    !lineColorInput ||
    !lineWidthInput ||
    !applyStyleButton ||
    !leftInput ||
    !topInput ||
    !widthInput ||
    !heightInput ||
    !applyGeometryButton ||
    !tableRowInput ||
    !tableColumnInput ||
    !tableCellTextInput ||
    !applyTableCellButton ||
    !undoButton ||
    !redoButton ||
    !exportButton ||
    !verifyExportButton ||
    !editStatus
  )
    return;

  fixtureSelect.innerHTML = fixtures
    .map(
      (fixture) =>
        `<option value="${fixture.index}"${fixture.index === selectedIndex ? ' selected' : ''}>${fixture.index + 1}. ${fixture.name}</option>`,
    )
    .join('');
  fixtureSelect.addEventListener('change', () => {
    const params = new URLSearchParams(location.search);
    params.set('surface', 'ppt');
    params.set('pptFixture', fixtureSelect.value);
    location.search = `?${params.toString()}`;
  });

  const response = await fetch(`/__coco_fixture__/pptx?index=${selectedIndex}`);
  if (!response.ok) throw new Error(`Failed to load pptx fixture: ${response.status}`);
  const fixturePath = decodeURIComponent(response.headers.get('X-Coco-Fixture-Path') ?? '');
  const bytes = new Uint8Array(await response.arrayBuffer());
  const view = await buildPptEditorViewFromPptx(bytes, {
    presentationId: fixtures[selectedIndex]?.name ?? `ppt-fixture-${selectedIndex}`,
    resolveImagePreview: resolvePptImagePreview,
  });
  const previewWidth = Number(new URLSearchParams(location.search).get('previewWidth'));
  if (Number.isFinite(previewWidth) && previewWidth > 0) {
    host.style.width = `${previewWidth}px`;
    host.style.marginInline = 'auto';
  }
  let lastObjectAction:
    | {
        action:
          | 'blocked-text-edit'
          | 'cycle-selection'
          | 'nudge-geometry'
          | 'snap-geometry'
          | 'copy-shape'
          | 'paste-shape'
          | 'delete-shape'
          | 'duplicate-shape'
          | 'bring-shape-to-front'
          | 'send-shape-to-back';
        slideId: string;
        shapeId: string;
        kind: PptSceneNode['kind'];
        name?: string;
        reason: string;
        candidateIndex?: number;
        candidateCount?: number;
      }
    | null = null;
  const editor = renderNativePptEditor(host, view, {
    styleIsolation: 'shadow',
    readOnly: selectedReadOnly(),
    showSlideList: selectedPptSlideListVisible(),
    onChange: () => render(),
    onObjectAction: (event) => {
      lastObjectAction = event;
      editStatus.textContent = event.reason;
      render();
    },
    onTextEditStateChange: (event) => {
      if (event.state === 'start') editStatus.textContent = '已进入文字编辑';
      if (event.state === 'commit') editStatus.textContent = '已提交到内存';
      if (event.state === 'cancel') editStatus.textContent = '已取消编辑';
      render();
    },
    onTableCellSelectionChange: (event) => {
      if (document.activeElement !== tableRowInput) tableRowInput.value = String(event.row);
      if (document.activeElement !== tableColumnInput) tableColumnInput.value = String(event.column);
      if (document.activeElement !== tableCellTextInput) tableCellTextInput.value = event.text;
      editStatus.textContent = `已选中单元格 ${event.row},${event.column}`;
    },
  });
  (window as unknown as { __cocoPptDev?: NativePptEditor }).__cocoPptDev = editor;

  const selectedShape = () => {
    const selection = editor.selection();
    if (selection?.kind !== 'shape') return null;
    const slide = editor
      .session()
      .scene()
      .slides.find((item) => item.id === selection.slideId);
    return (
      slide?.nodes.find((node) => node.id === selection.shapeId) ??
      slide?.groups.find((node) => node.id === selection.shapeId) ??
      null
    );
  };

  const validHexColor = (value: string | undefined, fallback: string): string =>
    value && /^#[0-9a-f]{6}$/i.test(value) ? value : fallback;

  const readNumber = (input: HTMLInputElement): number | undefined => {
    if (input.value.trim() === '') return undefined;
    const value = Number(input.value);
    return Number.isFinite(value) ? value : undefined;
  };

  const setInputValue = (input: HTMLInputElement, value: string): void => {
    if (document.activeElement !== input) input.value = value;
  };

  const disabledReason = (
    shape: PptSceneNode | null,
    capability: keyof PptSceneNode['capabilities']['reasons'],
    fallback: string,
  ): string => {
    if (!shape) return '请先选择一个对象';
    return shape.capabilities.reasons?.[capability] ?? shape.capabilities.reason ?? fallback;
  };

  const setControlState = (
    control: HTMLInputElement | HTMLButtonElement,
    disabled: boolean,
    enabledTitle: string,
    reason?: string,
  ): void => {
    control.disabled = disabled;
    control.title = disabled && reason ? reason : enabledTitle;
  };

  const setControlsState = (
    controls: (HTMLInputElement | HTMLButtonElement)[],
    disabled: boolean,
    enabledTitle: string,
    reason?: string,
  ): void => {
    for (const control of controls) {
      setControlState(control, disabled, enabledTitle, reason);
    }
  };

  const setToolbarHint = (shape: PptSceneNode | null): void => {
    if (!shape) {
      textInput.placeholder = '选中文本对象后编辑';
      editStatus.textContent = editStatus.textContent || '请选择一个对象';
      return;
    }
    textInput.placeholder = shape.capabilities.editText
      ? '选中对象文字'
      : disabledReason(shape, 'editText', '当前对象不支持文字编辑');
    if (
      !shape.capabilities.editText &&
      !shape.capabilities.editTextStyle &&
      !shape.capabilities.editShapeStyle &&
      !shape.capabilities.editGeometry &&
      !shape.capabilities.editTableCells
    ) {
      editStatus.textContent = '当前对象不支持常规编辑';
    }
  };

  const setShapeControlsDisabled = (disabled: boolean) => {
    const reason = disabled ? '请先选择一个对象' : undefined;
    setControlsState([fontSizeInput, textColorInput], disabled, '修改选中对象文字样式', reason);
    setControlsState(
      [fillColorInput, lineColorInput, lineWidthInput],
      disabled,
      '修改选中对象形状样式',
      reason,
    );
    setControlsState(
      [leftInput, topInput, widthInput, heightInput],
      disabled,
      '修改选中对象位置和大小',
      reason,
    );
    setControlState(applyStyleButton, disabled, '修改选中对象支持的文字/形状样式', reason);
    setControlState(applyGeometryButton, disabled, '修改选中对象的位置和大小', reason);
  };

  const setTableControlsDisabled = (disabled: boolean) => {
    const reason = disabled ? '请选择一个可编辑表格对象' : undefined;
    setControlsState(
      [tableRowInput, tableColumnInput],
      disabled,
      '选择表格单元格坐标',
      reason,
    );
    setControlsState(
      [tableCellTextInput, applyTableCellButton],
      disabled,
      '修改选中表格的指定单元格',
      reason,
    );
  };

  const readTableCellText = () => {
    const shape = selectedShape();
    if (!shape?.table) return '';
    const { row, column } = readSelectedTableCellCoordinates(shape);
    return shape.table.values[row]?.[column] ?? '';
  };

  const readSelectedTableCellCoordinates = (shape: PptSceneNode) => {
    const table = shape.table;
    if (!table) return { row: 0, column: 0 };
    const row = Math.max(
      0,
      Math.min(table.rowCount - 1, Math.trunc(readNumber(tableRowInput) ?? 0)),
    );
    const column = Math.max(
      0,
      Math.min(table.columnCount - 1, Math.trunc(readNumber(tableColumnInput) ?? 0)),
    );
    return { row, column };
  };

  const selectedTableCellMerge = (shape: PptSceneNode) => {
    if (!shape.table) return undefined;
    const { row, column } = readSelectedTableCellCoordinates(shape);
    return shape.table.mergedCells?.find((merge) => merge.row === row && merge.column === column);
  };

  const selectedTableCellEditReason = (shape: PptSceneNode): string | undefined => {
    if (!shape.table) return '当前对象不是表格';
    if (!shape.capabilities.editTableCells)
      return shape.capabilities.reasons?.editTableCells ?? '当前表格单元格不可编辑';
    const merge = selectedTableCellMerge(shape);
    if (merge?.hMerge || merge?.vMerge) return '合并延续单元格不可直接编辑，请选择合并主单元格';
    return undefined;
  };

  const render = () => {
    const scene = editor.session().scene();
    const shape = selectedShape();
    const table = shape?.table;
    const capabilities = shape?.capabilities;
    undoButton.disabled = !editor.canUndo();
    redoButton.disabled = !editor.canRedo();
    const textReason = capabilities?.editText
      ? undefined
      : disabledReason(shape, 'editText', '当前对象不支持文字编辑');
    setControlsState(
      [textInput, applyTextButton],
      !capabilities?.editText,
      '只修改选中对象的文字内容',
      textReason,
    );
    setShapeControlsDisabled(!shape);
    const textStyleReason = capabilities?.editTextStyle
      ? undefined
      : disabledReason(shape, 'editTextStyle', '当前对象不支持文字样式编辑');
    const shapeStyleReason = capabilities?.editShapeStyle
      ? undefined
      : disabledReason(shape, 'editShapeStyle', '当前对象不支持形状样式编辑');
    const geometryReason = capabilities?.editGeometry
      ? undefined
      : disabledReason(shape, 'editGeometry', '当前对象不支持位置尺寸编辑');
    setControlsState(
      [fontSizeInput, textColorInput],
      !capabilities?.editTextStyle,
      '修改选中对象文字样式',
      textStyleReason,
    );
    setControlsState(
      [fillColorInput, lineColorInput, lineWidthInput],
      !capabilities?.editShapeStyle,
      '修改选中对象形状样式',
      shapeStyleReason,
    );
    setControlState(
      applyStyleButton,
      !capabilities?.editTextStyle && !capabilities?.editShapeStyle,
      '修改选中对象支持的文字/形状样式',
      textStyleReason && shapeStyleReason ? `${textStyleReason}；${shapeStyleReason}` : undefined,
    );
    setControlsState(
      [leftInput, topInput, widthInput, heightInput],
      !capabilities?.editGeometry,
      '修改选中对象位置和大小',
      geometryReason,
    );
    setControlState(
      applyGeometryButton,
      !capabilities?.editGeometry,
      '修改选中对象的位置和大小',
      geometryReason,
    );
    setTableControlsDisabled(!capabilities?.editTableCells || !table);
    setToolbarHint(shape);
    if (shape) {
      setInputValue(textInput, shape.text ?? '');
      setInputValue(fontSizeInput, shape.textStyle?.fontSize ? String(shape.textStyle.fontSize) : '');
      setInputValue(textColorInput, validHexColor(shape.textStyle?.color, '#111827'));
      setInputValue(fillColorInput, validHexColor(shape.fillColor, '#ffffff'));
      setInputValue(lineColorInput, validHexColor(shape.lineColor, '#111827'));
      setInputValue(lineWidthInput, shape.lineWidth !== undefined ? String(shape.lineWidth) : '');
      setInputValue(leftInput, String(Math.round(shape.bounds.left)));
      setInputValue(topInput, String(Math.round(shape.bounds.top)));
      setInputValue(widthInput, String(Math.round(shape.bounds.width)));
      setInputValue(heightInput, String(Math.round(shape.bounds.height)));
    } else {
      for (const control of [
        textInput,
        fontSizeInput,
        lineWidthInput,
        leftInput,
        topInput,
        widthInput,
        heightInput,
      ]) {
        setInputValue(control, '');
      }
    }
    if (table) {
      tableRowInput.max = String(Math.max(0, table.rowCount - 1));
      tableColumnInput.max = String(Math.max(0, table.columnCount - 1));
      if (tableRowInput.value.trim() === '') tableRowInput.value = '0';
      if (tableColumnInput.value.trim() === '') tableColumnInput.value = '0';
      const cellEditReason = selectedTableCellEditReason(shape);
      const tableReason = capabilities?.editTableCells
        ? undefined
        : disabledReason(shape, 'editTableCells', '当前表格单元格不可编辑');
      setControlsState(
        [tableRowInput, tableColumnInput],
        !capabilities?.editTableCells,
        '选择表格单元格坐标',
        tableReason,
      );
      setControlsState(
        [tableCellTextInput, applyTableCellButton],
        Boolean(cellEditReason),
        '修改选中表格的指定单元格',
        cellEditReason,
      );
      tableCellTextInput.placeholder = cellEditReason ?? '选中单元格文字';
      if (cellEditReason && editStatus.textContent === '') editStatus.textContent = cellEditReason;
      if (document.activeElement !== tableCellTextInput)
        tableCellTextInput.value = readTableCellText();
    } else if (document.activeElement !== tableCellTextInput) {
      tableRowInput.removeAttribute('max');
      tableColumnInput.removeAttribute('max');
      tableRowInput.value = '';
      tableColumnInput.value = '';
      tableCellTextInput.value = '';
    }
    summary.textContent = `${fixtures[selectedIndex]?.name ?? 'fixture'} | ${scene.slides.length} slides | ${scene.slides.reduce((count, slide) => count + slide.nodes.length, 0)} shapes`;
    const inspectorPayload = {
      fixturePath,
      bytes: bytes.byteLength,
      revision: editor.session().revision(),
      selection: editor.selection(),
      selectedShape: shape
        ? {
            id: shape.id,
            name: shape.name,
            kind: shape.kind,
            bounds: shape.bounds,
            groupContext: shape.groupContext,
            editable: shape.editable,
            locked: shape.locked,
            capabilities: shape.capabilities,
            text: shape.text,
            table: shape.table
              ? {
                  rowCount: shape.table.rowCount,
                  columnCount: shape.table.columnCount,
                  values: shape.table.values,
                }
              : undefined,
          }
        : null,
      slides: scene.slides.map((slide) => ({
        id: slide.id,
        index: slide.index,
        title: slide.title,
        shapes: slide.nodes.length,
        preview: slide.nodes.slice(0, 8).map((node) => ({
          id: node.id,
          name: node.name,
          kind: node.kind,
          editable: node.editable,
          capabilities: node.capabilities,
          text: node.text?.slice(0, 80),
        })),
      })),
    };
    inspector.innerHTML = renderPptInspectorPanel({
      fixtureName: fixtures[selectedIndex]?.name ?? 'fixture',
      sceneSlideCount: scene.slides.length,
      sceneShapeCount: scene.slides.reduce((count, slide) => count + slide.nodes.length, 0),
      selection: editor.selection(),
      selectedShape: shape,
      lastObjectAction,
      debug: inspectorPayload,
    });
  };

  editor.session().subscribe('selectionChange', render);
  editor.session().subscribe('change', render);
  undoButton.addEventListener('click', () => void editor.undo().then(render));
  redoButton.addEventListener('click', () => void editor.redo().then(render));
  const dispatchSelectedShapeCommand = async (
    command: Parameters<NativePptEditor['dispatch']>[0],
  ): Promise<void> => {
    try {
      editStatus.textContent = '正在应用...';
      await editor.dispatch(command);
      editStatus.textContent = '已应用到内存';
      render();
    } catch (error) {
      editStatus.textContent = error instanceof Error ? error.message : 'Edit failed';
    }
  };
  const applySelectedTextStyle = async (
    shapeId: string,
    slideId: string,
    style: { fontSize?: number; color: string },
  ): Promise<void> => {
    await dispatchSelectedShapeCommand({
      kind: 'set-shape-text-style',
      slideIdOrIndex: slideId,
      shapeIdOrName: shapeId,
      style: {
        ...(style.fontSize !== undefined ? { fontSize: style.fontSize } : {}),
        color: style.color,
      },
    });
  };
  applyTextButton.addEventListener('click', () => {
    const shape = selectedShape();
    const selection = editor.selection();
    if (!shape?.capabilities.editText || selection?.kind !== 'shape') return;
    void dispatchSelectedShapeCommand({
      kind: 'set-shape-text',
      slideIdOrIndex: selection.slideId,
      shapeIdOrName: shape.id,
      text: textInput.value,
    });
  });
  applyStyleButton.addEventListener('click', async () => {
    const shape = selectedShape();
    const selection = editor.selection();
    if (!shape || selection?.kind !== 'shape') return;
    const fillColor = fillColorInput.value;
    const lineColor = lineColorInput.value;
    const lineWidth = readNumber(lineWidthInput);
    if (shape.capabilities.editTextStyle) {
      await applySelectedTextStyle(shape.id, selection.slideId, {
        fontSize: readNumber(fontSizeInput),
        color: textColorInput.value,
      });
    }
    if (shape.capabilities.editShapeStyle) {
      await dispatchSelectedShapeCommand({
        kind: 'set-shape-style',
        slideIdOrIndex: selection.slideId,
        shapeIdOrName: shape.id,
        style: {
          fillColor,
          lineColor,
          lineWidth,
        },
      });
    }
  });
  applyGeometryButton.addEventListener('click', () => {
    const shape = selectedShape();
    const selection = editor.selection();
    if (!shape?.capabilities.editGeometry || selection?.kind !== 'shape') return;
    void dispatchSelectedShapeCommand({
      kind: 'set-shape-geometry',
      slideIdOrIndex: selection.slideId,
      shapeIdOrName: shape.id,
      geometry: {
        left: readNumber(leftInput),
        top: readNumber(topInput),
        width: readNumber(widthInput),
        height: readNumber(heightInput),
      },
    });
  });
  tableRowInput.addEventListener('change', () => {
    if (document.activeElement !== tableCellTextInput)
      tableCellTextInput.value = readTableCellText();
    render();
  });
  tableColumnInput.addEventListener('change', () => {
    if (document.activeElement !== tableCellTextInput)
      tableCellTextInput.value = readTableCellText();
    render();
  });
  applyTableCellButton.addEventListener('click', () => {
    const shape = selectedShape();
    const selection = editor.selection();
    if (!shape?.table || !shape.capabilities.editTableCells || selection?.kind !== 'shape') return;
    const cellEditReason = selectedTableCellEditReason(shape);
    if (cellEditReason) {
      editStatus.textContent = cellEditReason;
      return;
    }
    const row = Math.trunc(readNumber(tableRowInput) ?? 0);
    const column = Math.trunc(readNumber(tableColumnInput) ?? 0);
    if (row < 0 || row >= shape.table.rowCount || column < 0 || column >= shape.table.columnCount) {
      editStatus.textContent = 'Cell out of range';
      return;
    }
    void dispatchSelectedShapeCommand({
      kind: 'set-table-cell-text',
      slideIdOrIndex: selection.slideId,
      shapeIdOrName: shape.id,
      row,
      column,
      text: tableCellTextInput.value,
    });
  });
  exportButton.addEventListener('click', async () => {
    const out = await editor.toPptx();
    downloadPptBytes(
      out,
      exportedPptFileName(fixtures[selectedIndex]?.name ?? 'presentation.pptx'),
    );
  });
  verifyExportButton.addEventListener('click', async () => {
    editStatus.textContent = '正在校验导出...';
    const verified = await editor.session().verifyExport();
    editStatus.textContent = `导出校验通过：${verified.slides.length} 页`;
  });
  render();
}

function renderPptInspectorPanel(options: {
  fixtureName: string;
  sceneSlideCount: number;
  sceneShapeCount: number;
  selection: ReturnType<NativePptEditor['selection']>;
  selectedShape: PptSceneNode | null;
  lastObjectAction: {
    action:
      | 'blocked-text-edit'
      | 'cycle-selection'
      | 'nudge-geometry'
      | 'snap-geometry'
      | 'copy-shape'
      | 'paste-shape'
      | 'delete-shape'
      | 'duplicate-shape'
      | 'bring-shape-to-front'
      | 'send-shape-to-back';
    slideId: string;
    shapeId: string;
    kind: PptSceneNode['kind'];
    name?: string;
    reason: string;
    candidateIndex?: number;
    candidateCount?: number;
  } | null;
  debug: unknown;
}): string {
  const selectionSlide =
    options.selection?.kind === 'slide' || options.selection?.kind === 'shape'
      ? options.selection.slideId
      : '无';
  const shape = options.selectedShape;
  const debugJson = escapeHtml(JSON.stringify(options.debug, null, 2));
  const hitHint = renderPptHitHint(options.lastObjectAction);
  if (!shape) {
    return `
      <div class="coco-ppt-object-title">未选择对象</div>
      <dl class="coco-ppt-object-meta">
        <dt>文档</dt><dd>${escapeHtml(options.fixtureName)}</dd>
        <dt>页数</dt><dd>${options.sceneSlideCount}</dd>
        <dt>对象数</dt><dd>${options.sceneShapeCount}</dd>
        <dt>当前页</dt><dd>${escapeHtml(selectionSlide)}</dd>
      </dl>
      <div class="coco-ppt-reason">请选择画布对象后查看可编辑能力。</div>
      ${hitHint}
      <details class="coco-ppt-json"><summary>调试 JSON</summary><pre>${debugJson}</pre></details>
    `;
  }
  const capabilities = shape.capabilities;
  const reason = capabilities.reason || pptObjectCapabilityReason(shape);
  const textPreview = shape.text ? escapeHtml(shape.text) : '无文本内容';
  const parentGroup = shape.groupContext
    ? `
      <dt>父组</dt><dd>${escapeHtml(shape.groupContext.parentName || shape.groupContext.parentId)}</dd>
      <dt>父组ID</dt><dd>${escapeHtml(shape.groupContext.parentId)}</dd>
    `
    : '';
  return `
    <div class="coco-ppt-object-title">${escapeHtml(shape.name || shape.id)}</div>
    <dl class="coco-ppt-object-meta">
      <dt>ID</dt><dd>${escapeHtml(shape.id)}</dd>
      <dt>类型</dt><dd>${escapeHtml(shape.kind)}</dd>
      ${parentGroup}
      <dt>选中页</dt><dd>${escapeHtml(selectionSlide)}</dd>
      <dt>可编辑</dt><dd>${shape.editable ? '是' : '否'}</dd>
      <dt>锁定</dt><dd>${shape.locked ? '是' : '否'}</dd>
    </dl>
    <div class="coco-ppt-capability-grid">
      ${pptCapabilityBadge('选择', capabilities.select)}
      ${pptCapabilityBadge('文字', capabilities.editText, capabilities.reasons?.editText)}
      ${pptCapabilityBadge('文本样式', capabilities.editTextStyle, capabilities.reasons?.editTextStyle)}
      ${pptCapabilityBadge('表格单元格', capabilities.editTableCells, capabilities.reasons?.editTableCells)}
      ${pptCapabilityBadge('位置尺寸', capabilities.editGeometry, capabilities.reasons?.editGeometry)}
      ${pptCapabilityBadge('形状样式', capabilities.editShapeStyle, capabilities.reasons?.editShapeStyle)}
    </div>
    ${reason ? `<div class="coco-ppt-reason">${escapeHtml(reason)}</div>` : ''}
    ${hitHint}
    <div class="coco-ppt-text-preview">${textPreview}</div>
    <details class="coco-ppt-json"><summary>调试 JSON</summary><pre>${debugJson}</pre></details>
  `;
}

function pptCapabilityBadge(label: string, enabled: boolean, reason?: string): string {
  const reasonMarkup =
    !enabled && reason
      ? `<span class="coco-ppt-capability-reason">${escapeHtml(reason)}</span>`
      : '';
  return `<div class="coco-ppt-capability" data-enabled="${enabled}">${escapeHtml(label)}：${enabled ? '可用' : '不可用'}${reasonMarkup}</div>`;
}

function renderPptHitHint(
  action: {
    action:
      | 'blocked-text-edit'
      | 'cycle-selection'
      | 'nudge-geometry'
      | 'snap-geometry'
      | 'copy-shape'
      | 'paste-shape'
      | 'delete-shape'
      | 'duplicate-shape'
      | 'bring-shape-to-front'
      | 'send-shape-to-back';
    slideId: string;
    shapeId: string;
    kind: PptSceneNode['kind'];
    name?: string;
    reason: string;
    candidateIndex?: number;
    candidateCount?: number;
  } | null,
): string {
  if (!action) return '';
  const target = `${action.name || action.shapeId} (${action.kind})`;
  const cycleText =
    action.action === 'cycle-selection' && action.candidateCount !== undefined
      ? `重叠候选：${(action.candidateIndex ?? 0) + 1}/${action.candidateCount}`
      : action.action === 'blocked-text-edit'
        ? '文字编辑受限'
        : '对象操作';
  return `
    <div class="coco-ppt-hit-hint">
      <strong>${action.action === 'nudge-geometry' ? '最近操作' : '最近命中'}</strong><br />
      ${escapeHtml(cycleText)}<br />
      ${escapeHtml(target)}<br />
      ${escapeHtml(action.reason)}
    </div>
  `;
}

function pptObjectCapabilityReason(shape: PptSceneNode): string {
  if (shape.capabilities.editText) return '';
  if (shape.kind === 'table') return '表格请使用单元格编辑';
  if (shape.kind === 'picture') return '图片对象不支持文字编辑';
  if (shape.kind === 'chart') return '图表当前仅支持选择和位置编辑';
  if (shape.kind === 'smartArt') return 'SmartArt 当前仅支持选择和位置编辑';
  if (shape.kind === 'group') return '组合对象当前仅支持选择和位置编辑';
  return '该对象不支持文字编辑';
}

async function renderSpreadsheetPublicDev(app: HTMLDivElement): Promise<void> {
  app.className = 'coco-spreadsheet-public-dev';
  app.innerHTML = `
    <style>
      .coco-spreadsheet-public-dev { height: 100vh; display: flex; flex-direction: column; background: #f8fafc; color: #111827; font-family: Arial, sans-serif; }
      .coco-spreadsheet-public-toolbar { display: flex; align-items: center; gap: 8px; min-height: 42px; padding: 6px 10px; background: #fff; border-bottom: 1px solid #d1d5db; }
      .coco-spreadsheet-public-toolbar select, .coco-spreadsheet-public-toolbar button { height: 28px; border: 1px solid #cbd5e1; border-radius: 4px; background: #fff; color: #111827; }
      .coco-spreadsheet-public-status { color: #475569; font-size: 12px; }
      .coco-spreadsheet-public-host { flex: 1; min-height: 0; margin: 10px; border: 1px solid #cbd5e1; background: #fff; }
    </style>
    <div class="coco-spreadsheet-public-toolbar">
      <select id="publicSpreadsheetFixture" aria-label="Spreadsheet fixture"></select>
      <button type="button" id="publicSpreadsheetExport">Export XLSX</button>
      <span id="publicSpreadsheetStatus" class="coco-spreadsheet-public-status">Loading...</span>
    </div>
    <div id="publicSpreadsheetHost" class="coco-spreadsheet-public-host"></div>
  `;

  const fixtureSelect = app.querySelector<HTMLSelectElement>('#publicSpreadsheetFixture');
  const exportButton = app.querySelector<HTMLButtonElement>('#publicSpreadsheetExport');
  const status = app.querySelector<HTMLSpanElement>('#publicSpreadsheetStatus');
  const host = app.querySelector<HTMLDivElement>('#publicSpreadsheetHost');
  if (!fixtureSelect || !exportButton || !status || !host) return;

  const response = await fetch('/__coco_fixture__/xlsx-files');
  const fixtures = (await response.json()) as FixtureFile[];
  const selectedIndex = selectedSpreadsheetFixtureIndex(fixtures);
  fixtureSelect.innerHTML = fixtures
    .map(
      (fixture) =>
        `<option value="${fixture.index}"${fixture.index === selectedIndex ? ' selected' : ''}>${fixture.index + 1}. ${fixture.name}</option>`,
    )
    .join('');

  let editor: SpreadsheetEditor | null = null;
  let unsubscribeState: (() => void) | null = null;
  let latestBytes: Uint8Array | null = null;

  const load = async (index: number): Promise<void> => {
    unsubscribeState?.();
    editor?.destroy();
    editor = null;
    const fileResponse = await fetch(`/__coco_fixture__/xlsx?index=${index}`);
    latestBytes = new Uint8Array(await fileResponse.arrayBuffer());
    editor = await renderSpreadsheetEditor(host, latestBytes, {
      workbookId: `editor-public-fixture-${index}`,
      readOnly: selectedReadOnly(),
      showSheetTabs: true,
      styleIsolation: 'shadow',
      enableHeaderFilters: false,
      onChange({ bytes }) {
        latestBytes = bytes;
        status.textContent = 'Changed';
      },
    });
    unsubscribeState = editor.subscribeState(({ activeSheetName, selection, editingCell }) => {
      const address = `${selection.active.row}:${selection.active.column}`;
      status.textContent = `${activeSheetName} · ${address}${editingCell ? ' · editing' : ''}`;
    });
  };

  fixtureSelect.addEventListener('change', () => {
    const params = new URLSearchParams(location.search);
    params.set('surface', 'spreadsheet-public');
    params.set('xlsxFixture', fixtureSelect.value);
    location.search = `?${params.toString()}`;
  });
  exportButton.addEventListener('click', async () => {
    if (!editor) return;
    latestBytes = await editor.toZip();
    const blob = new Blob([latestBytes.slice().buffer as ArrayBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'spreadsheet-public-test.xlsx';
    link.click();
    URL.revokeObjectURL(link.href);
  });
  await load(selectedIndex);
}

async function renderSpreadsheetDev(app: HTMLDivElement): Promise<void> {
  app.className = 'coco-spreadsheet-dev';
  app.innerHTML = `
    <style>
      .coco-spreadsheet-dev {
        height: 100vh;
        background: #f8fafc;
        color: #111827;
        font-family: Arial, sans-serif;
        overflow: hidden;
      }
      .coco-spreadsheet-shell {
        display: grid;
        grid-template-rows: auto auto auto 1fr;
        height: 100%;
      }
      .coco-spreadsheet-devbar {
        display: flex;
        align-items: center;
        gap: 8px;
        min-height: 42px;
        padding: 6px 10px;
        border-bottom: 1px solid #d1d5db;
        background: #ffffff;
      }
      .coco-spreadsheet-devbar select,
      .coco-spreadsheet-devbar button {
        height: 28px;
        border: 1px solid #cbd5e1;
        background: #ffffff;
        border-radius: 4px;
        color: #111827;
      }
      .coco-spreadsheet-summary {
        color: #475569;
        font-size: 12px;
      }
      .coco-spreadsheet-export-status {
        min-width: 120px;
        color: #64748b;
        font-size: 12px;
        white-space: nowrap;
      }
      .coco-spreadsheet-export-status.is-saving {
        color: #1d4ed8;
      }
      .coco-spreadsheet-export-status.is-ok {
        color: #047857;
      }
      .coco-spreadsheet-export-status.is-error {
        color: #b91c1c;
      }
      .coco-spreadsheet-formula-bar {
        display: grid;
        grid-template-columns: 92px minmax(0, 1fr);
        gap: 8px;
        align-items: center;
        min-height: 36px;
        padding: 4px 10px;
        border-bottom: 1px solid #d1d5db;
        background: #ffffff;
      }
      .coco-spreadsheet-name-box,
      .coco-spreadsheet-formula-input {
        width: 100%;
        min-width: 0;
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        padding: 0 8px;
        color: #111827;
        background: #ffffff;
        font: 13px Arial, sans-serif;
        box-sizing: border-box;
      }
      .coco-spreadsheet-findbar {
        display: none;
        grid-template-columns: minmax(160px, 260px) minmax(120px, 220px) auto auto auto auto;
        gap: 6px;
        align-items: center;
        min-height: 34px;
        padding: 4px 10px;
        border-bottom: 1px solid #d1d5db;
        background: #f8fafc;
      }
      .coco-spreadsheet-findbar.is-open {
        display: grid;
      }
      .coco-spreadsheet-findbar input {
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        padding: 0 8px;
        font: 13px Arial, sans-serif;
      }
      .coco-spreadsheet-findbar button {
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
      }
      .coco-spreadsheet-find-count {
        color: #475569;
        font-size: 12px;
        min-width: 64px;
      }
      .coco-spreadsheet-name-box {
        text-transform: uppercase;
      }
      .coco-spreadsheet-toolbar {
        display: flex;
        align-items: center;
        gap: 4px;
        min-height: 34px;
        padding: 4px 10px;
        border-bottom: 1px solid #d1d5db;
        background: #ffffff;
      }
      .coco-spreadsheet-toolbar button {
        min-width: 28px;
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
        color: #111827;
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-toolbar button:hover {
        background: #f8fafc;
      }
      .coco-spreadsheet-toolbar button.is-active {
        border-color: #2563eb;
        background: #eff6ff;
        color: #1d4ed8;
      }
      .coco-spreadsheet-toolbar button.is-mixed,
      .coco-spreadsheet-toolbar input.is-mixed,
      .coco-spreadsheet-toolbar select.is-mixed,
      .coco-spreadsheet-cell-settings button.is-mixed,
      .coco-spreadsheet-cell-settings input.is-mixed,
      .coco-spreadsheet-cell-settings select.is-mixed {
        border-style: dashed;
        border-color: #2563eb;
      }
      .coco-spreadsheet-toolbar input[type="color"] {
        width: 30px;
        height: 26px;
        padding: 2px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
      }
      .coco-spreadsheet-toolbar select {
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
        color: #111827;
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-toolbar-divider {
        width: 1px;
        height: 20px;
        margin: 0 4px;
        background: #e5e7eb;
      }
      .coco-spreadsheet-main {
        display: grid;
        grid-template-columns: minmax(0, 1fr) 320px;
        min-height: 0;
      }
      .coco-spreadsheet-work-area {
        display: grid;
        grid-template-rows: minmax(0, 1fr) auto;
        min-width: 0;
        min-height: 0;
      }
      .coco-spreadsheet-canvas-host {
        position: relative;
        min-width: 0;
        min-height: 0;
        background: #ffffff;
        overflow: hidden;
      }
      .coco-spreadsheet-canvas-host:focus {
        outline: none;
      }
      .coco-spreadsheet-canvas-host canvas,
      .coco-spreadsheet-cell-editor {
        position: absolute;
        inset: 0;
      }
      .coco-spreadsheet-canvas-host canvas {
        display: block;
      }
      .coco-spreadsheet-resize-guide {
        position: absolute;
        display: none;
        z-index: 6;
        pointer-events: none;
        background: #2563eb;
      }
      .coco-spreadsheet-resize-guide.is-column {
        top: 0;
        bottom: 12px;
        width: 2px;
      }
      .coco-spreadsheet-resize-guide.is-row {
        left: 0;
        right: 12px;
        height: 2px;
      }
      .coco-spreadsheet-cell-editor {
        display: none;
        z-index: 4;
        resize: none;
        border: 2px solid #2563eb;
        padding: 2px 5px;
        font: 12px Arial;
        outline: none;
        box-sizing: border-box;
      }
      .coco-spreadsheet-formula-suggest {
        position: fixed;
        display: none;
        z-index: 20;
        width: 280px;
        max-height: 230px;
        overflow: auto;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #ffffff;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-formula-suggest button {
        display: block;
        width: 100%;
        border: 0;
        padding: 7px 9px;
        background: transparent;
        color: #111827;
        text-align: left;
      }
      .coco-spreadsheet-formula-suggest button.is-active,
      .coco-spreadsheet-formula-suggest button:hover {
        background: #eff6ff;
      }
      .coco-spreadsheet-formula-suggest-name {
        font-weight: 700;
        color: #1d4ed8;
      }
      .coco-spreadsheet-formula-suggest-signature {
        margin-top: 2px;
        color: #64748b;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .coco-spreadsheet-context-menu {
        position: fixed;
        display: none;
        z-index: 30;
        min-width: 176px;
        padding: 4px;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #ffffff;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-context-menu button {
        display: block;
        width: 100%;
        border: 0;
        border-radius: 4px;
        padding: 7px 9px;
        background: transparent;
        color: #111827;
        text-align: left;
      }
      .coco-spreadsheet-context-menu button:hover {
        background: #eff6ff;
      }
      .coco-spreadsheet-filter-menu-title {
        padding: 6px 9px 5px;
        border-bottom: 1px solid #e5e7eb;
        color: #64748b;
        font-weight: 700;
      }
      .coco-spreadsheet-filter-menu-values {
        max-height: 280px;
        overflow-y: auto;
        padding: 4px;
      }
      .coco-spreadsheet-filter-choice {
        display: flex;
        align-items: center;
        gap: 7px;
        padding: 6px 5px;
        color: #111827;
        white-space: nowrap;
      }
      .coco-spreadsheet-filter-choice:hover {
        background: #eff6ff;
      }
      .coco-spreadsheet-filter-menu-actions {
        display: flex;
        gap: 4px;
        padding: 4px;
        border-top: 1px solid #e5e7eb;
      }
      .coco-spreadsheet-filter-menu-actions button {
        flex: 1;
        text-align: center;
      }
      .coco-spreadsheet-context-menu button.is-danger:hover {
        background: #fef2f2;
        color: #b91c1c;
      }
      .coco-spreadsheet-context-menu-separator {
        height: 1px;
        margin: 4px;
        background: #e5e7eb;
      }
      .coco-spreadsheet-confirm {
        position: fixed;
        inset: 0;
        z-index: 60;
        display: grid;
        place-items: center;
        padding: 20px;
        background: rgba(15, 23, 42, 0.32);
      }
      .coco-spreadsheet-confirm[hidden] {
        display: none;
      }
      .coco-spreadsheet-confirm-dialog {
        width: min(360px, 100%);
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        background: #ffffff;
        box-shadow: 0 20px 48px rgba(15, 23, 42, 0.24);
        overflow: hidden;
      }
      .coco-spreadsheet-confirm-body {
        padding: 18px;
      }
      .coco-spreadsheet-confirm-title {
        color: #111827;
        font: 600 15px Arial, sans-serif;
      }
      .coco-spreadsheet-confirm-message {
        margin: 8px 0 0;
        color: #475569;
        font: 13px/1.45 Arial, sans-serif;
        overflow-wrap: anywhere;
      }
      .coco-spreadsheet-confirm-actions {
        display: flex;
        justify-content: flex-end;
        gap: 8px;
        padding: 12px 18px;
        border-top: 1px solid #e2e8f0;
        background: #f8fafc;
      }
      .coco-spreadsheet-confirm-actions button {
        height: 30px;
        padding: 0 12px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
        color: #334155;
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-confirm-actions button.is-danger {
        border-color: #dc2626;
        background: #dc2626;
        color: #ffffff;
      }
      .coco-spreadsheet-cell-settings {
        position: fixed;
        display: none;
        z-index: 25;
        right: 336px;
        top: 126px;
        width: 260px;
        max-height: calc(100vh - 150px);
        overflow: auto;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        background: #ffffff;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-cell-settings-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 8px 10px;
        border-bottom: 1px solid #e5e7eb;
        font-weight: 700;
      }
      .coco-spreadsheet-cell-settings-body {
        display: grid;
        gap: 10px;
        padding: 10px;
      }
      .coco-spreadsheet-cell-settings-group {
        display: grid;
        gap: 6px;
      }
      .coco-spreadsheet-cell-settings-label {
        color: #64748b;
        font-size: 11px;
      }
      .coco-spreadsheet-cell-settings-row {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      .coco-spreadsheet-cell-settings button,
      .coco-spreadsheet-cell-settings select {
        height: 26px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
        color: #111827;
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-cell-settings button {
        min-width: 28px;
        padding: 0 8px;
      }
      .coco-spreadsheet-cell-settings button.is-active {
        border-color: #2563eb;
        background: #eff6ff;
        color: #1d4ed8;
      }
      .coco-spreadsheet-cell-settings input[type="color"] {
        width: 30px;
        height: 26px;
        padding: 2px;
        border: 1px solid #cbd5e1;
        border-radius: 4px;
        background: #ffffff;
      }
      .coco-spreadsheet-cell-settings select {
        width: 100%;
      }
      .coco-spreadsheet-scrollbar {
        position: absolute;
        z-index: 5;
        background: #f1f5f9;
        border: 1px solid #cbd5e1;
        box-sizing: border-box;
        user-select: none;
      }
      .coco-spreadsheet-scrollbar.is-horizontal {
        left: 48px;
        right: 12px;
        bottom: 0;
        height: 12px;
      }
      .coco-spreadsheet-scrollbar.is-vertical {
        top: 24px;
        right: 0;
        bottom: 12px;
        width: 12px;
      }
      .coco-spreadsheet-scrollbar-thumb {
        position: absolute;
        border-radius: 999px;
        background: #94a3b8;
      }
      .coco-spreadsheet-scrollbar-thumb:hover,
      .coco-spreadsheet-scrollbar-thumb.is-dragging {
        background: #64748b;
      }
      .coco-spreadsheet-tabs {
        display: flex;
        align-items: center;
        gap: 2px;
        min-height: 34px;
        padding: 4px 8px 0;
        border-top: 1px solid #d1d5db;
        background: #f8fafc;
        overflow-x: auto;
      }
      .coco-spreadsheet-tab {
        height: 29px;
        padding: 0 14px;
        border: 1px solid #cbd5e1;
        border-bottom: none;
        border-radius: 6px 6px 0 0;
        background: #e5e7eb;
        color: #334155;
        font-size: 12px;
        white-space: nowrap;
      }
      .coco-spreadsheet-tab.is-active {
        background: #ffffff;
        color: #0f172a;
        font-weight: 600;
      }
      .coco-spreadsheet-tab.is-add {
        min-width: 29px;
        padding: 0;
        border-bottom: 1px solid #cbd5e1;
        border-radius: 4px;
      }
      .coco-spreadsheet-tab-rename {
        box-sizing: border-box;
        width: 120px;
        height: 29px;
        padding: 0 8px;
        border: 1px solid #2563eb;
        border-bottom: none;
        border-radius: 6px 6px 0 0;
        outline: none;
        background: #ffffff;
        color: #0f172a;
        font: 12px Arial, sans-serif;
      }
      .coco-spreadsheet-inspector {
        min-width: 0;
        border-left: 1px solid #d1d5db;
        background: #ffffff;
        overflow: auto;
      }
      .coco-spreadsheet-inspector pre {
        margin: 0;
        padding: 10px;
        font-size: 11px;
        line-height: 1.45;
        white-space: pre-wrap;
      }
    </style>
    <div class="coco-spreadsheet-shell">
      <div class="coco-spreadsheet-devbar">
        <select id="spreadsheetFixtureSelect" aria-label="Spreadsheet fixture"></select>
        <select id="surfaceSelect" aria-label="Editor surface">
          <option value="word">Word</option>
          <option value="spreadsheet" selected>Spreadsheet</option>
        </select>
        <button type="button" id="exportXlsxButton">Export XLSX</button>
        <span id="spreadsheetExportStatus" class="coco-spreadsheet-export-status" aria-live="polite"></span>
        <span id="spreadsheetSummary" class="coco-spreadsheet-summary">Loading spreadsheet...</span>
      </div>
      <div class="coco-spreadsheet-formula-bar">
        <input id="spreadsheetNameBox" class="coco-spreadsheet-name-box" aria-label="Active cell address" />
        <input id="spreadsheetFormulaInput" class="coco-spreadsheet-formula-input" aria-label="Formula bar" />
      </div>
      <div id="spreadsheetFindBar" class="coco-spreadsheet-findbar">
        <input id="spreadsheetFindInput" type="search" placeholder="Find" aria-label="Find cell text" />
        <input id="spreadsheetReplaceInput" type="text" placeholder="Replace with" aria-label="Replace cell text" />
        <span id="spreadsheetFindCount" class="coco-spreadsheet-find-count">0 results</span>
        <button type="button" id="spreadsheetFindPreviousButton" title="Previous result">Prev</button>
        <button type="button" id="spreadsheetFindNextButton" title="Next result">Next</button>
        <button type="button" id="spreadsheetReplaceAllButton">Replace all</button>
        <button type="button" id="spreadsheetFindCloseButton" title="Close find">×</button>
      </div>
      <div class="coco-spreadsheet-toolbar" aria-label="Spreadsheet formatting toolbar">
        <button type="button" id="spreadsheetUndoButton" title="Undo">Undo</button>
        <button type="button" id="spreadsheetRedoButton" title="Redo">Redo</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" id="spreadsheetBoldButton" title="Bold"><strong>B</strong></button>
        <button type="button" id="spreadsheetItalicButton" title="Italic"><em>I</em></button>
        <button type="button" id="spreadsheetUnderlineButton" title="Underline"><u>U</u></button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <input id="spreadsheetFontColorInput" type="color" value="#111827" title="Font color" aria-label="Font color" />
        <input id="spreadsheetFillColorInput" type="color" value="#FFFFFF" title="Fill color" aria-label="Fill color" />
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <select id="spreadsheetNumberFormatSelect" title="Number format" aria-label="Number format">
          <option value="">General</option>
          <option value="#,##0.00">Number</option>
          <option value="0.00%">Percent</option>
          <option value="yyyy-mm-dd">Date</option>
          <option value="$#,##0.00">Currency</option>
        </select>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" data-spreadsheet-align="left" title="Align left">L</button>
        <button type="button" data-spreadsheet-align="center" title="Align center">C</button>
        <button type="button" data-spreadsheet-align="right" title="Align right">R</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" data-spreadsheet-vertical-align="top" title="Align top">T</button>
        <button type="button" data-spreadsheet-vertical-align="center" title="Align middle">M</button>
        <button type="button" data-spreadsheet-vertical-align="bottom" title="Align bottom">B</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" id="spreadsheetMergeButton" title="Merge cells">Merge</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <select id="spreadsheetFreezeSelect" title="Freeze panes" aria-label="Freeze panes">
          <option value="">Freeze</option>
          <option value="panes">Freeze panes</option>
          <option value="rows">Freeze rows above</option>
          <option value="columns">Freeze columns left</option>
          <option value="none">Unfreeze panes</option>
        </select>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <select id="spreadsheetBorderSelect" title="Borders" aria-label="Borders">
          <option value="">Borders</option>
          <option value="all">All borders</option>
          <option value="outer">Outer border</option>
          <option value="inner">Inner borders</option>
          <option value="clear">Clear borders</option>
        </select>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <label class="coco-spreadsheet-filter-toggle" title="Enable or disable header filters">
          <input id="spreadsheetFilterToggle" type="checkbox" />
          Filter
        </label>
        <label class="coco-spreadsheet-filter-toggle" title="Treat the first data row as column headers">
          <input id="spreadsheetFilterHeaderToggle" type="checkbox" checked />
          Headers
        </label>
        <select id="spreadsheetFilterValueSelect" title="Filter active column" aria-label="Filter active column">
          <option value="">Filter column</option>
        </select>
        <button type="button" id="spreadsheetApplyFilterButton" title="Apply filter">Filter</button>
        <button type="button" id="spreadsheetClearFilterButton" title="Clear filter">Clear filter</button>
        <button type="button" id="spreadsheetClearAllFiltersButton" title="Clear all filters">Clear all filters</button>
        <button type="button" id="spreadsheetSortAscendingButton" title="Sort active column ascending">A-Z</button>
        <button type="button" id="spreadsheetSortDescendingButton" title="Sort active column descending">Z-A</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" id="spreadsheetClearFormatButton" title="Clear format">Tx</button>
        <span class="coco-spreadsheet-toolbar-divider"></span>
        <button type="button" id="spreadsheetAgentPreviewButton" title="Preview agent edit">Preview Agent</button>
        <button type="button" id="spreadsheetAgentApplyButton" title="Apply last agent preview" disabled>Apply Preview</button>
      </div>
      <div id="spreadsheetFormulaSuggest" class="coco-spreadsheet-formula-suggest" role="listbox"></div>
      <div id="spreadsheetContextMenu" class="coco-spreadsheet-context-menu" role="menu"></div>
      <div id="spreadsheetConfirm" class="coco-spreadsheet-confirm" hidden>
        <section class="coco-spreadsheet-confirm-dialog" role="dialog" aria-modal="true" aria-labelledby="spreadsheetConfirmTitle">
          <div class="coco-spreadsheet-confirm-body">
            <div id="spreadsheetConfirmTitle" class="coco-spreadsheet-confirm-title">Confirm action</div>
            <p id="spreadsheetConfirmMessage" class="coco-spreadsheet-confirm-message"></p>
          </div>
          <div class="coco-spreadsheet-confirm-actions">
            <button type="button" id="spreadsheetConfirmCancel">Cancel</button>
            <button type="button" id="spreadsheetConfirmAccept" class="is-danger">Delete</button>
          </div>
        </section>
      </div>
      <section id="spreadsheetCellSettings" class="coco-spreadsheet-cell-settings" aria-label="Cell settings">
        <div class="coco-spreadsheet-cell-settings-header">
          <span id="spreadsheetCellSettingsTitle">Cell settings</span>
          <button type="button" id="spreadsheetCellSettingsClose" title="Close">x</button>
        </div>
        <div class="coco-spreadsheet-cell-settings-body">
          <div class="coco-spreadsheet-cell-settings-group">
            <div class="coco-spreadsheet-cell-settings-label">Font</div>
            <div class="coco-spreadsheet-cell-settings-row">
              <button type="button" data-cell-settings-action="bold"><strong>B</strong></button>
              <button type="button" data-cell-settings-action="italic"><em>I</em></button>
              <button type="button" data-cell-settings-action="underline"><u>U</u></button>
              <input id="spreadsheetCellSettingsFontColor" type="color" value="#111827" aria-label="Font color" />
              <input id="spreadsheetCellSettingsFillColor" type="color" value="#FFFFFF" aria-label="Fill color" />
            </div>
          </div>
          <div class="coco-spreadsheet-cell-settings-group">
            <div class="coco-spreadsheet-cell-settings-label">Number format</div>
            <select id="spreadsheetCellSettingsNumberFormat">
              <option value="">General</option>
              <option value="#,##0.00">Number</option>
              <option value="0.00%">Percent</option>
              <option value="yyyy-mm-dd">Date</option>
              <option value="$#,##0.00">Currency</option>
            </select>
          </div>
          <div class="coco-spreadsheet-cell-settings-group">
            <div class="coco-spreadsheet-cell-settings-label">Alignment</div>
            <div class="coco-spreadsheet-cell-settings-label">Horizontal</div>
            <div class="coco-spreadsheet-cell-settings-row">
              <button type="button" data-cell-settings-align="left">Left</button>
              <button type="button" data-cell-settings-align="center">Center</button>
              <button type="button" data-cell-settings-align="right">Right</button>
            </div>
            <div class="coco-spreadsheet-cell-settings-label">Vertical</div>
            <div class="coco-spreadsheet-cell-settings-row">
              <button type="button" data-cell-settings-vertical-align="top">Top</button>
              <button type="button" data-cell-settings-vertical-align="center">Middle</button>
              <button type="button" data-cell-settings-vertical-align="bottom">Bottom</button>
            </div>
          </div>
          <div class="coco-spreadsheet-cell-settings-group">
            <div class="coco-spreadsheet-cell-settings-label">Borders</div>
            <select id="spreadsheetCellSettingsBorder">
              <option value="">Choose border</option>
              <option value="all">All borders</option>
              <option value="outer">Outer border</option>
              <option value="inner">Inner borders</option>
              <option value="clear">Clear borders</option>
            </select>
          </div>
          <div class="coco-spreadsheet-cell-settings-group">
            <div class="coco-spreadsheet-cell-settings-label">Structure</div>
            <div class="coco-spreadsheet-cell-settings-row">
              <button type="button" data-cell-settings-action="merge">Merge</button>
              <button type="button" data-cell-settings-action="unmerge">Unmerge</button>
            </div>
          </div>
          <button type="button" data-cell-settings-action="clear-format">Clear format</button>
        </div>
      </section>
      <main class="coco-spreadsheet-main">
        <div class="coco-spreadsheet-work-area">
          <section id="spreadsheetCanvasHost" class="coco-spreadsheet-canvas-host" tabindex="0">
            <canvas id="spreadsheetBaseCanvas"></canvas>
            <canvas id="spreadsheetHeaderCanvas"></canvas>
            <canvas id="spreadsheetOverlayCanvas"></canvas>
            <div id="spreadsheetResizeGuide" class="coco-spreadsheet-resize-guide"></div>
            <textarea id="spreadsheetCellEditor" class="coco-spreadsheet-cell-editor" spellcheck="false"></textarea>
            <div id="spreadsheetHorizontalScrollbar" class="coco-spreadsheet-scrollbar is-horizontal">
              <div id="spreadsheetHorizontalThumb" class="coco-spreadsheet-scrollbar-thumb"></div>
            </div>
            <div id="spreadsheetVerticalScrollbar" class="coco-spreadsheet-scrollbar is-vertical">
              <div id="spreadsheetVerticalThumb" class="coco-spreadsheet-scrollbar-thumb"></div>
            </div>
          </section>
          <nav id="spreadsheetTabs" class="coco-spreadsheet-tabs" aria-label="Workbook sheets"></nav>
        </div>
        <aside class="coco-spreadsheet-inspector">
          <pre id="spreadsheetInspector">Loading...</pre>
        </aside>
      </main>
    </div>
  `;

  const fixtures = await loadSpreadsheetFixtures();
  const selectedIndex = selectedSpreadsheetFixtureIndex(fixtures);
  renderSpreadsheetFixtureSelect(fixtures, selectedIndex);
  wireSurfaceSelect();

  const response = await fetch(`/__coco_fixture__/xlsx?index=${selectedIndex}`);
  if (!response.ok) throw new Error(`Failed to load spreadsheet fixture: ${response.status}`);
  const fixturePath = decodeURIComponent(response.headers.get('X-Coco-Fixture-Path') ?? '');
  const bytes = new Uint8Array(await response.arrayBuffer());
  const runtime = await SpreadsheetEditorRuntime.fromZip(bytes, { mode: 'inline' });
  const host = document.querySelector<HTMLDivElement>('#spreadsheetCanvasHost');
  const baseCanvas = document.querySelector<HTMLCanvasElement>('#spreadsheetBaseCanvas');
  const headerCanvas = document.querySelector<HTMLCanvasElement>('#spreadsheetHeaderCanvas');
  const overlayCanvas = document.querySelector<HTMLCanvasElement>('#spreadsheetOverlayCanvas');
  const inspector = document.querySelector<HTMLPreElement>('#spreadsheetInspector');
  const summary = document.querySelector<HTMLSpanElement>('#spreadsheetSummary');
  const nameBox = document.querySelector<HTMLInputElement>('#spreadsheetNameBox');
  const formulaInput = document.querySelector<HTMLInputElement>('#spreadsheetFormulaInput');
  const findBar = document.querySelector<HTMLElement>('#spreadsheetFindBar');
  const findInput = document.querySelector<HTMLInputElement>('#spreadsheetFindInput');
  const replaceInput = document.querySelector<HTMLInputElement>('#spreadsheetReplaceInput');
  const findCount = document.querySelector<HTMLElement>('#spreadsheetFindCount');
  const findPreviousButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetFindPreviousButton',
  );
  const findNextButton = document.querySelector<HTMLButtonElement>('#spreadsheetFindNextButton');
  const replaceAllButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetReplaceAllButton',
  );
  const findCloseButton = document.querySelector<HTMLButtonElement>('#spreadsheetFindCloseButton');
  const filterValueSelect = document.querySelector<HTMLSelectElement>(
    '#spreadsheetFilterValueSelect',
  );
  const filterToggle = document.querySelector<HTMLInputElement>('#spreadsheetFilterToggle');
  const filterHeaderToggle = document.querySelector<HTMLInputElement>(
    '#spreadsheetFilterHeaderToggle',
  );
  const applyFilterButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetApplyFilterButton',
  );
  const clearFilterButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetClearFilterButton',
  );
  const clearAllFiltersButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetClearAllFiltersButton',
  );
  const sortAscendingButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetSortAscendingButton',
  );
  const sortDescendingButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetSortDescendingButton',
  );
  const editor = document.querySelector<HTMLTextAreaElement>('#spreadsheetCellEditor');
  const resizeGuide = document.querySelector<HTMLDivElement>('#spreadsheetResizeGuide');
  const horizontalScrollbar = document.querySelector<HTMLDivElement>(
    '#spreadsheetHorizontalScrollbar',
  );
  const verticalScrollbar = document.querySelector<HTMLDivElement>('#spreadsheetVerticalScrollbar');
  const horizontalThumb = document.querySelector<HTMLDivElement>('#spreadsheetHorizontalThumb');
  const verticalThumb = document.querySelector<HTMLDivElement>('#spreadsheetVerticalThumb');
  const tabs = document.querySelector<HTMLElement>('#spreadsheetTabs');
  const formulaSuggest = document.querySelector<HTMLDivElement>('#spreadsheetFormulaSuggest');
  const contextMenu = document.querySelector<HTMLDivElement>('#spreadsheetContextMenu');
  const confirmDialog = document.querySelector<HTMLDivElement>('#spreadsheetConfirm');
  const confirmTitle = document.querySelector<HTMLElement>('#spreadsheetConfirmTitle');
  const confirmMessage = document.querySelector<HTMLElement>('#spreadsheetConfirmMessage');
  const confirmCancelButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetConfirmCancel',
  );
  const confirmAcceptButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetConfirmAccept',
  );
  const cellSettings = document.querySelector<HTMLElement>('#spreadsheetCellSettings');
  const cellSettingsTitle = document.querySelector<HTMLElement>('#spreadsheetCellSettingsTitle');
  const cellSettingsClose = document.querySelector<HTMLButtonElement>(
    '#spreadsheetCellSettingsClose',
  );
  const cellSettingsFontColor = document.querySelector<HTMLInputElement>(
    '#spreadsheetCellSettingsFontColor',
  );
  const cellSettingsFillColor = document.querySelector<HTMLInputElement>(
    '#spreadsheetCellSettingsFillColor',
  );
  const cellSettingsNumberFormat = document.querySelector<HTMLSelectElement>(
    '#spreadsheetCellSettingsNumberFormat',
  );
  const cellSettingsBorder = document.querySelector<HTMLSelectElement>(
    '#spreadsheetCellSettingsBorder',
  );
  const cellSettingsActionButtons = Array.from(
    document.querySelectorAll<HTMLButtonElement>('[data-cell-settings-action]'),
  );
  const cellSettingsAlignButtons = Array.from(
    document.querySelectorAll<HTMLButtonElement>('[data-cell-settings-align]'),
  );
  const undoButton = document.querySelector<HTMLButtonElement>('#spreadsheetUndoButton');
  const redoButton = document.querySelector<HTMLButtonElement>('#spreadsheetRedoButton');
  const boldButton = document.querySelector<HTMLButtonElement>('#spreadsheetBoldButton');
  const italicButton = document.querySelector<HTMLButtonElement>('#spreadsheetItalicButton');
  const underlineButton = document.querySelector<HTMLButtonElement>('#spreadsheetUnderlineButton');
  const fontColorInput = document.querySelector<HTMLInputElement>('#spreadsheetFontColorInput');
  const fillColorInput = document.querySelector<HTMLInputElement>('#spreadsheetFillColorInput');
  const numberFormatSelect = document.querySelector<HTMLSelectElement>(
    '#spreadsheetNumberFormatSelect',
  );
  const mergeButton = document.querySelector<HTMLButtonElement>('#spreadsheetMergeButton');
  const freezeSelect = document.querySelector<HTMLSelectElement>('#spreadsheetFreezeSelect');
  const borderSelect = document.querySelector<HTMLSelectElement>('#spreadsheetBorderSelect');
  const clearFormatButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetClearFormatButton',
  );
  const exportButton = document.querySelector<HTMLButtonElement>('#exportXlsxButton');
  const exportStatus = document.querySelector<HTMLElement>('#spreadsheetExportStatus');
  const agentPreviewButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetAgentPreviewButton',
  );
  const agentApplyButton = document.querySelector<HTMLButtonElement>(
    '#spreadsheetAgentApplyButton',
  );
  const alignButtons = Array.from(
    document.querySelectorAll<HTMLButtonElement>('[data-spreadsheet-align]'),
  );
  const verticalAlignButtons = Array.from(
    document.querySelectorAll<HTMLButtonElement>('[data-spreadsheet-vertical-align]'),
  );
  const cellSettingsVerticalAlignButtons = Array.from(
    document.querySelectorAll<HTMLButtonElement>('[data-cell-settings-vertical-align]'),
  );
  if (
    !host ||
    !baseCanvas ||
    !headerCanvas ||
    !overlayCanvas ||
    !inspector ||
    !nameBox ||
    !formulaInput ||
    !findBar ||
    !findInput ||
    !replaceInput ||
    !findCount ||
    !findPreviousButton ||
    !findNextButton ||
    !replaceAllButton ||
    !findCloseButton ||
    !filterValueSelect ||
    !filterToggle ||
    !filterHeaderToggle ||
    !applyFilterButton ||
    !clearFilterButton ||
    !clearAllFiltersButton ||
    !sortAscendingButton ||
    !sortDescendingButton ||
    !editor ||
    !resizeGuide ||
    !horizontalScrollbar ||
    !verticalScrollbar ||
    !horizontalThumb ||
    !verticalThumb ||
    !tabs ||
    !formulaSuggest ||
    !contextMenu ||
    !confirmDialog ||
    !confirmTitle ||
    !confirmMessage ||
    !confirmCancelButton ||
    !confirmAcceptButton ||
    !cellSettings ||
    !cellSettingsTitle ||
    !cellSettingsClose ||
    !cellSettingsFontColor ||
    !cellSettingsFillColor ||
    !cellSettingsNumberFormat ||
    !cellSettingsBorder ||
    !undoButton ||
    !redoButton ||
    !boldButton ||
    !italicButton ||
    !underlineButton ||
    !fontColorInput ||
    !fillColorInput ||
    !numberFormatSelect ||
    !mergeButton ||
    !freezeSelect ||
    !borderSelect ||
    !clearFormatButton ||
    !exportButton ||
    !exportStatus ||
    !agentPreviewButton ||
    !agentApplyButton
  )
    return;

  const renderer = new SpreadsheetCanvasRenderer({ baseCanvas, headerCanvas, overlayCanvas });
  let scroll = { x: 0, y: 0 };
  let selection: SpreadsheetSelectionState = createSelection(1, 1);
  const sheetViewStates = new Map<string, SpreadsheetSheetViewState>();
  let isSwitchingSheet = false;
  let dashedRange = null as ReturnType<typeof normalizeSelectionRange> | null;
  let formulaReferenceRanges: SpreadsheetFormulaReferenceRange[] = [];
  let formulaReferenceEdit: SpreadsheetFormulaReferenceEditState | null = null;
  let formulaSuggestState: SpreadsheetFormulaSuggestState | null = null;
  let findMatches: SpreadsheetFindMatch[] = [];
  let activeFindIndex = -1;
  let spreadsheetClipboardPayload: SpreadsheetClipboardPayload | null = null;
  let contextMenuState: SpreadsheetContextMenuState | null = null;
  let headerFilterState: { column: number; values: Set<string> } | null = null;
  // Filter conditions are stored per column so applying one filter keeps the others.
  let activeColumnFilters = new Map<number, string[]>();
  let headerFilterColumnScope: Set<number> | null = null;
  let sheetContextMenuState: SpreadsheetSheetContextMenuState | null = null;
  let renamingSheetName: string | null = null;
  let spreadsheetConfirmationResolve: ((confirmed: boolean) => void) | null = null;
  let resizeDrag: SpreadsheetResizeDragState | null = null;
  let autoFillDrag: SpreadsheetAutoFillDragState | null = null;
  let autoFillPreviewRange: SpreadsheetRange | null = null;
  let selectionDragKind: 'cell' | 'row' | 'column' | null = null;
  let editingCell: { row: number; column: number } | null = null;
  let isSelecting = false;
  let dragDidMove = false;
  let suppressNextClick = false;
  let agentPreviewState: SpreadsheetAgentPreviewDevState | null = null;
  let exportState: SpreadsheetExportDevState = { status: 'idle' };
  let layout = new SpreadsheetLayout(runtime.activeSheet());
  let viewportManager = new SpreadsheetViewportManager(layout, {
    viewportSize: spreadsheetSheetViewportSize(hostSize(host)),
    scroll,
    freeze: freezeStateFromSheet(runtime.activeSheet().source),
  });

  const syncFormattingToolbar = () => {
    undoButton.disabled = !runtime.canUndo();
    redoButton.disabled = !runtime.canRedo();
    exportButton.disabled = exportState.status === 'saving';
    agentApplyButton.disabled = !canApplySpreadsheetAgentPreview(agentPreviewState);
    syncSpreadsheetFormattingToolbar({
      runtime,
      selection,
      boldButton,
      italicButton,
      underlineButton,
      fontColorInput,
      fillColorInput,
      numberFormatSelect,
      mergeButton,
      alignButtons,
      verticalAlignButtons,
      cellSettings,
      cellSettingsTitle,
      cellSettingsFontColor,
      cellSettingsFillColor,
      cellSettingsNumberFormat,
      cellSettingsActionButtons,
      cellSettingsAlignButtons,
      cellSettingsVerticalAlignButtons,
    });
  };

  const clearSpreadsheetTransientUiState = () => {
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    dashedRange = null;
    formulaReferenceRanges = [];
    formulaReferenceEdit = null;
    formulaSuggestState = null;
    contextMenuState = null;
    sheetContextMenuState = null;
    renamingSheetName = null;
    resizeDrag = null;
    autoFillDrag = null;
    autoFillPreviewRange = null;
    selectionDragKind = null;
    isSelecting = false;
    dragDidMove = false;
    suppressNextClick = false;
    hideSpreadsheetFormulaSuggest(formulaSuggest);
    hideSpreadsheetContextMenu(contextMenu);
    hideSpreadsheetCellSettings(cellSettings);
    hideSpreadsheetResizeGuide(resizeGuide);
  };

  const settleSpreadsheetConfirmation = (confirmed: boolean): void => {
    const resolve = spreadsheetConfirmationResolve;
    spreadsheetConfirmationResolve = null;
    confirmDialog.hidden = true;
    resolve?.(confirmed);
  };

  const requestSpreadsheetConfirmation = (title: string, message: string): Promise<boolean> => {
    if (spreadsheetConfirmationResolve) settleSpreadsheetConfirmation(false);
    confirmTitle.textContent = title;
    confirmMessage.textContent = message;
    confirmDialog.hidden = false;
    confirmAcceptButton.focus();
    return new Promise((resolve) => {
      spreadsheetConfirmationResolve = resolve;
    });
  };

  const syncSpreadsheetFind = () => {
    findMatches = findSpreadsheetMatches(runtime.activeSheet(), { query: findInput.value });
    const selectedAddress = formatCellAddress(selection.active.row, selection.active.column);
    activeFindIndex = findMatches.findIndex((match) => match.address === selectedAddress);
    renderSpreadsheetFindCount(findCount, findMatches, activeFindIndex);
    findPreviousButton.disabled = findMatches.length === 0;
    findNextButton.disabled = findMatches.length === 0;
  };

  const spreadsheetFilterRange = (): SpreadsheetRange => {
    const selectedRange = normalizeSelectionRange(selection.range);
    if (selectedRange.endRow > selectedRange.startRow) return selectedRange;
    return {
      startRow: 1,
      startColumn: 1,
      endRow: runtime.activeSheet().maxRow,
      endColumn: runtime.activeSheet().maxColumn,
    };
  };

  const syncSpreadsheetFilter = () => {
    const range = spreadsheetFilterRange();
    const values =
      range.endRow > range.startRow
        ? spreadsheetFilterValues(runtime.activeSheet(), range, selection.active.column)
        : [];
    filterValueSelect.replaceChildren(new Option('Filter column', ''));
    for (const value of values) {
      filterValueSelect.add(new Option(value || '(blank)', value || '__COCO_FILTER_BLANK__'));
    }
    const enabled = filterToggle.checked && values.length > 0;
    filterValueSelect.disabled = !enabled;
    applyFilterButton.disabled = !enabled;
    clearFilterButton.disabled = !enabled;
  };

  const openSpreadsheetFind = () => {
    findBar.classList.add('is-open');
    syncSpreadsheetFind();
    findInput.focus();
    findInput.select();
  };

  const closeSpreadsheetFind = () => {
    findBar.classList.remove('is-open');
    findInput.value = '';
    findMatches = [];
    activeFindIndex = -1;
    renderSpreadsheetFindCount(findCount, findMatches, activeFindIndex);
    host.focus();
  };

  const goToSpreadsheetFindMatch = (direction: 'forward' | 'backward') => {
    syncSpreadsheetFind();
    const match = nextSpreadsheetMatch(findMatches, selection.active, direction);
    if (!match) return;
    selection = expandSelectionForMergedRanges(
      createSelection(match.row, match.column),
      runtime.activeSheet(),
    );
    activeFindIndex = findMatches.findIndex((item) => item.address === match.address);
    scroll = scrollSpreadsheetSelectionIntoView(layout, viewportManager, selection, scroll);
    dashedRange = null;
    formulaReferenceRanges = [];
    render();
  };

  const replaceAllSpreadsheetMatches = async (): Promise<void> => {
    const query = findInput.value;
    if (!query) return;
    const count = await runtime.replaceAll(query, replaceInput.value);
    findMatches = [];
    activeFindIndex = -1;
    syncSpreadsheetFind();
    render();
    findCount.textContent = `${count} replaced`;
  };

  const applySpreadsheetHistoryNavigation = async (direction: 'undo' | 'redo'): Promise<void> => {
    clearSpreadsheetTransientUiState();
    const result = direction === 'undo' ? await runtime.undo() : await runtime.redo();
    activeColumnFilters.clear();
    if (!result) {
      syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
      syncFormattingToolbar();
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      return;
    }
    selection =
      selectionFromSpreadsheetHistoryResult(result.affectedRanges, runtime) ??
      createSelection(
        Math.min(selection.active.row, runtime.activeSheet().maxRow),
        Math.min(selection.active.column, runtime.activeSheet().maxColumn),
      );
    selection = expandSelectionForMergedRanges(selection, runtime.activeSheet());
    render();
    host.focus();
  };

  const render = () => {
    layout = new SpreadsheetLayout(runtime.activeSheet());
    const size = hostSize(host);
    const sheetViewportSize = spreadsheetSheetViewportSize(size);
    viewportManager = new SpreadsheetViewportManager(layout, {
      viewportSize: sheetViewportSize,
      scroll,
      freeze: freezeStateFromSheet(runtime.activeSheet().source),
    });
    scroll = viewportManager.scroll;
    renderer.resize(size.width, size.height);
    renderer.setHeaderFiltersEnabled(filterToggle.checked && filterHeaderToggle.checked);
    renderer.setHeaderFilterColumns(
      headerFilterColumnScope === null ? null : [...headerFilterColumnScope],
    );
    renderer.setActiveHeaderFilterColumns([...activeColumnFilters.keys()]);
    renderer.render(layout, viewportManager);
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
      autoFillPreviewRange,
      canAutoFillSpreadsheetSelection(selection, runtime),
    );
    renderSpreadsheetScrollbars({
      layout,
      viewportSize: sheetViewportSize,
      freeze: freezeStateFromSheet(runtime.activeSheet().source),
      scroll,
      horizontalScrollbar,
      verticalScrollbar,
      horizontalThumb,
      verticalThumb,
    });
    renderSpreadsheetInspector(inspector, {
      fixturePath,
      fixtureSize: bytes.byteLength,
      runtime,
      scroll,
      selection,
      agentPreviewStatus: agentPreviewState,
      exportStatus: exportState,
    });
    syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
    syncFormattingToolbar();
    if (findBar.classList.contains('is-open')) syncSpreadsheetFind();
    syncSpreadsheetFilter();
    renderSpreadsheetTabs(tabs, runtime, renamingSheetName);
    if (summary) {
      summary.textContent = `${fixtures[selectedIndex]?.name ?? 'fixture'} | ${runtime.activeSheetName()} | ${runtime.activeSheet().maxRow} rows x ${runtime.activeSheet().maxColumn} cols`;
    }
    renderSpreadsheetExportStatus(exportStatus, exportState);
  };

  const commitSpreadsheetSheetRename = async (sheetName: string, value: string): Promise<void> => {
    const name = value.trim();
    renamingSheetName = null;
    if (!name || name === sheetName) {
      render();
      host.focus();
      return;
    }
    if (isSwitchingSheet) return;
    isSwitchingSheet = true;
    try {
      const viewState = sheetViewStates.get(sheetName);
      await runtime.renameSheet(sheetName, name);
      if (viewState) {
        sheetViewStates.delete(sheetName);
        sheetViewStates.set(name, viewState);
      }
      render();
      host.focus();
    } catch (err) {
      window.alert(errorMessage(err));
      render();
    } finally {
      isSwitchingSheet = false;
    }
  };

  const runSpreadsheetSheetContextMenuAction = async (
    action: SpreadsheetSheetContextMenuAction,
    sheetName: string,
  ): Promise<void> => {
    if (isSwitchingSheet) return;
    isSwitchingSheet = true;
    const wasActive = sheetName === runtime.activeSheetName();
    try {
      if (action === 'rename-sheet') {
        renamingSheetName = sheetName;
        render();
        const input = tabs.querySelector<HTMLInputElement>('[data-sheet-rename-input]');
        input?.focus();
        input?.select();
        return;
      } else if (action === 'copy-sheet') {
        const name = nextSpreadsheetSheetName(
          runtime.model().sheets.map((sheet) => sheet.name),
          `${sheetName} Copy`,
        );
        await runtime.copySheet(sheetName, name);
        selection = createSelection(1, 1);
        scroll = { x: 0, y: 0 };
      } else if (action === 'delete-sheet') {
        if (
          !(await requestSpreadsheetConfirmation(
            'Delete worksheet?',
            `Delete worksheet "${sheetName}"?`,
          ))
        )
          return;
        await runtime.deleteSheet(sheetName);
        sheetViewStates.delete(sheetName);
        if (wasActive) {
          selection = createSelection(1, 1);
          scroll = { x: 0, y: 0 };
        }
      } else {
        const index = runtime.model().sheets.findIndex((sheet) => sheet.name === sheetName);
        const targetIndex = action === 'move-sheet-left' ? index - 1 : index + 1;
        if (index >= 0 && targetIndex >= 0 && targetIndex < runtime.model().sheets.length) {
          await runtime.moveSheet(sheetName, targetIndex);
        }
      }
      render();
      host.focus();
    } catch (err) {
      window.alert(errorMessage(err));
    } finally {
      isSwitchingSheet = false;
    }
  };

  agentPreviewButton.addEventListener('click', async () => {
    clearSpreadsheetTransientUiState();
    const target = normalizeSelectionRange(selection.range);
    const action = createSpreadsheetAgentDemoAction(target);
    const context = { selection: target };
    try {
      const preview = await previewSpreadsheetAgentAction(runtime, action, context);
      agentPreviewState = {
        action: 'preview',
        ok: true,
        target,
        agentAction: action,
        context,
        preview,
      };
    } catch (err) {
      agentPreviewState = {
        action: 'preview',
        ok: false,
        target,
        agentAction: action,
        context,
        error: errorMessage(err),
      };
    }
    render();
    host.focus();
  });

  agentApplyButton.addEventListener('click', async () => {
    const previewState = canApplySpreadsheetAgentPreview(agentPreviewState)
      ? agentPreviewState
      : null;
    if (!previewState) return;
    clearSpreadsheetTransientUiState();
    try {
      const apply = await applySpreadsheetAgentAction(
        runtime,
        previewState.agentAction,
        previewState.context,
      );
      const result = apply.transaction;
      agentPreviewState = {
        action: 'apply',
        ok: true,
        target: previewState.target,
        agentAction: previewState.agentAction,
        context: previewState.context,
        preview: previewState.preview,
        apply,
      };
      if (result) {
        selection =
          selectionFromSpreadsheetHistoryResult(result.affectedRanges, runtime) ??
          createSelection(previewState.target.startRow, previewState.target.startColumn);
        selection = expandSelectionForMergedRanges(selection, runtime.activeSheet());
      }
    } catch (err) {
      agentPreviewState = {
        ...previewState,
        action: 'apply',
        ok: false,
        error: errorMessage(err),
      };
    }
    render();
    host.focus();
  });

  host.addEventListener(
    'wheel',
    (event) => {
      event.preventDefault();
      closeSpreadsheetCellEditor(editor);
      editingCell = null;
      formulaReferenceRanges = [];
      contextMenuState = null;
      hideSpreadsheetContextMenu(contextMenu);
      scroll = applyWheelDelta(scroll, event.deltaX, event.deltaY);
      render();
    },
    { passive: false },
  );

  host.addEventListener('click', (event) => {
    if (event.target === editor) return;
    if (event.button !== 0) return;
    if ((event.target as HTMLElement).closest('.coco-spreadsheet-scrollbar')) return;
    if ((event.target as HTMLElement).closest('.coco-spreadsheet-context-menu')) return;
    contextMenuState = null;
    hideSpreadsheetContextMenu(contextMenu);
    if (suppressNextClick) {
      suppressNextClick = false;
      return;
    }
    host.focus();
    const rect = host.getBoundingClientRect();
    const hit = viewportManager.hitTest({
      x: event.clientX - rect.left,
      y: event.clientY - rect.top,
    });
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') return;
    const filterColumn = hit.kind === 'cell' && hit.row === 1 ? hit.column : null;
    if (
      filterToggle.checked &&
      filterHeaderToggle.checked &&
      filterColumn !== null &&
      spreadsheetFilterIconHit(host, layout, viewportManager, event)
    ) {
      selection = selectionFromSpreadsheetHit(hit, layout);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      headerFilterState = { column: hit.column, values: new Set() };
      headerFilterState.values = renderSpreadsheetHeaderFilterMenu(
        contextMenu,
        runtime.activeSheet(),
        filterColumn,
        event.clientX,
        event.clientY,
        activeColumnFilters.get(filterColumn),
      );
      return;
    }
    dashedRange = null;
    formulaReferenceRanges = [];
    selection = selectionFromSpreadsheetHit(hit, layout);
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
    renderSpreadsheetInspector(inspector, {
      fixturePath,
      fixtureSize: bytes.byteLength,
      runtime,
      scroll,
      selection,
    });
    syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
    syncFormattingToolbar();
  });

  host.addEventListener('mousedown', (event) => {
    if (event.target === editor) return;
    if ((event.target as HTMLElement).closest('.coco-spreadsheet-scrollbar')) return;
    if (event.button !== 0) return;
    contextMenuState = null;
    hideSpreadsheetContextMenu(contextMenu);
    const resizeHandle = hitSpreadsheetResizeHandle(host, layout, viewportManager, event);
    if (resizeHandle) {
      closeSpreadsheetCellEditor(editor);
      editingCell = null;
      formulaReferenceRanges = [];
      formulaSuggestState = null;
      hideSpreadsheetFormulaSuggest(formulaSuggest);
      resizeDrag = startSpreadsheetResizeDrag(resizeHandle, event);
      updateSpreadsheetResizeGuide(resizeGuide, resizeDrag);
      suppressNextClick = true;
      event.preventDefault();
      return;
    }
    const hit = hitSpreadsheetHost(host, viewportManager, event);
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') return;
    const formulaTarget = activeFormulaReferenceEditTarget(editor, formulaInput);
    if (formulaTarget && hit.kind === 'cell') {
      formulaReferenceEdit = startSpreadsheetFormulaReferenceEdit(
        formulaTarget,
        hit.row,
        hit.column,
      );
      formulaReferenceRanges = formulaReferenceRangesFromText(formulaTarget.value);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      suppressNextClick = true;
      event.preventDefault();
      return;
    }
    const pendingEditCell = editingCell;
    const hasPendingCellEdit = pendingEditCell !== null && editor.style.display !== 'none';
    const fillHandle = fillHandleForSpreadsheetSelection(
      renderer,
      layout,
      viewportManager,
      selection,
      runtime,
    );
    if (
      !hasPendingCellEdit &&
      fillHandle &&
      pointIsInsideSpreadsheetRect(eventPointInSpreadsheetHost(host, event), fillHandle, 3)
    ) {
      autoFillDrag = { source: normalizeSelectionRange(selection.range), preview: null };
      autoFillPreviewRange = null;
      suppressNextClick = true;
      event.preventDefault();
      return;
    }
    host.focus();
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    formulaReferenceRanges = [];
    formulaSuggestState = null;
    hideSpreadsheetFormulaSuggest(formulaSuggest);
    if (hasPendingCellEdit && pendingEditCell) {
      // Clicking another cell commits the active edit before changing selection,
      // matching normal spreadsheet behavior when Enter is not pressed.
      selection = selectionFromSpreadsheetHit(hit, layout);
      if (hit.kind === 'cell') selection = expandSelectionForMergedRanges(selection, layout.model);
      dashedRange = null;
      suppressNextClick = true;
      void commitSpreadsheetEditor(runtime, editor, pendingEditCell, render, () => {
        formulaReferenceRanges = [];
      });
      event.preventDefault();
      return;
    }
    isSelecting = true;
    selectionDragKind =
      hit.kind === 'rowHeader' ? 'row' : hit.kind === 'columnHeader' ? 'column' : 'cell';
    dragDidMove = false;
    dashedRange = null;
    selection = selectionFromSpreadsheetHit(hit, layout);
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
    renderSpreadsheetInspector(inspector, {
      fixturePath,
      fixtureSize: bytes.byteLength,
      runtime,
      scroll,
      selection,
    });
    syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
    syncFormattingToolbar();
    event.preventDefault();
  });

  host.addEventListener('contextmenu', (event) => {
    if (event.target === editor) return;
    if ((event.target as HTMLElement).closest('.coco-spreadsheet-scrollbar')) return;
    event.preventDefault();
    const hit = hitSpreadsheetHost(host, viewportManager, event);
    if (hit.kind !== 'cell' && hit.kind !== 'rowHeader' && hit.kind !== 'columnHeader') {
      contextMenuState = null;
      hideSpreadsheetContextMenu(contextMenu);
      return;
    }
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    formulaReferenceRanges = [];
    if (!selectionContainsHit(selection, hit)) {
      selection = selectionFromSpreadsheetHit(hit, layout);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      renderSpreadsheetInspector(inspector, {
        fixturePath,
        fixtureSize: bytes.byteLength,
        runtime,
        scroll,
        selection,
      });
      syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
      syncFormattingToolbar();
    }
    contextMenuState = {
      selection,
      hasLocalClipboardPayload: spreadsheetClipboardPayload !== null,
    };
    renderSpreadsheetContextMenu(contextMenu, contextMenuState, event.clientX, event.clientY);
  });

  contextMenu.addEventListener('mousedown', (event) => {
    if ((event.target as HTMLElement).closest('[data-spreadsheet-filter-choice]')) return;
    event.preventDefault();
  });

  contextMenu.addEventListener('click', (event) => {
    const filterChoice = (event.target as HTMLInputElement).closest<HTMLInputElement>(
      '[data-spreadsheet-filter-choice]',
    );
    const filterApplyButton = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-filter-apply]',
    );
    const filterClearButton = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-filter-clear]',
    );
    const filterSortButton = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-filter-sort]',
    );
    const filterState = headerFilterState;
    if (filterState && filterChoice) {
      const value = filterChoice.dataset.spreadsheetFilterChoice ?? '';
      if (filterChoice.checked) filterState.values.add(value);
      else filterState.values.delete(value);
      return;
    }
    if (filterState && filterSortButton) {
      const direction = filterSortButton.dataset.spreadsheetFilterSort;
      headerFilterState = null;
      hideSpreadsheetContextMenu(contextMenu);
      if (direction === 'asc' || direction === 'desc') {
        void runtime
          .sortRows(1, runtime.activeSheet().maxRow, filterState.column, direction)
          .then(render);
      }
      return;
    }
    if (filterState && (filterApplyButton || filterClearButton)) {
      headerFilterState = null;
      hideSpreadsheetContextMenu(contextMenu);
      if (filterClearButton) {
        activeColumnFilters.delete(filterState.column);
      } else {
        activeColumnFilters.set(
          filterState.column,
          [...filterState.values].map((value) => (value === '__COCO_FILTER_BLANK__' ? '' : value)),
        );
      }
      const filters = [...activeColumnFilters.entries()].map(([column, values]) => ({
        column,
        values,
      }));
      void runtime.filterRowsByColumns(1, runtime.activeSheet().maxRow, filters).then(render);
      return;
    }
    const sheetButton = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-sheet-action]',
    );
    const sheetState = sheetContextMenuState;
    if (sheetButton && sheetState) {
      const action = sheetButton.dataset.spreadsheetSheetAction as
        SpreadsheetSheetContextMenuAction | undefined;
      sheetContextMenuState = null;
      hideSpreadsheetContextMenu(contextMenu);
      if (action) void runSpreadsheetSheetContextMenuAction(action, sheetState.sheetName);
      return;
    }
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-action]',
    );
    const state = contextMenuState;
    if (!button || !state) return;
    const action = button.dataset.spreadsheetAction as SpreadsheetContextMenuAction | undefined;
    const contextMenuAnchor = contextMenu.getBoundingClientRect();
    contextMenuState = null;
    hideSpreadsheetContextMenu(contextMenu);
    if (!action) return;
    if (action === 'cell-settings') {
      openSpreadsheetCellSettings(cellSettings);
      syncFormattingToolbar();
      host.focus();
      return;
    }
    const pasteMode = spreadsheetPasteModeFromContextAction(action);
    if (pasteMode) {
      closeSpreadsheetCellEditor(editor);
      editingCell = null;
      formulaReferenceRanges = [];
      void pasteSpreadsheetSelection(
        runtime,
        state.selection,
        pasteMode,
        (nextSelection) => {
          selection = nextSelection;
          dashedRange = normalizeSelectionRange(nextSelection.range);
          render();
          host.focus();
        },
        inspector,
        {
          fixturePath,
          fixtureSize: bytes.byteLength,
          scroll,
        },
        spreadsheetClipboardPayload,
      );
      return;
    }
    if (action === 'filter-column') {
      headerFilterColumnScope = new Set([
        ...activeColumnFilters.keys(),
        state.selection.active.column,
      ]);
      filterToggle.checked = true;
      filterHeaderToggle.checked = true;
      headerFilterState = { column: state.selection.active.column, values: new Set() };
      render();
      headerFilterState.values = renderSpreadsheetHeaderFilterMenu(
        contextMenu,
        runtime.activeSheet(),
        state.selection.active.column,
        contextMenuAnchor.left,
        contextMenuAnchor.top,
        activeColumnFilters.get(state.selection.active.column),
      );
      return;
    }
    void runSpreadsheetContextMenuAction(
      runtime,
      action,
      state.selection,
      layout,
      activeColumnFilters,
    ).then((nextSelection) => {
      selection = nextSelection;
      dashedRange = null;
      formulaReferenceRanges = [];
      render();
      host.focus();
    });
  });

  boldButton.addEventListener('click', () => {
    void applySpreadsheetFormat(runtime, selection, 'bold', render);
  });
  italicButton.addEventListener('click', () => {
    void applySpreadsheetFormat(runtime, selection, 'italic', render);
  });
  underlineButton.addEventListener('click', () => {
    void applySpreadsheetFormat(runtime, selection, 'underline', render);
  });
  fontColorInput.addEventListener('input', () => {
    void applySpreadsheetFormat(runtime, selection, 'fontColor', render, fontColorInput.value);
  });
  fillColorInput.addEventListener('input', () => {
    void applySpreadsheetFormat(runtime, selection, 'fillColor', render, fillColorInput.value);
  });
  numberFormatSelect.addEventListener('change', () => {
    void applySpreadsheetNumberFormat(
      runtime,
      selection,
      numberFormatSelect.value || undefined,
      render,
    );
  });
  mergeButton.addEventListener('click', () => {
    void applySpreadsheetMergeCommand(runtime, selection, render);
  });
  freezeSelect.addEventListener('change', () => {
    const mode = freezeSelect.value;
    freezeSelect.value = '';
    if (mode === 'panes' || mode === 'rows' || mode === 'columns' || mode === 'none') {
      void applySpreadsheetFreezePanes(runtime, selection, mode, render);
    }
  });
  borderSelect.addEventListener('change', () => {
    const mode = borderSelect.value;
    borderSelect.value = '';
    if (mode === 'all' || mode === 'outer' || mode === 'inner' || mode === 'clear') {
      void applySpreadsheetBorder(runtime, selection, mode, render);
    }
  });
  filterToggle.addEventListener('change', () => {
    headerFilterState = null;
    headerFilterColumnScope = null;
    hideSpreadsheetContextMenu(contextMenu);
    render();
  });
  filterHeaderToggle.addEventListener('change', () => {
    headerFilterState = null;
    headerFilterColumnScope = null;
    hideSpreadsheetContextMenu(contextMenu);
    render();
  });
  clearFormatButton.addEventListener('click', () => {
    void applySpreadsheetFormat(runtime, selection, 'clearFormat', render);
  });
  alignButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const align = button.dataset.spreadsheetAlign;
      if (align === 'left' || align === 'center' || align === 'right') {
        void applySpreadsheetFormat(runtime, selection, 'horizontalAlign', render, align);
      }
    });
  });
  verticalAlignButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const align = button.dataset.spreadsheetVerticalAlign;
      if (align === 'top' || align === 'center' || align === 'bottom') {
        void applySpreadsheetFormat(runtime, selection, 'verticalAlign', render, align);
      }
    });
  });

  undoButton.addEventListener('click', () => {
    void applySpreadsheetHistoryNavigation('undo');
  });
  redoButton.addEventListener('click', () => {
    void applySpreadsheetHistoryNavigation('redo');
  });

  cellSettingsClose.addEventListener('click', () => {
    hideSpreadsheetCellSettings(cellSettings);
    host.focus();
  });
  confirmCancelButton.addEventListener('click', () => settleSpreadsheetConfirmation(false));
  confirmAcceptButton.addEventListener('click', () => settleSpreadsheetConfirmation(true));
  confirmDialog.addEventListener('click', (event) => {
    if (event.target === confirmDialog) settleSpreadsheetConfirmation(false);
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !confirmDialog.hidden) {
      event.preventDefault();
      settleSpreadsheetConfirmation(false);
    }
  });
  cellSettingsFontColor.addEventListener('input', () => {
    void applySpreadsheetFormat(
      runtime,
      selection,
      'fontColor',
      render,
      cellSettingsFontColor.value,
    );
  });
  cellSettingsFillColor.addEventListener('input', () => {
    void applySpreadsheetFormat(
      runtime,
      selection,
      'fillColor',
      render,
      cellSettingsFillColor.value,
    );
  });
  cellSettingsNumberFormat.addEventListener('change', () => {
    void applySpreadsheetNumberFormat(
      runtime,
      selection,
      cellSettingsNumberFormat.value || undefined,
      render,
    );
  });
  cellSettingsBorder.addEventListener('change', () => {
    const mode = cellSettingsBorder.value;
    cellSettingsBorder.value = '';
    if (mode === 'all' || mode === 'outer' || mode === 'inner' || mode === 'clear') {
      void applySpreadsheetBorder(runtime, selection, mode, render);
    }
  });
  cellSettingsActionButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const action = button.dataset.cellSettingsAction;
      if (action === 'bold' || action === 'italic' || action === 'underline') {
        void applySpreadsheetFormat(runtime, selection, action, render);
      } else if (action === 'merge') {
        void applySpreadsheetMergeCommand(runtime, selection, render);
      } else if (action === 'unmerge') {
        void unmergeSpreadsheetSelection(runtime, selection).then(render);
      } else if (action === 'clear-format') {
        void applySpreadsheetFormat(runtime, selection, 'clearFormat', render);
      }
    });
  });
  cellSettingsAlignButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const align = button.dataset.cellSettingsAlign;
      if (align === 'left' || align === 'center' || align === 'right') {
        void applySpreadsheetFormat(runtime, selection, 'horizontalAlign', render, align);
      }
    });
  });
  cellSettingsVerticalAlignButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const align = button.dataset.cellSettingsVerticalAlign;
      if (align === 'top' || align === 'center' || align === 'bottom') {
        void applySpreadsheetFormat(runtime, selection, 'verticalAlign', render, align);
      }
    });
  });

  host.addEventListener('mousemove', (event) => {
    if (resizeDrag || autoFillDrag) return;
    if (fillHandleForSpreadsheetSelection(renderer, layout, viewportManager, selection, runtime)) {
      const fillHandle = fillHandleForSpreadsheetSelection(
        renderer,
        layout,
        viewportManager,
        selection,
        runtime,
      );
      if (
        fillHandle &&
        pointIsInsideSpreadsheetRect(eventPointInSpreadsheetHost(host, event), fillHandle, 3)
      ) {
        host.style.cursor = 'crosshair';
        return;
      }
    }
    const resizeHandle = hitSpreadsheetResizeHandle(host, layout, viewportManager, event);
    host.style.cursor =
      resizeHandle?.kind === 'column'
        ? 'col-resize'
        : resizeHandle?.kind === 'row'
          ? 'row-resize'
          : '';
  });

  host.addEventListener('mouseleave', () => {
    if (!resizeDrag && !autoFillDrag) host.style.cursor = '';
  });

  window.addEventListener('mousemove', (event) => {
    if (resizeDrag) {
      resizeDrag = updateSpreadsheetResizeDrag(resizeDrag, event);
      updateSpreadsheetResizeGuide(resizeGuide, resizeDrag);
      event.preventDefault();
      return;
    }
    if (autoFillDrag) {
      const hit = hitSpreadsheetHost(host, viewportManager, event);
      const preview = autoFillPreviewRangeForHit(autoFillDrag.source, hit);
      autoFillDrag = { ...autoFillDrag, preview };
      autoFillPreviewRange = preview;
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
        autoFillPreviewRange,
        canAutoFillSpreadsheetSelection(selection, runtime),
      );
      event.preventDefault();
      return;
    }
    if (formulaReferenceEdit) {
      const hit = hitSpreadsheetHost(host, viewportManager, event);
      if (hit.kind !== 'cell') return;
      updateSpreadsheetFormulaReferenceEdit(formulaReferenceEdit, hit.row, hit.column);
      formulaReferenceRanges = formulaReferenceRangesFromText(formulaReferenceEdit.target.value);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      event.preventDefault();
      return;
    }
    if (!isSelecting) return;
    const hit = hitSpreadsheetHost(host, viewportManager, event);
    if (selectionDragKind === 'row') {
      if (hit.kind !== 'rowHeader' && hit.kind !== 'cell') return;
      const row = hit.kind === 'rowHeader' || hit.kind === 'cell' ? hit.row : selection.active.row;
      if (row !== selection.active.row) dragDidMove = true;
      selection = extendSelection(selection, row, layout.columnCount, {
        rowCount: layout.rowCount,
        columnCount: layout.columnCount,
      });
    } else if (selectionDragKind === 'column') {
      if (hit.kind !== 'columnHeader' && hit.kind !== 'cell') return;
      const column =
        hit.kind === 'columnHeader' || hit.kind === 'cell' ? hit.column : selection.active.column;
      if (column !== selection.active.column) dragDidMove = true;
      selection = extendSelection(selection, layout.rowCount, column, {
        rowCount: layout.rowCount,
        columnCount: layout.columnCount,
      });
    } else {
      if (hit.kind !== 'cell') return;
      if (hit.row !== selection.active.row || hit.column !== selection.active.column)
        dragDidMove = true;
      selection = expandSelectionForMergedRanges(
        extendSelection(selection, hit.row, hit.column, {
          rowCount: layout.rowCount,
          columnCount: layout.columnCount,
        }),
        layout.model,
      );
    }
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
    renderSpreadsheetInspector(inspector, {
      fixturePath,
      fixtureSize: bytes.byteLength,
      runtime,
      scroll,
      selection,
    });
    syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
    syncFormattingToolbar();
  });

  window.addEventListener('mouseup', (event) => {
    if (resizeDrag) {
      const completed = updateSpreadsheetResizeDrag(resizeDrag, event);
      resizeDrag = null;
      hideSpreadsheetResizeGuide(resizeGuide);
      host.style.cursor = '';
      suppressNextClick = true;
      void commitSpreadsheetResizeDrag(runtime, completed, render);
      return;
    }
    if (autoFillDrag) {
      const completed = autoFillDrag;
      autoFillDrag = null;
      autoFillPreviewRange = null;
      host.style.cursor = '';
      suppressNextClick = true;
      if (completed.preview) {
        void runtime.autoFillRange(completed.source, completed.preview).then(() => {
          selection = { ...selection, range: completed.preview! };
          dashedRange = null;
          render();
          host.focus();
        });
      } else {
        render();
      }
      return;
    }
    formulaReferenceEdit = null;
    if (isSelecting && dragDidMove) suppressNextClick = true;
    isSelecting = false;
    selectionDragKind = null;
    dragDidMove = false;
  });

  host.addEventListener('dblclick', (event) => {
    if (event.target === editor) return;
    const rect = host.getBoundingClientRect();
    const hit = viewportManager.hitTest({
      x: event.clientX - rect.left,
      y: event.clientY - rect.top,
    });
    if (hit.kind !== 'cell') return;
    selection = selectionFromSpreadsheetHit(hit, layout);
    editingCell = { row: selection.active.row, column: selection.active.column };
    const cell = getRenderCell(
      runtime.activeSheet(),
      selection.active.row,
      selection.active.column,
    );
    const editorInitialValue = cell?.formula ? `=${cell.formula}` : (cell?.displayText ?? '');
    formulaReferenceRanges = formulaReferenceRangesFromText(editorInitialValue);
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
    syncFormattingToolbar();
    openSpreadsheetCellEditor({
      editor,
      layout,
      viewportManager,
      row: selection.active.row,
      column: selection.active.column,
      initialValue: editorInitialValue,
      selectText: true,
    });
  });

  editor.addEventListener('keydown', (event) => {
    event.stopPropagation();
    if (
      handleSpreadsheetFormulaSuggestKeydown(event, formulaSuggestState, formulaSuggest, (next) => {
        formulaSuggestState = next;
      })
    ) {
      formulaReferenceRanges = formulaReferenceRangesFromText(editor.value);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      closeSpreadsheetCellEditor(editor);
      editingCell = null;
      formulaReferenceRanges = [];
      formulaSuggestState = null;
      hideSpreadsheetFormulaSuggest(formulaSuggest);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
    }
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      void commitSpreadsheetEditor(runtime, editor, editingCell ?? selection.active, render, () => {
        editingCell = null;
        formulaReferenceRanges = [];
        formulaSuggestState = null;
        hideSpreadsheetFormulaSuggest(formulaSuggest);
        selection = expandSelectionForMergedRanges(
          moveSelection(selection, 1, 0, {
            rowCount: layout.rowCount,
            columnCount: layout.columnCount,
          }),
          runtime.activeSheet(),
        );
      });
    }
    if (event.key === 'Tab') {
      event.preventDefault();
      const delta = event.shiftKey ? -1 : 1;
      void commitSpreadsheetEditor(runtime, editor, editingCell ?? selection.active, render, () => {
        editingCell = null;
        formulaReferenceRanges = [];
        formulaSuggestState = null;
        hideSpreadsheetFormulaSuggest(formulaSuggest);
        selection = expandSelectionForMergedRanges(
          moveSelection(selection, 0, delta, {
            rowCount: layout.rowCount,
            columnCount: layout.columnCount,
          }),
          runtime.activeSheet(),
        );
      });
    }
  });

  editor.addEventListener('input', () => {
    formulaReferenceRanges = formulaReferenceRangesFromText(editor.value);
    formulaSuggestState = updateSpreadsheetFormulaSuggest(
      editor,
      formulaSuggest,
      formulaSuggestState,
    );
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
  });

  nameBox.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    const target = parseSpreadsheetNameBox(nameBox.value);
    if (!target) {
      syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
      return;
    }
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    formulaReferenceRanges = [];
    formulaSuggestState = null;
    hideSpreadsheetFormulaSuggest(formulaSuggest);
    dashedRange = null;
    selection = expandSelectionForMergedRanges(
      clampSpreadsheetSelection(target, layout),
      runtime.activeSheet(),
    );
    scroll = scrollSpreadsheetSelectionIntoView(layout, viewportManager, selection, scroll);
    render();
    host.focus();
  });

  findInput.addEventListener('input', () => {
    syncSpreadsheetFind();
  });
  replaceInput.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    void replaceAllSpreadsheetMatches();
  });
  findInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      goToSpreadsheetFindMatch(event.shiftKey ? 'backward' : 'forward');
    } else if (event.key === 'Escape') {
      event.preventDefault();
      closeSpreadsheetFind();
    }
  });
  findPreviousButton.addEventListener('click', () => goToSpreadsheetFindMatch('backward'));
  findNextButton.addEventListener('click', () => goToSpreadsheetFindMatch('forward'));
  replaceAllButton.addEventListener('click', () => void replaceAllSpreadsheetMatches());
  findCloseButton.addEventListener('click', closeSpreadsheetFind);
  applyFilterButton.addEventListener('click', () => {
    const range = spreadsheetFilterRange();
    if (filterValueSelect.selectedIndex < 1) return;
    const value =
      filterValueSelect.value === '__COCO_FILTER_BLANK__' ? '' : filterValueSelect.value;
    activeColumnFilters.set(selection.active.column, [value]);
    const filters = [...activeColumnFilters.entries()].map(([column, values]) => ({
      column,
      values,
    }));
    void runtime.filterRowsByColumns(range.startRow, range.endRow, filters).then(render);
  });
  clearFilterButton.addEventListener('click', () => {
    const range = spreadsheetFilterRange();
    activeColumnFilters.delete(selection.active.column);
    const filters = [...activeColumnFilters.entries()].map(([column, values]) => ({
      column,
      values,
    }));
    void runtime.filterRowsByColumns(range.startRow, range.endRow, filters).then(render);
  });
  clearAllFiltersButton.addEventListener('click', () => {
    activeColumnFilters.clear();
    void runtime.filterRowsByColumns(1, runtime.activeSheet().maxRow, []).then(render);
  });
  const sortActiveColumn = (direction: 'asc' | 'desc') => {
    const startRow = filterHeaderToggle.checked ? 1 : 0;
    void runtime
      .sortRows(startRow, runtime.activeSheet().maxRow, selection.active.column, direction)
      .then(render);
  };
  sortAscendingButton.addEventListener('click', () => sortActiveColumn('asc'));
  sortDescendingButton.addEventListener('click', () => sortActiveColumn('desc'));

  formulaInput.addEventListener('keydown', (event) => {
    if (
      handleSpreadsheetFormulaSuggestKeydown(event, formulaSuggestState, formulaSuggest, (next) => {
        formulaSuggestState = next;
      })
    ) {
      formulaReferenceRanges = formulaReferenceRangesFromText(formulaInput.value);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      syncSpreadsheetFormulaBar(nameBox, formulaInput, runtime, selection);
      formulaReferenceRanges = [];
      formulaSuggestState = null;
      hideSpreadsheetFormulaSuggest(formulaSuggest);
      renderSpreadsheetOverlay(
        renderer,
        layout,
        viewportManager,
        selection,
        dashedRange,
        formulaReferenceRanges,
      );
      host.focus();
      return;
    }
    if (event.key !== 'Enter') return;
    event.preventDefault();
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    formulaReferenceRanges = [];
    formulaSuggestState = null;
    hideSpreadsheetFormulaSuggest(formulaSuggest);
    dashedRange = null;
    void commitSpreadsheetFormulaBar(runtime, selection, formulaInput.value, render).then(() => {
      host.focus();
    });
  });

  formulaInput.addEventListener('focus', () => {
    formulaReferenceRanges = formulaReferenceRangesFromText(formulaInput.value);
    formulaSuggestState = updateSpreadsheetFormulaSuggest(
      formulaInput,
      formulaSuggest,
      formulaSuggestState,
    );
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
  });

  formulaInput.addEventListener('input', () => {
    formulaReferenceRanges = formulaReferenceRangesFromText(formulaInput.value);
    formulaSuggestState = updateSpreadsheetFormulaSuggest(
      formulaInput,
      formulaSuggest,
      formulaSuggestState,
    );
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
  });

  formulaSuggest.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-formula-function]',
    );
    const state = formulaSuggestState;
    const target = state?.target;
    if (!button || !state || !target) return;
    formulaSuggestState = acceptSpreadsheetFormulaSuggestion(
      state,
      formulaSuggest,
      Number(button.dataset.formulaFunctionIndex ?? 0),
    );
    formulaReferenceRanges = formulaReferenceRangesFromText(target.value);
    renderSpreadsheetOverlay(
      renderer,
      layout,
      viewportManager,
      selection,
      dashedRange,
      formulaReferenceRanges,
    );
  });

  host.addEventListener('keydown', (event) => {
    if (event.defaultPrevented) return;
    if (event.key === 'Escape' && contextMenuState) {
      event.preventDefault();
      contextMenuState = null;
      hideSpreadsheetContextMenu(contextMenu);
      return;
    }
    if (event.key === 'Escape' && cellSettings.style.display !== 'none') {
      event.preventDefault();
      hideSpreadsheetCellSettings(cellSettings);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      void applySpreadsheetHistoryNavigation(event.shiftKey ? 'redo' : 'undo');
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 'y') {
      event.preventDefault();
      void applySpreadsheetHistoryNavigation('redo');
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openSpreadsheetFind();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c') {
      event.preventDefault();
      void copySpreadsheetSelection(runtime, selection, inspector, {
        fixturePath,
        fixtureSize: bytes.byteLength,
        scroll,
      }).then((result) => {
        spreadsheetClipboardPayload = result?.payload ?? null;
        dashedRange = result?.range ?? null;
        renderSpreadsheetOverlay(
          renderer,
          layout,
          viewportManager,
          selection,
          dashedRange,
          formulaReferenceRanges,
        );
      });
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v') {
      event.preventDefault();
      closeSpreadsheetCellEditor(editor);
      editingCell = null;
      formulaReferenceRanges = [];
      void pasteSpreadsheetSelection(
        runtime,
        selection,
        'all',
        (nextSelection) => {
          selection = nextSelection;
          dashedRange = normalizeSelectionRange(nextSelection.range);
          render();
        },
        inspector,
        {
          fixturePath,
          fixtureSize: bytes.byteLength,
          scroll,
        },
        spreadsheetClipboardPayload,
      );
      return;
    }
    if (event.metaKey || event.ctrlKey || event.altKey) return;
    if (editor.style.display !== 'none') {
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      dashedRange = null;
      formulaReferenceRanges = [];
      selection = expandSelectionForMergedRanges(
        moveSelection(selection, event.shiftKey ? -1 : 1, 0, {
          rowCount: layout.rowCount,
          columnCount: layout.columnCount,
        }),
        runtime.activeSheet(),
      );
      render();
      return;
    }
    if (event.key === 'Tab') {
      event.preventDefault();
      dashedRange = null;
      formulaReferenceRanges = [];
      selection = expandSelectionForMergedRanges(
        moveSelection(selection, 0, event.shiftKey ? -1 : 1, {
          rowCount: layout.rowCount,
          columnCount: layout.columnCount,
        }),
        runtime.activeSheet(),
      );
      render();
      return;
    }
    if (event.key === 'Backspace' || event.key === 'Delete') {
      event.preventDefault();
      dashedRange = null;
      formulaReferenceRanges = [];
      void clearSpreadsheetSelection(runtime, selection, render);
      return;
    }
    if (event.key.length === 1) {
      event.preventDefault();
      dashedRange = null;
      editingCell = { row: selection.active.row, column: selection.active.column };
      formulaReferenceRanges = formulaReferenceRangesFromText(event.key);
      openSpreadsheetCellEditor({
        editor,
        layout,
        viewportManager,
        row: selection.active.row,
        column: selection.active.column,
        initialValue: event.key,
        selectText: false,
      });
    }
  });

  exportButton.addEventListener('click', async () => {
    const fileName = exportedSpreadsheetFileName(fixtures[selectedIndex]?.name ?? 'workbook.xlsx');
    exportState = { status: 'saving', fileName, startedAt: Date.now() };
    render();
    try {
      const out = await runtime.toZip();
      downloadSpreadsheetBytes(out, fileName);
      exportState = {
        status: 'ok',
        fileName,
        byteLength: out.byteLength,
        finishedAt: Date.now(),
      };
    } catch (err) {
      exportState = {
        status: 'error',
        fileName,
        message: errorMessage(err),
        finishedAt: Date.now(),
      };
    }
    render();
    host.focus();
  });

  tabs.addEventListener('click', async (event) => {
    const addButton = (event.target as HTMLElement).closest<HTMLButtonElement>(
      '[data-spreadsheet-sheet-add]',
    );
    if (addButton) {
      if (isSwitchingSheet) return;
      isSwitchingSheet = true;
      try {
        const name = nextSpreadsheetSheetName(
          runtime.model().sheets.map((sheet) => sheet.name),
          'Sheet',
        );
        clearSpreadsheetTransientUiState();
        await runtime.addSheet(name);
        selection = createSelection(1, 1);
        scroll = { x: 0, y: 0 };
        render();
        host.focus();
      } finally {
        isSwitchingSheet = false;
      }
      return;
    }
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-sheet-name]');
    if (!button) return;
    const sheetName = button.dataset.sheetName;
    if (!sheetName || sheetName === runtime.activeSheetName() || isSwitchingSheet) return;
    isSwitchingSheet = true;
    sheetViewStates.set(runtime.activeSheetName(), { selection, scroll });
    closeSpreadsheetCellEditor(editor);
    editingCell = null;
    formulaReferenceRanges = [];
    isSelecting = false;
    selectionDragKind = null;
    dragDidMove = false;
    suppressNextClick = false;
    dashedRange = null;
    try {
      await runtime.setActiveSheet(sheetName);
      const viewState = sheetViewStates.get(sheetName);
      selection = viewState?.selection ?? createSelection(1, 1);
      scroll = viewState?.scroll ?? { x: 0, y: 0 };
      render();
      host.focus();
    } finally {
      isSwitchingSheet = false;
    }
  });

  tabs.addEventListener('contextmenu', (event) => {
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-sheet-name]');
    const sheetName = button?.dataset.sheetName;
    if (!sheetName) return;
    event.preventDefault();
    clearSpreadsheetTransientUiState();
    sheetContextMenuState = { sheetName };
    renderSpreadsheetSheetContextMenu(
      contextMenu,
      sheetContextMenuState,
      event.clientX,
      event.clientY,
    );
  });

  tabs.addEventListener('keydown', (event) => {
    const input = (event.target as HTMLElement).closest<HTMLInputElement>(
      '[data-sheet-rename-input]',
    );
    if (!input) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      renamingSheetName = null;
      render();
      host.focus();
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      const sheetName = input.dataset.sheetName;
      if (sheetName) void commitSpreadsheetSheetRename(sheetName, input.value);
    }
  });

  tabs.addEventListener('focusout', (event) => {
    const input = (event.target as HTMLElement).closest<HTMLInputElement>(
      '[data-sheet-rename-input]',
    );
    const sheetName = input?.dataset.sheetName;
    if (input && sheetName && renamingSheetName === sheetName) {
      void commitSpreadsheetSheetRename(sheetName, input.value);
    }
  });

  wireSpreadsheetScrollbarDrag({
    axis: 'horizontal',
    scrollbar: horizontalScrollbar,
    thumb: horizontalThumb,
    getGeometry: () =>
      computeScrollbarState(
        layout,
        hostSize(host),
        freezeStateFromSheet(runtime.activeSheet().source),
        scroll,
      ).horizontal,
    getScroll: () => scroll.x,
    setScroll: (value) => {
      scroll = { ...scroll, x: value };
      render();
    },
  });
  wireSpreadsheetScrollbarDrag({
    axis: 'vertical',
    scrollbar: verticalScrollbar,
    thumb: verticalThumb,
    getGeometry: () =>
      computeScrollbarState(
        layout,
        hostSize(host),
        freezeStateFromSheet(runtime.activeSheet().source),
        scroll,
      ).vertical,
    getScroll: () => scroll.y,
    setScroll: (value) => {
      scroll = { ...scroll, y: value };
      render();
    },
  });

  const resizeObserver = new ResizeObserver(render);
  resizeObserver.observe(host);
  render();
}

async function loadFixtures(): Promise<FixtureFile[]> {
  const response = await fetch('/__coco_fixture__/files');
  if (!response.ok) throw new Error(`Failed to list fixtures: ${response.status}`);
  return response.json();
}

async function loadSpreadsheetFixtures(): Promise<FixtureFile[]> {
  const response = await fetch('/__coco_fixture__/xlsx-files');
  if (!response.ok) throw new Error(`Failed to list spreadsheet fixtures: ${response.status}`);
  return response.json();
}

async function loadPptFixtures(): Promise<FixtureFile[]> {
  const response = await fetch('/__coco_fixture__/pptx-files');
  if (!response.ok) throw new Error(`Failed to list ppt fixtures: ${response.status}`);
  return response.json();
}

function exportedFileName(input: string): string {
  const base = input.toLowerCase().endsWith('.docx') ? input.slice(0, -5) : input;
  return `${base || 'document'}.edited.docx`;
}

function exportedSpreadsheetFileName(input: string): string {
  const base = input.toLowerCase().endsWith('.xlsx') ? input.slice(0, -5) : input;
  return `${base || 'workbook'}.edited.xlsx`;
}

function exportedPptFileName(input: string): string {
  const base = input.toLowerCase().endsWith('.pptx') ? input.slice(0, -5) : input;
  return `${base || 'presentation'}.edited.pptx`;
}

function isPptMetafilePreviewCandidate(image: PptEditorImagePreviewRequest): boolean {
  const contentType = image.contentType.toLowerCase();
  const path = image.path.toLowerCase();
  return (
    contentType.includes('emf') ||
    contentType.includes('wmf') ||
    path.endsWith('.emf') ||
    path.endsWith('.wmf')
  );
}

async function resolvePptImagePreview(
  image: PptEditorImagePreviewRequest,
): Promise<PptEditorImageResource | undefined> {
  if (!isPptMetafilePreviewCandidate(image)) return undefined;
  try {
    const response = await fetch('/__coco_fixture__/ppt-image-preview', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/octet-stream',
        'X-Coco-Image-Content-Type': image.contentType,
        'X-Coco-Image-Path': encodeURIComponent(image.path),
      },
      body: arrayBufferFromBytes(image.bytes),
    });
    if (!response.ok) return undefined;
    const blob = await response.blob();
    return {
      url: URL.createObjectURL(blob),
      contentType: blob.type || 'image/png',
    };
  } catch {
    return undefined;
  }
}

function downloadBytes(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([arrayBufferFromBytes(bytes)], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function downloadPptBytes(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([arrayBufferFromBytes(bytes)], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function downloadSpreadsheetBytes(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([arrayBufferFromBytes(bytes)], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function arrayBufferFromBytes(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function hostSize(host: HTMLElement): { width: number; height: number } {
  const rect = host.getBoundingClientRect();
  return {
    width: Math.max(1, rect.width),
    height: Math.max(1, rect.height),
  };
}

function spreadsheetSheetViewportSize(size: { width: number; height: number }): {
  width: number;
  height: number;
} {
  return {
    width: Math.max(1, size.width - 12),
    height: Math.max(1, size.height - 12),
  };
}

function renderSpreadsheetScrollbars(input: {
  layout: SpreadsheetLayout;
  viewportSize: { width: number; height: number };
  freeze: { frozenRows: number; frozenColumns: number };
  scroll: { x: number; y: number };
  horizontalScrollbar: HTMLDivElement;
  verticalScrollbar: HTMLDivElement;
  horizontalThumb: HTMLDivElement;
  verticalThumb: HTMLDivElement;
}): void {
  const state = computeScrollbarState(input.layout, input.viewportSize, input.freeze, input.scroll);
  const frozenWidth = input.layout.getColumnX(input.freeze.frozenColumns + 1);
  const frozenHeight = input.layout.getRowY(input.freeze.frozenRows + 1);
  input.horizontalScrollbar.style.left = `${input.layout.rowHeaderWidth + frozenWidth}px`;
  input.horizontalScrollbar.style.width = `${state.horizontal.trackLength}px`;
  input.horizontalScrollbar.style.display = state.horizontal.maxScroll > 0 ? 'block' : 'none';
  input.horizontalThumb.style.left = `${state.horizontal.thumbStart}px`;
  input.horizontalThumb.style.top = '2px';
  input.horizontalThumb.style.width = `${state.horizontal.thumbLength}px`;
  input.horizontalThumb.style.height = '6px';

  input.verticalScrollbar.style.top = `${input.layout.columnHeaderHeight + frozenHeight}px`;
  input.verticalScrollbar.style.height = `${state.vertical.trackLength}px`;
  input.verticalScrollbar.style.display = state.vertical.maxScroll > 0 ? 'block' : 'none';
  input.verticalThumb.style.left = '2px';
  input.verticalThumb.style.top = `${state.vertical.thumbStart}px`;
  input.verticalThumb.style.width = '6px';
  input.verticalThumb.style.height = `${state.vertical.thumbLength}px`;
}

function wireSpreadsheetScrollbarDrag(input: {
  axis: 'horizontal' | 'vertical';
  scrollbar: HTMLDivElement;
  thumb: HTMLDivElement;
  getGeometry: () => {
    trackLength: number;
    thumbStart: number;
    thumbLength: number;
    maxScroll: number;
  };
  getScroll: () => number;
  setScroll: (value: number) => void;
}): void {
  let dragging: {
    pointerStart: number;
    thumbStart: number;
  } | null = null;

  input.scrollbar.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const geometry = input.getGeometry();
    const rect = input.scrollbar.getBoundingClientRect();
    const pointer =
      input.axis === 'horizontal' ? event.clientX - rect.left : event.clientY - rect.top;
    if (event.target === input.thumb) {
      dragging = { pointerStart: pointer, thumbStart: geometry.thumbStart };
      input.thumb.classList.add('is-dragging');
      return;
    }
    input.setScroll(scrollFromThumbPosition(geometry, pointer - geometry.thumbLength / 2));
  });

  window.addEventListener('mousemove', (event) => {
    if (!dragging) return;
    const geometry = input.getGeometry();
    const rect = input.scrollbar.getBoundingClientRect();
    const pointer =
      input.axis === 'horizontal' ? event.clientX - rect.left : event.clientY - rect.top;
    input.setScroll(
      scrollFromThumbPosition(geometry, dragging.thumbStart + pointer - dragging.pointerStart),
    );
  });

  window.addEventListener('mouseup', () => {
    dragging = null;
    input.thumb.classList.remove('is-dragging');
  });
}

function hitSpreadsheetHost(
  host: HTMLElement,
  viewportManager: SpreadsheetViewportManager,
  event: MouseEvent,
) {
  return viewportManager.hitTest(eventPointInSpreadsheetHost(host, event));
}

function spreadsheetFilterIconHit(
  host: HTMLElement,
  layout: SpreadsheetLayout,
  viewportManager: SpreadsheetViewportManager,
  event: MouseEvent,
): boolean {
  const point = eventPointInSpreadsheetHost(host, event);
  const viewport = viewportManager.findViewport(point);
  if (!viewport) return false;
  const logicalX = point.x - viewport.screenRect.x + viewport.originX;
  const logicalY = point.y - viewport.screenRect.y + viewport.originY;
  if (viewport.kind === 'columnHeaderFrozen' || viewport.kind === 'columnHeaderScrollable')
    return false;
  if (
    viewport.kind === 'rowHeaderFrozen' ||
    viewport.kind === 'rowHeaderScrollable' ||
    layout.rowAt(logicalY) !== 1
  )
    return false;
  const column = layout.columnAt(logicalX);
  const rect = viewportManager.cellScreenRect(1, column, viewport);
  return rect.width >= 28 && point.x >= rect.x + rect.width - 22 && point.x <= rect.x + rect.width;
}

interface SpreadsheetAutoFillDragState {
  source: SpreadsheetRange;
  preview: SpreadsheetRange | null;
}

function canAutoFillSpreadsheetSelection(
  selection: SpreadsheetSelectionState,
  runtime: SpreadsheetEditorRuntime,
): boolean {
  return (
    selection.kind === 'cell' && supportsSpreadsheetAutoFill(runtime.activeSheet(), selection.range)
  );
}

function fillHandleForSpreadsheetSelection(
  renderer: SpreadsheetCanvasRenderer,
  layout: SpreadsheetLayout,
  viewportManager: SpreadsheetViewportManager,
  selection: SpreadsheetSelectionState,
  runtime: SpreadsheetEditorRuntime,
) {
  if (!canAutoFillSpreadsheetSelection(selection, runtime)) return null;
  return renderer.fillHandleRect(layout, viewportManager, normalizeSelectionRange(selection.range));
}

function autoFillPreviewRangeForHit(
  source: SpreadsheetRange,
  hit: ReturnType<SpreadsheetViewportManager['hitTest']>,
): SpreadsheetRange | null {
  if (hit.kind !== 'cell') return null;
  const down = Math.max(0, hit.row - source.endRow);
  const right = Math.max(0, hit.column - source.endColumn);
  if (!down && !right) return null;
  if (down >= right) return { ...source, endRow: source.endRow + down };
  return { ...source, endColumn: source.endColumn + right };
}

function eventPointInSpreadsheetHost(
  host: HTMLElement,
  event: MouseEvent,
): { x: number; y: number } {
  const rect = host.getBoundingClientRect();
  return { x: event.clientX - rect.left, y: event.clientY - rect.top };
}

function pointIsInsideSpreadsheetRect(
  point: { x: number; y: number },
  rect: { x: number; y: number; width: number; height: number },
  padding = 0,
): boolean {
  return (
    point.x >= rect.x - padding &&
    point.x <= rect.x + rect.width + padding &&
    point.y >= rect.y - padding &&
    point.y <= rect.y + rect.height + padding
  );
}

function selectionFromSpreadsheetHit(
  hit: ReturnType<SpreadsheetViewportManager['hitTest']>,
  layout: SpreadsheetLayout,
): SpreadsheetSelectionState {
  if (hit.kind === 'rowHeader')
    return createRowSelection(hit.row, { columnCount: layout.columnCount });
  if (hit.kind === 'columnHeader')
    return createColumnSelection(hit.column, { rowCount: layout.rowCount });
  if (hit.kind === 'cell') {
    const masterAddress = layout.model.mergeFollowerByAddress.get(hit.address) ?? hit.address;
    const merge = layout.model.mergeMasterByAddress.get(masterAddress);
    if (merge) {
      return {
        kind: 'cell',
        active: { row: merge.range.startRow, column: merge.range.startColumn },
        range: merge.range,
      };
    }
    return createSelection(hit.row, hit.column);
  }
  return createSelection(1, 1);
}

function selectionContainsHit(
  selection: SpreadsheetSelectionState,
  hit: ReturnType<SpreadsheetViewportManager['hitTest']>,
): boolean {
  const range = normalizeSelectionRange(selection.range);
  if (hit.kind === 'cell') {
    return (
      hit.row >= range.startRow &&
      hit.row <= range.endRow &&
      hit.column >= range.startColumn &&
      hit.column <= range.endColumn
    );
  }
  if (hit.kind === 'rowHeader') {
    return selection.kind === 'row' && hit.row >= range.startRow && hit.row <= range.endRow;
  }
  if (hit.kind === 'columnHeader') {
    return (
      selection.kind === 'column' &&
      hit.column >= range.startColumn &&
      hit.column <= range.endColumn
    );
  }
  return false;
}

async function commitSpreadsheetEditor(
  runtime: SpreadsheetEditorRuntime,
  editor: HTMLTextAreaElement,
  cell: { row: number; column: number },
  render: () => void,
  afterCommit?: () => void,
): Promise<void> {
  const value = editor.value;
  closeSpreadsheetCellEditor(editor);
  const address = formatCellAddress(cell.row, cell.column);
  const historyOptions = {
    mergeUndo: true,
    mergeWindowMs: SPREADSHEET_CELL_EDIT_MERGE_WINDOW_MS,
  };
  if (value.startsWith('=')) await runtime.setCellFormula(address, value.slice(1), historyOptions);
  else {
    // Keep numeric and boolean edits typed so default spreadsheet alignment does not change.
    await runtime.setCellValue(address, coerceSpreadsheetFormulaBarValue(value), historyOptions);
  }
  afterCommit?.();
  render();
  editor.closest<HTMLElement>('.coco-spreadsheet-canvas-host')?.focus();
}

function closeSpreadsheetCellEditor(editor: HTMLTextAreaElement): void {
  editor.style.display = 'none';
}

interface SpreadsheetFormulaReferenceRange {
  range: SpreadsheetRange;
  color: string;
}

interface SpreadsheetFormulaReferenceEditState {
  target: HTMLInputElement | HTMLTextAreaElement;
  anchor: { row: number; column: number };
  originalValue: string;
  replaceStart: number;
  replaceEnd: number;
}

interface SpreadsheetFormulaFunctionInfo {
  name: string;
  signature: string;
}

interface SpreadsheetFormulaSuggestState {
  target: HTMLInputElement | HTMLTextAreaElement;
  replaceStart: number;
  replaceEnd: number;
  activeIndex: number;
  items: SpreadsheetFormulaFunctionInfo[];
}

type SpreadsheetContextMenuAction =
  | 'insert-row-above'
  | 'insert-row-below'
  | 'delete-rows'
  | 'insert-column-left'
  | 'insert-column-right'
  | 'delete-columns'
  | 'paste'
  | 'paste-values'
  | 'paste-formats'
  | 'clear'
  | 'clear-format'
  | 'merge-cells'
  | 'unmerge-cells'
  | 'border-all'
  | 'border-outer'
  | 'border-inner'
  | 'border-clear'
  | 'freeze-panes'
  | 'freeze-rows'
  | 'freeze-columns'
  | 'unfreeze-panes'
  | 'filter-column'
  | 'cell-settings';

interface SpreadsheetContextMenuState {
  selection: SpreadsheetSelectionState;
  hasLocalClipboardPayload: boolean;
}

type SpreadsheetSheetContextMenuAction =
  'rename-sheet' | 'copy-sheet' | 'delete-sheet' | 'move-sheet-left' | 'move-sheet-right';

interface SpreadsheetSheetContextMenuState {
  sheetName: string;
}

type SpreadsheetFormatAction =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'fontColor'
  | 'fillColor'
  | 'horizontalAlign'
  | 'verticalAlign'
  | 'clearFormat';

function renderSpreadsheetOverlay(
  renderer: SpreadsheetCanvasRenderer,
  layout: SpreadsheetLayout,
  viewportManager: SpreadsheetViewportManager,
  selection: SpreadsheetSelectionState,
  dashedRange: ReturnType<typeof normalizeSelectionRange> | null,
  formulaReferenceRanges: SpreadsheetFormulaReferenceRange[] = [],
  autoFillPreviewRange: SpreadsheetRange | null = null,
  showFillHandle = selection.kind === 'cell' &&
    supportsSpreadsheetAutoFill(layout.model, selection.range),
): void {
  renderer.renderSelection(layout, viewportManager, selection.range, { showFillHandle });
  for (const reference of formulaReferenceRanges) {
    renderer.renderDashedRange(layout, viewportManager, reference.range, {
      stroke: reference.color,
      fill: colorWithAlpha(reference.color, 0.08),
    });
  }
  if (dashedRange) renderer.renderCopiedRange(layout, viewportManager, dashedRange);
  if (autoFillPreviewRange) {
    renderer.renderDashedRange(layout, viewportManager, autoFillPreviewRange, {
      stroke: '#2563EB',
      fill: 'rgba(37, 99, 235, 0.05)',
    });
  }
}

const FORMULA_REFERENCE_COLORS = ['#DC2626', '#16A34A', '#7C3AED', '#EA580C', '#0891B2'];
const SPREADSHEET_CELL_EDIT_MERGE_WINDOW_MS = 1000;

function renderSpreadsheetContextMenu(
  menu: HTMLElement,
  state: SpreadsheetContextMenuState,
  clientX: number,
  clientY: number,
): void {
  const items = spreadsheetContextMenuItems(state);
  menu.replaceChildren(
    ...items.map((item) => {
      if (item.kind === 'separator') {
        const separator = document.createElement('div');
        separator.className = 'coco-spreadsheet-context-menu-separator';
        return separator;
      }
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = item.label;
      button.dataset.spreadsheetAction = item.action;
      button.className = item.danger ? 'is-danger' : '';
      button.setAttribute('role', 'menuitem');
      return button;
    }),
  );
  menu.style.display = 'block';
  const width = menu.offsetWidth;
  const height = menu.offsetHeight;
  menu.style.left = `${Math.max(4, Math.min(clientX, window.innerWidth - width - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(clientY, window.innerHeight - height - 4))}px`;
}

function renderSpreadsheetHeaderFilterMenu(
  menu: HTMLElement,
  model: SpreadsheetSheetRenderModel,
  column: number,
  clientX: number,
  clientY: number,
  selectedValues?: readonly string[],
): Set<string> {
  const range = {
    startRow: 1,
    startColumn: 1,
    endRow: model.maxRow,
    endColumn: model.maxColumn,
  };
  const values = spreadsheetFilterValues(model, range, column);
  const selectedTokens = new Set(
    (selectedValues ?? values).map((value) => value || '__COCO_FILTER_BLANK__'),
  );
  const title = document.createElement('div');
  title.className = 'coco-spreadsheet-filter-menu-title';
  title.textContent = `Filter ${formatCellAddress(1, column).replace(/\d+$/, '')}`;
  const choices = document.createElement('div');
  choices.className = 'coco-spreadsheet-filter-menu-values';
  for (const value of values) {
    const label = document.createElement('label');
    label.className = 'coco-spreadsheet-filter-choice';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = selectedTokens.has(value || '__COCO_FILTER_BLANK__');
    checkbox.dataset.spreadsheetFilterChoice = value || '__COCO_FILTER_BLANK__';
    label.append(checkbox, document.createTextNode(value || '(blank)'));
    choices.append(label);
  }
  const actions = document.createElement('div');
  actions.className = 'coco-spreadsheet-filter-menu-actions';
  const sortAscending = document.createElement('button');
  sortAscending.type = 'button';
  sortAscending.textContent = 'A-Z';
  sortAscending.dataset.spreadsheetFilterSort = 'asc';
  sortAscending.setAttribute('role', 'menuitem');
  const sortDescending = document.createElement('button');
  sortDescending.type = 'button';
  sortDescending.textContent = 'Z-A';
  sortDescending.dataset.spreadsheetFilterSort = 'desc';
  sortDescending.setAttribute('role', 'menuitem');
  const apply = document.createElement('button');
  apply.type = 'button';
  apply.textContent = 'Apply';
  apply.dataset.spreadsheetFilterApply = 'true';
  apply.setAttribute('role', 'menuitem');
  const clear = document.createElement('button');
  clear.type = 'button';
  clear.textContent = 'Clear filter';
  clear.dataset.spreadsheetFilterClear = 'true';
  clear.setAttribute('role', 'menuitem');
  actions.append(sortAscending, sortDescending, apply, clear);
  menu.replaceChildren(title, choices, actions);
  menu.style.display = 'block';
  const width = menu.offsetWidth;
  const height = menu.offsetHeight;
  menu.style.left = `${Math.max(4, Math.min(clientX, window.innerWidth - width - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(clientY, window.innerHeight - height - 4))}px`;
  return selectedTokens;
}

function renderSpreadsheetSheetContextMenu(
  menu: HTMLElement,
  state: SpreadsheetSheetContextMenuState,
  clientX: number,
  clientY: number,
): void {
  const actions: Array<{
    label: string;
    action: SpreadsheetSheetContextMenuAction;
    danger?: boolean;
  }> = [
    { label: 'Rename sheet...', action: 'rename-sheet' },
    { label: 'Duplicate sheet', action: 'copy-sheet' },
    { label: 'Move sheet left', action: 'move-sheet-left' },
    { label: 'Move sheet right', action: 'move-sheet-right' },
    { label: 'Delete sheet', action: 'delete-sheet', danger: true },
  ];
  menu.replaceChildren(
    ...actions.map((item) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = item.label;
      button.dataset.spreadsheetSheetAction = item.action;
      button.className = item.danger ? 'is-danger' : '';
      button.setAttribute('role', 'menuitem');
      return button;
    }),
  );
  menu.style.display = 'block';
  const width = menu.offsetWidth;
  const height = menu.offsetHeight;
  menu.style.left = `${Math.max(4, Math.min(clientX, window.innerWidth - width - 4))}px`;
  menu.style.top = `${Math.max(4, Math.min(clientY, window.innerHeight - height - 4))}px`;
  menu.dataset.sheetName = state.sheetName;
}

function spreadsheetContextMenuItems(
  state: SpreadsheetContextMenuState,
): Array<
  | { kind: 'item'; label: string; action: SpreadsheetContextMenuAction; danger?: boolean }
  | { kind: 'separator' }
> {
  const selection = state.selection;
  const pasteItems = spreadsheetPasteContextMenuItems(state);
  if (selection.kind === 'row') {
    return [
      ...pasteItems,
      ...(pasteItems.length > 0 ? [{ kind: 'separator' as const }] : []),
      { kind: 'item', label: 'Insert row above', action: 'insert-row-above' },
      { kind: 'item', label: 'Insert row below', action: 'insert-row-below' },
      { kind: 'separator' },
      { kind: 'item', label: 'Delete selected rows', action: 'delete-rows', danger: true },
      { kind: 'item', label: 'Clear selected rows', action: 'clear' },
      { kind: 'separator' },
      { kind: 'item', label: 'Freeze rows above', action: 'freeze-rows' },
      { kind: 'item', label: 'Unfreeze panes', action: 'unfreeze-panes' },
      { kind: 'separator' },
      { kind: 'item', label: 'All borders', action: 'border-all' },
      { kind: 'item', label: 'Outer border', action: 'border-outer' },
      { kind: 'item', label: 'Clear borders', action: 'border-clear' },
      { kind: 'item', label: 'Clear format', action: 'clear-format' },
      { kind: 'separator' },
      { kind: 'item', label: 'Row settings...', action: 'cell-settings' },
    ];
  }
  if (selection.kind === 'column') {
    return [
      ...pasteItems,
      ...(pasteItems.length > 0 ? [{ kind: 'separator' as const }] : []),
      { kind: 'item', label: 'Insert column left', action: 'insert-column-left' },
      { kind: 'item', label: 'Insert column right', action: 'insert-column-right' },
      { kind: 'separator' },
      { kind: 'item', label: 'Delete selected columns', action: 'delete-columns', danger: true },
      { kind: 'item', label: 'Clear selected columns', action: 'clear' },
      { kind: 'separator' },
      { kind: 'item', label: 'Freeze columns left', action: 'freeze-columns' },
      { kind: 'item', label: 'Unfreeze panes', action: 'unfreeze-panes' },
      { kind: 'item', label: 'Filter this column...', action: 'filter-column' },
      { kind: 'separator' },
      { kind: 'item', label: 'All borders', action: 'border-all' },
      { kind: 'item', label: 'Outer border', action: 'border-outer' },
      { kind: 'item', label: 'Clear borders', action: 'border-clear' },
      { kind: 'item', label: 'Clear format', action: 'clear-format' },
      { kind: 'separator' },
      { kind: 'item', label: 'Column settings...', action: 'cell-settings' },
    ];
  }
  const mergeItems =
    selection.kind === 'cell'
      ? [
          { kind: 'item' as const, label: 'Merge cells', action: 'merge-cells' as const },
          { kind: 'item' as const, label: 'Unmerge cells', action: 'unmerge-cells' as const },
          { kind: 'separator' as const },
        ]
      : [];
  return [
    ...pasteItems,
    ...(pasteItems.length > 0 ? [{ kind: 'separator' as const }] : []),
    ...mergeItems,
    { kind: 'item', label: 'Insert row above', action: 'insert-row-above' },
    { kind: 'item', label: 'Insert row below', action: 'insert-row-below' },
    { kind: 'item', label: 'Insert column left', action: 'insert-column-left' },
    { kind: 'item', label: 'Insert column right', action: 'insert-column-right' },
    { kind: 'separator' },
    { kind: 'item', label: 'Freeze panes here', action: 'freeze-panes' },
    { kind: 'item', label: 'Freeze rows above', action: 'freeze-rows' },
    { kind: 'item', label: 'Freeze columns left', action: 'freeze-columns' },
    { kind: 'item', label: 'Unfreeze panes', action: 'unfreeze-panes' },
    { kind: 'separator' },
    { kind: 'item', label: 'All borders', action: 'border-all' },
    { kind: 'item', label: 'Outer border', action: 'border-outer' },
    { kind: 'item', label: 'Inner borders', action: 'border-inner' },
    { kind: 'item', label: 'Clear borders', action: 'border-clear' },
    { kind: 'separator' },
    { kind: 'item', label: 'Clear cells', action: 'clear' },
    { kind: 'item', label: 'Clear format', action: 'clear-format' },
    { kind: 'separator' },
    { kind: 'item', label: 'Cell settings...', action: 'cell-settings' },
  ];
}

function spreadsheetPasteContextMenuItems(
  state: SpreadsheetContextMenuState,
): Array<{ kind: 'item'; label: string; action: SpreadsheetContextMenuAction }> {
  const items: Array<{ kind: 'item'; label: string; action: SpreadsheetContextMenuAction }> = [
    { kind: 'item', label: 'Paste', action: 'paste' },
    { kind: 'item', label: 'Paste values', action: 'paste-values' },
  ];
  if (state.hasLocalClipboardPayload) {
    items.push({ kind: 'item', label: 'Paste formats', action: 'paste-formats' });
  }
  return items;
}

function spreadsheetPasteModeFromContextAction(
  action: SpreadsheetContextMenuAction,
): SpreadsheetPasteMode | null {
  switch (action) {
    case 'paste':
      return 'all';
    case 'paste-values':
      return 'values';
    case 'paste-formats':
      return 'formats';
    default:
      return null;
  }
}

function hideSpreadsheetContextMenu(menu: HTMLElement): void {
  menu.style.display = 'none';
  menu.replaceChildren();
}

function openSpreadsheetCellSettings(panel: HTMLElement): void {
  panel.style.display = 'block';
}

function hideSpreadsheetCellSettings(panel: HTMLElement): void {
  panel.style.display = 'none';
}

async function applySpreadsheetFormat(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  action: SpreadsheetFormatAction,
  render: () => void,
  value?: string,
): Promise<void> {
  const range = normalizeSelectionRange(selection.range);
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  const activeFormat = summarizeSpreadsheetSelectionFormat(runtime.activeSheet(), selection);
  if (action === 'bold') {
    await runtime.setRangeBold(
      range.startRow,
      range.startColumn,
      rowCount,
      columnCount,
      nextSpreadsheetToggleValue(activeFormat.font.bold),
    );
  } else if (action === 'italic') {
    await runtime.setRangeItalic(
      range.startRow,
      range.startColumn,
      rowCount,
      columnCount,
      nextSpreadsheetToggleValue(activeFormat.font.italic),
    );
  } else if (action === 'underline') {
    await runtime.setRangeUnderline(
      range.startRow,
      range.startColumn,
      rowCount,
      columnCount,
      nextSpreadsheetToggleValue(activeFormat.font.underline),
    );
  } else if (action === 'fontColor') {
    await runtime.setRangeFontColor(range.startRow, range.startColumn, rowCount, columnCount, {
      rgb: normalizeSpreadsheetHexColor(value, '#111827'),
    });
  } else if (action === 'fillColor') {
    await runtime.setRangeFillColor(range.startRow, range.startColumn, rowCount, columnCount, {
      rgb: normalizeSpreadsheetHexColor(value, '#FFFFFF'),
    });
  } else if (action === 'horizontalAlign') {
    const horizontal = value === 'center' || value === 'right' ? value : 'left';
    await runtime.setRangeHorizontalAlign(
      range.startRow,
      range.startColumn,
      rowCount,
      columnCount,
      horizontal,
    );
  } else if (action === 'verticalAlign') {
    const vertical = value === 'top' || value === 'center' ? value : 'bottom';
    await runtime.setRangeVerticalAlign(
      range.startRow,
      range.startColumn,
      rowCount,
      columnCount,
      vertical,
    );
  } else {
    await runtime.clearRangeFormat(range.startRow, range.startColumn, rowCount, columnCount);
  }
  render();
}

async function applySpreadsheetNumberFormat(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  code: string | undefined,
  render: () => void,
): Promise<void> {
  const range = normalizeSelectionRange(selection.range);
  await runtime.setRangeNumberFormat(
    range.startRow,
    range.startColumn,
    range.endRow - range.startRow + 1,
    range.endColumn - range.startColumn + 1,
    code,
  );
  render();
}

async function applySpreadsheetMergeCommand(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  render: () => void,
): Promise<void> {
  const merges = mergedRangesIntersectingSelection(runtime, selection);
  const range = normalizeSelectionRange(selection.range);
  const isSingleCell = range.startRow === range.endRow && range.startColumn === range.endColumn;
  if (merges.length > 0 && (isSingleCell || selectionMatchesAnyMerge(selection, merges))) {
    await runtime.unmergeRanges(merges.map((merge) => merge.address));
  } else {
    await mergeSpreadsheetSelection(runtime, selection);
  }
  render();
}

async function applySpreadsheetFreezePanes(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  mode: 'panes' | 'rows' | 'columns' | 'none',
  render: () => void,
): Promise<void> {
  const frozenRows =
    mode === 'columns' || mode === 'none' ? 0 : Math.max(0, selection.active.row - 1);
  const frozenColumns =
    mode === 'rows' || mode === 'none' ? 0 : Math.max(0, selection.active.column - 1);
  await runtime.setFreezePanes(frozenRows, frozenColumns);
  render();
}

async function applySpreadsheetBorder(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  mode: 'all' | 'outer' | 'inner' | 'clear',
  render: () => void,
): Promise<void> {
  await setSpreadsheetSelectionBorder(runtime, selection, mode);
  render();
}

async function setSpreadsheetSelectionBorder(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  mode: 'all' | 'outer' | 'inner' | 'clear',
): Promise<void> {
  const range = normalizeSelectionRange(selection.range);
  await runtime.setRangeBorder(
    range.startRow,
    range.startColumn,
    range.endRow - range.startRow + 1,
    range.endColumn - range.startColumn + 1,
    mode,
  );
}

async function mergeSpreadsheetSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
): Promise<void> {
  if (selection.kind !== 'cell') return;
  const range = normalizeSelectionRange(selection.range);
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  if (rowCount === 1 && columnCount === 1) return;
  const intersections = mergedRangesIntersectingSelection(runtime, selection);
  if (intersections.length > 0)
    await runtime.unmergeRanges(intersections.map((merge) => merge.address));
  await runtime.mergeRange(range.startRow, range.startColumn, rowCount, columnCount);
}

async function unmergeSpreadsheetSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
): Promise<void> {
  const merges = mergedRangesIntersectingSelection(runtime, selection);
  if (merges.length === 0) return;
  await runtime.unmergeRanges(merges.map((merge) => merge.address));
}

function mergedRangesIntersectingSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
): Array<{ address: string; range: SpreadsheetRange }> {
  const selectionRange = normalizeSelectionRange(selection.range);
  return [...runtime.activeSheet().mergeMasterByAddress.values()]
    .filter((merge) => rangesIntersect(selectionRange, merge.range))
    .map((merge) => ({ address: merge.address, range: merge.range }));
}

function selectionMatchesAnyMerge(
  selection: SpreadsheetSelectionState,
  merges: Array<{ range: SpreadsheetRange }>,
): boolean {
  const range = normalizeSelectionRange(selection.range);
  return merges.some(
    (merge) =>
      merge.range.startRow === range.startRow &&
      merge.range.startColumn === range.startColumn &&
      merge.range.endRow === range.endRow &&
      merge.range.endColumn === range.endColumn,
  );
}

function rangesIntersect(a: SpreadsheetRange, b: SpreadsheetRange): boolean {
  return !(
    a.endRow < b.startRow ||
    a.startRow > b.endRow ||
    a.endColumn < b.startColumn ||
    a.startColumn > b.endColumn
  );
}

function syncSpreadsheetFormattingToolbar(input: {
  runtime: SpreadsheetEditorRuntime;
  selection: SpreadsheetSelectionState;
  boldButton: HTMLButtonElement;
  italicButton: HTMLButtonElement;
  underlineButton: HTMLButtonElement;
  fontColorInput: HTMLInputElement;
  fillColorInput: HTMLInputElement;
  numberFormatSelect: HTMLSelectElement;
  mergeButton: HTMLButtonElement;
  alignButtons: HTMLButtonElement[];
  verticalAlignButtons: HTMLButtonElement[];
  cellSettings: HTMLElement;
  cellSettingsTitle: HTMLElement;
  cellSettingsFontColor: HTMLInputElement;
  cellSettingsFillColor: HTMLInputElement;
  cellSettingsNumberFormat: HTMLSelectElement;
  cellSettingsActionButtons: HTMLButtonElement[];
  cellSettingsAlignButtons: HTMLButtonElement[];
  cellSettingsVerticalAlignButtons: HTMLButtonElement[];
}): void {
  const format = summarizeSpreadsheetSelectionFormat(input.runtime.activeSheet(), input.selection);
  input.cellSettingsTitle.textContent =
    input.selection.kind === 'column'
      ? 'Column settings'
      : input.selection.kind === 'row'
        ? 'Row settings'
        : 'Cell settings';
  syncSpreadsheetToggleButton(input.boldButton, format.font.bold);
  syncSpreadsheetToggleButton(input.italicButton, format.font.italic);
  syncSpreadsheetToggleButton(input.underlineButton, format.font.underline);
  setSpreadsheetColorInputValue(input.fontColorInput, format.font.color, '#111827');
  setSpreadsheetColorInputValue(input.fillColorInput, format.fill.foregroundColor, '#FFFFFF');
  setSpreadsheetSelectValue(input.numberFormatSelect, format.numberFormatCode ?? '');
  setSpreadsheetColorInputValue(input.cellSettingsFontColor, format.font.color, '#111827');
  setSpreadsheetColorInputValue(
    input.cellSettingsFillColor,
    format.fill.foregroundColor,
    '#FFFFFF',
  );
  setSpreadsheetSelectValue(input.cellSettingsNumberFormat, format.numberFormatCode ?? '');
  const merges = mergedRangesIntersectingSelection(input.runtime, input.selection);
  input.mergeButton.disabled = input.selection.kind !== 'cell';
  input.mergeButton.classList.toggle('is-active', merges.length > 0);
  input.cellSettings.querySelector<HTMLButtonElement>(
    '[data-cell-settings-action="merge"]',
  )!.disabled = input.selection.kind !== 'cell';
  const stateByAction: Record<string, SpreadsheetMixedValue<boolean>> = {
    bold: format.font.bold,
    italic: format.font.italic,
    underline: format.font.underline,
  };
  for (const button of input.cellSettingsActionButtons) {
    const action = button.dataset.cellSettingsAction;
    syncSpreadsheetToggleButton(button, action ? stateByAction[action] : false);
  }
  for (const button of input.alignButtons) {
    syncSpreadsheetAlignButton(
      button,
      format.alignment.horizontal,
      button.dataset.spreadsheetAlign,
    );
  }
  for (const button of input.verticalAlignButtons) {
    syncSpreadsheetAlignButton(
      button,
      format.alignment.vertical,
      button.dataset.spreadsheetVerticalAlign,
    );
  }
  for (const button of input.cellSettingsAlignButtons) {
    syncSpreadsheetAlignButton(
      button,
      format.alignment.horizontal,
      button.dataset.cellSettingsAlign,
    );
  }
  for (const button of input.cellSettingsVerticalAlignButtons) {
    syncSpreadsheetAlignButton(
      button,
      format.alignment.vertical,
      button.dataset.cellSettingsVerticalAlign,
    );
  }
}

function syncSpreadsheetToggleButton(
  button: HTMLButtonElement,
  value: SpreadsheetMixedValue<boolean> | undefined,
): void {
  button.classList.toggle('is-active', value === true);
  button.classList.toggle('is-mixed', value === 'mixed');
}

function syncSpreadsheetAlignButton(
  button: HTMLButtonElement,
  value: SpreadsheetMixedValue<'left' | 'center' | 'right' | 'top' | 'bottom'>,
  buttonValue: string | undefined,
): void {
  button.classList.toggle('is-active', value !== 'mixed' && value === buttonValue);
  button.classList.toggle('is-mixed', value === 'mixed');
}

function nextSpreadsheetToggleValue(value: SpreadsheetMixedValue<boolean>): boolean {
  return value === true ? false : true;
}

function setSpreadsheetSelectValue(
  select: HTMLSelectElement,
  value: SpreadsheetMixedValue<string | undefined>,
): void {
  select.querySelector('option[data-runtime-value]')?.remove();
  select.classList.toggle('is-mixed', value === 'mixed');
  if (value === 'mixed') {
    const option = new Option('Mixed', '', true, true);
    option.dataset.runtimeValue = 'true';
    select.prepend(option);
    select.value = '';
    return;
  }
  const normalizedValue = value ?? '';
  if (!Array.from(select.options).some((option) => option.value === normalizedValue)) {
    const option = new Option(normalizedValue || 'General', normalizedValue, true, true);
    option.dataset.runtimeValue = 'true';
    select.prepend(option);
  }
  select.value = normalizedValue;
}

function setSpreadsheetColorInputValue(
  input: HTMLInputElement,
  value: SpreadsheetMixedValue<{ rgb?: string } | undefined>,
  fallback: string,
): void {
  input.classList.toggle('is-mixed', value === 'mixed');
  input.value = value === 'mixed' ? fallback : normalizeSpreadsheetHexColor(value?.rgb, fallback);
}

function normalizeSpreadsheetHexColor(value: string | undefined, fallback: string): string {
  if (value && /^#[0-9a-fA-F]{6}$/.test(value)) return value.toUpperCase();
  return fallback;
}

function shiftSpreadsheetColumnFilters(
  filters: Map<number, string[]> | undefined,
  startColumn: number,
  delta: number,
): void {
  if (!filters || delta === 0) return;
  const endColumn = delta < 0 ? startColumn - delta - 1 : startColumn - 1;
  const next = new Map<number, string[]>();
  for (const [column, values] of filters) {
    if (delta < 0 && column >= startColumn && column <= endColumn) continue;
    const nextColumn =
      column >= startColumn && (delta > 0 || column > endColumn) ? column + delta : column;
    if (nextColumn > 0) next.set(nextColumn, values);
  }
  filters.clear();
  for (const [column, values] of next) filters.set(column, values);
}

async function runSpreadsheetContextMenuAction(
  runtime: SpreadsheetEditorRuntime,
  action: SpreadsheetContextMenuAction,
  selection: SpreadsheetSelectionState,
  layout: SpreadsheetLayout,
  activeColumnFilters?: Map<number, string[]>,
): Promise<SpreadsheetSelectionState> {
  if (action === 'cell-settings') return selection;
  const range = normalizeSelectionRange(selection.range);
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  if (action === 'insert-row-above') {
    await runtime.insertRows(range.startRow, rowCount);
    return createRowRangeSelection(range.startRow, range.startRow + rowCount - 1, {
      columnCount: runtime.activeSheet().maxColumn,
    });
  }
  if (action === 'insert-row-below') {
    await runtime.insertRows(range.endRow + 1, rowCount);
    return createRowRangeSelection(range.endRow + 1, range.endRow + rowCount, {
      columnCount: runtime.activeSheet().maxColumn,
    });
  }
  if (action === 'delete-rows') {
    await runtime.deleteRows(range.startRow, rowCount);
    const sheet = runtime.activeSheet();
    return selectionAfterRowDeletion(
      selection,
      range.startRow,
      rowCount,
      { rowCount: sheet.maxRow, columnCount: sheet.maxColumn },
      sheet,
    );
  }
  if (action === 'insert-column-left') {
    await runtime.insertColumns(range.startColumn, columnCount);
    shiftSpreadsheetColumnFilters(activeColumnFilters, range.startColumn, columnCount);
    return createColumnRangeSelection(range.startColumn, range.startColumn + columnCount - 1, {
      rowCount: runtime.activeSheet().maxRow,
    });
  }
  if (action === 'insert-column-right') {
    await runtime.insertColumns(range.endColumn + 1, columnCount);
    shiftSpreadsheetColumnFilters(activeColumnFilters, range.endColumn + 1, columnCount);
    return createColumnRangeSelection(range.endColumn + 1, range.endColumn + columnCount, {
      rowCount: runtime.activeSheet().maxRow,
    });
  }
  if (action === 'delete-columns') {
    await runtime.deleteColumns(range.startColumn, columnCount);
    shiftSpreadsheetColumnFilters(activeColumnFilters, range.startColumn, -columnCount);
    const sheet = runtime.activeSheet();
    return selectionAfterColumnDeletion(
      selection,
      range.startColumn,
      columnCount,
      { rowCount: sheet.maxRow, columnCount: sheet.maxColumn },
      sheet,
    );
  }
  if (action === 'merge-cells') {
    await mergeSpreadsheetSelection(runtime, selection);
    return selection;
  }
  if (action === 'unmerge-cells') {
    await unmergeSpreadsheetSelection(runtime, selection);
    return selection;
  }
  if (
    action === 'freeze-panes' ||
    action === 'freeze-rows' ||
    action === 'freeze-columns' ||
    action === 'unfreeze-panes'
  ) {
    const mode =
      action === 'freeze-panes'
        ? 'panes'
        : action === 'freeze-rows'
          ? 'rows'
          : action === 'freeze-columns'
            ? 'columns'
            : 'none';
    await applySpreadsheetFreezePanes(runtime, selection, mode, () => undefined);
    return selection;
  }
  if (
    action === 'border-all' ||
    action === 'border-outer' ||
    action === 'border-inner' ||
    action === 'border-clear'
  ) {
    const mode =
      action === 'border-all'
        ? 'all'
        : action === 'border-outer'
          ? 'outer'
          : action === 'border-inner'
            ? 'inner'
            : 'clear';
    await setSpreadsheetSelectionBorder(runtime, selection, mode);
    return selection;
  }
  if (action === 'clear-format') {
    await runtime.clearRangeFormat(range.startRow, range.startColumn, rowCount, columnCount);
    return selection;
  }
  await runtime.clearRange(range.startRow, range.startColumn, rowCount, columnCount);
  return selection;
}

function hitSpreadsheetResizeHandle(
  host: HTMLElement,
  layout: SpreadsheetLayout,
  viewportManager: SpreadsheetViewportManager,
  event: MouseEvent,
): SpreadsheetResizeHandle | null {
  const rect = host.getBoundingClientRect();
  return hitSpreadsheetResizeHandleAtPoint(layout, viewportManager.viewports, {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top,
  });
}

function updateSpreadsheetResizeGuide(guide: HTMLElement, drag: SpreadsheetResizeDragState): void {
  guide.classList.toggle('is-column', drag.kind === 'column');
  guide.classList.toggle('is-row', drag.kind === 'row');
  guide.style.display = 'block';
  if (drag.kind === 'column') {
    guide.style.left = `${Math.round(drag.currentGuidePosition)}px`;
    guide.style.top = '';
  } else {
    guide.style.left = '';
    guide.style.top = `${Math.round(drag.currentGuidePosition)}px`;
  }
}

function hideSpreadsheetResizeGuide(guide: HTMLElement): void {
  guide.style.display = 'none';
  guide.classList.remove('is-column', 'is-row');
}

async function commitSpreadsheetResizeDrag(
  runtime: SpreadsheetEditorRuntime,
  drag: SpreadsheetResizeDragState,
  render: () => void,
): Promise<void> {
  if (drag.currentSize === drag.startSize) {
    render();
    return;
  }
  if (drag.kind === 'column') await runtime.setColumnWidth(drag.index, drag.currentSize);
  else await runtime.setRowHeight(drag.index, (drag.currentSize * 72) / 96);
  render();
}

function containsSpreadsheetPoint(
  rect: { x: number; y: number; width: number; height: number },
  point: { x: number; y: number },
): boolean {
  return (
    point.x >= rect.x &&
    point.y >= rect.y &&
    point.x < rect.x + rect.width &&
    point.y < rect.y + rect.height
  );
}

const SPREADSHEET_FORMULA_FUNCTIONS: SpreadsheetFormulaFunctionInfo[] = [
  { name: 'SUM', signature: 'SUM(number1, [number2], ...)' },
  { name: 'AVERAGE', signature: 'AVERAGE(number1, [number2], ...)' },
  { name: 'COUNT', signature: 'COUNT(value1, [value2], ...)' },
  { name: 'COUNTA', signature: 'COUNTA(value1, [value2], ...)' },
  { name: 'MIN', signature: 'MIN(number1, [number2], ...)' },
  { name: 'MAX', signature: 'MAX(number1, [number2], ...)' },
  { name: 'IF', signature: 'IF(logical_test, value_if_true, [value_if_false])' },
  { name: 'IFERROR', signature: 'IFERROR(value, value_if_error)' },
  { name: 'AND', signature: 'AND(logical1, [logical2], ...)' },
  { name: 'OR', signature: 'OR(logical1, [logical2], ...)' },
  { name: 'NOT', signature: 'NOT(logical)' },
  { name: 'SUMIF', signature: 'SUMIF(range, criteria, [sum_range])' },
  { name: 'COUNTIF', signature: 'COUNTIF(range, criteria)' },
  { name: 'AVERAGEIF', signature: 'AVERAGEIF(range, criteria, [average_range])' },
  { name: 'SUMIFS', signature: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)' },
  { name: 'COUNTIFS', signature: 'COUNTIFS(criteria_range1, criteria1, ...)' },
  { name: 'AVERAGEIFS', signature: 'AVERAGEIFS(average_range, criteria_range1, criteria1, ...)' },
  {
    name: 'VLOOKUP',
    signature: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
  },
  {
    name: 'HLOOKUP',
    signature: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
  },
  { name: 'INDEX', signature: 'INDEX(array, row_num, [column_num])' },
  { name: 'MATCH', signature: 'MATCH(lookup_value, lookup_array, [match_type])' },
  { name: 'CONCAT', signature: 'CONCAT(text1, [text2], ...)' },
  { name: 'CONCATENATE', signature: 'CONCATENATE(text1, [text2], ...)' },
  { name: 'TEXTJOIN', signature: 'TEXTJOIN(delimiter, ignore_empty, text1, ...)' },
  { name: 'LEN', signature: 'LEN(text)' },
  { name: 'TRIM', signature: 'TRIM(text)' },
  { name: 'LEFT', signature: 'LEFT(text, [num_chars])' },
  { name: 'RIGHT', signature: 'RIGHT(text, [num_chars])' },
  { name: 'MID', signature: 'MID(text, start_num, num_chars)' },
  { name: 'UPPER', signature: 'UPPER(text)' },
  { name: 'LOWER', signature: 'LOWER(text)' },
  { name: 'VALUE', signature: 'VALUE(text)' },
  { name: 'ROUND', signature: 'ROUND(number, num_digits)' },
  { name: 'ABS', signature: 'ABS(number)' },
  { name: 'INT', signature: 'INT(number)' },
  { name: 'MOD', signature: 'MOD(number, divisor)' },
  { name: 'TRUE', signature: 'TRUE()' },
  { name: 'FALSE', signature: 'FALSE()' },
];

function updateSpreadsheetFormulaSuggest(
  target: HTMLInputElement | HTMLTextAreaElement,
  panel: HTMLElement,
  previous: SpreadsheetFormulaSuggestState | null,
): SpreadsheetFormulaSuggestState | null {
  const span = spreadsheetFormulaFunctionPrefixSpan(
    target.value,
    target.selectionStart ?? target.value.length,
  );
  if (!span) {
    hideSpreadsheetFormulaSuggest(panel);
    return null;
  }
  const prefix = target.value.slice(span.start, span.end).toUpperCase();
  const items = SPREADSHEET_FORMULA_FUNCTIONS.filter((item) => item.name.startsWith(prefix)).slice(
    0,
    8,
  );
  if (items.length === 0) {
    hideSpreadsheetFormulaSuggest(panel);
    return null;
  }
  const state: SpreadsheetFormulaSuggestState = {
    target,
    replaceStart: span.start,
    replaceEnd: span.end,
    activeIndex: Math.min(previous?.activeIndex ?? 0, items.length - 1),
    items,
  };
  renderSpreadsheetFormulaSuggest(panel, state);
  return state;
}

function handleSpreadsheetFormulaSuggestKeydown(
  event: KeyboardEvent,
  state: SpreadsheetFormulaSuggestState | null,
  panel: HTMLElement,
  setState: (state: SpreadsheetFormulaSuggestState | null) => void,
): boolean {
  if (!state) return false;
  if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
    event.preventDefault();
    const delta = event.key === 'ArrowDown' ? 1 : -1;
    const nextState = {
      ...state,
      activeIndex: (state.activeIndex + delta + state.items.length) % state.items.length,
    };
    renderSpreadsheetFormulaSuggest(panel, nextState);
    setState(nextState);
    return true;
  }
  if (event.key === 'Enter' || event.key === 'Tab') {
    event.preventDefault();
    setState(acceptSpreadsheetFormulaSuggestion(state, panel, state.activeIndex));
    return true;
  }
  if (event.key === 'Escape') {
    hideSpreadsheetFormulaSuggest(panel);
    setState(null);
    return true;
  }
  return false;
}

function acceptSpreadsheetFormulaSuggestion(
  state: SpreadsheetFormulaSuggestState,
  panel: HTMLElement,
  index: number,
): null {
  const item = state.items[Math.max(0, Math.min(index, state.items.length - 1))];
  if (!item) return null;
  const insertion = `${item.name}(`;
  const nextValue =
    state.target.value.slice(0, state.replaceStart) +
    insertion +
    state.target.value.slice(state.replaceEnd);
  state.target.value = nextValue;
  const caret = state.replaceStart + insertion.length;
  state.target.focus();
  state.target.setSelectionRange(caret, caret);
  hideSpreadsheetFormulaSuggest(panel);
  return null;
}

function renderSpreadsheetFormulaSuggest(
  panel: HTMLElement,
  state: SpreadsheetFormulaSuggestState,
): void {
  const rect = state.target.getBoundingClientRect();
  panel.style.display = 'block';
  panel.style.left = `${rect.left}px`;
  panel.style.top = `${rect.bottom + 4}px`;
  panel.replaceChildren(
    ...state.items.map((item, index) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.dataset.formulaFunction = item.name;
      button.dataset.formulaFunctionIndex = String(index);
      button.className = index === state.activeIndex ? 'is-active' : '';
      button.setAttribute('role', 'option');
      button.setAttribute('aria-selected', String(index === state.activeIndex));
      button.innerHTML = `
        <div class="coco-spreadsheet-formula-suggest-name">${item.name}</div>
        <div class="coco-spreadsheet-formula-suggest-signature">${item.signature}</div>
      `;
      return button;
    }),
  );
}

function hideSpreadsheetFormulaSuggest(panel: HTMLElement): void {
  panel.style.display = 'none';
  panel.replaceChildren();
}

function spreadsheetFormulaFunctionPrefixSpan(
  value: string,
  caret: number,
): { start: number; end: number } | null {
  if (!value.trimStart().startsWith('=')) return null;
  let start = caret;
  while (start > 0 && /[A-Za-z]/.test(value[start - 1]!)) start--;
  if (start === caret) return null;
  if (start > 0 && /[A-Za-z0-9_]/.test(value[start - 1]!)) return null;
  const prefix = value.slice(start, caret);
  if (prefix.length < 1) return null;
  return { start, end: caret };
}

function activeFormulaReferenceEditTarget(
  editor: HTMLTextAreaElement,
  formulaInput: HTMLInputElement,
): HTMLInputElement | HTMLTextAreaElement | null {
  const active = document.activeElement;
  if (active === editor && editor.style.display !== 'none' && editor.value.trim().startsWith('='))
    return editor;
  if (active === formulaInput && formulaInput.value.trim().startsWith('=')) return formulaInput;
  return null;
}

function startSpreadsheetFormulaReferenceEdit(
  target: HTMLInputElement | HTMLTextAreaElement,
  row: number,
  column: number,
): SpreadsheetFormulaReferenceEditState {
  const [replaceStart, replaceEnd] = spreadsheetFormulaReferenceReplaceSpan(
    target.value,
    target.selectionStart ?? target.value.length,
    target.selectionEnd ?? target.value.length,
  );
  const state: SpreadsheetFormulaReferenceEditState = {
    target,
    anchor: { row, column },
    originalValue: target.value,
    replaceStart,
    replaceEnd,
  };
  updateSpreadsheetFormulaReferenceEdit(state, row, column);
  return state;
}

function updateSpreadsheetFormulaReferenceEdit(
  state: SpreadsheetFormulaReferenceEditState,
  row: number,
  column: number,
): void {
  const reference = formatSpreadsheetFormulaReferenceRange(state.anchor, { row, column });
  const nextValue =
    state.originalValue.slice(0, state.replaceStart) +
    reference +
    state.originalValue.slice(state.replaceEnd);
  state.target.value = nextValue;
  const caret = state.replaceStart + reference.length;
  state.target.focus();
  state.target.setSelectionRange(caret, caret);
}

function spreadsheetFormulaReferenceReplaceSpan(
  value: string,
  selectionStart: number,
  selectionEnd: number,
): [number, number] {
  if (selectionStart !== selectionEnd) return [selectionStart, selectionEnd];
  let start = selectionStart;
  let end = selectionEnd;
  while (start > 0 && isFormulaReferenceTokenChar(value[start - 1]!)) start--;
  while (end < value.length && isFormulaReferenceTokenChar(value[end]!)) end++;
  return [start, end];
}

function isFormulaReferenceTokenChar(char: string): boolean {
  return /[A-Za-z0-9_$:]/.test(char);
}

function formatSpreadsheetFormulaReferenceRange(
  start: { row: number; column: number },
  end: { row: number; column: number },
): string {
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startColumn = Math.min(start.column, end.column);
  const endColumn = Math.max(start.column, end.column);
  const first = formatCellAddress(startRow, startColumn);
  const last = formatCellAddress(endRow, endColumn);
  return first === last ? first : `${first}:${last}`;
}

function formulaReferenceRangesFromText(text: string): SpreadsheetFormulaReferenceRange[] {
  if (!text.trim().startsWith('=')) return [];
  return collectSpreadsheetFormulaReferences(text).flatMap((reference, index) => {
    const range = formulaReferenceToRange(reference);
    if (!range) return [];
    return [{ range, color: FORMULA_REFERENCE_COLORS[index % FORMULA_REFERENCE_COLORS.length]! }];
  });
}

function formulaReferenceToRange(
  reference: ReturnType<typeof collectSpreadsheetFormulaReferences>[number],
): SpreadsheetRange | null {
  try {
    if (reference.kind === 'cell') {
      const cell = parseCellAddress(reference.address);
      return {
        startRow: cell.row,
        startColumn: cell.column,
        endRow: cell.row,
        endColumn: cell.column,
      };
    }
    const range = parseRangeAddress(`${reference.startAddress}:${reference.endAddress}`);
    return {
      startRow: range.startRow,
      startColumn: range.startColumn,
      endRow: range.startRow + range.rowCount - 1,
      endColumn: range.startColumn + range.columnCount - 1,
    };
  } catch {
    return null;
  }
}

function colorWithAlpha(color: string, alpha: number): string {
  const normalized = /^#?([0-9a-f]{6})$/i.exec(color)?.[1];
  if (!normalized) return color;
  const value = Number.parseInt(normalized, 16);
  const red = (value >> 16) & 255;
  const green = (value >> 8) & 255;
  const blue = value & 255;
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function syncSpreadsheetFormulaBar(
  nameBox: HTMLInputElement,
  formulaInput: HTMLInputElement,
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
): void {
  const active = selection.active;
  const address = formatCellAddress(active.row, active.column);
  const cell = getRenderCell(runtime.activeSheet(), active.row, active.column);
  nameBox.value = address;
  formulaInput.value = cell?.formula ? `=${cell.formula}` : (cell?.displayText ?? '');
}

function selectionFromSpreadsheetHistoryResult(
  affectedRanges: string[],
  runtime: SpreadsheetEditorRuntime,
): SpreadsheetSelectionState | null {
  const activeSheetName = runtime.activeSheetName();
  for (const affectedRange of affectedRanges) {
    try {
      const qualified = splitSheetQualifiedAddress(affectedRange);
      if (qualified.sheetName && qualified.sheetName !== activeSheetName) continue;
      const range = parseRangeAddress(qualified.address);
      const startRow = Math.min(range.startRow, runtime.activeSheet().maxRow);
      const startColumn = Math.min(range.startColumn, runtime.activeSheet().maxColumn);
      const endRow = Math.min(range.startRow + range.rowCount - 1, runtime.activeSheet().maxRow);
      const endColumn = Math.min(
        range.startColumn + range.columnCount - 1,
        runtime.activeSheet().maxColumn,
      );
      return {
        kind: 'cell',
        active: { row: startRow, column: startColumn },
        range: { startRow, startColumn, endRow, endColumn },
      };
    } catch {
      continue;
    }
  }
  return null;
}

async function commitSpreadsheetFormulaBar(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  rawValue: string,
  render: () => void,
): Promise<void> {
  const address = formatCellAddress(selection.active.row, selection.active.column);
  const historyOptions = {
    mergeUndo: true,
    mergeWindowMs: SPREADSHEET_CELL_EDIT_MERGE_WINDOW_MS,
  };
  if (rawValue.startsWith('=')) {
    await runtime.setCellFormula(address, rawValue.slice(1), historyOptions);
  } else {
    await runtime.setCellValue(address, coerceSpreadsheetFormulaBarValue(rawValue), historyOptions);
  }
  render();
}

function parseSpreadsheetNameBox(value: string): SpreadsheetSelectionState | null {
  try {
    const trimmed = value.trim();
    if (trimmed.includes(':')) {
      const range = parseRangeAddress(trimmed);
      return {
        kind: 'cell',
        active: { row: range.startRow, column: range.startColumn },
        range: {
          startRow: range.startRow,
          startColumn: range.startColumn,
          endRow: range.startRow + range.rowCount - 1,
          endColumn: range.startColumn + range.columnCount - 1,
        },
      };
    }
    const parsed = parseCellAddress(trimmed);
    return createSelection(parsed.row, parsed.column);
  } catch {
    return null;
  }
}

function clampSpreadsheetSelection(
  selection: SpreadsheetSelectionState,
  layout: SpreadsheetLayout,
): SpreadsheetSelectionState {
  const range = normalizeSelectionRange(selection.range);
  const startRow = Math.max(1, Math.min(range.startRow, layout.rowCount));
  const startColumn = Math.max(1, Math.min(range.startColumn, layout.columnCount));
  const endRow = Math.max(1, Math.min(range.endRow, layout.rowCount));
  const endColumn = Math.max(1, Math.min(range.endColumn, layout.columnCount));
  return {
    ...selection,
    active: {
      row: Math.max(1, Math.min(selection.active.row, layout.rowCount)),
      column: Math.max(1, Math.min(selection.active.column, layout.columnCount)),
    },
    range: normalizeSelectionRange({ startRow, startColumn, endRow, endColumn }),
  };
}

function scrollSpreadsheetSelectionIntoView(
  layout: SpreadsheetLayout,
  viewportManager: SpreadsheetViewportManager,
  selection: SpreadsheetSelectionState,
  currentScroll: { x: number; y: number },
): { x: number; y: number } {
  const range = normalizeSelectionRange(selection.range);
  const rect = layout.getRangeLogicalRect(range);
  const frozenWidth = layout.getColumnX(viewportManager.freeze.frozenColumns + 1);
  const frozenHeight = layout.getRowY(viewportManager.freeze.frozenRows + 1);
  const viewportWidth = Math.max(
    0,
    viewportManager.viewportSize.width - layout.rowHeaderWidth - frozenWidth,
  );
  const viewportHeight = Math.max(
    0,
    viewportManager.viewportSize.height - layout.columnHeaderHeight - frozenHeight,
  );
  let x = currentScroll.x;
  let y = currentScroll.y;
  if (range.startColumn > viewportManager.freeze.frozenColumns) {
    const targetX = Math.max(0, rect.x - frozenWidth);
    if (targetX < x) x = targetX;
    else if (targetX + rect.width > x + viewportWidth)
      x = Math.max(0, targetX + rect.width - viewportWidth);
  }
  if (range.startRow > viewportManager.freeze.frozenRows) {
    const targetY = Math.max(0, rect.y - frozenHeight);
    if (targetY < y) y = targetY;
    else if (targetY + rect.height > y + viewportHeight)
      y = Math.max(0, targetY + rect.height - viewportHeight);
  }
  return { x, y };
}

function renderSpreadsheetFindCount(
  target: HTMLElement,
  matches: readonly SpreadsheetFindMatch[],
  activeIndex: number,
): void {
  target.textContent =
    matches.length === 0 ? '0 results' : `${Math.max(1, activeIndex + 1)} / ${matches.length}`;
}

function coerceSpreadsheetFormulaBarValue(value: string): string | number | boolean | null {
  if (value === '') return null;
  const number = Number(value);
  if (Number.isFinite(number) && value.trim() !== '') return number;
  if (value === 'TRUE') return true;
  if (value === 'FALSE') return false;
  return value;
}

function renderSpreadsheetTabs(
  container: HTMLElement,
  runtime: SpreadsheetEditorRuntime,
  renamingSheetName: string | null,
): void {
  const active = runtime.activeSheetName();
  container.replaceChildren(
    ...runtime.model().sheets.map((sheet) => {
      if (sheet.name === renamingSheetName) {
        const input = document.createElement('input');
        input.className = 'coco-spreadsheet-tab-rename';
        input.type = 'text';
        input.value = sheet.name;
        input.dataset.sheetName = sheet.name;
        input.dataset.sheetRenameInput = 'true';
        input.setAttribute('aria-label', 'Worksheet name');
        input.spellcheck = false;
        return input;
      }
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `coco-spreadsheet-tab${sheet.name === active ? ' is-active' : ''}`;
      button.dataset.sheetName = sheet.name;
      button.textContent = sheet.name;
      button.title = sheet.name;
      return button;
    }),
    (() => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'coco-spreadsheet-tab is-add';
      button.dataset.spreadsheetSheetAdd = 'true';
      button.title = 'Add worksheet';
      button.setAttribute('aria-label', 'Add worksheet');
      button.textContent = '+';
      return button;
    })(),
  );
}

function nextSpreadsheetSheetName(sheetNames: readonly string[], baseName: string): string {
  const names = new Set(sheetNames);
  const base = baseName.slice(0, 31) || 'Sheet';
  if (!names.has(base)) return base;
  for (let index = 2; index < 10_000; index++) {
    const suffix = ` (${index})`;
    const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    if (!names.has(candidate)) return candidate;
  }
  throw new Error('Spreadsheet editor: could not create a unique worksheet name');
}

interface SpreadsheetSheetViewState {
  selection: SpreadsheetSelectionState;
  scroll: { x: number; y: number };
}

async function copySpreadsheetSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  inspector: HTMLPreElement,
  context: { fixturePath: string; fixtureSize: number; scroll: { x: number; y: number } },
): Promise<{
  range: ReturnType<typeof normalizeSelectionRange>;
  payload: SpreadsheetClipboardPayload;
} | null> {
  const range = normalizeSelectionRange(selection.range);
  const payload = buildSpreadsheetClipboardPayload(runtime.activeSheet(), range);
  try {
    await navigator.clipboard.writeText(payload.text);
    renderSpreadsheetInspector(inspector, {
      ...context,
      runtime,
      selection,
      clipboardStatus: { action: 'copy', ok: true, cells: selectedCellCount(selection) },
    });
    return { range, payload };
  } catch (err) {
    renderSpreadsheetInspector(inspector, {
      ...context,
      runtime,
      selection,
      clipboardStatus: { action: 'copy', ok: false, error: errorMessage(err) },
    });
    return null;
  }
}

async function pasteSpreadsheetSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  mode: SpreadsheetPasteMode,
  onPasted: (selection: SpreadsheetSelectionState) => void,
  inspector: HTMLPreElement,
  context: { fixturePath: string; fixtureSize: number; scroll: { x: number; y: number } },
  localClipboardPayload: SpreadsheetClipboardPayload | null,
): Promise<void> {
  try {
    const text = await navigator.clipboard.readText();
    const payload = clipboardPayloadMatchesText(localClipboardPayload, text)
      ? localClipboardPayload
      : null;
    if (mode === 'formats' && !payload) {
      throw new Error('Paste formats requires a Coco internal clipboard payload');
    }
    const values = payload?.values ?? parseTsv(text);
    if (values.length === 0) return;
    const rowCount = values.length;
    const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
    await runtime.pasteRange(selection.active.row, selection.active.column, values, {
      mode,
      merges: mode === 'all' ? (payload?.merges ?? []) : [],
      styleIds: mode === 'all' || mode === 'formats' ? (payload?.styleIds ?? []) : [],
    });
    onPasted(
      expandSelectionForMergedRanges(
        {
          kind: 'cell',
          active: selection.active,
          range: {
            startRow: selection.active.row,
            startColumn: selection.active.column,
            endRow: selection.active.row + rowCount - 1,
            endColumn: selection.active.column + columnCount - 1,
          },
        },
        runtime.activeSheet(),
      ),
    );
    renderSpreadsheetInspector(inspector, {
      ...context,
      runtime,
      selection,
      clipboardStatus: {
        action: 'paste',
        ok: true,
        mode,
        cells: rowCount * columnCount,
        merges: mode === 'all' ? (payload?.merges.length ?? 0) : 0,
        formatted:
          (mode === 'all' || mode === 'formats') &&
          (payload?.styleIds.flat().some((styleId) => styleId !== undefined) ?? false),
        source: payload ? 'coco' : 'tsv',
      },
    });
  } catch (err) {
    renderSpreadsheetInspector(inspector, {
      ...context,
      runtime,
      selection,
      clipboardStatus: { action: 'paste', ok: false, error: errorMessage(err) },
    });
  }
}

function clipboardPayloadMatchesText(
  payload: SpreadsheetClipboardPayload | null,
  text: string,
): payload is SpreadsheetClipboardPayload {
  return !!payload && normalizeClipboardText(payload.text) === normalizeClipboardText(text);
}

function normalizeClipboardText(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

async function clearSpreadsheetSelection(
  runtime: SpreadsheetEditorRuntime,
  selection: SpreadsheetSelectionState,
  render: () => void,
): Promise<void> {
  const range = normalizeSelectionRange(selection.range);
  await runtime.clearRange(
    range.startRow,
    range.startColumn,
    range.endRow - range.startRow + 1,
    range.endColumn - range.startColumn + 1,
  );
  render();
}

function openSpreadsheetCellEditor(input: {
  editor: HTMLTextAreaElement;
  layout: SpreadsheetLayout;
  viewportManager: SpreadsheetViewportManager;
  row: number;
  column: number;
  initialValue: string;
  selectText: boolean;
}): void {
  const viewport = input.viewportManager.viewports.find((candidate) => {
    if (
      candidate.kind !== 'main' &&
      candidate.kind !== 'frozenTop' &&
      candidate.kind !== 'frozenLeft' &&
      candidate.kind !== 'frozenTopLeft'
    )
      return false;
    return (
      input.row >= candidate.rows.start &&
      input.row <= candidate.rows.end &&
      input.column >= candidate.columns.start &&
      input.column <= candidate.columns.end
    );
  });
  if (!viewport) return;
  const merge = input.layout.model.mergeMasterByAddress.get(
    formatCellAddress(input.row, input.column),
  );
  const cellRect = merge
    ? logicalRectToSpreadsheetEditorScreen(input.layout.getRangeLogicalRect(merge.range), viewport)
    : input.viewportManager.cellScreenRect(input.row, input.column, viewport);
  input.editor.value = input.initialValue;
  input.editor.style.display = 'block';
  input.editor.style.left = `${cellRect.x}px`;
  input.editor.style.top = `${cellRect.y}px`;
  input.editor.style.width = `${cellRect.width}px`;
  input.editor.style.height = `${cellRect.height}px`;
  input.editor.focus();
  if (input.selectText) input.editor.select();
  else input.editor.setSelectionRange(input.editor.value.length, input.editor.value.length);
}

function logicalRectToSpreadsheetEditorScreen(
  rect: { x: number; y: number; width: number; height: number },
  viewport: SpreadsheetViewport,
): { x: number; y: number; width: number; height: number } {
  return {
    x: viewport.screenRect.x + rect.x - viewport.originX,
    y: viewport.screenRect.y + rect.y - viewport.originY,
    width: rect.width,
    height: rect.height,
  };
}

function renderSpreadsheetInspector(
  inspector: HTMLPreElement,
  input: {
    fixturePath: string;
    fixtureSize: number;
    runtime: SpreadsheetEditorRuntime;
    scroll: { x: number; y: number };
    selection: SpreadsheetSelectionState;
    clipboardStatus?: unknown;
    agentPreviewStatus?: SpreadsheetAgentPreviewDevState | null;
    exportStatus?: SpreadsheetExportDevState;
  },
): void {
  const active = input.selection.active;
  const cell = getRenderCell(input.runtime.activeSheet(), active.row, active.column);
  inspector.textContent = JSON.stringify(
    {
      surface: 'spreadsheet',
      fixturePath: input.fixturePath,
      fixtureSize: input.fixtureSize,
      activeSheet: input.runtime.activeSheetName(),
      sheets: input.runtime.model().sheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        cells: sheet.cells.length,
        merges: sheet.mergedRanges.length,
      })),
      scroll: input.scroll,
      selection: input.selection,
      clipboard: input.clipboardStatus ?? null,
      agentPreview: summarizeSpreadsheetAgentPreviewState(input.agentPreviewStatus ?? null),
      export: input.exportStatus ?? null,
      history: summarizeSpreadsheetHistory(input.runtime),
      activeCell: {
        address: formatCellAddress(active.row, active.column),
        value: cell?.value ?? null,
        displayText: cell?.displayText ?? '',
        formula: cell?.formula ?? null,
        styleId: cell?.styleId ?? null,
      },
    },
    null,
    2,
  );
}

function renderSpreadsheetExportStatus(
  target: HTMLElement,
  state: SpreadsheetExportDevState,
): void {
  target.classList.remove('is-saving', 'is-ok', 'is-error');
  target.title = '';
  if (state.status === 'idle') {
    target.textContent = '';
    return;
  }
  target.classList.add(`is-${state.status}`);
  if (state.status === 'saving') {
    target.textContent = `Exporting ${state.fileName}...`;
    return;
  }
  if (state.status === 'ok') {
    target.textContent = `Exported ${formatByteLength(state.byteLength)}`;
    target.title = state.fileName;
    return;
  }
  target.textContent = `Export failed: ${state.message}`;
  target.title = state.message;
}

function formatByteLength(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function summarizeSpreadsheetAgentPreviewState(
  state: SpreadsheetAgentPreviewDevState | null,
): unknown {
  if (!state) return null;
  const action = {
    id: state.agentAction.id,
    label: state.agentAction.label,
    actor: state.agentAction.actor,
    agentRunId: state.agentAction.agentRunId,
    agentStepId: state.agentAction.agentStepId,
    agentActionId: state.agentAction.agentActionId,
    description: state.agentAction.description,
  };
  if (!state.ok) {
    return {
      action: state.action,
      ok: false,
      target: state.target,
      agentAction: action,
      context: state.context,
      error: state.error,
    };
  }
  if (state.action === 'preview') {
    return {
      action: 'preview',
      ok: true,
      target: state.target,
      agentAction: action,
      context: state.context,
      preview: {
        summary: state.preview.transaction.summary,
        diffSummary: state.preview.transaction.diffSummary,
      },
    };
  }
  return {
    action: 'apply',
    ok: true,
    target: state.target,
    agentAction: action,
    context: state.context,
    preview: {
      summary: state.preview.transaction.summary,
      diffSummary: state.preview.transaction.diffSummary,
    },
    apply: {
      summary: state.apply.transaction?.summary ?? null,
      affectedRanges: state.apply.transaction?.affectedRanges ?? [],
    },
  };
}

type SpreadsheetAgentPreviewDevState =
  | {
      action: 'preview';
      ok: true;
      target: SpreadsheetRange;
      agentAction: SpreadsheetAgentAction;
      context: { selection: SpreadsheetRange };
      preview: SpreadsheetAgentActionPreview;
    }
  | {
      action: 'preview';
      ok: false;
      target: SpreadsheetRange;
      agentAction: SpreadsheetAgentAction;
      context: { selection: SpreadsheetRange };
      error: string;
    }
  | {
      action: 'apply';
      ok: true;
      target: SpreadsheetRange;
      agentAction: SpreadsheetAgentAction;
      context: { selection: SpreadsheetRange };
      preview: SpreadsheetAgentActionPreview;
      apply: SpreadsheetAgentActionApplyResult;
    }
  | {
      action: 'apply';
      ok: false;
      target: SpreadsheetRange;
      agentAction: SpreadsheetAgentAction;
      context: { selection: SpreadsheetRange };
      preview: SpreadsheetAgentActionPreview;
      error: string;
    };

type SpreadsheetExportDevState =
  | {
      status: 'idle';
    }
  | {
      status: 'saving';
      fileName: string;
      startedAt: number;
    }
  | {
      status: 'ok';
      fileName: string;
      byteLength: number;
      finishedAt: number;
    }
  | {
      status: 'error';
      fileName: string;
      message: string;
      finishedAt: number;
    };

function canApplySpreadsheetAgentPreview(
  state: SpreadsheetAgentPreviewDevState | null,
): state is Extract<SpreadsheetAgentPreviewDevState, { action: 'preview'; ok: true }> {
  return state?.action === 'preview' && state.ok;
}

function createSpreadsheetAgentDemoAction(target: SpreadsheetRange): SpreadsheetAgentAction {
  return {
    id: 'dev-preview-agent-edit',
    label: 'Dev preview agent edit',
    actor: 'spreadsheet-dev-agent',
    description: `Writes a 2x2 demo block at ${formatCellAddress(target.startRow, target.startColumn)}`,
    agentRunId: `dev-preview:${Date.now()}`,
    agentStepId: 'write-demo-block',
    run: (runtime, context) => runSpreadsheetAgentDemoEdit(runtime, context.selection ?? target),
  };
}

async function runSpreadsheetAgentDemoEdit(
  runtime: SpreadsheetEditorRuntime,
  target: SpreadsheetRange,
): Promise<void> {
  const row = target.startRow;
  const column = target.startColumn;
  await runtime.setRangeValues(row, column, [
    ['Agent preview', 'pending'],
    ['Apply to commit', 'ready'],
  ]);
  await runtime.setRangeBold(row, column, 1, 2, true);
  await runtime.setRangeFillColor(row, column, 2, 2, { rgb: '#DBEAFE' });
}

function summarizeSpreadsheetHistory(runtime: SpreadsheetEditorRuntime): {
  canUndo: boolean;
  canRedo: boolean;
  undoDepth: number;
  redoDepth: number;
  containsSnapshot: boolean;
  undoTop: ReturnType<typeof summarizeExcelTransactionGroup>[];
  redoTop: ReturnType<typeof summarizeExcelTransactionGroup>[];
} {
  const history = runtime.workbook().history();
  const undoGroups = history.undoGroups();
  const redoGroups = history.redoGroups();
  return {
    canUndo: history.canUndo(),
    canRedo: history.canRedo(),
    undoDepth: undoGroups.length,
    redoDepth: redoGroups.length,
    containsSnapshot: [...undoGroups, ...redoGroups].some(
      (group) => summarizeExcelTransactionGroup(group).containsSnapshot,
    ),
    undoTop: undoGroups.slice(-5).reverse().map(summarizeExcelTransactionGroup),
    redoTop: redoGroups.slice(-5).reverse().map(summarizeExcelTransactionGroup),
  };
}

function selectedCellCount(selection: SpreadsheetSelectionState): number {
  const range = normalizeSelectionRange(selection.range);
  return (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1);
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

type NativeWordDocumentSnapshot = NativeWordEditor['getDocument'] extends () => infer T
  ? T | undefined
  : never;
type WordViewSnapshot = Awaited<ReturnType<typeof buildWordEditorViewFromPackage>>;

interface RenderedDomSnapshot {
  paragraphs: number;
  tables: number;
  images: number;
  anchoredImages: number;
  drawings: number;
  textBoxes: number;
  anchoredDrawings: number;
  numbered: number;
  pageBreaks: number;
  headers: number;
  footers: number;
  pages: number;
  overflowingPages: number;
  tableFragments: number;
  blockRowFragments: number;
  wrapTypes: string[];
  comments: number;
  commentBubbles: number;
  commentBubbleBodies: number;
  firstTableStyleId: string | null;
  pageLayouts: string[];
  sectionPageLayouts: string[];
  pageLayerAnchors: number;
  paragraphLayerAnchors: number;
  storyAnchors: number;
  bodyAnchors: number;
  visibleHeaderParts: string[];
  visibleFooterParts: string[];
  visibleHeaderAnchors: number;
  visibleFooterAnchors: number;
  anchorWrapTypes: string[];
  flowMode: 'natural' | 'grid';
  revisionInserts: number;
  revisionDeletes: number;
  styleProbe: StyleProbeSnapshot;
}

interface StyleProbeSnapshot {
  found: boolean;
  value: string;
  ok: boolean;
}

interface AcceptanceReport {
  status: 'pass' | 'warn';
  metrics: Array<{ label: string; value: string; tone?: 'good' | 'warn' }>;
  styleSignals: Array<{ label: string; value: string; tone?: 'good' | 'warn' }>;
  risks: string[];
}

function collectRenderedDomSnapshot(root: HTMLElement): RenderedDomSnapshot {
  return {
    paragraphs: root.querySelectorAll('coco-word-native-p, .ProseMirror p').length,
    tables: root.querySelectorAll('coco-word-native-table, table').length,
    images: root.querySelectorAll('coco-word-native-document img, .ProseMirror img').length,
    anchoredImages: root.querySelectorAll(
      'coco-word-native-document img[data-anchoring-mode="anchor"]',
    ).length,
    drawings: root.querySelectorAll('coco-word-native-document coco-word-native-drawing').length,
    textBoxes: root.querySelectorAll('coco-word-native-drawing[data-drawing-kind="textBox"]')
      .length,
    anchoredDrawings: root.querySelectorAll(
      'coco-word-native-document coco-word-native-drawing[data-anchoring-mode="anchor"]',
    ).length,
    numbered: root.querySelectorAll(
      'coco-word-native-p[data-numbering="true"], p[data-numbering="true"]',
    ).length,
    pageBreaks: root.querySelectorAll('.coco-word-page-break').length,
    headers: root.querySelectorAll('coco-word-native-story[data-story-kind="header"]').length,
    footers: root.querySelectorAll('coco-word-native-story[data-story-kind="footer"]').length,
    pages: root.querySelectorAll('coco-word-native-page').length,
    overflowingPages: countOverflowingNativePages(root),
    tableFragments: root.querySelectorAll('coco-word-native-table.is-table-fragment').length,
    blockRowFragments: root.querySelectorAll('tr[data-row-fragment="block"]').length,
    wrapTypes: Array.from(
      new Set(
        Array.from(
          root.querySelectorAll(
            'coco-word-native-document img[data-wrap], coco-word-native-drawing[data-wrap]',
          ),
        )
          .map((node) => node.getAttribute('data-wrap'))
          .filter((wrap): wrap is string => wrap !== null && wrap.length > 0),
      ),
    ).sort(),
    comments: root.querySelectorAll('.coco-word-comment[data-comment-id]').length,
    commentBubbles: root.querySelectorAll('.coco-word-native-comment-bubble[data-comment-id]')
      .length,
    commentBubbleBodies: Array.from(root.querySelectorAll('.coco-word-native-comment-body')).filter(
      (body) => (body.textContent ?? '').trim().length > 0,
    ).length,
    firstTableStyleId:
      root.querySelector('coco-word-native-table')?.getAttribute('data-style-id') ?? null,
    pageLayouts: collectDomPageLayouts(root),
    sectionPageLayouts: collectDomSectionPageLayouts(root),
    pageLayerAnchors: root.querySelectorAll(
      '.coco-word-native-page-overlay > .is-page-layer-anchor',
    ).length,
    paragraphLayerAnchors: root.querySelectorAll('.is-paragraph-relative-anchor').length,
    storyAnchors: root.querySelectorAll('coco-word-native-story [data-anchoring-mode="anchor"]')
      .length,
    bodyAnchors: countDomAnchorsBySource(root, undefined),
    visibleHeaderParts: collectVisibleStoryParts(root, 'header'),
    visibleFooterParts: collectVisibleStoryParts(root, 'footer'),
    visibleHeaderAnchors: countDomAnchorsBySource(root, 'header'),
    visibleFooterAnchors: countDomAnchorsBySource(root, 'footer'),
    anchorWrapTypes: collectDomAnchorWrapTypes(root),
    flowMode: root
      .querySelector('coco-word-native-document')
      ?.classList.contains('uses-natural-paragraph-flow')
      ? 'natural'
      : 'grid',
    revisionInserts: root.querySelectorAll('.coco-word-revision-insert').length,
    revisionDeletes: root.querySelectorAll('.coco-word-revision-delete').length,
    styleProbe: collectStyleProbe(root),
  };
}

function collectStyleProbe(root: HTMLElement): StyleProbeSnapshot {
  const paragraph = Array.from(root.querySelectorAll<HTMLElement>('coco-word-native-p')).find(
    (node) => (node.textContent ?? '').includes('COCO_STYLE_PROBE'),
  );
  if (!paragraph) return { found: false, value: 'missing', ok: false };
  const run = paragraph.querySelector<HTMLElement>('span') ?? paragraph;
  const paragraphStyle = getComputedStyle(paragraph);
  const runStyle = getComputedStyle(run);
  const checks = [
    approxPx(paragraphStyle.marginLeft, 48),
    approxPx(paragraphStyle.textIndent, 24),
    approxPx(paragraphStyle.lineHeight, 28),
    approxPx(runStyle.fontSize, 18),
    runStyle.fontFamily.toLowerCase().includes('arial'),
    colorMatches(runStyle.color, '34, 68, 136'),
  ];
  return {
    found: true,
    ok: checks.every(Boolean),
    value: [
      `ml ${paragraphStyle.marginLeft}`,
      `indent ${paragraphStyle.textIndent}`,
      `line ${paragraphStyle.lineHeight}`,
      `font ${runStyle.fontSize} ${runStyle.fontFamily}`,
      `color ${runStyle.color}`,
    ].join(' | '),
  };
}

function approxPx(value: string, expected: number): boolean {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && Math.abs(parsed - expected) <= 1.5;
}

function colorMatches(value: string, rgb: string): boolean {
  return value.replace(/\s+/g, ' ').includes(rgb);
}

function countOverflowingNativePages(root: HTMLElement): number {
  return Array.from(root.querySelectorAll('coco-word-native-page-body')).filter((body) => {
    const rect = body.getBoundingClientRect();
    return body.scrollHeight > rect.height + 24;
  }).length;
}

function buildAcceptanceReport(
  wordView: WordViewSnapshot,
  nativeDocument: NativeWordDocumentSnapshot,
  dom: RenderedDomSnapshot,
): AcceptanceReport {
  const coverage = nativeDocument ? summarizeNativeCoverage(wordView, nativeDocument) : null;
  const editorCoverage = summarizeEditorCoverage(wordView);
  const nativeTables = nativeDocument?.blocks.filter((block) => block.kind === 'table').length ?? 0;
  const tablePagination = nativeDocument ? summarizeNativeTablePagination(nativeDocument) : null;
  const numbered = wordView.paragraphs.filter((paragraph) => paragraph.format?.numbering).length;
  const comments = countCommentSpans(wordView);
  const anchoredImages = countAnchoredImages(wordView);
  const bodyAnchoredImages = countAnchoredImagesInParagraphs(wordView.paragraphs);
  const headerAnchoredImages = countAnchoredImagesInStories(wordView.headers);
  const footerAnchoredImages = countAnchoredImagesInStories(wordView.footers);
  const visibleHeaderAnchoredImages = countAnchoredImagesInVisibleStories(
    wordView.headers,
    dom.visibleHeaderParts,
  );
  const visibleFooterAnchoredImages = countAnchoredImagesInVisibleStories(
    wordView.footers,
    dom.visibleFooterParts,
  );
  const drawings = countDrawings(wordView);
  const textBoxes = countTextBoxes(wordView);
  const anchoredDrawings = countAnchoredDrawings(wordView);
  const revisions = countRevisionSpans(wordView);
  const pageBreakBefore = countPageBreakBefore(wordView);
  const sectionBreakParagraphs = countSectionBreakParagraphs(wordView);
  const wrapTypes = wrapTypesFromContent(wordView);
  const headerRefs = countHeaderRefs(wordView);
  const footerRefs = countFooterRefs(wordView);
  const referencedHeaderParts = referencedHeaderFooterParts(wordView, 'header');
  const referencedFooterParts = referencedHeaderFooterParts(wordView, 'footer');
  const parsedHeaderParts = new Set(wordView.headers.map((story) => story.partName));
  const parsedFooterParts = new Set(wordView.footers.map((story) => story.partName));
  const tableStyleHits = wordView.tables.filter((table) => table.tableStyleId).length;
  const tableBorderHits = wordView.tables.filter((table) => table.borders).length;
  const tableCellMarginHits = wordView.tables.filter((table) =>
    table.rows.some((row) => row.cells.some((cell) => cell.margin)),
  ).length;
  const tableCellShadingHits = wordView.tables.filter((table) =>
    table.rows.some((row) => row.cells.some((cell) => cell.shading)),
  ).length;
  const rowHeightHits = wordView.tables.filter((table) =>
    table.rows.some((row) => row.height),
  ).length;
  const hasNativeModel = nativeDocument !== undefined;
  const sectionLayoutMismatches = findSectionLayoutMismatches(wordView, dom);
  const expectsStyleProbe = allWordParagraphs(wordView).some((paragraph) =>
    paragraph.text.includes('COCO_STYLE_PROBE'),
  );
  const risks: string[] = [];

  if (!wordView.pageLayout) risks.push('No page layout was parsed.');
  if (coverage && coverage.missing.paragraphs !== 0)
    risks.push(`${coverage.missing.paragraphs} paragraphs missing in native model.`);
  if (coverage && coverage.missing.images !== 0)
    risks.push(`${coverage.missing.images} images missing in native model.`);
  if (coverage && coverage.missing.drawings !== 0)
    risks.push(`${coverage.missing.drawings} drawings missing in native model.`);
  if (coverage && coverage.missing.pageBreaks !== 0)
    risks.push(`${coverage.missing.pageBreaks} page breaks missing in native model.`);
  if (nativeDocument && nativeTables !== wordView.tables.length) {
    risks.push(`${wordView.tables.length - nativeTables} tables missing in native model.`);
  }
  if (!hasNativeModel && dom.paragraphs < editorCoverage.allParagraphs)
    risks.push('Rendered DOM paragraph count is lower than Word view.');
  if (!hasNativeModel && dom.images < editorCoverage.allImages)
    risks.push('Rendered DOM image count is lower than Word view.');
  if (!hasNativeModel && dom.anchoredImages < anchoredImages)
    risks.push('Anchored images parsed but not marked in DOM.');
  if (!hasNativeModel && dom.drawings < drawings)
    risks.push('Rendered DOM drawing count is lower than Word view.');
  if (!hasNativeModel && dom.textBoxes < textBoxes)
    risks.push('Text boxes parsed but not rendered in DOM.');
  if (!hasNativeModel && dom.anchoredDrawings < anchoredDrawings)
    risks.push('Anchored drawings parsed but not marked in DOM.');
  if (!hasNativeModel && dom.pageBreaks < editorCoverage.pageBreaks)
    risks.push('Rendered DOM page break count is lower than Word view.');
  if (!hasNativeModel && dom.comments < comments)
    risks.push('Comment spans parsed but not marked in DOM.');
  if (dom.commentBubbles < comments)
    risks.push('Comment spans parsed but not all comment bubbles were rendered.');
  if (dom.commentBubbles > 0 && dom.commentBubbleBodies < dom.commentBubbles)
    risks.push('Comment bubbles rendered without visible body text.');
  if (dom.revisionInserts + dom.revisionDeletes < revisions)
    risks.push('Revision spans parsed but not marked in DOM.');
  if (dom.overflowingPages > 0)
    risks.push(`${dom.overflowingPages} rendered pages overflow their page body.`);
  for (const partName of referencedHeaderParts) {
    if (!parsedHeaderParts.has(partName))
      risks.push(`Header part ${partName} is referenced but was not parsed.`);
  }
  for (const partName of referencedFooterParts) {
    if (!parsedFooterParts.has(partName))
      risks.push(`Footer part ${partName} is referenced but was not parsed.`);
  }
  for (const story of wordView.headers) {
    if (referencedHeaderParts.has(story.partName) && story.blocks.length === 0) {
      risks.push(`Header part ${story.partName} was parsed but contains no rendered blocks.`);
    }
  }
  for (const story of wordView.footers) {
    if (referencedFooterParts.has(story.partName) && story.blocks.length === 0) {
      risks.push(`Footer part ${story.partName} was parsed but contains no rendered blocks.`);
    }
  }
  if (headerRefs > 0 && dom.headers === 0)
    risks.push('Header references parsed but no rendered headers were found.');
  if (footerRefs > 0 && dom.footers === 0)
    risks.push('Footer references parsed but no rendered footers were found.');
  for (const mismatch of sectionLayoutMismatches) risks.push(mismatch);
  if (wordView.tables.length > 0 && tableBorderHits === 0) risks.push('No table borders resolved.');
  if (numbered > 0 && dom.numbered === 0)
    risks.push('Numbered paragraphs parsed but not marked in DOM.');
  if (expectsStyleProbe && !dom.styleProbe.found)
    risks.push('Style probe paragraph was not rendered.');
  if (expectsStyleProbe && dom.styleProbe.found && !dom.styleProbe.ok) {
    risks.push('Style probe computed styles do not match expected Word formatting.');
  }

  return {
    status: risks.length === 0 ? 'pass' : 'warn',
    metrics: [
      {
        label: 'Paragraphs',
        value: `${editorCoverage.allParagraphs} / ${nativeDocument ? coverage?.native.paragraphs : dom.paragraphs}`,
        tone:
          editorCoverage.allParagraphs !==
          (nativeDocument ? coverage?.native.paragraphs : dom.paragraphs)
            ? 'warn'
            : undefined,
      },
      {
        label: 'Tables',
        value: `${wordView.tables.length} / ${nativeDocument ? nativeTables : dom.tables}`,
      },
      {
        label: 'Table Pagination',
        value: tablePagination
          ? `${tablePagination.textFragmentableRows} text / ${tablePagination.blockFragmentableRows} block / ${tablePagination.unsplittableRows} locked`
          : 'n/a',
        tone:
          tablePagination &&
          tablePagination.textFragmentableRows + tablePagination.blockFragmentableRows > 0
            ? 'warn'
            : undefined,
      },
      {
        label: 'Images',
        value: `${editorCoverage.allImages} / ${nativeDocument ? coverage?.native.images : dom.images}`,
      },
      {
        label: 'Anchored Images',
        value: `${anchoredImages} model / ${dom.anchoredImages} rendered`,
        tone: !hasNativeModel && anchoredImages > dom.anchoredImages ? 'warn' : undefined,
      },
      {
        label: 'Drawings',
        value: `${drawings} / ${dom.drawings}`,
        tone: !hasNativeModel && drawings > dom.drawings ? 'warn' : undefined,
      },
      {
        label: 'TextBoxes',
        value: `${textBoxes} / ${dom.textBoxes}`,
        tone: !hasNativeModel && textBoxes > dom.textBoxes ? 'warn' : undefined,
      },
      { label: 'Numbered', value: `${numbered} / ${dom.numbered}` },
      {
        label: 'Page Breaks',
        value: `${editorCoverage.pageBreaks} / ${nativeDocument ? coverage?.native.pageBreaks : dom.pageBreaks}`,
      },
      { label: 'Page Break Before', value: String(pageBreakBefore) },
      { label: 'Section Break Paras', value: String(sectionBreakParagraphs) },
      {
        label: 'Pages',
        value: String(dom.pages || 1),
        tone: dom.overflowingPages > 0 ? 'warn' : undefined,
      },
      {
        label: 'Table Fragments',
        value: String(dom.tableFragments),
        tone: dom.tableFragments > 0 ? 'good' : undefined,
      },
      { label: 'Sections', value: String(wordView.sections.length) },
      {
        label: 'Headers',
        value: `${headerRefs} / ${wordView.headers.length} / ${dom.headers}`,
        tone:
          headerRefs > wordView.headers.length || (headerRefs > 0 && dom.headers === 0)
            ? 'warn'
            : undefined,
      },
      {
        label: 'Footers',
        value: `${footerRefs} / ${wordView.footers.length} / ${dom.footers}`,
        tone:
          footerRefs > wordView.footers.length || (footerRefs > 0 && dom.footers === 0)
            ? 'warn'
            : undefined,
      },
      {
        label: 'Comments',
        value: `${comments} / ${dom.comments} / ${dom.commentBubbles}`,
        tone:
          (!hasNativeModel && comments > dom.comments) ||
          dom.commentBubbles < comments ||
          (dom.commentBubbles > 0 && dom.commentBubbleBodies < dom.commentBubbles)
            ? 'warn'
            : undefined,
      },
      {
        label: 'Revisions',
        value: `${revisions} / ${dom.revisionInserts + dom.revisionDeletes}`,
        tone:
          revisions > dom.revisionInserts + dom.revisionDeletes
            ? 'warn'
            : revisions > 0
              ? 'good'
              : undefined,
      },
    ],
    styleSignals: [
      {
        label: 'Page Layout',
        value: wordView.pageLayout
          ? `${wordView.pageLayout.width} x ${wordView.pageLayout.height}`
          : 'missing',
        tone: wordView.pageLayout ? 'good' : 'warn',
      },
      { label: 'Flow Mode', value: dom.flowMode },
      {
        label: 'Table Styles',
        value: `${tableStyleHits}/${wordView.tables.length}`,
        tone: tableStyleHits > 0 || wordView.tables.length === 0 ? 'good' : 'warn',
      },
      {
        label: 'Table Borders',
        value: `${tableBorderHits}/${wordView.tables.length}`,
        tone: tableBorderHits > 0 || wordView.tables.length === 0 ? 'good' : 'warn',
      },
      {
        label: 'Cell Margins',
        value: `${tableCellMarginHits}/${wordView.tables.length}`,
        tone: tableCellMarginHits > 0 || wordView.tables.length === 0 ? 'good' : 'warn',
      },
      { label: 'Cell Shading', value: `${tableCellShadingHits}/${wordView.tables.length}` },
      { label: 'Row Heights', value: `${rowHeightHits}/${wordView.tables.length}` },
      {
        label: 'Block Row Fragments',
        value: String(dom.blockRowFragments),
        tone: dom.blockRowFragments > 0 ? 'good' : undefined,
      },
      {
        label: 'Image Wraps',
        value: wrapTypes.length > 0 ? wrapTypes.join(', ') : 'none',
        tone:
          wrapTypes.length > 0 && dom.wrapTypes.length === wrapTypes.length ? 'good' : undefined,
      },
      {
        label: 'DOM Anchor Wraps',
        value: dom.anchorWrapTypes.length > 0 ? dom.anchorWrapTypes.join(', ') : 'none',
        tone: dom.anchorWrapTypes.length > 0 ? 'good' : undefined,
      },
      {
        label: 'Body Anchors',
        value: `${bodyAnchoredImages} / ${dom.bodyAnchors}`,
        tone: bodyAnchoredImages === dom.bodyAnchors ? 'good' : undefined,
      },
      {
        label: 'Visible Header Anchors',
        value: `${visibleHeaderAnchoredImages} / ${dom.visibleHeaderAnchors}`,
        tone: visibleHeaderAnchoredImages === dom.visibleHeaderAnchors ? 'good' : undefined,
      },
      {
        label: 'Visible Footer Anchors',
        value: `${visibleFooterAnchoredImages} / ${dom.visibleFooterAnchors}`,
        tone: visibleFooterAnchoredImages === dom.visibleFooterAnchors ? 'good' : undefined,
      },
      {
        label: 'Unused Header Anchors',
        value: String(Math.max(0, headerAnchoredImages - visibleHeaderAnchoredImages)),
      },
      {
        label: 'Unused Footer Anchors',
        value: String(Math.max(0, footerAnchoredImages - visibleFooterAnchoredImages)),
      },
      {
        label: 'Page Layer Anchors',
        value: String(dom.pageLayerAnchors),
        tone: dom.pageLayerAnchors > 0 ? 'good' : undefined,
      },
      {
        label: 'Paragraph Anchors',
        value: String(dom.paragraphLayerAnchors),
        tone: dom.paragraphLayerAnchors > 0 ? 'good' : undefined,
      },
      {
        label: 'Story Anchors',
        value: String(dom.storyAnchors),
        tone: dom.storyAnchors > 0 ? 'good' : undefined,
      },
      {
        label: 'DOM Table Style',
        value: dom.firstTableStyleId ?? 'none',
        tone: dom.firstTableStyleId ? 'good' : undefined,
      },
      {
        label: 'Section Layouts',
        value: summarizeSectionLayouts(wordView),
        tone: wordView.sections.some((section) => section.pageLayout) ? 'good' : 'warn',
      },
      {
        label: 'DOM Page Layouts',
        value: dom.pageLayouts.length > 0 ? dom.pageLayouts.join(', ') : 'none',
        tone: dom.pageLayouts.length > 0 ? 'good' : undefined,
      },
      {
        label: 'Style Probe',
        value: expectsStyleProbe ? dom.styleProbe.value : 'not present',
        tone: expectsStyleProbe ? (dom.styleProbe.ok ? 'good' : 'warn') : undefined,
      },
      {
        label: 'DOM Section Layouts',
        value: dom.sectionPageLayouts.length > 0 ? dom.sectionPageLayouts.join('; ') : 'none',
        tone:
          sectionLayoutMismatches.length > 0
            ? 'warn'
            : dom.sectionPageLayouts.length > 0
              ? 'good'
              : undefined,
      },
    ],
    risks,
  };
}

function collectDomAnchorWrapTypes(root: HTMLElement): string[] {
  return Array.from(
    new Set(
      Array.from(root.querySelectorAll('[data-anchoring-mode="anchor"][data-wrap]'))
        .map((node) => node.getAttribute('data-wrap'))
        .filter((wrap): wrap is string => wrap !== null && wrap.length > 0),
    ),
  ).sort();
}

function collectVisibleStoryParts(root: HTMLElement, kind: 'header' | 'footer'): string[] {
  return Array.from(
    new Set(
      Array.from(root.querySelectorAll(`coco-word-native-story[data-story-kind="${kind}"]`))
        .map((story) => story.getAttribute('data-part-name'))
        .filter((partName): partName is string => partName !== null && partName.length > 0),
    ),
  ).sort();
}

function countDomAnchorsBySource(root: HTMLElement, kind: 'header' | 'footer' | undefined): number {
  const anchors = Array.from(
    root.querySelectorAll<HTMLElement>(
      'coco-word-native-document img[data-anchoring-mode="anchor"]',
    ),
  );
  return anchors.filter((anchor) => {
    const sourceKind =
      anchor.dataset.sourceStoryKind ??
      anchor.closest<HTMLElement>('coco-word-native-story')?.dataset.storyKind;
    return kind === undefined ? sourceKind === undefined : sourceKind === kind;
  }).length;
}

function collectDomPageLayouts(root: HTMLElement): string[] {
  return Array.from(
    new Set(
      Array.from(root.querySelectorAll('coco-word-native-page')).map((page) => {
        const rect = page.getBoundingClientRect();
        return `${Math.round(rect.width)}x${Math.round(rect.height)}`;
      }),
    ),
  );
}

function collectDomSectionPageLayouts(root: HTMLElement): string[] {
  const layoutsBySection = new Map<string, Set<string>>();
  for (const page of Array.from(root.querySelectorAll('coco-word-native-page'))) {
    const sectionIndex = page.getAttribute('data-section-index') ?? '?';
    const rect = page.getBoundingClientRect();
    const layouts = layoutsBySection.get(sectionIndex) ?? new Set<string>();
    layouts.add(`${Math.round(rect.width)}x${Math.round(rect.height)}`);
    layoutsBySection.set(sectionIndex, layouts);
  }
  return Array.from(layoutsBySection.entries()).map(([sectionIndex, layouts]) => {
    return `${sectionIndex}:${Array.from(layouts).join('/')}`;
  });
}

function findSectionLayoutMismatches(
  wordView: WordViewSnapshot,
  dom: RenderedDomSnapshot,
): string[] {
  if (wordView.sections.length === 0 || dom.sectionPageLayouts.length === 0) return [];
  const domLayouts = parseDomSectionPageLayouts(dom.sectionPageLayouts);
  return wordView.sections.flatMap((section) => {
    if (!section.pageLayout) return [];
    const expected = `${twipsToRoundedPx(section.pageLayout.width)}x${twipsToRoundedPx(section.pageLayout.height)}`;
    const actual = domLayouts.get(String(section.sectionIndex));
    if (!actual) return [`Section ${section.sectionIndex} has no rendered pages.`];
    if (!actual.has(expected)) {
      return [
        `Section ${section.sectionIndex} page layout ${expected} was parsed but DOM rendered ${Array.from(actual).join('/')}.`,
      ];
    }
    return [];
  });
}

function parseDomSectionPageLayouts(layouts: string[]): Map<string, Set<string>> {
  const parsed = new Map<string, Set<string>>();
  for (const item of layouts) {
    const separatorIndex = item.indexOf(':');
    if (separatorIndex < 0) continue;
    const sectionIndex = item.slice(0, separatorIndex);
    const pageLayouts = item
      .slice(separatorIndex + 1)
      .split('/')
      .filter((layout) => layout.length > 0);
    parsed.set(sectionIndex, new Set(pageLayouts));
  }
  return parsed;
}

function twipsToRoundedPx(value: number): number {
  return Math.round(value / 15);
}

function summarizeSectionLayouts(wordView: WordViewSnapshot): string {
  const layouts = wordView.sections
    .map((section) => {
      const layout = section.pageLayout;
      if (!layout) return `${section.sectionIndex}:missing`;
      return `${section.sectionIndex}:${layout.width}x${layout.height}`;
    })
    .join(', ');
  return layouts.length > 0 ? layouts : 'none';
}

function renderAcceptancePanel(panel: HTMLElement | null, report: AcceptanceReport): void {
  if (!panel) return;
  const statusText = report.status === 'pass' ? 'PASS' : 'CHECK';
  panel.innerHTML = `
    <div class="coco-word-acceptance-status is-${report.status}">${statusText}</div>
    <div class="coco-word-acceptance-grid">
      ${report.metrics.map(renderAcceptanceMetric).join('')}
    </div>
    <div class="coco-word-acceptance-section">Style Signals</div>
    <div class="coco-word-acceptance-grid">
      ${report.styleSignals.map(renderAcceptanceMetric).join('')}
    </div>
    <div class="coco-word-acceptance-section">Risks</div>
    ${
      report.risks.length > 0
        ? `<ul>${report.risks.map((risk) => `<li>${escapeHtml(risk)}</li>`).join('')}</ul>`
        : '<div class="coco-word-acceptance-empty">No obvious content loss detected.</div>'
    }
  `;
}

function renderAcceptanceMetric(metric: {
  label: string;
  value: string;
  tone?: 'good' | 'warn';
}): string {
  return `
    <div class="coco-word-acceptance-metric${metric.tone ? ` is-${metric.tone}` : ''}">
      <span>${escapeHtml(metric.label)}</span>
      <strong>${escapeHtml(metric.value)}</strong>
    </div>
  `;
}

function countHeaderRefs(wordView: WordViewSnapshot): number {
  return wordView.sections.reduce((count, section) => count + section.headerRefs.length, 0);
}

function countFooterRefs(wordView: WordViewSnapshot): number {
  return wordView.sections.reduce((count, section) => count + section.footerRefs.length, 0);
}

function countCommentSpans(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) => count + paragraph.spans.filter((span) => span.kind === 'comment').length,
    0,
  );
}

function countRevisionSpans(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) => count + paragraph.spans.filter((span) => span.kind === 'revision').length,
    0,
  );
}

function countAnchoredImages(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) =>
      count +
      paragraph.segments.filter(
        (segment) => segment.kind === 'image' && segment.anchoringMode === 'anchor',
      ).length,
    0,
  );
}

function countAnchoredImagesInParagraphs(paragraphs: WordViewSnapshot['paragraphs']): number {
  return paragraphs.reduce(
    (count, paragraph) =>
      count +
      paragraph.segments.filter(
        (segment) => segment.kind === 'image' && segment.anchoringMode === 'anchor',
      ).length,
    0,
  );
}

function countAnchoredImagesInStories(
  stories: WordViewSnapshot['headers'] | WordViewSnapshot['footers'],
): number {
  return stories.reduce(
    (count, story) => count + countAnchoredImagesInParagraphs(story.paragraphs),
    0,
  );
}

function countAnchoredImagesInVisibleStories(
  stories: WordViewSnapshot['headers'] | WordViewSnapshot['footers'],
  visibleParts: string[],
): number {
  const visible = new Set(visibleParts);
  return countAnchoredImagesInStories(stories.filter((story) => visible.has(story.partName)));
}

function countDrawings(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) =>
      count + paragraph.segments.filter((segment) => segment.kind === 'drawing').length,
    0,
  );
}

function countTextBoxes(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) =>
      count +
      paragraph.segments.filter(
        (segment) => segment.kind === 'drawing' && segment.drawingKind === 'textBox',
      ).length,
    0,
  );
}

function countPageBreakBefore(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).filter(
    (paragraph) => paragraph.format?.pageBreakBefore === true,
  ).length;
}

function countSectionBreakParagraphs(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).filter(
    (paragraph) => paragraph.format?.hasSectionBreak === true,
  ).length;
}

function countAnchoredDrawings(wordView: WordViewSnapshot): number {
  return allWordParagraphs(wordView).reduce(
    (count, paragraph) =>
      count +
      paragraph.segments.filter(
        (segment) => segment.kind === 'drawing' && segment.anchoringMode === 'anchor',
      ).length,
    0,
  );
}

function wrapTypesFromContent(wordView: WordViewSnapshot): string[] {
  return Array.from(
    new Set(
      allWordParagraphs(wordView).flatMap((paragraph) =>
        paragraph.segments
          .filter(
            (segment) => (segment.kind === 'image' || segment.kind === 'drawing') && segment.wrap,
          )
          .map((segment) =>
            segment.kind === 'image' || segment.kind === 'drawing' ? segment.wrap! : '',
          ),
      ),
    ),
  ).sort();
}

function allWordParagraphs(wordView: WordViewSnapshot): WordViewSnapshot['paragraphs'] {
  const hiddenTableParagraphIds = new Set(
    [
      ...wordView.tables,
      ...wordView.headers.flatMap((story) => story.tables),
      ...wordView.footers.flatMap((story) => story.tables),
    ].flatMap((table) =>
      table.rows.flatMap((row) =>
        row.cells.flatMap((cell) =>
          cell.vMerge === 'continue' ? cell.paragraphs.map((paragraph) => paragraph.paraId) : [],
        ),
      ),
    ),
  );
  return [
    ...wordView.paragraphs,
    ...wordView.headers.flatMap((story) => story.paragraphs),
    ...wordView.footers.flatMap((story) => story.paragraphs),
  ].filter((paragraph) => !hiddenTableParagraphIds.has(paragraph.paraId));
}

function referencedHeaderFooterParts(
  wordView: WordViewSnapshot,
  kind: 'header' | 'footer',
): Set<string> {
  const parts = new Set<string>();
  for (const section of wordView.sections) {
    const refs = kind === 'header' ? section.headerRefs : section.footerRefs;
    for (const ref of refs) parts.add(ref.partName);
  }
  return parts;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function summarizeNativeDocument(
  doc: NativeWordEditor['getDocument'] extends () => infer T ? T | undefined : unknown,
): unknown {
  if (!doc) return null;
  return {
    pageLayout: doc.pageLayout,
    headers: doc.headers.map((story) => ({
      partName: story.partName,
      blocks: story.blocks.length,
    })),
    footers: doc.footers.map((story) => ({
      partName: story.partName,
      blocks: story.blocks.length,
    })),
    blocks: doc.blocks.length,
    paragraphs: doc.blocks.reduce((count, block) => {
      if (block.kind === 'paragraph') return count + 1;
      return (
        count +
        block.rows.reduce(
          (rowCount, row) =>
            rowCount + row.cells.reduce((cellCount, cell) => cellCount + cell.paragraphs.length, 0),
          0,
        )
      );
    }, 0),
    images: doc.blocks.reduce((count, block) => {
      const countParagraphImages = (paragraph: { content: Array<{ kind: string }> }) =>
        paragraph.content.filter((child) => child.kind === 'image').length;
      if (block.kind === 'paragraph') return count + countParagraphImages(block);
      return (
        count +
        block.rows.reduce(
          (rowCount, row) =>
            rowCount +
            row.cells.reduce(
              (cellCount, cell) =>
                cellCount +
                cell.paragraphs.reduce(
                  (pCount, paragraph) => pCount + countParagraphImages(paragraph),
                  0,
                ),
              0,
            ),
          0,
        )
      );
    }, 0),
    preview: doc.blocks.slice(0, 8).map((block) =>
      block.kind === 'paragraph'
        ? {
            kind: 'paragraph',
            paraId: block.paraId,
            text: block.content
              .map((child) => (child.kind === 'text' ? child.text : '[image]'))
              .join('')
              .slice(0, 160),
          }
        : {
            kind: 'table',
            tableId: block.tableId,
            rows: block.rows.length,
          },
    ),
  };
}

function applyEditorPageLayout(
  page: HTMLElement,
  layout:
    | {
        width: number;
        height: number;
        margin: { top: number; right: number; bottom: number; left: number };
      }
    | undefined,
): void {
  if (!layout) return;
  page.style.width = twipsToPx(layout.width);
  page.style.minHeight = twipsToPx(layout.height);
  page.style.paddingTop = twipsToPx(layout.margin.top);
  page.style.paddingRight = twipsToPx(layout.margin.right);
  page.style.paddingBottom = twipsToPx(layout.margin.bottom);
  page.style.paddingLeft = twipsToPx(layout.margin.left);
}

function twipsToPx(twips: number): string {
  return `${twips / 15}px`;
}

function selectedFixtureIndex(fixtures: FixtureFile[]): number {
  const raw = new URLSearchParams(location.search).get('fixture');
  const index = raw === null ? 0 : Number(raw);
  return Number.isInteger(index) && index >= 0 && index < fixtures.length ? index : 0;
}

function selectedSpreadsheetFixtureIndex(fixtures: FixtureFile[]): number {
  const raw = new URLSearchParams(location.search).get('xlsxFixture');
  const index = raw === null ? 0 : Number(raw);
  return Number.isInteger(index) && index >= 0 && index < fixtures.length ? index : 0;
}

function selectedPptFixtureIndex(fixtures: FixtureFile[]): number {
  const raw = new URLSearchParams(location.search).get('pptFixture');
  const index = raw === null ? 0 : Number(raw);
  return Number.isInteger(index) && index >= 0 && index < fixtures.length ? index : 0;
}

function selectedSurface(): 'word' | 'spreadsheet' | 'spreadsheet-public' | 'ppt' {
  const raw = new URLSearchParams(location.search).get('surface');
  if (raw === 'spreadsheet-public') return 'spreadsheet-public';
  if (raw === 'ppt') return 'ppt';
  return raw === 'spreadsheet' ? 'spreadsheet' : 'word';
}

function selectedEngine(): 'prosemirror' | 'native' {
  const raw = new URLSearchParams(location.search).get('engine');
  return raw === 'native' ? 'native' : 'prosemirror';
}

function selectedReadOnly(): boolean {
  return new URLSearchParams(location.search).get('readOnly') === '1';
}

function selectedPptSlideListVisible(): boolean {
  return new URLSearchParams(location.search).get('showSlideList') !== '0';
}

function renderFixtureSelect(fixtures: FixtureFile[], selectedIndex: number): void {
  const select = document.querySelector<HTMLSelectElement>('#fixtureSelect');
  if (!select) return;
  select.innerHTML = fixtures
    .map(
      (fixture) =>
        `<option value="${fixture.index}"${fixture.index === selectedIndex ? ' selected' : ''}>${fixture.index + 1}. ${fixture.name}</option>`,
    )
    .join('');
  select.addEventListener('change', () => {
    const next = select.value;
    const params = new URLSearchParams(location.search);
    params.set('fixture', next);
    location.search = `?${params.toString()}`;
  });
}

function renderSpreadsheetFixtureSelect(fixtures: FixtureFile[], selectedIndex: number): void {
  const select = document.querySelector<HTMLSelectElement>('#spreadsheetFixtureSelect');
  if (!select) return;
  select.innerHTML = fixtures
    .map(
      (fixture) =>
        `<option value="${fixture.index}"${fixture.index === selectedIndex ? ' selected' : ''}>${fixture.index + 1}. ${fixture.name}</option>`,
    )
    .join('');
  select.addEventListener('change', () => {
    const params = new URLSearchParams(location.search);
    params.set('surface', 'spreadsheet');
    params.set('xlsxFixture', select.value);
    location.search = `?${params.toString()}`;
  });
}

function wireSurfaceSelect(): void {
  const select = document.querySelector<HTMLSelectElement>('#surfaceSelect');
  if (!select) return;
  select.addEventListener('change', () => {
    const params = new URLSearchParams(location.search);
    params.set('surface', select.value);
    location.search = `?${params.toString()}`;
  });
}

function renderEngineSelect(engine: 'prosemirror' | 'native'): void {
  const select = document.querySelector<HTMLSelectElement>('#engineSelect');
  if (!select) return;
  select.value = engine;
  select.addEventListener('change', () => {
    const params = new URLSearchParams(location.search);
    params.set('engine', select.value);
    location.search = `?${params.toString()}`;
  });
}

function wireToolbar(editorView: ProseMirrorEditorView): void {
  document.querySelector<HTMLButtonElement>('#boldButton')?.addEventListener('click', () => {
    toggleBold(editorView.state, editorView.dispatch, editorView);
  });

  document
    .querySelector<HTMLSelectElement>('#fontSizeSelect')
    ?.addEventListener('change', (event) => {
      const size = Number((event.currentTarget as HTMLSelectElement).value);
      setFontSize(size)(editorView.state, editorView.dispatch, editorView);
    });

  document
    .querySelector<HTMLInputElement>('#fontColorInput')
    ?.addEventListener('change', (event) => {
      const color = (event.currentTarget as HTMLInputElement).value;
      setColor(color)(editorView.state, editorView.dispatch, editorView);
    });

  document.querySelectorAll<HTMLButtonElement>('[data-align]').forEach((button) => {
    button.addEventListener('click', () => {
      const alignment = button.dataset.align;
      if (
        alignment === 'left' ||
        alignment === 'center' ||
        alignment === 'right' ||
        alignment === 'both'
      ) {
        setParagraphAlign(alignment)(editorView.state, editorView.dispatch);
      }
    });
  });
}

function wireNativeToolbar(editor: NativeWordEditor): void {
  const undoButton = document.querySelector<HTMLButtonElement>('#wordUndoButton');
  const redoButton = document.querySelector<HTMLButtonElement>('#wordRedoButton');
  const trackRevisionsButton = document.querySelector<HTMLButtonElement>('#trackRevisionsButton');
  const acceptRevisionButton = document.querySelector<HTMLButtonElement>('#acceptRevisionButton');
  const rejectRevisionButton = document.querySelector<HTMLButtonElement>('#rejectRevisionButton');
  const acceptAllRevisionsButton = document.querySelector<HTMLButtonElement>(
    '#acceptAllRevisionsButton',
  );
  const rejectAllRevisionsButton = document.querySelector<HTMLButtonElement>(
    '#rejectAllRevisionsButton',
  );
  const searchBar = document.querySelector<HTMLElement>('#wordSearchBar');
  const searchInput = document.querySelector<HTMLInputElement>('#wordSearchInput');
  const searchCount = document.querySelector<HTMLElement>('#wordSearchCount');
  const searchPreviousButton = document.querySelector<HTMLButtonElement>(
    '#wordSearchPreviousButton',
  );
  const searchNextButton = document.querySelector<HTMLButtonElement>('#wordSearchNextButton');
  const searchMatchCase = document.querySelector<HTMLInputElement>('#wordSearchMatchCase');
  const searchWholeWord = document.querySelector<HTMLInputElement>('#wordSearchWholeWord');
  const replaceInput = document.querySelector<HTMLInputElement>('#wordReplaceInput');
  const replaceCurrentButton = document.querySelector<HTMLButtonElement>(
    '#wordReplaceCurrentButton',
  );
  const replaceAllButton = document.querySelector<HTMLButtonElement>('#wordReplaceAllButton');
  const searchCloseButton = document.querySelector<HTMLButtonElement>('#wordSearchCloseButton');
  const boldButton = document.querySelector<HTMLButtonElement>('#boldButton');
  const italicButton = document.querySelector<HTMLButtonElement>('#italicButton');
  const underlineButton = document.querySelector<HTMLButtonElement>('#underlineButton');
  const strikeButton = document.querySelector<HTMLButtonElement>('#strikeButton');
  const fontFamilySelect = document.querySelector<HTMLSelectElement>('#fontFamilySelect');
  const fontSizeSelect = document.querySelector<HTMLSelectElement>('#fontSizeSelect');
  const fontColorInput = document.querySelector<HTMLInputElement>('#fontColorInput');
  const applyHighlightButton = document.querySelector<HTMLButtonElement>('#applyHighlightButton');
  const highlightColorInput = document.querySelector<HTMLInputElement>('#highlightColorInput');
  const clearHighlightButton = document.querySelector<HTMLButtonElement>('#clearHighlightButton');
  const clearFormatButton = document.querySelector<HTMLButtonElement>('#clearFormatButton');
  const hyperlinkInput = document.querySelector<HTMLInputElement>('#hyperlinkInput');
  const setHyperlinkButton = document.querySelector<HTMLButtonElement>('#setHyperlinkButton');
  const removeHyperlinkButton = document.querySelector<HTMLButtonElement>('#removeHyperlinkButton');
  const bulletListButton = document.querySelector<HTMLButtonElement>('#bulletListButton');
  const numberedListButton = document.querySelector<HTMLButtonElement>('#numberedListButton');
  const decreaseListLevelButton = document.querySelector<HTMLButtonElement>(
    '#decreaseListLevelButton',
  );
  const increaseListLevelButton = document.querySelector<HTMLButtonElement>(
    '#increaseListLevelButton',
  );
  const removeListButton = document.querySelector<HTMLButtonElement>('#removeListButton');
  const decreaseIndentButton = document.querySelector<HTMLButtonElement>('#decreaseIndentButton');
  const increaseIndentButton = document.querySelector<HTMLButtonElement>('#increaseIndentButton');
  const firstLineIndentSelect = document.querySelector<HTMLSelectElement>('#firstLineIndentSelect');
  const lineSpacingSelect = document.querySelector<HTMLSelectElement>('#lineSpacingSelect');
  const spacingBeforeSelect = document.querySelector<HTMLSelectElement>('#spacingBeforeSelect');
  const spacingAfterSelect = document.querySelector<HTMLSelectElement>('#spacingAfterSelect');
  const insertPageBreakButton = document.querySelector<HTMLButtonElement>('#insertPageBreakButton');
  const insertImageButton = document.querySelector<HTMLButtonElement>('#insertImageButton');
  const imageFileInput = document.querySelector<HTMLInputElement>('#imageFileInput');
  const imageWidthInput = document.querySelector<HTMLInputElement>('#imageWidthInput');
  const imageHeightInput = document.querySelector<HTMLInputElement>('#imageHeightInput');
  const deleteImageButton = document.querySelector<HTMLButtonElement>('#deleteImageButton');
  const newTableRowsInput = document.querySelector<HTMLInputElement>('#newTableRowsInput');
  const newTableColumnsInput = document.querySelector<HTMLInputElement>('#newTableColumnsInput');
  const insertTableButton = document.querySelector<HTMLButtonElement>('#insertTableButton');
  const deleteTableButton = document.querySelector<HTMLButtonElement>('#deleteTableButton');
  const insertTableRowBeforeButton = document.querySelector<HTMLButtonElement>(
    '#insertTableRowBeforeButton',
  );
  const insertTableRowAfterButton = document.querySelector<HTMLButtonElement>(
    '#insertTableRowAfterButton',
  );
  const deleteTableRowButton = document.querySelector<HTMLButtonElement>('#deleteTableRowButton');
  const insertTableColumnBeforeButton = document.querySelector<HTMLButtonElement>(
    '#insertTableColumnBeforeButton',
  );
  const insertTableColumnAfterButton = document.querySelector<HTMLButtonElement>(
    '#insertTableColumnAfterButton',
  );
  const deleteTableColumnButton = document.querySelector<HTMLButtonElement>(
    '#deleteTableColumnButton',
  );
  const mergeTableCellsButton = document.querySelector<HTMLButtonElement>('#mergeTableCellsButton');
  const splitTableCellButton = document.querySelector<HTMLButtonElement>('#splitTableCellButton');
  const tableCellBackgroundInput = document.querySelector<HTMLInputElement>(
    '#tableCellBackgroundInput',
  );
  const clearTableCellBackgroundButton = document.querySelector<HTMLButtonElement>(
    '#clearTableCellBackgroundButton',
  );
  const tableCellVerticalAlignmentSelect = document.querySelector<HTMLSelectElement>(
    '#tableCellVerticalAlignmentSelect',
  );
  const tableCellBorderColorInput = document.querySelector<HTMLInputElement>(
    '#tableCellBorderColorInput',
  );
  const applyTableCellBordersButton = document.querySelector<HTMLButtonElement>(
    '#applyTableCellBordersButton',
  );
  const clearTableCellBordersButton = document.querySelector<HTMLButtonElement>(
    '#clearTableCellBordersButton',
  );
  const alignmentButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-align]'));
  const updateSearchCount = () => {
    const state = editor.getSearchState();
    if (searchCount) {
      searchCount.textContent =
        state && state.matches.length > 0
          ? `${state.activeIndex + 1} / ${state.matches.length}`
          : '0 / 0';
    }
    if (searchPreviousButton) searchPreviousButton.disabled = !state?.matches.length;
    if (searchNextButton) searchNextButton.disabled = !state?.matches.length;
    if (replaceCurrentButton) replaceCurrentButton.disabled = !state?.matches.length;
    if (replaceAllButton) replaceAllButton.disabled = !state?.matches.length;
  };
  const runSearch = () => {
    editor.search({
      query: searchInput?.value ?? '',
      matchCase: searchMatchCase?.checked ?? false,
      wholeWord: searchWholeWord?.checked ?? false,
    });
    updateSearchCount();
  };
  const openSearch = () => {
    if (searchBar) searchBar.hidden = false;
    searchInput?.focus();
    searchInput?.select();
    runSearch();
  };
  const closeSearch = () => {
    editor.clearSearch();
    if (searchBar) searchBar.hidden = true;
    updateSearchCount();
    editor.focus();
  };
  searchInput?.addEventListener('input', runSearch);
  searchInput?.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      editor.findNext(event.shiftKey ? 'backward' : 'forward');
      updateSearchCount();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      closeSearch();
    }
  });
  searchMatchCase?.addEventListener('change', runSearch);
  searchWholeWord?.addEventListener('change', runSearch);
  searchPreviousButton?.addEventListener('click', () => {
    editor.findNext('backward');
    updateSearchCount();
  });
  searchNextButton?.addEventListener('click', () => {
    editor.findNext('forward');
    updateSearchCount();
  });
  replaceCurrentButton?.addEventListener('click', () => {
    editor.replaceCurrent(replaceInput?.value ?? '');
    updateSearchCount();
  });
  replaceAllButton?.addEventListener('click', () => {
    editor.replaceAll(replaceInput?.value ?? '');
    updateSearchCount();
  });
  searchCloseButton?.addEventListener('click', closeSearch);
  document.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openSearch();
    }
  });
  const textFormatButtons = [
    boldButton,
    italicButton,
    underlineButton,
    strikeButton,
    applyHighlightButton,
    clearHighlightButton,
    clearFormatButton,
    setHyperlinkButton,
    removeHyperlinkButton,
  ].filter((button): button is HTMLButtonElement => button !== null);
  textFormatButtons.forEach((button) =>
    button.addEventListener('mousedown', (event) => event.preventDefault()),
  );
  undoButton?.addEventListener('mousedown', (event) => event.preventDefault());
  redoButton?.addEventListener('mousedown', (event) => event.preventDefault());
  undoButton?.addEventListener('click', () => editor.undo());
  redoButton?.addEventListener('click', () => editor.redo());
  trackRevisionsButton?.addEventListener('mousedown', (event) => event.preventDefault());
  trackRevisionsButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'toggle-track-revisions' }),
  );
  acceptRevisionButton?.addEventListener('mousedown', (event) => event.preventDefault());
  rejectRevisionButton?.addEventListener('mousedown', (event) => event.preventDefault());
  acceptAllRevisionsButton?.addEventListener('mousedown', (event) => event.preventDefault());
  rejectAllRevisionsButton?.addEventListener('mousedown', (event) => event.preventDefault());
  acceptRevisionButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'accept-revision' }),
  );
  rejectRevisionButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'reject-revision' }),
  );
  acceptAllRevisionsButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'accept-all-revisions' }),
  );
  rejectAllRevisionsButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'reject-all-revisions' }),
  );
  boldButton?.addEventListener('click', () => editor.dispatch({ kind: 'toggle-bold' }));
  italicButton?.addEventListener('click', () => editor.dispatch({ kind: 'toggle-italic' }));
  underlineButton?.addEventListener('click', () => editor.dispatch({ kind: 'toggle-underline' }));
  strikeButton?.addEventListener('click', () => editor.dispatch({ kind: 'toggle-strike' }));
  applyHighlightButton?.addEventListener('click', () => {
    editor.dispatch({
      kind: 'set-highlight-color',
      color: highlightColorInput?.value ?? '#FFFF00',
    });
  });
  clearHighlightButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'set-highlight-color', color: null }),
  );
  clearFormatButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'clear-text-format' }),
  );
  setHyperlinkButton?.addEventListener('click', () => {
    const value = hyperlinkInput?.value.trim() ?? '';
    if (!value) return;
    editor.dispatch({
      kind: 'set-hyperlink',
      hyperlink: value.startsWith('#') ? { anchor: value.slice(1) } : { target: value },
    });
  });
  removeHyperlinkButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'remove-hyperlink' }),
  );

  fontFamilySelect?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-font-family',
      fontFamily: (event.currentTarget as HTMLSelectElement).value,
    });
  });

  fontSizeSelect?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-font-size',
      fontSize: Number((event.currentTarget as HTMLSelectElement).value),
    });
  });

  fontColorInput?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-text-color',
      color: (event.currentTarget as HTMLInputElement).value,
    });
  });

  let lastHighlightPickerValue: string | null = null;
  const applyHighlightPickerValue = (event: Event) => {
    const color = (event.currentTarget as HTMLInputElement).value.toLowerCase();
    if (color === lastHighlightPickerValue) return;
    lastHighlightPickerValue = color;
    editor.dispatch({ kind: 'set-highlight-color', color });
  };
  highlightColorInput?.addEventListener('pointerdown', () => {
    lastHighlightPickerValue = null;
  });
  highlightColorInput?.addEventListener('input', applyHighlightPickerValue);
  highlightColorInput?.addEventListener('change', applyHighlightPickerValue);

  alignmentButtons.forEach((button) => {
    button.addEventListener('mousedown', (event) => event.preventDefault());
    button.addEventListener('click', () => {
      const alignment = button.dataset.align;
      if (
        alignment === 'left' ||
        alignment === 'center' ||
        alignment === 'right' ||
        alignment === 'both'
      ) {
        editor.dispatch({ kind: 'set-paragraph-alignment', alignment });
      }
    });
  });

  const listButtons = [
    bulletListButton,
    numberedListButton,
    decreaseListLevelButton,
    increaseListLevelButton,
    removeListButton,
  ].filter((button): button is HTMLButtonElement => button !== null);
  listButtons.forEach((button) =>
    button.addEventListener('mousedown', (event) => event.preventDefault()),
  );
  bulletListButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'toggle-list', listType: 'bullet' }),
  );
  numberedListButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'toggle-list', listType: 'numbered' }),
  );
  decreaseListLevelButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'adjust-list-level', delta: -1 }),
  );
  increaseListLevelButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'adjust-list-level', delta: 1 }),
  );
  removeListButton?.addEventListener('click', () => editor.dispatch({ kind: 'remove-list' }));

  decreaseIndentButton?.addEventListener('mousedown', (event) => event.preventDefault());
  decreaseIndentButton?.addEventListener('click', () => {
    editor.dispatch({ kind: 'adjust-paragraph-left-indent', delta: -720 });
  });
  increaseIndentButton?.addEventListener('mousedown', (event) => event.preventDefault());
  increaseIndentButton?.addEventListener('click', () => {
    editor.dispatch({ kind: 'adjust-paragraph-left-indent', delta: 720 });
  });
  firstLineIndentSelect?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-paragraph-indent',
      indent: { firstLine: Number((event.currentTarget as HTMLSelectElement).value) },
    });
  });
  lineSpacingSelect?.addEventListener('change', (event) => {
    const line = Number((event.currentTarget as HTMLSelectElement).value);
    if (Number.isFinite(line))
      editor.dispatch({ kind: 'set-paragraph-line-spacing', line, lineRule: 'auto' });
  });
  spacingBeforeSelect?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-paragraph-spacing',
      spacing: { before: Number((event.currentTarget as HTMLSelectElement).value) },
    });
  });
  spacingAfterSelect?.addEventListener('change', (event) => {
    editor.dispatch({
      kind: 'set-paragraph-spacing',
      spacing: { after: Number((event.currentTarget as HTMLSelectElement).value) },
    });
  });

  insertPageBreakButton?.addEventListener('mousedown', (event) => event.preventDefault());
  insertPageBreakButton?.addEventListener('click', () => {
    editor.dispatch({ kind: 'insert-page-break' });
  });
  insertImageButton?.addEventListener('mousedown', (event) => event.preventDefault());
  insertImageButton?.addEventListener('click', () => imageFileInput?.click());
  imageFileInput?.addEventListener('change', () => {
    const file = imageFileInput.files?.[0];
    if (!file) return;
    void editor.insertImage(file).finally(() => {
      imageFileInput.value = '';
    });
  });
  deleteImageButton?.addEventListener('mousedown', (event) => event.preventDefault());
  deleteImageButton?.addEventListener('click', () => {
    const image = editor.getCommandState().image;
    if (image)
      editor.dispatch({ kind: 'delete-object', objectKind: 'image', objectId: image.objectId });
  });
  const resizeSelectedImage = (axis: 'width' | 'height') => {
    const image = editor.getCommandState().image;
    if (!image) return;
    const currentWidth = image.width ?? Number(imageWidthInput?.value);
    const currentHeight = image.height ?? Number(imageHeightInput?.value);
    const nextValue = Number(axis === 'width' ? imageWidthInput?.value : imageHeightInput?.value);
    if (!Number.isFinite(nextValue) || nextValue < 24 || !currentWidth || !currentHeight) return;
    const ratio = currentWidth / currentHeight;
    editor.dispatch({
      kind: 'set-image-size',
      objectId: image.objectId,
      width: axis === 'width' ? nextValue : nextValue * ratio,
      height: axis === 'height' ? nextValue : nextValue / ratio,
    });
  };
  imageWidthInput?.addEventListener('change', () => resizeSelectedImage('width'));
  imageHeightInput?.addEventListener('change', () => resizeSelectedImage('height'));
  const tableButtons = [
    insertTableButton,
    deleteTableButton,
    insertTableRowBeforeButton,
    insertTableRowAfterButton,
    deleteTableRowButton,
    insertTableColumnBeforeButton,
    insertTableColumnAfterButton,
    deleteTableColumnButton,
    mergeTableCellsButton,
    splitTableCellButton,
  ].filter((button): button is HTMLButtonElement => button !== null);
  tableButtons.forEach((button) =>
    button.addEventListener('mousedown', (event) => event.preventDefault()),
  );
  insertTableButton?.addEventListener('click', () => {
    const rows = Number(newTableRowsInput?.value ?? 2);
    const columns = Number(newTableColumnsInput?.value ?? 2);
    editor.dispatch({ kind: 'insert-table', rows, columns });
  });
  deleteTableButton?.addEventListener('click', () => editor.dispatch({ kind: 'delete-table' }));
  insertTableRowBeforeButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'insert-table-row', position: 'before' }),
  );
  insertTableRowAfterButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'insert-table-row', position: 'after' }),
  );
  deleteTableRowButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'delete-table-row' }),
  );
  insertTableColumnBeforeButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'insert-table-column', position: 'before' }),
  );
  insertTableColumnAfterButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'insert-table-column', position: 'after' }),
  );
  deleteTableColumnButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'delete-table-column' }),
  );
  mergeTableCellsButton?.addEventListener('mousedown', (event) => event.preventDefault());
  mergeTableCellsButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'merge-table-cells' }),
  );
  splitTableCellButton?.addEventListener('mousedown', (event) => event.preventDefault());
  splitTableCellButton?.addEventListener('click', () =>
    editor.dispatch({ kind: 'split-table-cell' }),
  );
  tableCellBackgroundInput?.addEventListener('change', () => {
    editor.dispatch({ kind: 'set-table-cell-background', color: tableCellBackgroundInput.value });
  });
  clearTableCellBackgroundButton?.addEventListener('mousedown', (event) => event.preventDefault());
  clearTableCellBackgroundButton?.addEventListener('click', () => {
    editor.dispatch({ kind: 'set-table-cell-background', color: null });
  });
  tableCellVerticalAlignmentSelect?.addEventListener('change', () => {
    const alignment = tableCellVerticalAlignmentSelect.value;
    if (alignment === 'top' || alignment === 'center' || alignment === 'bottom') {
      editor.dispatch({ kind: 'set-table-cell-vertical-alignment', alignment });
    }
  });
  applyTableCellBordersButton?.addEventListener('mousedown', (event) => event.preventDefault());
  applyTableCellBordersButton?.addEventListener('click', () => {
    editor.dispatch({
      kind: 'set-table-cell-borders',
      color: tableCellBorderColorInput?.value ?? '#000000',
    });
  });
  clearTableCellBordersButton?.addEventListener('mousedown', (event) => event.preventDefault());
  clearTableCellBordersButton?.addEventListener('click', () => {
    editor.dispatch({ kind: 'set-table-cell-borders', color: null });
  });

  const update = () => {
    const state = editor.getCommandState();
    if (undoButton) undoButton.disabled = !editor.canUndo();
    if (redoButton) redoButton.disabled = !editor.canRedo();
    if (trackRevisionsButton) {
      trackRevisionsButton.classList.toggle('is-active', state.trackRevisions);
      trackRevisionsButton.setAttribute('aria-pressed', String(state.trackRevisions));
    }
    const canFormatText = editor.canDispatch({ kind: 'toggle-bold' });
    updateToggleButton(boldButton, state.bold, canFormatText);
    updateToggleButton(italicButton, state.italic, canFormatText);
    updateToggleButton(underlineButton, state.underline, canFormatText);
    updateToggleButton(strikeButton, state.strike, canFormatText);
    if (fontFamilySelect) {
      fontFamilySelect.disabled = !canFormatText;
      setToolbarSelectValue(fontFamilySelect, state.fontFamily);
    }
    if (fontSizeSelect) {
      fontSizeSelect.disabled = !canFormatText;
      setToolbarSelectValue(fontSizeSelect, state.fontSize);
    }
    if (fontColorInput) {
      fontColorInput.disabled = !canFormatText;
      fontColorInput.classList.toggle('is-mixed', state.textColor === 'mixed');
      if (typeof state.textColor === 'string' && state.textColor.startsWith('#'))
        fontColorInput.value = state.textColor;
    }
    if (highlightColorInput) {
      highlightColorInput.disabled = !canFormatText;
      highlightColorInput.classList.toggle('is-mixed', state.highlightColor === 'mixed');
      if (typeof state.highlightColor === 'string' && state.highlightColor.startsWith('#')) {
        highlightColorInput.value = state.highlightColor;
      }
    }
    if (applyHighlightButton) {
      applyHighlightButton.disabled = !canFormatText;
      applyHighlightButton.classList.toggle(
        'is-active',
        state.highlightColor !== null && state.highlightColor !== 'mixed',
      );
      applyHighlightButton.classList.toggle('is-mixed', state.highlightColor === 'mixed');
      applyHighlightButton.style.setProperty(
        '--coco-highlight-color',
        highlightColorInput?.value ?? '#FFFF00',
      );
    }
    if (clearHighlightButton)
      clearHighlightButton.disabled = !canFormatText || state.highlightColor === null;
    if (clearFormatButton) clearFormatButton.disabled = !canFormatText;
    if (hyperlinkInput && state.hyperlink !== 'mixed' && state.hyperlink) {
      hyperlinkInput.value = state.hyperlink.target ?? `#${state.hyperlink.anchor ?? ''}`;
    }
    const hyperlinkValue = hyperlinkInput?.value.trim() ?? '';
    if (setHyperlinkButton) {
      setHyperlinkButton.disabled = !editor.canDispatch({
        kind: 'set-hyperlink',
        hyperlink: hyperlinkValue.startsWith('#')
          ? { anchor: hyperlinkValue.slice(1) }
          : { target: hyperlinkValue },
      });
    }
    if (removeHyperlinkButton) {
      removeHyperlinkButton.disabled = !editor.canDispatch({ kind: 'remove-hyperlink' });
    }
    for (const button of alignmentButtons) {
      const alignment = button.dataset.align;
      button.disabled = !editor.canDispatch({ kind: 'set-paragraph-alignment', alignment: 'left' });
      button.classList.toggle('is-active', state.paragraphAlignment === alignment);
      button.classList.toggle('is-mixed', state.paragraphAlignment === 'mixed');
      button.setAttribute('aria-pressed', String(state.paragraphAlignment === alignment));
    }
    const canFormatParagraph = editor.canDispatch({
      kind: 'set-paragraph-alignment',
      alignment: 'left',
    });
    const canToggleList = editor.canDispatch({ kind: 'toggle-list', listType: 'bullet' });
    const hasList = state.listType !== null;
    if (bulletListButton) {
      bulletListButton.disabled = !canToggleList;
      bulletListButton.classList.toggle('is-active', state.listType === 'bullet');
      bulletListButton.classList.toggle('is-mixed', state.listType === 'mixed');
      bulletListButton.setAttribute('aria-pressed', String(state.listType === 'bullet'));
    }
    if (numberedListButton) {
      numberedListButton.disabled = !canToggleList;
      numberedListButton.classList.toggle('is-active', state.listType === 'numbered');
      numberedListButton.classList.toggle('is-mixed', state.listType === 'mixed');
      numberedListButton.setAttribute('aria-pressed', String(state.listType === 'numbered'));
    }
    if (decreaseListLevelButton) {
      decreaseListLevelButton.disabled = !canToggleList || !hasList || state.listLevel === 0;
    }
    if (increaseListLevelButton) increaseListLevelButton.disabled = !canToggleList || !hasList;
    if (removeListButton) removeListButton.disabled = !canToggleList || !hasList;
    if (decreaseIndentButton) {
      decreaseIndentButton.disabled = !canFormatParagraph || state.leftIndent === 0;
    }
    if (increaseIndentButton) increaseIndentButton.disabled = !canFormatParagraph;
    if (firstLineIndentSelect) {
      firstLineIndentSelect.disabled = !canFormatParagraph;
      setToolbarSelectValue(firstLineIndentSelect, state.firstLineIndent);
    }
    if (lineSpacingSelect) {
      lineSpacingSelect.disabled = !canFormatParagraph;
      setLineSpacingSelectValue(lineSpacingSelect, state.lineSpacing);
    }
    if (spacingBeforeSelect) {
      spacingBeforeSelect.disabled = !canFormatParagraph;
      setToolbarSelectValue(spacingBeforeSelect, state.spacingBefore);
    }
    if (spacingAfterSelect) {
      spacingAfterSelect.disabled = !canFormatParagraph;
      setToolbarSelectValue(spacingAfterSelect, state.spacingAfter);
    }
    if (insertImageButton) {
      insertImageButton.disabled = !editor.canDispatch({
        kind: 'insert-image',
        src: 'data:image/png;base64,',
      });
    }
    if (insertPageBreakButton) {
      insertPageBreakButton.disabled = !editor.canDispatch({ kind: 'insert-page-break' });
    }
    if (imageWidthInput) {
      imageWidthInput.disabled = !state.image;
      imageWidthInput.value = state.image?.width ? String(Math.round(state.image.width)) : '';
    }
    if (imageHeightInput) {
      imageHeightInput.disabled = !state.image;
      imageHeightInput.value = state.image?.height ? String(Math.round(state.image.height)) : '';
    }
    if (deleteImageButton) deleteImageButton.disabled = !state.image;
    const rows = Number(newTableRowsInput?.value ?? 2);
    const columns = Number(newTableColumnsInput?.value ?? 2);
    if (insertTableButton) {
      insertTableButton.disabled = !editor.canDispatch({ kind: 'insert-table', rows, columns });
    }
    if (deleteTableButton)
      deleteTableButton.disabled = !editor.canDispatch({ kind: 'delete-table' });
    if (insertTableRowBeforeButton) {
      insertTableRowBeforeButton.disabled = !editor.canDispatch({
        kind: 'insert-table-row',
        position: 'before',
      });
    }
    if (insertTableRowAfterButton) {
      insertTableRowAfterButton.disabled = !editor.canDispatch({
        kind: 'insert-table-row',
        position: 'after',
      });
    }
    if (deleteTableRowButton) {
      deleteTableRowButton.disabled = !editor.canDispatch({ kind: 'delete-table-row' });
    }
    if (insertTableColumnBeforeButton) {
      insertTableColumnBeforeButton.disabled = !editor.canDispatch({
        kind: 'insert-table-column',
        position: 'before',
      });
    }
    if (insertTableColumnAfterButton) {
      insertTableColumnAfterButton.disabled = !editor.canDispatch({
        kind: 'insert-table-column',
        position: 'after',
      });
    }
    if (deleteTableColumnButton) {
      deleteTableColumnButton.disabled = !editor.canDispatch({ kind: 'delete-table-column' });
    }
    if (mergeTableCellsButton) {
      mergeTableCellsButton.disabled = !editor.canDispatch({ kind: 'merge-table-cells' });
    }
    if (splitTableCellButton) {
      splitTableCellButton.disabled = !editor.canDispatch({ kind: 'split-table-cell' });
    }
    const canFormatCell = editor.canDispatch({
      kind: 'set-table-cell-background',
      color: '#FFFFFF',
    });
    if (tableCellBackgroundInput) {
      tableCellBackgroundInput.disabled = !canFormatCell;
      if (state.table?.backgroundColor && state.table.backgroundColor !== 'mixed') {
        tableCellBackgroundInput.value = state.table.backgroundColor;
      }
    }
    if (clearTableCellBackgroundButton) {
      clearTableCellBackgroundButton.disabled = !canFormatCell || !state.table?.backgroundColor;
    }
    if (tableCellVerticalAlignmentSelect) {
      tableCellVerticalAlignmentSelect.disabled = !canFormatCell;
      tableCellVerticalAlignmentSelect.value =
        state.table?.verticalAlignment === 'bottom'
          ? 'bottom'
          : state.table?.verticalAlignment === 'center' || state.table?.verticalAlignment === 'both'
            ? 'center'
            : 'top';
    }
    if (tableCellBorderColorInput) {
      tableCellBorderColorInput.disabled = !canFormatCell;
      if (state.table?.borderColor && state.table.borderColor !== 'mixed') {
        tableCellBorderColorInput.value = state.table.borderColor;
      }
    }
    if (applyTableCellBordersButton) applyTableCellBordersButton.disabled = !canFormatCell;
    if (clearTableCellBordersButton) {
      clearTableCellBordersButton.disabled = !canFormatCell || !state.table?.bordersVisible;
    }
  };

  editor.subscribe('selectionChange', update);
  editor.subscribe('change', () => {
    update();
    updateSearchCount();
  });
  hyperlinkInput?.addEventListener('input', update);
  update();
  updateSearchCount();
}

function updateToggleButton(
  button: HTMLButtonElement | null,
  value: boolean | 'mixed',
  enabled: boolean,
): void {
  if (!button) return;
  button.disabled = !enabled;
  button.classList.toggle('is-active', value === true);
  button.classList.toggle('is-mixed', value === 'mixed');
  button.setAttribute('aria-pressed', String(value));
}

function setToolbarSelectValue(
  select: HTMLSelectElement,
  value: string | number | 'mixed' | null,
): void {
  select.querySelector('option[data-runtime-value]')?.remove();
  if (value === 'mixed' || value === null) {
    const option = new Option(value === 'mixed' ? '混合' : '-', '', true, true);
    option.dataset.runtimeValue = 'true';
    select.prepend(option);
    return;
  }
  if (!Array.from(select.options).some((option) => option.value === String(value))) {
    const option = new Option(String(value), String(value), true, true);
    option.dataset.runtimeValue = 'true';
    select.prepend(option);
  }
  select.value = String(value);
}

function setLineSpacingSelectValue(
  select: HTMLSelectElement,
  value: WordEditorLineSpacingState | 'mixed' | null,
): void {
  select.querySelector('option[data-runtime-value]')?.remove();
  if (value === 'mixed' || value === null) {
    setToolbarSelectValue(select, value);
    return;
  }
  if (value.lineRule === 'auto') {
    const optionValue = String(value.line);
    if (Array.from(select.options).some((option) => option.value === optionValue)) {
      select.value = optionValue;
      return;
    }
    const multiplier = Math.round((value.line / 240) * 100) / 100;
    const option = new Option(String(multiplier), optionValue, true, true);
    option.dataset.runtimeValue = 'true';
    select.prepend(option);
    return;
  }
  const label = `${value.line / 20} pt ${value.lineRule === 'exact' ? '固定' : '最小'}`;
  const option = new Option(label, `${value.lineRule}:${value.line}`, true, true);
  option.dataset.runtimeValue = 'true';
  select.prepend(option);
}

function summarizeWordView(
  fixturePath: string,
  byteLength: number,
  wordView: Awaited<ReturnType<typeof buildWordEditorViewFromPackage>>,
): unknown {
  return {
    fixturePath,
    bytes: byteLength,
    paragraphs: wordView.paragraphs.length,
    tables: wordView.tables.length,
    blocks: wordView.blocks.length,
    pageLayout: wordView.pageLayout,
    sections: wordView.sections,
    headers: wordView.headers.map((story) => ({
      partName: story.partName,
      paragraphs: story.paragraphs.length,
      tables: story.tables.length,
      preview: story.paragraphs.slice(0, 3).map((paragraph) => paragraph.text),
    })),
    footers: wordView.footers.map((story) => ({
      partName: story.partName,
      paragraphs: story.paragraphs.length,
      tables: story.tables.length,
      preview: story.paragraphs.slice(0, 3).map((paragraph) => paragraph.text),
    })),
    coverage: summarizeEditorCoverage(wordView),
    comments: wordView.paragraphs.reduce(
      (count, paragraph) =>
        count + paragraph.spans.filter((span) => span.kind === 'comment').length,
      0,
    ),
    revisions: wordView.paragraphs.reduce(
      (count, paragraph) =>
        count + paragraph.spans.filter((span) => span.kind === 'revision').length,
      0,
    ),
    images: wordView.paragraphs.reduce(
      (count, paragraph) =>
        count + paragraph.segments.filter((segment) => segment.kind === 'image').length,
      0,
    ),
    preview: wordView.paragraphs.slice(0, 12).map((paragraph) => ({
      ...paragraph,
      segments: paragraph.segments.map((segment) =>
        segment.kind === 'image'
          ? {
              ...segment,
              src: `${segment.src.slice(0, 40)}...`,
            }
          : segment,
      ),
    })),
  };
}

async function summarizeSavedDocx(bytes: Uint8Array, fixturePath: string): Promise<unknown> {
  const pkg = await fromZip(bytes);
  const wordView = buildWordEditorViewFromPackage(pkg, buildIndices(pkg));
  return {
    fixturePath,
    bytes: bytes.byteLength,
    paragraphs: wordView.paragraphs.length,
    tables: wordView.tables.length,
    pageLayout: wordView.pageLayout,
    images: wordView.paragraphs.reduce(
      (count, paragraph) =>
        count + paragraph.segments.filter((segment) => segment.kind === 'image').length,
      0,
    ),
    preview: wordView.paragraphs.slice(0, 6).map((paragraph) => ({
      paraId: paragraph.paraId,
      text: paragraph.text.slice(0, 160),
      segments: paragraph.segments.map((segment) =>
        segment.kind === 'image'
          ? {
              kind: 'image',
              from: segment.from,
              to: segment.to,
              drawingId: segment.drawingId,
              width: segment.width,
              height: segment.height,
            }
          : {
              kind: segment.kind,
              from: segment.from,
              to: segment.to,
              text:
                'text' in segment && typeof segment.text === 'string'
                  ? segment.text.slice(0, 80)
                  : undefined,
            },
      ),
    })),
  };
}

main().catch((err) => {
  const app = document.querySelector<HTMLDivElement>('#app');
  if (app) app.textContent = err instanceof Error ? err.stack || err.message : String(err);
});
