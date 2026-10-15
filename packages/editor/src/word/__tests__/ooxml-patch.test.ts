import { describe, expect, test } from 'vitest';
import { fromFlatOpc, getPart, isDirty, toFlatOpc } from '@coco-vexel/coco-office/pack';
import { buildIndices } from '@coco-vexel/coco-office/word-view';
import {
  applyWordOoxmlPatches,
  planWordOoxmlPatch,
  planWordOoxmlPatches,
  resolveWordOoxmlPatchTargets,
  type NativeWordChange,
} from '../native.js';
import { buildWordEditorViewFromPackage } from '../extract.js';

describe('@coco-vexel/coco-office-editor Word OOXML patch planning', () => {
  test('does not plan image insertion or resize as drawing deletion', () => {
    const changes: NativeWordChange[] = [
      {
        kind: 'image',
        op: 'insert',
        paraId: 'P1',
        offset: 0,
        drawingId: 'D1',
        src: 'data:image/png;base64,image',
      },
      {
        kind: 'image',
        op: 'resize',
        paraId: 'P1',
        offset: 0,
        drawingId: 'D1',
        width: 320,
        height: 180,
      },
    ];

    const plan = planWordOoxmlPatches(changes);

    expect(plan.patches).toEqual([
      {
        kind: 'drawing-resize',
        paraId: 'P1',
        drawingId: 'D1',
        offset: 0,
        width: 320,
        height: 180,
      },
    ]);
    expect(plan.unsupported.map((entry) => entry.change)).toEqual([changes[0]]);
  });

  test('plans supported table structure edits for export', () => {
    const change: NativeWordChange = {
      kind: 'table',
      op: 'insert-row',
      tableId: 'table-1',
      rowIndex: 1,
      columnIndex: 0,
    };

    const plan = planWordOoxmlPatches([change]);

    expect(plan.patches).toEqual([
      {
        kind: 'table-structure',
        op: 'insert-row',
        tableId: 'table-1',
        rowIndex: 1,
        columnIndex: 0,
      },
    ]);
    expect(plan.unsupported).toEqual([]);
  });

  test('does not silently commit a tracked table structure change', () => {
    const change: NativeWordChange = {
      kind: 'table',
      op: 'insert-row',
      tableId: 'table-1',
      rowIndex: 1,
      columnIndex: 0,
      revisionId: 'table-revision-1',
      revisionKind: 'insert',
    };
    const plan = planWordOoxmlPatches([change]);
    expect(plan.patches).toEqual([]);
    expect(plan.unsupported[0]?.reason).toContain('original table structure snapshot');
  });

  test('collapses tracked table changes by their final accept or reject action', () => {
    const initial: NativeWordChange = {
      kind: 'table', op: 'insert-row', tableId: 'table-1', rowIndex: 1, columnIndex: 0,
      revisionId: 'table-revision-2', revisionKind: 'insert',
    };
    const reject: NativeWordChange = {
      ...initial, op: 'insert-row', revisionAction: 'reject',
    };
    expect(planWordOoxmlPatches([initial, reject])).toMatchObject({ patches: [], unsupported: [] });

    const accept: NativeWordChange = {
      ...initial, op: 'insert-row', revisionAction: 'accept',
    };
    expect(planWordOoxmlPatches([initial, accept])).toMatchObject({
      unsupported: [],
      patches: [{ kind: 'table-structure', op: 'insert-row', tableId: 'table-1' }],
    });

    const deleteTable: NativeWordChange = {
      kind: 'table', op: 'delete-table', tableId: 'table-1', rowIndex: 0, columnIndex: 0,
      revisionId: 'table-revision-3', revisionKind: 'delete',
    };
    const acceptDelete: NativeWordChange = { ...deleteTable, revisionAction: 'accept' };
    const rejectDelete: NativeWordChange = { ...deleteTable, revisionAction: 'reject' };
    expect(planWordOoxmlPatches([deleteTable, rejectDelete])).toMatchObject({ patches: [], unsupported: [] });
    expect(planWordOoxmlPatches([deleteTable, acceptDelete])).toMatchObject({
      unsupported: [],
      patches: [{ kind: 'table-delete', tableId: 'table-1' }],
    });
  });

  test('maps native text changes to paragraph text patches', () => {
    expect(
      planWordOoxmlPatch({
        kind: 'text',
        op: 'insert',
        paraId: 'P0000001',
        offset: 3,
        text: 'A',
      }),
    ).toEqual({
      kind: 'paragraph-text',
      op: 'insert',
      paraId: 'P0000001',
      offset: 3,
      text: 'A',
    });

    expect(
      planWordOoxmlPatch({
        kind: 'text',
        op: 'replace',
        paraId: 'P0000001',
        offset: 3,
        text: 'B',
        deletedText: 'A',
      }),
    ).toEqual({
      kind: 'paragraph-text',
      op: 'replace',
      paraId: 'P0000001',
      offset: 3,
      text: 'B',
      deletedText: 'A',
    });
  });

  test('maps block and image changes to structural patches', () => {
    const changes: NativeWordChange[] = [
      {
        kind: 'block',
        op: 'split',
        paraId: 'P0000001',
        newParaId: 'P0000002',
        offset: 4,
      },
      {
        kind: 'block',
        op: 'merge',
        paraId: 'P0000001',
      },
      {
        kind: 'image',
        op: 'delete',
        paraId: 'P0000003',
        offset: 2,
        drawingId: 'drawing-1',
      },
    ];

    expect(planWordOoxmlPatches(changes)).toEqual({
      target: 'word/document.xml',
      unsupported: [],
      patches: [
        {
          kind: 'paragraph-split',
          paraId: 'P0000001',
          newParaId: 'P0000002',
          offset: 4,
        },
        {
          kind: 'paragraph-merge',
          paraId: 'P0000001',
        },
        {
          kind: 'drawing-delete',
          paraId: 'P0000003',
          drawingId: 'drawing-1',
          offset: 2,
        },
      ],
    });
  });

  test('reports incomplete split changes as unsupported', () => {
    const change: NativeWordChange = {
      kind: 'block',
      op: 'split',
      paraId: 'P0000001',
    };

    expect(planWordOoxmlPatches([change])).toEqual({
      target: 'word/document.xml',
      patches: [],
      unsupported: [
        {
          change,
          reason: 'paragraph split change is missing newParaId or offset',
        },
      ],
    });
  });

  test('resolves paragraph and drawing patch targets against package indices', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>
          <w:p w14:paraId="P0000002"><w:r><w:drawing/></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);
    const indices = buildIndices(pkg);
    const drawingId = Array.from(indices.drawingIdx.keys())[0]!;

    const result = resolveWordOoxmlPatchTargets(pkg, indices, [
      {
        kind: 'paragraph-text',
        op: 'insert',
        paraId: 'P0000001',
        offset: 5,
        text: '!',
      },
      {
        kind: 'drawing-delete',
        paraId: 'P0000002',
        drawingId,
        offset: 0,
      },
    ]);

    expect(result.unresolved).toEqual([]);
    expect(result.resolved).toHaveLength(2);
    expect(result.resolved[0]).toMatchObject({
      kind: 'paragraph',
      partPath: 'word/document.xml',
    });
    expect(result.resolved[1]).toMatchObject({
      kind: 'drawing',
      partPath: 'word/document.xml',
    });
  });

  test('reports unresolved patch targets', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);

    expect(
      resolveWordOoxmlPatchTargets(pkg, buildIndices(pkg), [
        {
          kind: 'paragraph-text',
          op: 'insert',
          paraId: 'P00000FF',
          offset: 0,
          text: 'x',
        },
      ]),
    ).toMatchObject({
      resolved: [],
      unresolved: [
        {
          reason: 'paragraph not found: P00000FF',
        },
      ],
    });
  });

  test('applies simple text and drawing patches to package XML', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>
          <w:p w14:paraId="P0000002"><w:r><w:drawing/></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);
    const indices = buildIndices(pkg);
    const drawingId = Array.from(indices.drawingIdx.keys())[0]!;

    const result = applyWordOoxmlPatches(pkg, indices, [
      {
        kind: 'paragraph-text',
        op: 'insert',
        paraId: 'P0000001',
        offset: 5,
        text: '!',
      },
      {
        kind: 'drawing-delete',
        paraId: 'P0000002',
        drawingId,
        offset: 0,
      },
    ]);

    expect(result).toMatchObject({
      skipped: [],
      unresolved: [],
      applied: [
        { kind: 'paragraph-text', paraId: 'P0000001' },
        { kind: 'drawing-delete', paraId: 'P0000002' },
      ],
    });
    expect(isDirty(getPart(pkg, 'word/document.xml')!)).toBe(true);
    const xml = toFlatOpc(pkg);
    expect(xml).toContain('Hello!');
    expect(xml).not.toContain('<w:drawing');
  });

  test('applies paragraph merge patches and preserves text order after reload', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p>
          <w:p w14:paraId="P0000002"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>World</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);

    expect(
      applyWordOoxmlPatches(pkg, buildIndices(pkg), [
        {
          kind: 'paragraph-merge',
          paraId: 'P0000001',
        },
      ]),
    ).toMatchObject({
      applied: [{ kind: 'paragraph-merge', paraId: 'P0000001' }],
      unresolved: [],
      skipped: [],
    });
    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    const view = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(view.paragraphs.map((paragraph) => ({ paraId: paragraph.paraId, text: paragraph.text }))).toEqual([
      { paraId: 'P0000001', text: 'HelloWorld' },
    ]);
  });

  test('skips paragraph merge when there is no following paragraph', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body><w:p w14:paraId="P0000001"><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);

    const result = applyWordOoxmlPatches(pkg, buildIndices(pkg), [
      {
        kind: 'paragraph-merge',
        paraId: 'P0000001',
      },
    ]);

    expect(result).toMatchObject({
      applied: [],
      unresolved: [],
      skipped: [{ reason: 'next paragraph not found for merge' }],
    });
  });

  test('applies paragraph split patches and preserves text order after reload', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);

    const result = applyWordOoxmlPatches(pkg, buildIndices(pkg), [
      {
        kind: 'paragraph-split',
        paraId: 'P0000001',
        newParaId: 'P0000002',
        offset: 2,
      },
    ]);

    expect(result).toMatchObject({
      applied: [{ kind: 'paragraph-split', paraId: 'P0000001', newParaId: 'P0000002', offset: 2 }],
      skipped: [],
      unresolved: [],
    });
    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    const view = buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded));
    expect(view.paragraphs.map((paragraph) => ({ paraId: paragraph.paraId, text: paragraph.text }))).toEqual([
      { paraId: 'P0000001', text: 'AB' },
      { paraId: 'P0000002', text: 'CD' },
    ]);
    expect(view.paragraphs[1]!.format).toMatchObject({ alignment: 'center' });
  });

  test('applies text replacement across multiple Word text runs', () => {
    const pkg = fromFlatOpc(`<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="P0000001"><w:r><w:t>AB</w:t></w:r><w:r><w:t>CD</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`);

    const result = applyWordOoxmlPatches(pkg, buildIndices(pkg), [
      {
        kind: 'paragraph-text',
        op: 'replace',
        paraId: 'P0000001',
        offset: 1,
        text: 'xy',
        deletedText: 'BC',
      },
    ]);

    expect(result.skipped).toEqual([]);
    const reloaded = fromFlatOpc(toFlatOpc(pkg));
    expect(buildWordEditorViewFromPackage(reloaded, buildIndices(reloaded)).paragraphs[0]!.text).toBe('AxyD');
  });
});
