import { describe, expect, test } from 'vitest';
import JSZip from 'jszip';
import { ExcelWorkbookCore as ExcelWorkbook, parseCellAddress } from '../index.js';

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
  <Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;

const SHEET_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/coco-excel" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

const DRAWING_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`;

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" state="hidden" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="Total">Sheet1!$C$2</definedName>
  </definedNames>
</workbook>`;

const SHARED_STRINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>hello</t></si>
  <si><r><t>rich</t></r><r><t> text</t></r></si>
</sst>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.00%"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><i/><u/><sz val="12"/><color rgb="FF38761D"/><name val="Aptos"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom style="dashed"/><diagonal/></border>
  </borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`;

const TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DemoTable" displayName="DemoTable" ref="A1:C2" totalsRowShown="0">
  <autoFilter ref="A1:C2"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Greeting"/>
    <tableColumn id="2" name="Inline"/>
    <tableColumn id="3" name="Flag"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Coco</author>
  </authors>
  <commentList>
    <comment ref="D2" authorId="0" shapeId="7">
      <text><t>legacy comment</t></text>
    </comment>
  </commentList>
</comments>`;

const THREADED_COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="D2" dT="2026-08-04T10:00:00" personId="{person}" id="{thread}">
    <text>threaded comment</text>
  </threadedComment>
</ThreadedComments>`;

const DRAWING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Demo Chart"/></xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Demo Image"/></xdr:nvPicPr>
      <xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/></xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>11</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="4" name="Demo Shape"/></xdr:nvSpPr>
      <xdr:spPr/>
      <xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>shape text</a:t></a:r></a:p></xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Demo Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx><c:val><c:numRef><c:f>Sheet1!$A$2:$A$10</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea>
  </c:chart>
</c:chartSpace>`;

const SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:D3"/>
  <sheetViews>
    <sheetView workbookViewId="0" view="normal" tabSelected="1" topLeftCell="B2" zoomScale="125">
      <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="C2" sqref="C2:D3"/>
    </sheetView>
  </sheetViews>
  <cols>
    <col min="1" max="1" width="18" customWidth="1"/>
    <col min="4" max="4" hidden="1" style="2"/>
  </cols>
  <sheetData>
    <row r="1" ht="28" customHeight="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="inlineStr"><is><t>inline</t></is></c>
      <c r="C1" t="b"><v>1</v></c>
    </row>
    <row r="2" hidden="1" s="3">
      <c r="A2"><v>21</v></c>
      <c r="B2"><v>2</v></c>
      <c r="C2" s="1"><f>A2*B2</f><v>42</v></c>
      <c r="D2" t="s"><v>1</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="B3:D3"/>
  </mergeCells>
  <conditionalFormatting sqref="A2:C2">
    <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="2" stopIfTrue="1">
      <formula>10</formula>
    </cfRule>
  </conditionalFormatting>
  <hyperlinks>
    <hyperlink ref="B1" r:id="rId1" display="Coco" tooltip="probe"/>
    <hyperlink ref="C1" location="Sheet1!A1" display="local"/>
  </hyperlinks>
  <dataValidations count="1">
    <dataValidation type="whole" operator="between" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2:A10" errorTitle="Invalid" error="Use 1-100" promptTitle="Amount" prompt="Enter amount">
      <formula1>1</formula1>
      <formula2>100</formula2>
    </dataValidation>
  </dataValidations>
  <tableParts count="1">
    <tablePart r:id="rId2"/>
  </tableParts>
  <drawing r:id="rId5"/>
</worksheet>`;

async function makeWorkbook(worksheetXml = SHEET_XML, workbookXml = WORKBOOK_XML): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('xl/workbook.xml', workbookXml);
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS);
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', SHEET_RELS);
  zip.file('xl/sharedStrings.xml', SHARED_STRINGS_XML);
  zip.file('xl/styles.xml', STYLES_XML);
  zip.file('xl/tables/table1.xml', TABLE_XML);
  zip.file('xl/comments1.xml', COMMENTS_XML);
  zip.file('xl/threadedComments/threadedComment1.xml', THREADED_COMMENTS_XML);
  zip.file('xl/drawings/drawing1.xml', DRAWING_XML);
  zip.file('xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS);
  zip.file('xl/charts/chart1.xml', CHART_XML);
  zip.file('xl/media/image1.png', new Uint8Array([137, 80, 78, 71]));
  zip.file('xl/worksheets/sheet1.xml', worksheetXml);
  return await zip.generateAsync({ type: 'uint8array' });
}

async function makeWorkbookWithoutLegacyComments(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await makeWorkbook());
  zip.remove('xl/comments1.xml');
  zip.file(
    '[Content_Types].xml',
    CT_XML.replace(
      /\n  <Override PartName="\/xl\/comments1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.comments\+xml"\/>/,
      '',
    ),
  );
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    SHEET_RELS.replace(
      /\n  <Relationship Id="rId3" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/comments" Target="\.\.\/comments1\.xml"\/>/,
      '',
    ),
  );
  return await zip.generateAsync({ type: 'uint8array' });
}

describe('ExcelWorkbook', () => {
  test('fromZip reads workbook sheets, shared strings, cells and formulas', async () => {
    const doc = await ExcelWorkbook.fromZip(await makeWorkbook(), { workbookId: 'test.xlsx' });
    const model = doc.toModel();

    expect(doc.workbookId).toBe('test.xlsx');
    expect(model.sharedStrings).toEqual(['hello', 'rich text']);
    expect(model.definedNames).toEqual([{ name: 'Total', text: 'Sheet1!$C$2' }]);
    expect(model.styles.cellFormats[1]).toMatchObject({
      numberFormatId: 164,
      numberFormatCode: '0.00%',
      fontId: 1,
      fillId: 2,
      borderId: 1,
      applyNumberFormat: true,
      applyFont: true,
      applyFill: true,
      applyBorder: true,
      applyAlignment: true,
      alignment: { horizontal: 'center', vertical: 'top', wrapText: true },
    });
    expect(model.styles.cellFormats[2]).toMatchObject({
      numberFormatId: 14,
      numberFormatCode: 'mm-dd-yy',
    });
    expect(model.styles.borders[0]).toEqual({});
    expect(model.sheets).toHaveLength(1);
    expect(model.sheets[0]).toMatchObject({
      id: '1',
      name: 'Sheet1',
      visibility: 'hidden',
      relationshipId: 'rId1',
      path: 'xl/worksheets/sheet1.xml',
      dimension: 'A1:D3',
      sheetViews: [
        {
          workbookViewId: 0,
          view: 'normal',
          tabSelected: true,
          topLeftCell: 'B2',
          zoomScale: 125,
          pane: {
            xSplit: 1,
            ySplit: 1,
            topLeftCell: 'B2',
            activePane: 'bottomRight',
            state: 'frozen',
          },
          selections: [
            {
              pane: 'bottomRight',
              activeCell: 'C2',
              sqref: 'C2:D3',
            },
          ],
        },
      ],
      columns: [
        { min: 1, max: 1, width: 18, customWidth: true },
        { min: 4, max: 4, hidden: true, styleId: 2 },
      ],
      rows: [
        { index: 1, height: 28, customHeight: true },
        { index: 2, hidden: true, styleId: 3 },
      ],
      mergedRanges: [{ address: 'B3:D3' }],
      hyperlinks: [
        {
          ref: 'B1',
          relationshipId: 'rId1',
          target: 'https://example.com/coco-excel',
          targetMode: 'External',
          display: 'Coco',
          tooltip: 'probe',
        },
        {
          ref: 'C1',
          location: 'Sheet1!A1',
          display: 'local',
        },
      ],
      tables: [
        {
          id: 1,
          name: 'DemoTable',
          displayName: 'DemoTable',
          relationshipId: 'rId2',
          path: 'xl/tables/table1.xml',
          ref: 'A1:C2',
          totalsRowShown: false,
          autoFilterRef: 'A1:C2',
          columns: [
            { id: 1, name: 'Greeting' },
            { id: 2, name: 'Inline' },
            { id: 3, name: 'Flag' },
          ],
          style: {
            name: 'TableStyleMedium2',
            showFirstColumn: false,
            showLastColumn: false,
            showRowStripes: true,
            showColumnStripes: false,
          },
        },
      ],
      dataValidations: [
        {
          sqref: 'A2:A10',
          type: 'whole',
          operator: 'between',
          allowBlank: true,
          showInputMessage: true,
          showErrorMessage: true,
          errorTitle: 'Invalid',
          error: 'Use 1-100',
          promptTitle: 'Amount',
          prompt: 'Enter amount',
          formula1: '1',
          formula2: '100',
        },
      ],
      conditionalFormats: [
        {
          sqref: 'A2:C2',
          rules: [
            {
              type: 'cellIs',
              priority: 1,
              operator: 'greaterThan',
              dxfId: 2,
              stopIfTrue: true,
              formulas: ['10'],
            },
          ],
        },
      ],
      comments: [
        {
          ref: 'D2',
          authorId: 0,
          author: 'Coco',
          text: 'legacy comment',
          shapeId: 7,
        },
      ],
      threadedComments: [
        {
          ref: 'D2',
          id: '{thread}',
          personId: '{person}',
          dateTime: '2026-08-04T10:00:00',
          text: 'threaded comment',
        },
      ],
      drawings: [
        {
          relationshipId: 'rId5',
          path: 'xl/drawings/drawing1.xml',
          anchors: [
            {
              kind: 'twoCell',
              from: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
              to: { column: 4, columnOffset: 0, row: 8, rowOffset: 0 },
              object: {
                kind: 'chart',
                name: 'Demo Chart',
                relationshipId: 'rId1',
                targetPath: 'xl/charts/chart1.xml',
              },
            },
            {
              kind: 'twoCell',
              from: { column: 6, columnOffset: 0, row: 1, rowOffset: 0 },
              to: { column: 8, columnOffset: 0, row: 6, rowOffset: 0 },
              object: {
                kind: 'image',
                name: 'Demo Image',
                relationshipId: 'rId2',
                targetPath: 'xl/media/image1.png',
              },
            },
            {
              kind: 'twoCell',
              from: { column: 9, columnOffset: 0, row: 2, rowOffset: 0 },
              to: { column: 11, columnOffset: 0, row: 5, rowOffset: 0 },
              object: {
                kind: 'shape',
                name: 'Demo Shape',
              },
            },
          ],
        },
      ],
      charts: [
        {
          relationshipId: 'rId1',
          drawingPath: 'xl/drawings/drawing1.xml',
          path: 'xl/charts/chart1.xml',
          title: 'Demo Chart',
          formulaRefs: ['Sheet1!$A$1', 'Sheet1!$A$2:$A$10'],
        },
      ],
      images: [
        {
          relationshipId: 'rId2',
          drawingPath: 'xl/drawings/drawing1.xml',
          path: 'xl/media/image1.png',
          name: 'Demo Image',
        },
      ],
    });

    const cells = Object.fromEntries(model.sheets[0]!.cells.map((cell) => [cell.address, cell]));
    expect(cells.A1).toMatchObject({ kind: 'string', value: 'hello', sharedStringIndex: 0 });
    expect(cells.B1).toMatchObject({ kind: 'string', value: 'inline' });
    expect(cells.C1).toMatchObject({ kind: 'boolean', value: true });
    expect(cells.A2).toMatchObject({ kind: 'number', value: 21 });
    expect(cells.C2).toMatchObject({ kind: 'formula', formula: 'A2*B2', value: 42, styleId: 1 });
    expect(cells.D2).toMatchObject({ kind: 'string', value: 'rich text', sharedStringIndex: 1 });
  });

  test('fromZip reads ISO date cells as Excel serial values without shifting offsets', async () => {
    const worksheetXml = SHEET_XML.replace(
      '</sheetData>',
      `<row r="4">
        <c r="A4" t="d"><v>2026-09-28T00:00:00.000Z</v></c>
        <c r="B4" t="d"><v>2026-10-01T15:30:00+08:00</v></c>
        <c r="C4" t="d"><v>1900-02-28</v></c>
        <c r="D4" t="d"><v>not-a-date</v></c>
      </row></sheetData>`,
    );
    const doc = await ExcelWorkbook.fromZip(await makeWorkbook(worksheetXml));
    const cells = Object.fromEntries(doc.toModel().sheets[0]!.cells.map((cell) => [cell.address, cell]));

    expect(cells.A4).toMatchObject({ kind: 'number', value: 46_293, dataType: 'd' });
    expect(cells.B4).toMatchObject({ kind: 'number', value: 46_296.64583333333, dataType: 'd' });
    expect(cells.C4).toMatchObject({ kind: 'number', value: 59, dataType: 'd' });
    expect(cells.D4).toMatchObject({ kind: 'string', value: 'not-a-date', dataType: 'd' });
  });

  test('fromZip applies and preserves the 1904 date system', async () => {
    const worksheetXml = SHEET_XML.replace(
      '</sheetData>',
      '<row r="4"><c r="A4" t="d"><v>1904-01-02</v></c></row></sheetData>',
    );
    const workbookXml = WORKBOOK_XML.replace('<sheets>', '<workbookPr date1904="1"/><sheets>');
    const doc = await ExcelWorkbook.fromZip(await makeWorkbook(worksheetXml, workbookXml));

    expect(doc.toModel().date1904).toBe(true);
    expect(doc.worksheets.getItem('Sheet1').getCell('A4').value()).toBe(1);

    doc.worksheets.getItem('Sheet1').insertRows(2);
    const roundTripped = await ExcelWorkbook.fromZip(await doc.toZip());
    expect(roundTripped.toModel().date1904).toBe(true);
  });

  test('parseCellAddress converts A1 notation to one-based coordinates', () => {
    expect(parseCellAddress('A1')).toEqual({ address: 'A1', row: 1, column: 1 });
    expect(parseCellAddress('Z9')).toEqual({ address: 'Z9', row: 9, column: 26 });
    expect(parseCellAddress('AA10')).toEqual({ address: 'AA10', row: 10, column: 27 });
    expect(parseCellAddress('$aa$10')).toEqual({ address: 'AA10', row: 10, column: 27 });
  });

  test('object API exposes workbook, worksheet, range and cell reads', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');
    const used = sheet.getUsedRange();

    expect(workbook.worksheets.items().map((item) => item.name())).toEqual(['Sheet1']);
    expect(workbook.names.getItem('Total').text()).toBe('Sheet1!$C$2');
    expect(sheet.id()).toBe('1');
    expect(sheet.visibility()).toBe('hidden');
    expect(sheet.dimension()).toBe('A1:D3');
    expect(sheet.sheetViews()).toEqual([
      {
        workbookViewId: 0,
        view: 'normal',
        tabSelected: true,
        topLeftCell: 'B2',
        zoomScale: 125,
        pane: {
          xSplit: 1,
          ySplit: 1,
          topLeftCell: 'B2',
          activePane: 'bottomRight',
          state: 'frozen',
        },
        selections: [
          {
            pane: 'bottomRight',
            activeCell: 'C2',
            sqref: 'C2:D3',
          },
        ],
      },
    ]);
    expect(workbook.styles().fonts[1]).toMatchObject({
      name: 'Aptos',
      size: 12,
      bold: true,
      italic: true,
      underline: true,
      color: { rgb: '#38761D' },
    });
    expect(workbook.styles().fills[2]).toEqual({
      patternType: 'solid',
      foregroundColor: { rgb: '#FFF2CC' },
      backgroundColor: { rgb: '#000000' },
    });
    expect(workbook.styles().borders[1]).toMatchObject({
      left: { style: 'thin', color: { rgb: '#000000' } },
      bottom: { style: 'dashed' },
    });
    expect(sheet.columns()).toEqual([
      { min: 1, max: 1, width: 18, customWidth: true },
      { min: 4, max: 4, hidden: true, styleId: 2 },
    ]);
    expect(sheet.rows()).toEqual([
      { index: 1, height: 28, customHeight: true },
      { index: 2, hidden: true, styleId: 3 },
    ]);
    expect(sheet.mergedRanges()).toEqual([{ address: 'B3:D3' }]);
    expect(sheet.hyperlinks.toModel()).toEqual([
      {
        ref: 'B1',
        relationshipId: 'rId1',
        target: 'https://example.com/coco-excel',
        targetMode: 'External',
        display: 'Coco',
        tooltip: 'probe',
      },
      {
        ref: 'C1',
        location: 'Sheet1!A1',
        display: 'local',
      },
    ]);
    expect(sheet.tables.toModel()).toEqual([
      {
        id: 1,
        name: 'DemoTable',
        displayName: 'DemoTable',
        relationshipId: 'rId2',
        path: 'xl/tables/table1.xml',
        ref: 'A1:C2',
        totalsRowShown: false,
        autoFilterRef: 'A1:C2',
        columns: [
          { id: 1, name: 'Greeting' },
          { id: 2, name: 'Inline' },
          { id: 3, name: 'Flag' },
        ],
        style: {
          name: 'TableStyleMedium2',
          showFirstColumn: false,
          showLastColumn: false,
          showRowStripes: true,
          showColumnStripes: false,
        },
      },
    ]);
    expect(sheet.dataValidations.toModel()).toEqual([
      {
        sqref: 'A2:A10',
        type: 'whole',
        operator: 'between',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        errorTitle: 'Invalid',
        error: 'Use 1-100',
        promptTitle: 'Amount',
        prompt: 'Enter amount',
        formula1: '1',
        formula2: '100',
      },
    ]);
    expect(sheet.conditionalFormats.toModel()).toEqual([
      {
        sqref: 'A2:C2',
        rules: [
          {
            type: 'cellIs',
            priority: 1,
            operator: 'greaterThan',
            dxfId: 2,
            stopIfTrue: true,
            formulas: ['10'],
          },
        ],
      },
    ]);
    expect(sheet.comments.toModel()).toEqual([
      {
        ref: 'D2',
        authorId: 0,
        author: 'Coco',
        text: 'legacy comment',
        shapeId: 7,
      },
    ]);
    expect(sheet.threadedComments.toModel()).toEqual([
      {
        ref: 'D2',
        id: '{thread}',
        personId: '{person}',
        dateTime: '2026-08-04T10:00:00',
        text: 'threaded comment',
      },
    ]);
    expect(sheet.drawings.toModel()).toEqual([
      {
        relationshipId: 'rId5',
        path: 'xl/drawings/drawing1.xml',
        anchors: [
          {
            kind: 'twoCell',
            from: { column: 1, columnOffset: 0, row: 1, rowOffset: 0 },
            to: { column: 4, columnOffset: 0, row: 8, rowOffset: 0 },
            object: {
              kind: 'chart',
              name: 'Demo Chart',
              relationshipId: 'rId1',
              targetPath: 'xl/charts/chart1.xml',
            },
          },
          {
            kind: 'twoCell',
            from: { column: 6, columnOffset: 0, row: 1, rowOffset: 0 },
            to: { column: 8, columnOffset: 0, row: 6, rowOffset: 0 },
            object: {
              kind: 'image',
              name: 'Demo Image',
              relationshipId: 'rId2',
              targetPath: 'xl/media/image1.png',
            },
          },
          {
            kind: 'twoCell',
            from: { column: 9, columnOffset: 0, row: 2, rowOffset: 0 },
            to: { column: 11, columnOffset: 0, row: 5, rowOffset: 0 },
            object: {
              kind: 'shape',
              name: 'Demo Shape',
            },
          },
        ],
      },
    ]);
    expect(sheet.charts.toModel()).toEqual([
      {
        relationshipId: 'rId1',
        drawingPath: 'xl/drawings/drawing1.xml',
        path: 'xl/charts/chart1.xml',
        title: 'Demo Chart',
        formulaRefs: ['Sheet1!$A$1', 'Sheet1!$A$2:$A$10'],
      },
    ]);
    expect(sheet.images.toModel()).toEqual([
      {
        relationshipId: 'rId2',
        drawingPath: 'xl/drawings/drawing1.xml',
        path: 'xl/media/image1.png',
        name: 'Demo Image',
      },
    ]);
    expect(sheet.images.items()[0]?.name()).toBe('Demo Image');
    expect(sheet.images.items()[0]?.path()).toBe('xl/media/image1.png');
    expect(sheet.shapes.toModel()).toEqual([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        name: 'Demo Shape',
        anchor: {
          kind: 'twoCell',
          from: { column: 9, columnOffset: 0, row: 2, rowOffset: 0 },
          to: { column: 11, columnOffset: 0, row: 5, rowOffset: 0 },
          object: {
            kind: 'shape',
            name: 'Demo Shape',
          },
        },
      },
    ]);
    expect(sheet.shapes.items()[0]?.name()).toBe('Demo Shape');
    expect(used?.address()).toBe('A1:D3');

    const range = sheet.getRange('A1:D2');
    expect(range.rowCount()).toBe(2);
    expect(range.columnCount()).toBe(4);
    expect(range.values()).toEqual([
      ['hello', 'inline', true, null],
      [21, 2, 42, 'rich text'],
    ]);
    expect(range.formulas()).toEqual([
      [null, null, null, null],
      [null, null, 'A2*B2', null],
    ]);
    expect(range.valueTypes()).toEqual([
      ['string', 'string', 'boolean', 'blank'],
      ['number', 'number', 'formula', 'string'],
    ]);
    expect(range.getCell(1, 2).formula()).toBe('A2*B2');
    expect(range.getRow(1).values()).toEqual([[21, 2, 42, 'rich text']]);
    expect(range.getColumn(0).values()).toEqual([['hello'], [21]]);

    const cell = sheet.getCell('C2');
    expect(cell.address()).toBe('C2');
    expect(cell.value()).toBe(42);
    expect(cell.formula()).toBe('A2*B2');
    expect(cell.valueType()).toBe('formula');
    expect(cell.toRange().values()).toEqual([[42]]);
  });

  test('fromModel exposes the same object API for live-object bridges', async () => {
    const source = await ExcelWorkbook.fromZip(await makeWorkbook());
    const model = source.toModel();
    const doc = ExcelWorkbook.fromModel(model, { workbookId: 'live-model' });

    model.sheets[0]!.cells[0]!.value = 'mutated after load';

    expect(doc.workbookId).toBe('live-model');
    expect(doc.getPackage().parts.size).toBe(0);
    expect(doc.worksheets.items().map((sheet) => sheet.name())).toEqual(['Sheet1']);
    expect(doc.worksheets.getItem('Sheet1').getRange('A1:C2').values()).toEqual([
      ['hello', 'inline', true],
      [21, 2, 42],
    ]);
    expect(doc.worksheets.getItem('Sheet1').tables.toModel()[0]?.name).toBe('DemoTable');
  });

  test('toZip writes package-backed workbook mutations back to xlsx parts', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.getCell('A1').setValue('changed');
    sheet.insertRows(2);

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(outputSheet.dimension()).toBe('A1:D3');
    expect(outputSheet.getCell('A1').value()).toBe('changed');
    expect(outputSheet.getCell('C3').formula()).toBe('A3*B3');
    expect(outputSheet.tables.toModel()[0]?.ref).toBe('A1:C3');
    expect(outputSheet.tables.toModel()[0]?.autoFilterRef).toBe('A1:C3');
    expect(outputSheet.comments.toModel()[0]?.ref).toBe('D3');
    expect(outputSheet.threadedComments.toModel()[0]?.ref).toBe('D3');
    expect(outputSheet.charts.toModel()[0]?.formulaRefs).toEqual([
      'Sheet1!$A$1',
      'Sheet1!$A$3:$A$11',
    ]);
    expect(outputSheet.drawings.toModel()[0]?.anchors[0]?.from).toMatchObject({
      row: 2,
      column: 1,
    });
    expect(outputSheet.drawings.toModel()[0]?.anchors[0]?.to).toMatchObject({ row: 9, column: 4 });
  });

  test('single-cell value edits preserve unrelated workbook, style and worksheet OOXML', async () => {
    const input = await JSZip.loadAsync(await makeWorkbook());
    const workbookXml = WORKBOOK_XML.replace(
      '  <sheets>',
      '  <bookViews><workbookView activeTab="0" showHorizontalScroll="1" showSheetTabs="1"/></bookViews>\n  <sheets>',
    ).replace('</workbook>', '  <calcPr calcId="191029" fullCalcOnLoad="1"/>\n</workbook>');
    const stylesXml = STYLES_XML.replace(
      '</styleSheet>',
      '  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>\n  <colors><indexedColors><rgbColor rgb="FF010203"/></indexedColors></colors>\n</styleSheet>',
    );
    const sheetXml = SHEET_XML.replace(
      '  <dimension ref="A1:D3"/>',
      '  <sheetPr><outlinePr summaryBelow="1"/><pageSetUpPr fitToPage="1"/></sheetPr>\n  <dimension ref="A1:D3"/>\n  <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>',
    )
      .replace('<c r="B2"><v>2</v></c>', '<c r="B2" s="1"><v>2</v></c>')
      .replace(
        '</worksheet>',
        '  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\n</worksheet>',
      );
    input.file('xl/workbook.xml', workbookXml);
    input.file('xl/styles.xml', stylesXml);
    input.file('xl/worksheets/sheet1.xml', sheetXml);
    const workbookRels = await input.file('xl/_rels/workbook.xml.rels')!.async('string');
    const bytes = await input.generateAsync({ type: 'uint8array' });

    const workbook = await ExcelWorkbook.fromZip(bytes);
    workbook.worksheets
      .getItem('Sheet1')
      .getRange('B2')
      .setValues([[0.8]]);
    const outputBytes = await workbook.toZip();
    const output = await JSZip.loadAsync(outputBytes);

    expect(await output.file('xl/workbook.xml')!.async('string')).toBe(workbookXml);
    expect(await output.file('xl/styles.xml')!.async('string')).toBe(stylesXml);
    expect(await output.file('xl/_rels/workbook.xml.rels')!.async('string')).toBe(workbookRels);

    const outputSheetXml = await output.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(outputSheetXml).toContain('<sheetPr>');
    expect(outputSheetXml).toContain('<dimension ref="A1:D3"/>');
    expect(outputSheetXml).toContain('<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>');
    expect(outputSheetXml).toContain('<pageMargins left="0.7"');

    const roundTripped = await ExcelWorkbook.fromZip(outputBytes);
    const cell = roundTripped.worksheets.getItem('Sheet1').getCell('B2');
    expect(cell.value()).toBe(0.8);
    expect(cell.styleId()).toBe(1);

    roundTripped.worksheets
      .getItem('Sheet1')
      .getRange('B2')
      .setFormulas([['=A2*2']]);
    const formulaBytes = await roundTripped.toZip();
    const formulaOutput = await JSZip.loadAsync(formulaBytes);
    expect(await formulaOutput.file('xl/workbook.xml')!.async('string')).toBe(workbookXml);
    expect(await formulaOutput.file('xl/styles.xml')!.async('string')).toBe(stylesXml);
    expect(await formulaOutput.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
      '<sheetPr>',
    );
    const formulaRoundTrip = await ExcelWorkbook.fromZip(formulaBytes);
    expect(formulaRoundTrip.worksheets.getItem('Sheet1').getCell('B2').formula()).toBe('A2*2');
    expect(formulaRoundTrip.worksheets.getItem('Sheet1').getCell('B2').styleId()).toBe(1);
  });

  test.each(['shared', 'array', 'dataTable'] as const)(
    'rejects a partial edit of an existing %s formula group',
    async (formulaType) => {
      const input = await JSZip.loadAsync(await makeWorkbook());
      input.file(
        'xl/worksheets/sheet1.xml',
        SHEET_XML.replace('<f>A2*B2</f>', `<f t="${formulaType}" ref="C2:C3" si="0">A2*B2</f>`),
      );
      const workbook = await ExcelWorkbook.fromZip(
        await input.generateAsync({ type: 'uint8array' }),
      );
      workbook.worksheets
        .getItem('Sheet1')
        .getRange('C2')
        .setValues([[99]]);

      await expect(workbook.toZip()).rejects.toThrow(
        `does not support ${formulaType} formula cell Sheet1!C2`,
      );
    },
  );

  test('invalidates calcChain and requests full recalculation after a cell patch', async () => {
    const input = await JSZip.loadAsync(await makeWorkbook());
    input.file(
      '[Content_Types].xml',
      CT_XML.replace(
        '</Types>',
        '  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>\n</Types>',
      ),
    );
    input.file(
      'xl/_rels/workbook.xml.rels',
      WORKBOOK_RELS.replace(
        '</Relationships>',
        '  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>\n</Relationships>',
      ),
    );
    input.file(
      'xl/workbook.xml',
      WORKBOOK_XML.replace('</workbook>', '  <calcPr calcId="191029"/>\n</workbook>'),
    );
    input.file(
      'xl/calcChain.xml',
      '<?xml version="1.0" encoding="UTF-8"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>',
    );
    const inputBytes = await input.generateAsync({ type: 'uint8array' });
    const cleanWorkbook = await ExcelWorkbook.fromZip(inputBytes);
    const cleanOutput = await JSZip.loadAsync(await cleanWorkbook.toZip());
    expect(cleanOutput.file('xl/calcChain.xml')).not.toBeNull();
    expect(await cleanOutput.file('xl/workbook.xml')!.async('string')).toBe(
      await input.file('xl/workbook.xml')!.async('string'),
    );

    const workbook = await ExcelWorkbook.fromZip(inputBytes);
    workbook.worksheets
      .getItem('Sheet1')
      .getRange('B2')
      .setValues([[3]]);
    const output = await JSZip.loadAsync(await workbook.toZip());

    expect(output.file('xl/calcChain.xml')).toBeNull();
    expect(await output.file('xl/_rels/workbook.xml.rels')!.async('string')).not.toContain(
      '/calcChain',
    );
    expect(await output.file('[Content_Types].xml')!.async('string')).not.toContain(
      '/xl/calcChain.xml',
    );
    const outputWorkbook = await output.file('xl/workbook.xml')!.async('string');
    expect(outputWorkbook).toContain('calcId="191029"');
    expect(outputWorkbook).toContain('fullCalcOnLoad="1"');
    expect(outputWorkbook).toContain('forceFullCalc="1"');
  });

  test('keeps every unrelated worksheet byte-identical during a cell patch', async () => {
    const input = await JSZip.loadAsync(await makeWorkbook());
    const secondSheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><outlinePr summaryBelow="0"/></sheetPr>
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>untouched</t></is></c></row></sheetData>
  <pageMargins left="1" right="1" top="1" bottom="1" header="0.5" footer="0.5"/>
</worksheet>`;
    input.file(
      '[Content_Types].xml',
      CT_XML.replace(
        '</Types>',
        '  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n</Types>',
      ),
    );
    input.file(
      'xl/_rels/workbook.xml.rels',
      WORKBOOK_RELS.replace(
        '</Relationships>',
        '  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>\n</Relationships>',
      ),
    );
    input.file(
      'xl/workbook.xml',
      WORKBOOK_XML.replace(
        '  </sheets>',
        '    <sheet name="Untouched" sheetId="2" r:id="rId8"/>\n  </sheets>',
      ),
    );
    input.file('xl/worksheets/sheet2.xml', secondSheet);
    const workbook = await ExcelWorkbook.fromZip(await input.generateAsync({ type: 'uint8array' }));
    workbook.worksheets
      .getItem('Sheet1')
      .getRange('B2')
      .setValues([[4]]);
    const output = await JSZip.loadAsync(await workbook.toZip());

    expect(await output.file('xl/worksheets/sheet2.xml')!.async('string')).toBe(secondSheet);
  });

  test('drawing anchors and chart formula refs expose live writeback operations', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    const anchor = sheet.drawings.items()[0]!.anchors.getItem(0);
    anchor.from.setRow(3);
    anchor.from.setColumn(2);
    anchor.to.setRow(10);
    anchor.to.setColumn(6);

    const chart = sheet.charts.items()[0]!;
    chart.setTitle('Renamed Chart');
    chart.formulaRefs.setItem(0, 'Sheet1!$B$1');
    chart.formulaRefs.setItem(1, 'Sheet1!$B$2:$B$10');
    sheet.shapes.getItem('Demo Shape').setName('Renamed Shape');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    const chartXml = await zip.file('xl/charts/chart1.xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(drawingXml).toContain('<xdr:col>2</xdr:col>');
    expect(drawingXml).toContain('<xdr:row>3</xdr:row>');
    expect(drawingXml).toContain('name="Renamed Shape"');
    expect(chartXml).toContain('<a:t>Renamed Chart</a:t>');
    expect(chartXml).toContain('<c:f>Sheet1!$B$1</c:f>');
    expect(chartXml).toContain('<c:f>Sheet1!$B$2:$B$10</c:f>');
    expect(outputSheet.drawings.items()[0]!.anchors.getItem(0).from.toModel()).toMatchObject({
      row: 3,
      column: 2,
    });
    expect(outputSheet.drawings.items()[0]!.anchors.getItem(0).to.toModel()).toMatchObject({
      row: 10,
      column: 6,
    });
    expect(outputSheet.charts.items()[0]!.title()).toBe('Renamed Chart');
    expect(outputSheet.charts.items()[0]!.formulaRefs.items()).toEqual([
      'Sheet1!$B$1',
      'Sheet1!$B$2:$B$10',
    ]);
    expect(outputSheet.shapes.items()[0]!.name()).toBe('Renamed Shape');
  });

  test('worksheet domain collections expose live object operations', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    const table = sheet.tables.getItem('DemoTable');
    table.setName('RenamedTable');
    table.setDisplayName('RenamedTable');
    table.setRange('A1:C3');
    table.columns.getItem('Inline').setName('Body');
    table.columns.getItem('Body').setTotalsRowLabel('Total body');
    table.style.setName('TableStyleMedium9');
    table.style.setShowFirstColumn(true);
    table.style.setShowColumnStripes(true);
    sheet.comments.getItem('D2').delete();
    sheet.comments.add({ ref: 'D3', author: 'Coco', text: 'new comment' });
    sheet.comments.getItem('D3').setText('updated comment');
    sheet.threadedComments.getItem('D2').setRef('D3');
    sheet.threadedComments.getItem('D3').setText('updated threaded comment');
    sheet.hyperlinks.delete('B1');
    sheet.hyperlinks.add({ ref: 'A1', location: 'Sheet1!D3', display: 'jump' });
    sheet.hyperlinks.add({
      ref: 'D1',
      target: 'https://example.com/new-link',
      targetMode: 'External',
      display: 'external',
    });
    sheet.dataValidations.getItem('A2:A10').delete();
    sheet.dataValidations.add({ sqref: 'A2:A3', type: 'whole', formula1: '1', formula2: '10' });
    const validation = sheet.dataValidations.getItem('A2:A3');
    validation.setOperator('between');
    validation.setAllowBlank(false);
    validation.setShowInputMessage(true);
    validation.setShowErrorMessage(true);
    validation.setPrompt('Small amount', 'Enter 1-10');
    validation.setError('Out of range', 'Use 1-10');
    sheet.conditionalFormats.getItem('A2:C2').delete();
    sheet.conditionalFormats.add({
      sqref: 'A2:C3',
      rules: [{ type: 'cellIs', priority: 9, operator: 'greaterThan', formulas: ['5'] }],
    });
    const conditionalFormat = sheet.conditionalFormats.getItem('A2:C3');
    const rule = conditionalFormat.rules.getItem(0);
    rule.setPriority(3);
    rule.setOperator('lessThan');
    rule.setFormula(0, '50');
    rule.setStopIfTrue(false);
    rule.addFormula('100');
    rule.deleteFormula(1);

    expect(sheet.tables.getItem('RenamedTable').rangeAddress()).toBe('A1:C3');
    expect(sheet.tables.getItem('RenamedTable').columns.getItem('Body').totalsRowLabel()).toBe(
      'Total body',
    );
    expect(sheet.tables.getItem('RenamedTable').style.showFirstColumn()).toBe(true);
    expect(sheet.toModel().comments).toEqual([
      { ref: 'D3', author: 'Coco', text: 'updated comment' },
    ]);
    expect(sheet.toModel().threadedComments[0]).toMatchObject({
      ref: 'D3',
      text: 'updated threaded comment',
    });
    expect(sheet.toModel().hyperlinks.map((hyperlink) => hyperlink.ref)).toEqual([
      'C1',
      'A1',
      'D1',
    ]);
    expect(sheet.toModel().dataValidations[0]?.sqref).toBe('A2:A3');
    expect(sheet.dataValidations.getItem('A2:A3').prompt()).toBe('Enter 1-10');
    expect(sheet.toModel().conditionalFormats[0]?.sqref).toBe('A2:C3');
    expect(sheet.conditionalFormats.getItem('A2:C3').rules.getItem(0).formulas()).toEqual(['50']);

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');
    expect(outputSheet.tables.getItem('RenamedTable').rangeAddress()).toBe('A1:C3');
    expect(outputSheet.tables.getItem('RenamedTable').columns.getItem('Body').toModel()).toEqual({
      id: 2,
      name: 'Body',
      totalsRowLabel: 'Total body',
    });
    expect(outputSheet.tables.getItem('RenamedTable').style.toModel()).toMatchObject({
      name: 'TableStyleMedium9',
      showFirstColumn: true,
      showColumnStripes: true,
    });
    expect(outputSheet.comments.getItem('D3').ref()).toBe('D3');
    expect(outputSheet.comments.getItem('D3').text()).toBe('updated comment');
    expect(outputSheet.threadedComments.getItem('D3').text()).toBe('updated threaded comment');
    expect(outputSheet.hyperlinks.toModel().map((hyperlink) => hyperlink.ref)).toEqual([
      'C1',
      'A1',
      'D1',
    ]);
    expect(outputSheet.hyperlinks.getItem('D1').target()).toBe('https://example.com/new-link');
    expect(outputSheet.hyperlinks.getItem('D1').targetMode()).toBe('External');
    expect(outputSheet.hyperlinks.getItemOrNullObject('B1')).toBeNull();
    expect(outputSheet.dataValidations.getItem('A2:A3').toModel()).toMatchObject({
      sqref: 'A2:A3',
      type: 'whole',
      operator: 'between',
      allowBlank: false,
      showInputMessage: true,
      showErrorMessage: true,
      promptTitle: 'Small amount',
      prompt: 'Enter 1-10',
      errorTitle: 'Out of range',
      error: 'Use 1-10',
      formula1: '1',
      formula2: '10',
    });
    expect(
      outputSheet.conditionalFormats.getItem('A2:C3').rules.getItem(0).toModel(),
    ).toMatchObject({
      type: 'cellIs',
      priority: 3,
      operator: 'lessThan',
      stopIfTrue: false,
      formulas: ['50'],
    });
  });

  test('toZip creates missing legacy comment part and worksheet relationship', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbookWithoutLegacyComments());
    const sheet = workbook.worksheets.getItem('Sheet1');

    expect(sheet.comments.items()).toEqual([]);

    sheet.comments.add({ ref: 'A3', author: 'Coco', text: 'created comment part' });

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(outputSheet.comments.toModel()).toEqual([
      {
        ref: 'A3',
        authorId: 0,
        author: 'Coco',
        text: 'created comment part',
      },
    ]);
  });

  test('toZip creates new table part and worksheet relationship', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.getRange('E1:F2').setValues([
      ['Name', 'Amount'],
      ['gamma', 64],
    ]);
    sheet.tables.add({
      name: 'CreatedTable',
      displayName: 'CreatedTable',
      ref: 'E1:F2',
      autoFilterRef: 'E1:F2',
      totalsRowShown: false,
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Amount' },
      ],
      style: { name: 'TableStyleMedium4', showRowStripes: true },
    });

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');
    const table = outputSheet.tables.getItem('CreatedTable');

    expect(table.id()).toBe(2);
    expect(table.rangeAddress()).toBe('E1:F2');
    expect(table.autoFilterRangeAddress()).toBe('E1:F2');
    expect(table.toModel()).toMatchObject({
      name: 'CreatedTable',
      displayName: 'CreatedTable',
      ref: 'E1:F2',
      autoFilterRef: 'E1:F2',
      relationshipId: expect.stringMatching(/^rId\d+$/),
      path: 'xl/tables/table2.xml',
      columns: [
        { id: 1, name: 'Name' },
        { id: 2, name: 'Amount' },
      ],
      style: { name: 'TableStyleMedium4', showRowStripes: true },
    });
  });

  test('toZip removes deleted table relationship, part and content type override', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.tables.delete('DemoTable');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').tables.items()).toEqual([]);
    expect(zip.file('xl/tables/table1.xml')).toBeNull();
    expect(contentTypes).not.toContain('/xl/tables/table1.xml');
  });

  test('toZip removes deleted comment relationship, part and content type override', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.comments.delete('D2');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').comments.items()).toEqual([]);
    expect(zip.file('xl/comments1.xml')).toBeNull();
    expect(contentTypes).not.toContain('/xl/comments1.xml');
  });

  test('toZip removes deleted threaded comment relationship, part and content type override', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.threadedComments.getItem('D2').delete();

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').threadedComments.items()).toEqual([]);
    expect(zip.file('xl/threadedComments/threadedComment1.xml')).toBeNull();
    expect(contentTypes).not.toContain('/xl/threadedComments/threadedComment1.xml');
    expect(relsXml).not.toContain('/threadedComment');
  });

  test('toZip removes deleted external hyperlink relationship while preserving local hyperlinks', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.hyperlinks.delete('B1');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(outputSheet.hyperlinks.toModel()).toEqual([
      { ref: 'C1', location: 'Sheet1!A1', display: 'local' },
    ]);
    expect(sheetXml).toContain('<hyperlinks>');
    expect(sheetXml).toContain('location="Sheet1!A1"');
    expect(sheetXml).not.toContain('r:id="rId1"');
    expect(relsXml).not.toContain('/hyperlink');
    expect(relsXml).not.toContain('https://example.com/coco-excel');
  });

  test('toZip omits worksheet hyperlinks node after deleting all hyperlinks', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.hyperlinks.delete('B1');
    sheet.hyperlinks.delete('C1');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const relsXml = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').hyperlinks.items()).toEqual([]);
    expect(sheetXml).not.toContain('<hyperlinks>');
    expect(relsXml).not.toContain('/hyperlink');
  });

  test('toZip omits worksheet dataValidations node after deleting all validations', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.dataValidations.delete('A2:A10');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').dataValidations.items()).toEqual([]);
    expect(sheetXml).not.toContain('<dataValidations');
  });

  test('toZip preserves remaining data validations when deleting one validation', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.dataValidations.add({
      sqref: 'B2:B3',
      type: 'decimal',
      operator: 'greaterThan',
      formula1: '0',
    });
    sheet.dataValidations.delete('A2:A10');

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(outputSheet.dataValidations.items()).toHaveLength(1);
    expect(outputSheet.dataValidations.getItem('B2:B3').toModel()).toMatchObject({
      sqref: 'B2:B3',
      type: 'decimal',
      operator: 'greaterThan',
      formula1: '0',
    });
    expect(outputSheet.dataValidations.getItemOrNullObject('A2:A10')).toBeNull();
  });

  test('toZip omits worksheet conditionalFormatting node after deleting all conditional formats', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.conditionalFormats.delete('A2:C2');

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);

    expect(roundTripped.worksheets.getItem('Sheet1').conditionalFormats.items()).toEqual([]);
    expect(sheetXml).not.toContain('<conditionalFormatting');
  });

  test('toZip preserves remaining conditional formats when deleting one conditional format', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.conditionalFormats.add({
      sqref: 'B2:B3',
      rules: [{ type: 'cellIs', priority: 2, operator: 'greaterThan', formulas: ['0'] }],
    });
    sheet.conditionalFormats.delete('A2:C2');

    const roundTripped = await ExcelWorkbook.fromZip(await workbook.toZip());
    const outputSheet = roundTripped.worksheets.getItem('Sheet1');

    expect(outputSheet.conditionalFormats.items()).toHaveLength(1);
    expect(
      outputSheet.conditionalFormats.getItem('B2:B3').rules.getItem(0).toModel(),
    ).toMatchObject({
      type: 'cellIs',
      priority: 2,
      operator: 'greaterThan',
      formulas: ['0'],
    });
    expect(outputSheet.conditionalFormats.getItemOrNullObject('A2:C2')).toBeNull();
  });

  test('toZip creates a standalone xlsx from a model-only workbook', async () => {
    const workbook = ExcelWorkbook.fromModel({
      sharedStrings: [],
      definedNames: [{ name: 'Total', text: 'Sheet1!$C$2' }],
      styles: { numberFormats: [], fonts: [], fills: [], borders: [], cellFormats: [] },
      sheets: [
        {
          id: '1',
          name: 'Sheet1',
          dimension: 'A1:C2',
          sheetViews: [],
          columns: [],
          rows: [],
          mergedRanges: [],
          hyperlinks: [
            { ref: 'A1', target: 'https://example.com/generated', targetMode: 'External' },
          ],
          tables: [
            {
              name: 'GeneratedTable',
              displayName: 'GeneratedTable',
              ref: 'A1:C2',
              autoFilterRef: 'A1:C2',
              columns: [
                { id: 1, name: 'Name' },
                { id: 2, name: 'Amount' },
                { id: 3, name: 'Total' },
              ],
            },
          ],
          dataValidations: [
            { sqref: 'B2:B2', type: 'whole', operator: 'between', formula1: '1', formula2: '100' },
          ],
          conditionalFormats: [
            {
              sqref: 'C2:C2',
              rules: [{ type: 'cellIs', priority: 1, operator: 'greaterThan', formulas: ['10'] }],
            },
          ],
          comments: [{ ref: 'C2', author: 'Coco', text: 'generated comment' }],
          threadedComments: [],
          drawings: [],
          charts: [],
          images: [],
          cells: [
            { address: 'A1', row: 1, column: 1, kind: 'string', value: 'Name' },
            { address: 'B1', row: 1, column: 2, kind: 'string', value: 'Amount' },
            { address: 'C1', row: 1, column: 3, kind: 'string', value: 'Total' },
            { address: 'A2', row: 2, column: 1, kind: 'string', value: 'alpha' },
            { address: 'B2', row: 2, column: 2, kind: 'number', value: 21 },
            { address: 'C2', row: 2, column: 3, kind: 'formula', formula: 'B2*2', value: 42 },
          ],
        },
      ],
    });

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const roundTripped = await ExcelWorkbook.fromZip(bytes);
    const sheet = roundTripped.worksheets.getItem('Sheet1');

    expect(zip.file('xl/workbook.xml')).not.toBeNull();
    expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();
    expect(zip.file('xl/tables/table1.xml')).not.toBeNull();
    expect(zip.file('xl/comments1.xml')).not.toBeNull();
    expect(sheet.getRange('A1:C2').values()).toEqual([
      ['Name', 'Amount', 'Total'],
      ['alpha', 21, 42],
    ]);
    expect(sheet.getCell('C2').formula()).toBe('B2*2');
    expect(sheet.tables.getItem('GeneratedTable').rangeAddress()).toBe('A1:C2');
    expect(sheet.comments.getItem('C2').text()).toBe('generated comment');
    expect(sheet.hyperlinks.getItem('A1').target()).toBe('https://example.com/generated');
    expect(sheet.dataValidations.getItem('B2:B2').formula2()).toBe('100');
    expect(sheet.conditionalFormats.getItem('C2:C2').rules.getItem(0).formulas()).toEqual(['10']);
  });

  test('sheet.charts.add creates basic chart parts', async () => {
    const workbook = ExcelWorkbook.fromModel({
      sharedStrings: [],
      definedNames: [],
      styles: { numberFormats: [], fonts: [], fills: [], borders: [], cellFormats: [] },
      sheets: [
        {
          id: '1',
          name: '图表分析',
          dimension: 'A1:F3',
          sheetViews: [],
          columns: [],
          rows: [],
          mergedRanges: [],
          hyperlinks: [],
          tables: [],
          dataValidations: [],
          conditionalFormats: [],
          comments: [],
          threadedComments: [],
          drawings: [],
          charts: [],
          images: [],
          cells: [
            { address: 'A1', row: 1, column: 1, kind: 'string', value: '月份' },
            { address: 'B1', row: 1, column: 2, kind: 'string', value: '合同金额' },
            { address: 'C1', row: 1, column: 3, kind: 'string', value: '收款金额' },
            { address: 'A2', row: 2, column: 1, kind: 'string', value: '2023-03' },
            { address: 'B2', row: 2, column: 2, kind: 'number', value: 150000 },
            { address: 'C2', row: 2, column: 3, kind: 'number', value: 50000 },
            { address: 'A3', row: 3, column: 1, kind: 'string', value: '2023-04' },
            { address: 'B3', row: 3, column: 2, kind: 'number', value: 89000 },
            { address: 'C3', row: 3, column: 3, kind: 'number', value: 189000 },
            { address: 'E1', row: 1, column: 5, kind: 'string', value: '区域' },
            { address: 'F1', row: 1, column: 6, kind: 'string', value: '合同金额' },
            { address: 'E2', row: 2, column: 5, kind: 'string', value: '东部' },
            { address: 'F2', row: 2, column: 6, kind: 'number', value: 242000 },
            { address: 'E3', row: 3, column: 5, kind: 'string', value: '北部' },
            { address: 'F3', row: 3, column: 6, kind: 'number', value: 89000 },
          ],
        },
      ],
    });
    const sheet = workbook.worksheets.getItem('图表分析');

    sheet.charts.add({
      type: 'columnClustered',
      title: '每月合同金额与收款金额',
      categoryRange: '图表分析!$A$2:$A$3',
      series: [
        { name: '合同金额', values: '图表分析!$B$2:$B$3' },
        { name: '收款金额', values: '图表分析!$C$2:$C$3' },
      ],
      position: { from: { row: 0, column: 7 }, to: { row: 12, column: 14 } },
    });
    sheet.charts.add({
      type: 'pie',
      title: '按销售区域统计合同金额',
      categoryRange: '图表分析!$E$2:$E$3',
      series: [{ name: '合同金额', values: '图表分析!$F$2:$F$3' }],
      position: { from: { row: 14, column: 7 }, to: { row: 27, column: 14 } },
    });

    const bytes = await workbook.toZip();
    const zip = await JSZip.loadAsync(bytes);
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const sheetRels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
    const drawingXml = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    const drawingRels = await zip.file('xl/drawings/_rels/drawing1.xml.rels')!.async('string');
    const columnChartXml = await zip.file('xl/charts/chart1.xml')!.async('string');
    const pieChartXml = await zip.file('xl/charts/chart2.xml')!.async('string');
    const roundTripped = await ExcelWorkbook.fromZip(bytes);
    const charts = roundTripped.worksheets.getItem('图表分析').charts.toModel();

    expect(sheetXml).toContain('<drawing r:id="');
    expect(sheetRels).toContain('/drawing');
    expect(drawingXml).toContain('<c:chart r:id="');
    expect(drawingRels).toContain('/chart');
    expect(columnChartXml).toContain('<c:barChart>');
    expect(columnChartXml).toContain('<c:barDir val="col"/>');
    expect(columnChartXml).toContain('图表分析!$B$2:$B$3');
    expect(pieChartXml).toContain('<c:pieChart>');
    expect(pieChartXml).toContain('图表分析!$F$2:$F$3');
    expect(charts).toHaveLength(2);
    expect(charts[0]?.title).toBe('每月合同金额与收款金额');
    expect(charts[1]?.title).toBe('按销售区域统计合同金额');
  });

  test('renames worksheets in the in-memory model', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    sheet.rename('Renamed');

    expect(sheet.name()).toBe('Renamed');
    expect(workbook.worksheets.getItem('Renamed')).toBe(sheet);
    expect(workbook.worksheets.getItemOrNullObject('Sheet1')).toBeNull();
    expect(sheet.getRange('Renamed!A1:B1').values()).toEqual([['hello', 'inline']]);
    expect(workbook.toModel().sheets[0]!.name).toBe('Renamed');
  });

  test('rejects invalid worksheet names', async () => {
    const workbook = await ExcelWorkbook.fromZip(await makeWorkbook());
    const sheet = workbook.worksheets.getItem('Sheet1');

    expect(() => sheet.rename('')).toThrow(/must not be empty/);
    expect(() => sheet.rename('A'.repeat(32))).toThrow(/31 characters/);
    expect(() => sheet.rename('Bad/Name')).toThrow(/invalid worksheet name/);
    expect(() => sheet.rename("'Quoted")).toThrow(/apostrophe/);
    expect(() => sheet.rename("Quoted'")).toThrow(/apostrophe/);
  });

  test('toModel returns a defensive copy', async () => {
    const doc = await ExcelWorkbook.fromZip(await makeWorkbook());
    const first = doc.toModel();
    first.sheets[0]!.cells[0]!.value = 'mutated';

    expect(doc.toModel().sheets[0]!.cells[0]!.value).toBe('hello');
  });
});
